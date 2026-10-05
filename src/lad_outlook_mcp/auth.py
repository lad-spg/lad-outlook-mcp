"""Delegierte Anmeldung per Device-Code (MSAL, Public Client – kein Client-Secret)."""

from __future__ import annotations

import os
import threading
from pathlib import Path

import msal

from .config import Config


class NotLoggedIn(Exception):
    pass


class Auth:
    def __init__(self, config: Config):
        self._config = config
        self._cache_path: Path = config.token_cache
        self._cache = msal.SerializableTokenCache()
        self._lock = threading.Lock()
        self._msal_app: msal.PublicClientApplication | None = None
        if self._cache_path.is_file():
            self._cache.deserialize(self._cache_path.read_text(encoding="utf-8"))

    @property
    def _app(self) -> msal.PublicClientApplication:
        # Lazy: MSAL fragt beim Anlegen die Tenant-Konfiguration online ab. Der MCP-Server
        # soll auch ohne Netz/mit Tippfehler in der Tenant-ID starten und sauber melden.
        if self._msal_app is None:
            try:
                self._msal_app = msal.PublicClientApplication(
                    self._config.client_id,
                    authority=f"https://login.microsoftonline.com/{self._config.tenant_id}",
                    token_cache=self._cache,
                )
            except ValueError as e:
                raise NotLoggedIn(
                    f"Anmeldedienst nicht erreichbar oder Tenant-ID falsch ({self._config.tenant_id}): "
                    f"{str(e).splitlines()[0][:200]}"
                ) from e
        return self._msal_app

    def _save_cache(self) -> None:
        if not self._cache.has_state_changed:
            return
        self._cache_path.parent.mkdir(parents=True, exist_ok=True)
        os.chmod(self._cache_path.parent, 0o700)
        tmp = self._cache_path.with_suffix(".tmp")
        fd = os.open(tmp, os.O_WRONLY | os.O_CREAT | os.O_TRUNC, 0o600)
        with os.fdopen(fd, "w", encoding="utf-8") as f:
            f.write(self._cache.serialize())
        os.replace(tmp, self._cache_path)
        self._cache.has_state_changed = False

    def account(self) -> dict | None:
        accounts = self._app.get_accounts()
        return accounts[0] if accounts else None

    def get_token(self) -> str:
        with self._lock:
            account = self.account()
            if account is None:
                raise NotLoggedIn("Nicht angemeldet. Bitte 'lad-outlook-mcp login' ausführen.")
            result = self._app.acquire_token_silent(self._config.scopes, account=account)
            self._save_cache()
            if not result or "access_token" not in result:
                detail = (result or {}).get("error_description", "kein gültiges Token im Cache")
                raise NotLoggedIn(
                    f"Anmeldung abgelaufen ({detail}). Bitte 'lad-outlook-mcp login' ausführen."
                )
            return result["access_token"]

    def login(self, print_fn=print) -> dict:
        flow = self._app.initiate_device_flow(scopes=self._config.scopes)
        if "user_code" not in flow:
            raise RuntimeError(f"Device-Code-Flow fehlgeschlagen: {flow.get('error_description', flow)}")
        print_fn(flow["message"])
        result = self._app.acquire_token_by_device_flow(flow)
        if "access_token" not in result:
            raise RuntimeError(
                f"Anmeldung fehlgeschlagen: {result.get('error')}: {result.get('error_description')}"
            )
        with self._lock:
            self._save_cache()
        return result.get("id_token_claims", {})

    def logout(self) -> None:
        with self._lock:
            for account in self._app.get_accounts():
                self._app.remove_account(account)
            if self._cache_path.exists():
                self._cache_path.unlink()
