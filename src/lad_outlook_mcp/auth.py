"""Delegierte Anmeldung (immer als der Nutzer selbst, nur /me).

Zwei Wege:
- ohne Client-Secret: Device-Code-Flow (App braucht "Öffentliche Clientflows zulassen").
- mit Client-Secret (bestehende Web-App-Registrierung): Browser-Anmeldung per
  Authorization-Code-Flow mit PKCE und Redirect auf http://localhost:<port>/...
"""

from __future__ import annotations

import os
import select
import sys
import threading
from http.server import BaseHTTPRequestHandler, HTTPServer
from urllib.parse import parse_qs, urlparse
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
        self._msal_app: msal.ClientApplication | None = None
        if self._cache_path.is_file():
            self._cache.deserialize(self._cache_path.read_text(encoding="utf-8"))

    @property
    def _app(self) -> msal.ClientApplication:
        # Lazy: MSAL fragt beim Anlegen die Tenant-Konfiguration online ab. Der MCP-Server
        # soll auch ohne Netz/mit Tippfehler in der Tenant-ID starten und sauber melden.
        if self._msal_app is None:
            try:
                authority = f"https://login.microsoftonline.com/{self._config.tenant_id}"
                if self._config.client_secret:
                    self._msal_app = msal.ConfidentialClientApplication(
                        self._config.client_id,
                        client_credential=self._config.client_secret,
                        authority=authority,
                        token_cache=self._cache,
                    )
                else:
                    self._msal_app = msal.PublicClientApplication(
                        self._config.client_id, authority=authority, token_cache=self._cache
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
        if self._config.client_secret:
            result = self._login_browser(print_fn)
        else:
            result = self._login_device_code(print_fn)
        if "access_token" not in result:
            raise RuntimeError(
                f"Anmeldung fehlgeschlagen: {result.get('error')}: {result.get('error_description')}"
            )
        with self._lock:
            self._save_cache()
        return result.get("id_token_claims", {})

    def _login_device_code(self, print_fn) -> dict:
        flow = self._app.initiate_device_flow(scopes=self._config.scopes)
        if "user_code" not in flow:
            raise RuntimeError(f"Device-Code-Flow fehlgeschlagen: {flow.get('error_description', flow)}")
        print_fn(flow["message"])
        return self._app.acquire_token_by_device_flow(flow)

    def _login_browser(self, print_fn, timeout: int = 600) -> dict:
        redirect = urlparse(self._config.redirect_uri)
        if redirect.hostname not in ("localhost", "127.0.0.1"):
            raise RuntimeError("OUTLOOK_REDIRECT_URI muss auf localhost zeigen.")
        flow = self._app.initiate_auth_code_flow(self._config.scopes, redirect_uri=self._config.redirect_uri)
        received: dict[str, dict] = {}

        class Handler(BaseHTTPRequestHandler):
            def do_GET(self):  # noqa: N802
                url = urlparse(self.path)
                if url.path != redirect.path:
                    self.send_response(404)
                    self.end_headers()
                    return
                received["params"] = {k: v[0] for k, v in parse_qs(url.query).items()}
                self.send_response(200)
                self.send_header("Content-Type", "text/html; charset=utf-8")
                self.end_headers()
                self.wfile.write(
                    "<h2>Anmeldung abgeschlossen.</h2><p>Du kannst dieses Fenster schließen.</p>".encode()
                )

            def log_message(self, *args):
                pass

        server = None
        try:
            server = HTTPServer(("127.0.0.1", redirect.port or 80), Handler)
            server.timeout = 1
        except OSError as e:
            print_fn(f"(Hinweis: Port {redirect.port} ist belegt ({e}) – nur Einfügen der Adresse möglich.)")

        print_fn(
            "\n1. Diese Adresse im Browser öffnen und mit dem Firmenkonto anmelden:\n\n"
            f"   {flow['auth_uri']}\n\n"
            "2. Läuft der Browser auf diesem Rechner, geht es danach automatisch weiter.\n"
            "   Zeigt der Browser stattdessen 'Seite nicht erreichbar', die komplette Adresse\n"
            "   aus der Adresszeile kopieren, hier einfügen und Enter drücken.\n"
        )
        print_fn("Warte auf Anmeldung ...")
        stdin_ok = sys.stdin is not None and sys.stdin.isatty()
        try:
            for _ in range(timeout):
                if server is not None:
                    server.handle_request()  # wartet max. 1 s
                elif not stdin_ok:
                    break
                if "params" in received:
                    break
                if stdin_ok and select.select([sys.stdin], [], [], 0 if server else 1)[0]:
                    line = sys.stdin.readline().strip()
                    if line:
                        received["params"] = {k: v[0] for k, v in parse_qs(urlparse(line).query).items()}
                        break
        finally:
            if server is not None:
                server.server_close()
        if "params" not in received:
            raise RuntimeError("Keine Anmeldung innerhalb der Wartezeit erhalten.")
        print_fn("")
        try:
            return self._app.acquire_token_by_auth_code_flow(flow, received["params"])
        except ValueError as e:  # z. B. state passt nicht
            raise RuntimeError(f"Antwort der Anmeldung ungültig: {e}") from e

    def logout(self) -> None:
        with self._lock:
            for account in self._app.get_accounts():
                self._app.remove_account(account)
            if self._cache_path.exists():
                self._cache_path.unlink()
