"""Konfiguration aus Umgebungsvariablen bzw. einer env-Datei."""

from __future__ import annotations

import os
from dataclasses import dataclass, field
from pathlib import Path

DEFAULT_ENV_FILE = Path("~/.config/lad-outlook-mcp/config.env").expanduser()

MODES = ("readonly", "draft", "send")
DEFAULT_MODE = "draft"
CONFIRM_MODES = ("auto", "always", "never")

# Delegierte Graph-Rechte je Modus (least privilege).
SCOPES_BY_MODE = {
    "readonly": ["User.Read", "Mail.Read"],
    "draft": ["User.Read", "Mail.ReadWrite"],
    "send": ["User.Read", "Mail.ReadWrite", "Mail.Send"],
}


# Microsoft-eigene, in jedem Tenant vorhandene Public-Client-App ("Microsoft Graph Command Line Tools",
# auch von Microsoft Graph PowerShell genutzt). Damit ist keine eigene App-Registrierung nötig.
GRAPH_CLI_CLIENT_ID = "14d82eec-204b-4c2f-b7e8-296a70dab67e"


class ConfigError(Exception):
    pass


def load_env_file(path: Path) -> dict[str, str]:
    """Einfacher KEY=VALUE-Parser (Kommentare mit #, optionale Anführungszeichen)."""
    values: dict[str, str] = {}
    if not path.is_file():
        return values
    for raw in path.read_text(encoding="utf-8").splitlines():
        line = raw.strip()
        if not line or line.startswith("#") or "=" not in line:
            continue
        key, _, value = line.partition("=")
        value = value.strip()
        if len(value) >= 2 and value[0] == value[-1] and value[0] in "\"'":
            value = value[1:-1]
        values[key.strip()] = value
    return values


def _split_list(value: str) -> frozenset[str]:
    return frozenset(v.strip().lower() for v in value.split(",") if v.strip())


@dataclass(frozen=True)
class Config:
    client_id: str
    tenant_id: str
    client_secret: str = field(default="", repr=False)
    redirect_uri: str = "http://localhost:3333/auth/callback"
    mode: str = DEFAULT_MODE
    confirm_send: str = "auto"
    allowed_domains: frozenset[str] = field(default_factory=frozenset)
    allowed_recipients: frozenset[str] = field(default_factory=frozenset)
    data_dir: Path = Path("~/.local/share/lad-outlook-mcp").expanduser()
    token_cache: Path = Path("~/.config/lad-outlook-mcp/token_cache.json").expanduser()
    max_body_chars: int = 20000
    max_attachment_mb: int = 25

    @property
    def uses_default_app(self) -> bool:
        return self.client_id == GRAPH_CLI_CLIENT_ID

    @property
    def scopes(self) -> list[str]:
        return SCOPES_BY_MODE[self.mode]

    @property
    def outbox_dir(self) -> Path:
        return self.data_dir / "outbox"

    @property
    def downloads_dir(self) -> Path:
        return self.data_dir / "downloads"

    @classmethod
    def load(cls, env: dict[str, str] | None = None) -> "Config":
        if env is None:
            env_file = Path(os.environ.get("LAD_OUTLOOK_ENV_FILE", DEFAULT_ENV_FILE)).expanduser()
            env = {**load_env_file(env_file), **os.environ}

        client_id = env.get("OUTLOOK_CLIENT_ID", "").strip()
        tenant_id = env.get("OUTLOOK_TENANT_ID", "").strip()
        if client_id and client_id != GRAPH_CLI_CLIENT_ID and not tenant_id:
            raise ConfigError("Bei eigener App (OUTLOOK_CLIENT_ID) muss auch OUTLOOK_TENANT_ID gesetzt sein.")
        client_id = client_id or GRAPH_CLI_CLIENT_ID
        # Ohne Tenant: beliebiges Firmenkonto (der Tenant ergibt sich aus der Anmeldung).
        tenant_id = tenant_id or "organizations"
        if tenant_id.lower() in ("common", "consumers"):
            raise ConfigError("OUTLOOK_TENANT_ID muss eine Tenant-ID oder 'organizations' sein.")

        mode = (env.get("OUTLOOK_MODE", "").strip() or DEFAULT_MODE).lower()
        if mode not in MODES:
            raise ConfigError(f"OUTLOOK_MODE muss einer von {MODES} sein, nicht {mode!r}.")
        # Senden nur über eine eigene App-Registrierung: Die Admin-Zustimmung für Mail.Send soll
        # gezielt für dieses Programm erteilt (und gesperrt) werden – nicht für die allgemeine
        # Microsoft-App "Graph Command Line Tools", die auch Graph PowerShell nutzt.
        if mode == "send" and client_id == GRAPH_CLI_CLIENT_ID:
            raise ConfigError(
                "OUTLOOK_MODE=send ist nur mit eigener App-Registrierung möglich "
                "(OUTLOOK_CLIENT_ID und OUTLOOK_TENANT_ID setzen, README Teil A). "
                "Ohne eigene App: OUTLOOK_MODE=draft oder readonly."
            )

        confirm = env.get("OUTLOOK_CONFIRM_SEND", "auto").strip().lower()
        if confirm not in CONFIRM_MODES:
            raise ConfigError(f"OUTLOOK_CONFIRM_SEND muss einer von {CONFIRM_MODES} sein.")

        kwargs = {}
        if env.get("OUTLOOK_DATA_DIR"):
            kwargs["data_dir"] = Path(env["OUTLOOK_DATA_DIR"]).expanduser()
        if env.get("OUTLOOK_TOKEN_CACHE"):
            kwargs["token_cache"] = Path(env["OUTLOOK_TOKEN_CACHE"]).expanduser()
        try:
            if env.get("OUTLOOK_MAX_BODY_CHARS"):
                kwargs["max_body_chars"] = int(env["OUTLOOK_MAX_BODY_CHARS"])
            if env.get("OUTLOOK_MAX_ATTACHMENT_MB"):
                kwargs["max_attachment_mb"] = int(env["OUTLOOK_MAX_ATTACHMENT_MB"])
        except ValueError as e:
            raise ConfigError(f"Ungültige Zahl in der Konfiguration: {e}") from e

        if env.get("OUTLOOK_REDIRECT_URI"):
            kwargs["redirect_uri"] = env["OUTLOOK_REDIRECT_URI"].strip()

        return cls(
            client_id=client_id,
            tenant_id=tenant_id,
            client_secret=env.get("OUTLOOK_CLIENT_SECRET", "").strip(),
            mode=mode,
            confirm_send=confirm,
            allowed_domains=_split_list(env.get("OUTLOOK_ALLOWED_RECIPIENT_DOMAINS", "")),
            allowed_recipients=_split_list(env.get("OUTLOOK_ALLOWED_RECIPIENTS", "")),
            **kwargs,
        )
