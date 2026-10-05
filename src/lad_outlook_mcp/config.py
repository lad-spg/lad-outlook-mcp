"""Konfiguration aus Umgebungsvariablen bzw. einer env-Datei."""

from __future__ import annotations

import os
from dataclasses import dataclass, field
from pathlib import Path

DEFAULT_ENV_FILE = Path("~/.config/lad-outlook-mcp/config.env").expanduser()

MODES = ("readonly", "draft", "send")
CONFIRM_MODES = ("auto", "always", "never")

# Delegierte Graph-Rechte je Modus (least privilege).
SCOPES_BY_MODE = {
    "readonly": ["User.Read", "Mail.Read"],
    "draft": ["User.Read", "Mail.ReadWrite"],
    "send": ["User.Read", "Mail.ReadWrite", "Mail.Send"],
}


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
    mode: str = "send"
    confirm_send: str = "auto"
    allowed_domains: frozenset[str] = field(default_factory=frozenset)
    allowed_recipients: frozenset[str] = field(default_factory=frozenset)
    data_dir: Path = Path("~/.local/share/lad-outlook-mcp").expanduser()
    token_cache: Path = Path("~/.config/lad-outlook-mcp/token_cache.json").expanduser()
    max_body_chars: int = 20000
    max_attachment_mb: int = 25

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
        if not client_id or not tenant_id:
            raise ConfigError(
                "OUTLOOK_CLIENT_ID und OUTLOOK_TENANT_ID müssen gesetzt sein "
                f"(Umgebung oder {DEFAULT_ENV_FILE})."
            )
        if tenant_id.lower() in ("common", "organizations", "consumers"):
            raise ConfigError("OUTLOOK_TENANT_ID muss die konkrete Tenant-ID sein, nicht 'common'.")

        mode = env.get("OUTLOOK_MODE", "send").strip().lower()
        if mode not in MODES:
            raise ConfigError(f"OUTLOOK_MODE muss einer von {MODES} sein, nicht {mode!r}.")

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

        return cls(
            client_id=client_id,
            tenant_id=tenant_id,
            mode=mode,
            confirm_send=confirm,
            allowed_domains=_split_list(env.get("OUTLOOK_ALLOWED_RECIPIENT_DOMAINS", "")),
            allowed_recipients=_split_list(env.get("OUTLOOK_ALLOWED_RECIPIENTS", "")),
            **kwargs,
        )
