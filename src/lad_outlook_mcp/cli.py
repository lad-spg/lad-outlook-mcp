"""Kommandozeile: `lad-outlook-mcp [serve|login|status|logout]`."""

from __future__ import annotations

import argparse
import sys

from . import __version__
from .auth import Auth, NotLoggedIn
from .config import Config, ConfigError


def _err(msg: str) -> int:
    print(f"Fehler: {msg}", file=sys.stderr)
    return 1


def main(argv: list[str] | None = None) -> int:
    parser = argparse.ArgumentParser(prog="lad-outlook-mcp", description="Outlook-MCP-Server über Microsoft Graph")
    parser.add_argument("--version", action="version", version=__version__)
    parser.add_argument(
        "command", nargs="?", default="serve", choices=["serve", "login", "status", "logout"],
        help="serve (Standard, stdio-MCP-Server), login (Device-Code-Anmeldung), status, logout",
    )
    args = parser.parse_args(argv)

    try:
        config = Config.load()
    except ConfigError as e:
        return _err(str(e))
    auth = Auth(config)

    if args.command == "login":
        try:
            claims = auth.login(print_fn=lambda *a, **k: print(*a, **k, flush=True))
        except (RuntimeError, NotLoggedIn) as e:
            return _err(str(e))
        print(f"Angemeldet als {claims.get('preferred_username', '?')} (Modus: {config.mode}).")
        return 0

    if args.command == "logout":
        try:
            auth.logout()
        except NotLoggedIn:  # Tenant nicht erreichbar – Cache trotzdem löschen
            config.token_cache.unlink(missing_ok=True)
        print("Abgemeldet, Token-Cache gelöscht.")
        return 0

    if args.command == "status":
        print(f"Modus:          {config.mode}  (Rechte: {', '.join(config.scopes)})")
        print(f"Empfänger:      {', '.join(sorted(config.allowed_domains | config.allowed_recipients)) or 'alle erlaubt'}")
        print(f"Bestätigung:    {config.confirm_send}")
        print(f"Anmeldeart:     {'Browser (App mit Client-Secret)' if config.client_secret else 'Device-Code'}")
        print(f"Daten-Ordner:   {config.data_dir}")
        try:
            account = auth.account()
            if account is None:
                print("Anmeldung:      nicht angemeldet")
                return 1
            auth.get_token()
            print(f"Anmeldung:      OK ({account.get('username')})")
            return 0
        except NotLoggedIn as e:
            print(f"Anmeldung:      {e}")
            return 1

    # serve
    from .graph import Graph
    from .server import build_server

    config.outbox_dir.mkdir(parents=True, exist_ok=True)
    config.downloads_dir.mkdir(parents=True, exist_ok=True)
    build_server(config, Graph(auth.get_token)).run("stdio")
    return 0


if __name__ == "__main__":
    sys.exit(main())
