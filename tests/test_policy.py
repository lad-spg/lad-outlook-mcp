import os

import pytest

from lad_outlook_mcp.config import Config, ConfigError
from lad_outlook_mcp.policy import (
    PolicyError,
    build_search,
    normalize_address,
    recipient_violations,
    resolve_inside,
    safe_filename,
)


def test_recipients_all_allowed_when_no_lists():
    assert recipient_violations(["a@extern.com"], frozenset(), frozenset()) == []


def test_recipients_domain_and_address_whitelist():
    domains = frozenset({"dension-dental.de"})
    addresses = frozenset({"steuerberater@kanzlei.de"})
    bad = recipient_violations(
        ["X@Dension-Dental.de", "steuerberater@kanzlei.de", "evil@dension-dental.de.attacker.com", "a@sub.dension-dental.de"],
        domains,
        addresses,
    )
    assert bad == ["evil@dension-dental.de.attacker.com", "a@sub.dension-dental.de"]


@pytest.mark.parametrize("addr", ["", "foo", "a@b", "a b@c.de", "<a@b.de>"])
def test_invalid_addresses(addr):
    with pytest.raises(PolicyError):
        normalize_address(addr)


def test_resolve_inside_blocks_escape(tmp_path):
    base = tmp_path / "outbox"
    base.mkdir()
    (base / "ok.pdf").write_text("x")
    (tmp_path / "secret").write_text("s")
    os.symlink(tmp_path / "secret", base / "link")
    assert resolve_inside(base, "ok.pdf") == (base / "ok.pdf").resolve()
    for name in ["../secret", "/etc/passwd", "link", "", ".", "a\x00b"]:
        with pytest.raises(PolicyError):
            resolve_inside(base, name)


def test_safe_filename():
    assert safe_filename("../../etc/passwd") == "_.._etc_passwd"
    assert safe_filename("Rechnung: 03/2026?.pdf") == "Rechnung_ 03_2026_.pdf"
    assert safe_filename("") == "anhang"


def test_build_search():
    assert build_search() is None
    q = build_search("Implantat", sender="Max Muster", subject='Angebot "neu"', since="2026-01-01", has_attachments=True)
    assert q == '"Implantat AND from:Max AND from:Muster AND subject:Angebot AND subject:neu AND received>=2026-01-01 AND hasattachments:true"'
    with pytest.raises(PolicyError):
        build_search(since="01.01.2026")


def test_config_load_and_validation(tmp_path):
    env = {"OUTLOOK_CLIENT_ID": "cid", "OUTLOOK_TENANT_ID": "tid", "OUTLOOK_ALLOWED_RECIPIENT_DOMAINS": "A.de, b.de",
           "OUTLOOK_DATA_DIR": str(tmp_path)}
    cfg = Config.load(env)
    assert cfg.mode == "send" and cfg.allowed_domains == {"a.de", "b.de"}
    assert "Mail.Send" in cfg.scopes
    assert Config.load({**env, "OUTLOOK_MODE": "readonly"}).scopes == ["User.Read", "Mail.Read"]
    with pytest.raises(ConfigError):
        Config.load({**env, "OUTLOOK_TENANT_ID": "common"})
    with pytest.raises(ConfigError):
        Config.load({**env, "OUTLOOK_MODE": "yolo"})
    with pytest.raises(ConfigError):
        Config.load({"OUTLOOK_CLIENT_ID": "cid"})  # eigene App ohne Tenant
    default = Config.load({})
    assert default.uses_default_app and default.tenant_id == "organizations"
    assert Config.load({"OUTLOOK_TENANT_ID": "tid"}).uses_default_app
