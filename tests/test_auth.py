import socket
import threading
import urllib.request

from lad_outlook_mcp.auth import Auth
from lad_outlook_mcp.config import Config


def free_port():
    with socket.socket() as s:
        s.bind(("127.0.0.1", 0))
        return s.getsockname()[1]


class FakeApp:
    def __init__(self):
        self.got = None

    def initiate_auth_code_flow(self, scopes, redirect_uri):
        return {"auth_uri": "https://login.example/authorize", "state": "abc"}

    def acquire_token_by_auth_code_flow(self, flow, params):
        self.got = params
        return {"access_token": "t", "id_token_claims": {"preferred_username": "me@firma.de"}}


def test_browser_login_via_local_redirect(tmp_path):
    port = free_port()
    cfg = Config.load({
        "OUTLOOK_CLIENT_ID": "cid", "OUTLOOK_TENANT_ID": "tid", "OUTLOOK_CLIENT_SECRET": "geheim",
        "OUTLOOK_REDIRECT_URI": f"http://localhost:{port}/auth/callback",
        "OUTLOOK_TOKEN_CACHE": str(tmp_path / "tc.json"),
    })
    assert "geheim" not in repr(cfg)
    auth = Auth(cfg)
    fake = FakeApp()
    auth._msal_app = fake
    out = {}
    t = threading.Thread(target=lambda: out.update(claims=auth.login(print_fn=lambda *a, **k: None, device_code=True)))
    t.start()
    for _ in range(50):
        try:
            urllib.request.urlopen(f"http://127.0.0.1:{port}/falsch", timeout=1)
        except urllib.error.HTTPError as e:
            assert e.code == 404
            break
        except OSError:
            import time; time.sleep(0.1)
    body = urllib.request.urlopen(f"http://127.0.0.1:{port}/auth/callback?code=XYZ&state=abc", timeout=5).read()
    t.join(10)
    assert "abgeschlossen" in body.decode()
    assert fake.got == {"code": "XYZ", "state": "abc"}
    assert out["claims"]["preferred_username"] == "me@firma.de"


class FakePublicApp:
    def __init__(self):
        self.used = None

    def acquire_token_interactive(self, scopes, **kw):
        self.used = "interactive"
        return {"access_token": "t", "id_token_claims": {"preferred_username": "me@firma.de"}}

    def initiate_device_flow(self, scopes):
        return {"user_code": "ABC", "message": "Code ABC eingeben"}

    def acquire_token_by_device_flow(self, flow):
        self.used = "device"
        return {"access_token": "t", "id_token_claims": {}}


def _public_auth(tmp_path):
    auth = Auth(Config.load({"OUTLOOK_TOKEN_CACHE": str(tmp_path / "tc.json")}))
    auth._msal_app = FakePublicApp()
    return auth


def test_login_picks_browser_with_display(tmp_path, monkeypatch):
    monkeypatch.setenv("DISPLAY", ":0")
    auth = _public_auth(tmp_path)
    auth.login(print_fn=lambda *a, **k: None)
    assert auth._msal_app.used == "interactive"


def test_login_picks_device_code_without_display_or_on_request(tmp_path, monkeypatch):
    monkeypatch.delenv("DISPLAY", raising=False)
    monkeypatch.delenv("WAYLAND_DISPLAY", raising=False)
    auth = _public_auth(tmp_path)
    auth.login(print_fn=lambda *a, **k: None)
    assert auth._msal_app.used == "device"

    monkeypatch.setenv("DISPLAY", ":0")
    auth = _public_auth(tmp_path)
    auth.login(print_fn=lambda *a, **k: None, device_code=True)
    assert auth._msal_app.used == "device"
