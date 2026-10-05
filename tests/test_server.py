"""Server-Tests gegen ein gemocktes Microsoft Graph über den In-Memory-MCP-Client."""

import asyncio
import json

import httpx
import mcp_types as types
import pytest
from mcp.client import Client

from lad_outlook_mcp.config import Config
from lad_outlook_mcp.graph import Graph
from lad_outlook_mcp.server import build_server

DRAFT = {
    "id": "draft/1=",
    "subject": "Angebot",
    "isDraft": True,
    "toRecipients": [{"emailAddress": {"address": "kollege@dension-dental.de", "name": "Kollege"}}],
    "ccRecipients": [],
    "bccRecipients": [],
    "body": {"contentType": "text", "content": "Hallo"},
}


class FakeGraph:
    """Minimaler Graph-Ersatz; protokolliert alle Requests."""

    def __init__(self, draft=None):
        self.calls: list[tuple[str, str, dict | None]] = []
        self.draft = dict(draft or DRAFT)

    def handler(self, request: httpx.Request) -> httpx.Response:
        path = request.url.raw_path.decode().split("?")[0].removeprefix("/v1.0")
        body = json.loads(request.content) if request.content else None
        self.calls.append((request.method, path, body))
        m = request.method
        if path == "/me/messages" and m == "GET":
            return httpx.Response(200, json={"value": [{"id": "m1", "subject": "Hi", "from": {"emailAddress": {"address": "a@b.de"}}}]})
        if path == "/me/messages" and m == "POST":
            return httpx.Response(201, json={**self.draft, **{"subject": body["subject"]}})
        if path.endswith("/attachments") and m == "GET":
            return httpx.Response(200, json={"value": []})
        if path.endswith("/attachments") and m == "POST":
            return httpx.Response(201, json={"id": "att1"})
        if path.startswith("/me/messages/draft%2F1%3D") and path.endswith("/send"):
            return httpx.Response(202)
        if path == "/me/messages/draft%2F1%3D" and m == "GET":
            return httpx.Response(200, json=self.draft)
        if path == "/me/mailFolders" and m == "GET":
            return httpx.Response(200, json={"value": [{"id": "F1", "displayName": "Projekte", "childFolderCount": 0}]})
        return httpx.Response(404, json={"error": {"code": "NotFound", "message": path}})

    def sends(self):
        return [c for c in self.calls if c[1].endswith("/send")]


def make(tmp_path, fake, **env):
    cfg = Config.load({
        "OUTLOOK_CLIENT_ID": "cid", "OUTLOOK_TENANT_ID": "tid", "OUTLOOK_DATA_DIR": str(tmp_path),
        "OUTLOOK_ALLOWED_RECIPIENT_DOMAINS": "dension-dental.de", **env,
    })
    graph = Graph(lambda: "token", transport=httpx.MockTransport(fake.handler))
    return cfg, build_server(cfg, graph)


def accept(value: bool, counter: list):
    async def cb(context, params):
        counter.append(params.message)
        return types.ElicitResult(action="accept", content={"senden": value})
    return cb


def run(coro):
    return asyncio.run(coro)


async def call(server, name, args, mode="auto", elicitation_callback=None):
    async with Client(server, mode=mode, elicitation_callback=elicitation_callback) as client:
        return await client.call_tool(name, args)


def payload(result):
    assert not result.is_error, result.content
    return json.loads(result.content[0].text) if result.content else result.structured_content


def test_tools_by_mode(tmp_path):
    async def names(server):
        async with Client(server) as c:
            return {t.name for t in (await c.list_tools()).tools}

    _, ro = make(tmp_path, FakeGraph(), OUTLOOK_MODE="readonly")
    _, dr = make(tmp_path, FakeGraph(), OUTLOOK_MODE="draft")
    _, se = make(tmp_path, FakeGraph())
    ro_n, dr_n, se_n = run(names(ro)), run(names(dr)), run(names(se))
    assert "search_mails" in ro_n and "create_draft" not in ro_n
    assert "create_draft" in dr_n and "send_draft" not in dr_n
    assert "send_draft" in se_n


def test_search_uses_kql(tmp_path):
    fake = FakeGraph()
    _, server = make(tmp_path, fake)
    res = run(call(server, "search_mails", {"sender": "max", "since": "2026-01-01"}))
    assert not res.is_error
    assert fake.calls[0][0] == "GET"


@pytest.mark.parametrize("mode", ["auto", "legacy"])
def test_send_with_confirmation(tmp_path, mode):
    fake, asked = FakeGraph(), []
    _, server = make(tmp_path, fake)
    res = run(call(server, "send_draft", {"draft_id": "draft/1="}, mode=mode, elicitation_callback=accept(True, asked)))
    assert not res.is_error, res.content
    assert len(asked) == 1 and "kollege@dension-dental.de" in asked[0]
    assert len(fake.sends()) == 1  # genau einmal gesendet


@pytest.mark.parametrize("mode", ["auto", "legacy"])
def test_send_declined(tmp_path, mode):
    fake, asked = FakeGraph(), []
    _, server = make(tmp_path, fake)
    res = run(call(server, "send_draft", {"draft_id": "draft/1="}, mode=mode, elicitation_callback=accept(False, asked)))
    assert fake.sends() == []
    assert "nicht bestätigt" in res.content[0].text


def test_send_blocked_by_whitelist(tmp_path):
    fake = FakeGraph({**DRAFT, "bccRecipients": [{"emailAddress": {"address": "x@extern.com"}}]})
    _, server = make(tmp_path, fake)
    res = run(call(server, "send_draft", {"draft_id": "draft/1="}, elicitation_callback=accept(True, [])))
    assert res.is_error and "x@extern.com" in res.content[0].text
    assert fake.sends() == []


def test_send_always_requires_elicitation(tmp_path):
    fake = FakeGraph()
    _, server = make(tmp_path, fake, OUTLOOK_CONFIRM_SEND="always")
    res = run(call(server, "send_draft", {"draft_id": "draft/1="}))  # Client ohne Elicitation
    assert res.is_error and fake.sends() == []


def test_create_draft_with_outbox_attachment(tmp_path):
    fake = FakeGraph()
    cfg, server = make(tmp_path, fake)
    cfg.outbox_dir.mkdir(parents=True)
    (cfg.outbox_dir / "angebot.pdf").write_bytes(b"%PDF-1.4 test")
    res = run(call(server, "create_draft", {
        "to": ["kollege@dension-dental.de"], "subject": "Angebot", "body": "Anbei", "attachments": ["angebot.pdf"],
    }))
    assert not res.is_error, res.content
    att_posts = [c for c in fake.calls if c[0] == "POST" and c[1].endswith("/attachments")]
    assert att_posts and att_posts[0][2]["name"] == "angebot.pdf"
    assert fake.sends() == []


def test_create_draft_rejects_path_escape(tmp_path):
    fake = FakeGraph()
    _, server = make(tmp_path, fake)
    (tmp_path / "geheim.txt").write_text("x")
    res = run(call(server, "create_draft", {
        "to": ["a@dension-dental.de"], "subject": "x", "body": "x", "attachments": ["../geheim.txt"],
    }))
    assert res.is_error
    assert not any(c[0] == "POST" for c in fake.calls)  # nicht einmal ein Entwurf angelegt


@pytest.mark.parametrize("mode", ["auto", "legacy"])
def test_send_decline_button(tmp_path, mode):
    fake = FakeGraph()
    _, server = make(tmp_path, fake)

    async def decline(context, params):
        return types.ElicitResult(action="decline")

    res = run(call(server, "send_draft", {"draft_id": "draft/1="}, mode=mode, elicitation_callback=decline))
    assert fake.sends() == []
    assert not res.is_error and "nicht bestätigt" in res.content[0].text


def test_send_auto_without_elicitation_relies_on_client_approval(tmp_path):
    fake = FakeGraph()
    _, server = make(tmp_path, fake)  # OUTLOOK_CONFIRM_SEND=auto
    res = run(call(server, "send_draft", {"draft_id": "draft/1="}))
    assert not res.is_error and len(fake.sends()) == 1
