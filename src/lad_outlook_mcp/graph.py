"""Minimaler async Client für Microsoft Graph (nur /me – eigenes Postfach)."""

from __future__ import annotations

import asyncio
from typing import Any, Callable
from urllib.parse import quote

import httpx

GRAPH_BASE = "https://graph.microsoft.com/v1.0"
SIMPLE_ATTACHMENT_LIMIT = 3 * 1024 * 1024  # darüber Upload-Session
UPLOAD_CHUNK = 320 * 1024 * 10  # Vielfaches von 320 KiB laut Graph-Doku


class GraphError(Exception):
    def __init__(self, status: int, message: str):
        super().__init__(f"Graph-Fehler {status}: {message}")
        self.status = status


def seg(value: str) -> str:
    """ID sicher als URL-Pfadsegment kodieren."""
    return quote(value, safe="")


class Graph:
    def __init__(
        self,
        token_provider: Callable[[], str],
        transport: httpx.AsyncBaseTransport | None = None,
    ):
        self._token = token_provider
        self._client = httpx.AsyncClient(
            base_url=GRAPH_BASE, timeout=httpx.Timeout(60.0), transport=transport
        )

    async def aclose(self) -> None:
        await self._client.aclose()

    async def request(
        self,
        method: str,
        url: str,
        *,
        params: dict[str, Any] | None = None,
        json: Any = None,
        headers: dict[str, str] | None = None,
        raw: bool = False,
    ) -> Any:
        hdrs = {"Authorization": f"Bearer {self._token()}"}
        if headers:
            hdrs.update(headers)
        for attempt in range(4):
            resp = await self._client.request(method, url, params=params, json=json, headers=hdrs)
            if resp.status_code in (429, 503, 504) and attempt < 3:
                delay = min(float(resp.headers.get("Retry-After", 2 ** attempt)), 30.0)
                await asyncio.sleep(delay)
                continue
            break
        if resp.status_code >= 400:
            try:
                err = resp.json().get("error", {})
                message = f"{err.get('code', '')}: {err.get('message', resp.text)}"
            except ValueError:
                message = resp.text[:500]
            raise GraphError(resp.status_code, message)
        if raw:
            return resp.content
        if resp.status_code == 204 or not resp.content:
            return None
        return resp.json()

    async def get(self, url: str, **kw: Any) -> Any:
        return await self.request("GET", url, **kw)

    async def post(self, url: str, **kw: Any) -> Any:
        return await self.request("POST", url, **kw)

    async def patch(self, url: str, **kw: Any) -> Any:
        return await self.request("PATCH", url, **kw)

    async def delete(self, url: str, **kw: Any) -> Any:
        return await self.request("DELETE", url, **kw)

    async def get_all(self, url: str, *, params: dict[str, Any] | None = None, limit: int = 1000) -> list[dict]:
        """Folgt @odata.nextLink bis `limit` Einträge erreicht sind."""
        items: list[dict] = []
        data = await self.get(url, params=params)
        while True:
            items.extend(data.get("value", []))
            nxt = data.get("@odata.nextLink")
            if not nxt or len(items) >= limit:
                return items[:limit]
            data = await self.get(nxt)

    async def upload_large_attachment(self, message_id: str, name: str, content: bytes) -> None:
        session = await self.post(
            f"/me/messages/{seg(message_id)}/attachments/createUploadSession",
            json={"AttachmentItem": {"attachmentType": "file", "name": name, "size": len(content)}},
        )
        upload_url = session["uploadUrl"]
        # Die Upload-URL ist vorsigniert – ohne Authorization-Header senden.
        async with httpx.AsyncClient(timeout=httpx.Timeout(120.0), transport=self._client._transport) as c:
            total = len(content)
            for start in range(0, total, UPLOAD_CHUNK):
                chunk = content[start : start + UPLOAD_CHUNK]
                end = start + len(chunk) - 1
                resp = await c.put(
                    upload_url,
                    content=chunk,
                    headers={"Content-Range": f"bytes {start}-{end}/{total}", "Content-Length": str(len(chunk))},
                )
                if resp.status_code >= 400:
                    raise GraphError(resp.status_code, f"Upload von {name!r} fehlgeschlagen: {resp.text[:300]}")
