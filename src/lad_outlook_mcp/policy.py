"""Sicherheitsregeln: Empfänger-Whitelist, Dateipfade, Such-Query."""

from __future__ import annotations

import re
import unicodedata
from datetime import date
from pathlib import Path

_EMAIL_RE = re.compile(r"^[^@\s<>\"]+@[^@\s<>\"]+\.[^@\s<>\"]+$")


class PolicyError(Exception):
    pass


# ---------------------------------------------------------------- Empfänger


def normalize_address(address: str) -> str:
    addr = address.strip()
    if not _EMAIL_RE.match(addr):
        raise PolicyError(f"Ungültige E-Mail-Adresse: {address!r}")
    return addr.lower()


def recipient_violations(
    addresses: list[str], allowed_domains: frozenset[str], allowed_recipients: frozenset[str]
) -> list[str]:
    """Adressen, die nicht erlaubt sind. Sind beide Listen leer, ist alles erlaubt."""
    if not allowed_domains and not allowed_recipients:
        return []
    bad = []
    for raw in addresses:
        addr = raw.strip().lower()
        domain = addr.rpartition("@")[2]
        if addr in allowed_recipients or domain in allowed_domains:
            continue
        bad.append(raw)
    return bad


# ---------------------------------------------------------------- Dateien


def resolve_inside(base: Path, name: str) -> Path:
    """Löst `name` relativ zu `base` auf und verhindert Ausbrüche (.., Symlinks, absolute Pfade)."""
    if not name or "\x00" in name:
        raise PolicyError("Leerer oder ungültiger Dateiname.")
    base_resolved = base.resolve()
    candidate = (base_resolved / name).resolve()
    if not candidate.is_relative_to(base_resolved) or candidate == base_resolved:
        raise PolicyError(f"Datei {name!r} liegt nicht im erlaubten Ordner {base}.")
    return candidate


def safe_filename(name: str, fallback: str = "anhang") -> str:
    name = unicodedata.normalize("NFC", name or "")
    name = name.replace("/", "_").replace("\\", "_").replace("\x00", "")
    name = re.sub(r"[\x00-\x1f<>:\"|?*]", "_", name).strip(" .")
    return (name or fallback)[:180]


def unique_path(directory: Path, filename: str) -> Path:
    path = directory / filename
    stem, suffix = path.stem, path.suffix
    n = 1
    while path.exists():
        path = directory / f"{stem} ({n}){suffix}"
        n += 1
    return path


# ---------------------------------------------------------------- Suche


def _clean(value: str) -> str:
    # Anführungszeichen/Backslashes würden den $search-Ausdruck aufbrechen.
    return re.sub(r"[\"'\\()]", " ", value).strip()


def _restrict(prop: str, value: str) -> str:
    """`prop:wert` – mehrere Wörter werden einzeln per AND verknüpft (kein Phrasen-Quoting nötig)."""
    words = _clean(value).split()
    if not words:
        raise PolicyError(f"Leerer Suchwert für {prop}.")
    return " AND ".join(f"{prop}:{w}" for w in words)


def _iso_date(value: str, field: str) -> str:
    try:
        return date.fromisoformat(value.strip()).isoformat()
    except ValueError as e:
        raise PolicyError(f"{field} muss im Format JJJJ-MM-TT sein, nicht {value!r}.") from e


def build_search(
    query: str | None = None,
    sender: str | None = None,
    recipient: str | None = None,
    subject: str | None = None,
    since: str | None = None,
    until: str | None = None,
    has_attachments: bool | None = None,
) -> str | None:
    """Baut einen KQL-Ausdruck für Graph `$search` (oder None, wenn nichts gesucht wird)."""
    parts: list[str] = []
    if query and _clean(query):
        parts.append(_clean(query))
    if sender:
        parts.append(_restrict("from", sender))
    if recipient:
        parts.append(_restrict("to", recipient))
    if subject:
        parts.append(_restrict("subject", subject))
    if since:
        parts.append(f"received>={_iso_date(since, 'since')}")
    if until:
        parts.append(f"received<={_iso_date(until, 'until')}")
    if has_attachments is not None:
        parts.append(f"hasattachments:{'true' if has_attachments else 'false'}")
    if not parts:
        return None
    return '"' + " AND ".join(parts) + '"'
