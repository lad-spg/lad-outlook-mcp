# lad-outlook-mcp

Schlanker MCP-Server für das **eigene** Outlook-Postfach (Microsoft 365) über Microsoft Graph.
Mails lesen, durchsuchen, Verläufe lesen, Anhänge auslesen, Entwürfe erstellen und senden.

Entstanden als bewusst kleine Alternative zu Community-Servern, die Postfächer lokal
spiegeln oder App-weite Rechte auf alle Postfächer im Tenant verlangen.

## Sicherheitsprinzipien

| Thema | Umsetzung |
|---|---|
| Anmeldung | Delegiert per Device-Code als **du selbst** – kein Client-Secret, kein Zugriff auf fremde Postfächer (nur `/me`). |
| Rechte | Je nach `OUTLOOK_MODE` nur so viel wie nötig: `readonly` → `Mail.Read`, `draft` → `Mail.ReadWrite`, `send` → zusätzlich `Mail.Send`. |
| Daten | Kein lokaler Mail-Speicher. Alles wird live abgefragt. Nur Anhänge, die ausdrücklich gespeichert werden, landen im Download-Ordner. |
| Senden | Immer zweistufig (Entwurf → `send_draft`). Empfänger-Whitelist (Domains/Adressen) wird vor dem Senden geprüft, auch für CC/BCC. |
| Bestätigung | Unterstützt der Client Elicitation, fragt `send_draft` den Nutzer direkt (inkl. Empfänger, Betreff, Text). |
| Anhänge senden | Nur Dateien aus dem Outbox- bzw. Download-Ordner – keine beliebigen Pfade (kein Abgreifen von `~/.ssh` o. Ä.). |
| Prompt-Injection | Mail-/Anhang-Inhalte werden als „externer Inhalt“ markiert; die Server-Instruktionen verbieten, Anweisungen daraus zu folgen. |
| Token | MSAL-Token-Cache als Datei mit Rechten `0600` (Ordner `0700`). |
| Abhängigkeiten | Exakt gepinnt (`pyproject.toml`). |

## Werkzeuge

| Tool | Modus | Zweck |
|---|---|---|
| `whoami` | alle | Angemeldetes Postfach, Modus, Whitelist |
| `list_folders` | alle | Ordnerbaum mit Anzahlen |
| `search_mails` | alle | Freitext + Absender/Empfänger/Betreff/Datum/Anhang, optional je Ordner |
| `get_mail` | alle | Mail komplett inkl. Anhangsliste |
| `get_thread` | alle | Ganzer Verlauf chronologisch |
| `list_attachments` | alle | Anhänge einer Mail |
| `read_attachment` | alle | PDF/Word/Excel/Text/HTML als Text; Rest wird gespeichert |
| `list_outbox` | draft, send | Dateien, die angehängt werden dürfen |
| `create_draft` | draft, send | Neuer Entwurf (optional mit Anhängen) |
| `create_reply_draft` | draft, send | Antwort(-an-alle)-Entwurf mit Verlauf |
| `create_forward_draft` | draft, send | Weiterleitung inkl. Original-Anhängen |
| `update_draft` | draft, send | Entwurf ändern / Anhänge ergänzen |
| `delete_draft` | draft, send | Entwurf verwerfen (nur Entwürfe) |
| `send_draft` | send | Entwurf senden (Whitelist + Bestätigung) |

## Einrichtung

### 1. App-Registrierung in Entra ID

Siehe [docs/entra-app-registrierung.md](docs/entra-app-registrierung.md).

### 2. Installation

```bash
git clone git@github.com:thelad-dev/lad-outlook-mcp.git ~/workspaces/lad-outlook-mcp
cd ~/workspaces/lad-outlook-mcp
python3 -m venv .venv
.venv/bin/pip install -e .
```

### 3. Konfiguration

```bash
mkdir -p ~/.config/lad-outlook-mcp
cp config.env.example ~/.config/lad-outlook-mcp/config.env
chmod 600 ~/.config/lad-outlook-mcp/config.env
# Client-ID, Tenant-ID und Whitelist eintragen
```

### 4. Anmelden

```bash
.venv/bin/lad-outlook-mcp login    # Code im Browser unter microsoft.com/devicelogin eingeben
.venv/bin/lad-outlook-mcp status
```

### 5. In Claude Code einbinden

```bash
claude mcp add lad-outlook -- ~/workspaces/lad-outlook-mcp/.venv/bin/lad-outlook-mcp
```

**Wichtig:** `send_draft` und `delete_draft` in Claude Code **nicht** dauerhaft freigeben
(kein „Always allow“). Dann fragt Claude Code bei jedem Senden nach – zusätzlich zur Whitelist.

## Konfiguration (`config.env` oder Umgebungsvariablen)

| Variable | Standard | Bedeutung |
|---|---|---|
| `OUTLOOK_CLIENT_ID` | – | Application (client) ID der App-Registrierung |
| `OUTLOOK_TENANT_ID` | – | Directory (tenant) ID (nicht `common`) |
| `OUTLOOK_MODE` | `send` | `readonly`, `draft` oder `send` |
| `OUTLOOK_ALLOWED_RECIPIENT_DOMAINS` | leer | Erlaubte Empfänger-Domains, kommagetrennt. Leer **und** keine Adressen = alle erlaubt |
| `OUTLOOK_ALLOWED_RECIPIENTS` | leer | Zusätzlich erlaubte Einzeladressen |
| `OUTLOOK_CONFIRM_SEND` | `auto` | `auto`: Rückfrage, wenn der Client sie kann; `always`: ohne Rückfrage-Fähigkeit kein Senden; `never` |
| `OUTLOOK_DATA_DIR` | `~/.local/share/lad-outlook-mcp` | enthält `outbox/` und `downloads/` |
| `OUTLOOK_TOKEN_CACHE` | `~/.config/lad-outlook-mcp/token_cache.json` | MSAL-Token-Cache |
| `OUTLOOK_MAX_BODY_CHARS` | `20000` | Kürzung langer Mails/Anhänge |
| `OUTLOOK_MAX_ATTACHMENT_MB` | `25` | Maximale Größe beim Lesen von Anhängen |
| `LAD_OUTLOOK_ENV_FILE` | `~/.config/lad-outlook-mcp/config.env` | Pfad der env-Datei |

Wird der Modus geändert, ist ein neues `login` nötig (andere Rechte).

## Anhänge versenden

Dateien in `~/.local/share/lad-outlook-mcp/outbox/` legen und im Tool nur den Dateinamen angeben.
Mit `read_attachment(save=true)` gespeicherte Anhänge lassen sich als `downloads/<name>` anhängen.
Bis 3 MB direkt, darüber per Upload-Session (max. 150 MB).

## Entwicklung

```bash
.venv/bin/pip install -e ".[dev]"
.venv/bin/pytest -q
```

Die Tests laufen komplett offline gegen ein gemocktes Graph und prüfen den Sendeweg über
beide MCP-Protokollvarianten (Elicitation per Rückkanal und per InputRequired-Runde).
