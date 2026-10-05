# App-Registrierung in Entra ID

Einmalig nötig. Dauert ca. 5 Minuten. Für die Zustimmung zu den Rechten
ist eventuell ein Admin nötig (je nach Tenant-Einstellung).

## Neue App anlegen (empfohlen)

1. <https://entra.microsoft.com> öffnen → **Anwendungen → App-Registrierungen → Neue Registrierung**.
2. Name: `lad-outlook-mcp`
3. Kontotypen: **Nur Konten in diesem Organisationsverzeichnis** (Single Tenant).
4. Umleitungs-URI: leer lassen → **Registrieren**.
5. Auf der Übersichtsseite notieren:
   - **Anwendungs-ID (Client)** → `OUTLOOK_CLIENT_ID`
   - **Verzeichnis-ID (Mandant)** → `OUTLOOK_TENANT_ID`

## Device-Code-Anmeldung erlauben

6. **Authentifizierung** → ganz unten **Öffentliche Clientflows zulassen** → **Ja** → Speichern.

## Rechte vergeben

7. **API-Berechtigungen → Berechtigung hinzufügen → Microsoft Graph → Delegierte Berechtigungen**:
   - `User.Read` (meist schon da)
   - `Mail.ReadWrite`
   - `Mail.Send`

   Für reinen Lesebetrieb (`OUTLOOK_MODE=readonly`) reicht `Mail.Read`.
8. Falls der Tenant es verlangt: **Administratorzustimmung für … erteilen**.

**Nicht** nötig und bewusst **nicht** verwenden:
- Anwendungsberechtigungen (Application permissions) – die würden Zugriff auf *alle* Postfächer geben.
- Geheime Clientschlüssel / Zertifikate.

## Vorhandene App wiederverwenden?

Auf `fm` existiert bereits eine Registrierung aus dem alten `outlook-mcp` (mit Client-Secret).
Technisch geht die Wiederverwendung (Schritte 6–8 dort ausführen). Eine eigene App ist aber
sauberer: Rechte und Anmeldungen bleiben getrennt nachvollziehbar und lassen sich einzeln
widerrufen.

## Zugriff widerrufen

- Lokal: `lad-outlook-mcp logout`
- Zentral: Entra → **Unternehmensanwendungen → lad-outlook-mcp → Benutzer und Gruppen / Berechtigungen** bzw. die App löschen.
