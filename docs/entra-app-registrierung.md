# App-Registrierung in Entra ID

**Standard: keine App-Registrierung nötig** – das Programm nutzt dann die Microsoft-eigene App
„Microsoft Graph Command Line Tools“.

Die Anleitung für eine eigene App (bestehend mit Client-Secret oder neu ohne Secret) steht in der [README, Teil A](../README.md#3-teil-a--die-app-registrierung-bei-microsoft).

Kurzfassung der Anforderungen an die App-Registrierung:

| | Variante 1 (bestehend, mit Secret) | Variante 2 (neu, ohne Secret) |
|---|---|---|
| Plattform | Web, Redirect `http://localhost:3333/auth/callback` | Mobil/Desktop, Redirect `http://localhost` |
| Öffentliche Clientflows | egal | **Ja** |
| Client-Secret | ja (`OUTLOOK_CLIENT_SECRET`) | nein |
| Delegierte Graph-Rechte | `User.Read`, `Mail.ReadWrite`, `Mail.Send` | gleich |
| Anwendungsberechtigungen | **keine** | **keine** |
