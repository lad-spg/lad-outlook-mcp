# App-Registrierung in Entra ID

Die vollständige Schritt-für-Schritt-Anleitung (beide Varianten: bestehende App mit
Client-Secret oder neue App ohne Secret) steht in der [README, Teil A](../README.md#3-teil-a--die-app-registrierung-bei-microsoft).

Kurzfassung der Anforderungen an die App-Registrierung:

| | Variante 1 (bestehend, mit Secret) | Variante 2 (neu, ohne Secret) |
|---|---|---|
| Plattform | Web, Redirect `http://localhost:3333/auth/callback` | keine Redirect-URI nötig |
| Öffentliche Clientflows | egal | **Ja** |
| Client-Secret | ja (`OUTLOOK_CLIENT_SECRET`) | nein |
| Delegierte Graph-Rechte | `User.Read`, `Mail.ReadWrite`, `Mail.Send` | gleich |
| Anwendungsberechtigungen | **keine** | **keine** |
