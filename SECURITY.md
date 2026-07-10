# Security Policy

## Umgang mit sensiblen Daten

TenantToolbox liest Daten aus Microsoft 365 / Entra ID und erzeugt Reports, die
**vertrauliche Informationen** enthalten können (Benutzer, Admins, MFA-Status,
Conditional-Access-Policies, App-Secrets, Rollenzuweisungen).

- Generierte Reports (`*.html`, `*.csv`, `*.xlsx`) und Snapshots (`*.json`) sind
  in `.gitignore` ausgeschlossen. **Committe niemals echte Tenant-Ausgaben.**
- Verwende für unbeaufsichtigte Läufe App-/Zertifikats-Authentifizierung mit
  minimalen (Read-)Berechtigungen.
- Alle Cmdlets, die verändern, unterstützen `-WhatIf`.

## Benötigte Graph-Berechtigungen (Read)

`User.Read.All`, `Group.Read.All`, `Directory.Read.All`, `Policy.Read.All`,
`Application.Read.All`, `UserAuthenticationMethod.Read.All`,
`RoleManagement.Read.Directory`, `AuditLog.Read.All`.
(Nur `Invoke-M365Offboarding` benötigt Schreibrechte.)

## Sicherheitslücke melden

Bitte **kein öffentliches Issue** für Sicherheitsprobleme. Nutze stattdessen die
[GitHub Security Advisories](../../security/advisories) („Report a vulnerability")
oder kontaktiere CloudNest365 direkt. Wir bemühen uns um eine Rückmeldung
innerhalb weniger Tage.
