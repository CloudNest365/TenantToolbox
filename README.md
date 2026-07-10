<div align="center">

# 🧰 TenantToolbox

**Eine Galerie kleiner, scharf geschnittener PowerShell-Cmdlets für die Microsoft 365 Tenant-Administration.**
Jedes Cmdlet macht genau *einen* Job – mit einheitlichem Auth-, Log- und `-WhatIf`-Rahmen und schönen, interaktiven HTML-Reports.

[![License: CC BY-NC-ND 4.0](https://img.shields.io/badge/License-CC%20BY--NC--ND%204.0-lightgrey.svg)](LICENSE)
[![PowerShell](https://img.shields.io/badge/PowerShell-7%2B-5391FE.svg?logo=powershell&logoColor=white)](https://learn.microsoft.com/powershell/)
[![Microsoft Graph](https://img.shields.io/badge/Microsoft-Graph-0078D4.svg?logo=microsoft&logoColor=white)](https://learn.microsoft.com/graph/)
[![CI](https://github.com/CloudNest365/TenantToolbox/actions/workflows/ci.yml/badge.svg)](https://github.com/CloudNest365/TenantToolbox/actions)

*by [CloudNest365](https://github.com/CloudNest365)*

</div>

---

> ⚠️ **Hinweis:** Frei zur **nicht-kommerziellen** Nutzung mit Namensnennung – Weiterveröffentlichung/Änderung untersagt. Siehe [Lizenz](#-lizenz).
> 🔒 Reports enthalten echte Tenant-Daten. Generierte Ausgaben (`*.html/.csv/.xlsx/.json`) sind in `.gitignore` – bitte **nie committen**.

## ✨ Highlights

- **11 Cmdlets** – von inaktiven Konten über MFA & PIM bis App-Secrets.
- **Interaktive HTML-Reports** – Live-Suche, Filter, Sortierung, Light/Dark-Theme, klappbare Karten – alles self-contained (kein CDN, läuft offline).
- **Security-Scorecard** mit Gesamtnote **A–F** über MFA, Conditional Access, inaktive Konten und App-Secrets.
- **Daten-Export** als CSV/Excel und **Change-Tracking** über CA-Snapshots.
- **Branding** – jeder Report per `-BrandName` anpassbar.

## 📦 Voraussetzungen

```powershell
Install-Module Microsoft.Graph      -Scope CurrentUser   # erforderlich
Install-Module ImportExcel          -Scope CurrentUser   # optional, für -Excel
```

PowerShell 7+.

## 🚀 Loslegen

```powershell
git clone https://github.com/CloudNest365/TenantToolbox.git
Import-Module ./TenantToolbox/TenantToolbox.psd1

# Anmelden (im VS-Code-Terminal am robustesten):
Connect-TenantToolbox -UseDeviceCode -LogPath ./tenanttoolbox.log

# Harmlos, nur lesen:
Get-M365StaleUsers -InactiveDays 90 | Format-Table
Export-M365SecurityScorecard -BrandName 'CloudNest365'

# Verändernde Aktion – immer erst als Dry-Run:
Invoke-M365Offboarding -User marta@contoso.ch -WhatIf
```

## 🧩 Cmdlets

| Cmdlet | Zweck | Verändert? |
|---|---|---|
| `Connect-TenantToolbox` | Auth (WAM / Device-Code / Browser) + Log-Pfad | – |
| `Get-M365StaleUsers` | Inaktive / nie angemeldete Benutzer finden | nein |
| `Get-M365MfaStatus` | MFA-/Registrierungsstatus aller Benutzer (Objekte) | nein |
| `Invoke-M365Offboarding` | Leaver-Kette (Entra + Exchange + OneDrive) | **ja** (`-WhatIf`) |
| `Backup-M365ConditionalAccess` | CA-Policies als JSON-Snapshot sichern | nein |
| `Compare-M365Snapshot` | Zwei Snapshots vergleichen → HTML-Diff | nein (lokal) |
| `Export-M365ConditionalAccessReport` | CA-Policies mit Impact-Hinweisen | nein |
| `Export-M365MfaReport` | MFA-Status (Tabelle, Admins-ohne-MFA) | nein |
| `Export-M365AppSecretReport` | Ablaufende/abgelaufene App-Secrets & Zertifikate | nein |
| `Export-M365SecurityScorecard` | Gesamtnote A–F über mehrere Signale | nein |
| `Export-M365PimReport` | Privilegierte Rollen: permanent vs. eligible vs. aktiviert | nein |

Volle Doku im **[Wiki](https://github.com/CloudNest365/TenantToolbox/wiki)**.

## 📤 Daten-Export (CSV / Excel)

Jeder `Export-*`-Report und `Compare-M365Snapshot` kann die Rohdaten zusätzlich exportieren:

```powershell
Export-M365MfaReport -Csv              # HTML + CSV
Export-M365PimReport -Excel            # HTML + XLSX (braucht ImportExcel)
Export-M365PimReport -Excel -NoHtml    # nur XLSX
Export-M365MfaReport -NoHtml           # nur CSV (Format automatisch)
```

- `-Csv` · `-Excel` · `-DataPath <pfad>` · `-NoHtml`
- Fehlt `ImportExcel`, fällt `-Excel` automatisch auf CSV zurück.

## 🔐 Benötigte Graph-Berechtigungen

Read-only für alle Reports: `Directory.Read.All`, `Policy.Read.All`, `Application.Read.All`,
`UserAuthenticationMethod.Read.All`, `RoleManagement.Read.Directory`, `AuditLog.Read.All`.
Nur `Invoke-M365Offboarding` benötigt Schreibrechte. Details in [SECURITY.md](SECURITY.md).

## 🗺️ Roadmap

- `Remove-M365StaleGuests` – verwaiste Gäste aufräumen
- `Export-M365GuestReport` – Gäste-Übersicht
- `Get-M365TenantSummary` – farbige Konsolen-Übersicht
- Zertifikats-Auth für unbeaufsichtigte Läufe
- PIM-Check in die Scorecard

Ideen? → [Feature Request](https://github.com/CloudNest365/TenantToolbox/issues/new/choose)

## 🤝 Mitmachen

Siehe [CONTRIBUTING.md](CONTRIBUTING.md) und den [Code of Conduct](CODE_OF_CONDUCT.md).
Bitte beachte die Lizenzgrenzen (keine kommerzielle Nutzung, keine Weiterveröffentlichung).

## 📄 Lizenz

**Creative Commons BY-NC-ND 4.0** – frei zur nicht-kommerziellen Nutzung mit Namensnennung.
Keine kommerzielle Verwendung, keine Weiterveröffentlichung veränderter Versionen. Details: [LICENSE](LICENSE).

Für kommerzielle Nutzung bitte CloudNest365 kontaktieren.

---

<div align="center"><sub>© 2026 CloudNest365 · TenantToolbox</sub></div>
