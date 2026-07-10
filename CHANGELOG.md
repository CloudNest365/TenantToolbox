# Changelog

Alle nennenswerten Änderungen an diesem Projekt werden hier dokumentiert.
Format orientiert sich an [Keep a Changelog](https://keepachangelog.com/de/1.0.0/),
Versionierung nach [SemVer](https://semver.org/lang/de/).

## [0.1.0] – 2026-07-10

### Added
- **Modul-Grundgerüst** mit gemeinsamem Auth-/Log-/`-WhatIf`-Rahmen.
- `Connect-TenantToolbox` – zentraler Login (WAM / Device-Code / Browser-OAuth).
- `Get-M365StaleUsers` – inaktive / nie angemeldete Benutzer.
- `Get-M365MfaStatus` – MFA-Registrierungsstatus.
- `Invoke-M365Offboarding` – Leaver-Kette (Entra + Exchange + OneDrive), mit `-WhatIf`.
- `Backup-M365ConditionalAccess` – CA-Policies als JSON-Snapshot.
- `Compare-M365Snapshot` – HTML-Diff zweier Snapshots.
- **HTML-Reports** (interaktiv: Suche, Filter, Sortierung, Theme, Klappen):
  - `Export-M365ConditionalAccessReport`
  - `Export-M365MfaReport`
  - `Export-M365AppSecretReport`
  - `Export-M365SecurityScorecard`
  - `Export-M365PimReport`
- **Gemeinsame Report-Engine** (`Private/Report-Html.ps1`).
- **Daten-Export** `-Csv` / `-Excel` / `-DataPath` / `-NoHtml` für alle Reports.
- **Branding** über `-BrandName` / `-BrandTagline` (z. B. CloudNest365).

[0.1.0]: https://github.com/CloudNest365/TenantToolbox/releases/tag/v0.1.0
