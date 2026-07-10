# Changelog

All notable changes to this project are documented here.
Format based on [Keep a Changelog](https://keepachangelog.com/en/1.0.0/),
versioning per [SemVer](https://semver.org/).

## [0.1.0] – 2026-07-10

### Added
- **Module skeleton** with a shared auth/log/`-WhatIf` frame.
- `Connect-TenantToolbox` – central sign-in (WAM / device code / browser OAuth).
- `Get-M365StaleUsers` – inactive / never-signed-in users.
- `Get-M365MfaStatus` – MFA registration status.
- `Invoke-M365Offboarding` – leaver chain (Entra + Exchange + OneDrive), with `-WhatIf`.
- `Backup-M365ConditionalAccess` – CA policies as a JSON snapshot.
- `Compare-M365Snapshot` – HTML diff of two snapshots.
- **HTML reports** (interactive: search, filters, sorting, theme, collapse):
  - `Export-M365ConditionalAccessReport`
  - `Export-M365MfaReport`
  - `Export-M365AppSecretReport`
  - `Export-M365SecurityScorecard`
  - `Export-M365PimReport`
- **Shared report engine** (`Private/Report-Html.ps1`).
- **Data export** `-Csv` / `-Excel` / `-DataPath` / `-NoHtml` for all reports.
- **Branding** via `-BrandName` / `-BrandTagline` (e.g. CloudNest365).

[0.1.0]: https://github.com/CloudNest365/TenantToolbox/releases/tag/v0.1.0
