# Changelog

All notable changes to this project are documented here.
Format based on [Keep a Changelog](https://keepachangelog.com/en/1.0.0/),
versioning per [SemVer](https://semver.org/).

## [Unreleased]

### Added
- `Get-M365IntuneDevice` + `Export-M365IntuneDeviceReport` – Intune managed devices with compliance, sync and encryption status (#6).
- `Get-M365IntuneApp` + `Export-M365IntuneAppReport` – software inventory (detected apps) with version and device count.
- `Get-M365IntuneAppDeployment` + `Export-M365IntuneAppDeploymentReport` – app assignment and install status (installed/failed/not installed).
- `Get-M365IntuneDeviceApp` + `Export-M365IntuneDeviceAppReport` – per-device app drilldown.
- `Remove-M365StaleDevices` – find and delete stale Intune devices, with `-WhatIf`.
- `Get-M365IntuneDevice` now returns `DeviceId`.
- `Get-M365TenantSummary` – colored console overview of the tenant's security posture (#2).
- `Remove-M365StaleGuests` – find and remove inactive guest accounts, with `-WhatIf` (#1).
- `Connect-TenantToolbox`: certificate/app-only auth (`-ClientId` / `-TenantId` / `-CertificateThumbprint`) for unattended runs (#4).
- `Export-M365SecurityScorecard`: new "Permanent Global Admins" check (PIM) (#3).
- `.gitattributes` for consistent LF line endings.

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
