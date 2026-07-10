# Changelog

All notable changes to this project are documented here.
Format based on [Keep a Changelog](https://keepachangelog.com/en/1.0.0/),
versioning per [SemVer](https://semver.org/).

## [Unreleased]

### Added — multi-tenant (MSP)
- `Set-M365MultiTenant` / `Get-M365MultiTenant` / `Invoke-M365MultiTenant` – define a set of tenants and run read cmdlets across all of them; results are tagged with a `Tenant` column. App-only (certificate) or interactive (device code) per tenant.
- `Connect-TenantToolbox` now accepts `-TenantId` for interactive sign-in too (target a specific tenant).

## [0.5.1] – 2026-07-10

### Changed
- Author/company/copyright set to CloudNest365 (removed personal name); published to the PowerShell Gallery.

## [0.5.0] – 2026-07-10

### Added — new reports
- `Export-M365SignInReport` – recent sign-ins flagged failed / legacy-auth / risky.
- `Export-M365AutopilotReport` – Windows Autopilot devices and profile assignment.
- `Export-M365DynamicGroupReport` – dynamic groups and their membership rules.
- `Export-M365BitLockerReport` – BitLocker recovery-key escrow coverage.
- `Export-M365MailboxSizeReport` – mailbox sizes and item counts (Exchange).
- `Export-M365SharingReport` – SharePoint external sharing posture (SPO).

### Added — remediation (with `-WhatIf`)
- `Disable-M365ExternalForwarding` – disable external inbox-forwarding rules.
- `Revoke-M365AppConsent` – revoke an app's delegated OAuth grants.

### Changed — robustness
- `Get-TTGraphCollection` now retries on throttling (429) and transient errors with backoff, and shows progress (new `Invoke-TTGraph` wrapper).
- New default scopes: `BitlockerKey.Read.All`, `DeviceManagementServiceConfig.Read.All`, `MailboxSettings.ReadWrite`, `DelegatedPermissionGrant.ReadWrite.All`.

## [0.4.0] – 2026-07-10

### Added — small admin reports
- `Export-M365TeamsReport` – Microsoft Teams with owner/member/guest counts and visibility.
- `Export-M365DomainReport` – domains (verified, default, auth type).
- `Export-M365ServiceHealthReport` – current M365 service issues/incidents.
- `Export-M365MessageCenterReport` – Message Center announcements & action-required items.
- `Export-M365PasswordPolicyReport` – password findings (never expires, weak, old).
- `Export-M365RegisteredDeviceReport` – Entra-registered/joined devices.
- `Export-M365DistributionListReport` – distribution lists (with empty-list flag).
- `Export-M365SharedMailboxReport` – shared mailboxes & Full Access delegates (Exchange Online).
- New default scopes: `ServiceHealth.Read.All`, `ServiceMessage.Read.All`.

## [0.3.0] – 2026-07-10

### Added — security detection
- `Get-M365MailForwarding` + `Export-M365MailForwardingReport` – inbox rules forwarding mail externally (account-compromise signal).
- `Get-M365EnterpriseApp` + `Export-M365EnterpriseAppReport` – OAuth consents with risky delegated scopes highlighted.
- `Get-M365RiskyUser` + `Export-M365RiskyUsersReport` – Entra Identity Protection risky users (P2).
- New default scopes: `MailboxSettings.Read`, `IdentityRiskyUser.Read.All`.

### Added — identity & governance
- `Get-M365AdminRole` + `Export-M365AdminRoleReport` – all admin role holders cross-referenced with MFA.
- `Get-M365Guest` + `Export-M365GuestReport` – guest accounts with domain, state and last sign-in.
- `Get-M365License` + `Export-M365LicenseReport` – license (SKU) assignment: total / assigned / available.
- `Get-M365Group` + `Export-M365GroupReport` – groups & teams with type, owners and orphan flag.

## [0.2.0] – 2026-07-10

### Added — Intune suite
- `Get-M365IntuneDevice` + `Export-M365IntuneDeviceReport` – managed devices with compliance, sync and encryption status (#6).
- `Get-M365IntuneApp` + `Export-M365IntuneAppReport` – software inventory (detected apps) with version and device count (#7).
- `Get-M365IntuneAppDeployment` + `Export-M365IntuneAppDeploymentReport` – app assignment and install status (installed/failed/not installed) (#8).
- `Get-M365IntuneDeviceApp` + `Export-M365IntuneDeviceAppReport` – per-device app drilldown (#10).
- `Remove-M365StaleDevices` – find and delete stale Intune devices, with `-WhatIf` (#9).

### Added — admin & security
- `Get-M365TenantSummary` – colored console overview of the tenant's security posture (#2).
- `Remove-M365StaleGuests` – find and remove inactive guest accounts, with `-WhatIf` (#1).
- `Connect-TenantToolbox`: certificate/app-only auth (`-ClientId` / `-TenantId` / `-CertificateThumbprint`) for unattended runs (#4).
- `Export-M365SecurityScorecard`: new "Permanent Global Admins" check (PIM) (#3).
- Report screenshots in README and Wiki (#5).

### Changed
- Repository is now fully English (docs, code, report UI, wiki, commits).
- `Get-M365IntuneDevice` now returns `DeviceId`.
- New default scopes: `DeviceManagementManagedDevices.ReadWrite.All`, `DeviceManagementApps.Read.All`.
- Graceful handling of missing-scope (403) errors.
- Added `.gitattributes` for consistent LF line endings.

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

[0.5.0]: https://github.com/CloudNest365/TenantToolbox/releases/tag/v0.5.0
[0.4.0]: https://github.com/CloudNest365/TenantToolbox/releases/tag/v0.4.0
[0.3.0]: https://github.com/CloudNest365/TenantToolbox/releases/tag/v0.3.0
[0.2.0]: https://github.com/CloudNest365/TenantToolbox/releases/tag/v0.2.0
[0.1.0]: https://github.com/CloudNest365/TenantToolbox/releases/tag/v0.1.0
