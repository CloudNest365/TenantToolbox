# Security Policy

## Handling sensitive data

TenantToolbox reads data from Microsoft 365 / Entra ID and generates reports that
may contain **confidential information** (users, admins, MFA status,
Conditional Access policies, app secrets, role assignments).

- Generated reports (`*.html`, `*.csv`, `*.xlsx`) and snapshots (`*.json`) are
  excluded via `.gitignore`. **Never commit real tenant output.**
- For unattended runs, use app/certificate authentication with minimal (read)
  permissions.
- All state-changing cmdlets support `-WhatIf`.

## Required Graph permissions (read)

`User.Read.All`, `Group.Read.All`, `Directory.Read.All`, `Policy.Read.All`,
`Application.Read.All`, `UserAuthenticationMethod.Read.All`,
`RoleManagement.Read.Directory`, `AuditLog.Read.All`.
(Only `Invoke-M365Offboarding` requires write permissions.)

## Reporting a vulnerability

Please **do not open a public issue** for security problems. Instead use
[GitHub Security Advisories](../../security/advisories) ("Report a vulnerability")
or contact CloudNest365 directly. We aim to respond within a few days.
