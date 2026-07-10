# Contributing to TenantToolbox

Thanks for your interest! 🎉 Contributions are welcome — but please note the
[license](LICENSE) (CC BY-NC-ND 4.0): the code may **not be used commercially**
and **not be redistributed as a standalone, modified release**. Contributions
here flow into the official repository under CloudNest365.

## How to contribute

1. **Issue first** — open an issue (bug or feature) before writing larger changes
   so we avoid duplicate work.
2. **Fork & branch** — work on a feature branch (`feature/…` or `fix/…`).
3. **Pull request** — against `main`, with the PR template filled out.

## Code guidelines

- **One cmdlet = one job.** `Verb-Noun`, approved verbs (`Get-`, `Export-`,
  `Invoke-`, `Backup-`, `Compare-` …).
- **`-WhatIf`** on anything that changes state (`SupportsShouldProcess`).
- **Emit objects, not text** — pipeline-friendly.
- **Comment-based help** with `.SYNOPSIS`, `.DESCRIPTION`, `.EXAMPLE`.
- **HTML reports** use the shared engine in `Private/Report-Html.ps1`.
- **No tenant data** committed — generated reports are in `.gitignore`.

## Before the PR

```powershell
Import-Module ./TenantToolbox.psd1 -Force        # must load without errors
Invoke-ScriptAnalyzer -Path . -Recurse           # no new warnings
```

## Adding a new cmdlet

1. Create a file under `Public/Verb-Noun.ps1`.
2. Add it to `FunctionsToExport` in `TenantToolbox.psd1`.
3. Report cmdlets: finish via `New-TTHtmlPage` / `Complete-TTReport`.
4. Update the README table + CHANGELOG.

Questions? Just open an issue with the `question` label.
