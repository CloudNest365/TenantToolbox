# Contributing to TenantToolbox

Danke für dein Interesse! 🎉 Beiträge sind willkommen – bitte beachte aber die
[Lizenz](LICENSE) (CC BY-NC-ND 4.0): der Code darf **nicht kommerziell** und
**nicht als eigenständige, veränderte Veröffentlichung** weitergegeben werden.
Beiträge hierher fliessen ins offizielle Repo unter CloudNest365 ein.

## Wie du beiträgst

1. **Issue zuerst** – eröffne ein Issue (Bug oder Feature), bevor du grösseren
   Code schreibst, damit wir Doppelarbeit vermeiden.
2. **Fork & Branch** – arbeite auf einem Feature-Branch (`feature/…` oder `fix/…`).
3. **Pull Request** – gegen `main`, mit ausgefülltem PR-Template.

## Code-Richtlinien

- **Ein Cmdlet = ein Job.** `Verb-Noun`, approved Verbs (`Get-`, `Export-`,
  `Invoke-`, `Backup-`, `Compare-` …).
- **`-WhatIf`** bei allem, was verändert (`SupportsShouldProcess`).
- **Objekte ausgeben, nicht Text** – weiterverarbeitbar via Pipeline.
- **Comment-Based Help** mit `.SYNOPSIS`, `.DESCRIPTION`, `.EXAMPLE`.
- **HTML-Reports** nutzen die gemeinsame Engine in `Private/Report-Html.ps1`.
- **Keine Tenant-Daten** committen – generierte Reports sind in `.gitignore`.

## Vor dem PR

```powershell
Import-Module ./TenantToolbox.psd1 -Force        # muss fehlerfrei laden
Invoke-ScriptAnalyzer -Path . -Recurse           # keine neuen Warnungen
```

## Neues Cmdlet hinzufügen

1. Datei unter `Public/Verb-Noun.ps1` anlegen.
2. In `TenantToolbox.psd1` unter `FunctionsToExport` eintragen.
3. Report-Cmdlets: über `New-TTHtmlPage` / `Complete-TTReport` abschliessen.
4. README-Tabelle + CHANGELOG ergänzen.

Fragen? Eröffne einfach ein Issue mit dem Label `question`.
