function Backup-M365ConditionalAccess {
    <#
    .SYNOPSIS
        Sichert alle Conditional-Access-Policies als JSON-Snapshot (fuer Versionierung/Vergleich).
    .DESCRIPTION
        Exportiert die relevanten Felder aller CA-Policies in eine JSON-Datei. Diese Snapshots
        lassen sich mit Compare-M365Snapshot vergleichen ("was hat sich geaendert?"). Reines Lesen.
    .PARAMETER Path
        Zielpfad der JSON-Datei. Standard: .\CA-Snapshot-<Zeitstempel>.json
    .PARAMETER PassThru
        Gibt das Snapshot-Objekt zusaetzlich auf die Pipeline aus.
    .EXAMPLE
        Backup-M365ConditionalAccess -Path .\snapshots\ca-2026-07-10.json
    #>
    [CmdletBinding()]
    param(
        [string]$Path = (Join-Path (Get-Location) ("CA-Snapshot-{0}.json" -f (Get-Date -Format 'yyyyMMdd-HHmmss'))),
        [switch]$PassThru
    )

    Assert-TTGraph
    Write-TTLog -Level INFO -Message "Sichere Conditional-Access-Policies ..."

    $pol = Get-MgIdentityConditionalAccessPolicy -All -ErrorAction Stop
    $snap = foreach ($p in $pol) {
        [pscustomobject]@{
            Id              = $p.Id
            DisplayName     = $p.DisplayName
            State           = $p.State
            Conditions      = $p.Conditions
            GrantControls   = $p.GrantControls
            SessionControls = $p.SessionControls
            ModifiedDateTime = $p.ModifiedDateTime
        }
    }

    $dir = Split-Path -Parent $Path
    if ($dir -and -not (Test-Path $dir)) { New-Item -ItemType Directory -Path $dir -Force | Out-Null }
    $snap | ConvertTo-Json -Depth 25 | Out-File -FilePath $Path -Encoding utf8

    Write-TTLog -Level INFO -Message "CA-Snapshot gesichert: $Path ($($snap.Count) Policies)."
    Write-Host "CA-Snapshot gesichert: $Path ($($snap.Count) Policies)" -ForegroundColor Green
    if ($PassThru) { $snap }
}
