function Backup-M365ConditionalAccess {
    <#
    .SYNOPSIS
        Backs up all Conditional Access policies as a JSON snapshot (for versioning/comparison).
    .DESCRIPTION
        Exports the relevant fields of all CA policies to a JSON file. These snapshots can be
        compared with Compare-M365Snapshot ("what changed?"). Read-only.
    .PARAMETER Path
        Target path of the JSON file. Default: .\CA-Snapshot-<timestamp>.json
    .PARAMETER PassThru
        Also emit the snapshot object on the pipeline.
    .EXAMPLE
        Backup-M365ConditionalAccess -Path .\snapshots\ca-2026-07-10.json
    #>
    [CmdletBinding()]
    param(
        [string]$Path = (Join-Path (Get-Location) ("CA-Snapshot-{0}.json" -f (Get-Date -Format 'yyyyMMdd-HHmmss'))),
        [switch]$PassThru
    )

    Assert-TTGraph
    Write-TTLog -Level INFO -Message "Backing up Conditional Access policies ..."

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

    Write-TTLog -Level INFO -Message "CA snapshot saved: $Path ($($snap.Count) policies)."
    Write-Host "CA snapshot saved: $Path ($($snap.Count) policies)" -ForegroundColor Green
    if ($PassThru) { $snap }
}
