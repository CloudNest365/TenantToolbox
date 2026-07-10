function Export-M365IntuneDeviceReport {
    <#
    .SYNOPSIS
        Generates an interactive HTML report of all Intune-managed devices.
    .DESCRIPTION
        Builds on Get-M365IntuneDevice and renders a self-contained HTML page: a KPI overview
        (compliant / non-compliant / stale / unencrypted) and a searchable, filterable, sortable
        table with compliance badges, ownership, last sync and encryption. Read-only.
    .PARAMETER Path
        Target path of the HTML file. Default: .\IntuneDevice-Report.html
    .PARAMETER StaleDays
        A device counts as stale if it has not synced for this many days. Default: 30.
    .PARAMETER BrandName
        Branding. For CloudNest365: -BrandName 'CloudNest365'
    .PARAMETER PassThru
        Also emit the device objects on the pipeline.
    .PARAMETER NoOpen
        Do not open the report automatically.
    .EXAMPLE
        Export-M365IntuneDeviceReport -BrandName 'CloudNest365'
    #>
    [CmdletBinding()]
    [OutputType([pscustomobject])]
    param(
        [string]$Path = (Join-Path (Get-Location) 'IntuneDevice-Report.html'),
        [int]$StaleDays = 30,
        [string]$BrandName = 'TenantToolbox',
        [string]$BrandTagline = 'M365 Tenant Administration',
        [switch]$Csv,
        [switch]$Excel,
        [string]$DataPath,
        [switch]$NoHtml,
        [switch]$PassThru,
        [switch]$NoOpen
    )

    Assert-TTGraph
    $data = @(Get-M365IntuneDevice -StaleDays $StaleDays)

    # Compliance display + badge class
    $compLabel = @{
        'compliant' = @{ t = 'Compliant'; c = 'b-ok' }
        'noncompliant' = @{ t = 'Non-compliant'; c = 'b-bad' }
        'error' = @{ t = 'Error'; c = 'b-bad' }
        'conflict' = @{ t = 'Conflict'; c = 'b-bad' }
        'inGracePeriod' = @{ t = 'In grace'; c = 'b-warn' }
        'configManager' = @{ t = 'ConfigMgr'; c = 'b-info' }
        'unknown' = @{ t = 'Unknown'; c = 'b-info' }
    }
    function Get-Os { param($os)
        $o = [string]$os
        if ($o -match 'windows') { 'windows' } elseif ($o -match 'ios|ipad') { 'ios' }
        elseif ($o -match 'android') { 'android' } elseif ($o -match 'mac') { 'macos' } else { 'other' }
    }

    # KPIs
    $total    = $data.Count
    $compliant = @($data | Where-Object Compliance -eq 'compliant').Count
    $noncomp  = @($data | Where-Object { $_.Compliance -in 'noncompliant', 'error', 'conflict' }).Count
    $stale    = @($data | Where-Object Stale).Count
    $unenc    = @($data | Where-Object { -not $_.Encrypted }).Count
    $genAt    = Get-Date -Format 'yyyy-MM-dd HH:mm'
    $tenant   = (Get-MgContext).TenantId

    $rows = foreach ($d in ($data | Sort-Object @{E = { $_.Compliance -eq 'compliant' } }, DeviceName)) {
        $cs = [string]$d.Compliance
        $comp = if ($compLabel.ContainsKey($cs)) { $compLabel[$cs] } else { @{ t = $cs; c = 'b-info' } }
        $osKey = Get-Os $d.OS
        $fComp = if ($cs -eq 'compliant') { '1' } else { '0' }
        $fNon  = if ($cs -in 'noncompliant', 'error', 'conflict') { '1' } else { '0' }
        $fStale = if ($d.Stale) { '1' } else { '0' }
        $fUnenc = if (-not $d.Encrypted) { '1' } else { '0' }
        $encBadge = if ($d.Encrypted) { "<span class='b b-ok'>encrypted</span>" } else { "<span class='b b-bad'>no</span>" }
        $syncTxt = if ($d.LastSync) { "$($d.LastSync.ToString('yyyy-MM-dd')) ($($d.DaysSinceSync)d)" } else { '<span class="muted">never</span>' }
        $syncSort = if ($d.LastSync) { $d.LastSync.ToString('yyyy-MM-dd') } else { '0000-00-00' }
        $searchAttr = TTEnc ("$($d.DeviceName) $($d.User) $($d.UPN) $($d.OS) $($d.Model)".ToLower())
        $nameAttr = TTEnc ([string]$d.DeviceName).ToLower()
        @"
      <tr class="item" data-f-compliant="$fComp" data-f-noncompliant="$fNon" data-f-stale="$fStale" data-f-unencrypted="$fUnenc" data-f-$osKey="1" data-name="$nameAttr" data-s-sync="$syncSort" data-search="$searchAttr">
        <td><div class="u"><b>$(TTEnc $d.DeviceName)</b><span class="upn">$(TTEnc $d.Model)</span></div></td>
        <td><div class="u"><b>$(if ($d.User) { TTEnc $d.User } else { '<span class="muted">–</span>' })</b><span class="upn">$(TTEnc $d.UPN)</span></div></td>
        <td>$(TTEnc $d.OS) $(TTEnc $d.OSVersion)</td>
        <td><span class="b $($comp.c)">$(TTEnc $comp.t)</span></td>
        <td>$(TTEnc $d.Owner)</td>
        <td data-s-sync="$syncSort">$syncTxt</td>
        <td>$encBadge</td>
      </tr>
"@
    }

    $kpiHtml = New-TTKpis @(
        @{ n = $total; l = 'Devices'; filter = 'all' }
        @{ n = $compliant; l = 'Compliant'; kind = 'ok'; filter = 'compliant' }
        @{ n = $noncomp; l = 'Non-compliant'; kind = 'bad'; filter = 'noncompliant' }
        @{ n = $stale; l = "Stale (>$StaleDays d)"; kind = 'warn'; filter = 'stale' }
        @{ n = $unenc; l = 'Unencrypted'; kind = 'bad'; filter = 'unencrypted' }
    )
    $toolbar = New-TTToolbar -SearchPlaceholder 'Search device, user or model ...' -Filters @(
        @{ label = 'All'; key = 'all' }, @{ label = 'Compliant'; key = 'compliant' }, @{ label = 'Non-compliant'; key = 'noncompliant' },
        @{ label = 'Stale'; key = 'stale' }, @{ label = 'Unencrypted'; key = 'unencrypted' },
        @{ label = 'Windows'; key = 'windows' }, @{ label = 'iOS'; key = 'ios' }, @{ label = 'Android'; key = 'android' }, @{ label = 'macOS'; key = 'macos' }
    )
    $body = @"
    <div class="panel">
      <table class="tbl">
        <thead><tr><th data-sort="name">Device</th><th>User</th><th>OS</th><th>Compliance</th><th>Owner</th><th data-sort="sync">Last sync</th><th>Encrypted</th></tr></thead>
        <tbody>
$($rows -join "`n")
        </tbody>
      </table>
      <div class="empty" id="empty">No device matches the filters.</div>
    </div>
"@

    $sub = "Tenant $tenant &middot; generated $genAt &middot; $total devices"
    $html = New-TTHtmlPage -Title 'Intune Device Report' -Heading 'Intune Device Report' -Sub $sub `
        -BrandName $BrandName -BrandTagline $BrandTagline -KpiHtml $kpiHtml -ToolbarHtml $toolbar -BodyHtml $body

    Write-TTLog -Level INFO -Message "Intune device report created: $Path ($total devices, $noncomp non-compliant, $stale stale)."
    if ($noncomp -gt 0) { Write-Host "  Note: $noncomp non-compliant device(s)." -ForegroundColor Yellow }

    $flat = $data | Select-Object DeviceName, User, UPN, OS, OSVersion, Compliance, Owner,
        @{N = 'LastSync'; E = { if ($_.LastSync) { $_.LastSync.ToString('yyyy-MM-dd') } } }, DaysSinceSync, Stale, Encrypted, Model, Manufacturer
    Complete-TTReport -Html $html -Path $Path -Data $flat -Csv:$Csv -Excel:$Excel -DataPath $DataPath -NoHtml:$NoHtml -NoOpen:$NoOpen -Kind 'Intune-Device-Report'
    if ($PassThru) { $data }
}
