function Export-M365AppSecretReport {
    <#
    .SYNOPSIS
        Finds app registrations with expiring/expired secrets and certificates.
    .DESCRIPTION
        Reads all app registrations via Graph and lists every credential (secret or
        certificate) with expiry date and remaining lifetime. Renders an interactive
        HTML report (search, filter by expiry status, sorting). Read-only.
    .PARAMETER Path
        Target path of the HTML file. Default: .\AppSecret-Report.html
    .PARAMETER WarnDays
        Threshold in days for the "expiring soon" warning. Default: 30.
    .PARAMETER BrandName
        Branding. For CloudNest365: -BrandName 'CloudNest365'
    .PARAMETER PassThru
        Also emit the objects on the pipeline.
    .PARAMETER NoOpen
        Do not open the report automatically.
    .EXAMPLE
        Export-M365AppSecretReport -Path .\secrets.html -BrandName 'CloudNest365'
    #>
    [CmdletBinding()]
    [OutputType([pscustomobject])]
    param(
        [string]$Path = (Join-Path (Get-Location) 'AppSecret-Report.html'),
        [int]$WarnDays = 30,
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
    Write-TTLog -Level INFO -Message "Reading app registrations and their credentials ..."
    $apps = Get-MgApplication -All -Property 'id,appId,displayName,passwordCredentials,keyCredentials' -ErrorAction Stop

    $now = Get-Date
    $records = foreach ($a in $apps) {
        foreach ($cred in @($a.PasswordCredentials)) {
            if (-not $cred.EndDateTime) { continue }
            [pscustomobject]@{ App = $a.DisplayName; AppId = $a.AppId; Kind = 'Secret'; Name = $cred.DisplayName; End = $cred.EndDateTime; Days = [math]::Floor(([datetime]$cred.EndDateTime - $now).TotalDays) }
        }
        foreach ($cred in @($a.KeyCredentials)) {
            if (-not $cred.EndDateTime) { continue }
            [pscustomobject]@{ App = $a.DisplayName; AppId = $a.AppId; Kind = 'Certificate'; Name = $cred.DisplayName; End = $cred.EndDateTime; Days = [math]::Floor(([datetime]$cred.EndDateTime - $now).TotalDays) }
        }
    }
    $records = @($records | Sort-Object Days)

    # Buckets
    function Get-Bucket { param($d)
        if ($d -lt 0) { 'expired' } elseif ($d -le $WarnDays) { 'soon' } elseif ($d -le 90) { 'warn' } else { 'ok' }
    }
    $total   = $records.Count
    $expired = @($records | Where-Object { $_.Days -lt 0 }).Count
    $soon    = @($records | Where-Object { $_.Days -ge 0 -and $_.Days -le $WarnDays }).Count
    $warn    = @($records | Where-Object { $_.Days -gt $WarnDays -and $_.Days -le 90 }).Count
    $ok      = @($records | Where-Object { $_.Days -gt 90 }).Count
    $genAt   = Get-Date -Format 'yyyy-MM-dd HH:mm'
    $tenant  = (Get-MgContext).TenantId

    $badgeMap = @{ expired = @{c = 'b-crit'; t = 'expired' }; soon = @{c = 'b-bad'; t = "$WarnDays days" }; warn = @{c = 'b-warn'; t = '90 days' }; ok = @{c = 'b-ok'; t = 'ok' } }

    $rows = foreach ($r in $records) {
        $bucket = Get-Bucket $r.Days
        $daysTxt = if ($r.Days -lt 0) { "$([math]::Abs($r.Days)) days ago" } else { "in $($r.Days) days" }
        $badge = $badgeMap[$bucket]
        $sortDays = ($r.Days + 100000)
        $endTxt = ([datetime]$r.End).ToString('yyyy-MM-dd')
        $searchAttr = TTEnc ("$($r.App) $($r.Name) $($r.Kind) $($r.AppId)".ToLower())
        $nameAttr = TTEnc ([string]$r.App).ToLower()
        @"
      <tr class="item" data-f-$bucket="1" data-name="$nameAttr" data-s-days="$sortDays" data-search="$searchAttr">
        <td><div class="u"><b>$(TTEnc $r.App)</b><span class="upn">$(TTEnc $r.AppId)</span></div></td>
        <td>$(TTEnc $r.Kind)</td>
        <td>$(if ($r.Name) { TTEnc $r.Name } else { '<span class="muted">–</span>' })</td>
        <td>$endTxt</td>
        <td><span class="b $($badge.c)">$daysTxt</span></td>
      </tr>
"@
    }

    $kpiHtml = New-TTKpis @(
        @{ n = $total; l = 'Credentials'; filter = 'all' }
        @{ n = $expired; l = 'Expired'; kind = 'bad'; filter = 'expired' }
        @{ n = $soon; l = "< $WarnDays days"; kind = 'bad'; filter = 'soon' }
        @{ n = $warn; l = '< 90 days'; kind = 'warn'; filter = 'warn' }
        @{ n = $ok; l = 'OK'; kind = 'ok'; filter = 'ok' }
    )
    $toolbar = New-TTToolbar -SearchPlaceholder 'Search app, secret or AppId ...' -Filters @(
        @{ label = 'All'; key = 'all' }, @{ label = 'Expired'; key = 'expired' }, @{ label = "< $WarnDays days"; key = 'soon' },
        @{ label = '< 90 days'; key = 'warn' }, @{ label = 'OK'; key = 'ok' }
    )
    $body = @"
    <div class="panel">
      <table class="tbl">
        <thead><tr><th data-sort="name">App</th><th>Type</th><th>Name</th><th>Expiry</th><th data-sort="days">Remaining</th></tr></thead>
        <tbody>
$($rows -join "`n")
        </tbody>
      </table>
      <div class="empty" id="empty">No credential matches the filters.</div>
    </div>
"@

    $sub = "Tenant $tenant &middot; generated $genAt &middot; $total credentials"
    $html = New-TTHtmlPage -Title 'App Secret & Certificate Report' -Heading 'App Secret & Certificate Report' -Sub $sub `
        -BrandName $BrandName -BrandTagline $BrandTagline -KpiHtml $kpiHtml -ToolbarHtml $toolbar -BodyHtml $body

    Write-TTLog -Level INFO -Message "App secret report created: $Path ($expired expired, $soon expiring soon)."
    if ($expired -gt 0) { Write-Host "  Warning: $expired expired credential(s)!" -ForegroundColor Red }
    if ($soon -gt 0) { Write-Host "  Note: $soon expire within the next $WarnDays days." -ForegroundColor Yellow }

    $flat = $records | Select-Object App, AppId, Kind, Name,
        @{N = 'End'; E = { if ($_.End) { ([datetime]$_.End).ToString('yyyy-MM-dd') } } },
        @{N = 'DaysRemaining'; E = { $_.Days } },
        @{N = 'Status'; E = { Get-Bucket $_.Days } }
    Complete-TTReport -Html $html -Path $Path -Data $flat -Csv:$Csv -Excel:$Excel -DataPath $DataPath -NoHtml:$NoHtml -NoOpen:$NoOpen -Kind 'App-Secret-Report'
    if ($PassThru) { $records }
}
