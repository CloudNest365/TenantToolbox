function Export-M365AppSecretReport {
    <#
    .SYNOPSIS
        Findet App-Registrierungen mit ablaufenden/abgelaufenen Secrets und Zertifikaten.
    .DESCRIPTION
        Liest alle App-Registrierungen via Graph und listet jede Anmeldeinformation (Secret
        oder Zertifikat) mit Ablaufdatum und Restlaufzeit. Rendert einen interaktiven
        HTML-Report (Suche, Filter nach Ablauf-Status, Sortierung). Reines Lesen.
    .PARAMETER Path
        Zielpfad der HTML-Datei. Standard: .\AppSecret-Report.html
    .PARAMETER WarnDays
        Schwelle in Tagen fuer die "bald ablaufend"-Warnung. Standard: 30.
    .PARAMETER BrandName
        Branding. Fuer CloudNest365: -BrandName 'CloudNest365'
    .PARAMETER PassThru
        Gibt zusaetzlich die Objekte auf die Pipeline aus.
    .PARAMETER NoOpen
        Report nicht automatisch oeffnen.
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
    Write-TTLog -Level INFO -Message "Lese App-Registrierungen und deren Anmeldeinformationen ..."
    $apps = Get-MgApplication -All -Property 'id,appId,displayName,passwordCredentials,keyCredentials' -ErrorAction Stop

    $now = Get-Date
    $records = foreach ($a in $apps) {
        foreach ($cred in @($a.PasswordCredentials)) {
            if (-not $cred.EndDateTime) { continue }
            [pscustomobject]@{ App = $a.DisplayName; AppId = $a.AppId; Kind = 'Secret'; Name = $cred.DisplayName; End = $cred.EndDateTime; Days = [math]::Floor(([datetime]$cred.EndDateTime - $now).TotalDays) }
        }
        foreach ($cred in @($a.KeyCredentials)) {
            if (-not $cred.EndDateTime) { continue }
            [pscustomobject]@{ App = $a.DisplayName; AppId = $a.AppId; Kind = 'Zertifikat'; Name = $cred.DisplayName; End = $cred.EndDateTime; Days = [math]::Floor(([datetime]$cred.EndDateTime - $now).TotalDays) }
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
    $genAt   = Get-Date -Format 'dd.MM.yyyy HH:mm'
    $tenant  = (Get-MgContext).TenantId

    $badgeMap = @{ expired = @{c = 'b-crit'; t = 'abgelaufen' }; soon = @{c = 'b-bad'; t = "$WarnDays Tage" }; warn = @{c = 'b-warn'; t = '90 Tage' }; ok = @{c = 'b-ok'; t = 'ok' } }

    $rows = foreach ($r in $records) {
        $bucket = Get-Bucket $r.Days
        $daysTxt = if ($r.Days -lt 0) { "$([math]::Abs($r.Days)) Tage her" } else { "in $($r.Days) Tagen" }
        $badge = $badgeMap[$bucket]
        $sortDays = ($r.Days + 100000)
        $endTxt = ([datetime]$r.End).ToString('dd.MM.yyyy')
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
        @{ n = $total; l = 'Anmeldeinfos'; filter = 'all' }
        @{ n = $expired; l = 'Abgelaufen'; kind = 'bad'; filter = 'expired' }
        @{ n = $soon; l = "< $WarnDays Tage"; kind = 'bad'; filter = 'soon' }
        @{ n = $warn; l = '< 90 Tage'; kind = 'warn'; filter = 'warn' }
        @{ n = $ok; l = 'OK'; kind = 'ok'; filter = 'ok' }
    )
    $toolbar = New-TTToolbar -SearchPlaceholder 'App, Secret oder AppId suchen ...' -Filters @(
        @{ label = 'Alle'; key = 'all' }, @{ label = 'Abgelaufen'; key = 'expired' }, @{ label = "< $WarnDays Tage"; key = 'soon' },
        @{ label = '< 90 Tage'; key = 'warn' }, @{ label = 'OK'; key = 'ok' }
    )
    $body = @"
    <div class="panel">
      <table class="tbl">
        <thead><tr><th data-sort="name">App</th><th>Typ</th><th>Bezeichnung</th><th>Ablauf</th><th data-sort="days">Restlaufzeit</th></tr></thead>
        <tbody>
$($rows -join "`n")
        </tbody>
      </table>
      <div class="empty" id="empty">Keine Anmeldeinformation entspricht den Filtern.</div>
    </div>
"@

    $sub = "Tenant $tenant &middot; erstellt am $genAt &middot; $total Anmeldeinformationen"
    $html = New-TTHtmlPage -Title 'App Secret & Zertifikat Report' -Heading 'App Secret & Zertifikat Report' -Sub $sub `
        -BrandName $BrandName -BrandTagline $BrandTagline -KpiHtml $kpiHtml -ToolbarHtml $toolbar -BodyHtml $body

    Write-TTLog -Level INFO -Message "App-Secret-Report erstellt: $Path ($expired abgelaufen, $soon bald ablaufend)."
    if ($expired -gt 0) { Write-Host "  Achtung: $expired abgelaufene Anmeldeinformation(en)!" -ForegroundColor Red }
    if ($soon -gt 0) { Write-Host "  Hinweis: $soon laufen in den naechsten $WarnDays Tagen ab." -ForegroundColor Yellow }

    $flat = $records | Select-Object App, AppId, Kind, Name,
        @{N = 'End'; E = { if ($_.End) { ([datetime]$_.End).ToString('yyyy-MM-dd') } } },
        @{N = 'DaysRemaining'; E = { $_.Days } },
        @{N = 'Status'; E = { Get-Bucket $_.Days } }
    Complete-TTReport -Html $html -Path $Path -Data $flat -Csv:$Csv -Excel:$Excel -DataPath $DataPath -NoHtml:$NoHtml -NoOpen:$NoOpen -Kind 'App-Secret-Report'
    if ($PassThru) { $records }
}
