function Export-M365MailForwardingReport {
    <#
    .SYNOPSIS
        Generates an interactive HTML report of inbox rules that forward mail externally.
    .DESCRIPTION
        Builds on Get-M365MailForwarding and renders a self-contained HTML page: a KPI overview
        (rules, affected users, enabled rules) and a searchable table with rule, action, external
        recipients and enabled state. Read-only. Scanning every mailbox can be slow on large tenants.
    .PARAMETER Path
        Target path of the HTML file. Default: .\MailForwarding-Report.html
    .PARAMETER BrandName
        Branding. For CloudNest365: -BrandName 'CloudNest365'
    .PARAMETER PassThru
        Also emit the objects on the pipeline.
    .PARAMETER NoOpen
        Do not open the report automatically.
    .EXAMPLE
        Export-M365MailForwardingReport -BrandName 'CloudNest365'
    #>
    [CmdletBinding()]
    [OutputType([pscustomobject])]
    param(
        [string]$Path = (Join-Path (Get-Location) 'MailForwarding-Report.html'),
        [string]$BrandName = 'TenantToolbox',
        [string]$BrandTagline = 'M365 Tenant Administration',
        [switch]$Csv, [switch]$Excel, [string]$DataPath, [switch]$NoHtml, [switch]$PassThru, [switch]$NoOpen
    )

    Assert-TTGraph
    Write-Host "Scanning mailboxes for external forwarding rules (may take a while) ..." -ForegroundColor DarkGray
    $data = @(Get-M365MailForwarding)

    $total   = $data.Count
    $users   = @($data | Select-Object -ExpandProperty UPN -Unique).Count
    $enabled = @($data | Where-Object Enabled).Count
    $genAt   = Get-Date -Format 'yyyy-MM-dd HH:mm'
    $tenant  = (Get-MgContext).TenantId

    $rows = foreach ($r in ($data | Sort-Object @{E = { -[int][bool]$_.Enabled } }, User)) {
        $fEnabled = if ($r.Enabled) { '1' } else { '0' }
        $enBadge = if ($r.Enabled) { "<span class='b b-bad'>enabled</span>" } else { "<span class='muted'>disabled</span>" }
        $extChips = (@($r.External) | ForEach-Object { "<span class='chip chip-exc'>$(TTEnc $_)</span>" }) -join ' '
        $searchAttr = TTEnc ("$($r.User) $($r.UPN) $($r.Rule) $(@($r.External) -join ' ')".ToLower())
        $nameAttr = TTEnc ([string]$r.User).ToLower()
        @"
      <tr class="item" data-f-enabled="$fEnabled" data-name="$nameAttr" data-search="$searchAttr">
        <td><div class="u"><b>$(TTEnc $r.User)</b><span class="upn">$(TTEnc $r.UPN)</span></div></td>
        <td>$(TTEnc $r.Rule)</td>
        <td>$(TTEnc $r.Action)</td>
        <td><div class="chips">$extChips</div></td>
        <td>$enBadge</td>
      </tr>
"@
    }

    $kpiHtml = New-TTKpis @(
        @{ n = $total; l = 'Forwarding rules'; filter = 'all' }
        @{ n = $users; l = 'Affected users'; kind = 'bad' }
        @{ n = $enabled; l = 'Enabled rules'; kind = 'bad'; filter = 'enabled' }
    )
    $toolbar = New-TTToolbar -SearchPlaceholder 'Search user, rule or recipient ...' -Filters @(
        @{ label = 'All'; key = 'all' }, @{ label = 'Enabled'; key = 'enabled' }
    )
    $body = @"
    <div class="panel">
      <table class="tbl">
        <thead><tr><th data-sort="name">User</th><th>Rule</th><th>Action</th><th>External recipients</th><th>State</th></tr></thead>
        <tbody>
$($rows -join "`n")
        </tbody>
      </table>
      <div class="empty" id="empty">No forwarding rule matches the filters.</div>
    </div>
"@

    $sub = "Tenant $tenant &middot; generated $genAt &middot; $total external forwarding rules on $users users"
    $html = New-TTHtmlPage -Title 'External Mail Forwarding' -Heading 'External Mail Forwarding' -Sub $sub `
        -BrandName $BrandName -BrandTagline $BrandTagline -KpiHtml $kpiHtml -ToolbarHtml $toolbar -BodyHtml $body

    Write-TTLog -Level INFO -Message "Mail forwarding report created: $Path ($total rules, $users users, $enabled enabled)."
    if ($enabled -gt 0) { Write-Host "  Warning: $enabled enabled external-forwarding rule(s) - review for compromise!" -ForegroundColor Red }

    $flat = $data | Select-Object User, UPN, Rule, Action, @{N = 'External'; E = { $_.External -join '; ' } }, Enabled
    Complete-TTReport -Html $html -Path $Path -Data $flat -Csv:$Csv -Excel:$Excel -DataPath $DataPath -NoHtml:$NoHtml -NoOpen:$NoOpen -Kind 'Mail-Forwarding-Report'
    if ($PassThru) { $data }
}
