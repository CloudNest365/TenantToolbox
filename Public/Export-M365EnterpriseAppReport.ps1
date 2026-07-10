function Export-M365EnterpriseAppReport {
    <#
    .SYNOPSIS
        Generates an interactive HTML report of enterprise apps and their consented permissions.
    .DESCRIPTION
        Builds on Get-M365EnterpriseApp and renders a self-contained HTML page: a KPI overview
        (apps, tenant-wide consents, apps with risky scopes) and a searchable, filterable table
        with the consented scopes and risky scopes highlighted. Read-only.
    .PARAMETER Path
        Target path of the HTML file. Default: .\EnterpriseApp-Report.html
    .PARAMETER BrandName
        Branding. For CloudNest365: -BrandName 'CloudNest365'
    .PARAMETER PassThru
        Also emit the objects on the pipeline.
    .PARAMETER NoOpen
        Do not open the report automatically.
    .EXAMPLE
        Export-M365EnterpriseAppReport -BrandName 'CloudNest365'
    #>
    [CmdletBinding()]
    [OutputType([pscustomobject])]
    param(
        [string]$Path = (Join-Path (Get-Location) 'EnterpriseApp-Report.html'),
        [string]$BrandName = 'TenantToolbox',
        [string]$BrandTagline = 'M365 Tenant Administration',
        [switch]$Csv, [switch]$Excel, [string]$DataPath, [switch]$NoHtml, [switch]$PassThru, [switch]$NoOpen
    )

    Assert-TTGraph
    $data = @(Get-M365EnterpriseApp)

    $total      = $data.Count
    $tenantWide = @($data | Where-Object TenantWide).Count
    $withRisky  = @($data | Where-Object { $_.RiskyScopes.Count -gt 0 }).Count
    $genAt      = Get-Date -Format 'yyyy-MM-dd HH:mm'
    $tenant     = (Get-MgContext).TenantId

    $rows = foreach ($a in ($data | Sort-Object @{E = { -$_.RiskyScopes.Count } }, App)) {
        $fRisky = if ($a.RiskyScopes.Count -gt 0) { '1' } else { '0' }
        $fTw = if ($a.TenantWide) { '1' } else { '0' }
        $consent = if ($a.TenantWide) { "<span class='b b-warn'>Tenant-wide</span>" } else { "<span class='b b-info'>User</span>" }
        $riskyChips = if ($a.RiskyScopes.Count) { (@($a.RiskyScopes) | ForEach-Object { "<span class='chip chip-exc'>$(TTEnc $_)</span>" }) -join ' ' } else { "<span class='muted'>–</span>" }
        $riskSort = ('{0:D3}' -f [int]$a.RiskyScopes.Count)
        $searchAttr = TTEnc ("$($a.App) $(@($a.Scopes) -join ' ')".ToLower())
        $nameAttr = TTEnc ([string]$a.App).ToLower()
        @"
      <tr class="item" data-f-risky="$fRisky" data-f-tenantwide="$fTw" data-name="$nameAttr" data-s-risky="$riskSort" data-search="$searchAttr">
        <td><b>$(TTEnc $a.App)</b></td>
        <td>$consent</td>
        <td>$($a.ScopeCount)</td>
        <td data-s-risky="$riskSort"><div class="chips">$riskyChips</div></td>
      </tr>
"@
    }

    $kpiHtml = New-TTKpis @(
        @{ n = $total; l = 'Apps with grants'; filter = 'all' }
        @{ n = $withRisky; l = 'With risky scopes'; kind = 'bad'; filter = 'risky' }
        @{ n = $tenantWide; l = 'Tenant-wide consent'; kind = 'warn'; filter = 'tenantwide' }
    )
    $toolbar = New-TTToolbar -SearchPlaceholder 'Search app or scope ...' -Filters @(
        @{ label = 'All'; key = 'all' }, @{ label = 'Risky'; key = 'risky' }, @{ label = 'Tenant-wide'; key = 'tenantwide' }
    )
    $body = @"
    <div class="panel">
      <table class="tbl">
        <thead><tr><th data-sort="name">Application</th><th>Consent</th><th>Scopes</th><th data-sort="risky">Risky scopes</th></tr></thead>
        <tbody>
$($rows -join "`n")
        </tbody>
      </table>
      <div class="empty" id="empty">No app matches the filters.</div>
    </div>
"@

    $sub = "Tenant $tenant &middot; generated $genAt &middot; $total apps with delegated grants"
    $html = New-TTHtmlPage -Title 'Enterprise Apps' -Heading 'Enterprise Apps & Consent' -Sub $sub `
        -BrandName $BrandName -BrandTagline $BrandTagline -KpiHtml $kpiHtml -ToolbarHtml $toolbar -BodyHtml $body

    Write-TTLog -Level INFO -Message "Enterprise app report created: $Path ($total apps, $withRisky with risky scopes)."
    if ($withRisky -gt 0) { Write-Host "  Note: $withRisky app(s) hold risky delegated scopes." -ForegroundColor Yellow }

    $flat = $data | Select-Object App, TenantWide, ScopeCount,
        @{N = 'Scopes'; E = { $_.Scopes -join '; ' } }, @{N = 'RiskyScopes'; E = { $_.RiskyScopes -join '; ' } }, ClientId
    Complete-TTReport -Html $html -Path $Path -Data $flat -Csv:$Csv -Excel:$Excel -DataPath $DataPath -NoHtml:$NoHtml -NoOpen:$NoOpen -Kind 'Enterprise-App-Report'
    if ($PassThru) { $data }
}
