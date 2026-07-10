function Export-M365RiskyUsersReport {
    <#
    .SYNOPSIS
        Generates an interactive HTML report of Entra Identity Protection risky users.
    .DESCRIPTION
        Builds on Get-M365RiskyUser and renders a self-contained HTML page: a KPI overview and a
        searchable, filterable, sortable table with risk level and state per user. Read-only.
        Requires Entra ID P2 and the IdentityRiskyUser.Read.All scope.
    .PARAMETER Path
        Target path of the HTML file. Default: .\RiskyUsers-Report.html
    .PARAMETER BrandName
        Branding. For CloudNest365: -BrandName 'CloudNest365'
    .PARAMETER PassThru
        Also emit the objects on the pipeline.
    .PARAMETER NoOpen
        Do not open the report automatically.
    .EXAMPLE
        Export-M365RiskyUsersReport -BrandName 'CloudNest365'
    #>
    [CmdletBinding()]
    [OutputType([pscustomobject])]
    param(
        [string]$Path = (Join-Path (Get-Location) 'RiskyUsers-Report.html'),
        [string]$BrandName = 'TenantToolbox',
        [string]$BrandTagline = 'M365 Tenant Administration',
        [switch]$Csv, [switch]$Excel, [string]$DataPath, [switch]$NoHtml, [switch]$PassThru, [switch]$NoOpen
    )

    Assert-TTGraph
    $data = @(Get-M365RiskyUser)

    $levelBadge = @{ high = 'b-bad'; medium = 'b-warn'; low = 'b-info' }
    $total  = $data.Count
    $high   = @($data | Where-Object RiskLevel -eq 'high').Count
    $comp   = @($data | Where-Object RiskState -eq 'confirmedCompromised').Count
    $atRisk = @($data | Where-Object RiskState -eq 'atRisk').Count
    $genAt  = Get-Date -Format 'yyyy-MM-dd HH:mm'
    $tenant = (Get-MgContext).TenantId

    $order = @{ high = 0; medium = 1; low = 2 }
    $rows = foreach ($u in ($data | Sort-Object @{E = { $order[[string]$_.RiskLevel] } }, User)) {
        $lvl = [string]$u.RiskLevel
        $bc = if ($levelBadge.ContainsKey($lvl)) { $levelBadge[$lvl] } else { 'b-info' }
        $fHigh = if ($lvl -eq 'high') { '1' } else { '0' }
        $fMed = if ($lvl -eq 'medium') { '1' } else { '0' }
        $fAt = if ($u.RiskState -eq 'atRisk') { '1' } else { '0' }
        $fComp = if ($u.RiskState -eq 'confirmedCompromised') { '1' } else { '0' }
        $upd = if ($u.LastUpdated) { ([datetime]$u.LastUpdated).ToString('yyyy-MM-dd') } else { '–' }
        $searchAttr = TTEnc ("$($u.User) $($u.UPN) $lvl $($u.RiskState)".ToLower())
        $nameAttr = TTEnc ([string]$u.User).ToLower()
        @"
      <tr class="item" data-f-high="$fHigh" data-f-medium="$fMed" data-f-atrisk="$fAt" data-f-compromised="$fComp" data-name="$nameAttr" data-search="$searchAttr">
        <td><div class="u"><b>$(TTEnc $u.User)</b><span class="upn">$(TTEnc $u.UPN)</span></div></td>
        <td><span class="b $bc">$(TTEnc $lvl)</span></td>
        <td>$(TTEnc $u.RiskState)</td>
        <td>$upd</td>
      </tr>
"@
    }

    $kpiHtml = New-TTKpis @(
        @{ n = $total; l = 'Risky users'; filter = 'all' }
        @{ n = $high; l = 'High risk'; kind = 'bad'; filter = 'high' }
        @{ n = $comp; l = 'Compromised'; kind = 'bad'; filter = 'compromised' }
        @{ n = $atRisk; l = 'At risk'; kind = 'warn'; filter = 'atrisk' }
    )
    $toolbar = New-TTToolbar -SearchPlaceholder 'Search user ...' -Filters @(
        @{ label = 'All'; key = 'all' }, @{ label = 'High'; key = 'high' }, @{ label = 'Medium'; key = 'medium' },
        @{ label = 'At risk'; key = 'atrisk' }, @{ label = 'Compromised'; key = 'compromised' }
    )
    $note = if ($total -eq 0) { '<p class="muted" style="padding:0 0 12px">No risky users returned. This report requires Entra ID P2 and the IdentityRiskyUser.Read.All scope.</p>' } else { '' }
    $body = @"
    $note
    <div class="panel">
      <table class="tbl">
        <thead><tr><th data-sort="name">User</th><th>Risk level</th><th>State</th><th>Last updated</th></tr></thead>
        <tbody>
$($rows -join "`n")
        </tbody>
      </table>
      <div class="empty" id="empty">No user matches the filters.</div>
    </div>
"@

    $sub = "Tenant $tenant &middot; generated $genAt &middot; $total risky users"
    $html = New-TTHtmlPage -Title 'Risky Users' -Heading 'Risky Users (Identity Protection)' -Sub $sub `
        -BrandName $BrandName -BrandTagline $BrandTagline -KpiHtml $kpiHtml -ToolbarHtml $toolbar -BodyHtml $body

    Write-TTLog -Level INFO -Message "Risky users report created: $Path ($total risky, $high high, $comp compromised)."
    if ($comp -gt 0) { Write-Host "  Warning: $comp confirmed compromised user(s)!" -ForegroundColor Red }

    $flat = $data | Select-Object User, UPN, RiskLevel, RiskState, LastUpdated
    Complete-TTReport -Html $html -Path $Path -Data $flat -Csv:$Csv -Excel:$Excel -DataPath $DataPath -NoHtml:$NoHtml -NoOpen:$NoOpen -Kind 'Risky-Users-Report'
    if ($PassThru) { $data }
}
