function Export-M365DomainReport {
    <#
    .SYNOPSIS
        Generates an HTML report of the tenant's domains (verified, default, auth type).
    .DESCRIPTION
        Lists domains via Microsoft Graph with verification status, default flag and authentication
        type (managed / federated). Read-only.
    .PARAMETER Path
        Target path of the HTML file. Default: .\Domain-Report.html
    .EXAMPLE
        Export-M365DomainReport -BrandName 'CloudNest365'
    #>
    [CmdletBinding()]
    [OutputType([pscustomobject])]
    param(
        [string]$Path = (Join-Path (Get-Location) 'Domain-Report.html'),
        [string]$BrandName = 'TenantToolbox',
        [string]$BrandTagline = 'M365 Tenant Administration',
        [switch]$Csv, [switch]$Excel, [string]$DataPath, [switch]$NoHtml, [switch]$PassThru, [switch]$NoOpen
    )

    Assert-TTGraph
    Write-TTLog -Level INFO -Message "Reading domains ..."
    $doms = Get-TTGraphCollection 'https://graph.microsoft.com/v1.0/domains'

    $data = foreach ($d in $doms) {
        [pscustomobject]@{
            Domain = $d.id; IsDefault = [bool]$d.isDefault; IsVerified = [bool]$d.isVerified
            AuthType = $d.authenticationType; Services = (@($d.supportedServices) -join ', ')
        }
    }
    $data = @($data)
    $total = $data.Count
    $verified = @($data | Where-Object IsVerified).Count
    $federated = @($data | Where-Object { $_.AuthType -eq 'Federated' }).Count
    $genAt = Get-Date -Format 'yyyy-MM-dd HH:mm'; $tenant = (Get-MgContext).TenantId

    $rows = foreach ($d in ($data | Sort-Object @{E = { -[int][bool]$_.IsDefault } }, Domain)) {
        $def = if ($d.IsDefault) { "<span class='b b-info'>default</span>" } else { '<span class="muted">–</span>' }
        $ver = if ($d.IsVerified) { "<span class='b b-ok'>verified</span>" } else { "<span class='b b-warn'>unverified</span>" }
        $auth = if ($d.AuthType -eq 'Federated') { "<span class='b b-warn'>Federated</span>" } else { "$(TTEnc $d.AuthType)" }
        $fFed = if ($d.AuthType -eq 'Federated') { '1' } else { '0' }
        $fUnv = if (-not $d.IsVerified) { '1' } else { '0' }
        $searchAttr = TTEnc ([string]$d.Domain).ToLower()
        @"
      <tr class="item" data-f-federated="$fFed" data-f-unverified="$fUnv" data-name="$searchAttr" data-search="$searchAttr">
        <td><b>$(TTEnc $d.Domain)</b></td><td>$def</td><td>$ver</td><td>$auth</td><td><span class="muted">$(TTEnc $d.Services)</span></td>
      </tr>
"@
    }
    $kpiHtml = New-TTKpis @(
        @{ n = $total; l = 'Domains'; filter = 'all' }
        @{ n = $verified; l = 'Verified'; kind = 'ok' }
        @{ n = ($total - $verified); l = 'Unverified'; kind = 'warn'; filter = 'unverified' }
        @{ n = $federated; l = 'Federated'; kind = 'info'; filter = 'federated' }
    )
    $toolbar = New-TTToolbar -SearchPlaceholder 'Search domain ...' -Filters @(
        @{ label = 'All'; key = 'all' }, @{ label = 'Unverified'; key = 'unverified' }, @{ label = 'Federated'; key = 'federated' }
    )
    $body = @"
    <div class="panel"><table class="tbl"><thead><tr><th data-sort="name">Domain</th><th>Default</th><th>Verified</th><th>Auth type</th><th>Services</th></tr></thead>
      <tbody>
$($rows -join "`n")
      </tbody></table><div class="empty" id="empty">No domain matches the filters.</div></div>
"@
    $html = New-TTHtmlPage -Title 'Domains' -Heading 'Domains' -Sub "Tenant $tenant &middot; generated $genAt &middot; $total domains" `
        -BrandName $BrandName -BrandTagline $BrandTagline -KpiHtml $kpiHtml -ToolbarHtml $toolbar -BodyHtml $body

    Write-TTLog -Level INFO -Message "Domain report created: $Path ($total domains, $verified verified)."
    $flat = $data | Select-Object Domain, IsDefault, IsVerified, AuthType, Services
    Complete-TTReport -Html $html -Path $Path -Data $flat -Csv:$Csv -Excel:$Excel -DataPath $DataPath -NoHtml:$NoHtml -NoOpen:$NoOpen -Kind 'Domain-Report'
    if ($PassThru) { $data }
}
