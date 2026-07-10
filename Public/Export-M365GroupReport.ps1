function Export-M365GroupReport {
    <#
    .SYNOPSIS
        Generates an interactive HTML report of groups (type, owners, orphans).
    .DESCRIPTION
        Builds on Get-M365Group and renders a self-contained HTML page: a KPI overview (groups,
        Microsoft 365 groups, orphaned, security groups) and a searchable, filterable table with
        type, owner count and visibility. Read-only.
    .PARAMETER Path
        Target path of the HTML file. Default: .\Group-Report.html
    .PARAMETER BrandName
        Branding. For CloudNest365: -BrandName 'CloudNest365'
    .PARAMETER PassThru
        Also emit the objects on the pipeline.
    .PARAMETER NoOpen
        Do not open the report automatically.
    .EXAMPLE
        Export-M365GroupReport -BrandName 'CloudNest365'
    #>
    [CmdletBinding()]
    [OutputType([pscustomobject])]
    param(
        [string]$Path = (Join-Path (Get-Location) 'Group-Report.html'),
        [string]$BrandName = 'TenantToolbox',
        [string]$BrandTagline = 'M365 Tenant Administration',
        [switch]$Csv, [switch]$Excel, [string]$DataPath, [switch]$NoHtml, [switch]$PassThru, [switch]$NoOpen
    )

    Assert-TTGraph
    $data = @(Get-M365Group)

    $total    = $data.Count
    $m365     = @($data | Where-Object Type -eq 'Microsoft 365').Count
    $orphaned = @($data | Where-Object Orphaned).Count
    $security = @($data | Where-Object Type -eq 'Security').Count
    $genAt    = Get-Date -Format 'yyyy-MM-dd HH:mm'
    $tenant   = (Get-MgContext).TenantId

    $rows = foreach ($g in ($data | Sort-Object @{E = { -[int][bool]$_.Orphaned } }, DisplayName)) {
        $fOrphan = if ($g.Orphaned) { '1' } else { '0' }
        $fM365 = if ($g.Type -eq 'Microsoft 365') { '1' } else { '0' }
        $fSec = if ($g.Type -eq 'Security') { '1' } else { '0' }
        $ownerBadge = if ($g.Orphaned) { "<span class='b b-bad'>0 (orphaned)</span>" } else { "$($g.Owners)" }
        $created = if ($g.Created) { ([datetime]$g.Created).ToString('yyyy-MM-dd') } else { '–' }
        $searchAttr = TTEnc ("$($g.DisplayName) $($g.Type)".ToLower())
        $nameAttr = TTEnc ([string]$g.DisplayName).ToLower()
        @"
      <tr class="item" data-f-orphaned="$fOrphan" data-f-m365="$fM365" data-f-security="$fSec" data-name="$nameAttr" data-search="$searchAttr">
        <td><b>$(TTEnc $g.DisplayName)</b></td>
        <td>$(TTEnc $g.Type)</td>
        <td>$ownerBadge</td>
        <td>$(if ($g.Visibility) { TTEnc $g.Visibility } else { '<span class="muted">–</span>' })</td>
        <td>$created</td>
      </tr>
"@
    }

    $kpiHtml = New-TTKpis @(
        @{ n = $total; l = 'Groups'; filter = 'all' }
        @{ n = $m365; l = 'Microsoft 365'; kind = 'info'; filter = 'm365' }
        @{ n = $orphaned; l = 'Orphaned'; kind = 'bad'; filter = 'orphaned' }
        @{ n = $security; l = 'Security'; kind = 'info'; filter = 'security' }
    )
    $toolbar = New-TTToolbar -SearchPlaceholder 'Search group ...' -Filters @(
        @{ label = 'All'; key = 'all' }, @{ label = 'Orphaned'; key = 'orphaned' },
        @{ label = 'Microsoft 365'; key = 'm365' }, @{ label = 'Security'; key = 'security' }
    )
    $body = @"
    <div class="panel">
      <table class="tbl">
        <thead><tr><th data-sort="name">Group</th><th>Type</th><th>Owners</th><th>Visibility</th><th>Created</th></tr></thead>
        <tbody>
$($rows -join "`n")
        </tbody>
      </table>
      <div class="empty" id="empty">No group matches the filters.</div>
    </div>
"@

    $sub = "Tenant $tenant &middot; generated $genAt &middot; $total groups"
    $html = New-TTHtmlPage -Title 'Groups' -Heading 'Groups & Teams' -Sub $sub `
        -BrandName $BrandName -BrandTagline $BrandTagline -KpiHtml $kpiHtml -ToolbarHtml $toolbar -BodyHtml $body

    Write-TTLog -Level INFO -Message "Group report created: $Path ($total groups, $orphaned orphaned)."
    if ($orphaned -gt 0) { Write-Host "  Note: $orphaned orphaned Microsoft 365 group(s) without an owner." -ForegroundColor Yellow }

    $flat = $data | Select-Object DisplayName, Type, Owners, Orphaned, Visibility,
        @{N = 'Created'; E = { if ($_.Created) { ([datetime]$_.Created).ToString('yyyy-MM-dd') } } }
    Complete-TTReport -Html $html -Path $Path -Data $flat -Csv:$Csv -Excel:$Excel -DataPath $DataPath -NoHtml:$NoHtml -NoOpen:$NoOpen -Kind 'Group-Report'
    if ($PassThru) { $data }
}
