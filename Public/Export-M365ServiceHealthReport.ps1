function Export-M365ServiceHealthReport {
    <#
    .SYNOPSIS
        Generates an HTML report of current Microsoft 365 service health issues.
    .DESCRIPTION
        Reads active service issues via Graph (admin/serviceAnnouncement/issues) and shows service,
        title, classification (incident/advisory), status and start time. Read-only. Requires
        ServiceHealth.Read.All.
    .PARAMETER Path
        Target path of the HTML file. Default: .\ServiceHealth-Report.html
    .EXAMPLE
        Export-M365ServiceHealthReport -BrandName 'CloudNest365'
    #>
    [CmdletBinding()]
    [OutputType([pscustomobject])]
    param(
        [string]$Path = (Join-Path (Get-Location) 'ServiceHealth-Report.html'),
        [string]$BrandName = 'TenantToolbox',
        [string]$BrandTagline = 'M365 Tenant Administration',
        [switch]$Csv, [switch]$Excel, [string]$DataPath, [switch]$NoHtml, [switch]$PassThru, [switch]$NoOpen
    )

    Assert-TTGraph
    Write-TTLog -Level INFO -Message "Reading service health issues ..."
    try { $issues = Get-TTGraphCollection "https://graph.microsoft.com/v1.0/admin/serviceAnnouncement/issues?`$top=100" }
    catch {
        if ("$_" -match 'Forbidden|403') { Write-Warning "Access denied. Needs the 'ServiceHealth.Read.All' scope. Reconnect: Connect-TenantToolbox -UseDeviceCode" }
        else { Write-Warning "Could not read service issues: $_" }
        $issues = @()
    }

    $data = foreach ($i in $issues) {
        [pscustomobject]@{
            Service = $i.service; Title = $i.title; Classification = $i.classification
            Status = $i.status; Started = $i.startDateTime; Feature = $i.feature; Id = $i.id
        }
    }
    $data = @($data)
    $active = @($data | Where-Object { $_.Status -notin 'serviceRestored', 'resolved', 'falsePositive', 'postIncidentReviewPublished' })
    $total = $data.Count
    $activeN = $active.Count
    $incidents = @($data | Where-Object { $_.Classification -eq 'incident' }).Count
    $advisories = @($data | Where-Object { $_.Classification -eq 'advisory' }).Count
    $genAt = Get-Date -Format 'yyyy-MM-dd HH:mm'; $tenant = (Get-MgContext).TenantId

    $activeSet = @{}; foreach ($a in $active) { $activeSet[$a.Id] = $true }
    $rows = foreach ($i in ($data | Sort-Object Started -Descending)) {
        $isActive = $activeSet.ContainsKey($i.Id)
        $fInc = if ($i.Classification -eq 'incident') { '1' } else { '0' }
        $fActive = if ($isActive) { '1' } else { '0' }
        $cls = if ($i.Classification -eq 'incident') { "<span class='b b-bad'>incident</span>" } else { "<span class='b b-warn'>advisory</span>" }
        $st = if ($isActive) { "<span class='b b-warn'>$(TTEnc $i.Status)</span>" } else { "<span class='muted'>$(TTEnc $i.Status)</span>" }
        $started = if ($i.Started) { ([datetime]$i.Started).ToString('yyyy-MM-dd HH:mm') } else { '–' }
        $searchAttr = TTEnc ("$($i.Service) $($i.Title)".ToLower()); $nameAttr = TTEnc ([string]$i.Title).ToLower()
        @"
      <tr class="item" data-f-incident="$fInc" data-f-active="$fActive" data-name="$nameAttr" data-search="$searchAttr">
        <td><b>$(TTEnc $i.Service)</b></td><td>$(TTEnc $i.Title)</td><td>$cls</td><td>$st</td><td>$started</td>
      </tr>
"@
    }
    $kpiHtml = New-TTKpis @(
        @{ n = $total; l = 'Issues'; filter = 'all' }
        @{ n = $activeN; l = 'Active'; kind = 'warn'; filter = 'active' }
        @{ n = $incidents; l = 'Incidents'; kind = 'bad'; filter = 'incident' }
        @{ n = $advisories; l = 'Advisories'; kind = 'info' }
    )
    $toolbar = New-TTToolbar -SearchPlaceholder 'Search service or title ...' -Filters @(
        @{ label = 'All'; key = 'all' }, @{ label = 'Active'; key = 'active' }, @{ label = 'Incidents'; key = 'incident' }
    )
    $body = @"
    <div class="panel"><table class="tbl"><thead><tr><th data-sort="name">Service</th><th>Title</th><th>Type</th><th>Status</th><th>Started</th></tr></thead>
      <tbody>
$($rows -join "`n")
      </tbody></table><div class="empty" id="empty">No issue matches the filters.</div></div>
"@
    $html = New-TTHtmlPage -Title 'Service Health' -Heading 'Microsoft 365 Service Health' -Sub "Tenant $tenant &middot; generated $genAt &middot; $activeN active issues" `
        -BrandName $BrandName -BrandTagline $BrandTagline -KpiHtml $kpiHtml -ToolbarHtml $toolbar -BodyHtml $body

    Write-TTLog -Level INFO -Message "Service health report created: $Path ($activeN active, $incidents incidents)."
    $flat = $data | Select-Object Service, Title, Classification, Status,
        @{N = 'Started'; E = { if ($_.Started) { ([datetime]$_.Started).ToString('yyyy-MM-dd HH:mm') } } }
    Complete-TTReport -Html $html -Path $Path -Data $flat -Csv:$Csv -Excel:$Excel -DataPath $DataPath -NoHtml:$NoHtml -NoOpen:$NoOpen -Kind 'Service-Health-Report'
    if ($PassThru) { $data }
}
