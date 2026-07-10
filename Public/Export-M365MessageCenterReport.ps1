function Export-M365MessageCenterReport {
    <#
    .SYNOPSIS
        Generates an HTML report of Microsoft 365 Message Center announcements.
    .DESCRIPTION
        Reads Message Center posts via Graph (admin/serviceAnnouncement/messages): upcoming changes,
        action-required items and major changes. Read-only. Requires ServiceMessage.Read.All.
    .PARAMETER Path
        Target path of the HTML file. Default: .\MessageCenter-Report.html
    .EXAMPLE
        Export-M365MessageCenterReport -BrandName 'CloudNest365'
    #>
    [CmdletBinding()]
    [OutputType([pscustomobject])]
    param(
        [string]$Path = (Join-Path (Get-Location) 'MessageCenter-Report.html'),
        [string]$BrandName = 'TenantToolbox',
        [string]$BrandTagline = 'M365 Tenant Administration',
        [switch]$Csv, [switch]$Excel, [string]$DataPath, [switch]$NoHtml, [switch]$PassThru, [switch]$NoOpen
    )

    Assert-TTGraph
    Write-TTLog -Level INFO -Message "Reading Message Center posts ..."
    try { $msgs = Get-TTGraphCollection "https://graph.microsoft.com/v1.0/admin/serviceAnnouncement/messages?`$top=100" }
    catch {
        if ("$_" -match 'Forbidden|403') { Write-Warning "Access denied. Needs the 'ServiceMessage.Read.All' scope. Reconnect: Connect-TenantToolbox -UseDeviceCode" }
        else { Write-Warning "Could not read Message Center: $_" }
        $msgs = @()
    }

    $catLabel = @{ planForChange = 'Plan for change'; preventOrFixIssue = 'Prevent/fix'; stayInformed = 'Stay informed' }
    $data = foreach ($m in $msgs) {
        [pscustomobject]@{
            Title = $m.title; Services = (@($m.services) -join ', '); Category = $m.category
            Major = [bool]$m.isMajorChange; ActionBy = $m.actionRequiredByDateTime; Updated = $m.lastModifiedDateTime; Id = $m.id
        }
    }
    $data = @($data)
    $total = $data.Count
    $action = @($data | Where-Object { $_.ActionBy }).Count
    $major = @($data | Where-Object Major).Count
    $genAt = Get-Date -Format 'yyyy-MM-dd HH:mm'; $tenant = (Get-MgContext).TenantId

    $rows = foreach ($m in ($data | Sort-Object Updated -Descending)) {
        $fAction = if ($m.ActionBy) { '1' } else { '0' }
        $fMajor = if ($m.Major) { '1' } else { '0' }
        $cat = if ($catLabel.ContainsKey([string]$m.Category)) { $catLabel[[string]$m.Category] } else { [string]$m.Category }
        $catBadge = if ($m.Category -eq 'planForChange') { "<span class='b b-warn'>$cat</span>" } elseif ($m.Category -eq 'preventOrFixIssue') { "<span class='b b-bad'>$cat</span>" } else { "<span class='b b-info'>$cat</span>" }
        $major = if ($m.Major) { " <span class='b b-bad'>major</span>" } else { '' }
        $actBy = if ($m.ActionBy) { "<span class='b b-warn'>$(([datetime]$m.ActionBy).ToString('yyyy-MM-dd'))</span>" } else { '<span class="muted">–</span>' }
        $upd = if ($m.Updated) { ([datetime]$m.Updated).ToString('yyyy-MM-dd') } else { '–' }
        $searchAttr = TTEnc ("$($m.Title) $($m.Services)".ToLower()); $nameAttr = TTEnc ([string]$m.Title).ToLower()
        @"
      <tr class="item" data-f-action="$fAction" data-f-major="$fMajor" data-name="$nameAttr" data-search="$searchAttr">
        <td><div class="u"><b>$(TTEnc $m.Title)$major</b><span class="upn">$(TTEnc $m.Services)</span></div></td>
        <td>$catBadge</td><td>$actBy</td><td>$upd</td>
      </tr>
"@
    }
    $kpiHtml = New-TTKpis @(
        @{ n = $total; l = 'Messages'; filter = 'all' }
        @{ n = $action; l = 'Action required'; kind = 'warn'; filter = 'action' }
        @{ n = $major; l = 'Major changes'; kind = 'bad'; filter = 'major' }
    )
    $toolbar = New-TTToolbar -SearchPlaceholder 'Search message ...' -Filters @(
        @{ label = 'All'; key = 'all' }, @{ label = 'Action required'; key = 'action' }, @{ label = 'Major'; key = 'major' }
    )
    $body = @"
    <div class="panel"><table class="tbl"><thead><tr><th data-sort="name">Title</th><th>Category</th><th>Action by</th><th>Updated</th></tr></thead>
      <tbody>
$($rows -join "`n")
      </tbody></table><div class="empty" id="empty">No message matches the filters.</div></div>
"@
    $html = New-TTHtmlPage -Title 'Message Center' -Heading 'Message Center' -Sub "Tenant $tenant &middot; generated $genAt &middot; $total messages" `
        -BrandName $BrandName -BrandTagline $BrandTagline -KpiHtml $kpiHtml -ToolbarHtml $toolbar -BodyHtml $body

    Write-TTLog -Level INFO -Message "Message Center report created: $Path ($total messages, $action action-required)."
    $flat = $data | Select-Object Title, Services, Category, Major,
        @{N = 'ActionBy'; E = { if ($_.ActionBy) { ([datetime]$_.ActionBy).ToString('yyyy-MM-dd') } } },
        @{N = 'Updated'; E = { if ($_.Updated) { ([datetime]$_.Updated).ToString('yyyy-MM-dd') } } }
    Complete-TTReport -Html $html -Path $Path -Data $flat -Csv:$Csv -Excel:$Excel -DataPath $DataPath -NoHtml:$NoHtml -NoOpen:$NoOpen -Kind 'Message-Center-Report'
    if ($PassThru) { $data }
}
