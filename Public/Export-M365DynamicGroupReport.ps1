function Export-M365DynamicGroupReport {
    <#
    .SYNOPSIS
        Generates an HTML report of dynamic groups and their membership rules.
    .DESCRIPTION
        Lists groups with dynamic membership via Graph and shows the membership rule and its
        processing state (On / Paused). Read-only.
    .PARAMETER Path
        Target path of the HTML file. Default: .\DynamicGroup-Report.html
    .EXAMPLE
        Export-M365DynamicGroupReport -BrandName 'CloudNest365'
    #>
    [CmdletBinding()]
    [OutputType([pscustomobject])]
    param(
        [string]$Path = (Join-Path (Get-Location) 'DynamicGroup-Report.html'),
        [string]$BrandName = 'TenantToolbox',
        [string]$BrandTagline = 'M365 Tenant Administration',
        [switch]$Csv, [switch]$Excel, [string]$DataPath, [switch]$NoHtml, [switch]$PassThru, [switch]$NoOpen
    )

    Assert-TTGraph
    Write-TTLog -Level INFO -Message "Reading dynamic groups ..."
    $groups = Get-TTGraphCollection "https://graph.microsoft.com/v1.0/groups?`$filter=groupTypes/any(c:c eq 'DynamicMembership')&`$select=displayName,membershipRule,membershipRuleProcessingState,securityEnabled,groupTypes&`$top=100"

    $data = foreach ($g in $groups) {
        $kind = if (@($g.groupTypes) -contains 'Unified') { 'Microsoft 365' } elseif ($g.securityEnabled) { 'Security' } else { 'Other' }
        [pscustomobject]@{
            Group = $g.displayName; Kind = $kind; Rule = $g.membershipRule
            State = $g.membershipRuleProcessingState; Paused = ($g.membershipRuleProcessingState -eq 'Paused')
        }
    }
    $data = @($data)
    $total = $data.Count
    $paused = @($data | Where-Object Paused).Count
    $genAt = Get-Date -Format 'yyyy-MM-dd HH:mm'; $tenant = (Get-MgContext).TenantId

    $rows = foreach ($g in ($data | Sort-Object @{E = { -[int][bool]$_.Paused } }, Group)) {
        $fPaused = if ($g.Paused) { '1' } else { '0' }
        $st = if ($g.Paused) { "<span class='b b-warn'>Paused</span>" } else { "<span class='b b-ok'>On</span>" }
        $searchAttr = TTEnc ("$($g.Group) $($g.Rule)".ToLower()); $nameAttr = TTEnc ([string]$g.Group).ToLower()
        @"
      <tr class="item" data-f-paused="$fPaused" data-name="$nameAttr" data-search="$searchAttr">
        <td><b>$(TTEnc $g.Group)</b></td><td>$(TTEnc $g.Kind)</td><td>$st</td>
        <td><code style="font-size:11px">$(TTEnc $g.Rule)</code></td>
      </tr>
"@
    }
    $kpiHtml = New-TTKpis @(
        @{ n = $total; l = 'Dynamic groups'; filter = 'all' }
        @{ n = $paused; l = 'Paused'; kind = 'warn'; filter = 'paused' }
        @{ n = ($total - $paused); l = 'Active'; kind = 'ok' }
    )
    $toolbar = New-TTToolbar -SearchPlaceholder 'Search group or rule ...' -Filters @(
        @{ label = 'All'; key = 'all' }, @{ label = 'Paused'; key = 'paused' }
    )
    $body = @"
    <div class="panel"><table class="tbl"><thead><tr><th data-sort="name">Group</th><th>Type</th><th>State</th><th>Membership rule</th></tr></thead>
      <tbody>
$($rows -join "`n")
      </tbody></table><div class="empty" id="empty">No group matches the filters.</div></div>
"@
    $html = New-TTHtmlPage -Title 'Dynamic Groups' -Heading 'Dynamic Groups' -Sub "Tenant $tenant &middot; generated $genAt &middot; $total dynamic groups" `
        -BrandName $BrandName -BrandTagline $BrandTagline -KpiHtml $kpiHtml -ToolbarHtml $toolbar -BodyHtml $body

    Write-TTLog -Level INFO -Message "Dynamic group report created: $Path ($total groups, $paused paused)."
    $flat = $data | Select-Object Group, Kind, State, Rule
    Complete-TTReport -Html $html -Path $Path -Data $flat -Csv:$Csv -Excel:$Excel -DataPath $DataPath -NoHtml:$NoHtml -NoOpen:$NoOpen -Kind 'Dynamic-Group-Report'
    if ($PassThru) { $data }
}
