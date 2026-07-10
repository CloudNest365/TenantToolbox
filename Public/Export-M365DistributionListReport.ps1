function Export-M365DistributionListReport {
    <#
    .SYNOPSIS
        Generates an HTML report of distribution lists (members, empty lists).
    .DESCRIPTION
        Lists mail-enabled, non-security, non-Microsoft-365 groups (classic distribution lists) via
        Graph with member count. Read-only. One membership call per list.
    .PARAMETER Path
        Target path of the HTML file. Default: .\DistributionList-Report.html
    .EXAMPLE
        Export-M365DistributionListReport -BrandName 'CloudNest365'
    #>
    [CmdletBinding()]
    [OutputType([pscustomobject])]
    param(
        [string]$Path = (Join-Path (Get-Location) 'DistributionList-Report.html'),
        [string]$BrandName = 'TenantToolbox',
        [string]$BrandTagline = 'M365 Tenant Administration',
        [switch]$Csv, [switch]$Excel, [string]$DataPath, [switch]$NoHtml, [switch]$PassThru, [switch]$NoOpen
    )

    Assert-TTGraph
    Write-TTLog -Level INFO -Message "Reading distribution lists ..."
    $groups = Get-TTGraphCollection "https://graph.microsoft.com/v1.0/groups?`$select=id,displayName,mail,mailEnabled,securityEnabled,groupTypes&`$top=100"
    $dls = @($groups | Where-Object { $_.mailEnabled -and -not $_.securityEnabled -and (@($_.groupTypes) -notcontains 'Unified') })

    $data = foreach ($g in $dls) {
        $members = 0
        try { $members = @(Get-TTGraphCollection "https://graph.microsoft.com/v1.0/groups/$($g.id)/members?`$select=id&`$top=999").Count } catch { }
        [pscustomobject]@{ Name = $g.displayName; Mail = $g.mail; Members = $members; Empty = ($members -eq 0); Id = $g.id }
    }
    $data = @($data)
    $total = $data.Count
    $empty = @($data | Where-Object Empty).Count
    $genAt = Get-Date -Format 'yyyy-MM-dd HH:mm'; $tenant = (Get-MgContext).TenantId

    $rows = foreach ($g in ($data | Sort-Object @{E = { -[int][bool]$_.Empty } }, Name)) {
        $fEmpty = if ($g.Empty) { '1' } else { '0' }
        $memBadge = if ($g.Empty) { "<span class='b b-warn'>0 (empty)</span>" } else { "$($g.Members)" }
        $memSort = ('{0:D6}' -f [int]$g.Members)
        $searchAttr = TTEnc ("$($g.Name) $($g.Mail)".ToLower()); $nameAttr = TTEnc ([string]$g.Name).ToLower()
        @"
      <tr class="item" data-f-empty="$fEmpty" data-name="$nameAttr" data-s-members="$memSort" data-search="$searchAttr">
        <td><div class="u"><b>$(TTEnc $g.Name)</b><span class="upn">$(TTEnc $g.Mail)</span></div></td>
        <td data-s-members="$memSort">$memBadge</td>
      </tr>
"@
    }
    $kpiHtml = New-TTKpis @(
        @{ n = $total; l = 'Distribution lists'; filter = 'all' }
        @{ n = $empty; l = 'Empty'; kind = 'warn'; filter = 'empty' }
        @{ n = ($total - $empty); l = 'With members'; kind = 'ok' }
    )
    $toolbar = New-TTToolbar -SearchPlaceholder 'Search list ...' -Filters @(
        @{ label = 'All'; key = 'all' }, @{ label = 'Empty'; key = 'empty' }
    )
    $body = @"
    <div class="panel"><table class="tbl"><thead><tr><th data-sort="name">Distribution list</th><th data-sort="members">Members</th></tr></thead>
      <tbody>
$($rows -join "`n")
      </tbody></table><div class="empty" id="empty">No list matches the filters.</div></div>
"@
    $html = New-TTHtmlPage -Title 'Distribution Lists' -Heading 'Distribution Lists' -Sub "Tenant $tenant &middot; generated $genAt &middot; $total lists" `
        -BrandName $BrandName -BrandTagline $BrandTagline -KpiHtml $kpiHtml -ToolbarHtml $toolbar -BodyHtml $body

    Write-TTLog -Level INFO -Message "Distribution list report created: $Path ($total lists, $empty empty)."
    $flat = $data | Select-Object Name, Mail, Members, Empty
    Complete-TTReport -Html $html -Path $Path -Data $flat -Csv:$Csv -Excel:$Excel -DataPath $DataPath -NoHtml:$NoHtml -NoOpen:$NoOpen -Kind 'Distribution-List-Report'
    if ($PassThru) { $data }
}
