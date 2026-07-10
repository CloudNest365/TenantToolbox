function Export-M365TeamsReport {
    <#
    .SYNOPSIS
        Generates an interactive HTML report of all Microsoft Teams (owners, members, guests).
    .DESCRIPTION
        Lists team-enabled Microsoft 365 groups via Graph and, per team, the owner/member/guest
        counts and visibility. Read-only. One membership call per team (slower on large tenants).
    .PARAMETER Path
        Target path of the HTML file. Default: .\Teams-Report.html
    .PARAMETER BrandName
        Branding. For CloudNest365: -BrandName 'CloudNest365'
    .PARAMETER PassThru / NoOpen / Csv / Excel / DataPath / NoHtml
        Standard report options.
    .EXAMPLE
        Export-M365TeamsReport -BrandName 'CloudNest365'
    #>
    [CmdletBinding()]
    [OutputType([pscustomobject])]
    param(
        [string]$Path = (Join-Path (Get-Location) 'Teams-Report.html'),
        [string]$BrandName = 'TenantToolbox',
        [string]$BrandTagline = 'M365 Tenant Administration',
        [switch]$Csv, [switch]$Excel, [string]$DataPath, [switch]$NoHtml, [switch]$PassThru, [switch]$NoOpen
    )

    Assert-TTGraph
    Write-TTLog -Level INFO -Message "Reading Microsoft Teams ..."
    try {
        $teams = Get-TTGraphCollection "https://graph.microsoft.com/v1.0/groups?`$filter=resourceProvisioningOptions/Any(x:x eq 'Team')&`$select=id,displayName,visibility&`$expand=owners(`$select=id)&`$top=100"
    }
    catch {
        $teams = Get-TTGraphCollection "https://graph.microsoft.com/v1.0/groups?`$filter=groupTypes/any(c:c eq 'Unified')&`$select=id,displayName,visibility&`$expand=owners(`$select=id)&`$top=100"
    }

    $data = foreach ($t in $teams) {
        $owners = @($t.owners).Count
        $members = 0; $guests = 0
        try {
            $m = Get-TTGraphCollection "https://graph.microsoft.com/v1.0/groups/$($t.id)/members?`$select=id,userType&`$top=999"
            $members = @($m).Count
            $guests = @($m | Where-Object { $_.userType -eq 'Guest' }).Count
        }
        catch { }
        [pscustomobject]@{
            Team = $t.displayName; Visibility = $t.visibility; Owners = $owners; Members = $members; Guests = $guests
            Ownerless = ($owners -eq 0); Id = $t.id
        }
    }
    $data = @($data)

    $total = $data.Count
    $withGuests = @($data | Where-Object { $_.Guests -gt 0 }).Count
    $ownerless = @($data | Where-Object Ownerless).Count
    $public = @($data | Where-Object Visibility -eq 'Public').Count
    $genAt = Get-Date -Format 'yyyy-MM-dd HH:mm'; $tenant = (Get-MgContext).TenantId

    $rows = foreach ($t in ($data | Sort-Object @{E = { -[int][bool]$_.Ownerless } }, Team)) {
        $fGuests = if ($t.Guests -gt 0) { '1' } else { '0' }
        $fOwnerless = if ($t.Ownerless) { '1' } else { '0' }
        $fPublic = if ($t.Visibility -eq 'Public') { '1' } else { '0' }
        $ownerBadge = if ($t.Ownerless) { "<span class='b b-bad'>0</span>" } else { "$($t.Owners)" }
        $guestBadge = if ($t.Guests -gt 0) { "<span class='b b-warn'>$($t.Guests)</span>" } else { "<span class='muted'>0</span>" }
        $vis = if ($t.Visibility -eq 'Public') { "<span class='b b-info'>Public</span>" } else { "<span class='muted'>Private</span>" }
        $searchAttr = TTEnc ([string]$t.Team).ToLower(); $nameAttr = $searchAttr
        @"
      <tr class="item" data-f-guests="$fGuests" data-f-ownerless="$fOwnerless" data-f-public="$fPublic" data-name="$nameAttr" data-search="$searchAttr">
        <td><b>$(TTEnc $t.Team)</b></td>
        <td>$vis</td>
        <td>$ownerBadge</td>
        <td>$($t.Members)</td>
        <td>$guestBadge</td>
      </tr>
"@
    }

    $kpiHtml = New-TTKpis @(
        @{ n = $total; l = 'Teams'; filter = 'all' }
        @{ n = $withGuests; l = 'With guests'; kind = 'warn'; filter = 'guests' }
        @{ n = $ownerless; l = 'Ownerless'; kind = 'bad'; filter = 'ownerless' }
        @{ n = $public; l = 'Public'; kind = 'info'; filter = 'public' }
    )
    $toolbar = New-TTToolbar -SearchPlaceholder 'Search team ...' -Filters @(
        @{ label = 'All'; key = 'all' }, @{ label = 'With guests'; key = 'guests' }, @{ label = 'Ownerless'; key = 'ownerless' }, @{ label = 'Public'; key = 'public' }
    )
    $body = @"
    <div class="panel"><table class="tbl"><thead><tr><th data-sort="name">Team</th><th>Visibility</th><th>Owners</th><th>Members</th><th>Guests</th></tr></thead>
      <tbody>
$($rows -join "`n")
      </tbody></table><div class="empty" id="empty">No team matches the filters.</div></div>
"@
    $html = New-TTHtmlPage -Title 'Microsoft Teams' -Heading 'Microsoft Teams' -Sub "Tenant $tenant &middot; generated $genAt &middot; $total teams" `
        -BrandName $BrandName -BrandTagline $BrandTagline -KpiHtml $kpiHtml -ToolbarHtml $toolbar -BodyHtml $body

    Write-TTLog -Level INFO -Message "Teams report created: $Path ($total teams, $ownerless ownerless, $withGuests with guests)."
    $flat = $data | Select-Object Team, Visibility, Owners, Members, Guests, Ownerless
    Complete-TTReport -Html $html -Path $Path -Data $flat -Csv:$Csv -Excel:$Excel -DataPath $DataPath -NoHtml:$NoHtml -NoOpen:$NoOpen -Kind 'Teams-Report'
    if ($PassThru) { $data }
}
