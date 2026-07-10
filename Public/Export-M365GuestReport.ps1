function Export-M365GuestReport {
    <#
    .SYNOPSIS
        Generates an interactive HTML report of guest (external) accounts.
    .DESCRIPTION
        Builds on Get-M365Guest and renders a self-contained HTML page: a KPI overview (guests,
        pending, stale, never signed in) and a searchable, filterable table with domain, state and
        last sign-in. Read-only.
    .PARAMETER Path
        Target path of the HTML file. Default: .\Guest-Report.html
    .PARAMETER InactiveDays
        Threshold in days for the stale flag. Default: 90.
    .PARAMETER BrandName
        Branding. For CloudNest365: -BrandName 'CloudNest365'
    .PARAMETER PassThru
        Also emit the objects on the pipeline.
    .PARAMETER NoOpen
        Do not open the report automatically.
    .EXAMPLE
        Export-M365GuestReport -BrandName 'CloudNest365'
    #>
    [CmdletBinding()]
    [OutputType([pscustomobject])]
    param(
        [string]$Path = (Join-Path (Get-Location) 'Guest-Report.html'),
        [int]$InactiveDays = 90,
        [string]$BrandName = 'TenantToolbox',
        [string]$BrandTagline = 'M365 Tenant Administration',
        [switch]$Csv, [switch]$Excel, [string]$DataPath, [switch]$NoHtml, [switch]$PassThru, [switch]$NoOpen
    )

    Assert-TTGraph
    $data = @(Get-M365Guest -InactiveDays $InactiveDays)

    $total   = $data.Count
    $pending = @($data | Where-Object { $_.State -eq 'PendingAcceptance' }).Count
    $stale   = @($data | Where-Object Stale).Count
    $never   = @($data | Where-Object NeverSignedIn).Count
    $genAt   = Get-Date -Format 'yyyy-MM-dd HH:mm'
    $tenant  = (Get-MgContext).TenantId

    $rows = foreach ($g in ($data | Sort-Object @{E = { -[int][bool]$_.Stale } }, DisplayName)) {
        $fStale = if ($g.Stale) { '1' } else { '0' }
        $fPending = if ($g.State -eq 'PendingAcceptance') { '1' } else { '0' }
        $stateBadge = if ($g.State -eq 'Accepted') { "<span class='b b-ok'>accepted</span>" } elseif ($g.State -eq 'PendingAcceptance') { "<span class='b b-warn'>pending</span>" } else { "<span class='muted'>$(TTEnc $g.State)</span>" }
        $syncTxt = if ($g.NeverSignedIn) { "<span class='b b-bad'>never</span>" } elseif ($g.LastSignIn) { "$($g.LastSignIn.ToString('yyyy-MM-dd')) ($($g.DaysInactive)d)" } else { '–' }
        $created = if ($g.Created) { ([datetime]$g.Created).ToString('yyyy-MM-dd') } else { '–' }
        $searchAttr = TTEnc ("$($g.DisplayName) $($g.Mail) $($g.Domain)".ToLower())
        $nameAttr = TTEnc ([string]$g.DisplayName).ToLower()
        @"
      <tr class="item" data-f-stale="$fStale" data-f-pending="$fPending" data-name="$nameAttr" data-search="$searchAttr">
        <td><div class="u"><b>$(TTEnc $g.DisplayName)</b><span class="upn">$(TTEnc $g.Mail)</span></div></td>
        <td>$(TTEnc $g.Domain)</td>
        <td>$stateBadge</td>
        <td>$created</td>
        <td>$syncTxt</td>
      </tr>
"@
    }

    $kpiHtml = New-TTKpis @(
        @{ n = $total; l = 'Guests'; filter = 'all' }
        @{ n = $pending; l = 'Pending'; kind = 'warn'; filter = 'pending' }
        @{ n = $stale; l = "Stale (>$InactiveDays d)"; kind = 'bad'; filter = 'stale' }
        @{ n = $never; l = 'Never signed in'; kind = 'bad' }
    )
    $toolbar = New-TTToolbar -SearchPlaceholder 'Search guest or domain ...' -Filters @(
        @{ label = 'All'; key = 'all' }, @{ label = 'Pending'; key = 'pending' }, @{ label = 'Stale'; key = 'stale' }
    )
    $body = @"
    <div class="panel">
      <table class="tbl">
        <thead><tr><th data-sort="name">Guest</th><th>Domain</th><th>State</th><th>Created</th><th>Last sign-in</th></tr></thead>
        <tbody>
$($rows -join "`n")
        </tbody>
      </table>
      <div class="empty" id="empty">No guest matches the filters.</div>
    </div>
"@

    $sub = "Tenant $tenant &middot; generated $genAt &middot; $total guests"
    $html = New-TTHtmlPage -Title 'Guest Accounts' -Heading 'Guest Accounts' -Sub $sub `
        -BrandName $BrandName -BrandTagline $BrandTagline -KpiHtml $kpiHtml -ToolbarHtml $toolbar -BodyHtml $body

    Write-TTLog -Level INFO -Message "Guest report created: $Path ($total guests, $stale stale, $pending pending)."

    $flat = $data | Select-Object DisplayName, Mail, Domain, State,
        @{N = 'Created'; E = { if ($_.Created) { ([datetime]$_.Created).ToString('yyyy-MM-dd') } } },
        @{N = 'LastSignIn'; E = { if ($_.LastSignIn) { $_.LastSignIn.ToString('yyyy-MM-dd') } } }, DaysInactive, Stale
    Complete-TTReport -Html $html -Path $Path -Data $flat -Csv:$Csv -Excel:$Excel -DataPath $DataPath -NoHtml:$NoHtml -NoOpen:$NoOpen -Kind 'Guest-Report'
    if ($PassThru) { $data }
}
