function Export-M365AdminRoleReport {
    <#
    .SYNOPSIS
        Generates an interactive HTML report of admin role assignments (with MFA status).
    .DESCRIPTION
        Builds on Get-M365AdminRole and renders a self-contained HTML page: a KPI overview
        (assignments, unique admins, without MFA, Global Admins) and a searchable, filterable
        table of role holders with MFA badges. Read-only.
    .PARAMETER Path
        Target path of the HTML file. Default: .\AdminRole-Report.html
    .PARAMETER BrandName
        Branding. For CloudNest365: -BrandName 'CloudNest365'
    .PARAMETER PassThru
        Also emit the objects on the pipeline.
    .PARAMETER NoOpen
        Do not open the report automatically.
    .EXAMPLE
        Export-M365AdminRoleReport -BrandName 'CloudNest365'
    #>
    [CmdletBinding()]
    [OutputType([pscustomobject])]
    param(
        [string]$Path = (Join-Path (Get-Location) 'AdminRole-Report.html'),
        [string]$BrandName = 'TenantToolbox',
        [string]$BrandTagline = 'M365 Tenant Administration',
        [switch]$Csv, [switch]$Excel, [string]$DataPath, [switch]$NoHtml, [switch]$PassThru, [switch]$NoOpen
    )

    Assert-TTGraph
    $data = @(Get-M365AdminRole)

    $total    = $data.Count
    $admins   = @($data | Where-Object Type -eq 'user' | Select-Object -ExpandProperty UPN -Unique).Count
    $noMfa    = @($data | Where-Object { $_.Type -eq 'user' -and $_.Privileged -and $_.MfaRegistered -eq $false }).Count
    $ga       = @($data | Where-Object Role -eq 'Global Administrator').Count
    $genAt    = Get-Date -Format 'yyyy-MM-dd HH:mm'
    $tenant   = (Get-MgContext).TenantId

    $rows = foreach ($r in ($data | Sort-Object @{E = { -[int][bool]$_.Privileged } }, Role, Member)) {
        $isGa = $r.Role -eq 'Global Administrator'
        $rolePill = if ($isGa) { " <span class='b b-crit'>&#9888; GA</span>" } elseif ($r.Privileged) { " <span class='b b-info'>privileged</span>" } else { '' }
        $mfaBadge = if ($r.Type -ne 'user') { "<span class='muted'>–</span>" }
                    elseif ($r.MfaRegistered -eq $true) { "<span class='b b-ok'>MFA</span>" }
                    elseif ($r.MfaRegistered -eq $false) { "<span class='b b-bad'>no MFA</span>" }
                    else { "<span class='muted'>?</span>" }
        $fPriv = if ($r.Privileged) { '1' } else { '0' }
        $fGa = if ($isGa) { '1' } else { '0' }
        $fNoMfa = if ($r.Type -eq 'user' -and $r.MfaRegistered -eq $false) { '1' } else { '0' }
        $searchAttr = TTEnc ("$($r.Member) $($r.UPN) $($r.Role)".ToLower())
        $nameAttr = TTEnc ([string]$r.Member).ToLower()
        $roleAttr = TTEnc ([string]$r.Role).ToLower()
        @"
      <tr class="item" data-f-privileged="$fPriv" data-f-globaladmin="$fGa" data-f-nomfa="$fNoMfa" data-name="$nameAttr" data-s-role="$roleAttr" data-search="$searchAttr">
        <td><div class="u"><b>$(TTEnc $r.Member)</b><span class="upn">$(TTEnc $r.UPN)</span></div></td>
        <td>$(TTEnc $r.Role)$rolePill</td>
        <td>$(TTEnc $r.Type)</td>
        <td>$mfaBadge</td>
      </tr>
"@
    }

    $kpiHtml = New-TTKpis @(
        @{ n = $total; l = 'Assignments'; filter = 'all' }
        @{ n = $admins; l = 'Unique admins'; kind = 'info' }
        @{ n = $noMfa; l = 'Priv. without MFA'; kind = 'bad'; filter = 'nomfa' }
        @{ n = $ga; l = 'Global Admins'; kind = 'adm'; filter = 'globaladmin' }
    )
    $toolbar = New-TTToolbar -SearchPlaceholder 'Search admin or role ...' -Filters @(
        @{ label = 'All'; key = 'all' }, @{ label = 'Privileged'; key = 'privileged' },
        @{ label = 'Without MFA'; key = 'nomfa' }, @{ label = 'Global Admins'; key = 'globaladmin' }
    )
    $body = @"
    <div class="panel">
      <table class="tbl">
        <thead><tr><th data-sort="name">Member</th><th data-sort="role">Role</th><th>Type</th><th>MFA</th></tr></thead>
        <tbody>
$($rows -join "`n")
        </tbody>
      </table>
      <div class="empty" id="empty">No assignment matches the filters.</div>
    </div>
"@

    $sub = "Tenant $tenant &middot; generated $genAt &middot; $total role assignments"
    $html = New-TTHtmlPage -Title 'Admin Roles' -Heading 'Admin Role Assignments' -Sub $sub `
        -BrandName $BrandName -BrandTagline $BrandTagline -KpiHtml $kpiHtml -ToolbarHtml $toolbar -BodyHtml $body

    Write-TTLog -Level INFO -Message "Admin role report created: $Path ($total assignments, $noMfa privileged without MFA, $ga Global Admins)."
    if ($noMfa -gt 0) { Write-Host "  Warning: $noMfa privileged admin(s) without MFA!" -ForegroundColor Red }

    $flat = $data | Select-Object Member, UPN, Role, Privileged, Type, MfaRegistered
    Complete-TTReport -Html $html -Path $Path -Data $flat -Csv:$Csv -Excel:$Excel -DataPath $DataPath -NoHtml:$NoHtml -NoOpen:$NoOpen -Kind 'Admin-Role-Report'
    if ($PassThru) { $data }
}
