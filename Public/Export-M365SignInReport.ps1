function Export-M365SignInReport {
    <#
    .SYNOPSIS
        Generates an HTML report of recent sign-ins (failed, legacy-auth, risky).
    .DESCRIPTION
        Reads recent sign-in logs via Graph (auditLogs/signIns) and flags failures, legacy
        authentication and risk. Read-only. Requires AuditLog.Read.All (and Entra ID P1+ for logs).
    .PARAMETER Top
        Number of most recent sign-ins to analyze. Default: 500.
    .PARAMETER Path
        Target path of the HTML file. Default: .\SignIn-Report.html
    .EXAMPLE
        Export-M365SignInReport -BrandName 'CloudNest365'
    #>
    [CmdletBinding()]
    [OutputType([pscustomobject])]
    param(
        [int]$Top = 500,
        [string]$Path = (Join-Path (Get-Location) 'SignIn-Report.html'),
        [string]$BrandName = 'TenantToolbox',
        [string]$BrandTagline = 'M365 Tenant Administration',
        [switch]$Csv, [switch]$Excel, [string]$DataPath, [switch]$NoHtml, [switch]$PassThru, [switch]$NoOpen
    )

    Assert-TTGraph
    Write-TTLog -Level INFO -Message "Reading recent sign-ins ..."
    try { $signins = Get-TTGraphCollection "https://graph.microsoft.com/v1.0/auditLogs/signIns?`$top=$([math]::Min($Top,1000))" }
    catch {
        if ("$_" -match 'Forbidden|403') { Write-Warning "Access denied. Needs AuditLog.Read.All (and Entra ID P1+). Reconnect: Connect-TenantToolbox -UseDeviceCode" }
        else { Write-Warning "Could not read sign-ins: $_" }
        $signins = @()
    }

    $modern = @('Browser', 'Mobile Apps and Desktop clients')
    $data = foreach ($s in $signins) {
        $failed = [int]$s.status.errorCode -ne 0
        $client = [string]$s.clientAppUsed
        $legacy = $client -and ($client -notin $modern)
        $risk = [string]$s.riskLevelAggregated
        $risky = $risk -in 'low', 'medium', 'high'
        [pscustomobject]@{
            User = $s.userPrincipalName; App = $s.appDisplayName; Client = $client
            Failed = $failed; ErrorCode = $s.status.errorCode; Reason = $s.status.failureReason
            Legacy = $legacy; Risk = $risk; Risky = $risky; IP = $s.ipAddress; Time = $s.createdDateTime
        }
    }
    $data = @($data)
    $total = $data.Count
    $failed = @($data | Where-Object Failed).Count
    $legacy = @($data | Where-Object Legacy).Count
    $risky = @($data | Where-Object Risky).Count
    $genAt = Get-Date -Format 'yyyy-MM-dd HH:mm'; $tenant = (Get-MgContext).TenantId

    $rows = foreach ($s in ($data | Sort-Object Time -Descending)) {
        $fFail = if ($s.Failed) { '1' } else { '0' }
        $fLeg = if ($s.Legacy) { '1' } else { '0' }
        $fRisk = if ($s.Risky) { '1' } else { '0' }
        $res = if ($s.Failed) { "<span class='b b-bad'>fail $($s.ErrorCode)</span>" } else { "<span class='b b-ok'>success</span>" }
        $riskB = if ($s.Risk -eq 'high') { "<span class='b b-bad'>high</span>" } elseif ($s.Risk -in 'low', 'medium') { "<span class='b b-warn'>$($s.Risk)</span>" } else { "<span class='muted'>&#8211;</span>" }
        $clientB = if ($s.Legacy) { "<span class='b b-warn'>$(TTEnc $s.Client)</span>" } else { "$(TTEnc $s.Client)" }
        $t = if ($s.Time) { ([datetime]$s.Time).ToString('yyyy-MM-dd HH:mm') } else { '' }
        $searchAttr = TTEnc ("$($s.User) $($s.App) $($s.Client) $($s.IP)".ToLower()); $nameAttr = TTEnc ([string]$s.User).ToLower()
        @"
      <tr class="item" data-f-failed="$fFail" data-f-legacy="$fLeg" data-f-risky="$fRisk" data-name="$nameAttr" data-search="$searchAttr">
        <td><div class="u"><b>$(TTEnc $s.User)</b><span class="upn">$(TTEnc $s.IP)</span></div></td>
        <td>$(TTEnc $s.App)</td><td>$clientB</td><td>$res</td><td>$riskB</td><td>$t</td>
      </tr>
"@
    }
    $kpiHtml = New-TTKpis @(
        @{ n = $total; l = 'Sign-ins'; filter = 'all' }
        @{ n = $failed; l = 'Failed'; kind = 'bad'; filter = 'failed' }
        @{ n = $legacy; l = 'Legacy auth'; kind = 'warn'; filter = 'legacy' }
        @{ n = $risky; l = 'Risky'; kind = 'bad'; filter = 'risky' }
    )
    $toolbar = New-TTToolbar -SearchPlaceholder 'Search user, app or IP ...' -Filters @(
        @{ label = 'All'; key = 'all' }, @{ label = 'Failed'; key = 'failed' }, @{ label = 'Legacy'; key = 'legacy' }, @{ label = 'Risky'; key = 'risky' }
    )
    $body = @"
    <div class="panel"><table class="tbl"><thead><tr><th data-sort="name">User</th><th>App</th><th>Client</th><th>Result</th><th>Risk</th><th>Time</th></tr></thead>
      <tbody>
$($rows -join "`n")
      </tbody></table><div class="empty" id="empty">No sign-in matches the filters.</div></div>
"@
    $html = New-TTHtmlPage -Title 'Sign-ins' -Heading 'Sign-in Analysis' -Sub "Tenant $tenant &middot; generated $genAt &middot; last $total sign-ins" `
        -BrandName $BrandName -BrandTagline $BrandTagline -KpiHtml $kpiHtml -ToolbarHtml $toolbar -BodyHtml $body

    Write-TTLog -Level INFO -Message "Sign-in report created: $Path ($failed failed, $legacy legacy, $risky risky)."
    $flat = $data | Select-Object User, App, Client, Failed, ErrorCode, Reason, Legacy, Risk, IP,
        @{N = 'Time'; E = { if ($_.Time) { ([datetime]$_.Time).ToString('yyyy-MM-dd HH:mm') } } }
    Complete-TTReport -Html $html -Path $Path -Data $flat -Csv:$Csv -Excel:$Excel -DataPath $DataPath -NoHtml:$NoHtml -NoOpen:$NoOpen -Kind 'Sign-In-Report'
    if ($PassThru) { $data }
}
