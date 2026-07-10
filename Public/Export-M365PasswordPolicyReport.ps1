function Export-M365PasswordPolicyReport {
    <#
    .SYNOPSIS
        Generates an HTML report of password-policy findings (never expires, weak, old).
    .DESCRIPTION
        Reads enabled member users via Graph and flags those whose password never expires
        ('DisablePasswordExpiration'), has strong-password disabled, or is older than a threshold.
        Read-only.
    .PARAMETER OldDays
        Flag passwords older than this many days. Default: 365.
    .PARAMETER Path
        Target path of the HTML file. Default: .\PasswordPolicy-Report.html
    .EXAMPLE
        Export-M365PasswordPolicyReport -BrandName 'CloudNest365'
    #>
    [CmdletBinding()]
    [OutputType([pscustomobject])]
    param(
        [int]$OldDays = 365,
        [string]$Path = (Join-Path (Get-Location) 'PasswordPolicy-Report.html'),
        [string]$BrandName = 'TenantToolbox',
        [string]$BrandTagline = 'M365 Tenant Administration',
        [switch]$Csv, [switch]$Excel, [string]$DataPath, [switch]$NoHtml, [switch]$PassThru, [switch]$NoOpen
    )

    Assert-TTGraph
    Write-TTLog -Level INFO -Message "Reading password-policy findings ..."
    $users = Get-MgUser -All -Filter "accountEnabled eq true and userType eq 'Member'" -Property 'displayName,userPrincipalName,passwordPolicies,lastPasswordChangeDateTime' -ErrorAction Stop

    $data = foreach ($u in $users) {
        $pp = [string]$u.PasswordPolicies
        $never = $pp -match 'DisablePasswordExpiration'
        $weak = $pp -match 'DisableStrongPassword'
        $age = if ($u.LastPasswordChangeDateTime) { [math]::Floor(((Get-Date) - [datetime]$u.LastPasswordChangeDateTime).TotalDays) } else { $null }
        $old = ($null -ne $age) -and ($age -gt $OldDays)
        if (-not ($never -or $weak -or $old)) { continue }
        [pscustomobject]@{ User = $u.DisplayName; UPN = $u.UserPrincipalName; NeverExpires = $never; WeakAllowed = $weak; PasswordAge = $age; Old = $old }
    }
    $data = @($data)
    $total = $data.Count
    $never = @($data | Where-Object NeverExpires).Count
    $weak = @($data | Where-Object WeakAllowed).Count
    $old = @($data | Where-Object Old).Count
    $genAt = Get-Date -Format 'yyyy-MM-dd HH:mm'; $tenant = (Get-MgContext).TenantId

    $rows = foreach ($u in ($data | Sort-Object @{E = { -[int][bool]$_.NeverExpires } }, User)) {
        $fNever = if ($u.NeverExpires) { '1' } else { '0' }
        $fWeak = if ($u.WeakAllowed) { '1' } else { '0' }
        $fOld = if ($u.Old) { '1' } else { '0' }
        $neverB = if ($u.NeverExpires) { "<span class='b b-bad'>never expires</span>" } else { '<span class="muted">–</span>' }
        $weakB = if ($u.WeakAllowed) { "<span class='b b-warn'>weak allowed</span>" } else { '<span class="muted">–</span>' }
        $ageB = if ($null -eq $u.PasswordAge) { '<span class="muted">?</span>' } elseif ($u.Old) { "<span class='b b-warn'>$($u.PasswordAge)d</span>" } else { "$($u.PasswordAge)d" }
        $ageSort = ('{0:D7}' -f [int]([math]::Max(0, [int]$u.PasswordAge)))
        $searchAttr = TTEnc ("$($u.User) $($u.UPN)".ToLower()); $nameAttr = TTEnc ([string]$u.User).ToLower()
        @"
      <tr class="item" data-f-neverexpires="$fNever" data-f-weak="$fWeak" data-f-old="$fOld" data-name="$nameAttr" data-s-age="$ageSort" data-search="$searchAttr">
        <td><div class="u"><b>$(TTEnc $u.User)</b><span class="upn">$(TTEnc $u.UPN)</span></div></td>
        <td>$neverB</td><td>$weakB</td><td data-s-age="$ageSort">$ageB</td>
      </tr>
"@
    }
    $kpiHtml = New-TTKpis @(
        @{ n = $total; l = 'Findings'; filter = 'all' }
        @{ n = $never; l = 'Never expires'; kind = 'bad'; filter = 'neverexpires' }
        @{ n = $weak; l = 'Weak allowed'; kind = 'warn'; filter = 'weak' }
        @{ n = $old; l = "Older than $OldDays d"; kind = 'warn'; filter = 'old' }
    )
    $toolbar = New-TTToolbar -SearchPlaceholder 'Search user ...' -Filters @(
        @{ label = 'All'; key = 'all' }, @{ label = 'Never expires'; key = 'neverexpires' }, @{ label = 'Weak'; key = 'weak' }, @{ label = 'Old'; key = 'old' }
    )
    $body = @"
    <div class="panel"><table class="tbl"><thead><tr><th data-sort="name">User</th><th>Expiry</th><th>Strong password</th><th data-sort="age">Password age</th></tr></thead>
      <tbody>
$($rows -join "`n")
      </tbody></table><div class="empty" id="empty">No user matches the filters.</div></div>
"@
    $html = New-TTHtmlPage -Title 'Password Policy' -Heading 'Password Policy Findings' -Sub "Tenant $tenant &middot; generated $genAt &middot; $total findings" `
        -BrandName $BrandName -BrandTagline $BrandTagline -KpiHtml $kpiHtml -ToolbarHtml $toolbar -BodyHtml $body

    Write-TTLog -Level INFO -Message "Password policy report created: $Path ($never never-expires, $old old)."
    $flat = $data | Select-Object User, UPN, NeverExpires, WeakAllowed, PasswordAge, Old
    Complete-TTReport -Html $html -Path $Path -Data $flat -Csv:$Csv -Excel:$Excel -DataPath $DataPath -NoHtml:$NoHtml -NoOpen:$NoOpen -Kind 'Password-Policy-Report'
    if ($PassThru) { $data }
}
