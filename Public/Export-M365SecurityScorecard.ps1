function Export-M365SecurityScorecard {
    <#
    .SYNOPSIS
        Generates a Security Scorecard (grade A-F) as HTML - the executive one-pager.
    .DESCRIPTION
        Bundles several signals (MFA coverage, unprotected admins, Conditional Access
        baseline, legacy-auth block, stale accounts, expired app secrets) into an overall
        grade and shows each check with status and recommendation. Read-only.
    .PARAMETER Path
        Target path of the HTML file. Default: .\Security-Scorecard.html
    .PARAMETER BrandName
        Branding. For CloudNest365: -BrandName 'CloudNest365'
    .PARAMETER PassThru
        Return the check results as objects.
    .PARAMETER NoOpen
        Do not open the report automatically.
    .EXAMPLE
        Export-M365SecurityScorecard -BrandName 'CloudNest365'
    #>
    [CmdletBinding()]
    [OutputType([pscustomobject])]
    param(
        [string]$Path = (Join-Path (Get-Location) 'Security-Scorecard.html'),
        [string]$BrandName = 'TenantToolbox',
        [string]$BrandTagline = 'M365 Tenant Administration',
        [switch]$Csv,
        [switch]$Excel,
        [string]$DataPath,
        [switch]$NoHtml,
        [switch]$PassThru,
        [switch]$NoOpen
    )

    Assert-TTGraph
    Write-TTLog -Level INFO -Message "Collecting signals for the Security Scorecard ..."

    $checks = @()
    function New-Check { param($Name, $Status, $Value, $Rec)
        $score = switch ($Status) { 'pass' { 1 } 'warn' { 0.5 } default { 0 } }
        [pscustomobject]@{ Name = $Name; Status = $Status; Value = $Value; Rec = $Rec; Score = $score }
    }

    # 1) MFA coverage
    try {
        $mfa = @(Get-M365MfaStatus)
        $mfaTotal = $mfa.Count
        $mfaReg = @($mfa | Where-Object MfaRegistered).Count
        $pct = if ($mfaTotal) { [math]::Round(100 * $mfaReg / $mfaTotal) } else { 0 }
        $st = if ($pct -ge 90) { 'pass' } elseif ($pct -ge 70) { 'warn' } else { 'fail' }
        $checks += New-Check 'MFA coverage' $st "$pct%" "$mfaReg of $mfaTotal users have MFA registered. Target: >= 90%."

        # 2) Admins protected
        $adminBad = @($mfa | Where-Object { $_.IsAdmin -and -not $_.MfaRegistered }).Count
        $st = if ($adminBad -eq 0) { 'pass' } else { 'fail' }
        $checks += New-Check 'Admins with MFA' $st "$adminBad without" $(if ($adminBad) { "$adminBad privileged accounts without MFA - secure immediately." } else { 'All admin accounts have MFA. Great.' })
    }
    catch { Write-TTLog -Level WARN -Message "MFA signal skipped: $_" }

    # 3+4) Conditional Access
    try {
        $pol = @(Get-MgIdentityConditionalAccessPolicy -All -ErrorAction Stop)
        $enabled = @($pol | Where-Object { $_.State -eq 'enabled' })
        $mfaBaseline = @($enabled | Where-Object { $_.GrantControls.BuiltInControls -contains 'mfa' }).Count -gt 0
        $st = if ($mfaBaseline) { 'pass' } else { 'fail' }
        $checks += New-Check 'CA: MFA baseline' $st $(if ($mfaBaseline) { 'present' } else { 'missing' }) $(if ($mfaBaseline) { 'At least one enabled policy enforces MFA.' } else { 'No enabled CA policy enforces MFA.' })

        $legacy = @($enabled | Where-Object { ($_.Conditions.ClientAppTypes | Where-Object { $_ -in 'exchangeActiveSync', 'other' }) -and ($_.GrantControls.BuiltInControls -contains 'block') }).Count -gt 0
        $st = if ($legacy) { 'pass' } else { 'fail' }
        $checks += New-Check 'CA: Legacy auth blocked' $st $(if ($legacy) { 'yes' } else { 'no' }) $(if ($legacy) { 'Legacy authentication is blocked.' } else { 'Legacy auth is not blocked - major attack surface.' })
    }
    catch { Write-TTLog -Level WARN -Message "CA signal skipped: $_" }

    # 5) Stale accounts
    try {
        $stale = @(Get-M365StaleUsers -InactiveDays 90)
        $sc = $stale.Count
        $st = if ($sc -eq 0) { 'pass' } elseif ($sc -le 5) { 'warn' } else { 'fail' }
        $checks += New-Check 'Stale accounts (90 days)' $st "$sc" $(if ($sc) { "$sc enabled accounts without sign-in for 90 days - review/disable." } else { 'No stale enabled accounts.' })
    }
    catch { Write-TTLog -Level WARN -Message "Stale signal skipped: $_" }

    # 6) App secret hygiene
    try {
        $now = Get-Date
        $apps = Get-MgApplication -All -Property 'displayName,passwordCredentials,keyCredentials' -ErrorAction Stop
        $expired = 0
        foreach ($a in $apps) {
            foreach ($c in @($a.PasswordCredentials) + @($a.KeyCredentials)) {
                if ($c.EndDateTime -and ([datetime]$c.EndDateTime -lt $now)) { $expired++ }
            }
        }
        $st = if ($expired -eq 0) { 'pass' } else { 'fail' }
        $checks += New-Check 'App secret hygiene' $st "$expired expired" $(if ($expired) { "$expired expired secrets/certificates - rotate or remove." } else { 'No expired app secrets.' })
    }
    catch { Write-TTLog -Level WARN -Message "App secret signal skipped: $_" }

    # 7) Permanent Global Admins (PIM)
    try {
        $base = 'https://graph.microsoft.com/v1.0/roleManagement/directory'
        $active = Get-TTGraphCollection "$base/roleAssignmentScheduleInstances?`$expand=roleDefinition&`$top=100"
        $permGa = @($active | Where-Object { $_.assignmentType -eq 'Assigned' -and $_.roleDefinition.displayName -eq 'Global Administrator' }).Count
        $st = if ($permGa -le 2) { 'pass' } elseif ($permGa -le 5) { 'warn' } else { 'fail' }
        $checks += New-Check 'Permanent Global Admins' $st "$permGa" $(if ($permGa -le 2) { 'Few standing Global Admins (break-glass). Good.' } else { "$permGa permanent Global Admins - prefer eligible (PIM/JIT) assignments." })
    }
    catch { Write-TTLog -Level WARN -Message "PIM signal skipped: $_" }

    # --- Overall grade -------------------------------------------------------
    $overall = if ($checks.Count) { [math]::Round(100 * (($checks | Measure-Object Score -Sum).Sum) / $checks.Count) } else { 0 }
    $grade = if ($overall -ge 90) { 'A' } elseif ($overall -ge 80) { 'B' } elseif ($overall -ge 70) { 'C' } elseif ($overall -ge 60) { 'D' } else { 'F' }
    $gradeColor = if ($overall -ge 80) { 'var(--on)' } elseif ($overall -ge 60) { 'var(--warn)' } else { 'var(--bad)' }
    $genAt  = Get-Date -Format 'yyyy-MM-dd HH:mm'
    $tenant = (Get-MgContext).TenantId

    $statusBadge = @{ pass = "<span class='b b-ok'>pass</span>"; warn = "<span class='b b-warn'>warning</span>"; fail = "<span class='b b-bad'>critical</span>" }

    $checkCards = foreach ($c in ($checks | Sort-Object Score, Name)) {
        @"
      <div class="check $($c.Status)">
        <div class="ct"><h4>$(TTEnc $c.Name)</h4>$($statusBadge[$c.Status])</div>
        <div class="val">$(TTEnc $c.Value)</div>
        <div class="rec">$(TTEnc $c.Rec)</div>
      </div>
"@
    }

    $body = @"
    <div class="score-hero">
      <div class="ring" style="--p:$overall;--c:$gradeColor">
        <div class="inner"><div class="grade" style="color:$gradeColor">$grade</div><div class="pct">$overall / 100</div></div>
      </div>
      <div class="score-meta">
        <h2>Security score: $overall %</h2>
        <p>$($checks.Count) checks &middot; Tenant $tenant &middot; $genAt</p>
      </div>
    </div>
    <div class="checks">
$($checkCards -join "`n")
    </div>
"@

    $sub = "Tenant $tenant &middot; generated $genAt"
    $html = New-TTHtmlPage -Title 'Security Scorecard' -Heading 'Security Scorecard' -Sub $sub `
        -BrandName $BrandName -BrandTagline $BrandTagline -BodyHtml $body

    Write-TTLog -Level INFO -Message "Security Scorecard created: $Path (grade $grade, $overall %)."
    Write-Host "Overall grade: $grade ($overall %)" -ForegroundColor Green

    $flat = @([pscustomobject]@{ Name = 'OVERALL'; Status = "Grade $grade"; Value = "$overall %"; Rec = "$($checks.Count) checks"; Score = '' }) +
            ($checks | Select-Object Name, Status, Value, Rec, Score)
    Complete-TTReport -Html $html -Path $Path -Data $flat -Csv:$Csv -Excel:$Excel -DataPath $DataPath -NoHtml:$NoHtml -NoOpen:$NoOpen -Kind 'Security-Scorecard'
    if ($PassThru) { $checks }
}
