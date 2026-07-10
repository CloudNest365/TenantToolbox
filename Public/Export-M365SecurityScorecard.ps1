function Export-M365SecurityScorecard {
    <#
    .SYNOPSIS
        Erzeugt eine Security-Scorecard (Note A-F) als HTML - der Executive-One-Pager.
    .DESCRIPTION
        Buendelt mehrere Signale (MFA-Abdeckung, ungeschuetzte Admins, Conditional-Access-
        Baseline, Legacy-Auth-Blockade, inaktive Konten, abgelaufene App-Secrets) zu einer
        Gesamtnote und zeigt jede Pruefung mit Status und Empfehlung. Reines Lesen.
    .PARAMETER Path
        Zielpfad der HTML-Datei. Standard: .\Security-Scorecard.html
    .PARAMETER BrandName
        Branding. Fuer CloudNest365: -BrandName 'CloudNest365'
    .PARAMETER PassThru
        Gibt die Pruef-Ergebnisse als Objekte zurueck.
    .PARAMETER NoOpen
        Report nicht automatisch oeffnen.
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
    Write-TTLog -Level INFO -Message "Sammle Signale fuer die Security-Scorecard ..."

    $checks = @()
    function New-Check { param($Name, $Status, $Value, $Rec)
        $score = switch ($Status) { 'pass' { 1 } 'warn' { 0.5 } default { 0 } }
        [pscustomobject]@{ Name = $Name; Status = $Status; Value = $Value; Rec = $Rec; Score = $score }
    }

    # 1) MFA-Abdeckung
    try {
        $mfa = @(Get-M365MfaStatus)
        $mfaTotal = $mfa.Count
        $mfaReg = @($mfa | Where-Object MfaRegistered).Count
        $pct = if ($mfaTotal) { [math]::Round(100 * $mfaReg / $mfaTotal) } else { 0 }
        $st = if ($pct -ge 90) { 'pass' } elseif ($pct -ge 70) { 'warn' } else { 'fail' }
        $checks += New-Check 'MFA-Abdeckung' $st "$pct%" "$mfaReg von $mfaTotal Benutzern haben MFA registriert. Ziel: >= 90%."

        # 2) Admins geschuetzt
        $adminBad = @($mfa | Where-Object { $_.IsAdmin -and -not $_.MfaRegistered }).Count
        $st = if ($adminBad -eq 0) { 'pass' } else { 'fail' }
        $checks += New-Check 'Admins mit MFA' $st "$adminBad ohne" $(if ($adminBad) { "$adminBad privilegierte Konten ohne MFA - sofort absichern." } else { 'Alle Admin-Konten haben MFA. Sehr gut.' })
    }
    catch { Write-TTLog -Level WARN -Message "MFA-Signal uebersprungen: $_" }

    # 3+4) Conditional Access
    try {
        $pol = @(Get-MgIdentityConditionalAccessPolicy -All -ErrorAction Stop)
        $enabled = @($pol | Where-Object { $_.State -eq 'enabled' })
        $mfaBaseline = @($enabled | Where-Object { $_.GrantControls.BuiltInControls -contains 'mfa' }).Count -gt 0
        $st = if ($mfaBaseline) { 'pass' } else { 'fail' }
        $checks += New-Check 'CA: MFA-Baseline' $st $(if ($mfaBaseline) { 'vorhanden' } else { 'fehlt' }) $(if ($mfaBaseline) { 'Mindestens eine aktive Policy erzwingt MFA.' } else { 'Keine aktive CA-Policy erzwingt MFA.' })

        $legacy = @($enabled | Where-Object { ($_.Conditions.ClientAppTypes | Where-Object { $_ -in 'exchangeActiveSync', 'other' }) -and ($_.GrantControls.BuiltInControls -contains 'block') }).Count -gt 0
        $st = if ($legacy) { 'pass' } else { 'fail' }
        $checks += New-Check 'CA: Legacy-Auth blockiert' $st $(if ($legacy) { 'ja' } else { 'nein' }) $(if ($legacy) { 'Legacy-Authentifizierung wird blockiert.' } else { 'Legacy-Auth ist nicht blockiert - grosses Angriffsrisiko.' })
    }
    catch { Write-TTLog -Level WARN -Message "CA-Signal uebersprungen: $_" }

    # 5) Inaktive Konten
    try {
        $stale = @(Get-M365StaleUsers -InactiveDays 90)
        $sc = $stale.Count
        $st = if ($sc -eq 0) { 'pass' } elseif ($sc -le 5) { 'warn' } else { 'fail' }
        $checks += New-Check 'Inaktive Konten (90 Tage)' $st "$sc" $(if ($sc) { "$sc aktive Konten ohne Anmeldung seit 90 Tagen - pruefen/deaktivieren." } else { 'Keine inaktiven aktiven Konten.' })
    }
    catch { Write-TTLog -Level WARN -Message "Stale-Signal uebersprungen: $_" }

    # 6) App-Secret-Hygiene
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
        $checks += New-Check 'App-Secret-Hygiene' $st "$expired abgelaufen" $(if ($expired) { "$expired abgelaufene Secrets/Zertifikate - rotieren oder entfernen." } else { 'Keine abgelaufenen App-Secrets.' })
    }
    catch { Write-TTLog -Level WARN -Message "App-Secret-Signal uebersprungen: $_" }

    # --- Gesamtnote ----------------------------------------------------------
    $overall = if ($checks.Count) { [math]::Round(100 * (($checks | Measure-Object Score -Sum).Sum) / $checks.Count) } else { 0 }
    $grade = if ($overall -ge 90) { 'A' } elseif ($overall -ge 80) { 'B' } elseif ($overall -ge 70) { 'C' } elseif ($overall -ge 60) { 'D' } else { 'F' }
    $gradeColor = if ($overall -ge 80) { 'var(--on)' } elseif ($overall -ge 60) { 'var(--warn)' } else { 'var(--bad)' }
    $genAt  = Get-Date -Format 'dd.MM.yyyy HH:mm'
    $tenant = (Get-MgContext).TenantId

    $statusBadge = @{ pass = "<span class='b b-ok'>bestanden</span>"; warn = "<span class='b b-warn'>Achtung</span>"; fail = "<span class='b b-bad'>kritisch</span>" }

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
        <h2>Sicherheits-Score: $overall %</h2>
        <p>$($checks.Count) Pruefungen &middot; Tenant $tenant &middot; $genAt</p>
      </div>
    </div>
    <div class="checks">
$($checkCards -join "`n")
    </div>
"@

    $sub = "Tenant $tenant &middot; erstellt am $genAt"
    $html = New-TTHtmlPage -Title 'Security Scorecard' -Heading 'Security Scorecard' -Sub $sub `
        -BrandName $BrandName -BrandTagline $BrandTagline -BodyHtml $body

    Write-TTLog -Level INFO -Message "Security-Scorecard erstellt: $Path (Note $grade, $overall %)."
    Write-Host "Gesamtnote: $grade ($overall %)" -ForegroundColor Green

    $flat = @([pscustomobject]@{ Name = 'GESAMT'; Status = "Note $grade"; Value = "$overall %"; Rec = "$($checks.Count) Pruefungen"; Score = '' }) +
            ($checks | Select-Object Name, Status, Value, Rec, Score)
    Complete-TTReport -Html $html -Path $Path -Data $flat -Csv:$Csv -Excel:$Excel -DataPath $DataPath -NoHtml:$NoHtml -NoOpen:$NoOpen -Kind 'Security-Scorecard'
    if ($PassThru) { $checks }
}
