function Get-M365TenantSummary {
    <#
    .SYNOPSIS
        Prints a compact, colored console overview of the tenant's security posture.
    .DESCRIPTION
        Gathers the key signals (users, MFA coverage, admins without MFA, Conditional Access
        policies, stale accounts, expired app secrets, permanent Global Admins) and prints a
        colored summary. Also returns a structured object for further processing. Read-only.
    .PARAMETER PassThru
        Return the summary object (it is returned anyway; use to suppress nothing extra).
    .EXAMPLE
        Get-M365TenantSummary
    #>
    [CmdletBinding()]
    [OutputType([pscustomobject])]
    param()

    Assert-TTGraph
    Write-TTLog -Level INFO -Message "Building tenant summary ..."
    $tenant = (Get-MgContext).TenantId

    # MFA
    $mfa = @(); try { $mfa = @(Get-M365MfaStatus) } catch { }
    $users    = $mfa.Count
    $mfaReg   = @($mfa | Where-Object MfaRegistered).Count
    $mfaPct   = if ($users) { [math]::Round(100 * $mfaReg / $users) } else { 0 }
    $adminBad = @($mfa | Where-Object { $_.IsAdmin -and -not $_.MfaRegistered }).Count

    # Conditional Access
    $caTotal = 0; $caEnabled = 0
    try {
        $pol = @(Get-MgIdentityConditionalAccessPolicy -All -ErrorAction Stop)
        $caTotal = $pol.Count
        $caEnabled = @($pol | Where-Object State -eq 'enabled').Count
    } catch { }

    # Stale users
    $stale = 0; try { $stale = @(Get-M365StaleUsers -InactiveDays 90).Count } catch { }

    # Expired app secrets
    $expiredSecrets = 0
    try {
        $now = Get-Date
        $apps = Get-MgApplication -All -Property 'passwordCredentials,keyCredentials' -ErrorAction Stop
        foreach ($a in $apps) {
            foreach ($c in @($a.PasswordCredentials) + @($a.KeyCredentials)) {
                if ($c.EndDateTime -and ([datetime]$c.EndDateTime -lt $now)) { $expiredSecrets++ }
            }
        }
    } catch { }

    # Permanent Global Admins (PIM)
    $permGa = 0
    try {
        $base = 'https://graph.microsoft.com/v1.0/roleManagement/directory'
        $active = Get-TTGraphCollection "$base/roleAssignmentScheduleInstances?`$expand=roleDefinition&`$top=100"
        $permGa = @($active | Where-Object { $_.assignmentType -eq 'Assigned' -and $_.roleDefinition.displayName -eq 'Global Administrator' }).Count
    } catch { }

    $summary = [pscustomobject]@{
        Tenant             = $tenant
        Users              = $users
        MfaCoveragePct     = $mfaPct
        AdminsWithoutMfa   = $adminBad
        CaPoliciesTotal    = $caTotal
        CaPoliciesEnabled  = $caEnabled
        StaleUsers90d      = $stale
        ExpiredAppSecrets  = $expiredSecrets
        PermanentGlobalAdmins = $permGa
        GeneratedAt        = (Get-Date -Format 'yyyy-MM-dd HH:mm')
    }

    # --- Console output -----------------------------------------------------
    function Write-Row { param($Label, $Value, $Good)
        $color = if ($Good -eq $true) { 'Green' } elseif ($Good -eq $false) { 'Red' } else { 'Gray' }
        Write-Host ("  {0,-26}" -f $Label) -NoNewline
        Write-Host $Value -ForegroundColor $color
    }

    Write-Host ""
    Write-Host "  TenantToolbox - Tenant Summary" -ForegroundColor Cyan
    Write-Host "  Tenant $tenant  -  $($summary.GeneratedAt)" -ForegroundColor DarkGray
    Write-Host ("  " + ("-" * 44)) -ForegroundColor DarkGray
    Write-Row 'Users'                 $users                                $null
    Write-Row 'MFA coverage'          "$mfaPct %"                            ($mfaPct -ge 90)
    Write-Row 'Admins without MFA'    $adminBad                             ($adminBad -eq 0)
    Write-Row 'CA policies (enabled)' "$caEnabled / $caTotal"               ($caEnabled -gt 0)
    Write-Row 'Stale users (90d)'     $stale                                ($stale -eq 0)
    Write-Row 'Expired app secrets'   $expiredSecrets                       ($expiredSecrets -eq 0)
    Write-Row 'Permanent Global Admins' $permGa                            ($permGa -le 2)
    Write-Host ("  " + ("-" * 44)) -ForegroundColor DarkGray
    Write-Host ""

    $summary
}
