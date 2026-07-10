function Connect-TenantToolbox {
    <#
    .SYNOPSIS
        Signs in to Microsoft Graph once and sets the shared log path.
    .DESCRIPTION
        Central auth entry point for all TenantToolbox cmdlets. After calling it, all
        cmdlets reuse the same Graph session. Optionally sets a log path where
        state-changing actions (ACTION) are recorded.
    .PARAMETER Scopes
        Graph permissions. The default covers the current cmdlets.
    .PARAMETER LogPath
        Path to a log file. If omitted, logging goes to the console only.
    .PARAMETER UseDeviceCode
        More robust in embedded terminals (VS Code): no hidden WAM window, instead a
        code to enter at microsoft.com/devicelogin.
    .PARAMETER UseBrowser
        Regular OAuth sign-in in the default browser (with SSO) instead of the WAM popup.
    .PARAMETER ClientId
        App (client) id for unattended certificate authentication (app-only).
    .PARAMETER TenantId
        Tenant id for certificate authentication.
    .PARAMETER CertificateThumbprint
        Thumbprint of a certificate in the user's cert store. Together with -ClientId and
        -TenantId this signs in app-only (no interaction) - ideal for scheduled/unattended runs.
    .EXAMPLE
        Connect-TenantToolbox -LogPath .\tenanttoolbox.log
    .EXAMPLE
        Connect-TenantToolbox -UseDeviceCode
    .EXAMPLE
        Connect-TenantToolbox -ClientId <appid> -TenantId <tenantid> -CertificateThumbprint <thumb>
    #>
    [CmdletBinding(DefaultParameterSetName = 'Interactive')]
    param(
        [Parameter(ParameterSetName = 'Interactive')]
        [string[]]$Scopes = @(
            'User.ReadWrite.All',
            'Group.ReadWrite.All',
            'AuditLog.Read.All',
            'Directory.Read.All',
            'Policy.Read.All',
            'Application.Read.All',
            'UserAuthenticationMethod.Read.All',
            'RoleManagement.Read.Directory',
            'DeviceManagementManagedDevices.Read.All'
        ),

        [string]$LogPath,

        [Parameter(ParameterSetName = 'Interactive')]
        [switch]$UseDeviceCode,

        [Parameter(ParameterSetName = 'Interactive')]
        [switch]$UseBrowser,

        [Parameter(Mandatory, ParameterSetName = 'Certificate')]
        [string]$ClientId,

        [Parameter(Mandatory, ParameterSetName = 'Certificate')]
        [string]$TenantId,

        [Parameter(Mandatory, ParameterSetName = 'Certificate')]
        [string]$CertificateThumbprint
    )

    if (-not (Get-Module -ListAvailable -Name Microsoft.Graph.Authentication)) {
        throw "Microsoft Graph SDK is missing. Please run once: Install-Module Microsoft.Graph -Scope CurrentUser"
    }

    if ($LogPath) {
        $script:TTLogPath = $LogPath
        if (-not (Test-Path $LogPath)) { New-Item -ItemType File -Path $LogPath -Force | Out-Null }
    }

    # Disable WAM -> classic browser OAuth flow (with SSO).
    # The parameter name differs by SDK version -> detect automatically.
    if ($UseBrowser) {
        $opt = Get-Command Set-MgGraphOption -ErrorAction SilentlyContinue
        if ($opt) {
            try {
                if ($opt.Parameters.ContainsKey('DisableLoginByWAM')) {
                    Set-MgGraphOption -DisableLoginByWAM $true
                }
                elseif ($opt.Parameters.ContainsKey('EnableLoginByWAM')) {
                    Set-MgGraphOption -EnableLoginByWAM $false
                }
                else {
                    Write-Warning "This SDK version offers no WAM switch. Consider -UseDeviceCode."
                }
            }
            catch { Write-Warning "Could not disable WAM: $_" }
        }
    }

    try {
        if ($PSCmdlet.ParameterSetName -eq 'Certificate') {
            Connect-MgGraph -ClientId $ClientId -TenantId $TenantId -CertificateThumbprint $CertificateThumbprint -ErrorAction Stop | Out-Null
        }
        elseif ($UseDeviceCode) {
            Connect-MgGraph -Scopes $Scopes -UseDeviceCode -ErrorAction Stop | Out-Null
        }
        else {
            Connect-MgGraph -Scopes $Scopes -ErrorAction Stop | Out-Null
        }
    }
    catch {
        $script:TTConnected = $false
        Write-TTLog -Level ERROR -Message "Sign-in failed: $_"
        Write-Host "Sign-in failed: $($_.Exception.Message)" -ForegroundColor Red
        Write-Host "Tip: in the VS Code terminal WAM often gets in the way - use:  Connect-TenantToolbox -UseDeviceCode" -ForegroundColor Yellow
        return
    }

    # Only report as connected if a context really exists.
    $ctx = Get-MgContext
    $identity = if ($ctx.Account) { $ctx.Account } elseif ($ctx.AppName) { "$($ctx.AppName) (app-only)" } else { $null }
    if (-not $ctx -or -not $identity) {
        $script:TTConnected = $false
        Write-Host "No valid session obtained. Use:  Connect-TenantToolbox -UseDeviceCode" -ForegroundColor Yellow
        return
    }

    $script:TTConnected = $true
    Write-TTLog -Level INFO -Message "Connected to tenant '$($ctx.TenantId)' as '$identity'."
    Write-Host "TenantToolbox connected: $identity ($($ctx.TenantId))" -ForegroundColor Green
}
