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
    .EXAMPLE
        Connect-TenantToolbox -LogPath .\tenanttoolbox.log
    .EXAMPLE
        Connect-TenantToolbox -UseDeviceCode
    #>
    [CmdletBinding()]
    param(
        [string[]]$Scopes = @(
            'User.ReadWrite.All',
            'Group.ReadWrite.All',
            'AuditLog.Read.All',
            'Directory.Read.All',
            'Policy.Read.All',
            'Application.Read.All',
            'UserAuthenticationMethod.Read.All',
            'RoleManagement.Read.Directory'
        ),

        [string]$LogPath,

        [switch]$UseDeviceCode,

        [switch]$UseBrowser
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
        if ($UseDeviceCode) {
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
    if (-not $ctx -or -not $ctx.Account) {
        $script:TTConnected = $false
        Write-Host "No valid session obtained. Use:  Connect-TenantToolbox -UseDeviceCode" -ForegroundColor Yellow
        return
    }

    $script:TTConnected = $true
    Write-TTLog -Level INFO -Message "Connected to tenant '$($ctx.TenantId)' as '$($ctx.Account)'."
    Write-Host "TenantToolbox connected: $($ctx.Account) ($($ctx.TenantId))" -ForegroundColor Green
}
