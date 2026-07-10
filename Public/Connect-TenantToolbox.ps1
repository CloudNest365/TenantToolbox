function Connect-TenantToolbox {
    <#
    .SYNOPSIS
        Meldet sich einmalig an Microsoft Graph an und setzt den gemeinsamen Log-Pfad.
    .DESCRIPTION
        Zentraler Auth-Einstieg fuer alle TenantToolbox-Cmdlets. Nach dem Aufruf nutzen
        alle Cmdlets dieselbe Graph-Session. Optional wird ein Log-Pfad gesetzt, in den
        veraendernde Aktionen (ACTION) protokolliert werden.
    .PARAMETER Scopes
        Graph-Berechtigungen. Standard deckt die aktuellen Cmdlets ab.
    .PARAMETER LogPath
        Pfad zu einer Log-Datei. Wird sie nicht angegeben, wird nur in die Konsole geloggt.
    .EXAMPLE
        Connect-TenantToolbox -LogPath .\tenanttoolbox.log
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

        # Robuster im eingebetteten Terminal (VS Code): kein verstecktes WAM-Fenster,
        # stattdessen ein Code zum Eingeben unter microsoft.com/devicelogin.
        [switch]$UseDeviceCode,

        # Normale OAuth-Anmeldung im Standardbrowser (mit SSO) statt WAM-Popup.
        [switch]$UseBrowser
    )

    if (-not (Get-Module -ListAvailable -Name Microsoft.Graph.Authentication)) {
        throw "Microsoft Graph SDK fehlt. Bitte einmalig: Install-Module Microsoft.Graph -Scope CurrentUser"
    }

    if ($LogPath) {
        $script:TTLogPath = $LogPath
        if (-not (Test-Path $LogPath)) { New-Item -ItemType File -Path $LogPath -Force | Out-Null }
    }

    # WAM ausschalten -> klassischer Browser-OAuth-Flow (mit SSO).
    # Parametername unterscheidet sich je nach SDK-Version -> automatisch erkennen.
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
                    Write-Warning "Diese SDK-Version bietet keinen WAM-Schalter. Nutze ggf. -UseDeviceCode."
                }
            }
            catch { Write-Warning "WAM konnte nicht deaktiviert werden: $_" }
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
        Write-TTLog -Level ERROR -Message "Anmeldung fehlgeschlagen: $_"
        Write-Host "Anmeldung fehlgeschlagen: $($_.Exception.Message)" -ForegroundColor Red
        Write-Host "Tipp: Im VS-Code-Terminal ist WAM oft im Weg - nutze:  Connect-TenantToolbox -UseDeviceCode" -ForegroundColor Yellow
        return
    }

    # Nur als verbunden melden, wenn wirklich ein Kontext existiert.
    $ctx = Get-MgContext
    if (-not $ctx -or -not $ctx.Account) {
        $script:TTConnected = $false
        Write-Host "Keine gueltige Sitzung erhalten. Nutze:  Connect-TenantToolbox -UseDeviceCode" -ForegroundColor Yellow
        return
    }

    $script:TTConnected = $true
    Write-TTLog -Level INFO -Message "Verbunden mit Tenant '$($ctx.TenantId)' als '$($ctx.Account)'."
    Write-Host "TenantToolbox verbunden: $($ctx.Account) ($($ctx.TenantId))" -ForegroundColor Green
}
