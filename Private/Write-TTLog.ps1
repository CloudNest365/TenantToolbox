function Write-TTLog {
    <#
    .SYNOPSIS
        Einheitliches Logging fuer alle TenantToolbox-Cmdlets.
    .DESCRIPTION
        Schreibt eine Zeile mit Zeitstempel und Level in die Konsole und - falls
        via Connect-TenantToolbox ein Log-Pfad gesetzt wurde - zusaetzlich in die Datei.
        ACTION-Eintraege sind fuer nachvollziehbare, veraendernde Schritte gedacht (Audit).
    #>
    [CmdletBinding()]
    param(
        [Parameter(Mandatory)]
        [string]$Message,

        [ValidateSet('INFO', 'WARN', 'ERROR', 'ACTION')]
        [string]$Level = 'INFO'
    )

    $ts   = Get-Date -Format 'yyyy-MM-dd HH:mm:ss'
    $line = "[$ts] [$Level] $Message"

    if ($script:TTLogPath) {
        try { Add-Content -Path $script:TTLogPath -Value $line -Encoding utf8 }
        catch { Write-Warning "TenantToolbox: Log konnte nicht geschrieben werden ($($script:TTLogPath)): $_" }
    }

    switch ($Level) {
        'ERROR'  { Write-Host $line -ForegroundColor Red }
        'WARN'   { Write-Host $line -ForegroundColor Yellow }
        'ACTION' { Write-Host $line -ForegroundColor Cyan }
        default  { Write-Verbose $line }
    }
}

function Assert-TTGraph {
    <#
    .SYNOPSIS
        Stellt sicher, dass eine Graph-Verbindung besteht - sonst klare Fehlermeldung.
    #>
    [CmdletBinding()]
    param()

    if (-not (Get-Module -ListAvailable -Name Microsoft.Graph.Authentication)) {
        throw "Microsoft Graph SDK fehlt. Bitte einmalig: Install-Module Microsoft.Graph -Scope CurrentUser"
    }
    $ctx = $null
    try { $ctx = Get-MgContext } catch { }
    if (-not $ctx) {
        throw "Keine Graph-Verbindung. Bitte zuerst: Connect-TenantToolbox"
    }
}

function Get-TTGraphCollection {
    <#
    .SYNOPSIS
        Holt eine komplette Graph-Collection (mit Paging) via Invoke-MgGraphRequest.
    .DESCRIPTION
        Robust und unabhaengig davon, welche Microsoft.Graph-Submodule installiert sind -
        braucht nur Microsoft.Graph.Authentication. Folgt @odata.nextLink bis zum Ende.
    #>
    [CmdletBinding()]
    param([Parameter(Mandatory)][string]$Uri)

    $items = New-Object System.Collections.Generic.List[object]
    $next = $Uri
    while ($next) {
        $resp = Invoke-MgGraphRequest -Method GET -Uri $next -OutputType PSObject -ErrorAction Stop
        if ($resp.value) { foreach ($v in $resp.value) { $items.Add($v) } }
        $next = $resp.'@odata.nextLink'
    }
    , $items
}

function Assert-TTExchange {
    <#
    .SYNOPSIS
        Stellt sicher, dass eine Exchange-Online-Verbindung besteht - sonst klare Fehlermeldung.
    #>
    [CmdletBinding()]
    param()

    if (-not (Get-Command -Name Get-Mailbox -ErrorAction SilentlyContinue)) {
        throw "ExchangeOnlineManagement fehlt oder ist nicht verbunden. Bitte: Install-Module ExchangeOnlineManagement -Scope CurrentUser; dann Connect-ExchangeOnline"
    }
    # Get-Mailbox ohne Session wirft -> als Verbindungstest nutzen
    try { Get-EXOMailbox -ResultSize 1 -ErrorAction Stop | Out-Null }
    catch { throw "Keine aktive Exchange-Online-Verbindung. Bitte: Connect-ExchangeOnline" }
}
