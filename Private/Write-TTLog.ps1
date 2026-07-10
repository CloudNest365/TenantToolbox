function Write-TTLog {
    <#
    .SYNOPSIS
        Unified logging for all TenantToolbox cmdlets.
    .DESCRIPTION
        Writes a line with timestamp and level to the console and - if a log path was
        set via Connect-TenantToolbox - additionally to the file. ACTION entries are
        meant for traceable, state-changing steps (audit).
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
        catch { Write-Warning "TenantToolbox: could not write log ($($script:TTLogPath)): $_" }
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
        Ensures a Graph connection exists - otherwise a clear error.
    #>
    [CmdletBinding()]
    param()

    if (-not (Get-Module -ListAvailable -Name Microsoft.Graph.Authentication)) {
        throw "Microsoft Graph SDK is missing. Please run once: Install-Module Microsoft.Graph -Scope CurrentUser"
    }
    $ctx = $null
    try { $ctx = Get-MgContext } catch { }
    if (-not $ctx) {
        throw "No Graph connection. Please run first: Connect-TenantToolbox"
    }
}

function Get-TTGraphCollection {
    <#
    .SYNOPSIS
        Fetches a full Graph collection (with paging) via Invoke-MgGraphRequest.
    .DESCRIPTION
        Robust and independent of which Microsoft.Graph submodules are installed -
        only needs Microsoft.Graph.Authentication. Follows @odata.nextLink to the end.
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
        Ensures an Exchange Online connection exists - otherwise a clear error.
    #>
    [CmdletBinding()]
    param()

    if (-not (Get-Command -Name Get-Mailbox -ErrorAction SilentlyContinue)) {
        throw "ExchangeOnlineManagement is missing or not connected. Please: Install-Module ExchangeOnlineManagement -Scope CurrentUser; then Connect-ExchangeOnline"
    }
    # Get-Mailbox without a session throws -> use as a connection test
    try { Get-EXOMailbox -ResultSize 1 -ErrorAction Stop | Out-Null }
    catch { throw "No active Exchange Online connection. Please: Connect-ExchangeOnline" }
}
