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

function Invoke-TTGraph {
    <#
    .SYNOPSIS
        Single Graph request with automatic retry on throttling (429) and transient errors (503/504).
    .DESCRIPTION
        Wraps Invoke-MgGraphRequest with exponential backoff. Honors a Retry-After hint when present
        in the error, otherwise backs off 2, 4, 8, ... seconds (capped). Up to -MaxRetries attempts.
    #>
    [CmdletBinding()]
    param(
        [Parameter(Mandatory)][string]$Uri,
        [string]$Method = 'GET',
        $Body,
        [int]$MaxRetries = 5
    )
    $attempt = 0
    while ($true) {
        try {
            $params = @{ Method = $Method; Uri = $Uri; OutputType = 'PSObject'; ErrorAction = 'Stop' }
            if ($null -ne $Body) { $params.Body = $Body }
            return Invoke-MgGraphRequest @params
        }
        catch {
            $msg = "$_"
            $throttled = $msg -match '429|Too Many Requests|throttl|503|504|Service Unavailable|Gateway Timeout'
            if (-not $throttled -or $attempt -ge $MaxRetries) { throw }
            $attempt++
            $wait = [math]::Min([math]::Pow(2, $attempt), 60)
            if ($msg -match 'Retry-After[:\s]+(\d+)') { $wait = [int]$Matches[1] }
            Write-TTLog -Level WARN -Message "Graph throttled/transient - retry $attempt/$MaxRetries in ${wait}s."
            Start-Sleep -Seconds $wait
        }
    }
}

function Get-TTGraphCollection {
    <#
    .SYNOPSIS
        Fetches a full Graph collection (with paging) via Invoke-MgGraphRequest.
    .DESCRIPTION
        Robust and independent of which Microsoft.Graph submodules are installed - only needs
        Microsoft.Graph.Authentication. Follows @odata.nextLink to the end, retries on throttling
        (429) and shows progress. Use -NoProgress to suppress the progress bar.
    #>
    [CmdletBinding()]
    param([Parameter(Mandatory)][string]$Uri, [string]$Activity = 'Reading from Microsoft Graph', [switch]$NoProgress)

    $items = New-Object System.Collections.Generic.List[object]
    $next = $Uri
    $page = 0
    while ($next) {
        $resp = Invoke-TTGraph -Uri $next
        if ($resp.value) { foreach ($v in $resp.value) { $items.Add($v) } }
        $next = $resp.'@odata.nextLink'
        $page++
        if (-not $NoProgress) { Write-Progress -Activity $Activity -Status "$($items.Count) items (page $page)" }
    }
    if (-not $NoProgress) { Write-Progress -Activity $Activity -Completed }
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

function Assert-TTSpo {
    <#
    .SYNOPSIS
        Ensures a SharePoint Online (SPO Management Shell) connection exists.
    #>
    [CmdletBinding()]
    param()
    if (-not (Get-Command -Name Get-SPOSite -ErrorAction SilentlyContinue)) {
        throw "Microsoft.Online.SharePoint.PowerShell is missing or not connected. Please: Install-Module Microsoft.Online.SharePoint.PowerShell; then Connect-SPOService -Url https://<tenant>-admin.sharepoint.com"
    }
    try { Get-SPOTenant -ErrorAction Stop | Out-Null }
    catch { throw "No active SharePoint Online connection. Please: Connect-SPOService -Url https://<tenant>-admin.sharepoint.com" }
}
