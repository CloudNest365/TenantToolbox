# TenantToolbox root module
# Loads all Private and Public functions and exports only the Public cmdlets.

# --- Module state (script-scoped) ---------------------------------------------
$script:TTConnected  = $false
$script:TTLogPath    = $null
$script:TTMultiTenant = $null

# --- Load functions -----------------------------------------------------------
$private = @( Get-ChildItem -Path (Join-Path $PSScriptRoot 'Private') -Filter '*.ps1' -ErrorAction SilentlyContinue )
$public  = @( Get-ChildItem -Path (Join-Path $PSScriptRoot 'Public')  -Filter '*.ps1' -ErrorAction SilentlyContinue )

foreach ($file in @($private + $public)) {
    try {
        . $file.FullName
    }
    catch {
        Write-Error "TenantToolbox: could not load '$($file.FullName)': $_"
    }
}

Export-ModuleMember -Function $public.BaseName
