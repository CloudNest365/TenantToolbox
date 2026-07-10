# TenantToolbox root module
# Laedt alle Private- und Public-Funktionen und exportiert nur die Public-Cmdlets.

# --- Modul-Zustand (script-scoped) --------------------------------------------
$script:TTConnected = $false
$script:TTLogPath   = $null

# --- Funktionen laden ----------------------------------------------------------
$private = @( Get-ChildItem -Path (Join-Path $PSScriptRoot 'Private') -Filter '*.ps1' -ErrorAction SilentlyContinue )
$public  = @( Get-ChildItem -Path (Join-Path $PSScriptRoot 'Public')  -Filter '*.ps1' -ErrorAction SilentlyContinue )

foreach ($file in @($private + $public)) {
    try {
        . $file.FullName
    }
    catch {
        Write-Error "TenantToolbox: Konnte '$($file.FullName)' nicht laden: $_"
    }
}

Export-ModuleMember -Function $public.BaseName
