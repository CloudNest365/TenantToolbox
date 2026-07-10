@{
    RootModule        = 'TenantToolbox.psm1'
    ModuleVersion     = '0.1.0'
    GUID              = 'b3f1c2a4-7d6e-4c8b-9a2f-1e5d3c7a9b41'
    Author            = 'Reto Binz'
    CompanyName       = 'znib'
    Copyright         = '(c) Reto Binz. All rights reserved.'
    Description       = 'Eine Galerie kleiner, scharf geschnittener PowerShell-Cmdlets fuer die Microsoft 365 Tenant-Administration. Jedes Cmdlet macht genau einen Job - mit einheitlichem Auth-, Log- und -WhatIf-Rahmen.'
    PowerShellVersion = '7.2'

    # Bewusst NICHT als RequiredModules erzwungen, damit das Modul auch ohne
    # installiertes Graph-SDK importiert werden kann. Die Cmdlets pruefen zur Laufzeit.
    # Zum Nutzen benoetigt:
    #   Install-Module Microsoft.Graph -Scope CurrentUser
    RequiredModules   = @()

    FunctionsToExport = @(
        'Connect-TenantToolbox',
        'Get-M365StaleUsers',
        'Get-M365MfaStatus',
        'Invoke-M365Offboarding',
        'Backup-M365ConditionalAccess',
        'Compare-M365Snapshot',
        'Export-M365ConditionalAccessReport',
        'Export-M365MfaReport',
        'Export-M365AppSecretReport',
        'Export-M365SecurityScorecard',
        'Export-M365PimReport'
    )
    CmdletsToExport   = @()
    VariablesToExport = @()
    AliasesToExport   = @()

    PrivateData = @{
        PSData = @{
            Tags         = @('Microsoft365','Graph','EntraID','Administration','Offboarding','TenantToolbox')
            ProjectUri   = ''
            ReleaseNotes = 'Erste Version: Connect-TenantToolbox, Get-M365StaleUsers, Invoke-M365Offboarding.'
        }
    }
}
