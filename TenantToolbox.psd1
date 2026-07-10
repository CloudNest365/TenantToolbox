@{
    RootModule        = 'TenantToolbox.psm1'
    ModuleVersion     = '0.2.0'
    GUID              = 'b3f1c2a4-7d6e-4c8b-9a2f-1e5d3c7a9b41'
    Author            = 'Reto Binz'
    CompanyName       = 'znib'
    Copyright         = '(c) Reto Binz. All rights reserved.'
    Description       = 'A gallery of small, sharply-scoped PowerShell cmdlets for Microsoft 365 tenant administration. Each cmdlet does exactly one job - with a unified auth, log and -WhatIf frame.'
    PowerShellVersion = '7.2'

    # Intentionally NOT enforced as RequiredModules so the module can be imported
    # even without the Graph SDK installed. The cmdlets check at runtime.
    # Required to use:
    #   Install-Module Microsoft.Graph -Scope CurrentUser
    RequiredModules   = @()

    FunctionsToExport = @(
        'Connect-TenantToolbox',
        'Get-M365StaleUsers',
        'Get-M365MfaStatus',
        'Get-M365TenantSummary',
        'Get-M365IntuneDevice',
        'Get-M365IntuneApp',
        'Get-M365IntuneAppDeployment',
        'Get-M365IntuneDeviceApp',
        'Invoke-M365Offboarding',
        'Remove-M365StaleGuests',
        'Remove-M365StaleDevices',
        'Backup-M365ConditionalAccess',
        'Compare-M365Snapshot',
        'Export-M365ConditionalAccessReport',
        'Export-M365MfaReport',
        'Export-M365AppSecretReport',
        'Export-M365SecurityScorecard',
        'Export-M365PimReport',
        'Export-M365IntuneDeviceReport',
        'Export-M365IntuneAppReport',
        'Export-M365IntuneAppDeploymentReport',
        'Export-M365IntuneDeviceAppReport'
    )
    CmdletsToExport   = @()
    VariablesToExport = @()
    AliasesToExport   = @()

    PrivateData = @{
        PSData = @{
            Tags         = @('Microsoft365','Graph','EntraID','Administration','Offboarding','TenantToolbox')
            ProjectUri   = ''
            ReleaseNotes = 'v0.2.0: 22 cmdlets incl. full Intune suite (devices, software inventory, deployment, per-device drilldown, stale-device cleanup), tenant summary, cert auth. See CHANGELOG.md.'
        }
    }
}
