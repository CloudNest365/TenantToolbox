function Disable-M365ExternalForwarding {
    <#
    .SYNOPSIS
        Disables inbox rules that forward mail to external recipients (remediation for the
        mail-forwarding report).
    .DESCRIPTION
        Finds enabled inbox rules that forward/redirect to external recipients (via
        Get-M365MailForwarding) and disables them via Microsoft Graph. Every change goes through
        ShouldProcess: with -WhatIf NOTHING is changed (dry run). Returns one object per rule.
        Requires MailboxSettings.ReadWrite.
    .PARAMETER UserId
        Restrict to a single user (UPN or id).
    .PARAMETER WhatIf
        Dry run - only show what would be disabled.
    .EXAMPLE
        Disable-M365ExternalForwarding -WhatIf
    .EXAMPLE
        Disable-M365ExternalForwarding -UserId marta@contoso.ch
    #>
    [CmdletBinding(SupportsShouldProcess, ConfirmImpact = 'High')]
    [OutputType([pscustomobject])]
    param([string]$UserId)

    Assert-TTGraph
    $rules = @(Get-M365MailForwarding -UserId $UserId | Where-Object Enabled)
    Write-TTLog -Level INFO -Message "Found $($rules.Count) enabled external-forwarding rule(s)."

    foreach ($r in $rules) {
        $done = $false
        if ($PSCmdlet.ShouldProcess("$($r.UPN): rule '$($r.Rule)'", "Disable external-forwarding rule (to: $(@($r.External) -join ', '))")) {
            try {
                Invoke-MgGraphRequest -Method PATCH -Uri "https://graph.microsoft.com/v1.0/users/$($r.UserId)/mailFolders/inbox/messageRules/$($r.RuleId)" -Body (@{ isEnabled = $false } | ConvertTo-Json) -ErrorAction Stop | Out-Null
                Write-TTLog -Level ACTION -Message "Disabled forwarding rule '$($r.Rule)' for '$($r.UPN)'."
                $done = $true
            }
            catch { Write-TTLog -Level WARN -Message "Could not disable rule for '$($r.UPN)': $_" }
        }
        [pscustomobject]@{ User = $r.User; UPN = $r.UPN; Rule = $r.Rule; External = $r.External; Disabled = $done; WhatIf = [bool]$WhatIfPreference }
    }
}
