function Get-M365MailForwarding {
    <#
    .SYNOPSIS
        Finds inbox rules that forward or redirect mail to external recipients.
    .DESCRIPTION
        Scans users' inbox message rules via Microsoft Graph and returns rules whose forward /
        redirect action targets a recipient outside the tenant's verified domains - a classic
        account-compromise indicator. Read-only. Requires MailboxSettings.Read.

        Note: scanning every mailbox makes one call per user and can be slow on large tenants.
        Use -UserId to scope to a single user.
    .PARAMETER UserId
        Restrict the scan to a single user (UPN or id).
    .EXAMPLE
        Get-M365MailForwarding
    .EXAMPLE
        Get-M365MailForwarding -UserId marta@contoso.ch
    #>
    [CmdletBinding()]
    [OutputType([pscustomobject])]
    param([string]$UserId)

    Assert-TTGraph
    Write-TTLog -Level INFO -Message "Scanning inbox rules for external forwarding ..."

    # Tenant verified domains
    $domains = @()
    try { $domains = @((Get-TTGraphCollection 'https://graph.microsoft.com/v1.0/domains?$select=id').id) } catch { }
    $domainSet = @{}; foreach ($d in $domains) { if ($d) { $domainSet[$d.ToLower()] = $true } }

    if ($UserId) {
        $users = @([pscustomobject]@{ Id = $UserId; UPN = $UserId; Name = $UserId })
    }
    else {
        $users = Get-MgUser -All -Filter "accountEnabled eq true" -Property 'id,displayName,userPrincipalName,mail' -ErrorAction Stop |
            Where-Object { $_.Mail } | Select-Object @{N = 'Id'; E = { $_.Id } }, @{N = 'UPN'; E = { $_.UserPrincipalName } }, @{N = 'Name'; E = { $_.DisplayName } }
    }

    function Get-Recips { param($arr)
        @($arr) | ForEach-Object { $_.emailAddress.address } | Where-Object { $_ }
    }
    function Test-External { param($addr)
        if (-not $addr -or $addr -notmatch '@') { return $false }
        $dom = ($addr -split '@')[-1].ToLower()
        return -not $domainSet.ContainsKey($dom)
    }

    foreach ($u in $users) {
        $rules = $null
        try { $rules = Get-TTGraphCollection "https://graph.microsoft.com/v1.0/users/$($u.Id)/mailFolders/inbox/messageRules" }
        catch { continue }

        foreach ($rule in @($rules)) {
            $act = $rule.actions
            if (-not $act) { continue }
            $recips = @()
            $action = @()
            if ($act.forwardTo) { $recips += Get-Recips $act.forwardTo; $action += 'forward' }
            if ($act.redirectTo) { $recips += Get-Recips $act.redirectTo; $action += 'redirect' }
            if ($act.forwardAsAttachmentTo) { $recips += Get-Recips $act.forwardAsAttachmentTo; $action += 'forwardAsAttachment' }
            $ext = @($recips | Where-Object { Test-External $_ } | Select-Object -Unique)
            if (-not $ext) { continue }

            [pscustomobject]@{
                User      = $u.Name
                UPN       = $u.UPN
                Rule      = $rule.displayName
                Action    = ($action -join ', ')
                External  = $ext
                Enabled   = [bool]$rule.isEnabled
            }
        }
    }
}
