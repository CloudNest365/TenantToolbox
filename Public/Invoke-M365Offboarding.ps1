function Invoke-M365Offboarding {
    <#
    .SYNOPSIS
        Runs the complete leaver checklist for a user in one step.
    .DESCRIPTION
        Covers Entra, Exchange Online and OneDrive:
          1. Disable the account
          2. Revoke all sessions/tokens
          3. (optional) Set an auto-reply
          4. (optional) Convert the mailbox to a shared mailbox
          5. (optional) Delegate the mailbox to the manager (FullAccess)
          6. (optional) Grant the manager access to the user's OneDrive
          7. Remove from all groups (logged beforehand -> recoverable)
          8. (optional) Remove licenses

        Everything goes through ShouldProcess: with -WhatIf NOTHING happens (dry run). Only
        without -WhatIf are changes applied. Entra steps need Graph (Connect-TenantToolbox),
        mailbox steps need Exchange Online (Connect-ExchangeOnline), OneDrive needs
        SharePoint Online (Connect-SPOService). If a connection is missing, that step is
        cleanly skipped and reported - the rest continues.
    .PARAMETER User
        UPN or object id of the user to offboard.
    .PARAMETER Manager
        UPN of the manager for delegation/OneDrive. If omitted, the cmdlet tries to resolve
        the manager stored in Entra automatically.
    .PARAMETER ConvertToShared
        Convert the mailbox to a shared mailbox (usually no license needed afterwards).
    .PARAMETER AutoReplyMessage
        If set, enables auto-reply (internal and external) with this text.
    .PARAMETER GrantOneDriveToManager
        Grants the manager access to the user's OneDrive.
    .PARAMETER RemoveLicenses
        Remove all directly assigned licenses (runs last, after convert-to-shared).
    .EXAMPLE
        Invoke-M365Offboarding -User marta@contoso.ch -WhatIf
        Full dry run of the Entra core without changing anything.
    .EXAMPLE
        Invoke-M365Offboarding -User marta@contoso.ch -ConvertToShared -GrantOneDriveToManager `
            -AutoReplyMessage 'I am no longer with the company. Please contact info@contoso.ch.' `
            -RemoveLicenses
    #>
    [CmdletBinding(SupportsShouldProcess, ConfirmImpact = 'High')]
    param(
        [Parameter(Mandatory, ValueFromPipeline, ValueFromPipelineByPropertyName)]
        [Alias('UserPrincipalName', 'Id')]
        [string]$User,

        [string]$Manager,

        [switch]$ConvertToShared,

        [string]$AutoReplyMessage,

        [switch]$GrantOneDriveToManager,

        [switch]$RemoveLicenses
    )

    begin { Assert-TTGraph }

    process {
        # --- Resolve user ------------------------------------------------------
        try {
            $mgUser = Get-MgUser -UserId $User -Property 'id,displayName,userPrincipalName,accountEnabled,assignedLicenses' -ErrorAction Stop
        }
        catch {
            Write-TTLog -Level ERROR -Message "User '$User' not found: $_"
            return
        }

        $upn    = $mgUser.UserPrincipalName
        $result = [ordered]@{
            User = $upn; Id = $mgUser.Id; Deactivated = $false; SessionsRevoked = $false
            AutoReply = $false; ConvertedToShared = $false; MailboxDelegated = $false
            OneDriveGranted = $false; GroupsRemoved = 0; LicenseRemoved = $false
            WhatIf = [bool]$WhatIfPreference
        }
        Write-TTLog -Level INFO -Message "Starting offboarding for '$upn' ($($mgUser.Id))."

        # --- Manager (if needed for delegation/OneDrive) ----------------------
        $needsManager = $ConvertToShared -or $GrantOneDriveToManager
        if (-not $Manager -and $needsManager) {
            try {
                $mgr = Get-MgUserManager -UserId $mgUser.Id -ErrorAction Stop
                $Manager = $mgr.AdditionalProperties['userPrincipalName']
                Write-TTLog -Level INFO -Message "[$upn] Manager from Entra: '$Manager'."
            }
            catch {
                Write-TTLog -Level WARN -Message "[$upn] No manager found in Entra - delegation/OneDrive skipped."
            }
        }

        # --- 1) Disable account -----------------------------------------------
        if ($PSCmdlet.ShouldProcess($upn, 'Disable account (accountEnabled = false)')) {
            Update-MgUser -UserId $mgUser.Id -AccountEnabled:$false -ErrorAction Stop
            Write-TTLog -Level ACTION -Message "[$upn] Account disabled."
            $result.Deactivated = $true
        }

        # --- 2) Revoke sessions / tokens --------------------------------------
        if ($PSCmdlet.ShouldProcess($upn, 'Revoke all sessions and refresh tokens')) {
            Revoke-MgUserSignInSession -UserId $mgUser.Id -ErrorAction Stop | Out-Null
            Write-TTLog -Level ACTION -Message "[$upn] Sessions/tokens revoked."
            $result.SessionsRevoked = $true
        }

        # --- 3) Auto-reply (Exchange) -----------------------------------------
        if ($AutoReplyMessage) {
            if ($PSCmdlet.ShouldProcess($upn, 'Enable auto-reply')) {
                try {
                    Assert-TTExchange
                    Set-MailboxAutoReplyConfiguration -Identity $upn -AutoReplyState Enabled `
                        -InternalMessage $AutoReplyMessage -ExternalMessage $AutoReplyMessage -ErrorAction Stop
                    Write-TTLog -Level ACTION -Message "[$upn] Auto-reply enabled."
                    $result.AutoReply = $true
                }
                catch { Write-TTLog -Level WARN -Message "[$upn] Auto-reply skipped: $_" }
            }
        }

        # --- 4) Mailbox -> shared (Exchange) ----------------------------------
        if ($ConvertToShared) {
            if ($PSCmdlet.ShouldProcess($upn, 'Convert mailbox to shared mailbox')) {
                try {
                    Assert-TTExchange
                    Set-Mailbox -Identity $upn -Type Shared -ErrorAction Stop
                    Write-TTLog -Level ACTION -Message "[$upn] Mailbox converted to shared mailbox."
                    $result.ConvertedToShared = $true
                }
                catch { Write-TTLog -Level WARN -Message "[$upn] Convert to shared skipped: $_" }
            }
        }

        # --- 5) Delegate mailbox to manager (Exchange) ------------------------
        if ($Manager -and ($ConvertToShared -or $GrantOneDriveToManager)) {
            if ($PSCmdlet.ShouldProcess($upn, "Delegate mailbox to '$Manager' (FullAccess)")) {
                try {
                    Assert-TTExchange
                    Add-MailboxPermission -Identity $upn -User $Manager -AccessRights FullAccess `
                        -InheritanceType All -AutoMapping:$true -ErrorAction Stop | Out-Null
                    Write-TTLog -Level ACTION -Message "[$upn] Mailbox delegated to '$Manager'."
                    $result.MailboxDelegated = $true
                }
                catch { Write-TTLog -Level WARN -Message "[$upn] Delegation skipped: $_" }
            }
        }

        # --- 6) Grant OneDrive access to manager (SharePoint) -----------------
        if ($GrantOneDriveToManager -and $Manager) {
            if ($PSCmdlet.ShouldProcess($upn, "Grant OneDrive access to '$Manager'")) {
                try {
                    $drive = Get-MgUserDefaultDrive -UserId $mgUser.Id -ErrorAction Stop
                    $siteUrl = ($drive.WebUrl -replace '/Documents/?$', '')
                    if (Get-Command -Name Set-SPOUser -ErrorAction SilentlyContinue) {
                        Set-SPOUser -Site $siteUrl -LoginName $Manager -IsSiteCollectionAdmin $true -ErrorAction Stop | Out-Null
                        Write-TTLog -Level ACTION -Message "[$upn] OneDrive access granted to '$Manager' ($siteUrl)."
                        $result.OneDriveGranted = $true
                    }
                    else {
                        Write-TTLog -Level WARN -Message "[$upn] SPO not connected - grant OneDrive manually: $siteUrl (admin: $Manager)."
                    }
                }
                catch { Write-TTLog -Level WARN -Message "[$upn] OneDrive grant skipped: $_" }
            }
        }

        # --- 7) Remove from all groups (with log) -----------------------------
        $groups = Get-MgUserMemberOf -UserId $mgUser.Id -All -ErrorAction SilentlyContinue |
            Where-Object { $_.AdditionalProperties['@odata.type'] -eq '#microsoft.graph.group' }

        foreach ($g in $groups) {
            $gName = $g.AdditionalProperties['displayName']
            if ($PSCmdlet.ShouldProcess($upn, "Remove from group '$gName'")) {
                try {
                    Remove-MgGroupMemberByRef -GroupId $g.Id -DirectoryObjectId $mgUser.Id -ErrorAction Stop
                    Write-TTLog -Level ACTION -Message "[$upn] Removed from group '$gName' ($($g.Id))."
                    $result.GroupsRemoved++
                }
                catch {
                    Write-TTLog -Level WARN -Message "[$upn] Group '$gName' skipped (possibly dynamic/synced): $_"
                }
            }
        }

        # --- 8) Remove licenses (optional, last) ------------------------------
        if ($RemoveLicenses -and $mgUser.AssignedLicenses.Count -gt 0) {
            $skus = @($mgUser.AssignedLicenses.SkuId)
            if ($PSCmdlet.ShouldProcess($upn, "Remove licenses ($($skus.Count))")) {
                Set-MgUserLicense -UserId $mgUser.Id -AddLicenses @() -RemoveLicenses $skus -ErrorAction Stop | Out-Null
                Write-TTLog -Level ACTION -Message "[$upn] $($skus.Count) license(s) removed."
                $result.LicenseRemoved = $true
            }
        }

        Write-TTLog -Level INFO -Message "[$upn] Offboarding complete."
        [pscustomobject]$result
    }
}
