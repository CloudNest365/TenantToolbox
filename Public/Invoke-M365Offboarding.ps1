function Invoke-M365Offboarding {
    <#
    .SYNOPSIS
        Fuehrt die komplette Leaver-Checkliste fuer einen Benutzer in einem Schritt aus.
    .DESCRIPTION
        Deckt Entra, Exchange Online und OneDrive ab:
          1. Konto deaktivieren
          2. alle Sessions/Tokens widerrufen
          3. (optional) Auto-Reply setzen
          4. (optional) Postfach in Shared Mailbox umwandeln
          5. (optional) Postfach an Vorgesetzten delegieren (FullAccess)
          6. (optional) OneDrive-Zugriff an Vorgesetzten geben
          7. aus allen Gruppen entfernen (vorher protokolliert -> wiederherstellbar)
          8. (optional) Lizenzen entziehen

        Alles laeuft ueber ShouldProcess: mit -WhatIf passiert NICHTS (Dry-Run). Erst ohne
        -WhatIf wird veraendert. Entra-Schritte brauchen Graph (Connect-TenantToolbox),
        Postfach-Schritte brauchen Exchange Online (Connect-ExchangeOnline), OneDrive braucht
        SharePoint Online (Connect-SPOService). Fehlt eine Verbindung, wird der jeweilige
        Schritt sauber uebersprungen bzw. mit klarer Meldung gemeldet - der Rest laeuft weiter.
    .PARAMETER User
        UPN oder Objekt-Id des zu offboardenden Benutzers.
    .PARAMETER Manager
        UPN des Vorgesetzten fuer Delegierung/OneDrive. Wird er weggelassen, versucht das
        Cmdlet den in Entra hinterlegten Manager automatisch aufzuloesen.
    .PARAMETER ConvertToShared
        Postfach in eine Shared Mailbox umwandeln (braucht dann i.d.R. keine Lizenz mehr).
    .PARAMETER AutoReplyMessage
        Wird gesetzt, aktiviert Auto-Reply (intern und extern) mit diesem Text.
    .PARAMETER GrantOneDriveToManager
        Gibt dem Vorgesetzten Zugriff auf das OneDrive des Benutzers.
    .PARAMETER RemoveLicenses
        Alle direkt zugewiesenen Lizenzen entziehen (laeuft zuletzt, nach Convert-to-Shared).
    .EXAMPLE
        Invoke-M365Offboarding -User marta@contoso.ch -WhatIf
        Kompletter Dry-Run des Entra-Kerns, ohne etwas zu veraendern.
    .EXAMPLE
        Invoke-M365Offboarding -User marta@contoso.ch -ConvertToShared -GrantOneDriveToManager `
            -AutoReplyMessage 'Ich bin nicht mehr im Unternehmen. Bitte wenden Sie sich an info@contoso.ch.' `
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
        # --- Benutzer aufloesen ------------------------------------------------
        try {
            $mgUser = Get-MgUser -UserId $User -Property 'id,displayName,userPrincipalName,accountEnabled,assignedLicenses' -ErrorAction Stop
        }
        catch {
            Write-TTLog -Level ERROR -Message "Benutzer '$User' nicht gefunden: $_"
            return
        }

        $upn    = $mgUser.UserPrincipalName
        $result = [ordered]@{
            User = $upn; Id = $mgUser.Id; Deactivated = $false; SessionsRevoked = $false
            AutoReply = $false; ConvertedToShared = $false; MailboxDelegated = $false
            OneDriveGranted = $false; GroupsRemoved = 0; LicenseRemoved = $false
            WhatIf = [bool]$WhatIfPreference
        }
        Write-TTLog -Level INFO -Message "Starte Offboarding fuer '$upn' ($($mgUser.Id))."

        # --- Manager (falls fuer Delegierung/OneDrive noetig) ------------------
        $needsManager = $ConvertToShared -or $GrantOneDriveToManager
        if (-not $Manager -and $needsManager) {
            try {
                $mgr = Get-MgUserManager -UserId $mgUser.Id -ErrorAction Stop
                $Manager = $mgr.AdditionalProperties['userPrincipalName']
                Write-TTLog -Level INFO -Message "[$upn] Vorgesetzter aus Entra: '$Manager'."
            }
            catch {
                Write-TTLog -Level WARN -Message "[$upn] Kein Vorgesetzter in Entra gefunden - Delegierung/OneDrive werden uebersprungen."
            }
        }

        # --- 1) Konto deaktivieren --------------------------------------------
        if ($PSCmdlet.ShouldProcess($upn, 'Konto deaktivieren (accountEnabled = false)')) {
            Update-MgUser -UserId $mgUser.Id -AccountEnabled:$false -ErrorAction Stop
            Write-TTLog -Level ACTION -Message "[$upn] Konto deaktiviert."
            $result.Deactivated = $true
        }

        # --- 2) Sessions / Tokens widerrufen ----------------------------------
        if ($PSCmdlet.ShouldProcess($upn, 'Alle Sessions und Refresh-Tokens widerrufen')) {
            Revoke-MgUserSignInSession -UserId $mgUser.Id -ErrorAction Stop | Out-Null
            Write-TTLog -Level ACTION -Message "[$upn] Sessions/Tokens widerrufen."
            $result.SessionsRevoked = $true
        }

        # --- 3) Auto-Reply (Exchange) -----------------------------------------
        if ($AutoReplyMessage) {
            if ($PSCmdlet.ShouldProcess($upn, 'Auto-Reply aktivieren')) {
                try {
                    Assert-TTExchange
                    Set-MailboxAutoReplyConfiguration -Identity $upn -AutoReplyState Enabled `
                        -InternalMessage $AutoReplyMessage -ExternalMessage $AutoReplyMessage -ErrorAction Stop
                    Write-TTLog -Level ACTION -Message "[$upn] Auto-Reply aktiviert."
                    $result.AutoReply = $true
                }
                catch { Write-TTLog -Level WARN -Message "[$upn] Auto-Reply uebersprungen: $_" }
            }
        }

        # --- 4) Postfach -> Shared (Exchange) ---------------------------------
        if ($ConvertToShared) {
            if ($PSCmdlet.ShouldProcess($upn, 'Postfach in Shared Mailbox umwandeln')) {
                try {
                    Assert-TTExchange
                    Set-Mailbox -Identity $upn -Type Shared -ErrorAction Stop
                    Write-TTLog -Level ACTION -Message "[$upn] Postfach in Shared Mailbox umgewandelt."
                    $result.ConvertedToShared = $true
                }
                catch { Write-TTLog -Level WARN -Message "[$upn] Umwandlung in Shared uebersprungen: $_" }
            }
        }

        # --- 5) Postfach an Vorgesetzten delegieren (Exchange) ----------------
        if ($Manager -and ($ConvertToShared -or $GrantOneDriveToManager)) {
            if ($PSCmdlet.ShouldProcess($upn, "Postfach an '$Manager' delegieren (FullAccess)")) {
                try {
                    Assert-TTExchange
                    Add-MailboxPermission -Identity $upn -User $Manager -AccessRights FullAccess `
                        -InheritanceType All -AutoMapping:$true -ErrorAction Stop | Out-Null
                    Write-TTLog -Level ACTION -Message "[$upn] Postfach an '$Manager' delegiert."
                    $result.MailboxDelegated = $true
                }
                catch { Write-TTLog -Level WARN -Message "[$upn] Delegierung uebersprungen: $_" }
            }
        }

        # --- 6) OneDrive-Zugriff an Vorgesetzten (SharePoint) -----------------
        if ($GrantOneDriveToManager -and $Manager) {
            if ($PSCmdlet.ShouldProcess($upn, "OneDrive-Zugriff an '$Manager' geben")) {
                try {
                    $drive = Get-MgUserDefaultDrive -UserId $mgUser.Id -ErrorAction Stop
                    $siteUrl = ($drive.WebUrl -replace '/Documents/?$', '')
                    if (Get-Command -Name Set-SPOUser -ErrorAction SilentlyContinue) {
                        Set-SPOUser -Site $siteUrl -LoginName $Manager -IsSiteCollectionAdmin $true -ErrorAction Stop | Out-Null
                        Write-TTLog -Level ACTION -Message "[$upn] OneDrive-Zugriff an '$Manager' erteilt ($siteUrl)."
                        $result.OneDriveGranted = $true
                    }
                    else {
                        Write-TTLog -Level WARN -Message "[$upn] SPO nicht verbunden - OneDrive manuell freigeben: $siteUrl (Admin: $Manager)."
                    }
                }
                catch { Write-TTLog -Level WARN -Message "[$upn] OneDrive-Freigabe uebersprungen: $_" }
            }
        }

        # --- 7) Aus allen Gruppen entfernen (mit Protokoll) -------------------
        $groups = Get-MgUserMemberOf -UserId $mgUser.Id -All -ErrorAction SilentlyContinue |
            Where-Object { $_.AdditionalProperties['@odata.type'] -eq '#microsoft.graph.group' }

        foreach ($g in $groups) {
            $gName = $g.AdditionalProperties['displayName']
            if ($PSCmdlet.ShouldProcess($upn, "Aus Gruppe '$gName' entfernen")) {
                try {
                    Remove-MgGroupMemberByRef -GroupId $g.Id -DirectoryObjectId $mgUser.Id -ErrorAction Stop
                    Write-TTLog -Level ACTION -Message "[$upn] Aus Gruppe '$gName' ($($g.Id)) entfernt."
                    $result.GroupsRemoved++
                }
                catch {
                    Write-TTLog -Level WARN -Message "[$upn] Gruppe '$gName' uebersprungen (evtl. dynamisch/synchronisiert): $_"
                }
            }
        }

        # --- 8) Lizenzen entziehen (optional, zuletzt) ------------------------
        if ($RemoveLicenses -and $mgUser.AssignedLicenses.Count -gt 0) {
            $skus = @($mgUser.AssignedLicenses.SkuId)
            if ($PSCmdlet.ShouldProcess($upn, "Lizenzen entziehen ($($skus.Count) Stueck)")) {
                Set-MgUserLicense -UserId $mgUser.Id -AddLicenses @() -RemoveLicenses $skus -ErrorAction Stop | Out-Null
                Write-TTLog -Level ACTION -Message "[$upn] $($skus.Count) Lizenz(en) entzogen."
                $result.LicenseRemoved = $true
            }
        }

        Write-TTLog -Level INFO -Message "[$upn] Offboarding abgeschlossen."
        [pscustomobject]$result
    }
}
