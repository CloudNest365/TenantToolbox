function Export-M365SharingReport {
    <#
    .SYNOPSIS
        Generates an HTML report of SharePoint sites and their external sharing setting.
    .DESCRIPTION
        Lists SharePoint Online sites with their SharingCapability (external sharing posture) and
        storage usage. Read-only. Requires a SharePoint Online connection (Connect-SPOService); if
        missing, the report is created empty with a note.
    .PARAMETER Path
        Target path of the HTML file. Default: .\Sharing-Report.html
    .EXAMPLE
        Export-M365SharingReport -BrandName 'CloudNest365'
    #>
    [CmdletBinding()]
    [OutputType([pscustomobject])]
    param(
        [string]$Path = (Join-Path (Get-Location) 'Sharing-Report.html'),
        [string]$BrandName = 'TenantToolbox',
        [string]$BrandTagline = 'M365 Tenant Administration',
        [switch]$Csv, [switch]$Excel, [string]$DataPath, [switch]$NoHtml, [switch]$PassThru, [switch]$NoOpen
    )

    $data = @(); $spoOk = $true
    try { Assert-TTSpo } catch { $spoOk = $false; Write-Warning $_ }

    $capLabel = @{
        Disabled = @{ t = 'Internal only'; c = 'b-ok' }
        ExistingExternalUserSharingOnly = @{ t = 'Existing guests'; c = 'b-info' }
        ExternalUserSharingOnly = @{ t = 'New guests'; c = 'b-warn' }
        ExternalUserAndGuestSharing = @{ t = 'Anyone (anonymous)'; c = 'b-bad' }
    }

    if ($spoOk) {
        Write-TTLog -Level INFO -Message "Reading SharePoint sites (SPO) ..."
        $sites = Get-SPOSite -Limit All -ErrorAction Stop
        $data = foreach ($s in $sites) {
            [pscustomobject]@{ Title = $s.Title; Url = $s.Url; Sharing = [string]$s.SharingCapability; StorageMB = [int]$s.StorageUsageCurrent }
        }
        $data = @($data)
    }

    $total = $data.Count
    $anon = @($data | Where-Object { $_.Sharing -eq 'ExternalUserAndGuestSharing' }).Count
    $external = @($data | Where-Object { $_.Sharing -in 'ExternalUserSharingOnly', 'ExternalUserAndGuestSharing', 'ExistingExternalUserSharingOnly' }).Count
    $genAt = Get-Date -Format 'yyyy-MM-dd HH:mm'; $tenant = try { (Get-MgContext).TenantId } catch { '' }

    $rows = foreach ($s in ($data | Sort-Object @{E = { switch ($_.Sharing) { 'ExternalUserAndGuestSharing' { 0 } 'ExternalUserSharingOnly' { 1 } 'ExistingExternalUserSharingOnly' { 2 } default { 3 } } } }, Title)) {
        $cap = if ($capLabel.ContainsKey($s.Sharing)) { $capLabel[$s.Sharing] } else { @{ t = $s.Sharing; c = 'b-info' } }
        $fAnon = if ($s.Sharing -eq 'ExternalUserAndGuestSharing') { '1' } else { '0' }
        $fExt = if ($s.Sharing -in 'ExternalUserSharingOnly', 'ExternalUserAndGuestSharing', 'ExistingExternalUserSharingOnly') { '1' } else { '0' }
        $stor = if ($s.StorageMB -ge 1024) { "$([math]::Round($s.StorageMB/1024,1)) GB" } else { "$($s.StorageMB) MB" }
        $searchAttr = TTEnc ("$($s.Title) $($s.Url)".ToLower()); $nameAttr = TTEnc ([string]$s.Title).ToLower()
        @"
      <tr class="item" data-f-anon="$fAnon" data-f-external="$fExt" data-name="$nameAttr" data-search="$searchAttr">
        <td><div class="u"><b>$(TTEnc $s.Title)</b><span class="upn">$(TTEnc $s.Url)</span></div></td>
        <td><span class="b $($cap.c)">$(TTEnc $cap.t)</span></td><td>$stor</td>
      </tr>
"@
    }
    $kpiHtml = New-TTKpis @(
        @{ n = $total; l = 'Sites'; filter = 'all' }
        @{ n = $external; l = 'External sharing'; kind = 'warn'; filter = 'external' }
        @{ n = $anon; l = 'Anonymous links'; kind = 'bad'; filter = 'anon' }
    )
    $toolbar = New-TTToolbar -SearchPlaceholder 'Search site ...' -Filters @( @{ label = 'All'; key = 'all' }, @{ label = 'External'; key = 'external' }, @{ label = 'Anonymous'; key = 'anon' } )
    $note = if (-not $spoOk) { '<p class="muted" style="padding:0 0 12px">Not connected to SharePoint Online. Run Connect-SPOService -Url https://&lt;tenant&gt;-admin.sharepoint.com, then re-run this report.</p>' } else { '' }
    $body = @"
    $note
    <div class="panel"><table class="tbl"><thead><tr><th data-sort="name">Site</th><th>External sharing</th><th>Storage</th></tr></thead>
      <tbody>
$($rows -join "`n")
      </tbody></table><div class="empty" id="empty">No site matches the filters.</div></div>
"@
    $html = New-TTHtmlPage -Title 'SharePoint Sharing' -Heading 'SharePoint External Sharing' -Sub "Tenant $tenant &middot; generated $genAt &middot; $total sites" `
        -BrandName $BrandName -BrandTagline $BrandTagline -KpiHtml $kpiHtml -ToolbarHtml $toolbar -BodyHtml $body

    Write-TTLog -Level INFO -Message "Sharing report created: $Path ($total sites, $anon anonymous)."
    $flat = $data | Select-Object Title, Url, Sharing, StorageMB
    Complete-TTReport -Html $html -Path $Path -Data $flat -Csv:$Csv -Excel:$Excel -DataPath $DataPath -NoHtml:$NoHtml -NoOpen:$NoOpen -Kind 'Sharing-Report'
    if ($PassThru) { $data }
}
