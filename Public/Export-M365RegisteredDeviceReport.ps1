function Export-M365RegisteredDeviceReport {
    <#
    .SYNOPSIS
        Generates an HTML report of Entra-registered/joined devices (not Intune).
    .DESCRIPTION
        Reads directory devices via Graph and shows OS, trust type (Entra joined / registered /
        hybrid), managed/compliant state and last activity. Read-only.
    .PARAMETER StaleDays
        Flag devices with no activity for this many days. Default: 90.
    .PARAMETER Path
        Target path of the HTML file. Default: .\RegisteredDevice-Report.html
    .EXAMPLE
        Export-M365RegisteredDeviceReport -BrandName 'CloudNest365'
    #>
    [CmdletBinding()]
    [OutputType([pscustomobject])]
    param(
        [int]$StaleDays = 90,
        [string]$Path = (Join-Path (Get-Location) 'RegisteredDevice-Report.html'),
        [string]$BrandName = 'TenantToolbox',
        [string]$BrandTagline = 'M365 Tenant Administration',
        [switch]$Csv, [switch]$Excel, [string]$DataPath, [switch]$NoHtml, [switch]$PassThru, [switch]$NoOpen
    )

    Assert-TTGraph
    Write-TTLog -Level INFO -Message "Reading registered devices ..."
    $devs = Get-TTGraphCollection "https://graph.microsoft.com/v1.0/devices?`$select=displayName,operatingSystem,operatingSystemVersion,trustType,isCompliant,isManaged,approximateLastSignInDateTime,accountEnabled&`$top=100"

    $trustLabel = @{ AzureAd = 'Entra joined'; Workplace = 'Registered'; ServerAd = 'Hybrid joined' }
    $cutoff = (Get-Date).AddDays(-$StaleDays)
    $data = foreach ($d in $devs) {
        $last = $d.approximateLastSignInDateTime
        $days = if ($last) { [math]::Floor(((Get-Date) - [datetime]$last).TotalDays) } else { $null }
        $stale = ($null -ne $last) -and ([datetime]$last -lt $cutoff)
        [pscustomobject]@{
            Device = $d.displayName; OS = "$($d.operatingSystem) $($d.operatingSystemVersion)"
            Trust = if ($trustLabel.ContainsKey([string]$d.trustType)) { $trustLabel[[string]$d.trustType] } else { [string]$d.trustType }
            Managed = [bool]$d.isManaged; Compliant = [bool]$d.isCompliant; LastSignIn = if ($last) { [datetime]$last } else { $null }
            DaysSince = $days; Stale = $stale; Enabled = [bool]$d.accountEnabled
        }
    }
    $data = @($data)
    $total = $data.Count
    $unmanaged = @($data | Where-Object { -not $_.Managed }).Count
    $stale = @($data | Where-Object Stale).Count
    $disabled = @($data | Where-Object { -not $_.Enabled }).Count
    $genAt = Get-Date -Format 'yyyy-MM-dd HH:mm'; $tenant = (Get-MgContext).TenantId

    $rows = foreach ($d in ($data | Sort-Object @{E = { -[int][bool]$_.Stale } }, Device)) {
        $fUn = if (-not $d.Managed) { '1' } else { '0' }
        $fStale = if ($d.Stale) { '1' } else { '0' }
        $man = if ($d.Managed) { "<span class='b b-ok'>managed</span>" } else { "<span class='muted'>no</span>" }
        $syncTxt = if ($d.LastSignIn) { "$($d.LastSignIn.ToString('yyyy-MM-dd')) ($($d.DaysSince)d)" } else { '<span class="muted">–</span>' }
        $syncSort = if ($d.LastSignIn) { $d.LastSignIn.ToString('yyyy-MM-dd') } else { '0000-00-00' }
        $searchAttr = TTEnc ("$($d.Device) $($d.OS) $($d.Trust)".ToLower()); $nameAttr = TTEnc ([string]$d.Device).ToLower()
        @"
      <tr class="item" data-f-unmanaged="$fUn" data-f-stale="$fStale" data-name="$nameAttr" data-s-sync="$syncSort" data-search="$searchAttr">
        <td><b>$(TTEnc $d.Device)</b></td><td>$(TTEnc $d.OS)</td><td>$(TTEnc $d.Trust)</td><td>$man</td><td data-s-sync="$syncSort">$syncTxt</td>
      </tr>
"@
    }
    $kpiHtml = New-TTKpis @(
        @{ n = $total; l = 'Devices'; filter = 'all' }
        @{ n = $unmanaged; l = 'Unmanaged'; kind = 'warn'; filter = 'unmanaged' }
        @{ n = $stale; l = "Stale (>$StaleDays d)"; kind = 'bad'; filter = 'stale' }
        @{ n = $disabled; l = 'Disabled'; kind = 'info' }
    )
    $toolbar = New-TTToolbar -SearchPlaceholder 'Search device or OS ...' -Filters @(
        @{ label = 'All'; key = 'all' }, @{ label = 'Unmanaged'; key = 'unmanaged' }, @{ label = 'Stale'; key = 'stale' }
    )
    $body = @"
    <div class="panel"><table class="tbl"><thead><tr><th data-sort="name">Device</th><th>OS</th><th>Trust type</th><th>Managed</th><th data-sort="sync">Last activity</th></tr></thead>
      <tbody>
$($rows -join "`n")
      </tbody></table><div class="empty" id="empty">No device matches the filters.</div></div>
"@
    $html = New-TTHtmlPage -Title 'Registered Devices' -Heading 'Entra Registered Devices' -Sub "Tenant $tenant &middot; generated $genAt &middot; $total devices" `
        -BrandName $BrandName -BrandTagline $BrandTagline -KpiHtml $kpiHtml -ToolbarHtml $toolbar -BodyHtml $body

    Write-TTLog -Level INFO -Message "Registered device report created: $Path ($total devices, $stale stale, $unmanaged unmanaged)."
    $flat = $data | Select-Object Device, OS, Trust, Managed, Compliant,
        @{N = 'LastSignIn'; E = { if ($_.LastSignIn) { $_.LastSignIn.ToString('yyyy-MM-dd') } } }, DaysSince, Stale, Enabled
    Complete-TTReport -Html $html -Path $Path -Data $flat -Csv:$Csv -Excel:$Excel -DataPath $DataPath -NoHtml:$NoHtml -NoOpen:$NoOpen -Kind 'Registered-Device-Report'
    if ($PassThru) { $data }
}
