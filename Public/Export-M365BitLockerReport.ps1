function Export-M365BitLockerReport {
    <#
    .SYNOPSIS
        Generates an HTML report of BitLocker recovery-key escrow coverage.
    .DESCRIPTION
        Reads BitLocker recovery keys via Graph (informationProtection/bitlocker/recoveryKeys) and
        shows which devices have escrowed keys, volume type and creation date (the key value itself
        is NOT retrieved). Read-only. Requires BitlockerKey.Read.All.
    .PARAMETER Path
        Target path of the HTML file. Default: .\BitLocker-Report.html
    .EXAMPLE
        Export-M365BitLockerReport -BrandName 'CloudNest365'
    #>
    [CmdletBinding()]
    [OutputType([pscustomobject])]
    param(
        [string]$Path = (Join-Path (Get-Location) 'BitLocker-Report.html'),
        [string]$BrandName = 'TenantToolbox',
        [string]$BrandTagline = 'M365 Tenant Administration',
        [switch]$Csv, [switch]$Excel, [string]$DataPath, [switch]$NoHtml, [switch]$PassThru, [switch]$NoOpen
    )

    Assert-TTGraph
    Write-TTLog -Level INFO -Message "Reading BitLocker recovery keys ..."
    try { $keys = Get-TTGraphCollection "https://graph.microsoft.com/v1.0/informationProtection/bitlocker/recoveryKeys?`$top=100" }
    catch {
        if ("$_" -match 'Forbidden|403') { Write-Warning "Access denied. Needs BitlockerKey.Read.All. Reconnect: Connect-TenantToolbox -UseDeviceCode" }
        else { Write-Warning "Could not read BitLocker keys: $_" }
        $keys = @()
    }

    $data = foreach ($k in $keys) {
        [pscustomobject]@{ DeviceId = $k.deviceId; KeyId = $k.id; VolumeType = $k.volumeType; Created = $k.createdDateTime }
    }
    $data = @($data)
    $total = $data.Count
    $devices = @($data | Select-Object -ExpandProperty DeviceId -Unique | Where-Object { $_ }).Count
    $os = @($data | Where-Object { $_.VolumeType -eq 'operatingSystemVolume' }).Count
    $genAt = Get-Date -Format 'yyyy-MM-dd HH:mm'; $tenant = (Get-MgContext).TenantId

    $rows = foreach ($k in ($data | Sort-Object Created -Descending)) {
        $created = if ($k.Created) { ([datetime]$k.Created).ToString('yyyy-MM-dd') } else { '–' }
        $searchAttr = TTEnc ("$($k.DeviceId) $($k.VolumeType)".ToLower())
        @"
      <tr class="item" data-name="$searchAttr" data-search="$searchAttr">
        <td><b>$(TTEnc $k.DeviceId)</b></td><td>$(TTEnc $k.VolumeType)</td><td><span class="b b-ok">escrowed</span></td><td>$created</td>
      </tr>
"@
    }
    $kpiHtml = New-TTKpis @(
        @{ n = $total; l = 'Recovery keys'; filter = 'all' }
        @{ n = $devices; l = 'Devices covered'; kind = 'ok' }
        @{ n = $os; l = 'OS volumes'; kind = 'info' }
    )
    $toolbar = New-TTToolbar -SearchPlaceholder 'Search device id ...' -Filters @( @{ label = 'All'; key = 'all' } )
    $body = @"
    <div class="panel"><table class="tbl"><thead><tr><th data-sort="name">Device id</th><th>Volume</th><th>Key</th><th>Created</th></tr></thead>
      <tbody>
$($rows -join "`n")
      </tbody></table><div class="empty" id="empty">No key matches the search.</div></div>
"@
    $html = New-TTHtmlPage -Title 'BitLocker' -Heading 'BitLocker Recovery Coverage' -Sub "Tenant $tenant &middot; generated $genAt &middot; $total keys on $devices devices" `
        -BrandName $BrandName -BrandTagline $BrandTagline -KpiHtml $kpiHtml -ToolbarHtml $toolbar -BodyHtml $body

    Write-TTLog -Level INFO -Message "BitLocker report created: $Path ($total keys, $devices devices)."
    $flat = $data | Select-Object DeviceId, VolumeType, @{N = 'Created'; E = { if ($_.Created) { ([datetime]$_.Created).ToString('yyyy-MM-dd') } } }
    Complete-TTReport -Html $html -Path $Path -Data $flat -Csv:$Csv -Excel:$Excel -DataPath $DataPath -NoHtml:$NoHtml -NoOpen:$NoOpen -Kind 'BitLocker-Report'
    if ($PassThru) { $data }
}
