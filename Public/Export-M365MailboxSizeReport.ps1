function Export-M365MailboxSizeReport {
    <#
    .SYNOPSIS
        Generates an HTML report of mailbox sizes and item counts.
    .DESCRIPTION
        Lists user mailboxes via Exchange Online with total size and item count (one statistics call
        per mailbox). Read-only. Requires an Exchange Online connection (Connect-ExchangeOnline); if
        missing, the report is created empty with a note.
    .PARAMETER Path
        Target path of the HTML file. Default: .\MailboxSize-Report.html
    .EXAMPLE
        Export-M365MailboxSizeReport -BrandName 'CloudNest365'
    #>
    [CmdletBinding()]
    [OutputType([pscustomobject])]
    param(
        [string]$Path = (Join-Path (Get-Location) 'MailboxSize-Report.html'),
        [string]$BrandName = 'TenantToolbox',
        [string]$BrandTagline = 'M365 Tenant Administration',
        [switch]$Csv, [switch]$Excel, [string]$DataPath, [switch]$NoHtml, [switch]$PassThru, [switch]$NoOpen
    )

    $data = @(); $exOk = $true
    try { Assert-TTExchange } catch { $exOk = $false; Write-Warning $_ }

    if ($exOk) {
        Write-TTLog -Level INFO -Message "Reading mailbox sizes (Exchange Online) ..."
        $boxes = Get-EXOMailbox -RecipientTypeDetails UserMailbox -ResultSize Unlimited -ErrorAction Stop
        $i = 0; $n = @($boxes).Count
        $data = foreach ($b in $boxes) {
            $i++; Write-Progress -Activity 'Reading mailbox statistics' -Status "$i / $n" -PercentComplete ($(if ($n) { 100 * $i / $n } else { 0 }))
            $sizeMB = $null; $items = $null
            try {
                $st = Get-EXOMailboxStatistics -Identity $b.Identity -ErrorAction Stop
                if ($st.TotalItemSize -and "$($st.TotalItemSize)" -match '\(([\d,]+) bytes\)') { $sizeMB = [math]::Round(([double]($Matches[1] -replace ',', '')) / 1MB, 1) }
                $items = [int]$st.ItemCount
            }
            catch { }
            [pscustomobject]@{ Mailbox = $b.DisplayName; Email = $b.PrimarySmtpAddress; SizeMB = $sizeMB; Items = $items }
        }
        Write-Progress -Activity 'Reading mailbox statistics' -Completed
        $data = @($data)
    }

    $total = $data.Count
    $sumGB = [math]::Round((($data | Measure-Object SizeMB -Sum).Sum) / 1024, 1)
    $large = @($data | Where-Object { $_.SizeMB -ge 20000 }).Count
    $genAt = Get-Date -Format 'yyyy-MM-dd HH:mm'; $tenant = try { (Get-MgContext).TenantId } catch { '' }

    $rows = foreach ($m in ($data | Sort-Object SizeMB -Descending)) {
        $fLarge = if ($m.SizeMB -ge 20000) { '1' } else { '0' }
        $sizeTxt = if ($null -ne $m.SizeMB) { if ($m.SizeMB -ge 1024) { "$([math]::Round($m.SizeMB/1024,1)) GB" } else { "$($m.SizeMB) MB" } } else { '<span class="muted">–</span>' }
        $sizeBadge = if ($m.SizeMB -ge 20000) { "<span class='b b-warn'>$sizeTxt</span>" } else { $sizeTxt }
        $sizeSort = ('{0:D9}' -f [int]([math]::Max(0, [int]$m.SizeMB)))
        $searchAttr = TTEnc ("$($m.Mailbox) $($m.Email)".ToLower()); $nameAttr = TTEnc ([string]$m.Mailbox).ToLower()
        @"
      <tr class="item" data-f-large="$fLarge" data-name="$nameAttr" data-s-size="$sizeSort" data-search="$searchAttr">
        <td><div class="u"><b>$(TTEnc $m.Mailbox)</b><span class="upn">$(TTEnc $m.Email)</span></div></td>
        <td data-s-size="$sizeSort"><b>$sizeBadge</b></td><td>$(if ($null -ne $m.Items) { $m.Items } else { '–' })</td>
      </tr>
"@
    }
    $kpiHtml = New-TTKpis @(
        @{ n = $total; l = 'Mailboxes'; filter = 'all' }
        @{ n = "$sumGB GB"; l = 'Total size'; kind = 'info' }
        @{ n = $large; l = 'Large (>=20 GB)'; kind = 'warn'; filter = 'large' }
    )
    $toolbar = New-TTToolbar -SearchPlaceholder 'Search mailbox ...' -Filters @( @{ label = 'All'; key = 'all' }, @{ label = 'Large'; key = 'large' } )
    $note = if (-not $exOk) { '<p class="muted" style="padding:0 0 12px">Not connected to Exchange Online. Run Connect-ExchangeOnline, then re-run this report.</p>' } else { '' }
    $body = @"
    $note
    <div class="panel"><table class="tbl"><thead><tr><th data-sort="name">Mailbox</th><th data-sort="size">Size</th><th>Items</th></tr></thead>
      <tbody>
$($rows -join "`n")
      </tbody></table><div class="empty" id="empty">No mailbox matches the filters.</div></div>
"@
    $html = New-TTHtmlPage -Title 'Mailbox Sizes' -Heading 'Mailbox Sizes' -Sub "Tenant $tenant &middot; generated $genAt &middot; $total mailboxes, $sumGB GB" `
        -BrandName $BrandName -BrandTagline $BrandTagline -KpiHtml $kpiHtml -ToolbarHtml $toolbar -BodyHtml $body

    Write-TTLog -Level INFO -Message "Mailbox size report created: $Path ($total mailboxes, $sumGB GB)."
    $flat = $data | Select-Object Mailbox, Email, SizeMB, Items
    Complete-TTReport -Html $html -Path $Path -Data $flat -Csv:$Csv -Excel:$Excel -DataPath $DataPath -NoHtml:$NoHtml -NoOpen:$NoOpen -Kind 'Mailbox-Size-Report'
    if ($PassThru) { $data }
}
