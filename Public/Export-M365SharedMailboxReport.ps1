function Export-M365SharedMailboxReport {
    <#
    .SYNOPSIS
        Generates an HTML report of shared mailboxes and who has Full Access.
    .DESCRIPTION
        Lists shared mailboxes via Exchange Online and their Full Access delegates. Read-only.
        Requires an Exchange Online connection (Connect-ExchangeOnline); if missing, the report is
        created empty with a note.
    .PARAMETER Path
        Target path of the HTML file. Default: .\SharedMailbox-Report.html
    .EXAMPLE
        Export-M365SharedMailboxReport -BrandName 'CloudNest365'
    #>
    [CmdletBinding()]
    [OutputType([pscustomobject])]
    param(
        [string]$Path = (Join-Path (Get-Location) 'SharedMailbox-Report.html'),
        [string]$BrandName = 'TenantToolbox',
        [string]$BrandTagline = 'M365 Tenant Administration',
        [switch]$Csv, [switch]$Excel, [string]$DataPath, [switch]$NoHtml, [switch]$PassThru, [switch]$NoOpen
    )

    $data = @()
    $exOk = $true
    try { Assert-TTExchange } catch { $exOk = $false; Write-Warning $_ }

    if ($exOk) {
        Write-TTLog -Level INFO -Message "Reading shared mailboxes (Exchange Online) ..."
        $boxes = Get-EXOMailbox -RecipientTypeDetails SharedMailbox -ResultSize Unlimited -ErrorAction Stop
        $data = foreach ($b in $boxes) {
            $delegates = @()
            try {
                $delegates = @(Get-EXOMailboxPermission -Identity $b.Identity -ErrorAction Stop |
                    Where-Object { $_.AccessRights -contains 'FullAccess' -and -not $_.IsInherited -and $_.User -notlike 'NT AUTHORITY*' } |
                    ForEach-Object { $_.User })
            }
            catch { }
            [pscustomobject]@{ Mailbox = $b.DisplayName; Email = $b.PrimarySmtpAddress; Delegates = $delegates }
        }
        $data = @($data)
    }

    $total = $data.Count
    $noDel = @($data | Where-Object { @($_.Delegates).Count -eq 0 }).Count
    $genAt = Get-Date -Format 'yyyy-MM-dd HH:mm'
    $tenant = try { (Get-MgContext).TenantId } catch { '' }

    $rows = foreach ($m in ($data | Sort-Object Mailbox)) {
        $fNo = if (@($m.Delegates).Count -eq 0) { '1' } else { '0' }
        $delChips = if (@($m.Delegates).Count) { (@($m.Delegates) | ForEach-Object { "<span class='chip'>$(TTEnc $_)</span>" }) -join ' ' } else { "<span class='b b-warn'>none</span>" }
        $searchAttr = TTEnc ("$($m.Mailbox) $($m.Email)".ToLower()); $nameAttr = TTEnc ([string]$m.Mailbox).ToLower()
        @"
      <tr class="item" data-f-nodelegate="$fNo" data-name="$nameAttr" data-search="$searchAttr">
        <td><div class="u"><b>$(TTEnc $m.Mailbox)</b><span class="upn">$(TTEnc $m.Email)</span></div></td>
        <td><div class="chips">$delChips</div></td>
      </tr>
"@
    }
    $kpiHtml = New-TTKpis @(
        @{ n = $total; l = 'Shared mailboxes'; filter = 'all' }
        @{ n = $noDel; l = 'Without delegate'; kind = 'warn'; filter = 'nodelegate' }
    )
    $toolbar = New-TTToolbar -SearchPlaceholder 'Search mailbox ...' -Filters @(
        @{ label = 'All'; key = 'all' }, @{ label = 'Without delegate'; key = 'nodelegate' }
    )
    $note = if (-not $exOk) { '<p class="muted" style="padding:0 0 12px">Not connected to Exchange Online. Run Connect-ExchangeOnline, then re-run this report.</p>' } else { '' }
    $body = @"
    $note
    <div class="panel"><table class="tbl"><thead><tr><th data-sort="name">Shared mailbox</th><th>Full Access delegates</th></tr></thead>
      <tbody>
$($rows -join "`n")
      </tbody></table><div class="empty" id="empty">No mailbox matches the filters.</div></div>
"@
    $html = New-TTHtmlPage -Title 'Shared Mailboxes' -Heading 'Shared Mailboxes' -Sub "Tenant $tenant &middot; generated $genAt &middot; $total shared mailboxes" `
        -BrandName $BrandName -BrandTagline $BrandTagline -KpiHtml $kpiHtml -ToolbarHtml $toolbar -BodyHtml $body

    Write-TTLog -Level INFO -Message "Shared mailbox report created: $Path ($total mailboxes)."
    $flat = $data | Select-Object Mailbox, Email, @{N = 'Delegates'; E = { $_.Delegates -join '; ' } }
    Complete-TTReport -Html $html -Path $Path -Data $flat -Csv:$Csv -Excel:$Excel -DataPath $DataPath -NoHtml:$NoHtml -NoOpen:$NoOpen -Kind 'Shared-Mailbox-Report'
    if ($PassThru) { $data }
}
