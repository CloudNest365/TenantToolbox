function Compare-M365Snapshot {
    <#
    .SYNOPSIS
        Vergleicht zwei CA-Snapshots (JSON) und zeigt die Aenderungen als HTML-Report.
    .DESCRIPTION
        Nimmt zwei mit Backup-M365ConditionalAccess erzeugte Snapshots (Referenz = alt,
        Difference = neu) und ermittelt hinzugefuegte, entfernte, geaenderte und unveraenderte
        Policies. Bei Aenderungen wird pro Abschnitt (State/Conditions/GrantControls/
        SessionControls) alt gegen neu gestellt. Rein lokal - kein Graph noetig.
    .PARAMETER Reference
        Pfad zum aelteren Snapshot (Ausgangszustand).
    .PARAMETER Difference
        Pfad zum neueren Snapshot (aktueller Zustand).
    .PARAMETER Path
        Zielpfad der HTML-Datei. Standard: .\Snapshot-Diff.html
    .PARAMETER BrandName
        Branding. Fuer CloudNest365: -BrandName 'CloudNest365'
    .PARAMETER PassThru
        Gibt die Diff-Objekte zusaetzlich auf die Pipeline aus.
    .PARAMETER NoOpen
        Report nicht automatisch oeffnen.
    .EXAMPLE
        Compare-M365Snapshot -Reference .\ca-alt.json -Difference .\ca-neu.json
    #>
    [CmdletBinding()]
    [OutputType([pscustomobject])]
    param(
        [Parameter(Mandatory)][string]$Reference,
        [Parameter(Mandatory)][string]$Difference,
        [string]$Path = (Join-Path (Get-Location) 'Snapshot-Diff.html'),
        [string]$BrandName = 'TenantToolbox',
        [string]$BrandTagline = 'M365 Tenant Administration',
        [switch]$Csv,
        [switch]$Excel,
        [string]$DataPath,
        [switch]$NoHtml,
        [switch]$PassThru,
        [switch]$NoOpen
    )

    foreach ($f in @($Reference, $Difference)) {
        if (-not (Test-Path $f)) { throw "Snapshot nicht gefunden: $f" }
    }

    $ref = @(Get-Content $Reference -Raw | ConvertFrom-Json)
    $dif = @(Get-Content $Difference -Raw | ConvertFrom-Json)

    $refMap = @{}; foreach ($p in $ref) { if ($p.Id) { $refMap[$p.Id] = $p } }
    $difMap = @{}; foreach ($p in $dif) { if ($p.Id) { $difMap[$p.Id] = $p } }
    $allIds = @($refMap.Keys) + @($difMap.Keys) | Select-Object -Unique

    $sections = 'State', 'Conditions', 'GrantControls', 'SessionControls'
    function Get-Json { param($o) if ($null -eq $o) { '' } else { ($o | ConvertTo-Json -Depth 25 -Compress) } }

    $results = foreach ($id in $allIds) {
        $r = $refMap[$id]; $d = $difMap[$id]
        if ($r -and -not $d) {
            [pscustomobject]@{ Id = $id; Name = $r.DisplayName; Change = 'removed'; Sections = @(); Ref = $r; Dif = $null }
        }
        elseif ($d -and -not $r) {
            [pscustomobject]@{ Id = $id; Name = $d.DisplayName; Change = 'added'; Sections = @(); Ref = $null; Dif = $d }
        }
        else {
            $changed = foreach ($s in $sections) {
                if ((Get-Json $r.$s) -ne (Get-Json $d.$s)) { $s }
            }
            $type = if ($changed) { 'changed' } else { 'unchanged' }
            [pscustomobject]@{ Id = $id; Name = $d.DisplayName; Change = $type; Sections = @($changed); Ref = $r; Dif = $d }
        }
    }

    $added   = @($results | Where-Object Change -eq 'added').Count
    $removed = @($results | Where-Object Change -eq 'removed').Count
    $changed = @($results | Where-Object Change -eq 'changed').Count
    $unch    = @($results | Where-Object Change -eq 'unchanged').Count
    $total   = $results.Count
    $genAt   = Get-Date -Format 'dd.MM.yyyy HH:mm'

    $changeBadge = @{
        added     = "<span class='b b-ok'>neu</span>"
        removed   = "<span class='b b-bad'>entfernt</span>"
        changed   = "<span class='b b-warn'>geändert</span>"
        unchanged = "<span class='b b-info'>unverändert</span>"
    }

    function Pretty { param($o) if ($null -eq $o) { '(nicht vorhanden)' } else { ($o | ConvertTo-Json -Depth 25) } }

    $cards = foreach ($c in ($results | Sort-Object @{ E = { @{added = 0; removed = 1; changed = 2; unchanged = 3 }[$_.Change] } }, Name)) {
        $detail = ''
        if ($c.Change -eq 'changed') {
            $secBlocks = foreach ($s in $c.Sections) {
                @"
          <div class="block-title">$(TTEnc $s)</div>
          <div class="row"><span class="lbl">Vorher</span><div class="vals"><pre class="diff old">$(TTEnc (Pretty $c.Ref.$s))</pre></div></div>
          <div class="row"><span class="lbl">Nachher</span><div class="vals"><pre class="diff new">$(TTEnc (Pretty $c.Dif.$s))</pre></div></div>
"@
            }
            $detail = ($secBlocks -join "`n")
        }
        elseif ($c.Change -eq 'added') { $detail = "<div class='row'><span class='lbl'>Neu</span><div class='vals'><pre class='diff new'>$(TTEnc (Pretty $c.Dif))</pre></div></div>" }
        elseif ($c.Change -eq 'removed') { $detail = "<div class='row'><span class='lbl'>Entfernt</span><div class='vals'><pre class='diff old'>$(TTEnc (Pretty $c.Ref))</pre></div></div>" }
        else { $detail = "<div class='row'><div class='vals muted'>Keine Aenderung.</div></div>" }

        $secTxt = if ($c.Sections.Count) { "Geänderte Abschnitte: " + ($c.Sections -join ', ') } else { '' }
        $searchAttr = TTEnc ("$($c.Name) $($c.Change) $($c.Sections -join ' ')".ToLower())
        $nameAttr = TTEnc ([string]$c.Name).ToLower()
        $collapsed = if ($c.Change -eq 'unchanged') { ' collapsed' } else { '' }
        @"
    <article class="card item$collapsed" data-f-$($c.Change)="1" data-name="$nameAttr" data-search="$searchAttr">
      <header class="card-head">
        <div class="title-wrap"><span class="chevron">&#9662;</span><h3>$(TTEnc $c.Name)</h3>$($changeBadge[$c.Change])</div>
        <div class="rec muted" style="padding-left:24px;margin-top:6px;font-size:12.5px">$(TTEnc $secTxt)</div>
      </header>
      <div class="card-body"><div style="padding:16px 20px">
$detail
      </div></div>
    </article>
"@
    }

    $kpiHtml = New-TTKpis @(
        @{ n = $total; l = 'Policies'; filter = 'all' }
        @{ n = $added; l = 'Neu'; kind = 'ok'; filter = 'added' }
        @{ n = $removed; l = 'Entfernt'; kind = 'bad'; filter = 'removed' }
        @{ n = $changed; l = 'Geändert'; kind = 'warn'; filter = 'changed' }
        @{ n = $unch; l = 'Unverändert'; kind = 'off'; filter = 'unchanged' }
    )
    $toolbar = New-TTToolbar -SearchPlaceholder 'Policy suchen ...' -WithToggleAll -Filters @(
        @{ label = 'Alle'; key = 'all' }, @{ label = 'Neu'; key = 'added' }, @{ label = 'Entfernt'; key = 'removed' },
        @{ label = 'Geändert'; key = 'changed' }, @{ label = 'Unverändert'; key = 'unchanged' }
    )
    $body = @"
    <div class="cards">
$($cards -join "`n")
    </div>
    <div class="empty" id="empty">Keine Policy entspricht den Filtern.</div>
    <style>.diff{margin:0;white-space:pre-wrap;word-break:break-word;font-size:11.5px;font-family:ui-monospace,SFMono-Regular,Menlo,Consolas,monospace;padding:10px 12px;border-radius:8px;max-height:280px;overflow:auto}
      .diff.old{background:color-mix(in srgb,var(--bad) 8%,transparent);border:1px solid color-mix(in srgb,var(--bad) 22%,transparent)}
      .diff.new{background:color-mix(in srgb,var(--on) 8%,transparent);border:1px solid color-mix(in srgb,var(--on) 22%,transparent)}
      .block-title{font-size:10.5px;font-weight:700;color:var(--muted);text-transform:uppercase;letter-spacing:.07em;margin:12px 0 6px}
      .row{display:flex;gap:12px;padding:4px 0;align-items:flex-start}.lbl{flex:0 0 70px;font-size:12.5px;color:var(--muted)}.vals{flex:1;min-width:0}</style>
"@

    $refName = Split-Path $Reference -Leaf
    $difName = Split-Path $Difference -Leaf
    $sub = "Referenz: $refName &rarr; Difference: $difName &middot; $genAt"
    $html = New-TTHtmlPage -Title 'Snapshot-Vergleich' -Heading 'Conditional Access - Snapshot-Vergleich' -Sub $sub `
        -BrandName $BrandName -BrandTagline $BrandTagline -KpiHtml $kpiHtml -ToolbarHtml $toolbar -BodyHtml $body

    Write-TTLog -Level INFO -Message "Snapshot-Diff erstellt: $Path (+$added / -$removed / ~$changed)."
    Write-Host "Snapshot-Vergleich: neu $added, entfernt $removed, geaendert $changed" -ForegroundColor Green

    $flat = $results | Select-Object Name, Id, Change, @{N = 'ChangedSections'; E = { $_.Sections -join '; ' } }
    Complete-TTReport -Html $html -Path $Path -Data $flat -Csv:$Csv -Excel:$Excel -DataPath $DataPath -NoHtml:$NoHtml -NoOpen:$NoOpen -Kind 'Snapshot-Vergleich'
    if ($PassThru) { $results }
}
