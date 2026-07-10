function Compare-M365Snapshot {
    <#
    .SYNOPSIS
        Compares two CA snapshots (JSON) and shows the changes as an HTML report.
    .DESCRIPTION
        Takes two snapshots created with Backup-M365ConditionalAccess (Reference = old,
        Difference = new) and determines added, removed, changed and unchanged policies.
        For changes, each section (State/Conditions/GrantControls/SessionControls) is shown
        old vs. new. Fully local - no Graph needed.
    .PARAMETER Reference
        Path to the older snapshot (baseline).
    .PARAMETER Difference
        Path to the newer snapshot (current state).
    .PARAMETER Path
        Target path of the HTML file. Default: .\Snapshot-Diff.html
    .PARAMETER BrandName
        Branding. For CloudNest365: -BrandName 'CloudNest365'
    .PARAMETER PassThru
        Also emit the diff objects on the pipeline.
    .PARAMETER NoOpen
        Do not open the report automatically.
    .EXAMPLE
        Compare-M365Snapshot -Reference .\ca-old.json -Difference .\ca-new.json
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
        if (-not (Test-Path $f)) { throw "Snapshot not found: $f" }
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
    $genAt   = Get-Date -Format 'yyyy-MM-dd HH:mm'

    $changeBadge = @{
        added     = "<span class='b b-ok'>new</span>"
        removed   = "<span class='b b-bad'>removed</span>"
        changed   = "<span class='b b-warn'>changed</span>"
        unchanged = "<span class='b b-info'>unchanged</span>"
    }

    function Pretty { param($o) if ($null -eq $o) { '(not present)' } else { ($o | ConvertTo-Json -Depth 25) } }

    $cards = foreach ($c in ($results | Sort-Object @{ E = { @{added = 0; removed = 1; changed = 2; unchanged = 3 }[$_.Change] } }, Name)) {
        $detail = ''
        if ($c.Change -eq 'changed') {
            $secBlocks = foreach ($s in $c.Sections) {
                @"
          <div class="block-title">$(TTEnc $s)</div>
          <div class="row"><span class="lbl">Before</span><div class="vals"><pre class="diff old">$(TTEnc (Pretty $c.Ref.$s))</pre></div></div>
          <div class="row"><span class="lbl">After</span><div class="vals"><pre class="diff new">$(TTEnc (Pretty $c.Dif.$s))</pre></div></div>
"@
            }
            $detail = ($secBlocks -join "`n")
        }
        elseif ($c.Change -eq 'added') { $detail = "<div class='row'><span class='lbl'>New</span><div class='vals'><pre class='diff new'>$(TTEnc (Pretty $c.Dif))</pre></div></div>" }
        elseif ($c.Change -eq 'removed') { $detail = "<div class='row'><span class='lbl'>Removed</span><div class='vals'><pre class='diff old'>$(TTEnc (Pretty $c.Ref))</pre></div></div>" }
        else { $detail = "<div class='row'><div class='vals muted'>No change.</div></div>" }

        $secTxt = if ($c.Sections.Count) { "Changed sections: " + ($c.Sections -join ', ') } else { '' }
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
        @{ n = $added; l = 'New'; kind = 'ok'; filter = 'added' }
        @{ n = $removed; l = 'Removed'; kind = 'bad'; filter = 'removed' }
        @{ n = $changed; l = 'Changed'; kind = 'warn'; filter = 'changed' }
        @{ n = $unch; l = 'Unchanged'; kind = 'off'; filter = 'unchanged' }
    )
    $toolbar = New-TTToolbar -SearchPlaceholder 'Search policy ...' -WithToggleAll -Filters @(
        @{ label = 'All'; key = 'all' }, @{ label = 'New'; key = 'added' }, @{ label = 'Removed'; key = 'removed' },
        @{ label = 'Changed'; key = 'changed' }, @{ label = 'Unchanged'; key = 'unchanged' }
    )
    $body = @"
    <div class="cards">
$($cards -join "`n")
    </div>
    <div class="empty" id="empty">No policy matches the filters.</div>
    <style>.diff{margin:0;white-space:pre-wrap;word-break:break-word;font-size:11.5px;font-family:ui-monospace,SFMono-Regular,Menlo,Consolas,monospace;padding:10px 12px;border-radius:8px;max-height:280px;overflow:auto}
      .diff.old{background:color-mix(in srgb,var(--bad) 8%,transparent);border:1px solid color-mix(in srgb,var(--bad) 22%,transparent)}
      .diff.new{background:color-mix(in srgb,var(--on) 8%,transparent);border:1px solid color-mix(in srgb,var(--on) 22%,transparent)}
      .block-title{font-size:10.5px;font-weight:700;color:var(--muted);text-transform:uppercase;letter-spacing:.07em;margin:12px 0 6px}
      .row{display:flex;gap:12px;padding:4px 0;align-items:flex-start}.lbl{flex:0 0 70px;font-size:12.5px;color:var(--muted)}.vals{flex:1;min-width:0}</style>
"@

    $refName = Split-Path $Reference -Leaf
    $difName = Split-Path $Difference -Leaf
    $sub = "Reference: $refName &rarr; Difference: $difName &middot; $genAt"
    $html = New-TTHtmlPage -Title 'Snapshot Comparison' -Heading 'Conditional Access - Snapshot Comparison' -Sub $sub `
        -BrandName $BrandName -BrandTagline $BrandTagline -KpiHtml $kpiHtml -ToolbarHtml $toolbar -BodyHtml $body

    Write-TTLog -Level INFO -Message "Snapshot diff created: $Path (+$added / -$removed / ~$changed)."
    Write-Host "Snapshot comparison: new $added, removed $removed, changed $changed" -ForegroundColor Green

    $flat = $results | Select-Object Name, Id, Change, @{N = 'ChangedSections'; E = { $_.Sections -join '; ' } }
    Complete-TTReport -Html $html -Path $Path -Data $flat -Csv:$Csv -Excel:$Excel -DataPath $DataPath -NoHtml:$NoHtml -NoOpen:$NoOpen -Kind 'Snapshot-Comparison'
    if ($PassThru) { $results }
}
