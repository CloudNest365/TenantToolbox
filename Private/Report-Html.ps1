# Gemeinsame HTML-Report-Engine fuer TenantToolbox.
# Liefert CSS + generisches JS (Suche/Filter/Sortierung/Theme/Klappen) und baut die Seite zusammen.
# Konventionen fuers JS:
#   - Ein filter-/suchbares Element traegt class="item", data-search="<lowercase>", data-name="<lowercase>".
#   - Filter-Zugehoerigkeit ueber data-f-<KEY>="1|0". Filter-Button/KPI traegt data-filter="<KEY>" ("all" = alles).
#   - Sortierbare Tabellenkoepfe: <th data-sort="KEY">. Element traegt data-name bzw. data-s-<KEY>.

function TTEnc { param($s) [System.Net.WebUtility]::HtmlEncode([string]$s) }

function Get-TTExportTarget {
    # Ermittelt Format+Pfad aus den -Csv/-Excel-Schaltern eines Report-Cmdlets. Gibt $null zurueck, wenn nichts gewuenscht.
    param([switch]$Csv, [switch]$Excel, [string]$DataPath, [Parameter(Mandatory)][string]$HtmlPath)
    if (-not ($Csv -or $Excel)) { return $null }
    $fmt = if ($Excel) { 'Excel' } else { 'Csv' }
    $ext = if ($Excel) { 'xlsx' } else { 'csv' }
    $p = if ($DataPath) { $DataPath } else { [System.IO.Path]::ChangeExtension($HtmlPath, $ext) }
    @{ Format = $fmt; Path = $p }
}

function Complete-TTReport {
    # Einheitlicher Abschluss aller Report-Cmdlets: HTML schreiben (ausser -NoHtml), optional
    # CSV/XLSX exportieren, passende Datei oeffnen (ausser -NoOpen).
    param(
        [Parameter(Mandatory)][string]$Html,
        [Parameter(Mandatory)][string]$Path,
        $Data,
        [switch]$Csv, [switch]$Excel, [string]$DataPath,
        [switch]$NoHtml, [switch]$NoOpen,
        [string]$Kind = 'Report'
    )
    $openTarget = $null
    if (-not $NoHtml) {
        $Html | Out-File -FilePath $Path -Encoding utf8
        Write-Host "$Kind erstellt: $Path" -ForegroundColor Green
        $openTarget = $Path
    }
    # -NoHtml ohne Datenformat -> mindestens CSV, sonst gaebe es keine Ausgabe.
    if ($NoHtml -and -not ($Csv -or $Excel)) { $Csv = [switch]::Present }

    $tgt = Get-TTExportTarget -Csv:$Csv -Excel:$Excel -DataPath $DataPath -HtmlPath $Path
    if ($tgt -and $null -ne $Data) {
        $written = Export-TTData -Data $Data -Path $tgt.Path -Format $tgt.Format
        Write-Host "Daten exportiert: $written" -ForegroundColor Green
        if (-not $openTarget) { $openTarget = $written }
    }
    if (-not $NoOpen -and $openTarget) { try { Invoke-Item -Path $openTarget } catch { } }
}

function Export-TTData {
    # Schreibt flache Objekte als CSV oder (falls ImportExcel vorhanden) als XLSX. Faellt sonst auf CSV zurueck.
    param(
        [Parameter(Mandatory)][object]$Data,
        [Parameter(Mandatory)][string]$Path,
        [ValidateSet('Csv', 'Excel')][string]$Format = 'Csv'
    )
    $arr = @($Data)
    $dir = Split-Path -Parent $Path
    if ($dir -and -not (Test-Path $dir)) { New-Item -ItemType Directory -Path $dir -Force | Out-Null }

    if ($Format -eq 'Excel') {
        if (Get-Module -ListAvailable -Name ImportExcel) {
            Import-Module ImportExcel -ErrorAction SilentlyContinue
            $arr | Export-Excel -Path $Path -AutoSize -BoldTopRow -FreezeTopRow -TableName 'Data' -WorksheetName 'Report' -ClearSheet
            return $Path
        }
        Write-Warning "Modul 'ImportExcel' fehlt (Install-Module ImportExcel -Scope CurrentUser) - schreibe CSV statt XLSX."
        $Path = [System.IO.Path]::ChangeExtension($Path, 'csv')
    }
    $arr | Export-Csv -Path $Path -NoTypeInformation -Encoding UTF8
    return $Path
}

function New-TTKpis {
    # $Kpis: Array von @{ n=Zahl; l=Label; kind='ok|bad|adm|report|off|on|block|mfa'; filter='KEY' }
    param([object[]]$Kpis)
    $items = foreach ($k in $Kpis) {
        $cls = "kpi"
        if ($k.kind) { $cls += " $($k.kind)" }
        $fa = if ($k.filter) { " data-filter=""$($k.filter)""" } else { "" }
        $active = if ($k.filter -eq 'all') { " active" } else { "" }
        "<div class=""$cls$active""$fa><div class=""n"">$(TTEnc $k.n)</div><div class=""l"">$(TTEnc $k.l)</div></div>"
    }
    "<div class=""kpis"">`n$($items -join "`n")`n</div>"
}

function New-TTToolbar {
    # $Filters: Array von @{ label='...'; key='KEY' }
    param([object[]]$Filters, [string]$SearchPlaceholder = 'Suchen ...', [switch]$WithToggleAll)
    $fbtns = foreach ($f in $Filters) {
        $active = if ($f.key -eq 'all') { ' active' } else { '' }
        "<button class=""fbtn$active"" data-filter=""$($f.key)"">$(TTEnc $f.label)</button>"
    }
    $toggle = if ($WithToggleAll) { '<button class="tbtn" id="toggleAll">Zuklappen</button>' } else { '' }
    @"
    <div class="toolbar">
      <div class="search">
        <svg width="16" height="16" viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="2" stroke-linecap="round"><circle cx="11" cy="11" r="7"></circle><path d="m21 21-4.3-4.3"></path></svg>
        <input id="q" type="search" placeholder="$(TTEnc $SearchPlaceholder)" autocomplete="off">
      </div>
      <div class="filters">
$($fbtns -join "`n")
      </div>
      $toggle
      <button class="tbtn" id="theme">&#9681; Design</button>
      <span class="count" id="count"></span>
    </div>
"@
}

function New-TTHtmlPage {
    param(
        [string]$Title,
        [string]$Heading,
        [string]$Sub,
        [string]$BrandName = 'TenantToolbox',
        [string]$BrandTagline = 'M365 Tenant Administration',
        [string]$KpiHtml = '',
        [string]$ToolbarHtml = '',
        [Parameter(Mandatory)][string]$BodyHtml
    )
    $bn = TTEnc $BrandName
    $bt = TTEnc $BrandTagline
    $bi = TTEnc ($BrandName.Substring(0, 1).ToUpper())
    $ti = TTEnc $Title

    @"
<!doctype html>
<html lang="de">
<head>
<meta charset="utf-8">
<meta name="viewport" content="width=device-width, initial-scale=1">
<title>$bn &middot; $ti</title>
<style>$(Get-TTReportStyle)</style>
</head>
<body>
  <div class="hero"><div class="wrap">
    <div class="brand"><div class="logo">$bi</div><div><span class="bn">$bn</span> <span class="bt">$bt</span></div></div>
    <h1>$(TTEnc $Heading)</h1>
    <p class="sub">$Sub</p>
  </div></div>
  <div class="wrap">
$KpiHtml
$ToolbarHtml
$BodyHtml
    <div class="footer"><div class="logo">$bi</div><div>Erstellt mit <b>$bn</b> &middot; $bt</div></div>
  </div>
<script>$(Get-TTReportScript)</script>
</body>
</html>
"@
}

function Get-TTReportStyle {
    @'
  :root{
    --bg:#f5f6fa; --card:#ffffff; --text:#171923; --muted:#7b8394; --border:#e7e9f0; --border-strong:#d6d9e4;
    --on:#16a34a; --off:#e5484d; --warn:#e08600; --bad:#e5484d; --crit:#e5484d; --adm:#6366f1; --info:#6366f1; --report:#e08600;
    --brand1:#4f46e5; --brand2:#7c3aed; --brand3:#db2777;
    --shadow:0 1px 2px rgba(16,24,40,.04),0 6px 20px rgba(16,24,40,.06); --shadow-lg:0 8px 30px rgba(16,24,40,.10); --radius:16px;
  }
  @media (prefers-color-scheme: dark){
    :root:not([data-theme="light"]){ --bg:#0c0e14; --card:#161922; --text:#e8eaf0; --muted:#8a90a1; --border:#242835; --border-strong:#2e3342;
      --shadow:0 1px 2px rgba(0,0,0,.4),0 8px 24px rgba(0,0,0,.35); --shadow-lg:0 12px 40px rgba(0,0,0,.5); }
  }
  :root[data-theme="dark"]{ --bg:#0c0e14; --card:#161922; --text:#e8eaf0; --muted:#8a90a1; --border:#242835; --border-strong:#2e3342;
    --shadow:0 1px 2px rgba(0,0,0,.4),0 8px 24px rgba(0,0,0,.35); --shadow-lg:0 12px 40px rgba(0,0,0,.5); }
  *{box-sizing:border-box}
  body{margin:0;background:var(--bg);color:var(--text);font-family:-apple-system,BlinkMacSystemFont,"Segoe UI",Roboto,Helvetica,Arial,sans-serif;line-height:1.5;-webkit-font-smoothing:antialiased}
  .wrap{max-width:1200px;margin:0 auto;padding:0 24px}
  .hero{position:relative;padding:30px 0;color:#fff;overflow:hidden;background:linear-gradient(120deg,var(--brand1),var(--brand2) 52%,var(--brand3))}
  .hero::after{content:"";position:absolute;inset:0;opacity:.18;background:radial-gradient(600px 200px at 85% -20%,#fff,transparent)}
  .hero .wrap{position:relative;z-index:1}
  .brand{display:flex;align-items:center;gap:11px;margin-bottom:18px}
  .brand .logo{width:34px;height:34px;border-radius:9px;background:rgba(255,255,255,.16);display:grid;place-items:center;font-weight:800;font-size:17px;border:1px solid rgba(255,255,255,.25)}
  .brand .bn{font-weight:700;font-size:15px} .brand .bt{font-size:12px;opacity:.8;margin-left:2px}
  .hero h1{margin:0;font-size:27px;font-weight:750;letter-spacing:-.022em} .hero .sub{margin:6px 0 0;opacity:.86;font-size:13px}
  .kpis{display:grid;grid-template-columns:repeat(auto-fit,minmax(148px,1fr));gap:13px;margin:-24px auto 20px;position:relative;z-index:3}
  .kpi{background:var(--card);border:1px solid var(--border);border-radius:14px;padding:15px 17px;box-shadow:var(--shadow);transition:transform .12s,box-shadow .12s,border-color .12s}
  .kpi[data-filter]{cursor:pointer} .kpi[data-filter]:hover{transform:translateY(-2px);box-shadow:var(--shadow-lg)}
  .kpi.active{border-color:var(--brand2);box-shadow:0 0 0 2px color-mix(in srgb,var(--brand2) 35%,transparent)}
  .kpi .n{font-size:29px;font-weight:780;letter-spacing:-.03em;line-height:1}
  .kpi .l{font-size:11.5px;color:var(--muted);text-transform:uppercase;letter-spacing:.05em;margin-top:5px;font-weight:600}
  .kpi.ok .n,.kpi.on .n{color:var(--on)} .kpi.bad .n,.kpi.block .n{color:var(--bad)} .kpi.warn .n,.kpi.report .n{color:var(--warn)} .kpi.adm .n,.kpi.mfa .n{color:var(--adm)} .kpi.off .n{color:var(--muted)}
  .toolbar{position:sticky;top:0;z-index:20;display:flex;flex-wrap:wrap;gap:10px;align-items:center;padding:12px 0;margin-bottom:14px;background:color-mix(in srgb,var(--bg) 88%,transparent);backdrop-filter:blur(10px);border-bottom:1px solid var(--border)}
  .search{flex:1 1 240px;position:relative}
  .search input{width:100%;padding:9px 12px 9px 34px;border-radius:10px;border:1px solid var(--border-strong);background:var(--card);color:var(--text);font-size:14px;outline:none}
  .search input:focus{border-color:var(--brand2);box-shadow:0 0 0 3px color-mix(in srgb,var(--brand2) 22%,transparent)}
  .search svg{position:absolute;left:10px;top:50%;transform:translateY(-50%);opacity:.5}
  .filters{display:flex;flex-wrap:wrap;gap:6px}
  .fbtn{font-size:12.5px;font-weight:600;padding:7px 12px;border-radius:9px;border:1px solid var(--border-strong);background:var(--card);color:var(--text);cursor:pointer}
  .fbtn:hover{border-color:var(--brand2)} .fbtn.active{background:var(--brand2);border-color:var(--brand2);color:#fff}
  select,.tbtn{font-size:13px;padding:8px 11px;border-radius:9px;border:1px solid var(--border-strong);background:var(--card);color:var(--text);cursor:pointer;outline:none}
  .tbtn{font-weight:600} .tbtn:hover{border-color:var(--brand2)}
  .count{font-size:12.5px;color:var(--muted);margin-left:auto;white-space:nowrap;font-weight:600}
  .panel{background:var(--card);border:1px solid var(--border);border-radius:16px;box-shadow:var(--shadow);overflow:hidden;margin-bottom:24px}
  table.tbl{width:100%;border-collapse:collapse;font-size:13px}
  .tbl thead th{text-align:left;padding:12px 16px;font-size:11px;text-transform:uppercase;letter-spacing:.05em;color:var(--muted);font-weight:700;border-bottom:1px solid var(--border);position:sticky;top:57px;background:var(--card);white-space:nowrap}
  .tbl thead th[data-sort]{cursor:pointer;user-select:none} .tbl thead th[data-sort]:hover{color:var(--text)}
  .tbl tbody td{padding:11px 16px;border-bottom:1px solid var(--border);vertical-align:middle}
  .tbl tbody tr:hover{background:color-mix(in srgb,var(--text) 3%,transparent)}
  .tbl tbody tr.hidden{display:none}
  .u{display:flex;flex-direction:column;gap:1px} .u b{font-weight:600} .upn{font-size:12px;color:var(--muted)}
  .b{font-size:11px;font-weight:700;padding:3px 9px;border-radius:999px;white-space:nowrap;display:inline-block}
  .b-on,.b-ok{background:color-mix(in srgb,var(--on) 15%,transparent);color:var(--on)}
  .b-off,.b-bad{background:color-mix(in srgb,var(--bad) 15%,transparent);color:var(--bad)}
  .b-warn{background:color-mix(in srgb,var(--warn) 16%,transparent);color:var(--warn)}
  .b-adm,.b-info{background:color-mix(in srgb,var(--adm) 15%,transparent);color:var(--adm)}
  .b-crit{background:var(--crit);color:#fff}
  .chips{display:flex;flex-wrap:wrap;gap:4px}
  .chip{font-size:11.5px;padding:2px 8px;border-radius:7px;background:color-mix(in srgb,var(--text) 6%,transparent);border:1px solid var(--border)}
  .chip-inc{background:color-mix(in srgb,var(--on) 11%,transparent);border-color:color-mix(in srgb,var(--on) 28%,transparent)}
  .chip-exc{background:color-mix(in srgb,var(--bad) 11%,transparent);border-color:color-mix(in srgb,var(--bad) 28%,transparent)}
  .muted{color:var(--muted)}
  .empty{display:none;text-align:center;padding:50px 20px;color:var(--muted)} .empty.show{display:block}
  .footer{border-top:1px solid var(--border);margin-top:8px;padding:22px 0 40px;color:var(--muted);font-size:12.5px;display:flex;align-items:center;gap:10px;flex-wrap:wrap}
  .footer .logo{width:24px;height:24px;border-radius:7px;color:#fff;display:grid;place-items:center;font-weight:800;font-size:12px;background:linear-gradient(135deg,var(--brand1),var(--brand3))}
  .footer b{color:var(--text)}
  /* Cards */
  .cards{display:grid;gap:16px;padding-bottom:20px}
  .card{background:var(--card);border:1px solid var(--border);border-radius:var(--radius);box-shadow:var(--shadow);overflow:hidden}
  .card.hidden{display:none}
  .card-head{padding:16px 20px;cursor:pointer;user-select:none}
  .card-head:hover{background:color-mix(in srgb,var(--text) 3%,transparent)}
  .title-wrap{display:flex;align-items:center;gap:11px;flex-wrap:wrap}
  .chevron{color:var(--muted);font-size:13px;transition:transform .18s} .card.collapsed .chevron{transform:rotate(-90deg)}
  .title-wrap h3{margin:0;font-size:16.5px;font-weight:650;flex:1 1 auto}
  .card.collapsed .card-body{display:none}
  /* Scorecard */
  .score-hero{display:flex;gap:26px;align-items:center;flex-wrap:wrap;background:var(--card);border:1px solid var(--border);border-radius:20px;box-shadow:var(--shadow);padding:26px;margin:-24px 0 22px;position:relative;z-index:3}
  .ring{--p:0;--c:var(--on);width:132px;height:132px;border-radius:50%;flex:0 0 auto;display:grid;place-items:center;
    background:conic-gradient(var(--c) calc(var(--p)*1%), color-mix(in srgb,var(--text) 8%,transparent) 0)}
  .ring .inner{width:104px;height:104px;border-radius:50%;background:var(--card);display:grid;place-items:center;flex-direction:column;text-align:center}
  .ring .grade{font-size:40px;font-weight:800;line-height:1;color:var(--c)} .ring .pct{font-size:12px;color:var(--muted);font-weight:600;margin-top:2px}
  .score-meta h2{margin:0 0 4px;font-size:20px} .score-meta p{margin:0;color:var(--muted);font-size:13px}
  .checks{display:grid;grid-template-columns:repeat(auto-fill,minmax(320px,1fr));gap:14px;padding-bottom:24px}
  .check{background:var(--card);border:1px solid var(--border);border-radius:14px;box-shadow:var(--shadow);padding:16px 18px;border-left:4px solid var(--muted)}
  .check.pass{border-left-color:var(--on)} .check.warn{border-left-color:var(--warn)} .check.fail{border-left-color:var(--bad)}
  .check .ct{display:flex;align-items:center;gap:8px;justify-content:space-between}
  .check h4{margin:0;font-size:14.5px;font-weight:650} .check .val{font-size:22px;font-weight:750;margin:8px 0 2px}
  .check.pass .val{color:var(--on)} .check.warn .val{color:var(--warn)} .check.fail .val{color:var(--bad)}
  .check .rec{font-size:12.5px;color:var(--muted)}
  @media (max-width:720px){ .ring{width:110px;height:110px} }
'@
}

function Get-TTReportScript {
    @'
(function(){
  var items=[].slice.call(document.querySelectorAll('.item'));
  var q=document.getElementById('q'), count=document.getElementById('count'), empty=document.getElementById('empty');
  var active='all';
  function matches(el){ return active==='all' || el.getAttribute('data-f-'+active)==='1'; }
  function apply(){
    var term=(q&&q.value||'').trim().toLowerCase(), vis=0;
    items.forEach(function(el){
      var s=el.getAttribute('data-search')||'';
      var ok=matches(el)&&(term===''||s.indexOf(term)!==-1);
      el.classList.toggle('hidden',!ok); if(ok)vis++;
    });
    if(count)count.textContent=vis+' von '+items.length+' sichtbar';
    if(empty)empty.classList.toggle('show',vis===0);
  }
  function setFilter(f){ active=f;
    document.querySelectorAll('.fbtn').forEach(function(b){b.classList.toggle('active',b.dataset.filter===f);});
    document.querySelectorAll('.kpi[data-filter]').forEach(function(k){k.classList.toggle('active',k.dataset.filter===f);});
    apply();
  }
  if(q)q.addEventListener('input',apply);
  document.querySelectorAll('.fbtn').forEach(function(b){b.addEventListener('click',function(){setFilter(b.dataset.filter);});});
  document.querySelectorAll('.kpi[data-filter]').forEach(function(k){k.addEventListener('click',function(){setFilter(k.dataset.filter);});});
  var sortState={};
  document.querySelectorAll('th[data-sort]').forEach(function(th){
    th.addEventListener('click',function(){
      var key=th.dataset.sort, dir=sortState[key]=!sortState[key];
      var body=th.closest('table').querySelector('tbody');
      var arr=[].slice.call(body.querySelectorAll('tr.item'));
      arr.sort(function(a,b){
        var av=key==='name'?(a.dataset.name||''):(a.getAttribute('data-s-'+key)||'');
        var bv=key==='name'?(b.dataset.name||''):(b.getAttribute('data-s-'+key)||'');
        var r=av.localeCompare(bv,undefined,{numeric:true}); return dir?r:-r;
      });
      arr.forEach(function(r){body.appendChild(r);});
    });
  });
  document.querySelectorAll('.card .card-head').forEach(function(h){
    h.setAttribute('tabindex','0');
    function t(){h.closest('.card').classList.toggle('collapsed');}
    h.addEventListener('click',t);
    h.addEventListener('keydown',function(e){if(e.key==='Enter'||e.key===' '){e.preventDefault();t();}});
  });
  var ta=document.getElementById('toggleAll');
  if(ta){var col=false;ta.addEventListener('click',function(){col=!col;document.querySelectorAll('.card').forEach(function(c){c.classList.toggle('collapsed',col);});ta.textContent=col?'Aufklappen':'Zuklappen';});}
  var themes=['auto','light','dark'],ti=0,KEY='tt_theme';
  try{var s=localStorage.getItem(KEY);if(s){ti=themes.indexOf(s);if(ti<0)ti=0;}}catch(_){}
  function applyTheme(){var t=themes[ti];
    if(t==='auto')document.documentElement.removeAttribute('data-theme');else document.documentElement.setAttribute('data-theme',t);
    var tb=document.getElementById('theme');if(tb)tb.innerHTML='&#9681; '+(t==='auto'?'Auto':(t==='light'?'Hell':'Dunkel'));
    try{localStorage.setItem(KEY,t);}catch(_){}
  }
  var tb=document.getElementById('theme');if(tb)tb.addEventListener('click',function(){ti=(ti+1)%themes.length;applyTheme();});
  applyTheme();apply();
})();
'@
}
