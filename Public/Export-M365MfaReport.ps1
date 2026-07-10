function Export-M365MfaReport {
    <#
    .SYNOPSIS
        Erzeugt einen modernen, interaktiven HTML-Report zum MFA-Status aller Benutzer.
    .DESCRIPTION
        Baut auf Get-M365MfaStatus auf und rendert eine self-contained HTML-Seite: KPI-Uebersicht
        (registriert / nicht registriert / ungeschuetzte Admins ...) und eine durchsuch-, filter-
        und sortierbare Tabelle mit Status-Badges und registrierten Methoden. Reines Lesen.
    .PARAMETER Path
        Zielpfad der HTML-Datei. Standard: .\MFA-Report.html
    .PARAMETER IncludeGuests
        Auch Gastkonten aufnehmen (Standard: nur Mitglieder).
    .PARAMETER BrandName
        Branding im Kopf/Fuss. Fuer CloudNest365: -BrandName 'CloudNest365'
    .PARAMETER BrandTagline
        Untertitel neben dem Markennamen.
    .PARAMETER PassThru
        Gibt zusaetzlich die Status-Objekte auf die Pipeline aus.
    .PARAMETER NoOpen
        Report nach dem Erzeugen NICHT im Browser oeffnen.
    .EXAMPLE
        Export-M365MfaReport -Path .\mfa.html -BrandName 'CloudNest365'
    #>
    [CmdletBinding()]
    [OutputType([pscustomobject])]
    param(
        [string]$Path = (Join-Path (Get-Location) 'MFA-Report.html'),
        [switch]$IncludeGuests,
        [string]$BrandName = 'TenantToolbox',
        [string]$BrandTagline = 'M365 Tenant Administration',
        [switch]$Csv,
        [switch]$Excel,
        [string]$DataPath,
        [switch]$NoHtml,
        [switch]$PassThru,
        [switch]$NoOpen
    )

    Assert-TTGraph
    $data = @(Get-M365MfaStatus -IncludeGuests:$IncludeGuests)

    $enc = { param($s) [System.Net.WebUtility]::HtmlEncode([string]$s) }

    $methodLabel = @{
        'microsoftAuthenticatorPush'         = 'Authenticator (Push)'
        'microsoftAuthenticatorPasswordless' = 'Authenticator (Passwordless)'
        'softwareOneTimePasscode'            = 'Software-OTP'
        'hardwareOneTimePasscode'            = 'Hardware-OTP'
        'mobilePhone'                        = 'SMS / Anruf'
        'alternateMobilePhone'               = 'Alt. Telefon'
        'officePhone'                        = 'Bürotelefon'
        'windowsHelloForBusiness'            = 'Windows Hello'
        'fido2SecurityKey'                   = 'FIDO2-Key'
        'passKeyDeviceBound'                 = 'Passkey'
        'email'                              = 'E-Mail (SSPR)'
        'securityQuestion'                   = 'Sicherheitsfragen'
        'temporaryAccessPass'                = 'Temp. Access Pass'
    }
    $defaultLabel = @{
        'push' = 'Authenticator (Push)'; 'oath' = 'OTP-Code'; 'sms' = 'SMS'; 'voice' = 'Anruf'
        'mobilePhone' = 'SMS / Anruf'; 'none' = '–'; '' = '–'
    }

    # --- KPIs ---------------------------------------------------------------
    $total    = $data.Count
    $reg      = @($data | Where-Object MfaRegistered).Count
    $unreg    = $total - $reg
    $capable  = @($data | Where-Object MfaCapable).Count
    $admins   = @($data | Where-Object IsAdmin).Count
    $adminBad = @($data | Where-Object { $_.IsAdmin -and -not $_.MfaRegistered }).Count
    $pwless   = @($data | Where-Object PasswordlessCapable).Count
    $genAt    = Get-Date -Format 'dd.MM.yyyy HH:mm'
    $tenant   = (Get-MgContext).TenantId

    # --- Tabellenzeilen -----------------------------------------------------
    $rows = foreach ($u in ($data | Sort-Object { -not $_.MfaRegistered }, DisplayName)) {
        $regFlag   = if ($u.MfaRegistered) { '1' } else { '0' }
        $adminFlag = if ($u.IsAdmin) { '1' } else { '0' }
        $pwFlag    = if ($u.PasswordlessCapable) { '1' } else { '0' }
        $typ       = if ($u.UserType -eq 'guest') { 'guest' } else { 'member' }
        $typLabel  = if ($typ -eq 'guest') { 'Gast' } else { 'Mitglied' }

        $mfaBadge = if ($u.MfaRegistered) { "<span class='b b-on'>registriert</span>" } else { "<span class='b b-off'>fehlt</span>" }
        $adminBadge = if ($u.IsAdmin) {
            if ($u.MfaRegistered) { "<span class='b b-adm'>Admin</span>" } else { "<span class='b b-crit'>Admin &#9888;</span>" }
        } else { "<span class='muted'>–</span>" }

        $defRaw = [string]$u.DefaultMethod
        $defTxt = if ($defaultLabel.ContainsKey($defRaw)) { $defaultLabel[$defRaw] } elseif ($defRaw) { $defRaw } else { '–' }

        $methodChips = if (@($u.Methods).Count) {
            (@($u.Methods) | Where-Object { $_ } | ForEach-Object {
                $lbl = if ($methodLabel.ContainsKey($_)) { $methodLabel[$_] } else { $_ }
                "<span class='chip'>$(& $enc $lbl)</span>"
            }) -join ' '
        } else { "<span class='muted'>keine</span>" }

        $searchRaw  = "$($u.DisplayName) $($u.UserPrincipalName) $defTxt $(@($u.Methods) -join ' ')"
        $searchAttr = & $enc ($searchRaw.ToLower())
        $nameAttr   = & $enc ([string]$u.DisplayName).ToLower()
        $modAttr    = & $enc ([string]$u.LastUpdated)

        @"
      <tr class="item" data-reg="$regFlag" data-admin="$adminFlag" data-type="$typ" data-pwless="$pwFlag" data-name="$nameAttr" data-modified="$modAttr" data-search="$searchAttr">
        <td><div class="u"><b>$(& $enc $u.DisplayName)</b><span class="upn">$(& $enc $u.UserPrincipalName)</span></div></td>
        <td>$typLabel</td>
        <td>$adminBadge</td>
        <td>$mfaBadge</td>
        <td>$(& $enc $defTxt)</td>
        <td><div class="chips">$methodChips</div></td>
      </tr>
"@
    }

    $brandName    = & $enc $BrandName
    $brandTagline = & $enc $BrandTagline
    $brandInitial = & $enc ($BrandName.Substring(0, 1).ToUpper())

    $html = @"
<!doctype html>
<html lang="de">
<head>
<meta charset="utf-8">
<meta name="viewport" content="width=device-width, initial-scale=1">
<title>$brandName · MFA Report</title>
<style>
  :root{
    --bg:#f5f6fa; --card:#ffffff; --text:#171923; --muted:#7b8394; --border:#e7e9f0; --border-strong:#d6d9e4;
    --on:#16a34a; --off:#e5484d; --crit:#e5484d; --adm:#6366f1; --brand1:#4f46e5; --brand2:#7c3aed; --brand3:#db2777;
    --shadow:0 1px 2px rgba(16,24,40,.04),0 6px 20px rgba(16,24,40,.06); --shadow-lg:0 8px 30px rgba(16,24,40,.10);
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
  .hero h1{margin:0;font-size:27px;font-weight:750;letter-spacing:-.022em}
  .hero .sub{margin:6px 0 0;opacity:.86;font-size:13px}
  .kpis{display:grid;grid-template-columns:repeat(auto-fit,minmax(148px,1fr));gap:13px;margin:-24px auto 20px;position:relative;z-index:3}
  .kpi{background:var(--card);border:1px solid var(--border);border-radius:14px;padding:15px 17px;box-shadow:var(--shadow);cursor:pointer;transition:transform .12s,box-shadow .12s,border-color .12s}
  .kpi:hover{transform:translateY(-2px);box-shadow:var(--shadow-lg)}
  .kpi.active{border-color:var(--brand2);box-shadow:0 0 0 2px color-mix(in srgb,var(--brand2) 35%,transparent)}
  .kpi .n{font-size:29px;font-weight:780;letter-spacing:-.03em;line-height:1}
  .kpi .l{font-size:11.5px;color:var(--muted);text-transform:uppercase;letter-spacing:.05em;margin-top:5px;font-weight:600}
  .kpi.ok .n{color:var(--on)} .kpi.bad .n{color:var(--off)} .kpi.adm .n{color:var(--adm)}
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
  .tbl thead th{text-align:left;padding:12px 16px;font-size:11px;text-transform:uppercase;letter-spacing:.05em;color:var(--muted);font-weight:700;border-bottom:1px solid var(--border);position:sticky;top:57px;background:var(--card);cursor:pointer;user-select:none;white-space:nowrap}
  .tbl thead th:hover{color:var(--text)}
  .tbl tbody td{padding:11px 16px;border-bottom:1px solid var(--border);vertical-align:middle}
  .tbl tbody tr:hover{background:color-mix(in srgb,var(--text) 3%,transparent)}
  .tbl tbody tr.hidden{display:none}
  .u{display:flex;flex-direction:column;gap:1px} .u b{font-weight:600} .upn{font-size:12px;color:var(--muted)}
  .b{font-size:11px;font-weight:700;padding:3px 9px;border-radius:999px;white-space:nowrap}
  .b-on{background:color-mix(in srgb,var(--on) 15%,transparent);color:var(--on)}
  .b-off{background:color-mix(in srgb,var(--off) 15%,transparent);color:var(--off)}
  .b-adm{background:color-mix(in srgb,var(--adm) 15%,transparent);color:var(--adm)}
  .b-crit{background:var(--crit);color:#fff}
  .chips{display:flex;flex-wrap:wrap;gap:4px}
  .chip{font-size:11.5px;padding:2px 8px;border-radius:7px;background:color-mix(in srgb,var(--text) 6%,transparent);border:1px solid var(--border)}
  .muted{color:var(--muted)}
  .empty{display:none;text-align:center;padding:50px 20px;color:var(--muted)} .empty.show{display:block}
  .footer{border-top:1px solid var(--border);margin-top:8px;padding:22px 0 40px;color:var(--muted);font-size:12.5px;display:flex;align-items:center;gap:10px;flex-wrap:wrap}
  .footer .logo{width:24px;height:24px;border-radius:7px;color:#fff;display:grid;place-items:center;font-weight:800;font-size:12px;background:linear-gradient(135deg,var(--brand1),var(--brand3))}
  .footer b{color:var(--text)}
</style>
</head>
<body>
  <div class="hero"><div class="wrap">
    <div class="brand"><div class="logo">$brandInitial</div><div><span class="bn">$brandName</span> <span class="bt">$brandTagline</span></div></div>
    <h1>MFA-Status Report</h1>
    <p class="sub">Tenant $tenant &middot; erstellt am $genAt &middot; $total Benutzer</p>
  </div></div>

  <div class="wrap">
    <div class="kpis">
      <div class="kpi active" data-filter="all"><div class="n">$total</div><div class="l">Benutzer</div></div>
      <div class="kpi ok" data-filter="reg"><div class="n">$reg</div><div class="l">MFA registriert</div></div>
      <div class="kpi bad" data-filter="unreg"><div class="n">$unreg</div><div class="l">Ohne MFA</div></div>
      <div class="kpi adm" data-filter="admin"><div class="n">$admins</div><div class="l">Admins</div></div>
      <div class="kpi bad" data-filter="adminbad"><div class="n">$adminBad</div><div class="l">Admins ohne MFA</div></div>
      <div class="kpi ok" data-filter="pwless"><div class="n">$pwless</div><div class="l">Passwordless-fähig</div></div>
    </div>

    <div class="toolbar">
      <div class="search">
        <svg width="16" height="16" viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="2" stroke-linecap="round"><circle cx="11" cy="11" r="7"></circle><path d="m21 21-4.3-4.3"></path></svg>
        <input id="q" type="search" placeholder="Benutzer suchen (Name, UPN, Methode ...)" autocomplete="off">
      </div>
      <div class="filters" id="filters">
        <button class="fbtn active" data-filter="all">Alle</button>
        <button class="fbtn" data-filter="reg">Registriert</button>
        <button class="fbtn" data-filter="unreg">Ohne MFA</button>
        <button class="fbtn" data-filter="admin">Admins</button>
        <button class="fbtn" data-filter="adminbad">Admins ohne MFA</button>
        <button class="fbtn" data-filter="pwless">Passwordless</button>
      </div>
      <button class="tbtn" id="theme">&#9681; Design</button>
      <span class="count" id="count"></span>
    </div>

    <div class="panel">
      <table class="tbl">
        <thead><tr>
          <th data-sort="name">Benutzer</th>
          <th data-sort="type">Typ</th>
          <th data-sort="admin">Admin</th>
          <th data-sort="reg">MFA</th>
          <th>Standardmethode</th>
          <th>Registrierte Methoden</th>
        </tr></thead>
        <tbody id="rows">
$($rows -join "`n")
        </tbody>
      </table>
      <div class="empty" id="empty">Kein Benutzer entspricht den Filtern.</div>
    </div>

    <div class="footer"><div class="logo">$brandInitial</div><div>Erstellt mit <b>$brandName</b> &middot; $brandTagline &middot; $genAt</div></div>
  </div>

<script>
(function(){
  var rows = Array.prototype.slice.call(document.querySelectorAll('tbody#rows tr.item'));
  var q = document.getElementById('q'), count = document.getElementById('count'), empty = document.getElementById('empty');
  var tbody = document.getElementById('rows'), activeFilter = 'all';

  function matchesFilter(r){
    switch(activeFilter){
      case 'all': return true;
      case 'reg': return r.dataset.reg === '1';
      case 'unreg': return r.dataset.reg === '0';
      case 'admin': return r.dataset.admin === '1';
      case 'adminbad': return r.dataset.admin === '1' && r.dataset.reg === '0';
      case 'pwless': return r.dataset.pwless === '1';
    }
    return true;
  }
  function apply(){
    var term = (q.value||'').trim().toLowerCase(), visible = 0;
    rows.forEach(function(r){
      var ok = matchesFilter(r) && (term==='' || r.dataset.search.indexOf(term)!==-1);
      r.classList.toggle('hidden', !ok); if(ok) visible++;
    });
    count.textContent = visible + ' von ' + rows.length + ' sichtbar';
    empty.classList.toggle('show', visible === 0);
  }
  function setFilter(f){
    activeFilter = f;
    document.querySelectorAll('#filters .fbtn').forEach(function(b){ b.classList.toggle('active', b.dataset.filter===f); });
    document.querySelectorAll('.kpi').forEach(function(k){ k.classList.toggle('active', k.dataset.filter===f); });
    apply();
  }
  q.addEventListener('input', apply);
  document.querySelectorAll('#filters .fbtn').forEach(function(b){ b.addEventListener('click', function(){ setFilter(b.dataset.filter); }); });
  document.querySelectorAll('.kpi').forEach(function(k){ k.addEventListener('click', function(){ setFilter(k.dataset.filter); }); });

  // Sortierung per Spaltenkopf
  var sortDir = {};
  document.querySelectorAll('th[data-sort]').forEach(function(th){
    th.addEventListener('click', function(){
      var key = th.dataset.sort; var dir = sortDir[key] = !sortDir[key];
      var sorted = rows.slice().sort(function(a,b){
        var av, bv;
        if(key==='name'){ av=a.dataset.name; bv=b.dataset.name; return dir?av.localeCompare(bv):bv.localeCompare(av); }
        av = a.dataset[key]||''; bv = b.dataset[key]||'';
        return dir ? av.localeCompare(bv) : bv.localeCompare(av);
      });
      sorted.forEach(function(r){ tbody.appendChild(r); });
    });
  });

  // Theme
  var themes=['auto','light','dark'], ti=0;
  try{ var s=localStorage.getItem('mfa_theme'); if(s){ ti=themes.indexOf(s); if(ti<0)ti=0; } }catch(_){}
  function applyTheme(){ var t=themes[ti];
    if(t==='auto') document.documentElement.removeAttribute('data-theme'); else document.documentElement.setAttribute('data-theme',t);
    document.getElementById('theme').innerHTML='&#9681; '+(t==='auto'?'Auto':(t==='light'?'Hell':'Dunkel'));
    try{ localStorage.setItem('mfa_theme',t); }catch(_){}
  }
  document.getElementById('theme').addEventListener('click', function(){ ti=(ti+1)%themes.length; applyTheme(); });
  applyTheme(); apply();
})();
</script>
</body>
</html>
"@

    Write-TTLog -Level INFO -Message "MFA-Report erstellt: $Path ($total Benutzer, $unreg ohne MFA, $adminBad Admins ungeschuetzt)."
    if ($adminBad -gt 0) { Write-Host "  Achtung: $adminBad Admin(s) ohne registrierte MFA!" -ForegroundColor Red }

    $flat = $data | Select-Object DisplayName, UserPrincipalName, UserType, IsAdmin, MfaRegistered, MfaCapable,
        SsprRegistered, PasswordlessCapable, DefaultMethod, @{N = 'Methods'; E = { $_.Methods -join '; ' } }, LastUpdated
    Complete-TTReport -Html $html -Path $Path -Data $flat -Csv:$Csv -Excel:$Excel -DataPath $DataPath -NoHtml:$NoHtml -NoOpen:$NoOpen -Kind 'MFA-Report'
    if ($PassThru) { $data }
}
