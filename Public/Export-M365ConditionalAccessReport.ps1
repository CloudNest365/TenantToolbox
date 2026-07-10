function Export-M365ConditionalAccessReport {
    <#
    .SYNOPSIS
        Erzeugt einen modernen, self-contained HTML-Report aller Conditional-Access-Policies.
    .DESCRIPTION
        Liest alle CA-Policies ueber Microsoft Graph, loest die GUIDs zu lesbaren Namen auf
        (Benutzer, Gruppen, Rollen, Apps, Named Locations) und rendert einen huebschen
        HTML-Report: KPI-Uebersicht, eine Karte pro Policy mit Assignments und Controls sowie
        automatisch berechnete Impact-/Hinweis-Pills (z. B. "Blockiert Zugriff", "Erzwingt MFA",
        "Gilt fuer ALLE Benutzer", "N Ausnahmen", "Report-only").

        Der Report ist eine einzelne HTML-Datei (Inline-CSS, Light/Dark, keine externen Abhaengigkeiten).
        Reines Lesen - veraendert nichts.
    .PARAMETER Path
        Zielpfad der HTML-Datei. Standard: .\ConditionalAccess-Report.html
    .PARAMETER PassThru
        Gibt zusaetzlich die strukturierten Policy-Objekte auf die Pipeline aus.
    .PARAMETER NoOpen
        Oeffnet den Report nach dem Erzeugen NICHT im Standardbrowser.
    .EXAMPLE
        Export-M365ConditionalAccessReport
    .EXAMPLE
        Export-M365ConditionalAccessReport -Path C:\Reports\ca.html -PassThru
    #>
    [CmdletBinding()]
    [OutputType([pscustomobject])]
    param(
        [string]$Path = (Join-Path (Get-Location) 'ConditionalAccess-Report.html'),

        # Branding im Kopf/Fuss des Reports. Fuer CloudNest365 einfach: -BrandName 'CloudNest365'
        [string]$BrandName = 'TenantToolbox',
        [string]$BrandTagline = 'M365 Tenant Administration',

        # Zusaetzlich die Rohdaten als CSV bzw. XLSX exportieren (XLSX braucht Modul ImportExcel).
        [switch]$Csv,
        [switch]$Excel,
        [string]$DataPath,
        [switch]$NoHtml,

        [switch]$PassThru,
        [switch]$NoOpen
    )

    Assert-TTGraph
    Write-TTLog -Level INFO -Message "Lese Conditional-Access-Policies ..."

    # --- Rohdaten + Lookups --------------------------------------------------
    $policies = Get-MgIdentityConditionalAccessPolicy -All -ErrorAction Stop

    $locMap = @{}
    try { Get-MgIdentityConditionalAccessNamedLocation -All -ErrorAction Stop | ForEach-Object { $locMap[$_.Id] = $_.DisplayName } } catch { }

    $roleMap = @{}
    try { Get-MgDirectoryRoleTemplate -All -ErrorAction Stop | ForEach-Object { $roleMap[$_.Id] = $_.DisplayName } } catch { }

    $userCache = @{}; $groupCache = @{}; $appCache = @{}

    $enc = { param($s) [System.Net.WebUtility]::HtmlEncode([string]$s) }

    function Get-Cached { param($Cache, $Id, [scriptblock]$Fetch)
        if ([string]::IsNullOrWhiteSpace($Id)) { return $Id }
        if ($Cache.ContainsKey($Id)) { return $Cache[$Id] }
        $name = $Id
        try { $r = & $Fetch $Id; if ($r) { $name = $r } } catch { }
        $Cache[$Id] = $name; return $name
    }

    function Resolve-Users { param($Ids)
        foreach ($id in @($Ids)) {
            if ([string]::IsNullOrWhiteSpace($id)) { continue }
            switch ($id) {
                'All'                    { 'Alle Benutzer'; break }
                'None'                   { 'Keine'; break }
                'GuestsOrExternalUsers'  { 'Gäste / externe Benutzer'; break }
                default { Get-Cached $userCache $id { param($x) (Get-MgUser -UserId $x -Property displayName -ErrorAction Stop).DisplayName } }
            }
        }
    }
    function Resolve-Groups { param($Ids)
        foreach ($id in @($Ids)) {
            if ([string]::IsNullOrWhiteSpace($id)) { continue }
            Get-Cached $groupCache $id { param($x) (Get-MgGroup -GroupId $x -Property displayName -ErrorAction Stop).DisplayName }
        }
    }
    function Resolve-Roles { param($Ids)
        foreach ($id in @($Ids)) {
            if ([string]::IsNullOrWhiteSpace($id)) { continue }
            if ($roleMap.ContainsKey($id)) { $roleMap[$id] } else { $id }
        }
    }
    function Resolve-Apps { param($Ids)
        foreach ($id in @($Ids)) {
            if ([string]::IsNullOrWhiteSpace($id)) { continue }
            switch ($id) {
                'All'                                    { 'Alle Cloud-Apps'; break }
                'None'                                   { 'Keine'; break }
                'Office365'                              { 'Office 365'; break }
                'MicrosoftAdminPortals'                  { 'Microsoft Admin-Portale'; break }
                default { Get-Cached $appCache $id { param($x) (Get-MgServicePrincipal -Filter "appId eq '$x'" -Property displayName -ErrorAction Stop | Select-Object -First 1).DisplayName } }
            }
        }
    }
    function Resolve-Locations { param($Ids)
        foreach ($id in @($Ids)) {
            if ([string]::IsNullOrWhiteSpace($id)) { continue }
            switch ($id) {
                'All'         { 'Alle Standorte'; break }
                'AllTrusted'  { 'Alle vertrauenswürdigen Standorte'; break }
                default { if ($locMap.ContainsKey($id)) { $locMap[$id] } else { $id } }
            }
        }
    }

    $stateLabel = @{ 'enabled' = 'Aktiv'; 'disabled' = 'Aus'; 'enabledForReportingButNotEnforced' = 'Report-only' }
    $ctrlLabel  = @{
        'mfa' = 'MFA'; 'block' = 'Zugriff blockieren'; 'compliantDevice' = 'Konformes Gerät'
        'domainJoinedDevice' = 'Domänenbeitritt'; 'approvedApplication' = 'Genehmigte App'
        'compliantApplication' = 'App-Schutzrichtlinie'; 'passwordChange' = 'Passwortänderung'
    }

    # --- Policies zu Objekten verarbeiten -----------------------------------
    $report = foreach ($p in $policies) {
        $c = $p.Conditions
        $incUsers  = @(Resolve-Users  $c.Users.IncludeUsers)  + @(Resolve-Groups $c.Users.IncludeGroups) + @(Resolve-Roles $c.Users.IncludeRoles)
        $excUsers  = @(Resolve-Users  $c.Users.ExcludeUsers)  + @(Resolve-Groups $c.Users.ExcludeGroups) + @(Resolve-Roles $c.Users.ExcludeRoles)
        $incApps   = @(Resolve-Apps   $c.Applications.IncludeApplications)
        $excApps   = @(Resolve-Apps   $c.Applications.ExcludeApplications)
        $platforms = @($c.Platforms.IncludePlatforms)
        $locations = @(Resolve-Locations $c.Locations.IncludeLocations)
        $grants    = @(@($p.GrantControls.BuiltInControls) | Where-Object { $_ } | ForEach-Object { if ($ctrlLabel.ContainsKey($_)) { $ctrlLabel[$_] } else { $_ } })

        # Impact / Hinweise
        $insights = @()
        if ($p.GrantControls.BuiltInControls -contains 'block') { $insights += @{ k = 'block'; t = 'Blockiert Zugriff' } }
        if ($p.GrantControls.BuiltInControls -contains 'mfa')   { $insights += @{ k = 'good';  t = 'Erzwingt MFA' } }
        if ($c.Users.IncludeUsers -contains 'All')              { $insights += @{ k = 'warn';  t = 'Gilt für ALLE Benutzer' } }
        $exCount = @($c.Users.ExcludeUsers).Count + @($c.Users.ExcludeGroups).Count + @($c.Users.ExcludeRoles).Count
        if ($exCount -gt 0)                                     { $insights += @{ k = 'warn';  t = "$exCount Ausnahme(n)" } }
        if ($p.State -eq 'enabledForReportingButNotEnforced')   { $insights += @{ k = 'report'; t = 'Nicht erzwungen (Report-only)' } }
        if ($p.State -eq 'disabled')                            { $insights += @{ k = 'off';   t = 'Deaktiviert' } }
        if (@($c.SignInRiskLevels).Count -or @($c.UserRiskLevels).Count) { $insights += @{ k = 'info'; t = 'Risikobasiert' } }
        $legacy = @($c.ClientAppTypes | Where-Object { $_ -in 'exchangeActiveSync', 'other' })
        if ($legacy.Count -and ($p.GrantControls.BuiltInControls -contains 'block')) { $insights += @{ k = 'good'; t = 'Blockiert Legacy-Auth' } }

        [pscustomobject]@{
            Name          = $p.DisplayName
            State         = $stateLabel[$p.State]
            StateRaw      = $p.State
            IncludeUsers  = $incUsers
            ExcludeUsers  = $excUsers
            IncludeApps   = $incApps
            ExcludeApps   = $excApps
            Platforms     = $platforms
            Locations     = $locations
            ClientApps    = @($c.ClientAppTypes)
            SignInRisk    = @($c.SignInRiskLevels)
            UserRisk      = @($c.UserRiskLevels)
            GrantOperator = $p.GrantControls.Operator
            Grants        = $grants
            Session       = @(
                if ($p.SessionControls.SignInFrequency.IsEnabled) { 'Anmeldehäufigkeit' }
                if ($p.SessionControls.PersistentBrowser.IsEnabled) { 'Persistenter Browser' }
                if ($p.SessionControls.ApplicationEnforcedRestrictions.IsEnabled) { 'App-erzwungene Einschränkungen' }
                if ($p.SessionControls.CloudAppSecurity.IsEnabled) { 'Defender for Cloud Apps' }
            )
            Insights      = $insights
            Modified      = $p.ModifiedDateTime
        }
    }

    # --- HTML bauen ----------------------------------------------------------
    $total    = @($report).Count
    $onCount  = @($report | Where-Object StateRaw -eq 'enabled').Count
    $repCount = @($report | Where-Object StateRaw -eq 'enabledForReportingButNotEnforced').Count
    $offCount = @($report | Where-Object StateRaw -eq 'disabled').Count
    $mfaCount = @($report | Where-Object { $_.Grants -contains 'MFA' }).Count
    $blkCount = @($report | Where-Object { $_.Grants -contains 'Zugriff blockieren' }).Count
    $genAt    = Get-Date -Format 'dd.MM.yyyy HH:mm'
    $tenant   = (Get-MgContext).TenantId

    function New-Pill { param($Text, $Kind = 'info') "<span class='pill pill-$Kind'>$(& $enc $Text)</span>" }
    function New-Chips { param($Items, $Kind = 'chip')
        $arr = @($Items) | Where-Object { $_ }
        if (-not $arr) { return "<span class='muted'>–</span>" }
        ($arr | ForEach-Object { "<span class='$Kind'>$(& $enc $_)</span>" }) -join ' '
    }

    $cards = foreach ($r in $report) {
        $stateKind = switch ($r.StateRaw) { 'enabled' { 'on' } 'enabledForReportingButNotEnforced' { 'report' } default { 'off' } }
        $insightPills = ($r.Insights | ForEach-Object { New-Pill $_.t $_.k }) -join ' '
        $grantText = if ($r.Grants) { (@($r.Grants) -join " <span class='op'>$($r.GrantOperator)</span> ") } else { "<span class='muted'>Keine Grant-Controls</span>" }

        $mfaFlag   = if ($r.Grants -contains 'MFA') { '1' } else { '0' }
        $blockFlag = if ($r.Grants -contains 'Zugriff blockieren') { '1' } else { '0' }
        $searchRaw = (@($r.Name) + @($r.IncludeUsers) + @($r.ExcludeUsers) + @($r.IncludeApps) + @($r.ExcludeApps) + @($r.Locations) + @($r.Grants) + @($r.Platforms)) -join ' '
        $searchAttr = & $enc ($searchRaw.ToLower())
        $nameAttr   = & $enc ([string]$r.Name).ToLower()
        $modAttr    = & $enc ([string]$r.Modified)

        @"
    <article class="card" data-state="$stateKind" data-mfa="$mfaFlag" data-block="$blockFlag" data-name="$nameAttr" data-modified="$modAttr" data-search="$searchAttr">
      <header class="card-head" role="button" tabindex="0">
        <div class="title-wrap">
          <span class="chevron" aria-hidden="true">&#9662;</span>
          <h3>$(& $enc $r.Name)</h3>
          <span class="state state-$stateKind">$(& $enc $r.State)</span>
        </div>
        <div class="insights">$insightPills</div>
      </header>
      <div class="card-body">
      <div class="grid">
        <section>
          <div class="block-title">Wer</div>
          <div class="row"><span class="lbl">Eingeschlossen</span><div class="vals">$(New-Chips $r.IncludeUsers 'chip chip-inc')</div></div>
          <div class="row"><span class="lbl">Ausgeschlossen</span><div class="vals">$(New-Chips $r.ExcludeUsers 'chip chip-exc')</div></div>
          <div class="block-title">Was (Apps)</div>
          <div class="row"><span class="lbl">Eingeschlossen</span><div class="vals">$(New-Chips $r.IncludeApps 'chip chip-inc')</div></div>
          <div class="row"><span class="lbl">Ausgeschlossen</span><div class="vals">$(New-Chips $r.ExcludeApps 'chip chip-exc')</div></div>
        </section>
        <section>
          <div class="block-title">Bedingungen</div>
          <div class="row"><span class="lbl">Plattformen</span><div class="vals">$(New-Chips $r.Platforms)</div></div>
          <div class="row"><span class="lbl">Standorte</span><div class="vals">$(New-Chips $r.Locations)</div></div>
          <div class="row"><span class="lbl">Client-Apps</span><div class="vals">$(New-Chips $r.ClientApps)</div></div>
          <div class="row"><span class="lbl">Anmelderisiko</span><div class="vals">$(New-Chips $r.SignInRisk)</div></div>
          <div class="row"><span class="lbl">Benutzerrisiko</span><div class="vals">$(New-Chips $r.UserRisk)</div></div>
          <div class="block-title">Kontrollen</div>
          <div class="row"><span class="lbl">Gewähren</span><div class="vals grant">$grantText</div></div>
          <div class="row"><span class="lbl">Sitzung</span><div class="vals">$(New-Chips $r.Session)</div></div>
        </section>
      </div>
      <footer class="card-foot">Zuletzt geändert: $(& $enc $r.Modified)</footer>
      </div>
    </article>
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
<title>$brandName · Conditional Access Report</title>
<style>
  :root{
    --bg:#f5f6fa; --card:#ffffff; --text:#171923; --muted:#7b8394; --border:#e7e9f0; --border-strong:#d6d9e4;
    --on:#16a34a; --report:#e08600; --off:#9aa0ad; --block:#e5484d; --good:#0ea5e9; --warn:#e08600; --info:#6366f1;
    --brand1:#4f46e5; --brand2:#7c3aed; --brand3:#db2777;
    --shadow:0 1px 2px rgba(16,24,40,.04),0 6px 20px rgba(16,24,40,.06);
    --shadow-lg:0 8px 30px rgba(16,24,40,.10);
    --radius:16px;
  }
  :root[data-theme="dark"], html:not([data-theme="light"]) {}
  @media (prefers-color-scheme: dark){
    :root:not([data-theme="light"]){
      --bg:#0c0e14; --card:#161922; --text:#e8eaf0; --muted:#8a90a1; --border:#242835; --border-strong:#2e3342;
      --shadow:0 1px 2px rgba(0,0,0,.4),0 8px 24px rgba(0,0,0,.35); --shadow-lg:0 12px 40px rgba(0,0,0,.5);
    }
  }
  :root[data-theme="dark"]{
    --bg:#0c0e14; --card:#161922; --text:#e8eaf0; --muted:#8a90a1; --border:#242835; --border-strong:#2e3342;
    --shadow:0 1px 2px rgba(0,0,0,.4),0 8px 24px rgba(0,0,0,.35); --shadow-lg:0 12px 40px rgba(0,0,0,.5);
  }
  *{box-sizing:border-box}
  html{scroll-behavior:smooth}
  body{margin:0;background:var(--bg);color:var(--text);
    font-family:-apple-system,BlinkMacSystemFont,"Segoe UI",Roboto,Helvetica,Arial,sans-serif;
    line-height:1.5;-webkit-font-smoothing:antialiased;text-rendering:optimizeLegibility}
  .wrap{max-width:1200px;margin:0 auto;padding:0 24px}

  /* Hero + Branding */
  .hero{position:relative;padding:30px 0 30px;color:#fff;overflow:hidden;
    background:linear-gradient(120deg,var(--brand1) 0%,var(--brand2) 52%,var(--brand3) 100%)}
  .hero::after{content:"";position:absolute;inset:0;opacity:.18;
    background:radial-gradient(600px 200px at 85% -20%,#fff,transparent)}
  .hero .wrap{position:relative;z-index:1}
  .brand{display:flex;align-items:center;gap:11px;margin-bottom:18px}
  .brand .logo{width:34px;height:34px;border-radius:9px;background:rgba(255,255,255,.16);
    display:grid;place-items:center;font-weight:800;font-size:17px;backdrop-filter:blur(4px);
    border:1px solid rgba(255,255,255,.25)}
  .brand .bn{font-weight:700;font-size:15px;letter-spacing:.01em}
  .brand .bt{font-size:12px;opacity:.8;margin-left:2px}
  .hero h1{margin:0;font-size:27px;font-weight:750;letter-spacing:-.022em}
  .hero .sub{margin:6px 0 0;opacity:.86;font-size:13px}

  /* KPIs */
  .kpis{display:grid;grid-template-columns:repeat(auto-fit,minmax(148px,1fr));gap:13px;margin:-24px auto 20px;position:relative;z-index:3}
  .kpi{background:var(--card);border:1px solid var(--border);border-radius:14px;padding:15px 17px;box-shadow:var(--shadow);
    cursor:pointer;transition:transform .12s ease,box-shadow .12s ease,border-color .12s}
  .kpi:hover{transform:translateY(-2px);box-shadow:var(--shadow-lg)}
  .kpi.active{border-color:var(--brand2);box-shadow:0 0 0 2px color-mix(in srgb,var(--brand2) 35%,transparent)}
  .kpi .n{font-size:29px;font-weight:780;letter-spacing:-.03em;line-height:1}
  .kpi .l{font-size:11.5px;color:var(--muted);text-transform:uppercase;letter-spacing:.05em;margin-top:5px;font-weight:600}
  .kpi.on .n{color:var(--on)} .kpi.report .n{color:var(--report)} .kpi.off .n{color:var(--off)}
  .kpi.mfa .n{color:var(--good)} .kpi.block .n{color:var(--block)}

  /* Toolbar */
  .toolbar{position:sticky;top:0;z-index:20;display:flex;flex-wrap:wrap;gap:10px;align-items:center;
    padding:12px 0;margin-bottom:14px;background:color-mix(in srgb,var(--bg) 88%,transparent);
    backdrop-filter:blur(10px);border-bottom:1px solid var(--border)}
  .search{flex:1 1 240px;position:relative}
  .search input{width:100%;padding:9px 12px 9px 34px;border-radius:10px;border:1px solid var(--border-strong);
    background:var(--card);color:var(--text);font-size:14px;outline:none;transition:border-color .12s,box-shadow .12s}
  .search input:focus{border-color:var(--brand2);box-shadow:0 0 0 3px color-mix(in srgb,var(--brand2) 22%,transparent)}
  .search svg{position:absolute;left:10px;top:50%;transform:translateY(-50%);opacity:.5}
  .filters{display:flex;flex-wrap:wrap;gap:6px}
  .fbtn{font-size:12.5px;font-weight:600;padding:7px 12px;border-radius:9px;border:1px solid var(--border-strong);
    background:var(--card);color:var(--text);cursor:pointer;transition:all .12s}
  .fbtn:hover{border-color:var(--brand2)}
  .fbtn.active{background:var(--brand2);border-color:var(--brand2);color:#fff}
  select,.tbtn{font-size:13px;padding:8px 11px;border-radius:9px;border:1px solid var(--border-strong);
    background:var(--card);color:var(--text);cursor:pointer;outline:none}
  .tbtn{font-weight:600}
  .tbtn:hover{border-color:var(--brand2)}
  .count{font-size:12.5px;color:var(--muted);margin-left:auto;white-space:nowrap;font-weight:600}

  /* Cards */
  .cards{display:grid;gap:16px;padding-bottom:20px}
  .card{background:var(--card);border:1px solid var(--border);border-radius:var(--radius);box-shadow:var(--shadow);
    overflow:hidden;transition:box-shadow .12s,border-color .12s}
  .card:hover{box-shadow:var(--shadow-lg)}
  .card.hidden{display:none}
  .card-head{padding:16px 20px;cursor:pointer;user-select:none;transition:background .12s}
  .card-head:hover{background:color-mix(in srgb,var(--text) 3%,transparent)}
  .card-head:focus-visible{outline:2px solid var(--brand2);outline-offset:-2px}
  .title-wrap{display:flex;align-items:center;gap:11px;flex-wrap:wrap}
  .chevron{color:var(--muted);font-size:13px;transition:transform .18s ease;flex:0 0 auto}
  .card.collapsed .chevron{transform:rotate(-90deg)}
  .title-wrap h3{margin:0;font-size:16.5px;font-weight:650;letter-spacing:-.01em;flex:1 1 auto}
  .state{font-size:10.5px;font-weight:700;padding:3px 10px;border-radius:999px;text-transform:uppercase;letter-spacing:.05em;flex:0 0 auto}
  .state-on{background:color-mix(in srgb,var(--on) 15%,transparent);color:var(--on)}
  .state-report{background:color-mix(in srgb,var(--report) 16%,transparent);color:var(--report)}
  .state-off{background:color-mix(in srgb,var(--off) 20%,transparent);color:var(--off)}
  .insights{margin-top:10px;display:flex;flex-wrap:wrap;gap:6px;padding-left:24px}
  .pill{font-size:11px;font-weight:600;padding:3px 9px;border-radius:999px;border:1px solid transparent}
  .pill-block{background:color-mix(in srgb,var(--block) 13%,transparent);color:var(--block)}
  .pill-good{background:color-mix(in srgb,var(--good) 15%,transparent);color:var(--good)}
  .pill-warn{background:color-mix(in srgb,var(--warn) 15%,transparent);color:var(--warn)}
  .pill-info{background:color-mix(in srgb,var(--info) 15%,transparent);color:var(--info)}
  .pill-report{background:color-mix(in srgb,var(--report) 15%,transparent);color:var(--report)}
  .pill-off{background:color-mix(in srgb,var(--off) 18%,transparent);color:var(--off)}
  .card-body{border-top:1px solid var(--border)}
  .card.collapsed .card-body{display:none}
  .grid{display:grid;grid-template-columns:1fr 1fr;gap:0}
  .grid section{padding:16px 20px}
  .grid section:first-child{border-right:1px solid var(--border)}
  .block-title{font-size:10.5px;font-weight:700;color:var(--muted);text-transform:uppercase;letter-spacing:.07em;margin:14px 0 8px}
  .block-title:first-child{margin-top:0}
  .row{display:flex;gap:12px;padding:4px 0;align-items:flex-start}
  .lbl{flex:0 0 116px;font-size:12.5px;color:var(--muted)}
  .vals{flex:1;display:flex;flex-wrap:wrap;gap:5px;font-size:12.5px}
  .chip{font-size:12px;padding:2px 8px;border-radius:7px;background:color-mix(in srgb,var(--text) 6%,transparent);border:1px solid var(--border)}
  .chip-inc{background:color-mix(in srgb,var(--on) 11%,transparent);border-color:color-mix(in srgb,var(--on) 28%,transparent)}
  .chip-exc{background:color-mix(in srgb,var(--block) 11%,transparent);border-color:color-mix(in srgb,var(--block) 28%,transparent)}
  .grant{font-weight:600}
  .op{font-size:10px;color:var(--muted);font-weight:700;padding:0 2px}
  .muted{color:var(--muted)}
  .card-foot{padding:10px 20px;font-size:12px;color:var(--muted);border-top:1px solid var(--border)}
  .empty{display:none;text-align:center;padding:60px 20px;color:var(--muted)}
  .empty.show{display:block}

  /* Footer */
  .footer{border-top:1px solid var(--border);margin-top:24px;padding:22px 0 40px;color:var(--muted);font-size:12.5px;
    display:flex;align-items:center;gap:10px;flex-wrap:wrap}
  .footer .logo{width:24px;height:24px;border-radius:7px;color:#fff;display:grid;place-items:center;font-weight:800;font-size:12px;
    background:linear-gradient(135deg,var(--brand1),var(--brand3))}
  .footer b{color:var(--text)}
  @media (max-width:720px){ .grid{grid-template-columns:1fr} .grid section:first-child{border-right:none;border-bottom:1px solid var(--border)} .count{margin-left:0} }
</style>
</head>
<body>
  <div class="hero">
    <div class="wrap">
      <div class="brand">
        <div class="logo">$brandInitial</div>
        <div><span class="bn">$brandName</span> <span class="bt">$brandTagline</span></div>
      </div>
      <h1>Conditional Access Report</h1>
      <p class="sub">Tenant $tenant &middot; erstellt am $genAt &middot; $total Policies</p>
    </div>
  </div>

  <div class="wrap">
    <div class="kpis">
      <div class="kpi active" data-filter="all"><div class="n">$total</div><div class="l">Policies gesamt</div></div>
      <div class="kpi on" data-filter="on"><div class="n">$onCount</div><div class="l">Aktiv</div></div>
      <div class="kpi report" data-filter="report"><div class="n">$repCount</div><div class="l">Report-only</div></div>
      <div class="kpi off" data-filter="off"><div class="n">$offCount</div><div class="l">Deaktiviert</div></div>
      <div class="kpi mfa" data-filter="mfa"><div class="n">$mfaCount</div><div class="l">Erzwingen MFA</div></div>
      <div class="kpi block" data-filter="block"><div class="n">$blkCount</div><div class="l">Blockieren</div></div>
    </div>

    <div class="toolbar">
      <div class="search">
        <svg width="16" height="16" viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="2" stroke-linecap="round"><circle cx="11" cy="11" r="7"></circle><path d="m21 21-4.3-4.3"></path></svg>
        <input id="q" type="search" placeholder="Policies durchsuchen (Name, Benutzer, App, Standort ...)" autocomplete="off">
      </div>
      <div class="filters" id="filters">
        <button class="fbtn active" data-filter="all">Alle</button>
        <button class="fbtn" data-filter="on">Aktiv</button>
        <button class="fbtn" data-filter="report">Report-only</button>
        <button class="fbtn" data-filter="off">Deaktiviert</button>
        <button class="fbtn" data-filter="mfa">MFA</button>
        <button class="fbtn" data-filter="block">Blockieren</button>
      </div>
      <select id="sort" title="Sortierung">
        <option value="name">Name (A–Z)</option>
        <option value="name-desc">Name (Z–A)</option>
        <option value="modified-desc">Zuletzt geändert</option>
        <option value="state">Status</option>
      </select>
      <button class="tbtn" id="toggleAll" title="Alle ein-/ausklappen">Zuklappen</button>
      <button class="tbtn" id="theme" title="Design wechseln">&#9681; Design</button>
      <span class="count" id="count"></span>
    </div>

    <div class="cards" id="cards">
$($cards -join "`n")
    </div>
    <div class="empty" id="empty">Keine Policy entspricht den Filtern.</div>

    <div class="footer">
      <div class="logo">$brandInitial</div>
      <div>Erstellt mit <b>$brandName</b> &middot; $brandTagline &middot; $genAt</div>
    </div>
  </div>

<script>
(function(){
  var cards = Array.prototype.slice.call(document.querySelectorAll('.card'));
  var q = document.getElementById('q');
  var count = document.getElementById('count');
  var empty = document.getElementById('empty');
  var container = document.getElementById('cards');
  var activeFilter = 'all';

  function matchesFilter(c){
    switch(activeFilter){
      case 'all': return true;
      case 'on': case 'report': case 'off': return c.dataset.state === activeFilter;
      case 'mfa': return c.dataset.mfa === '1';
      case 'block': return c.dataset.block === '1';
    }
    return true;
  }
  function apply(){
    var term = (q.value || '').trim().toLowerCase();
    var visible = 0;
    cards.forEach(function(c){
      var ok = matchesFilter(c) && (term === '' || c.dataset.search.indexOf(term) !== -1 || c.dataset.name.indexOf(term) !== -1);
      c.classList.toggle('hidden', !ok);
      if(ok) visible++;
    });
    count.textContent = visible + ' von ' + cards.length + ' sichtbar';
    empty.classList.toggle('show', visible === 0);
  }
  function setFilter(f){
    activeFilter = f;
    document.querySelectorAll('#filters .fbtn').forEach(function(b){ b.classList.toggle('active', b.dataset.filter === f); });
    document.querySelectorAll('.kpi').forEach(function(k){ k.classList.toggle('active', k.dataset.filter === f); });
    apply();
  }

  // Suche + Filter
  q.addEventListener('input', apply);
  document.querySelectorAll('#filters .fbtn').forEach(function(b){ b.addEventListener('click', function(){ setFilter(b.dataset.filter); }); });
  document.querySelectorAll('.kpi').forEach(function(k){ k.addEventListener('click', function(){ setFilter(k.dataset.filter); }); });

  // Sortierung
  var stateOrder = { on:0, report:1, off:2 };
  document.getElementById('sort').addEventListener('change', function(e){
    var v = e.target.value;
    var sorted = cards.slice().sort(function(a,b){
      if(v === 'name') return a.dataset.name.localeCompare(b.dataset.name);
      if(v === 'name-desc') return b.dataset.name.localeCompare(a.dataset.name);
      if(v === 'modified-desc') return (b.dataset.modified||'').localeCompare(a.dataset.modified||'');
      if(v === 'state') return (stateOrder[a.dataset.state]-stateOrder[b.dataset.state]) || a.dataset.name.localeCompare(b.dataset.name);
      return 0;
    });
    sorted.forEach(function(c){ container.appendChild(c); });
  });

  // Karten ein-/ausklappen
  function toggleCard(c){ c.classList.toggle('collapsed'); }
  cards.forEach(function(c){
    var head = c.querySelector('.card-head');
    head.addEventListener('click', function(){ toggleCard(c); });
    head.addEventListener('keydown', function(e){ if(e.key==='Enter'||e.key===' '){ e.preventDefault(); toggleCard(c); } });
  });
  var allCollapsed = false;
  document.getElementById('toggleAll').addEventListener('click', function(e){
    allCollapsed = !allCollapsed;
    cards.forEach(function(c){ c.classList.toggle('collapsed', allCollapsed); });
    e.target.textContent = allCollapsed ? 'Aufklappen' : 'Zuklappen';
  });

  // Theme-Toggle (auto -> hell -> dunkel)
  var themes = ['auto','light','dark'];
  var ti = 0;
  try { var saved = localStorage.getItem('ca_theme'); if(saved){ ti = themes.indexOf(saved); if(ti<0) ti=0; } } catch(_){}
  function applyTheme(){
    var t = themes[ti];
    if(t === 'auto') document.documentElement.removeAttribute('data-theme');
    else document.documentElement.setAttribute('data-theme', t);
    document.getElementById('theme').innerHTML = '&#9681; ' + (t==='auto'?'Auto':(t==='light'?'Hell':'Dunkel'));
    try { localStorage.setItem('ca_theme', t); } catch(_){}
  }
  document.getElementById('theme').addEventListener('click', function(){ ti = (ti+1)%themes.length; applyTheme(); });

  applyTheme();
  apply();
})();
</script>
</body>
</html>
"@

    Write-TTLog -Level INFO -Message "Report erstellt: $Path ($total Policies)."
    $flat = $report | Select-Object Name, State,
        @{N = 'IncludeUsers'; E = { $_.IncludeUsers -join '; ' } }, @{N = 'ExcludeUsers'; E = { $_.ExcludeUsers -join '; ' } },
        @{N = 'IncludeApps'; E = { $_.IncludeApps -join '; ' } }, @{N = 'ExcludeApps'; E = { $_.ExcludeApps -join '; ' } },
        @{N = 'Platforms'; E = { $_.Platforms -join '; ' } }, @{N = 'Locations'; E = { $_.Locations -join '; ' } },
        @{N = 'ClientApps'; E = { $_.ClientApps -join '; ' } }, @{N = 'SignInRisk'; E = { $_.SignInRisk -join '; ' } },
        @{N = 'UserRisk'; E = { $_.UserRisk -join '; ' } }, GrantOperator,
        @{N = 'Grants'; E = { $_.Grants -join '; ' } }, @{N = 'Session'; E = { $_.Session -join '; ' } }, Modified
    Complete-TTReport -Html $html -Path $Path -Data $flat -Csv:$Csv -Excel:$Excel -DataPath $DataPath -NoHtml:$NoHtml -NoOpen:$NoOpen -Kind 'Conditional-Access-Report'
    if ($PassThru) { $report }
}
