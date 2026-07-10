function Export-M365PimReport {
    <#
    .SYNOPSIS
        Report ueber privilegierte Rollenzuweisungen (PIM): dauerhaft vs. eligible vs. aktiviert.
    .DESCRIPTION
        Liest aktive und (PIM-)eligible Rollenzuweisungen via Graph und zeigt pro Prinzipal/Rolle
        den Status: Permanent (stehender Zugriff - riskant), Eligible (Just-in-Time - gut) oder
        Aktiviert (gerade aktiv). Privilegierte Rollen (z. B. Global Administrator) werden markiert.
        Interaktiver HTML-Report. Reines Lesen.

        Hinweis: Eligible-Zuweisungen erfordern Entra ID P2. Ohne P2 werden nur aktive/dauerhafte
        Zuweisungen angezeigt (Eligible bleibt leer) - der Report laeuft trotzdem.
    .PARAMETER Path
        Zielpfad der HTML-Datei. Standard: .\PIM-Report.html
    .PARAMETER BrandName
        Branding. Fuer CloudNest365: -BrandName 'CloudNest365'
    .PARAMETER PassThru
        Gibt die Zuweisungs-Objekte zusaetzlich auf die Pipeline aus.
    .PARAMETER NoOpen
        Report nicht automatisch oeffnen.
    .EXAMPLE
        Export-M365PimReport -BrandName 'CloudNest365'
    #>
    [CmdletBinding()]
    [OutputType([pscustomobject])]
    param(
        [string]$Path = (Join-Path (Get-Location) 'PIM-Report.html'),
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
    Write-TTLog -Level INFO -Message "Lese Rollenzuweisungen (PIM) ..."

    $privileged = @(
        'Global Administrator', 'Privileged Role Administrator', 'Privileged Authentication Administrator',
        'Security Administrator', 'Conditional Access Administrator', 'Exchange Administrator',
        'SharePoint Administrator', 'User Administrator', 'Application Administrator',
        'Cloud Application Administrator', 'Hybrid Identity Administrator', 'Intune Administrator',
        'Authentication Administrator', 'Helpdesk Administrator'
    )

    # Prinzipal aus dem $expand-Objekt lesen (kein Extra-Call noetig)
    function Convert-Principal { param($p, $Id)
        $res = @{ Name = $Id; Sub = ''; Type = 'unknown' }
        if (-not $p) { return $res }
        $res.Name = [string]$p.displayName
        $t = [string]$p.'@odata.type'
        if ($t -match 'user') { $res.Type = 'user'; $res.Sub = [string]$p.userPrincipalName }
        elseif ($t -match 'group') { $res.Type = 'group'; $res.Sub = 'Gruppe' }
        elseif ($t -match 'servicePrincipal') { $res.Type = 'sp'; $res.Sub = 'Dienstprinzipal' }
        return $res
    }

    $base = 'https://graph.microsoft.com/v1.0/roleManagement/directory'
    $expand = '?$expand=principal,roleDefinition&$top=100'
    $records = New-Object System.Collections.Generic.List[object]
    $activatedKeys = @{}

    # Aktive Zuweisungen (Permanent / Aktiviert)
    try {
        $active = Get-TTGraphCollection "$base/roleAssignmentScheduleInstances$expand"
        foreach ($a in $active) {
            $roleName = if ($a.roleDefinition.displayName) { $a.roleDefinition.displayName } else { $a.roleDefinitionId }
            $status = if ($a.assignmentType -eq 'Activated') { 'Activated' } else { 'Permanent' }
            if ($status -eq 'Activated') { $activatedKeys["$($a.principalId)|$($a.roleDefinitionId)"] = $true }
            $p = Convert-Principal $a.principal $a.principalId
            $records.Add([pscustomobject]@{ Principal = $p.Name; Sub = $p.Sub; PType = $p.Type; Role = $roleName; Status = $status; End = $a.endDateTime; MemberType = $a.memberType })
        }
    }
    catch { Write-TTLog -Level WARN -Message "Aktive Zuweisungen uebersprungen: $_" }

    # Eligible Zuweisungen (JIT) - benoetigt P2
    try {
        $elig = Get-TTGraphCollection "$base/roleEligibilityScheduleInstances$expand"
        foreach ($e in $elig) {
            if ($activatedKeys.ContainsKey("$($e.principalId)|$($e.roleDefinitionId)")) { continue }
            $roleName = if ($e.roleDefinition.displayName) { $e.roleDefinition.displayName } else { $e.roleDefinitionId }
            $p = Convert-Principal $e.principal $e.principalId
            $records.Add([pscustomobject]@{ Principal = $p.Name; Sub = $p.Sub; PType = $p.Type; Role = $roleName; Status = 'Eligible'; End = $e.endDateTime; MemberType = $e.memberType })
        }
    }
    catch { Write-TTLog -Level WARN -Message "Eligible-Zuweisungen uebersprungen (evtl. keine P2-Lizenz): $_" }

    $rec = @($records | Sort-Object @{E = { $_.Status -ne 'Permanent' } }, Role, Principal)

    # KPIs
    $total     = $rec.Count
    $permanent = @($rec | Where-Object Status -eq 'Permanent').Count
    $eligible  = @($rec | Where-Object Status -eq 'Eligible').Count
    $activated = @($rec | Where-Object Status -eq 'Activated').Count
    $privPerm  = @($rec | Where-Object { $_.Status -eq 'Permanent' -and $privileged -contains $_.Role }).Count
    $ga        = @($rec | Where-Object Role -eq 'Global Administrator').Count
    $genAt     = Get-Date -Format 'dd.MM.yyyy HH:mm'
    $tenant    = (Get-MgContext).TenantId

    $statusBadge = @{ Permanent = "<span class='b b-bad'>permanent</span>"; Eligible = "<span class='b b-ok'>eligible (JIT)</span>"; Activated = "<span class='b b-warn'>aktiviert</span>" }
    $typeLabel   = @{ user = 'Benutzer'; group = 'Gruppe'; sp = 'Dienstprinzipal'; unknown = '–' }

    $rows = foreach ($r in $rec) {
        $isPriv = $privileged -contains $r.Role
        $isGa   = $r.Role -eq 'Global Administrator'
        $sKey   = $r.Status.ToLower()
        $endTxt = if ($r.End) { ([datetime]$r.End).ToString('dd.MM.yyyy') } else { '<span class="muted">unbefristet</span>' }
        $rolePill = if ($isGa) { " <span class='b b-crit'>&#9888; GA</span>" } elseif ($isPriv) { " <span class='b b-info'>privilegiert</span>" } else { '' }
        $searchAttr = TTEnc ("$($r.Principal) $($r.Sub) $($r.Role) $($r.Status)".ToLower())
        $nameAttr = TTEnc ([string]$r.Principal).ToLower()
        $roleAttr = TTEnc ([string]$r.Role).ToLower()
        $fPriv = if ($isPriv) { '1' } else { '0' }
        $fGa   = if ($isGa) { '1' } else { '0' }
        @"
      <tr class="item" data-f-$sKey="1" data-f-privileged="$fPriv" data-f-globaladmin="$fGa" data-name="$nameAttr" data-s-role="$roleAttr" data-search="$searchAttr">
        <td><div class="u"><b>$(TTEnc $r.Principal)</b><span class="upn">$(if ($r.Sub) { TTEnc $r.Sub } else { $typeLabel[$r.PType] })</span></div></td>
        <td>$(TTEnc $r.Role)$rolePill</td>
        <td>$($typeLabel[$r.PType])</td>
        <td>$($statusBadge[$r.Status])</td>
        <td>$endTxt</td>
      </tr>
"@
    }

    $kpiHtml = New-TTKpis @(
        @{ n = $total; l = 'Zuweisungen'; filter = 'all' }
        @{ n = $permanent; l = 'Permanent'; kind = 'bad'; filter = 'permanent' }
        @{ n = $eligible; l = 'Eligible (JIT)'; kind = 'ok'; filter = 'eligible' }
        @{ n = $activated; l = 'Aktiviert'; kind = 'warn'; filter = 'activated' }
        @{ n = $privPerm; l = 'Priv. permanent'; kind = 'bad'; filter = 'privileged' }
        @{ n = $ga; l = 'Global Admins'; kind = 'adm'; filter = 'globaladmin' }
    )
    $toolbar = New-TTToolbar -SearchPlaceholder 'Prinzipal oder Rolle suchen ...' -Filters @(
        @{ label = 'Alle'; key = 'all' }, @{ label = 'Permanent'; key = 'permanent' }, @{ label = 'Eligible'; key = 'eligible' },
        @{ label = 'Aktiviert'; key = 'activated' }, @{ label = 'Privilegiert'; key = 'privileged' }, @{ label = 'Global Admins'; key = 'globaladmin' }
    )
    $body = @"
    <div class="panel">
      <table class="tbl">
        <thead><tr><th data-sort="name">Prinzipal</th><th data-sort="role">Rolle</th><th>Typ</th><th>Status</th><th>Ablauf</th></tr></thead>
        <tbody>
$($rows -join "`n")
        </tbody>
      </table>
      <div class="empty" id="empty">Keine Zuweisung entspricht den Filtern.</div>
    </div>
"@

    $sub = "Tenant $tenant &middot; erstellt am $genAt &middot; $total Rollenzuweisungen"
    $html = New-TTHtmlPage -Title 'PIM / Rollen-Report' -Heading 'Privilegierte Rollen (PIM)' -Sub $sub `
        -BrandName $BrandName -BrandTagline $BrandTagline -KpiHtml $kpiHtml -ToolbarHtml $toolbar -BodyHtml $body

    Write-TTLog -Level INFO -Message "PIM-Report erstellt: $Path ($permanent permanent, $privPerm privilegiert-permanent, $ga Global Admins)."
    if ($privPerm -gt 0) { Write-Host "  Achtung: $privPerm privilegierte Rolle(n) mit dauerhaftem Zugriff!" -ForegroundColor Red }

    $flat = $rec | Select-Object Principal, @{N = 'Detail'; E = { $_.Sub } }, @{N = 'Type'; E = { $_.PType } },
        Role, Status, MemberType, @{N = 'End'; E = { if ($_.End) { ([datetime]$_.End).ToString('yyyy-MM-dd') } else { 'unbefristet' } } }
    Complete-TTReport -Html $html -Path $Path -Data $flat -Csv:$Csv -Excel:$Excel -DataPath $DataPath -NoHtml:$NoHtml -NoOpen:$NoOpen -Kind 'PIM-Report'
    if ($PassThru) { $rec }
}
