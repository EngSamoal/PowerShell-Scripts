#Requires -Version 5.1
<#
.SYNOPSIS
    Backup & Storage Weekly Health Check - reads the Backup Summary workbook and produces a
    single combined HTML dashboard, in the same visual style as Network_Weekly_HealthCheck.ps1.

.DESCRIPTION
    This dashboard covers Backup job health today. Storage is a separate data source that has
    not been supplied yet - every site's Storage panel shows "There is no data provided yet"
    until that workbook format is provided and wired in; this script only needs re-running once
    that happens, nothing here needs to change in the meantime.

    Reads ONE workbook (-BackupWorkbookPath), with one tab per site - the real file seen so far
    uses the short tab names AQ / SF / TABUK / ABHA (SceneCinema and SEVEN Alhamra don't have
    data yet, so their tab names are a best guess below - update $SiteTabCandidates once their
    real tab name is known, nothing else needs to change). Each site tab has a "Protection Group"
    table (Protection Group / Source / Runs / Last Run Success-Error / Data Read / SLA Violation
    / Last Run Status / Last Run Replication-Archival Status (icon, not read) / Bandwidth), read
    by locating the "Protection Group" / "Source" / "Runs" header row and then every column after
    "Runs" by fixed offset, since the real header wraps across two lines in Excel.

    A protection group with 0 runs this period (shown as "-" across every other column in the
    source sheet) is its own neutral status - it doesn't count as Healthy, Warning or Critical,
    the same way a template site with no workbook supplied doesn't skew the site-level counts.
    Every other row's status comes from its own Last Run Status and SLA Violation columns:
    Error -> Critical; Warning status OR a failed SLA Violation -> Warning; Success + Pass ->
    Healthy. A site's overall health is Critical if ANY protection group is Critical, Warning if
    any is Warning, else Healthy - unlike the Network dashboard's rack-health threshold (a
    handful of bad racks out of 15-20 is normal noise), a site typically only has a handful of
    protection groups, so any single failed backup job is itself worth flagging immediately.

    Every tab is located by searching for its known column headers rather than assuming a fixed
    row/column position, since real workbooks can have inconsistent header rows.

    This dashboard template covers 6 sites (AquaArabia, SixFlags, SceneCinema, SEVEN Tabuk,
    SEVEN ABHA, SEVEN Alhamra) so it's ready to use as each site's data becomes available - a
    site whose tab isn't found in the workbook still gets its own tab and page, showing
    "There is no data provided yet" instead of being silently left out.

    NEVER writes/modifies the source workbook - opened read-only, closed without saving.

.PARAMETER BackupWorkbookPath
    Path to the single "Backup Summary - <Month> <Year>.xlsx" workbook (one tab per site). Omit
    to leave every site's Backup section showing "There is no data provided yet".
.PARAMETER OutputPath
    Folder the dashboard + log are written to.
#>

[CmdletBinding()]
param(
    [string]$BackupWorkbookPath = '',

    [string]$OutputPath = (Join-Path $PSScriptRoot "BackupStorage_HealthCheck_Reports")
)

$ScriptBuild = '2026-10-04-01-backupstorage-initial'
Write-Host "BackupStorage_Weekly_HealthCheck.ps1 - build $ScriptBuild" -ForegroundColor Magenta

$ErrorActionPreference = 'Stop'
$ScriptStart = Get-Date
$RunDate = $ScriptStart.ToString('yyyy-MM-dd')
$RunDateDisplay = $ScriptStart.ToString('d-MMMM-yyyy', [System.Globalization.CultureInfo]::InvariantCulture)
if (-not (Test-Path $OutputPath)) { New-Item -ItemType Directory -Path $OutputPath -Force | Out-Null }

# ============================================================================
# 0. LOGGING / SAFE-EXECUTION HELPERS (same pattern as Network_Weekly_HealthCheck.ps1)
# ============================================================================
$Global:FailureLog = [System.Collections.Generic.List[object]]::new()
$Global:AllResults = [System.Collections.Generic.List[object]]::new()

function Write-CheckLog {
    param([string]$Site, [string]$Object, [string]$CheckName, [string]$ErrorMessage)
    $Global:FailureLog.Add([pscustomobject]@{
        Timestamp = (Get-Date).ToString('s')
        Site      = $Site
        Object    = $Object
        Check     = $CheckName
        Error     = $ErrorMessage
    })
    Write-Warning "[$Site] $CheckName on '$Object' failed: $ErrorMessage"
}

function Invoke-SafeCheck {
    param(
        [Parameter(Mandatory)][scriptblock]$Script,
        [Parameter(Mandatory)][string]$CheckName,
        [string]$Site = 'n/a',
        [string]$ObjectName = 'n/a'
    )
    try {
        & $Script
    } catch {
        Write-CheckLog -Site $Site -Object $ObjectName -CheckName $CheckName -ErrorMessage $_.Exception.Message
        return $null
    }
}

function New-Finding {
    param(
        [string]$Site, [string]$Area, [string]$Group, [string]$Object,
        [string]$Item, [string]$Value, [string]$Status, [string]$Notes = ''
    )
    # Area is always 'Backup' today - 'Storage' is reserved for once that data source is wired in.
    $obj = [pscustomobject]@{
        Site = $Site; Area = $Area; Group = $Group; Object = $Object
        Item = $Item; Value = $Value; Status = $Status; Notes = $Notes
    }
    $Global:AllResults.Add($obj) | Out-Null
}

function ConvertTo-CompactNote {
    param([string]$Text)
    if ([string]::IsNullOrWhiteSpace($Text)) { return '' }
    return ($Text -replace '\r?\n', ' | ').Trim()
}

# ============================================================================
# 1. EXCEL HEADER-ANCHORING HELPERS
# ============================================================================
function Find-HeaderRow {
    param($Worksheet, [string[]]$Labels, [int]$MaxRow = 40, [int]$MaxCol = 20)
    for ($r = 1; $r -le $MaxRow; $r++) {
        $found = @{}
        for ($c = 1; $c -le $MaxCol; $c++) {
            $v = $Worksheet.Cells.Item($r, $c).Value2
            if ($null -eq $v) { continue }
            $vs = "$v".Trim()
            foreach ($lab in $Labels) {
                if (-not $found.ContainsKey($lab) -and $vs -ieq $lab) { $found[$lab] = $c }
            }
        }
        if ($found.Count -eq $Labels.Count) {
            return [pscustomobject]@{ Row = $r; Columns = $found }
        }
    }
    return $null
}

# ============================================================================
# 2. BACKUP SUMMARY TAB COLLECTION
# ============================================================================
function Read-BackupSiteTab {
    param($Worksheet, [string]$Site)
    $hdr = Find-HeaderRow -Worksheet $Worksheet -Labels @('Protection Group', 'Source', 'Runs')
    if (-not $hdr) { throw "Could not locate the Protection Group / Source / Runs header on this tab." }
    $cols = $hdr.Columns
    $pgCol = $cols['Protection Group']; $sourceCol = $cols['Source']; $runsCol = $cols['Runs']
    # Everything after "Runs" is read by fixed offset rather than header text match - the real
    # header wraps across two lines in Excel ("Last Run Success/" + "Error" in one cell), which
    # makes an exact text match unreliable; the column ORDER is consistent across every site tab.
    $lastRunObjCol = $runsCol + 1
    $dataReadCol   = $runsCol + 2
    $slaCol        = $runsCol + 3
    $lastStatusCol = $runsCol + 4
    # $runsCol + 5 is the Replication/Archival Status icon column - not read, it carries no text.
    $bandwidthCol  = $runsCol + 6

    $used = $Worksheet.UsedRange
    $lastRow = $used.Row + $used.Rows.Count - 1
    for ($r = $hdr.Row + 1; $r -le $lastRow; $r++) {
        $pgVal = $Worksheet.Cells.Item($r, $pgCol).Value2
        if (-not $pgVal -or "$pgVal".Trim() -eq '') { continue }
        $pgName = "$pgVal".Trim()

        $source = "$($Worksheet.Cells.Item($r, $sourceCol).Value2)".Trim()
        $runsVal = $Worksheet.Cells.Item($r, $runsCol).Value2
        $runsN = 0
        if ($runsVal) { [int]::TryParse("$runsVal", [ref]$runsN) | Out-Null }
        $lastRunObjText = "$($Worksheet.Cells.Item($r, $lastRunObjCol).Value2)".Trim()
        $dataRead = "$($Worksheet.Cells.Item($r, $dataReadCol).Value2)".Trim()
        $sla = "$($Worksheet.Cells.Item($r, $slaCol).Value2)".Trim()
        $lastStatus = "$($Worksheet.Cells.Item($r, $lastStatusCol).Value2)".Trim()
        $bandwidth = "$($Worksheet.Cells.Item($r, $bandwidthCol).Value2)".Trim()

        # 0 runs this period (shown as "-" across the sheet) is its own neutral status - not
        # counted as Healthy, Warning or Critical, same as a template site with no data supplied.
        $status = if ($runsN -eq 0 -or -not $lastStatus -or $lastStatus -eq '-') { 'Information' }
            elseif ($lastStatus -ieq 'Error') { 'Critical' }
            elseif ($lastStatus -ieq 'Warning' -or $sla -ieq 'Fail') { 'Warning' }
            elseif ($lastStatus -ieq 'Success' -and $sla -ieq 'Pass') { 'Healthy' }
            else { 'Information' }

        $value = "$runsN|$lastRunObjText|$dataRead|$sla|$lastStatus|$bandwidth"
        New-Finding -Site $Site -Area 'Backup' -Group $pgName -Object $source -Item 'Backup Job' -Value $value -Status $status
    }
}

# ============================================================================
# 3. RUN COLLECTION
# ============================================================================
# Each site's candidate tab name(s) inside the single Backup Summary workbook, tried in order,
# case-insensitive. AQ / SF / TABUK / ABHA are the real, confirmed tab names. SceneCinema and
# SEVEN Alhamra have no data yet, so their entries are a best guess - once that site's tab shows
# up in a real workbook, just correct its candidate list here, nothing else needs to change.
$SiteTabCandidates = [ordered]@{
    'AquaArabia'    = @('AQ')
    'SixFlags'      = @('SF')
    'SceneCinema'   = @('SC', 'SCC', 'SceneCinema')
    'SEVEN Tabuk'   = @('TABUK')
    'SEVEN ABHA'    = @('ABHA')
    'SEVEN Alhamra' = @('ALHAMRA', 'ALH', 'Alhamra')
}

if ($BackupWorkbookPath -and (Test-Path $BackupWorkbookPath)) {
    $Excel = $null
    try {
        $Excel = New-Object -ComObject Excel.Application
        $Excel.Visible = $false
        $Excel.DisplayAlerts = $false
    } catch {
        throw "Microsoft Excel is not available via COM automation on this machine - cannot read the source workbook. $($_.Exception.Message)"
    }
    $wb = $Excel.Workbooks.Open($BackupWorkbookPath, 0, $true)
    try {
        foreach ($site in $SiteTabCandidates.Keys) {
            Write-Host "`n=== Collecting: $site ===" -ForegroundColor Green
            $ws = $null
            foreach ($candidate in $SiteTabCandidates[$site]) {
                $ws = $wb.Worksheets | Where-Object { $_.Name.Trim() -ieq $candidate } | Select-Object -First 1
                if ($ws) { break }
            }
            if (-not $ws) {
                Write-CheckLog -Site $site -Object $BackupWorkbookPath -CheckName 'Backup tab' -ErrorMessage "No tab matching $($SiteTabCandidates[$site] -join '/') found - this site shows 'There is no data provided yet'."
                continue
            }
            Invoke-SafeCheck -Site $site -ObjectName $ws.Name -CheckName 'Backup Summary tab' -Script {
                Read-BackupSiteTab -Worksheet $ws -Site $site
            }
        }
    } finally {
        $wb.Close($false)
        [System.Runtime.Interopservices.Marshal]::ReleaseComObject($wb) | Out-Null
        $Excel.Quit()
        [System.Runtime.Interopservices.Marshal]::ReleaseComObject($Excel) | Out-Null
    }
} else {
    if ($BackupWorkbookPath) { Write-CheckLog -Site 'n/a' -Object $BackupWorkbookPath -CheckName 'Backup workbook' -ErrorMessage 'File not found.' }
    Write-Host "No -BackupWorkbookPath supplied - every site will show 'There is no data provided yet'." -ForegroundColor Yellow
}

# ============================================================================
# 4. DASHBOARD GENERATION HELPERS
# ============================================================================
function ConvertTo-HtmlSafe {
    param([string]$Text)
    if (-not $Text) { return '' }
    return $Text -replace '&', '&amp;' -replace '<', '&lt;' -replace '>', '&gt;' -replace '"', '&quot;'
}

function ConvertTo-Slug {
    param([string]$Text)
    $slug = ($Text -replace '[^a-zA-Z0-9]+', '-').Trim('-').ToLower()
    if (-not $slug) { return 'site' }
    return $slug
}

function Get-SiteDashboardSummary {
    param([string]$SiteLabel)
    $SiteFindings = @($Global:AllResults | Where-Object { $_.Site -eq $SiteLabel -and $_.Area -eq 'Backup' })
    if ($SiteFindings.Count -eq 0) {
        return [pscustomobject]@{
            Site           = $SiteLabel
            HasData        = $false
            OverallHealth  = 'NoData'
            CriticalCount  = 0
            WarningCount   = 0
            HealthyCount   = 0
            NoRunsCount    = 0
            TotalGroups    = 0
            BackupRows     = @()
            SummaryText    = ''
        }
    }

    $rows = @(foreach ($f in $SiteFindings) {
        $p = $f.Value -split '\|'
        [pscustomobject]@{
            ProtectionGroup = $f.Group
            Source          = $f.Object
            Runs            = $p[0]
            LastRunObjects  = $p[1]
            DataRead        = $p[2]
            SLAViolation    = $p[3]
            LastRunStatus   = $p[4]
            Bandwidth       = $p[5]
            Status          = $f.Status
        }
    })

    $criticalCount = @($rows | Where-Object { $_.Status -eq 'Critical' }).Count
    $warningCount  = @($rows | Where-Object { $_.Status -eq 'Warning' }).Count
    $healthyCount  = @($rows | Where-Object { $_.Status -eq 'Healthy' }).Count
    $noRunsCount   = @($rows | Where-Object { $_.Status -eq 'Information' }).Count

    # Unlike the Network dashboard's rack-health threshold (a handful of bad racks out of 15-20
    # is normal noise), a site here typically only has a handful of protection groups - so any
    # single failed/SLA-missed backup job is itself worth flagging at the site level immediately.
    $OverallHealth = if ($criticalCount -gt 0) { 'Critical' }
        elseif ($warningCount -gt 0) { 'Warning' }
        else { 'Healthy' }

    $SummaryText = if ($criticalCount -gt 0) {
        "$SiteLabel has $criticalCount critical and $warningCount warning backup job(s) across $($rows.Count) protection group(s)."
    } elseif ($warningCount -gt 0) {
        "$SiteLabel is stable overall, with $warningCount backup job(s) flagged for a missed SLA or warning status across $($rows.Count) protection group(s)."
    } else {
        "$SiteLabel's $($rows.Count) protection group(s) all completed their last backup run successfully within SLA."
    }

    [pscustomobject]@{
        Site          = $SiteLabel
        HasData       = $true
        OverallHealth = $OverallHealth
        CriticalCount = $criticalCount
        WarningCount  = $warningCount
        HealthyCount  = $healthyCount
        NoRunsCount   = $noRunsCount
        TotalGroups   = $rows.Count
        BackupRows    = $rows
        SummaryText   = $SummaryText
    }
}

function Write-DashboardHtml {
    param([string[]]$SiteLabels, [string]$OutputPath, [string]$RunDateDisplay)

    $summaries = @($SiteLabels | ForEach-Object { Get-SiteDashboardSummary -SiteLabel $_ })
    if ($summaries.Count -eq 0) { return }
    $dataSummaries = @($summaries | Where-Object { $_.HasData })

    $healthColor = @{ Healthy = '#2e7d32'; Warning = '#e6a100'; Critical = '#c62828'; NoData = '#8a8f98' }
    $healthLabelText = @{ Healthy = 'Healthy - No Issues Detected'; Warning = 'Healthy - Minor Issues Detected'; Critical = 'Attention Required - Critical Issues'; NoData = 'There is no data provided yet' }
    $healthyNames  = @($dataSummaries | Where-Object { $_.OverallHealth -eq 'Healthy' }  | Select-Object -ExpandProperty Site)
    $warningNames  = @($dataSummaries | Where-Object { $_.OverallHealth -eq 'Warning' }  | Select-Object -ExpandProperty Site)
    $criticalNames = @($dataSummaries | Where-Object { $_.OverallHealth -eq 'Critical' } | Select-Object -ExpandProperty Site)
    $healthCounts = @{ Healthy = $healthyNames.Count; Warning = $warningNames.Count; Critical = $criticalNames.Count }

    function Get-SiteListNoteHtml {
        param([string[]]$Names)
        if ($Names.Count -eq 0) { return '<div class="stat-note">No sites in this bucket</div>' }
        return "<div class=`"stat-note`">$(($Names | ForEach-Object { ConvertTo-HtmlSafe $_ }) -join ', ')</div>"
    }
    function Get-SiteBreakdownNoteHtml {
        param([string]$Property)
        $parts = @($dataSummaries | Where-Object { $_.$Property -gt 0 } | ForEach-Object { "$(ConvertTo-HtmlSafe $_.Site): $($_.$Property)" })
        if ($parts.Count -eq 0) { return '<div class="stat-note">None reported</div>' }
        return "<div class=`"stat-note`">$($parts -join ', ')</div>"
    }
    $healthyNoteHtml  = Get-SiteListNoteHtml -Names $healthyNames
    $warningNoteHtml  = Get-SiteListNoteHtml -Names $warningNames
    $criticalNoteHtml = Get-SiteListNoteHtml -Names $criticalNames
    $totalGroups   = ($dataSummaries | Measure-Object -Property TotalGroups -Sum).Sum
    $totalCritical = ($dataSummaries | Measure-Object -Property CriticalCount -Sum).Sum
    $totalWarning  = ($dataSummaries | Measure-Object -Property WarningCount -Sum).Sum
    $groupsNoteHtml   = Get-SiteBreakdownNoteHtml -Property TotalGroups
    $criticalJobsNote = Get-SiteBreakdownNoteHtml -Property CriticalCount
    $warningJobsNote  = Get-SiteBreakdownNoteHtml -Property WarningCount

    $tabPalette = @('#2C5577', '#3F6652', '#6B3F42', '#5B4B77', '#7A5C3E', '#45586B', '#3E6B6B', '#5A5240')
    $tileBlue = '#1565C0'; $tilePurple = '#6A1B9A'; $tileTeal = '#00897B'

    function Get-StatusPillHtml {
        param([string]$Status, [string]$Text)
        $variant = switch ($Status) { 'Healthy' { 'ok' }; 'Warning' { 'warn' }; 'Critical' { 'bad' }; default { 'info' } }
        "<span class=`"status-pill $variant`"><span class=`"dot`"></span>$(ConvertTo-HtmlSafe $Text)</span>"
    }

    # --- Overview page ---
    $overviewCards = ($summaries | ForEach-Object {
        $s = $_
        $slug = ConvertTo-Slug $s.Site
        $color = $healthColor[$s.OverallHealth]
        if (-not $s.HasData) {
@"
      <div class="ov-card ov-card-nodata" onclick="showPage('$slug')" style="border-top-color:$color">
        <div class="ov-head"><h2>$(ConvertTo-HtmlSafe $s.Site)</h2><span class="badge" style="background:$color">$(ConvertTo-HtmlSafe $healthLabelText.NoData)</span></div>
        <p class="nodata-text">There is no data provided yet</p>
        <span class="ov-link">View Page &rarr;</span>
      </div>
"@
            return
        }
@"
      <div class="ov-card" onclick="showPage('$slug')" style="border-top-color:$color">
        <div class="ov-head"><h2>$(ConvertTo-HtmlSafe $s.Site)</h2><span class="badge" style="background:$color">$(ConvertTo-HtmlSafe $healthLabelText[$s.OverallHealth])</span></div>
        <div class="risk-row"><span class="risk risk-high">Critical Jobs: $($s.CriticalCount)</span><span class="risk risk-med">Warning Jobs: $($s.WarningCount)</span></div>
        <table class="metrics">
          <tr><td>Protection Groups</td><td>$($s.TotalGroups)</td></tr>
          <tr><td>Healthy / No Runs This Period</td><td>$($s.HealthyCount) / $($s.NoRunsCount)</td></tr>
        </table>
        <span class="ov-link">View Full Details &rarr;</span>
      </div>
"@
    }) -join "`n"

    $siteTabButtons = $(for ($i = 0; $i -lt $summaries.Count; $i++) {
        $s = $summaries[$i]
        $slug = ConvertTo-Slug $s.Site
        $tabColor = $tabPalette[$i % $tabPalette.Count]
        "<button class=`"tab`" id=`"tab-$slug`" onclick=`"showPage('$slug')`" style=`"background:$tabColor`">$(ConvertTo-HtmlSafe $s.Site)</button>"
    }) -join "`n    "

    # --- Per-site pages ---
    $sitePages = ($summaries | ForEach-Object {
        $s = $_
        $slug = ConvertTo-Slug $s.Site
        $color = $healthColor[$s.OverallHealth]

        $storagePanel = @"
        <div class="panel panel-full" style="border-top-color:#8a8f98">
          <h3><span class="n" style="background:#8a8f98">&#128451;&#65039;</span>Storage</h3>
          <div class="nodata-panel">
            <p class="nodata-text-big">There is no data provided yet</p>
          </div>
        </div>
"@

        if (-not $s.HasData) {
@"
      <section class="page" id="page-$slug" data-site="$(ConvertTo-HtmlSafe $s.Site)">
        <div class="site-hero" style="border-left-color:$color">
          <h1>$(ConvertTo-HtmlSafe $s.Site)</h1>
          <span class="badge big" style="background:$color">$(ConvertTo-HtmlSafe $healthLabelText.NoData)</span>
          <div class="meta">Report generated $(ConvertTo-HtmlSafe $RunDateDisplay)</div>
        </div>
        <div class="panel panel-full nodata-panel">
          <p class="nodata-text-big">There is no data provided yet</p>
          <p class="nodata-text-small">This site's Backup Summary tab hasn't been found in the supplied workbook yet. Once it's available, this page fills in the same way as AquaArabia and SixFlags.</p>
        </div>
$storagePanel
      </section>
"@
            return
        }

        $backupCards = ($s.BackupRows | ForEach-Object {
            $row = $_
            $pillText = switch ($row.Status) { 'Healthy' { 'Healthy' }; 'Warning' { 'Warning' }; 'Critical' { 'Critical' }; default { 'No Runs This Period' } }
@"
          <div class="backup-card">
            <div class="backup-card-head"><span class="pg-name">$(ConvertTo-HtmlSafe $row.ProtectionGroup)</span>$(Get-StatusPillHtml -Status $row.Status -Text $pillText)</div>
            <div class="backup-source">$(ConvertTo-HtmlSafe $row.Source)</div>
            <table class="metrics">
              <tr><td>Runs</td><td>$(ConvertTo-HtmlSafe $row.Runs)</td></tr>
              <tr><td>Objects (Success / Error)</td><td>$(ConvertTo-HtmlSafe $row.LastRunObjects)</td></tr>
              <tr><td>Data Read</td><td>$(ConvertTo-HtmlSafe $row.DataRead)</td></tr>
              <tr><td>SLA Violation</td><td>$(ConvertTo-HtmlSafe $row.SLAViolation)</td></tr>
              <tr><td>Bandwidth</td><td>$(ConvertTo-HtmlSafe $row.Bandwidth)</td></tr>
            </table>
          </div>
"@
        }) -join "`n"

@"
      <section class="page" id="page-$slug" data-site="$(ConvertTo-HtmlSafe $s.Site)">
        <div class="site-hero" style="border-left-color:$color">
          <h1>$(ConvertTo-HtmlSafe $s.Site)</h1>
          <span class="badge big" style="background:$color">$(ConvertTo-HtmlSafe $healthLabelText[$s.OverallHealth])</span>
          <div class="meta">Report generated $(ConvertTo-HtmlSafe $RunDateDisplay)</div>
        </div>

        <div class="risk-row big">
          <span class="risk risk-high">Critical Jobs: $($s.CriticalCount)</span>
          <span class="risk risk-med">Warning Jobs: $($s.WarningCount)</span>
        </div>

        <div class="tile-row">
          <div class="tile" style="background:$tileBlue"><span class="num">$($s.TotalGroups)</span><span class="label">Protection Groups</span></div>
          <div class="tile" style="background:#2e7d32"><span class="num">$($s.HealthyCount)</span><span class="label">Healthy Jobs</span></div>
          <div class="tile" style="background:#e6a100"><span class="num">$($s.WarningCount)</span><span class="label">Warning Jobs</span></div>
          <div class="tile" style="background:#c62828"><span class="num">$($s.CriticalCount)</span><span class="label">Critical Jobs</span></div>
          <div class="tile" style="background:$tilePurple"><span class="num">$($s.NoRunsCount)</span><span class="label">No Runs This Period</span></div>
        </div>

        <div class="panel panel-full" style="border-top-color:#1E3A5F">
          <h3><span class="n" style="background:#1E3A5F">&#128203;</span>Summary</h3>
          <p class="summary-text">$(ConvertTo-HtmlSafe $s.SummaryText)</p>
        </div>

        <div class="panel panel-full" style="border-top-color:$tileTeal">
          <h3><span class="n" style="background:$tileTeal">&#128190;</span>Backup Jobs <span style="font-weight:normal;font-size:14px;color:#999">($($s.TotalGroups) protection group(s))</span></h3>
          <div class="backup-grid">
$backupCards
          </div>
        </div>

$storagePanel
      </section>
"@
    }) -join "`n"

    $html = @"
<!DOCTYPE html>
<html lang="en">
<head>
<meta charset="UTF-8">
<title>Backup & Storage Weekly Health Check - Dashboard</title>
<style>
  body { font-family: Calibri, Arial, sans-serif; background:#eef1f5; color:#1a1a1a; margin:0; padding:0 32px 32px; font-size:16px; line-height:1.4; }
  h1 { color:#1E3A5F; margin:0; font-size:36px; }
  .subtitle { color:#555; margin:6px 0 20px; font-size:16px; }
  .tab { display:flex; align-items:center; justify-content:center; text-align:center; cursor:pointer; border:3px solid transparent; border-radius:10px; padding:28px 16px; font-family:inherit; font-weight:bold; font-size:18px; color:#fff; background:#1E3A5F; box-shadow:0 1px 4px rgba(0,0,0,0.2); min-height:90px; box-sizing:border-box; }
  .tab.active { border-color:#1a1a1a; box-shadow:0 0 0 3px rgba(0,0,0,0.15), 0 1px 4px rgba(0,0,0,0.2); }
  .dashboard-header { width:100%; text-align:left; background:#1E3A5F; border-radius:10px; padding:26px 32px; margin:24px 0 0; box-sizing:border-box; }
  .dashboard-header h1 { color:#fff; font-size:36px; }
  .dashboard-header .subtitle { color:rgba(255,255,255,0.8); margin:8px 0 0; font-size:16px; }
  .overview-box { display:flex; justify-content:center; align-items:center; width:100%; text-align:center; margin-top:16px; font-size:52px; letter-spacing:1px; padding:34px 32px; min-height:0; box-sizing:border-box; }
  .tabs { display:grid; grid-template-columns: repeat(auto-fit, minmax(170px, 1fr)); gap:20px; padding:24px 0 20px; position:sticky; top:0; background:#eef1f5; z-index:10; border-bottom:1px solid #dfe3e8; margin-bottom:32px; }
  .page { display:none; }
  .page.active { display:block; }

  .stat-row { display:grid; grid-template-columns: repeat(auto-fit, minmax(170px, 1fr)); gap:20px; margin-bottom:28px; }
  .stat { border-radius:10px; padding:18px 20px; text-align:center; box-shadow:0 1px 4px rgba(0,0,0,0.14); }
  .stat .num { font-size:32px; font-weight:bold; display:block; color:#fff; }
  .stat .label { font-size:14px; margin-top:4px; display:block; color:#fff; }
  .stat .stat-note { font-size:11px; margin-top:8px; color:rgba(255,255,255,0.85); line-height:1.4; }

  .ov-grid { display:grid; grid-template-columns: repeat(auto-fit, minmax(420px, 1fr)); gap:24px; }
  .ov-card { background:#fff; border-radius:10px; padding:24px 26px; box-shadow:0 1px 4px rgba(0,0,0,0.14); border-top:6px solid; cursor:pointer; }
  .ov-card:hover { box-shadow:0 4px 14px rgba(0,0,0,0.18); }
  .ov-head { display:flex; flex-direction:column; align-items:flex-start; gap:10px; margin-bottom:14px; }
  .ov-head h2 { margin:0; font-size:24px; color:#1E3A5F; }
  .ov-link { display:inline-block; margin-top:14px; color:#1E3A5F; font-weight:bold; font-size:14px; }
  .ov-card-nodata { opacity:0.85; }
  .nodata-text { color:#999; font-style:italic; margin:4px 0 0; }
  .badge { color:#fff; padding:6px 14px; border-radius:6px; font-size:14px; font-weight:bold; white-space:nowrap; display:inline-block; text-align:center; width:280px; }
  .badge.big { font-size:20px; padding:10px 22px; width:420px; }
  .risk-row { display:flex; gap:12px; margin-bottom:16px; flex-wrap:wrap; }
  .risk-row.big { margin:20px 0 28px; }
  .risk-row.big .risk { font-size:18px; padding:10px 20px; }
  .risk { font-size:14px; padding:5px 12px; border-radius:6px; font-weight:bold; }
  .risk-high { background:#fdecea; color:#c62828; }
  .risk-med  { background:#fff6e0; color:#8a6100; }
  .site-hero { border-left:8px solid; padding:14px 0 14px 26px; margin-bottom:8px; }
  .site-hero h1 { font-size:40px; }
  .site-hero .meta { color:#888; font-size:14px; margin-top:10px; }

  .tile-row { display:grid; grid-template-columns: repeat(auto-fit, minmax(170px, 1fr)); gap:20px; margin-bottom:28px; }
  .tile { border-radius:10px; padding:20px 18px; text-align:center; box-shadow:0 1px 4px rgba(0,0,0,0.14); }
  .tile .num { font-size:24px; font-weight:bold; display:block; color:#fff; }
  .tile .label { font-size:14px; margin-top:6px; display:block; color:#fff; }

  .panel { background:#fff; border-radius:10px; padding:26px 28px; box-shadow:0 1px 4px rgba(0,0,0,0.14); margin-bottom:24px; border-top:5px solid #ccc; }
  .status-pill { display:inline-flex; align-items:center; gap:6px; padding:5px 12px; border-radius:14px; font-size:13px; font-weight:bold; }
  .status-pill.ok { background:#e8f5e9; color:#1b5e20; }
  .status-pill.warn { background:#fff6e0; color:#8a6100; }
  .status-pill.bad { background:#fdecea; color:#c62828; }
  .status-pill.info { background:#eef2f5; color:#555; }
  .status-pill .dot { width:8px; height:8px; border-radius:50%; background:currentColor; }
  .panel h3 { margin:0 0 18px; color:#1E3A5F; font-size:19px; display:flex; align-items:center; gap:10px; }
  .panel h3 .n { color:#fff; width:34px; height:34px; border-radius:50%; display:inline-flex; align-items:center; justify-content:center; font-size:17px; flex-shrink:0; box-shadow:0 1px 3px rgba(0,0,0,0.25); }
  table.metrics { width:100%; border-collapse:collapse; font-size:16px; }
  table.metrics td { padding:10px 6px; border-bottom:1px solid #eee; vertical-align:top; }
  table.metrics td:first-child { color:#666; width:55%; }
  table.metrics td:last-child:not(:first-child) { text-align:right; font-weight:600; }
  .panel-full { grid-column: 1 / -1; }
  .summary-text { color:#444; font-size:15px; line-height:1.6; }

  /* Backup Jobs - one card per protection group */
  .backup-grid { display:grid; grid-template-columns: repeat(auto-fill, minmax(260px, 1fr)); gap:14px; }
  .backup-card { background:#fff; border:1px solid #eee; border-radius:8px; padding:16px 18px; box-shadow:0 1px 3px rgba(0,0,0,0.06); }
  .backup-card-head { display:flex; justify-content:space-between; align-items:center; gap:8px; margin-bottom:4px; }
  .backup-card .pg-name { font-weight:bold; font-size:15px; color:#1a1a1a; }
  .backup-card .backup-source { font-size:12px; color:#888; margin-bottom:10px; word-break:break-all; }
  .backup-card table.metrics td { font-size:13px; padding:7px 4px; }

  /* No-data placeholder (template site, or the Storage panel on every site until it's wired in) */
  .nodata-panel { text-align:center; padding:48px 24px; }
  .nodata-text-big { font-size:22px; font-weight:bold; color:#8a8f98; margin:0 0 10px; }
  .nodata-text-small { font-size:14px; color:#999; max-width:520px; margin:0 auto; line-height:1.6; }
</style>
</head>
<body>
  <div class="dashboard-header">
    <h1>Backup &amp; Storage Weekly Health Check - Dashboard</h1>
    <p class="subtitle">Generated $(ConvertTo-HtmlSafe $RunDateDisplay) - $($summaries.Count) site(s)</p>
  </div>

  <button type="button" class="tab overview-box active" id="tab-overview" onclick="showPage('overview')">Overview</button>

  <nav class="tabs">
    $siteTabButtons
  </nav>

  <section class="page active" id="page-overview">
    <div class="stat-row">
      <div class="stat" style="background:#2e7d32"><span class="num">$($healthCounts.Healthy)</span><span class="label">Healthy Sites</span>$healthyNoteHtml</div>
      <div class="stat" style="background:#e6a100"><span class="num">$($healthCounts.Warning)</span><span class="label">Sites with Warnings</span>$warningNoteHtml</div>
      <div class="stat" style="background:#c62828"><span class="num">$($healthCounts.Critical)</span><span class="label">Critical Sites</span>$criticalNoteHtml</div>
      <div class="stat" style="background:$tileBlue"><span class="num">$totalGroups</span><span class="label">Total Protection Groups</span>$groupsNoteHtml</div>
      <div class="stat" style="background:#c62828"><span class="num">$totalCritical</span><span class="label">Total Critical Backup Jobs</span>$criticalJobsNote</div>
      <div class="stat" style="background:#e6a100"><span class="num">$totalWarning</span><span class="label">Total Warning Backup Jobs</span>$warningJobsNote</div>
    </div>

    <div class="ov-grid">
$overviewCards
    </div>
  </section>

$sitePages

<script>
function showPage(slug) {
  document.querySelectorAll('.page').forEach(function(el){ el.classList.remove('active'); });
  document.querySelectorAll('.tab').forEach(function(el){ el.classList.remove('active'); });
  var page = document.getElementById('page-' + slug);
  var tab = document.getElementById('tab-' + slug);
  if (page) { page.classList.add('active'); }
  if (tab) { tab.classList.add('active'); }
  window.scrollTo(0, 0);
}
</script>
</body>
</html>
"@

    $dashboardPath = Join-Path $OutputPath "BackupStorage_HealthCheck_Dashboard_$RunDate.html"
    $html | Out-File -FilePath $dashboardPath -Encoding UTF8
    Write-Host "Dashboard written: $dashboardPath" -ForegroundColor Cyan
}

# ============================================================================
# 5. OUTPUT: DASHBOARD + LOG
# ============================================================================
$SiteLabels = @($SiteTabCandidates.Keys)
Invoke-SafeCheck -CheckName 'Dashboard generation' -Site 'n/a' -ObjectName 'Dashboard' -Script {
    Write-DashboardHtml -SiteLabels $SiteLabels -OutputPath $OutputPath -RunDateDisplay $RunDateDisplay
}

$LogPath = Join-Path $OutputPath "BackupStorage_Weekly_HealthCheck_$RunDate.log"
$LogLines = @()
$LogLines += "Backup & Storage Weekly Health Check run - $($ScriptStart.ToString('u')) - build $ScriptBuild"
$LogLines += "Sites processed: $($SiteLabels -join ', ')"
$LogLines += "Total findings collected: $($Global:AllResults.Count)"
$LogLines += "Total collection failures: $($Global:FailureLog.Count)"
$LogLines += "----------------------------------------------------------------"
foreach ($f in $Global:FailureLog) {
    $LogLines += "$($f.Timestamp) | $($f.Site) | $($f.Object) | $($f.Check) | $($f.Error)"
}
$LogLines | Out-File -FilePath $LogPath -Encoding UTF8
Write-Host "Log written: $LogPath" -ForegroundColor Cyan

Write-Host "`nDone. Findings: $($Global:AllResults.Count) | Failures: $($Global:FailureLog.Count) | Duration: $([Math]::Round(((Get-Date)-$ScriptStart).TotalMinutes,1)) min" -ForegroundColor Green
