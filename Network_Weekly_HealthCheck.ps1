#Requires -Version 5.1
<#
.SYNOPSIS
    Network Weekly Health Check - reads the IDF Room Report workbooks for AquaArabia and
    SixFlags and produces a single combined HTML dashboard, in the same visual style as
    VMware_Weekly_HealthCheck.ps1's dashboard.

.DESCRIPTION
    This dashboard covers Network (IDF room/cabinet/access-control) data only. The DC RackWise
    Health Check data (compute racks) and Backup & Storage are separate dashboards, not part of
    this script.

    Reads 1 source workbook per site - the IDF Room Report - whose single "Summary" tab lists
    ONLY the issues found (an exceptions list, not a full per-switch checklist): one row per
    issue, each naming the affected zone/location code(s). The two companies word this slightly
    differently (AquaArabia: "Issues" + a free-text "Zone N: loc1, loc2" cell per issue type;
    SixFlags: "No." + "Observation" + a flat comma-separated "Zone - IDF(s)" list per
    observation) but both are read the same way: find the issue-label column, the affected-
    location column is always the one immediately to its right.

    Every tab is located by searching for its known column headers rather than assuming a
    fixed row/column position, since the real workbooks have inconsistent header rows (a
    banner/logo above the real header, blank spacer rows) - a tab whose headers can't be found
    is skipped and logged, never guessed at.

    Produces ONE combined HTML dashboard (no Word/PDF) covering both sites, matching the
    VMware dashboard's look: rectangular site tabs (sized to align with the KPI tile row below
    them), KPI stat tiles, and an IDF Room Issues panel per site.

    This script never embeds evidence photos - that is handled by a separate, optional script
    (Network_Evidence_Photos.ps1) run afterward against the dashboard this script produces, so
    a month with no photos supplied needs no change here.

    NEVER writes/modifies the source workbooks - opened read-only, closed without saving.

.PARAMETER AquaArabiaIdfPath
    Path to AquaArabia's IDF Room Report workbook. Omit to skip AquaArabia.
.PARAMETER SixFlagsIdfPath
    Path to SixFlags' IDF Room Report workbook. Omit to skip SixFlags.
.PARAMETER OutputPath
    Folder the dashboard + log are written to.
#>

[CmdletBinding()]
param(
    [string]$AquaArabiaIdfPath = '',
    [string]$SixFlagsIdfPath = '',

    [string]$OutputPath = (Join-Path $PSScriptRoot "Network_HealthCheck_Reports")
)

$ScriptBuild = '2026-10-01-03-network-only'
Write-Host "Network_Weekly_HealthCheck.ps1 - build $ScriptBuild" -ForegroundColor Magenta

$ErrorActionPreference = 'Stop'
$ScriptStart = Get-Date
$RunDate = $ScriptStart.ToString('yyyy-MM-dd')
$RunDateDisplay = $ScriptStart.ToString('d-MMMM-yyyy', [System.Globalization.CultureInfo]::InvariantCulture)
if (-not (Test-Path $OutputPath)) { New-Item -ItemType Directory -Path $OutputPath -Force | Out-Null }

# ============================================================================
# 0. LOGGING / SAFE-EXECUTION HELPERS (same pattern as VMware_Weekly_HealthCheck.ps1)
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
        [string]$Item, [string]$Value, [string]$Status, [string]$Notes = '', [string]$RowNum = ''
    )
    # RowNum: the IDF Summary row this finding came from (used only to match evidence photos,
    # which are supplied separately and named by row number - see Network_Evidence_Photos.ps1).
    $obj = [pscustomobject]@{
        Site = $Site; Area = $Area; Group = $Group; Object = $Object
        Item = $Item; Value = $Value; Status = $Status; Notes = $Notes; RowNum = $RowNum
    }
    $Global:AllResults.Add($obj) | Out-Null
}

# Collapses a multi-line cell (real Comments cells contain embedded newlines) into one compact
# line for display.
function ConvertTo-CompactNote {
    param([string]$Text)
    if ([string]::IsNullOrWhiteSpace($Text)) { return '' }
    return ($Text -replace '\r?\n', ' | ').Trim()
}

# ============================================================================
# 1. EXCEL HEADER-ANCHORING HELPERS
#    Every tab is located by searching for its known header text rather than assuming a fixed
#    row/column - the real workbooks have inconsistent header rows (a banner/logo above the
#    real header, blank spacer rows) that make a fixed-position read silently wrong instead of
#    simply failing.
# ============================================================================
function Find-HeaderCellContains {
    param($Worksheet, [string]$Contains, [int]$MaxRow = 20, [int]$MaxCol = 40)
    for ($r = 1; $r -le $MaxRow; $r++) {
        for ($c = 1; $c -le $MaxCol; $c++) {
            $v = $Worksheet.Cells.Item($r, $c).Value2
            if ($null -ne $v -and "$v" -match [regex]::Escape($Contains)) {
                return [pscustomobject]@{ Row = $r; Col = $c }
            }
        }
    }
    return $null
}

# ============================================================================
# 2. IDF ROOM REPORT COLLECTION (exceptions-list format)
#    The real IDF Room Report workbooks have no per-zone checklist tabs at all - just a single
#    "Summary" tab listing only the issues found, each naming the affected zone/location
#    code(s). AquaArabia words the issue-label column "Issues" (issue TYPE in each row, e.g.
#    "Rodent Trap Missing", with a multi-line "Zone N: loc1, loc2" cell listing every affected
#    location); SixFlags words it "Observation" (a free-text description per row, e.g. "Access
#    control is not working", with a flat comma-separated zone/IDF-code list and no "Zone N:"
#    grouping). Both are read the same way: locate the issue-label column, and the affected-
#    location list is always the column immediately to its right.
# ============================================================================
function Read-IdfIssuesWorkbook {
    param($Excel, [string]$Site, [string]$Path)
    if (-not $Path -or -not (Test-Path $Path)) {
        if ($Path) { Write-CheckLog -Site $Site -Object $Path -CheckName 'IDF workbook' -ErrorMessage 'File not found.' }
        return
    }
    $wb = $Excel.Workbooks.Open($Path, 0, $true)
    try {
        $ws = $wb.Worksheets | Where-Object { $_.Name.Trim() -ieq 'Summary' } | Select-Object -First 1
        if (-not $ws) { $ws = $wb.Worksheets.Item(1) }
        Invoke-SafeCheck -Site $Site -ObjectName $ws.Name -CheckName 'IDF issues tab' -Script {
            $issueCell = Find-HeaderCellContains -Worksheet $ws -Contains 'Issues'
            if (-not $issueCell) { $issueCell = Find-HeaderCellContains -Worksheet $ws -Contains 'Observation' }
            if (-not $issueCell) { throw "Could not locate an 'Issues' or 'Observation' column header on this tab." }
            $issueCol = $issueCell.Col
            $locCol = $issueCol + 1
            $hdrRow = $issueCell.Row

            # Optional 'No.' row-number column, one column to the left of the issue column.
            $noCol = $null
            if ($issueCol -gt 1) {
                $noVal0 = $ws.Cells.Item($hdrRow, $issueCol - 1).Value2
                if ($null -ne $noVal0 -and "$noVal0".Trim() -ieq 'No.') { $noCol = $issueCol - 1 }
            }

            # Optional 'Comments' column, searched a few columns right of the location column.
            $commentsCol = $null
            for ($c = $locCol + 1; $c -le ($locCol + 5); $c++) {
                $v = $ws.Cells.Item($hdrRow, $c).Value2
                if ($null -ne $v -and "$v".Trim() -ieq 'Comments') { $commentsCol = $c; break }
            }

            $used = $ws.UsedRange
            $lastRow = $used.Row + $used.Rows.Count - 1
            $rowCounter = 0
            for ($r = $hdrRow + 1; $r -le $lastRow; $r++) {
                $issueVal = $ws.Cells.Item($r, $issueCol).Value2
                if (-not $issueVal -or "$issueVal".Trim() -eq '') { continue }
                $issueText = "$issueVal".Trim()
                $rowCounter++
                $rowLabel = "$rowCounter"
                if ($noCol) {
                    $nv = $ws.Cells.Item($r, $noCol).Value2
                    if ($nv -and "$nv".Trim() -ne '') { $rowLabel = "$nv".Trim() }
                }

                $locVal = $ws.Cells.Item($r, $locCol).Value2
                $locText = if ($locVal) { "$locVal" } else { '' }
                $commentsVal = if ($commentsCol) { ConvertTo-CompactNote "$($ws.Cells.Item($r, $commentsCol).Value2)" } else { '' }

                if (-not $locText.Trim()) {
                    New-Finding -Site $Site -Area 'IDF' -Group '' -Object '(unspecified location)' `
                        -Item $issueText -Value 'Reported' -Status 'Warning' -Notes $commentsVal -RowNum $rowLabel
                    continue
                }

                # AquaArabia-style: one "Zone N: loc1, loc2" line per affected zone (multi-line cell).
                # SixFlags-style: a single line of flat comma-separated location codes, no zone prefix.
                $lines = $locText -split '\r?\n' | Where-Object { $_.Trim() -ne '' }
                foreach ($line in $lines) {
                    if ($line -match '^\s*([^:]+):\s*(.+)$') {
                        $zoneLabel = $Matches[1].Trim()
                        $locations = $Matches[2] -split ',' | ForEach-Object { $_.Trim() } | Where-Object { $_ -ne '' }
                    } else {
                        $zoneLabel = ''
                        $locations = $line -split ',' | ForEach-Object { $_.Trim() } | Where-Object { $_ -ne '' }
                    }
                    foreach ($loc in $locations) {
                        New-Finding -Site $Site -Area 'IDF' -Group $zoneLabel -Object $loc `
                            -Item $issueText -Value 'Reported' -Status 'Warning' -Notes $commentsVal -RowNum $rowLabel
                    }
                }
            }
        }
    } finally {
        $wb.Close($false)
        [System.Runtime.Interopservices.Marshal]::ReleaseComObject($wb) | Out-Null
    }
}

# ============================================================================
# 3. RUN COLLECTION
# ============================================================================
$Sites = @(
    [pscustomobject]@{ Site = 'AquaArabia'; IdfPath = $AquaArabiaIdfPath }
    [pscustomobject]@{ Site = 'SixFlags';   IdfPath = $SixFlagsIdfPath }
)

$Excel = $null
try {
    $Excel = New-Object -ComObject Excel.Application
    $Excel.Visible = $false
    $Excel.DisplayAlerts = $false
} catch {
    throw "Microsoft Excel is not available via COM automation on this machine - cannot read the source workbooks. $($_.Exception.Message)"
}

foreach ($s in $Sites) {
    Write-Host "`n=== Collecting: $($s.Site) ===" -ForegroundColor Green
    Read-IdfIssuesWorkbook -Excel $Excel -Site $s.Site -Path $s.IdfPath
}

$Excel.Quit()
[System.Runtime.Interopservices.Marshal]::ReleaseComObject($Excel) | Out-Null

# ============================================================================
# 4. DASHBOARD GENERATION HELPERS
# ============================================================================
function ConvertTo-HtmlSafe {
    param([string]$Text)
    if (-not $Text) { return '' }
    return $Text -replace '&','&amp;' -replace '<','&lt;' -replace '>','&gt;' -replace '"','&quot;'
}

function ConvertTo-Slug {
    param([string]$Text)
    $slug = ($Text -replace '[^a-zA-Z0-9]+', '-').Trim('-').ToLower()
    if (-not $slug) { return 'site' }
    return $slug
}

# Computes every per-site aggregate the dashboard needs directly from $Global:AllResults - the
# IDF issues list, grouped back up by the Summary row each finding came from.
function Get-SiteDashboardSummary {
    param([string]$SiteLabel)
    $SiteFindings = $Global:AllResults | Where-Object { $_.Site -eq $SiteLabel }
    if (-not $SiteFindings) { return $null }

    $WarnCount = @($SiteFindings | Where-Object { $_.Status -eq 'Warning' }).Count
    $OverallHealth = if ($WarnCount -gt 0) { 'Warning' } else { 'Healthy' }

    $idfIssueRows = @($SiteFindings | Group-Object RowNum | ForEach-Object {
        $grp = $_.Group
        $first = $grp[0]
        $locText = ($grp | ForEach-Object {
            if ($_.Object -eq '(unspecified location)') { $null }
            elseif ($_.Group) { "$($_.Group): $($_.Object)" }
            else { $_.Object }
        } | Where-Object { $_ }) -join '; '
        if (-not $locText) { $locText = '(no location specified)' }
        [pscustomobject]@{
            RowNum    = $first.RowNum
            Issue     = $first.Item
            Locations = $locText
            Notes     = $first.Notes
        }
    } | Sort-Object { [int]($_.RowNum -replace '\D','0') })
    $idfIssueCount = $idfIssueRows.Count
    $idfLocationCount = @($SiteFindings | Where-Object { $_.Object -ne '(unspecified location)' } | Select-Object -ExpandProperty Object -Unique).Count

    $SummaryText = if ($WarnCount -gt 0) {
        "$SiteLabel has $idfIssueCount IDF issue(s) reported across $idfLocationCount location(s) - see the list below."
    } else {
        "$SiteLabel's IDF rooms are all reporting healthy with no issues flagged."
    }

    [pscustomobject]@{
        Site             = $SiteLabel
        OverallHealth    = $OverallHealth
        MediumRisk       = $WarnCount
        IdfIssueCount    = $idfIssueCount
        IdfLocationCount = $idfLocationCount
        IdfIssueRows     = $idfIssueRows
        SummaryText      = $SummaryText
    }
}

function Write-DashboardHtml {
    param([string[]]$SiteLabels, [string]$OutputPath, [string]$RunDateDisplay)

    $summaries = @($SiteLabels | ForEach-Object { Get-SiteDashboardSummary -SiteLabel $_ } | Where-Object { $_ } | Sort-Object Site)
    if ($summaries.Count -eq 0) { return }

    $healthColor = @{ Healthy = '#2e7d32'; Warning = '#e6a100' }
    $healthLabelText = @{ Healthy = 'Healthy - No Issues Detected'; Warning = 'Healthy - Minor Issues Detected' }
    $healthCounts = @{
        Healthy = @($summaries | Where-Object { $_.OverallHealth -eq 'Healthy' }).Count
        Warning = @($summaries | Where-Object { $_.OverallHealth -eq 'Warning' }).Count
    }
    $totalMedium    = ($summaries | Measure-Object -Property MediumRisk -Sum).Sum
    $totalIdfIssues = ($summaries | Measure-Object -Property IdfIssueCount -Sum).Sum
    $totalLocations = ($summaries | Measure-Object -Property IdfLocationCount -Sum).Sum

    $tabPalette = @('#2563EB','#7C3AED','#0D9488','#C026D3','#EA580C','#4F46E5','#DB2777','#0EA5E9')
    $tileBlue = '#1565C0'; $tilePurple = '#6A1B9A'

    # --- Overview page ---
    $overviewCards = ($summaries | ForEach-Object {
        $s = $_
        $slug = ConvertTo-Slug $s.Site
        $color = $healthColor[$s.OverallHealth]
@"
      <div class="ov-card" onclick="showPage('$slug')" style="border-top-color:$color">
        <div class="ov-head"><h2>$(ConvertTo-HtmlSafe $s.Site)</h2><span class="badge" style="background:$color">$(ConvertTo-HtmlSafe $healthLabelText[$s.OverallHealth])</span></div>
        <table class="metrics">
          <tr><td>IDF Issues Reported</td><td>$($s.IdfIssueCount)</td></tr>
          <tr><td>Locations Affected</td><td>$($s.IdfLocationCount)</td></tr>
        </table>
        <span class="ov-link">View Full Details &rarr;</span>
      </div>
"@
    }) -join "`n"

    $overviewTabButton = '<button class="tab overview-tab active" id="tab-overview" onclick="showPage(''overview'')">Overview</button>'
    $siteTabButtons = $(for ($i = 0; $i -lt $summaries.Count; $i++) {
        $s = $summaries[$i]
        $slug = ConvertTo-Slug $s.Site
        $tabColor = $tabPalette[$i % $tabPalette.Count]
        "<button class=`"tab`" id=`"tab-$slug`" onclick=`"showPage('$slug')`" style=`"background:$tabColor`">$(ConvertTo-HtmlSafe $s.Site)</button>"
    }) -join "`n    "
    $tabButtons = @($overviewTabButton, $siteTabButtons) -join "`n    "

    # --- Per-site pages ---
    $sitePages = ($summaries | ForEach-Object {
        $s = $_
        $slug = ConvertTo-Slug $s.Site
        $color = $healthColor[$s.OverallHealth]

        # Each row carries a data-idf-row attribute so Network_Evidence_Photos.ps1 (a separate,
        # optional script) can find the right insertion point when photos are supplied for a
        # given month - this script itself never depends on photos existing.
        $idfRows = if ($s.IdfIssueRows.Count -gt 0) {
            ($s.IdfIssueRows | ForEach-Object {
@"
          <tr data-idf-row="$(ConvertTo-HtmlSafe $_.RowNum)"><td>$(ConvertTo-HtmlSafe $_.RowNum)</td><td>$(ConvertTo-HtmlSafe $_.Issue)</td><td>$(ConvertTo-HtmlSafe $_.Locations)</td>
            <td class="idf-evidence-cell"></td></tr>
"@
            }) -join "`n"
        } else { "<tr><td colspan='4' style='color:#999'>No IDF issues reported</td></tr>" }

@"
      <section class="page" id="page-$slug" data-site="$(ConvertTo-HtmlSafe $s.Site)">
        <div class="site-hero" style="border-left-color:$color">
          <h1>$(ConvertTo-HtmlSafe $s.Site)</h1>
          <span class="badge big" style="background:$color">$(ConvertTo-HtmlSafe $healthLabelText[$s.OverallHealth])</span>
          <div class="meta">Report generated $(ConvertTo-HtmlSafe $RunDateDisplay)</div>
        </div>

        <div class="tile-row">
          <div class="tile" style="background:$tileBlue"><span class="num">$($s.IdfIssueCount)</span><span class="label">IDF Issues Reported</span></div>
          <div class="tile" style="background:$tilePurple"><span class="num">$($s.IdfLocationCount)</span><span class="label">Locations Affected</span></div>
        </div>

        <div class="panel-grid">
          <div class="panel panel-full" style="border-top-color:$tilePurple">
            <h3><span class="n" style="background:$tilePurple">&#127968;</span>IDF Room Issues <span style="font-weight:normal;font-size:14px;color:#999">($($s.IdfIssueCount) reported)</span></h3>
            <div class="scroll-box">
              <table class="metrics wide">
                <thead><tr><th>No.</th><th>Issue</th><th>Affected Location(s)</th><th>Evidence</th></tr></thead>
                <tbody>
$idfRows
                </tbody>
              </table>
            </div>
          </div>

          <div class="panel" style="border-top-color:#1E3A5F">
            <h3><span class="n" style="background:#1E3A5F">&#128203;</span>Summary</h3>
            <p class="summary-text">$(ConvertTo-HtmlSafe $s.SummaryText)</p>
          </div>
        </div>
      </section>
"@
    }) -join "`n"

    $html = @"
<!DOCTYPE html>
<html lang="en">
<head>
<meta charset="UTF-8">
<title>Network Weekly Health Check - Dashboard</title>
<style>
  body { font-family: Calibri, Arial, sans-serif; background:#eef1f5; color:#1a1a1a; margin:0; padding:0 32px 32px; font-size:16px; line-height:1.4; }
  h1 { color:#1E3A5F; margin:0; font-size:36px; }
  .subtitle { color:#555; margin:6px 0 20px; font-size:16px; }
  .tabs { display:grid; grid-template-columns: repeat(auto-fit, minmax(170px, 1fr)); gap:20px; padding:24px 0 20px; position:sticky; top:0; background:#eef1f5; z-index:10; border-bottom:1px solid #dfe3e8; margin-bottom:32px; }
  .tab { display:flex; align-items:center; justify-content:center; text-align:center; cursor:pointer; border:3px solid transparent; border-radius:10px; padding:28px 16px; font-family:inherit; font-weight:bold; font-size:18px; color:#fff; background:#1E3A5F; box-shadow:0 1px 4px rgba(0,0,0,0.2); min-height:90px; box-sizing:border-box; }
  .tab.overview-tab { background:#1E3A5F; }
  .tab.active { border-color:#1a1a1a; box-shadow:0 0 0 3px rgba(0,0,0,0.15), 0 1px 4px rgba(0,0,0,0.2); }
  .page { display:none; }
  .page.active { display:block; }

  .stat-row { display:grid; grid-template-columns: repeat(auto-fit, minmax(170px, 1fr)); gap:20px; margin-bottom:28px; }
  .stat { border-radius:10px; padding:18px 20px; text-align:center; box-shadow:0 1px 4px rgba(0,0,0,0.14); }
  .stat .num { font-size:32px; font-weight:bold; display:block; color:#fff; }
  .stat .label { font-size:14px; margin-top:4px; display:block; color:#fff; }

  .ov-grid { display:grid; grid-template-columns: repeat(auto-fit, minmax(420px, 1fr)); gap:24px; }
  .ov-card { background:#fff; border-radius:10px; padding:24px 26px; box-shadow:0 1px 4px rgba(0,0,0,0.14); border-top:6px solid; cursor:pointer; }
  .ov-card:hover { box-shadow:0 4px 14px rgba(0,0,0,0.18); }
  .ov-head { display:flex; flex-direction:column; align-items:flex-start; gap:10px; margin-bottom:14px; }
  .ov-head h2 { margin:0; font-size:24px; color:#1E3A5F; }
  .ov-link { display:inline-block; margin-top:14px; color:#1E3A5F; font-weight:bold; font-size:14px; }
  .badge { color:#fff; padding:6px 14px; border-radius:6px; font-size:14px; font-weight:bold; white-space:nowrap; display:inline-block; text-align:center; width:280px; }
  .badge.big { font-size:20px; padding:10px 22px; width:420px; }
  .site-hero { border-left:8px solid; padding:14px 0 14px 26px; margin-bottom:28px; }
  .site-hero h1 { font-size:40px; }
  .site-hero .meta { color:#888; font-size:14px; margin-top:10px; }

  .tile-row { display:grid; grid-template-columns: repeat(auto-fit, minmax(170px, 1fr)); gap:20px; margin-bottom:28px; }
  .tile { border-radius:10px; padding:20px 18px; text-align:center; box-shadow:0 1px 4px rgba(0,0,0,0.14); }
  .tile .num { font-size:24px; font-weight:bold; display:block; color:#fff; }
  .tile .label { font-size:14px; margin-top:6px; display:block; color:#fff; }

  .panel-grid { display:grid; grid-template-columns: repeat(auto-fit, minmax(380px, 1fr)); gap:24px; margin-bottom:24px; }
  .panel { background:#fff; border-radius:10px; padding:26px 28px; box-shadow:0 1px 4px rgba(0,0,0,0.14); margin-bottom:24px; border-top:5px solid #ccc; }
  .panel h3 { margin:0 0 18px; color:#1E3A5F; font-size:19px; display:flex; align-items:center; gap:10px; }
  .panel h3 .n { color:#fff; width:34px; height:34px; border-radius:50%; display:inline-flex; align-items:center; justify-content:center; font-size:17px; flex-shrink:0; box-shadow:0 1px 3px rgba(0,0,0,0.25); }
  table.metrics { width:100%; border-collapse:collapse; font-size:16px; }
  table.metrics td, table.metrics th { padding:10px 6px; border-bottom:1px solid #eee; vertical-align:top; }
  table.metrics td:first-child { color:#666; width:55%; }
  table.metrics td:last-child:not(:first-child) { text-align:right; font-weight:600; }
  table.metrics.wide th { text-align:center; color:#999; font-size:13px; text-transform:uppercase; padding-bottom:10px; }
  table.metrics.wide th:first-child { text-align:left; }
  table.metrics.wide td { text-align:center; }
  table.metrics.wide td:first-child { text-align:left; color:#333; font-weight:600; width:auto; }
  table.metrics.wide td:last-child:not(:first-child) { text-align:center; font-weight:normal; }
  .scroll-box { max-height:420px; overflow-y:auto; border:1px solid #f0f0f0; border-radius:6px; }
  .scroll-box table.metrics.wide { font-size:15px; }
  .scroll-box thead th { position:sticky; top:0; background:#fff; }
  .panel-full { grid-column: 1 / -1; }
  .summary-text { color:#444; font-size:15px; line-height:1.6; }
  .idf-evidence-cell img { max-width:90px; max-height:90px; border-radius:4px; box-shadow:0 1px 3px rgba(0,0,0,0.3); cursor:zoom-in; }
</style>
</head>
<body>
  <h1>Network Weekly Health Check - Dashboard</h1>
  <p class="subtitle">Generated $(ConvertTo-HtmlSafe $RunDateDisplay) - $($summaries.Count) site(s)</p>

  <nav class="tabs">
    $tabButtons
  </nav>

  <section class="page active" id="page-overview">
    <div class="stat-row">
      <div class="stat" style="background:#2e7d32"><span class="num">$($healthCounts.Healthy)</span><span class="label">Healthy Sites</span></div>
      <div class="stat" style="background:#e6a100"><span class="num">$($healthCounts.Warning)</span><span class="label">Sites with Issues</span></div>
      <div class="stat" style="background:#4F46E5"><span class="num">$totalMedium</span><span class="label">Total Issues Reported</span></div>
      <div class="stat" style="background:#1565C0"><span class="num">$totalIdfIssues</span><span class="label">Total IDF Issues</span></div>
      <div class="stat" style="background:#6A1B9A"><span class="num">$totalLocations</span><span class="label">Total Locations Affected</span></div>
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

    $dashboardPath = Join-Path $OutputPath "Network_HealthCheck_Dashboard_$RunDate.html"
    $html | Out-File -FilePath $dashboardPath -Encoding UTF8
    Write-Host "Dashboard written: $dashboardPath" -ForegroundColor Cyan
}

# ============================================================================
# 5. OUTPUT: DASHBOARD + LOG
# ============================================================================
$SiteLabels = $Global:AllResults | Select-Object -ExpandProperty Site -Unique
Invoke-SafeCheck -CheckName 'Dashboard generation' -Site 'n/a' -ObjectName 'Dashboard' -Script {
    Write-DashboardHtml -SiteLabels $SiteLabels -OutputPath $OutputPath -RunDateDisplay $RunDateDisplay
}

$LogPath = Join-Path $OutputPath "Network_Weekly_HealthCheck_$RunDate.log"
$LogLines = @()
$LogLines += "Network Weekly Health Check run - $($ScriptStart.ToString('u')) - build $ScriptBuild"
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
