<#
.SYNOPSIS
    Generates a professional, self-contained HTML executive dashboard for a
    monthly Windows Server patching activity, using an Excel workbook as the
    ONLY source of truth.

.DESCRIPTION
    The workflow is:  Excel sheet  ->  this script  ->  Executive dashboard.

    The script never asks you to type the Site Name, Engineer / Implementer,
    server counts, statuses, dates or any other reporting figure. Every value
    is read and calculated directly from the workbook.

        * Site Name              -> read from the "Site Name" column
        * Engineer / Implementer -> read from the "Update Engineer" column
        * Patching month / date  -> derived from the "Start Patching Date" column
        * Server list, OS, owners -> read row by row
        * All KPIs and %          -> calculated dynamically

    The original workbook is opened READ-ONLY (via the ImportExcel module) and
    is never modified. If the file is locked (open in Excel) the script
    transparently works on a temporary copy.

.PARAMETER InputExcel
    Mandatory. Full path to the monthly patching workbook (.xlsx / .xlsm).

.PARAMETER OutputFolder
    Optional. Folder for the generated report(s) and log file.
    Default: a "PatchingReports" sub-folder next to the Excel file.

.PARAMETER WorksheetName
    Optional. Force a specific worksheet. If omitted the script auto-selects the
    sheet that actually contains the patching data (HostName + Status columns).

.PARAMETER HeaderRow
    Optional. Row number that holds the column headers. Default: 1.

.PARAMETER SiteName
    Optional OVERRIDE only. Use it if the workbook has no Site column at all.
    Never required for normal runs.

.PARAMETER EngineerName
    Optional OVERRIDE only. Same idea as -SiteName.

.PARAMETER PatchingMonth
    Optional OVERRIDE only, e.g. "September 2026". Normally derived from the dates.

.PARAMETER ComplianceThreshold
    Completion %% at or above which the site is considered "On Track". This only
    drives the wording of the "Overall Status" card - no compliance figure or
    gauge is shown on this per-site dashboard (that belongs in the separate
    consolidated dashboard). Default: 95.

.PARAMETER ColumnMapPath
    Optional path to a JSON file that adds/overrides the column-name aliases,
    e.g.  { "HostName": ["Node","CI Name"], "Status": ["Patch State"] }

.PARAMETER CombineSites
    If the workbook contains more than one site, produce ONE combined dashboard
    titled "MULTIPLE SITES" instead of one dashboard per site (the default).

.PARAMETER ExportPdf
    Also produce a PDF next to the HTML (uses headless Microsoft Edge or Chrome
    if available; best effort).

.PARAMETER Open
    Open each generated HTML report when finished.

.EXAMPLE
    .\Generate-PatchingDashboard.ps1 -InputExcel "C:\Reports\September-Patching.xlsx"

.EXAMPLE
    .\Generate-PatchingDashboard.ps1 -InputExcel "C:\Patching\September.xlsx" -OutputFolder "C:\Patching\Reports" -ExportPdf

.NOTES
    Requires PowerShell 5.1+ and the ImportExcel module (no Microsoft Excel /
    COM needed).
        Install-Module ImportExcel -Scope CurrentUser
#>

#Requires -Version 5.1
[CmdletBinding()]
param(
    [Parameter(Mandatory = $true, Position = 0)]
    [ValidateNotNullOrEmpty()]
    [string]$InputExcel,

    [string]$OutputFolder,

    [string]$WorksheetName,

    [ValidateRange(1, 100)]
    [int]$HeaderRow = 1,

    [string]$SiteName,        # optional override only
    [string]$EngineerName,    # optional override only
    [string]$PatchingMonth,   # optional override only  (e.g. "September 2026")

    [ValidateRange(0, 100)]
    [double]$ComplianceThreshold = 95.0,

    [string]$ColumnMapPath,

    [switch]$CombineSites,
    [switch]$ExportPdf,
    [switch]$Open
)

$ErrorActionPreference = 'Stop'
$script:LogFile = $null
$script:ExitCode = 0

# ======================================================================
# 1. COLUMN NAME ALIASES  --  adjust here if your headers differ
# ----------------------------------------------------------------------
# Canonical field  =  list of header names that may appear in the sheet.
# Matching is case-insensitive and ignores spaces / punctuation, and also
# matches when the real header simply STARTS WITH the alias (handy for
# long headers such as "OS(Provide Image) || VA(Provide OVA)").
# ======================================================================
$ColumnAliases = [ordered]@{
    SiteName       = @('Site Name', 'Site', 'SiteName', 'Location', 'Data Center', 'Datacenter', 'DC', 'Region', 'Customer', 'Client', 'Account', 'Project')
    BusinessOwner  = @('Business Owner', 'Business Owners', 'App Owner', 'Application Owner', 'Service Owner', 'Product Owner')
    TechnicalOwner = @('Technical Owner', 'Tech Owner', 'System Owner', 'Server Owner', 'Infra Owner', 'Administrator', 'Admin', 'Support Owner')
    HostName       = @('HostName', 'Host Name', 'Host', 'Server Name', 'Server', 'Servername', 'Computer Name', 'ComputerName', 'Machine', 'Node', 'CI Name', 'Device')
    IPAddress      = @('IP Address', 'IPAddress', 'IP', 'IP Addr', 'Address', 'Mgmt IP', 'Management IP', 'Primary IP')
    OS             = @('OS', 'Operating System', 'OS Name', 'OS Version', 'Platform', 'Image', 'OS(Provide Image)', 'OS (Provide Image)')
    Engineer       = @('Update Engineer', 'Engineer', 'Implementer', 'Implemented By', 'Patching Engineer', 'Patch Engineer', 'Performed By', 'Executed By', 'Assigned To', 'Owner Engineer', 'Resource', 'Applied By')
    StartDate      = @('Start Patching Date', 'Patching Date', 'Patch Date', 'Start Date', 'Scheduled Date', 'Patching Start Date', 'Completion Date', 'Date', 'Activity Date', 'Maintenance Date')
    Status         = @('Patching Status', 'Patch Status', 'Status', 'Result', 'Outcome', 'State', 'Compliance Status', 'Patch State', 'Activity Status')
}

# Fields that MUST be present for the script to run.
$RequiredFields = @('HostName', 'Status')

# ======================================================================
# 1b. SITE CODE -> FULL SITE NAME  --  add / edit your codes here
# ----------------------------------------------------------------------
# If the workbook (Site column) or the Excel file name contains a short
# code, it is expanded to the full site name for the dashboard header.
# Matching is case-insensitive; the code may appear on its own ("SF") or
# as a separate token ("SF - Prod", "Site: SF", "SixFlags(SF)").
# Unknown values are used exactly as they appear in the sheet.
# ======================================================================
$SiteNameMap = @{
    'SF'  = 'Six Flags'
    'AQ'  = 'Aquarabia'
    'AQA' = 'Aquarabia'
    'AMC' = 'AMC'
    'TB'  = 'SEVEN Tabuk'
    'AB'  = 'SEVEN Abha'
    'HA'  = 'SEVEN Alhamra'
}

# ======================================================================
# 2. STATUS NORMALISATION  --  adjust the keyword lists if needed
# ----------------------------------------------------------------------
# Every raw status is lower-cased and stripped of non-alphanumerics, then
# tested (as a substring) against these lists, in this priority order:
#   negative-completed  ->  Pending
#   Failed keywords     ->  Failed
#   Completed keywords  ->  Completed
#   Pending keywords    ->  Pending
#   anything else / blank ->  Unknown   (NEVER counted as completed)
# ======================================================================
$StatusKeywords = @{
    Failed    = @('failed', 'fail', 'failure', 'error', 'errored', 'unsuccessful', 'aborted', 'rollback', 'rolledback', 'cancelled', 'canceled')
    Completed = @('completed', 'complete', 'done', 'success', 'successful', 'successfullypatched', 'patchedsuccessfully', 'patched', 'patchingcompleted', 'installed', 'applied', 'closed', 'resolved', 'compliant', 'uptodate', 'fullypatched', 'ok', 'pass', 'passed', 'green')
    Pending   = @('pending', 'notstarted', 'yettostart', 'tobestarted', 'tobedone', 'scheduled', 'planned', 'inprogress', 'wip', 'workinprogress', 'ongoing', 'incomplete', 'partiallycompleted', 'partial', 'onhold', 'hold', 'deferred', 'postponed', 'rescheduled', 'excluded', 'skipped', 'queued', 'inqueue', 'open', 'new', 'rebootpending', 'pendingreboot', 'amber', 'yellow')
}

# ======================================================================
# 3. PALETTE
# ======================================================================
$Palette = @{
    Green   = '#1F8A4C'
    Amber   = '#C77700'
    Red     = '#C0392B'
    Grey    = '#6B7683'
    Ink     = '#1F2933'
    Line    = '#E3E8EF'
    SlateHi = '#1D4E79'
    SlateLo = '#0F2A43'
    Bg      = '#F4F6F9'
}

# ======================================================================
#  HELPER FUNCTIONS
# ======================================================================
function Write-Log {
    param(
        [Parameter(Mandatory)] [string]$Message,
        [ValidateSet('INFO', 'OK', 'WARN', 'ERROR')] [string]$Level = 'INFO'
    )
    $ts   = (Get-Date).ToString('yyyy-MM-dd HH:mm:ss')
    $line = '[{0}] [{1,-5}] {2}' -f $ts, $Level, $Message
    switch ($Level) {
        'ERROR' { Write-Host $line -ForegroundColor Red }
        'WARN'  { Write-Host $line -ForegroundColor Yellow }
        'OK'    { Write-Host $line -ForegroundColor Green }
        default { Write-Host $line -ForegroundColor Gray }
    }
    if ($script:LogFile) {
        try { Add-Content -LiteralPath $script:LogFile -Value $line -Encoding UTF8 } catch { }
    }
}

function Get-Prop {
    # Safe property read from a PSCustomObject row.
    param($Row, [string]$Name)
    if ([string]::IsNullOrEmpty($Name)) { return $null }
    $p = $Row.PSObject.Properties[$Name]
    if ($p) { return $p.Value }
    return $null
}

function ConvertTo-Text {
    param($Value)
    if ($null -eq $Value) { return '' }
    return ([string]$Value).Trim()
}

function ConvertTo-HtmlSafe {
    param([string]$Text)
    if ([string]::IsNullOrEmpty($Text)) { return '' }
    return $Text.Replace('&', '&amp;').Replace('<', '&lt;').Replace('>', '&gt;').Replace('"', '&quot;').Replace("'", '&#39;')
}

function Get-NormalKey {
    param([string]$Text)
    if ($null -eq $Text) { return '' }
    return ($Text.ToLowerInvariant() -replace '[^a-z0-9]', '')
}

function ConvertTo-DateOrNull {
    param($Value)
    if ($null -eq $Value) { return $null }
    if ($Value -is [datetime]) { return [datetime]$Value }
    $s = ([string]$Value).Trim()
    if ($s -eq '') { return $null }

    # OADate serial number (Excel stores dates as numbers)
    $num = 0.0
    if ([double]::TryParse($s, [ref]$num) -and $num -gt 30000 -and $num -lt 80000) {
        try { return [datetime]::FromOADate($num) } catch { }
    }

    $formats = @(
        'yyyy-MM-dd', 'dd/MM/yyyy', 'MM/dd/yyyy', 'd/M/yyyy', 'M/d/yyyy',
        'dd-MM-yyyy', 'MM-dd-yyyy', 'dd.MM.yyyy', 'yyyy/MM/dd',
        'dd MMM yyyy', 'dd MMMM yyyy', 'MMM dd, yyyy', 'MMMM dd, yyyy',
        'dd-MMM-yyyy', 'dd-MMM-yy', 'yyyy-MM-dd HH:mm:ss'
    )
    foreach ($f in $formats) {
        try {
            return [datetime]::ParseExact($s, $f, [System.Globalization.CultureInfo]::InvariantCulture)
        } catch { }
    }
    $d = [datetime]::MinValue
    if ([datetime]::TryParse($s, [System.Globalization.CultureInfo]::InvariantCulture, [System.Globalization.DateTimeStyles]::None, [ref]$d)) { return $d }
    if ([datetime]::TryParse($s, [System.Globalization.CultureInfo]::CurrentCulture,  [System.Globalization.DateTimeStyles]::None, [ref]$d)) { return $d }
    return $null
}

function Get-StatusBucket {
    param([string]$Raw)
    $n = Get-NormalKey $Raw
    if ($n -eq '') { return 'Unknown' }
    if ($n -match 'not(completed|complete|done|patched|installed|applied|started|successful)') { return 'Pending' }
    foreach ($kw in $StatusKeywords.Failed)    { if ($n.Contains($kw)) { return 'Failed' } }
    foreach ($kw in $StatusKeywords.Completed) { if ($n.Contains($kw)) { return 'Completed' } }
    foreach ($kw in $StatusKeywords.Pending)   { if ($n.Contains($kw)) { return 'Pending' } }
    return 'Unknown'
}

function Resolve-Columns {
    <#
        Maps every canonical field to the real header found in the sheet.
        Returns a hashtable  canonical -> actual header (or $null).
    #>
    param(
        [string[]]$PropertyNames,
        [System.Collections.Specialized.OrderedDictionary]$Aliases
    )
    $lookup = @{}
    foreach ($p in $PropertyNames) {
        $k = Get-NormalKey $p
        if ($k -and -not $lookup.ContainsKey($k)) { $lookup[$k] = $p }
    }
    $result = @{}
    foreach ($field in $Aliases.Keys) {
        $match = $null
        foreach ($alias in $Aliases[$field]) {
            $ak = Get-NormalKey $alias
            if ($ak -eq '') { continue }
            if ($lookup.ContainsKey($ak)) { $match = $lookup[$ak]; break }
        }
        if (-not $match) {
            # start-with / contains fallback for long headers
            foreach ($alias in $Aliases[$field]) {
                $ak = Get-NormalKey $alias
                if ($ak.Length -lt 2) { continue }
                foreach ($k in $lookup.Keys) {
                    if ($k.StartsWith($ak) -or ($ak.Length -ge 4 -and $k.Contains($ak))) { $match = $lookup[$k]; break }
                }
                if ($match) { break }
            }
        }
        $result[$field] = $match
    }
    return $result
}

function New-Slug {
    param([string]$Text, [string]$Fallback = 'Report')
    if ([string]::IsNullOrWhiteSpace($Text)) { $Text = $Fallback }
    $parts = ($Text -replace '[^A-Za-z0-9]+', ' ').Trim() -split '\s+'
    $tc = [System.Globalization.CultureInfo]::InvariantCulture.TextInfo
    $parts = $parts | Where-Object { $_ -ne '' } | ForEach-Object {
        if ($_ -cmatch '^[A-Z0-9]+$') { $_ } else { $tc.ToTitleCase($_.ToLowerInvariant()) }
    }
    if (-not $parts) { return $Fallback }
    return ($parts -join '-')
}

function Resolve-SiteName {
    # Expand a short site code (SF, AQ, AMC, TB, AB, HA ...) to its full name
    # using $SiteNameMap. Falls back to the original text when nothing matches.
    param([string]$Raw)
    $r = ([string]$Raw).Trim()
    if ($r -eq '') { return $r }
    if ($SiteNameMap.ContainsKey($r.ToUpperInvariant())) { return $SiteNameMap[$r.ToUpperInvariant()] }
    foreach ($tok in ($r -split '[^A-Za-z0-9]+')) {
        if ($tok -and $SiteNameMap.ContainsKey($tok.ToUpperInvariant())) { return $SiteNameMap[$tok.ToUpperInvariant()] }
    }
    return $r
}

function Get-MonthLabel {
    <# Most common month/year across a set of dates -> "September 2026" #>
    param([datetime[]]$Dates)
    $valid = @($Dates | Where-Object { $_ -is [datetime] -and $_ -ne [datetime]::MinValue })
    if ($valid.Count -eq 0) { return $null }
    $grp = $valid | Group-Object { '{0:0000}-{1:00}' -f $_.Year, $_.Month } |
        Sort-Object Count, Name -Descending | Select-Object -First 1
    $y, $m = $grp.Name -split '-'
    $name = [System.Globalization.CultureInfo]::CurrentCulture.DateTimeFormat.GetMonthName([int]$m)
    return ('{0} {1}' -f $name, $y)
}

function Get-MonthLabelFromString {
    param([string]$Text)
    if ([string]::IsNullOrWhiteSpace($Text)) { return $null }
    $months = 'January|February|March|April|May|June|July|August|September|October|November|December'
    if ($Text -match "($months)[ _-]?(\d{4})") { return ('{0} {1}' -f $Matches[1], $Matches[2]) }
    if ($Text -match '(\d{4})[ _-](\d{2})') {
        $name = [System.Globalization.CultureInfo]::CurrentCulture.DateTimeFormat.GetMonthName([int]$Matches[2])
        return ('{0} {1}' -f $name, $Matches[1])
    }
    return $null
}

# ======================================================================
#  HTML BUILDER
# ======================================================================
function Build-DashboardHtml {
    param([hashtable]$Ctx)

    $P = $Palette
    $safeSite = ConvertTo-HtmlSafe $Ctx.SiteDisplay
    $safeEng  = ConvertTo-HtmlSafe $Ctx.EngineerDisplay
    $safeMon  = ConvertTo-HtmlSafe $Ctx.MonthDisplay

    $total     = [int]$Ctx.Total
    $completed = [int]$Ctx.Completed
    $failed    = [int]$Ctx.Failed
    $pending   = [int]$Ctx.Pending
    $unknown   = [int]$Ctx.Unknown

    $statusClass = $Ctx.StatusClass
    $statusText  = ConvertTo-HtmlSafe $Ctx.StatusText
    $statusColor = switch ($statusClass) { 'good' { $P.Green } 'warn' { $P.Amber } default { $P.Red } }

    # ---- donut segments -------------------------------------------------
    $r = 80; $cx = 100; $cy = 100
    $circ = [math]::Round(2 * [math]::PI * $r, 3)
    $segDefs = @(
        @{ Label = 'Completed'; Value = $completed; Color = $P.Green },
        @{ Label = 'Pending';   Value = $pending;   Color = $P.Amber },
        @{ Label = 'Failed';    Value = $failed;    Color = $P.Red },
        @{ Label = 'Unknown';   Value = $unknown;   Color = $P.Grey }
    ) | Where-Object { $_.Value -gt 0 }

    $svgSegs = ''
    $accum = 0.0
    if ($total -gt 0 -and $segDefs.Count -gt 0) {
        foreach ($s in $segDefs) {
            $len = [math]::Round($circ * ($s.Value / $total), 3)
            $gap = [math]::Round($circ - $len, 3)
            $off = [math]::Round(-$accum, 3)
            $svgSegs += "      <circle cx='$cx' cy='$cy' r='$r' fill='none' stroke='$($s.Color)' stroke-width='34' stroke-dasharray='$len $gap' stroke-dashoffset='$off' transform='rotate(-90 $cx $cy)'></circle>`n"
            $accum += $len
        }
    } else {
        $svgSegs = "      <circle cx='$cx' cy='$cy' r='$r' fill='none' stroke='$($P.Line)' stroke-width='34'></circle>`n"
    }

    # ---- legend ------------------------------------------------------
    $legend = @"
        <div class="legend">
          <span><i style="background:$($P.Green)"></i>Completed <b>$completed</b></span>
          <span><i style="background:$($P.Amber)"></i>Pending / Not started <b>$pending</b></span>
          <span><i style="background:$($P.Red)"></i>Failed <b>$failed</b></span>
$([string]$(if ($unknown -gt 0) { "          <span><i style=""background:$($P.Grey)""></i>Unknown <b>$unknown</b></span>`n" }))        </div>
"@

    # ---- KPI cards --------------------------------------------------
    $pendNote = if ($unknown -gt 0) { "incl. $unknown unknown" } else { 'Not started / in progress' }
    $kpiCards = @"
      <div class="kpi">
        <div class="kpi-card"><div class="kpi-accent" style="background:$($P.SlateHi)"></div>
          <div class="kpi-label">Total Servers</div><div class="kpi-value">$total</div>
          <div class="kpi-sub">Scheduled this cycle</div></div>
        <div class="kpi-card"><div class="kpi-accent" style="background:$($P.Green)"></div>
          <div class="kpi-label">Completed</div><div class="kpi-value" style="color:$($P.Green)">$completed</div>
          <div class="kpi-sub">Successfully patched</div></div>
        <div class="kpi-card"><div class="kpi-accent" style="background:$($P.Red)"></div>
          <div class="kpi-label">Failed</div><div class="kpi-value" style="color:$($P.Red)">$failed</div>
          <div class="kpi-sub">Patching errors</div></div>
        <div class="kpi-card"><div class="kpi-accent" style="background:$($P.Amber)"></div>
          <div class="kpi-label">Pending</div><div class="kpi-value" style="color:$($P.Amber)">$pending</div>
          <div class="kpi-sub">$pendNote</div></div>
        <div class="kpi-card"><div class="kpi-accent" style="background:$statusColor"></div>
          <div class="kpi-label">Overall Status</div><div class="kpi-status $statusClass">$statusText</div>
          <div class="kpi-sub">Completed vs. failed / pending</div></div>
      </div>
"@

    # ---- OS breakdown --------------------------------------------
    $osRows = ''
    if ($Ctx.OsBreakdown.Count -gt 0) {
        $osMax = ($Ctx.OsBreakdown | Measure-Object -Property Count -Maximum).Maximum
        foreach ($o in $Ctx.OsBreakdown) {
            $w = if ($osMax -gt 0) { [math]::Round(100 * $o.Count / $osMax, 1) } else { 0 }
            $osName = ConvertTo-HtmlSafe $o.Name
            $osRows += @"
          <div class="os-row">
            <div class="os-name" title="$osName">$osName</div>
            <div class="os-bar"><div class="os-fill" style="width:$w%"></div></div>
            <div class="os-count">$($o.Count)</div>
          </div>
"@
        }
    } else {
        $osRows = '<div class="muted">No OS information available in the workbook.</div>'
    }

    # ---- exceptions table -----------------------------------------
    $exBody = ''
    if ($Ctx.Exceptions.Count -gt 0) {
        foreach ($e in $Ctx.Exceptions) {
            $bClass = switch ($e.Bucket) { 'Failed' { 'b-red' } 'Pending' { 'b-amber' } default { 'b-grey' } }
            $exBody += @"
            <tr>
              <td>$([string](ConvertTo-HtmlSafe $e.HostName))</td>
              <td>$([string](ConvertTo-HtmlSafe $e.IP))</td>
              <td>$([string](ConvertTo-HtmlSafe $e.OS))</td>
              <td>$([string](ConvertTo-HtmlSafe $e.TechOwner))</td>
              <td>$([string](ConvertTo-HtmlSafe $e.Engineer))</td>
              <td><span class="badge $bClass">$([string](ConvertTo-HtmlSafe $e.StatusText))</span></td>
            </tr>
"@
        }
        $exSection = @"
        <table class="ex-table">
          <thead><tr><th>HostName</th><th>IP Address</th><th>OS</th><th>Technical Owner</th><th>Engineer</th><th>Status</th></tr></thead>
          <tbody>
$exBody          </tbody>
        </table>
"@
    } else {
        $exSection = '<div class="all-clear">&#10004;&nbsp; No Exceptions &ndash; All Scheduled Servers Successfully Patched</div>'
    }

    # ---- compact server list (VM name + status) under the donut --
    $srvItems = ''
    foreach ($sv in $Ctx.Servers) {
        $dotClass = switch ($sv.Bucket) {
            'Completed' { 's-green' } 'Failed' { 's-red' } 'Pending' { 's-amber' } default { 's-grey' }
        }
        $svName = [string](ConvertTo-HtmlSafe $sv.HostName)
        $svStat = [string](ConvertTo-HtmlSafe $sv.StatusText)
        $srvItems += "            <div class=""srv-item""><span class=""srv-dot $dotClass""></span><span class=""srv-name"" title=""$svName"">$svName</span><span class=""srv-stat"">$svStat</span></div>`n"
    }
    $serverList = @"
        <div class="srv-wrap">
          <div class="srv-title">Servers in this cycle ($total)</div>
          <div class="srv-grid">
$srvItems          </div>
        </div>
"@

    # ---- data-quality notes -------------------------------------
    $dq = ''
    if ($Ctx.DataQuality.Count -gt 0) {
        foreach ($n in $Ctx.DataQuality) { $dq += "            <li>$([string](ConvertTo-HtmlSafe $n))</li>`n" }
        $dqSection = "        <details class=""dq""><summary>Data-quality notes ($($Ctx.DataQuality.Count))</summary>`n          <ul>`n$dq          </ul>`n        </details>"
    } else {
        $dqSection = '        <div class="dq-clean">No data-quality issues detected in the workbook.</div>'
    }

    $genStamp   = ConvertTo-HtmlSafe $Ctx.Generated
    $srcName    = ConvertTo-HtmlSafe $Ctx.SourceFile
    $sheetName  = ConvertTo-HtmlSafe $Ctx.Worksheet

$css = @'
  :root{--ink:#1F2933;--line:#E3E8EF;--muted:#6B7683;--bg:#F4F6F9;}
  *{box-sizing:border-box;}
  body{margin:0;background:var(--bg);color:var(--ink);
       font-family:-apple-system,"Segoe UI",Roboto,Helvetica,Arial,sans-serif;line-height:1.45;}
  .wrap{max-width:1180px;margin:0 auto;padding:28px;}
  header.hero{background:linear-gradient(135deg,#0F2A43 0%,#1D4E79 100%);color:#fff;
       border-radius:14px;padding:40px 44px;box-shadow:0 10px 30px rgba(15,42,67,.18);}
  .site-name{font-size:56px;font-weight:800;letter-spacing:2px;text-transform:uppercase;margin:0;line-height:1.05;}
  .engineer{font-size:26px;font-weight:600;margin:10px 0 0;color:#DCE7F2;}
  .engineer b{color:#fff;}
  .report-title{font-size:18px;font-weight:600;margin:22px 0 2px;letter-spacing:.5px;color:#AFC6DD;text-transform:uppercase;}
  .report-month{font-size:22px;font-weight:700;margin:0;color:#fff;}
  .kpi{display:grid;grid-template-columns:repeat(5,1fr);gap:16px;margin:26px 0 10px;}
  .kpi-card{position:relative;background:#fff;border:1px solid var(--line);border-radius:12px;
       padding:18px 16px 16px;overflow:hidden;box-shadow:0 2px 6px rgba(31,41,51,.04);}
  .kpi-accent{position:absolute;top:0;left:0;right:0;height:5px;}
  .kpi-label{font-size:12px;font-weight:700;text-transform:uppercase;letter-spacing:.6px;color:var(--muted);}
  .kpi-value{font-size:34px;font-weight:800;margin-top:6px;}
  .kpi-status{font-size:17px;font-weight:800;margin-top:10px;}
  .kpi-sub{font-size:11px;color:var(--muted);margin-top:6px;}
  .kpi-status.good{color:#1F8A4C;} .kpi-status.warn{color:#C77700;} .kpi-status.bad{color:#C0392B;}
  section.panel{background:#fff;border:1px solid var(--line);border-radius:12px;padding:22px 24px;margin-top:20px;
       box-shadow:0 2px 6px rgba(31,41,51,.04);}
  h2{font-size:15px;text-transform:uppercase;letter-spacing:.7px;margin:0 0 16px;color:#334155;}
  .charts{display:grid;grid-template-columns:1fr 1fr;gap:20px;}
  .charts.single{grid-template-columns:1fr;}
  .chart-box{display:flex;flex-direction:column;align-items:center;}
  .legend{display:flex;flex-wrap:wrap;gap:14px 20px;margin-top:14px;font-size:13px;justify-content:center;}
  .legend i{display:inline-block;width:12px;height:12px;border-radius:3px;margin-right:7px;vertical-align:middle;}
  .legend b{margin-left:5px;}
  .srv-wrap{width:100%;margin-top:22px;border-top:1px solid var(--line);padding-top:16px;}
  .srv-title{font-size:12px;font-weight:700;text-transform:uppercase;letter-spacing:.5px;color:var(--muted);margin-bottom:10px;text-align:center;}
  .srv-grid{display:grid;grid-template-columns:repeat(auto-fill,minmax(240px,1fr));gap:2px 20px;}
  .srv-item{display:flex;align-items:center;gap:8px;font-size:11.5px;padding:3px 0;border-bottom:1px dotted #EDF1F6;}
  .srv-dot{flex:0 0 auto;width:8px;height:8px;border-radius:50%;}
  .srv-dot.s-green{background:#1F8A4C;} .srv-dot.s-amber{background:#C77700;}
  .srv-dot.s-red{background:#C0392B;} .srv-dot.s-grey{background:#6B7683;}
  .srv-name{flex:1 1 auto;white-space:nowrap;overflow:hidden;text-overflow:ellipsis;color:#334155;font-weight:600;}
  .srv-stat{flex:0 0 auto;color:var(--muted);}
  .os-row{display:grid;grid-template-columns:230px 1fr 46px;align-items:center;gap:12px;margin-bottom:9px;font-size:13px;}
  .os-name{white-space:nowrap;overflow:hidden;text-overflow:ellipsis;color:#334155;}
  .os-bar{background:#EEF2F7;border-radius:6px;height:16px;overflow:hidden;}
  .os-fill{height:100%;background:#1D4E79;border-radius:6px;}
  .os-count{text-align:right;font-weight:700;}
  table.ex-table{width:100%;border-collapse:collapse;font-size:13px;}
  .ex-table th{background:#0F2A43;color:#fff;text-align:left;padding:10px 12px;font-weight:600;}
  .ex-table td{padding:9px 12px;border-bottom:1px solid var(--line);}
  .ex-table tr:nth-child(even){background:#FAFBFD;}
  .badge{display:inline-block;padding:3px 10px;border-radius:20px;font-size:12px;font-weight:700;color:#fff;}
  .badge.b-red{background:#C0392B;} .badge.b-amber{background:#C77700;} .badge.b-grey{background:#6B7683;}
  .all-clear{background:#E9F6EE;border:1px solid #B7E1C6;color:#1F8A4C;font-weight:700;font-size:15px;
       padding:16px 18px;border-radius:10px;text-align:center;}
  .muted{color:var(--muted);font-size:13px;}
  footer{margin-top:24px;font-size:12px;color:var(--muted);}
  .dq{margin-top:10px;} .dq summary{cursor:pointer;font-weight:700;color:#C77700;}
  .dq ul{margin:10px 0 0;padding-left:20px;} .dq li{margin-bottom:4px;}
  .dq-clean{margin-top:10px;color:#1F8A4C;font-size:13px;font-weight:600;}
  @media (max-width:960px){.kpi{grid-template-columns:repeat(2,1fr);}.charts{grid-template-columns:1fr;}
       .site-name{font-size:40px;}.os-row{grid-template-columns:140px 1fr 40px;}}
  @media print{
     body{background:#fff;}
     .wrap{max-width:100%;padding:0;}
     header.hero,.kpi-card,section.panel{box-shadow:none;}
     header.hero{-webkit-print-color-adjust:exact;print-color-adjust:exact;}
     section.panel,.kpi-card,.ex-table tr{page-break-inside:avoid;}
     .srv-grid{grid-template-columns:repeat(3,1fr);}
     .srv-item{page-break-inside:avoid;}
     *{-webkit-print-color-adjust:exact;print-color-adjust:exact;}
  }
'@

$html = @"
<!DOCTYPE html>
<html lang="en">
<head>
<meta charset="utf-8">
<meta name="viewport" content="width=device-width, initial-scale=1">
<title>$safeSite - Patching Executive Report - $safeMon</title>
<style>
$css
</style>
</head>
<body>
<div class="wrap">

  <header class="hero">
    <h1 class="site-name">$safeSite</h1>
    <p class="engineer">Implemented by: <b>$safeEng</b></p>
    <p class="report-title">Monthly Infrastructure Patching Executive Report</p>
    <p class="report-month">$safeMon</p>
  </header>

$kpiCards

  <section class="panel">
    <h2>Patching Status Overview</h2>
    <div class="charts single">
      <div class="chart-box">
        <svg viewBox="0 0 200 200" width="240" height="240" role="img" aria-label="Patching status donut chart">
$svgSegs          <text x="100" y="94" text-anchor="middle" font-size="30" font-weight="800" fill="#1F2933">$total</text>
          <text x="100" y="118" text-anchor="middle" font-size="12" fill="#6B7683">SERVERS</text>
        </svg>
$legend
$serverList      </div>
    </div>
  </section>

  <section class="panel">
    <h2>Operating System Breakdown</h2>
$osRows
  </section>

  <section class="panel">
    <h2>Exceptions / Attention Required</h2>
$exSection
  </section>

  <footer>
$dqSection
    <p style="margin-top:14px;">
      Generated $genStamp from <b>$srcName</b> (worksheet: $sheetName).
      All figures are calculated directly from the source workbook. This report is read-only; the workbook was not modified.
    </p>
  </footer>

</div>
</body>
</html>
"@

    return $html
}

# ======================================================================
#  MAIN
# ======================================================================
try {
    # ---- 0. resolve paths & logging -------------------------------
    if (-not (Test-Path -LiteralPath $InputExcel)) {
        throw "Input Excel file not found: $InputExcel"
    }
    $InputExcel = (Resolve-Path -LiteralPath $InputExcel).Path
    $ext = [System.IO.Path]::GetExtension($InputExcel)
    if ($ext -notmatch '^\.xls[xm]?$') {
        throw "Unsupported file type '$ext'. Provide an .xlsx or .xlsm workbook."
    }

    if (-not $OutputFolder) {
        $OutputFolder = Join-Path (Split-Path -Parent $InputExcel) 'PatchingReports'
    }
    if (-not (Test-Path -LiteralPath $OutputFolder)) {
        New-Item -ItemType Directory -Path $OutputFolder -Force | Out-Null
    }
    $OutputFolder = (Resolve-Path -LiteralPath $OutputFolder).Path
    $script:LogFile = Join-Path $OutputFolder ('Generate-PatchingDashboard_{0:yyyyMMdd_HHmmss}.log' -f (Get-Date))

    Write-Log "==== Patching Dashboard generation started ===="
    Write-Log "Input workbook : $InputExcel"
    Write-Log "Output folder  : $OutputFolder"

    # ---- 1. module check ----------------------------------------
    if (-not (Get-Module -ListAvailable -Name ImportExcel)) {
        throw "The 'ImportExcel' module is required but not installed. Run:  Install-Module ImportExcel -Scope CurrentUser"
    }
    Import-Module ImportExcel -ErrorAction Stop
    Write-Log ("ImportExcel module loaded (v{0})." -f (Get-Module ImportExcel).Version) 'OK'

    # ---- 1b. optional column-map override -----------------------
    if ($ColumnMapPath) {
        if (-not (Test-Path -LiteralPath $ColumnMapPath)) { throw "ColumnMapPath not found: $ColumnMapPath" }
        $override = Get-Content -LiteralPath $ColumnMapPath -Raw | ConvertFrom-Json
        foreach ($k in $override.PSObject.Properties.Name) {
            if ($ColumnAliases.Contains($k)) {
                $ColumnAliases[$k] = @($override.$k) + $ColumnAliases[$k]
                Write-Log "Column aliases for '$k' extended from $ColumnMapPath."
            } else {
                Write-Log "Ignoring unknown field '$k' in column map." 'WARN'
            }
        }
    }

    # ---- 2. read workbook (READ-ONLY) --------------------------
    function Import-Sheet {
        param([string]$Path, [string]$Sheet)
        $p = @{ Path = $Path; DataOnly = $true; ErrorAction = 'Stop' }
        if ($Sheet)      { $p.WorksheetName = $Sheet }
        if ($HeaderRow -gt 1) { $p.StartRow = $HeaderRow }
        Import-Excel @p
    }

    $usedTempCopy = $false
    $readPath = $InputExcel
    try {
        $sheetInfo = Get-ExcelSheetInfo -Path $readPath
    } catch {
        Write-Log "Direct open failed ($($_.Exception.Message)). Retrying on a temporary copy..." 'WARN'
        $readPath = Join-Path $env:TEMP ('patchdash_{0}{1}' -f ([guid]::NewGuid().ToString('N')), $ext)
        Copy-Item -LiteralPath $InputExcel -Destination $readPath -Force
        $usedTempCopy = $true
        $sheetInfo = Get-ExcelSheetInfo -Path $readPath
    }

    $visibleSheets = @($sheetInfo | Where-Object { $_.Hidden -notin @('Hidden', 'VeryHidden') })
    if (-not $visibleSheets) { $visibleSheets = @($sheetInfo) }

    $data = $null; $chosenSheet = $null; $cols = $null
    $candidates = if ($WorksheetName) { @($WorksheetName) } else { $visibleSheets.Name }

    foreach ($sn in $candidates) {
        try { $try = Import-Sheet -Path $readPath -Sheet $sn } catch { Write-Log "Could not read sheet '$sn': $($_.Exception.Message)" 'WARN'; continue }
        if (-not $try) { continue }
        $first = $try | Select-Object -First 1
        $propNames = $first.PSObject.Properties.Name
        $c = Resolve-Columns -PropertyNames $propNames -Aliases $ColumnAliases
        $missing = $RequiredFields | Where-Object { -not $c[$_] }
        if ($missing.Count -eq 0) {
            $data = $try; $chosenSheet = $sn; $cols = $c
            Write-Log "Using worksheet '$sn' ($([array]::IndexOf($visibleSheets.Name,$sn)+1) of $($visibleSheets.Count))." 'OK'
            break
        } else {
            Write-Log "Sheet '$sn' skipped - missing required column(s): $($missing -join ', ')." 'WARN'
        }
    }

    if ($usedTempCopy) { Remove-Item -LiteralPath $readPath -Force -ErrorAction SilentlyContinue }

    if (-not $data) {
        throw ("No worksheet contained all required columns ({0}). " -f ($RequiredFields -join ', ')) +
              "Check the header row (-HeaderRow) or extend the aliases in section 1 of the script / via -ColumnMapPath."
    }

    # ---- 3. report resolved columns ---------------------------
    Write-Log "---- Column mapping ----"
    foreach ($f in $ColumnAliases.Keys) {
        if ($cols[$f]) { Write-Log ("  {0,-15} -> '{1}'" -f $f, $cols[$f]) }
        else           { Write-Log ("  {0,-15} -> (not found)" -f $f) 'WARN' }
    }

    # ---- 4. project rows & validate --------------------------
    $dq = New-Object System.Collections.Generic.List[string]
    if (-not $cols.SiteName)  { $dq.Add("No Site column found - using $([string]$(if($SiteName){"override '$SiteName'"}else{'file name / placeholder'})).") }
    if (-not $cols.Engineer)  { $dq.Add("No Engineer / Implementer column found.") }
    if (-not $cols.StartDate) { $dq.Add("No Start Patching Date column found - month derived from file name / file date.") }
    if (-not $cols.OS)        { $dq.Add("No OS column found - OS breakdown will be empty.") }
    if (-not $cols.TechnicalOwner) { $dq.Add("No Technical Owner column found.") }
    if (-not $cols.IPAddress) { $dq.Add("No IP Address column found.") }

    $records   = New-Object System.Collections.Generic.List[object]
    $blankRows = 0; $noHost = 0; $dateFail = 0; $unknownStatus = 0; $rowNo = $HeaderRow

    foreach ($row in $data) {
        $rowNo++
        $h  = ConvertTo-Text (Get-Prop $row $cols.HostName)
        $st = ConvertTo-Text (Get-Prop $row $cols.Status)
        $ip = ConvertTo-Text (Get-Prop $row $cols.IPAddress)
        $os = ConvertTo-Text (Get-Prop $row $cols.OS)
        $to = ConvertTo-Text (Get-Prop $row $cols.TechnicalOwner)
        $en = ConvertTo-Text (Get-Prop $row $cols.Engineer)
        $sr = Resolve-SiteName (ConvertTo-Text (Get-Prop $row $cols.SiteName))
        $dr = Get-Prop $row $cols.StartDate

        if (($h -eq '') -and ($st -eq '') -and ($ip -eq '') -and ($os -eq '') -and ($to -eq '') -and ($en -eq '') -and ($sr -eq '')) {
            $blankRows++; continue
        }
        if ($h -eq '') { $noHost++; $dq.Add("Row $rowNo has no HostName (status '$st').") ; $h = '(missing hostname)' }

        $bucket = Get-StatusBucket $st
        if ($bucket -eq 'Unknown') {
            $unknownStatus++
            if ($st -ne '') { $dq.Add("Row $rowNo - unrecognised status '$st' (not counted as completed).") }
            else            { $dq.Add("Row $rowNo ($h) - empty Patching Status (not counted as completed).") }
        }

        $date = $null
        if ($cols.StartDate) {
            $date = ConvertTo-DateOrNull $dr
            if (-not $date -and (ConvertTo-Text $dr) -ne '') { $dateFail++; $dq.Add("Row $rowNo ($h) - unparseable date '$(ConvertTo-Text $dr)'.") }
        }

        $records.Add([pscustomobject]@{
            RowNo      = $rowNo
            HostName   = $h
            IP         = $ip
            OS         = $os
            TechOwner  = $to
            Engineer   = $en
            SiteRaw    = $sr
            StatusText = if ($st -ne '') { $st } else { '(blank)' }
            Bucket     = $bucket
            Date       = $date
        })
    }

    if ($blankRows -gt 0) { Write-Log "Skipped $blankRows blank row(s)." }
    if ($records.Count -eq 0) { throw "No data rows found after validation. Nothing to report." }

    # duplicate hostnames
    $dups = $records | Where-Object { $_.HostName -ne '(missing hostname)' } |
        Group-Object { $_.HostName.ToLowerInvariant() } | Where-Object { $_.Count -gt 1 }
    foreach ($d in $dups) { $dq.Add("Duplicate HostName '$($d.Group[0].HostName)' appears $($d.Count) times (rows $((($d.Group.RowNo) -join ', '))).") }

    Write-Log ("Valid server records: {0} (unique hosts: {1})" -f $records.Count, (($records.HostName | Sort-Object -Unique).Count)) 'OK'

    # ---- 5. group by site -----------------------------------
    if ($SiteName) {
        $ovr = Resolve-SiteName $SiteName
        $groups = @([pscustomobject]@{ Key = 'override'; Records = $records; DisplayName = $ovr })
        Write-Log "Site name overridden via -SiteName: '$SiteName' -> '$ovr'." 'WARN'
    }
    else {
        $siteGroups = $records | Group-Object {
            $k = Get-NormalKey $_.SiteRaw
            if ($k) { $k } else { '__none__' }
        }
        $distinctNames = @($records | ForEach-Object { $_.SiteRaw } | Where-Object { $_ -ne '' } | Sort-Object -Unique)

        if ($siteGroups.Count -le 1 -or $CombineSites) {
            if ($distinctNames.Count -gt 1 -and $CombineSites) {
                $disp = 'MULTIPLE SITES'
                $dq.Add("Workbook contains $($distinctNames.Count) sites ($($distinctNames -join ', ')) - combined into one report by -CombineSites.")
            }
            elseif ($distinctNames.Count -eq 1) { $disp = $distinctNames[0] }
            elseif ($distinctNames.Count -gt 1) { $disp = 'MULTIPLE SITES' }
            else {
                $baseName = [System.IO.Path]::GetFileNameWithoutExtension($InputExcel)
                $codeHit  = $null
                foreach ($tok in ($baseName -split '[^A-Za-z0-9]+')) {
                    if ($tok -and $SiteNameMap.ContainsKey($tok.ToUpperInvariant())) { $codeHit = $SiteNameMap[$tok.ToUpperInvariant()]; break }
                }
                if ($codeHit) {
                    $disp = $codeHit
                    $dq.Add("Site Name absent from workbook - resolved site code in file name to '$disp'.")
                }
                else {
                    $fileGuess = New-Slug $baseName 'UNKNOWN SITE'
                    $disp = if ($fileGuess -and $fileGuess -ne 'Report') { ($fileGuess -replace '-', ' ') } else { 'SITE NAME NOT SPECIFIED' }
                    $dq.Add("Site Name absent from workbook - using '$disp' (from file name). Pass -SiteName to override.")
                }
            }
            $groups = @([pscustomobject]@{ Key = 'all'; Records = $records; DisplayName = $disp })
        }
        else {
            Write-Log "Multiple sites detected ($($distinctNames -join ', ')) - generating one dashboard per site." 'WARN'
            $dq.Add("Workbook contains $($siteGroups.Count) sites - one dashboard generated per site.")
            $groups = foreach ($g in $siteGroups) {
                $nm = @($g.Group | ForEach-Object { $_.SiteRaw } | Where-Object { $_ -ne '' } | Select-Object -First 1)
                [pscustomobject]@{
                    Key         = $g.Name
                    Records     = $g.Group
                    DisplayName = if ($nm) { $nm[0] } else { 'SITE NAME NOT SPECIFIED' }
                }
            }
        }
    }

    # ---- 6. build a dashboard per group -------------------
    $generated = New-Object System.Collections.Generic.List[object]

    foreach ($grp in $groups) {
        $recs = @($grp.Records)
        $siteDisplay = $grp.DisplayName

        # engineers
        if ($EngineerName) {
            $engDisplay = $EngineerName
        } else {
            $engs = @($recs | ForEach-Object { $_.Engineer } | Where-Object { $_ -ne '' } |
                      ForEach-Object { $_.Trim() } | Sort-Object -Unique)
            if ($engs.Count -eq 0)      { $engDisplay = 'Not specified in workbook' ; $dq.Add("[$siteDisplay] No engineer / implementer value found.") }
            elseif ($engs.Count -eq 1)  { $engDisplay = $engs[0] }
            else                        { $engDisplay = ($engs -join ', ') ; Write-Log "[$siteDisplay] Multiple engineers: $engDisplay" }
        }

        # month
        if ($PatchingMonth) {
            $monthDisplay = $PatchingMonth
        } else {
            $monthDisplay = Get-MonthLabel ($recs | ForEach-Object { $_.Date })
            if (-not $monthDisplay) { $monthDisplay = Get-MonthLabelFromString ([System.IO.Path]::GetFileNameWithoutExtension($InputExcel)) }
            if (-not $monthDisplay) {
                $monthDisplay = (Get-Item -LiteralPath $InputExcel).LastWriteTime.ToString('MMMM yyyy')
                $dq.Add("[$siteDisplay] Patching month not found in data or file name - using workbook file date ($monthDisplay).")
            }
        }

        # KPIs
        $total     = $recs.Count
        $completed = @($recs | Where-Object { $_.Bucket -eq 'Completed' }).Count
        $failed    = @($recs | Where-Object { $_.Bucket -eq 'Failed'    }).Count
        $pending   = @($recs | Where-Object { $_.Bucket -eq 'Pending'   }).Count
        $unknown   = @($recs | Where-Object { $_.Bucket -eq 'Unknown'   }).Count
        $compliance = if ($total -gt 0) { [math]::Round(100.0 * $completed / $total, 1) } else { 0 }

        if     ($total -eq 0)                        { $sTxt = 'No Data';                    $sCls = 'bad' }
        elseif ($failed -gt 0)                       { $sTxt = 'Attention Required';         $sCls = 'bad' }
        elseif ($unknown -gt 0)                      { $sTxt = 'Attention Required';         $sCls = 'bad' }
        elseif ($compliance -ge 100)                 { $sTxt = 'Fully Compliant';            $sCls = 'good' }
        elseif ($compliance -ge $ComplianceThreshold){ $sTxt = 'On Track - Minor Exceptions'; $sCls = 'warn' }
        else                                         { $sTxt = 'Attention Required';         $sCls = 'bad' }

        # OS breakdown
        $osBreak = $recs | ForEach-Object { if ($_.OS -ne '') { $_.OS } else { '(not specified)' } } |
            Group-Object | Sort-Object Count -Descending |
            ForEach-Object { [pscustomobject]@{ Name = $_.Name; Count = $_.Count } }
        if ($osBreak.Count -gt 8) {
            $top = $osBreak | Select-Object -First 7
            $rest = ($osBreak | Select-Object -Skip 7 | Measure-Object -Property Count -Sum).Sum
            $osBreak = @($top) + [pscustomobject]@{ Name = 'Other'; Count = $rest }
        }

        # exceptions
        $order = @{ 'Failed' = 0; 'Pending' = 1; 'Unknown' = 2 }
        $exceptions = $recs | Where-Object { $_.Bucket -ne 'Completed' } |
            Sort-Object @{ E = { $order[$_.Bucket] } }, HostName

        $ctx = @{
            SiteDisplay     = $siteDisplay
            EngineerDisplay = $engDisplay
            MonthDisplay    = $monthDisplay
            Total = $total; Completed = $completed; Failed = $failed; Pending = $pending; Unknown = $unknown
            Compliance = $compliance; Threshold = $ComplianceThreshold
            StatusText = $sTxt; StatusClass = $sCls
            OsBreakdown = @($osBreak)
            Exceptions  = @($exceptions)
            Servers     = @($recs | Sort-Object HostName)
            DataQuality = @($dq | Select-Object -Unique)
            Generated   = (Get-Date).ToString('yyyy-MM-dd HH:mm')
            SourceFile  = [System.IO.Path]::GetFileName($InputExcel)
            Worksheet   = $chosenSheet
        }

        $html = Build-DashboardHtml -Ctx $ctx

        $monthSlug = New-Slug $monthDisplay ('Report-{0:yyyyMMdd}' -f (Get-Date))
        $siteSlug  = New-Slug $siteDisplay 'Site'
        $fileName  = '{0}-Patching-Executive-Report-{1}.html' -f $siteSlug, $monthSlug
        $outFile   = Join-Path $OutputFolder $fileName

        [System.IO.File]::WriteAllText($outFile, $html, (New-Object System.Text.UTF8Encoding($false)))
        Write-Log "Dashboard written: $outFile" 'OK'
        Write-Log ("  Site='{0}'  Engineer='{1}'  Month='{2}'  Total={3} Completed={4} Failed={5} Pending={6} Unknown={7} Compliance={8}%  Status='{9}'" -f `
                    $siteDisplay, $engDisplay, $monthDisplay, $total, $completed, $failed, $pending, $unknown, $compliance, $sTxt)

        $pdfFile = $null
        if ($ExportPdf) {
            $pdfFile = [System.IO.Path]::ChangeExtension($outFile, '.pdf')
            $browser = $null
            foreach ($cand in @(
                    "$env:ProgramFiles\Microsoft\Edge\Application\msedge.exe",
                    "${env:ProgramFiles(x86)}\Microsoft\Edge\Application\msedge.exe",
                    "$env:ProgramFiles\
