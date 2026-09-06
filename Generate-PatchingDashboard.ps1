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
    Remarks        = @('Issues in the source excel sheet', 'Issue', 'Issues', 'Remarks', 'Remark', 'Comments', 'Comment', 'Notes', 'Note', 'Reason', 'Root Cause', 'RootCause', 'Failure Reason', 'FailureReason', 'Error', 'Error Details', 'Details', 'Description', 'Observation', 'Findings', 'Justification')
    Excluded       = @('Excluded', 'Exclude', 'Exclusion', 'Excluded by Management', 'Management Decision', 'Descoped', 'Out of Scope', 'Exempted', 'Exception')
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

# Text shown directly under the report title.  -EngineerName overrides it.
$ImplementedBy = 'Asset IT Operations'

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
    Excluded  = @('excluded', 'exclude', 'exclusion', 'exempt', 'exempted', 'exemption', 'descoped', 'descope', 'outofscope', 'notinscope', 'removedfromscope', 'waived', 'waiver', 'perManagement', 'managementexclusion')
    Completed = @('completed', 'complete', 'done', 'success', 'successful', 'successfullypatched', 'patchedsuccessfully', 'patched', 'patchingcompleted', 'installed', 'applied', 'closed', 'resolved', 'compliant', 'uptodate', 'fullypatched', 'ok', 'pass', 'passed', 'green')
    Pending   = @('pending', 'notstarted', 'yettostart', 'tobestarted', 'tobedone', 'scheduled', 'planned', 'inprogress', 'wip', 'workinprogress', 'ongoing', 'incomplete', 'partiallycompleted', 'partial', 'onhold', 'hold', 'deferred', 'postponed', 'rescheduled', 'skipped', 'queued', 'inqueue', 'open', 'new', 'rebootpending', 'pendingreboot', 'amber', 'yellow')
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
    foreach ($kw in $StatusKeywords.Excluded)  { if ($n.Contains($kw)) { return 'Excluded' } }
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

function Get-OsFamily {
    # Classify an OS string into Windows / Linux / Other.  Adjust keywords freely.
    param([string]$OsText)
    $n = Get-NormalKey $OsText
    if ($n -eq '') { return 'Other' }
    if ($n -match 'windows|microsoft|winsrv|windowsserver|hyperv|w2k|wserver') { return 'Windows' }
    if ($n -match 'linux|ubuntu|debian|rhel|redhat|centos|rocky|almalinux|alma|oraclelinux|amazonlinux|amzn|suse|sles|opensuse|fedora|photon|coreos') { return 'Linux' }
    return 'Other'
}

function Get-DonutSvg {
    param([int]$Completed, [int]$Pending, [int]$Failed, [int]$Unknown, [int]$Excluded, [int]$Total)
    $P = $Palette
    $cx = 100; $cy = 100; $rr = 80
    $circ = [math]::Round(2 * [math]::PI * $rr, 3)
    $defs = @(
        @{ v = $Completed; c = $P.Green },
        @{ v = $Pending;   c = $P.Amber },
        @{ v = $Failed;    c = $P.Red },
        @{ v = $Excluded;  c = $P.Grey },
        @{ v = $Unknown;   c = '#B0B7C0' }
    ) | Where-Object { $_.v -gt 0 }
    $segs = ''
    $acc = 0.0
    if ($Total -gt 0 -and $defs.Count -gt 0) {
        foreach ($s in $defs) {
            $len = [math]::Round($circ * ($s.v / $Total), 3)
            $gap = [math]::Round($circ - $len, 3)
            $off = [math]::Round(-$acc, 3)
            $segs += "      <circle cx='$cx' cy='$cy' r='$rr' fill='none' stroke='$($s.c)' stroke-width='34' stroke-dasharray='$len $gap' stroke-dashoffset='$off' transform='rotate(-90 $cx $cy)'></circle>`n"
            $acc += $len
        }
    } else {
        $segs = "      <circle cx='$cx' cy='$cy' r='$rr' fill='none' stroke='$($P.Line)' stroke-width='34'></circle>`n"
    }
    return "<svg viewBox='0 0 200 200' width='300' height='300' role='img' aria-label='Patching status donut'>`n$segs      <text x='100' y='92' text-anchor='middle' font-size='32' font-weight='800' fill='#1F2933'>$Total</text>`n      <text x='100' y='116' text-anchor='middle' font-size='11' fill='#6B7683'>VMs</text>`n    </svg>"
}

# ======================================================================
#  HTML BUILDER
# ======================================================================
function Build-DashboardHtml {
    param([hashtable]$Ctx)

    $P = $Palette
    $safeTitle = ConvertTo-HtmlSafe $Ctx.Title
    $safeEng   = ConvertTo-HtmlSafe $Ctx.EngineerDisplay
    $safeMon   = ConvertTo-HtmlSafe $Ctx.MonthDisplay

    # ================================================================
    #  One self-contained block per OS family (Windows, then Linux ...)
    # ================================================================
    $blocksHtml = ''
    foreach ($b in $Ctx.Blocks) {
        $total = [int]$b.Total; $completed = [int]$b.Completed; $failed = [int]$b.Failed
        $pending = [int]$b.Pending; $unknown = [int]$b.Unknown; $excluded = [int]$b.Excluded
        $sCls = $b.StatusClass
        $sTxt = ConvertTo-HtmlSafe $b.StatusText
        $sCol = switch ($sCls) { 'good' { $P.Green } 'warn' { $P.Amber } default { $P.Red } }
        $famU = ([string]$b.Family).ToUpperInvariant()
        $pendLbl = if ($unknown -gt 0) { "incl. $unknown unknown" } else { 'Not started / in progress' }
        $donut = Get-DonutSvg -Completed $completed -Pending $pending -Failed $failed -Unknown $unknown -Excluded $excluded -Total $total
        $exclLegend = if ($excluded -gt 0) { "`n            <span><i style=""background:$($P.Grey)""></i>Excluded <b>$excluded</b></span>" } else { '' }

        # -- completed VMs : scrollable table  VM Name | Patching Date --
        $doneRows = ''
        foreach ($v in ($b.CompletedVMs | Sort-Object HostName)) {
            $d = if ($v.Date -is [datetime]) { $v.Date.ToString('yyyy-MM-dd') } else { 'NA' }
            $doneRows += "            <tr><td>$([string](ConvertTo-HtmlSafe $v.HostName))</td><td>$d</td></tr>`n"
        }
        if ($doneRows) {
            $doneSection = @"
        <p class="mini-cap">Completed VMs ($completed)</p>
        <div class="scroll-wrap">
          <table class="lst"><thead><tr><th>VM Name</th><th>Patching Date</th></tr></thead>
          <tbody>
$doneRows          </tbody></table>
        </div>
"@
        } else {
            $doneSection = '        <div class="muted">No completed VMs.</div>'
        }

        # -- pending / not-started VMs : scrollable table  VM Name  ("NA" when none) --
        $pendRows = ''
        foreach ($v in ($b.PendingVMs | Sort-Object HostName)) {
            $pendRows += "          <tr><td>$([string](ConvertTo-HtmlSafe $v.HostName))</td></tr>`n"
        }
        if ($pendRows) {
            $pendCount = $b.PendingVMs.Count
            $pendSection = @"
      <p class="mini-cap">$pendCount VM(s)</p>
      <div class="scroll-wrap">
        <table class="lst"><thead><tr><th>VM Name</th></tr></thead>
        <tbody>
$pendRows        </tbody></table>
      </div>
"@
        } else {
            $pendSection = '      <div class="na-box">NA</div>'
        }

        # -- OS breakdown bars --
        $osRows = ''
        if ($b.OsBreakdown.Count -gt 0) {
            $osMax = ($b.OsBreakdown | Measure-Object -Property Count -Maximum).Maximum
            foreach ($o in $b.OsBreakdown) {
                $w = if ($osMax -gt 0) { [math]::Round(100 * $o.Count / $osMax, 1) } else { 0 }
                $on = ConvertTo-HtmlSafe $o.Name
                $osRows += "      <div class=""os-row""><div class=""os-name"" title=""$on"">$on</div><div class=""os-bar""><div class=""os-fill"" style=""width:$w%""></div></div><div class=""os-count"">$($o.Count)</div></div>`n"
            }
        } else {
            $osRows = '      <div class="muted">No OS information available.</div>'
        }

        # -- exceptions : FAILED VMs only (pending/not-started are shown above) --
        $exRows = ''
        foreach ($e in $b.Exceptions) {
            $issue = [string](ConvertTo-HtmlSafe $e.Remarks)
            if (-not $issue) {
                $rawTrim = ([string]$e.StatusText).Trim()
                $rawKey  = Get-NormalKey $rawTrim
                if ($rawTrim -and $rawKey -ne 'failed' -and $rawKey -ne 'fail' -and $rawKey -ne 'failure') {
                    $issue = [string](ConvertTo-HtmlSafe $rawTrim)
                } else {
                    $issue = '<span class="muted">Not specified in sheet</span>'
                }
            }
            $exRows += "            <tr><td>$([string](ConvertTo-HtmlSafe $e.HostName))</td><td>$([string](ConvertTo-HtmlSafe $e.IP))</td><td>$([string](ConvertTo-HtmlSafe $e.OS))</td><td>$([string](ConvertTo-HtmlSafe $e.TechOwner))</td><td><span class=""badge b-red"">Failed</span></td><td>$issue</td></tr>`n"
        }
        if ($exRows) {
            $exSection = @"
      <table class="ex-table"><thead><tr><th>HostName</th><th>IP Address</th><th>OS</th><th>Technical Owner</th><th>Status</th><th>Issue</th></tr></thead>
        <tbody>
$exRows        </tbody></table>
"@
        } else {
            $exSection = '      <div class="all-clear">&#10004;&nbsp; No Failed VMs &ndash; nothing requires attention in this section</div>'
        }

        $blocksHtml += @"
  <section class="os-block">
    <div class="os-band">OS &ndash; $famU</div>

    <div class="kpi">
      <div class="kpi-card"><div class="kpi-accent" style="background:$($P.SlateHi)"></div>
        <div class="kpi-label">Total VMs (All)</div><div class="kpi-value">$total</div>
        <div class="kpi-sub">Scheduled this cycle</div></div>
      <div class="kpi-card"><div class="kpi-accent" style="background:$($P.Green)"></div>
        <div class="kpi-label">Completed</div><div class="kpi-value" style="color:$($P.Green)">$completed</div>
        <div class="kpi-sub">Successfully patched</div></div>
      <div class="kpi-card"><div class="kpi-accent" style="background:$($P.Red)"></div>
        <div class="kpi-label">Failed</div><div class="kpi-value" style="color:$($P.Red)">$failed</div>
        <div class="kpi-sub">Patching errors</div></div>
      <div class="kpi-card"><div class="kpi-accent" style="background:$($P.Amber)"></div>
        <div class="kpi-label">Pending</div><div class="kpi-value" style="color:$($P.Amber)">$pending</div>
        <div class="kpi-sub">$pendLbl</div></div>
      <div class="kpi-card"><div class="kpi-accent" style="background:$($P.Grey)"></div>
        <div class="kpi-label">Excluded</div><div class="kpi-value" style="color:$($P.Grey)">$excluded</div>
        <div class="kpi-sub">Excluded as per management</div></div>
      <div class="kpi-card"><div class="kpi-accent" style="background:$sCol"></div>
        <div class="kpi-label">Overall Status</div><div class="kpi-status $sCls">$sTxt</div>
        <div class="kpi-sub">On in-scope VMs (Total &minus; Excluded)</div></div>
    </div>

    <div class="panel">
      <h2>Patching Status Overview</h2>
      <div class="so-grid">
        <div class="donut-box">
    $donut
          <div class="legend">
            <span><i style="background:$($P.Green)"></i>Completed <b>$completed</b></span>
            <span><i style="background:$($P.Amber)"></i>Pending <b>$pending</b></span>
            <span><i style="background:$($P.Red)"></i>Failed <b>$failed</b></span>$exclLegend
          </div>
        </div>
        <div class="done-box">
$doneSection        </div>
      </div>
    </div>

    <div class="panel">
      <h2>Operating System Breakdown</h2>
$osRows    </div>

    <div class="panel">
      <h2>Pending / Not Started VMs</h2>
$pendSection
    </div>

    <div class="panel">
      <h2>Exceptions / Attention Required</h2>
$exSection
    </div>
  </section>

"@
    }

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
       border-radius:14px;padding:34px 40px;box-shadow:0 10px 30px rgba(15,42,67,.18);}
  .report-h1{font-size:34px;font-weight:800;letter-spacing:.4px;margin:0;line-height:1.15;}
  .engineer{font-size:20px;font-weight:600;margin:12px 0 0;color:#DCE7F2;}
  .engineer b{color:#fff;}
  .report-month{font-size:20px;font-weight:700;margin:6px 0 0;color:#fff;}
  .os-band{font-size:30px;font-weight:800;letter-spacing:2px;text-transform:uppercase;color:#0F2A43;
       margin:40px 0 6px;padding-bottom:8px;border-bottom:3px solid #1D4E79;}
  .os-block:first-of-type .os-band{margin-top:14px;}
  .kpi{display:grid;grid-template-columns:repeat(6,1fr);gap:14px;margin:14px 0 10px;}
  .kpi-card{position:relative;background:#fff;border:1px solid var(--line);border-radius:12px;
       padding:18px 16px 16px;overflow:hidden;box-shadow:0 2px 6px rgba(31,41,51,.04);}
  .kpi-accent{position:absolute;top:0;left:0;right:0;height:5px;}
  .kpi-label{font-size:12px;font-weight:700;text-transform:uppercase;letter-spacing:.6px;color:var(--muted);}
  .kpi-value{font-size:34px;font-weight:800;margin-top:6px;}
  .kpi-status{font-size:17px;font-weight:800;margin-top:10px;}
  .kpi-sub{font-size:11px;color:var(--muted);margin-top:6px;}
  .kpi-status.good{color:#1F8A4C;} .kpi-status.warn{color:#C77700;} .kpi-status.bad{color:#C0392B;}
  .panel{background:#fff;border:1px solid var(--line);border-radius:12px;padding:20px 22px;margin-top:18px;
       box-shadow:0 2px 6px rgba(31,41,51,.04);}
  h2{font-size:15px;text-transform:uppercase;letter-spacing:.7px;margin:0 0 16px;color:#334155;}
  .so-grid{display:grid;grid-template-columns:330px 1fr;gap:28px;align-items:start;}
  .donut-box{display:flex;flex-direction:column;align-items:center;}
  .done-box{display:flex;flex-direction:column;min-width:0;}
  .legend{display:flex;flex-wrap:wrap;gap:10px 18px;margin-top:14px;font-size:13px;justify-content:center;}
  .legend i{display:inline-block;width:12px;height:12px;border-radius:3px;margin-right:6px;vertical-align:middle;}
  .legend b{margin-left:4px;}
  .mini-cap{font-size:11.5px;font-weight:700;text-transform:uppercase;letter-spacing:.4px;color:var(--muted);margin:0 0 10px;}
  .scroll-wrap{max-height:300px;overflow-y:auto;border:1px solid var(--line);border-radius:8px;}
  table.lst{width:100%;border-collapse:collapse;font-size:12px;}
  .lst th{position:sticky;top:0;background:#EEF2F7;color:#334155;text-align:left;padding:7px 10px;font-weight:700;}
  .lst td{padding:5px 10px;border-bottom:1px solid var(--line);white-space:nowrap;}
  .lst tr:nth-child(even){background:#FAFBFD;}
  .na-box{background:#F4F6F9;border:1px solid var(--line);color:var(--muted);font-weight:800;font-size:15px;
       letter-spacing:1px;padding:16px;border-radius:10px;text-align:center;}
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
  .all-clear{background:#E9F6EE;border:1px solid #B7E1C6;color:#1F8A4C;font-weight:700;font-size:14px;
       padding:14px 16px;border-radius:10px;text-align:center;}
  .muted{color:var(--muted);font-size:13px;}
  footer{margin-top:24px;font-size:12px;color:var(--muted);}
  .dq{margin-top:10px;} .dq summary{cursor:pointer;font-weight:700;color:#C77700;}
  .dq ul{margin:10px 0 0;padding-left:20px;} .dq li{margin-bottom:4px;}
  .dq-clean{margin-top:10px;color:#1F8A4C;font-size:13px;font-weight:600;}
  @media (max-width:960px){.kpi{grid-template-columns:repeat(3,1fr);}.so-grid{grid-template-columns:1fr;}
       .report-h1{font-size:26px;}.os-row{grid-template-columns:140px 1fr 40px;}}
  @media print{
     body{background:#fff;}
     .wrap{max-width:100%;padding:0;}
     header.hero,.kpi-card,.panel{box-shadow:none;}
     header.hero{-webkit-print-color-adjust:exact;print-color-adjust:exact;}
     .panel,.kpi-card,.ex-table tr,.lst tr{page-break-inside:avoid;}
     .os-block + .os-block{page-break-before:always;}
     .scroll-wrap{max-height:none;overflow:visible;}
     *{-webkit-print-color-adjust:exact;print-color-adjust:exact;}
  }
'@

$html = @"
<!DOCTYPE html>
<html lang="en">
<head>
<meta charset="utf-8">
<meta name="viewport" content="width=device-width, initial-scale=1">
<title>$safeTitle - $safeMon</title>
<style>
$css
</style>
</head>
<body>
<div class="wrap">

  <header class="hero">
    <h1 class="report-h1">$safeTitle</h1>
    <p class="engineer"><b>$safeEng</b></p>
    <p class="report-month">$safeMon</p>
  </header>

$blocksHtml
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
    if (-not $cols.Remarks) { $dq.Add("No Issue / Remarks column found - the 'Issues in the source Excel sheet' column in Exceptions falls back to the raw status text.") }

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
        $rm = ConvertTo-Text (Get-Prop $row $cols.Remarks)
        $xc = ConvertTo-Text (Get-Prop $row $cols.Excluded)
        $sr = Resolve-SiteName (ConvertTo-Text (Get-Prop $row $cols.SiteName))
        $dr = Get-Prop $row $cols.StartDate

        if (($h -eq '') -and ($st -eq '') -and ($ip -eq '') -and ($os -eq '') -and ($to -eq '') -and ($en -eq '') -and ($sr -eq '')) {
            $blankRows++; continue
        }
        if ($h -eq '') { $noHost++; $dq.Add("Row $rowNo has no HostName (status '$st').") ; $h = '(missing hostname)' }

        $bucket = Get-StatusBucket $st
        # an explicit "Excluded" column (truthy value) overrides the status bucket
        $xk = Get-NormalKey $xc
        if ($xk -in @('y', 'yes', 'true', '1', 'x', 'excluded', 'exclude', 'exempt', 'exempted', 'outofscope', 'descoped', 'notinscope', 'permanagement')) { $bucket = 'Excluded' }
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
            Remarks    = $rm
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

        # ---- overall KPIs (all OS) for the run summary + log ----
        $total     = $recs.Count
        $completed = @($recs | Where-Object { $_.Bucket -eq 'Completed' }).Count
        $failed    = @($recs | Where-Object { $_.Bucket -eq 'Failed'    }).Count
        $pending   = @($recs | Where-Object { $_.Bucket -eq 'Pending'   }).Count
        $unknown   = @($recs | Where-Object { $_.Bucket -eq 'Unknown'   }).Count
        $excluded  = @($recs | Where-Object { $_.Bucket -eq 'Excluded'  }).Count
        $inScope   = $total - $excluded
        $compliance = if ($inScope -gt 0) { [math]::Round(100.0 * $completed / $inScope, 1) } else { 0 }

        if     ($inScope -eq 0)                        { $sTxt = 'No In-Scope VMs';             $sCls = 'warn' }
        elseif ($failed -gt 0 -or $unknown -gt 0)      { $sTxt = 'Attention Required';          $sCls = 'bad' }
        elseif ($compliance -ge 100)                   { $sTxt = 'Fully Compliant';             $sCls = 'good' }
        elseif ($compliance -ge $ComplianceThreshold)  { $sTxt = 'On Track - Minor Exceptions'; $sCls = 'warn' }
        else                                           { $sTxt = 'Attention Required';          $sCls = 'bad' }

        # ---- one block per OS family : Windows first, then Linux, then Other ----
        $ord = @{ 'Failed' = 0; 'Pending' = 1; 'Unknown' = 2 }
        $blocks = New-Object System.Collections.Generic.List[object]
        foreach ($fam in @('Windows', 'Linux', 'Other')) {
            $fr = @($recs | Where-Object { (Get-OsFamily $_.OS) -eq $fam })
            if ($fr.Count -eq 0) { continue }

            $fTotal     = $fr.Count
            $fCompleted = @($fr | Where-Object { $_.Bucket -eq 'Completed' }).Count
            $fFailed    = @($fr | Where-Object { $_.Bucket -eq 'Failed'    }).Count
            $fPending   = @($fr | Where-Object { $_.Bucket -eq 'Pending'   }).Count
            $fUnknown   = @($fr | Where-Object { $_.Bucket -eq 'Unknown'   }).Count
            $fExcluded  = @($fr | Where-Object { $_.Bucket -eq 'Excluded'  }).Count
            $fInScope   = $fTotal - $fExcluded
            $fComp      = if ($fInScope -gt 0) { [math]::Round(100.0 * $fCompleted / $fInScope, 1) } else { 0 }

            if     ($fInScope -eq 0)                       { $fTxt = 'No In-Scope VMs';             $fCls = 'warn' }
            elseif ($fFailed -gt 0 -or $fUnknown -gt 0)    { $fTxt = 'Attention Required';          $fCls = 'bad' }
            elseif ($fComp -ge 100)                        { $fTxt = 'Fully Compliant';             $fCls = 'good' }
            elseif ($fComp -ge $ComplianceThreshold)       { $fTxt = 'On Track - Minor Exceptions'; $fCls = 'warn' }
            else                                           { $fTxt = 'Attention Required';          $fCls = 'bad' }

            $fOs = $fr | ForEach-Object { if ($_.OS -ne '') { $_.OS } else { '(not specified)' } } |
                Group-Object | Sort-Object Count -Descending |
                ForEach-Object { [pscustomobject]@{ Name = $_.Name; Count = $_.Count } }
            if ($fOs.Count -gt 8) {
                $fTop  = $fOs | Select-Object -First 7
                $fRest = ($fOs | Select-Object -Skip 7 | Measure-Object -Property Count -Sum).Sum
                $fOs   = @($fTop) + [pscustomobject]@{ Name = 'Other'; Count = $fRest }
            }

            $blocks.Add([pscustomobject]@{
                Family       = $fam
                Total = $fTotal; Completed = $fCompleted; Failed = $fFailed; Pending = $fPending; Unknown = $fUnknown; Excluded = $fExcluded
                Compliance   = $fComp; StatusText = $fTxt; StatusClass = $fCls
                OsBreakdown  = @($fOs)
                CompletedVMs = @($fr | Where-Object { $_.Bucket -eq 'Completed' } | Sort-Object HostName)
                PendingVMs   = @($fr | Where-Object {
                        $k = Get-NormalKey $_.StatusText
                        ($_.Bucket -eq 'Unknown') -or ($k -like '*pending*') -or ($k -like '*notstarted*') -or ($k -like '*notyetstarted*') -or ($k -like '*yettostart*') -or ($k -like '*tobestarted*')
                    } | Sort-Object HostName)
                Exceptions   = @($fr | Where-Object { $_.Bucket -eq 'Failed' } | Sort-Object HostName)
            })
            Write-Log ("  [{0,-7}] Total={1} Completed={2} Failed={3} Pending={4} Unknown={5} Excluded={6}  in-scope compliance {7}%" -f $fam, $fTotal, $fCompleted, $fFailed, $fPending, $fUnknown, $fExcluded, $fComp)
        }
        if ($blocks.Count -eq 0) { $dq.Add('No VMs could be classified by operating system.'); continue }

        $ctx = @{
            Title           = 'Monthly Infrastructure Patching Executive Report'
            EngineerDisplay = $(if ($EngineerName) { $EngineerName } else { $ImplementedBy })
            MonthDisplay    = $monthDisplay
            Threshold       = $ComplianceThreshold
            Blocks          = $blocks.ToArray()
            DataQuality     = @($dq | Select-Object -Unique)
            Generated       = (Get-Date).ToString('yyyy-MM-dd HH:mm')
            SourceFile      = [System.IO.Path]::GetFileName($InputExcel)
            Worksheet       = $chosenSheet
        }

        $html = Build-DashboardHtml -Ctx $ctx

        $monthSlug = New-Slug $monthDisplay ('Report-{0:yyyyMMdd}' -f (Get-Date))
        if ($siteDisplay -match 'NOT SPECIFIED|UNKNOWN SITE|MULTIPLE SITES') {
            $fileName = 'Monthly-Infrastructure-Patching-Executive-Report-{0}.html' -f $monthSlug
        } else {
            $fileName = '{0}-Monthly-Infrastructure-Patching-Executive-Report-{1}.html' -f (New-Slug $siteDisplay 'Site'), $monthSlug
        }
        $outFile   = Join-Path $OutputFolder $fileName

        [System.IO.File]::WriteAllText($outFile, $html, (New-Object System.Text.UTF8Encoding($false)))
        Write-Log "Dashboard written: $outFile" 'OK'
        Write-Log ("  Site='{0}'  Engineer='{1}'  Month='{2}'  Total={3} Completed={4} Failed={5} Pending={6} Unknown={7} Compliance={8}%  Status='{9}'" -f `
                    $siteDisplay, $engDisplay, $monthDisplay, $total, $completed, $failed, $pending, $unknown, $compliance, $sTxt)

        $pdfFile = $null
        if ($ExportPdf) {
            $pdfFile = [System.IO.Path]::ChangeExtension($outFile, '.pdf')
            $browser = $null
            $pfRoots = @($env:ProgramFiles, [Environment]::GetEnvironmentVariable('ProgramFiles(x86)'), $env:LOCALAPPDATA) |
                Where-Object { $_ }
            foreach ($root in $pfRoots) {
                foreach ($rel in @('Microsoft\Edge\Application\msedge.exe', 'Google\Chrome\Application\chrome.exe')) {
                    $cand = Join-Path $root $rel
                    if (Test-Path -LiteralPath $cand) { $browser = $cand; break }
                }
                if ($browser) { break }
            }
            if ($browser) {
                try {
                    $uri = ([System.Uri](Resolve-Path -LiteralPath $outFile).Path).AbsoluteUri
                    & $browser --headless=new --disable-gpu --no-pdf-header-footer "--print-to-pdf=$pdfFile" $uri 2>$null
                    Start-Sleep -Seconds 2
                    if (Test-Path -LiteralPath $pdfFile) { Write-Log "PDF written: $pdfFile" 'OK' }
                    else { Write-Log "PDF was not produced by $browser." 'WARN'; $pdfFile = $null }
                } catch { Write-Log "PDF export failed: $($_.Exception.Message)" 'WARN'; $pdfFile = $null }
            } else {
                Write-Log "No headless Edge/Chrome found - skipping PDF. Open the HTML and 'Print > Save as PDF'." 'WARN'
                $pdfFile = $null
            }
        }

        $generated.Add([pscustomobject]@{ Site = $siteDisplay; Html = $outFile; Pdf = $pdfFile; Compliance = $compliance; Status = $sTxt })
        if ($Open) { Invoke-Item -LiteralPath $outFile }
    }

    Write-Log "==== Completed. $($generated.Count) dashboard(s) generated. ===="
    Write-Host ''
    $generated | Format-Table Site, Compliance, Status, Html -AutoSize | Out-Host
    $generated  # return objects to the pipeline
}
catch {
    $script:ExitCode = 1
    Write-Log $_.Exception.Message 'ERROR'
    Write-Log ($_.ScriptStackTrace) 'ERROR'
    Write-Host ''
    Write-Host "FAILED: $($_.Exception.Message)" -ForegroundColor Red
}
finally {
    if ($script:LogFile) { Write-Host "Log file: $script:LogFile" -ForegroundColor DarkGray }
}

exit $script:ExitCode
