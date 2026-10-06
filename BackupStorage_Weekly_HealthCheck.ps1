#Requires -Version 5.1
<#
.SYNOPSIS
    Backup & Storage Weekly Health Check - reads the Backup Summary workbook and produces a
    single combined HTML dashboard, in the same visual style as Network_Weekly_HealthCheck.ps1.

.DESCRIPTION
    This dashboard covers Backup job health and Storage capacity.

    Backup is read directly from Cohesity's own native "Backup Summary" CSV export(s)
    (-BackupCsvPaths, one or more files) - no manual data entry, no intermediate template, no
    screenshot. Earlier data sources for this dashboard were a pasted-screenshot workbook (no
    readable cells at all) and then a hand-typed template; both are obsolete now that a real
    Cohesity export is available. Real column names: Protection Group, Registered Source, Number
    of Runs, Last Run Successful Objects, Last Run Error Objects, Bytes Read - GiB, Sla, Last Run
    Status, Last Run Time, Replication Copy Tasks Status, Archival Copy Tasks Status, Organization
    Names, Bandwidth (/sec). Each export has a couple of report-header lines before the real
    header row, located by searching for it rather than assuming a fixed line number.

    A single export file does not necessarily cover exactly one site - the real "SFAQ" export
    seen so far covers BOTH AquaArabia and SixFlags in one file, distinguished only by Protection
    Group name (AquaArabia's groups are "AQ-..."; everything else in that file is SixFlags,
    including a couple of utility groups like "SQL" with no site-identifying name or source at
    all). Get-SiteForBackupRow resolves each row's site with, in order: (1) an "AQ-" Protection
    Group prefix - checked first because AquaArabia's rows share SixFlags' own source hostname;
    (2) the Registered Source hostname (sevenab/seventb/sixflags); (3) an AB-/TB-/SF- Protection
    Group prefix as a fallback. A row that still can't be resolved (like "SQL") falls back to
    whichever site the previous resolved row in the SAME file belonged to - every genuinely
    unresolvable row seen in the real exports sits immediately after a resolved row from the
    correct site, so this works in practice; a row that can't be resolved at all (no prior row in
    that file either) is logged and skipped rather than guessed at.

    Storage is read from an Excel workbook (-StorageWorkbookPath) - that one DOES contain real
    cells (confirmed from a screenshot of it showing a normal selectable Excel grid with a
    formula-bar value, not a pasted picture) - ONE "Storage Capacity by Site" table covering
    every site in a single sheet (Site / Total Capacity / Used Capacity / Free Capacity / Model),
    read by locating that header row on whichever worksheet has it, then every row until the
    "Grand Total" row. The real file seen so far only has rows for SEVEN ABHA ("Abha"), SixFlags
    ("Six Flags"), SEVEN Tabuk ("Tabuk") and SEVEN Alhamra ("Alhamra") - AquaArabia and
    SceneCinema aren't in it yet, so their Storage panel shows "There is no data provided yet"
    until a row for them appears; update $StorageSiteNameMap if a site's real label in that sheet
    differs from what's guessed below. A site's Storage status is Critical at >=90% used, Warning
    at >=75%, else Healthy - these thresholds are a starting assumption, easy to change in
    Read-StorageWorkbook if a different bar is wanted.

    Storage is still read from an Excel workbook (-StorageWorkbookPath) - that one DOES contain
    real cells (confirmed from a screenshot of it showing a normal selectable Excel grid with a
    formula-bar value, not a pasted picture) - ONE "Storage Capacity by Site" table covering
    every site in a single sheet (Site / Total Capacity / Used Capacity / Free Capacity / Model),
    read by locating that header row on whichever worksheet has it, then every row until the
    "Grand Total" row. The real file seen so far only has rows for SEVEN ABHA ("Abha"), SixFlags
    ("Six Flags"), SEVEN Tabuk ("Tabuk") and SEVEN Alhamra ("Alhamra") - AquaArabia and
    SceneCinema aren't in it yet, so their Storage panel shows "There is no data provided yet"
    until a row for them appears; update $StorageSiteNameMap if a site's real label in that sheet
    differs from what's guessed below. A site's Storage status is Critical at >=90% used, Warning
    at >=75%, else Healthy - these thresholds are a starting assumption, easy to change in
    Read-StorageWorkbook if a different bar is wanted.

    A protection group with 0 runs this period (shown as "-" across every other column in the
    source) is its own neutral status - it doesn't count as Healthy, Warning or Critical, the
    same way a template site with no data supplied doesn't skew the site-level counts. Every
    other row's status comes from its own Last Run Status and SLA Violation columns: Error ->
    Critical; Warning status OR a failed SLA Violation -> Warning; Success + Pass -> Healthy. A
    site's overall health is Critical if ANY protection group is Critical, Warning if any is
    Warning, else Healthy - unlike the Network dashboard's rack-health threshold (a handful of
    bad racks out of 15-20 is normal noise), a site typically only has a handful of protection
    groups, so any single failed backup job is itself worth flagging immediately.

    This dashboard template covers 6 sites (AquaArabia, SixFlags, SceneCinema, SEVEN Tabuk,
    SEVEN ABHA, SEVEN Alhamra) so it's ready to use as each site's data becomes available - a
    site with no CSV rows and no Storage row still gets its own tab and page, showing "There is
    no data provided yet" instead of being silently left out.

    NEVER writes/modifies the source Storage workbook - opened read-only, closed without saving.

.PARAMETER BackupCsvPaths
    One or more paths to Cohesity's native "Backup Summary" CSV export. A single file may cover
    more than one site (site is resolved per row, not per file). Omit to leave every site's
    Backup section showing "There is no data provided yet".
.PARAMETER StorageWorkbookPath
    Path to the single "HPE Storage Alletra Capacity Overview - All Sites.xlsx" workbook (one
    "Storage Capacity by Site" table covering every site). Omit to leave every site's Storage
    section showing "There is no data provided yet".
.PARAMETER OutputPath
    Folder the dashboard + log are written to.
#>

[CmdletBinding()]
param(
    [string[]]$BackupCsvPaths = @(),
    [string]$StorageWorkbookPath = '',

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
    # Area is 'Backup' (one row per protection group) or 'Storage' (one row per site).
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
# 2. BACKUP CSV COLLECTION (Cohesity's own native "Backup Summary" export)
# ============================================================================
# Resolves which of this dashboard's 6 sites a single export row belongs to. Checked in this
# order because AquaArabia's Protection Groups share SixFlags' own source hostname, so the name
# prefix has to be checked before the hostname - see the script's own .DESCRIPTION.
function Get-SiteForBackupRow {
    param([string]$ProtectionGroup, [string]$Source)
    if ($ProtectionGroup -match '^AQ-?') { return 'AquaArabia' }
    if ($Source -match 'sevenab')  { return 'SEVEN ABHA' }
    if ($Source -match 'seventb')  { return 'SEVEN Tabuk' }
    if ($Source -match 'sixflags') { return 'SixFlags' }
    if ($ProtectionGroup -match '^AB-') { return 'SEVEN ABHA' }
    if ($ProtectionGroup -match '^TB-') { return 'SEVEN Tabuk' }
    if ($ProtectionGroup -match '^SF-') { return 'SixFlags' }
    return $null
}

# Cohesity's "Bytes Read - GiB" column is a plain number, always in GiB - reformatted here into
# the same "N.NN GiB"/"N.NN TiB" style the rest of this dashboard family uses.
function ConvertTo-DataReadText {
    param([string]$GiBText)
    $v = $null
    if (-not [double]::TryParse($GiBText, [ref]$v) -or $v -le 0) { return '-' }
    if ($v -ge 1024) { return "{0:N2} TiB" -f ($v / 1024) }
    return "{0:N2} GiB" -f $v
}

function Read-BackupCsvReport {
    param([string]$Path)
    if (-not $Path -or -not (Test-Path $Path)) {
        if ($Path) { Write-CheckLog -Site 'n/a' -Object $Path -CheckName 'Backup CSV' -ErrorMessage 'File not found.' }
        return
    }
    Invoke-SafeCheck -Site 'n/a' -ObjectName $Path -CheckName 'Backup CSV' -Script {
        # A couple of report-header lines ("Report Name: ...", "Parameters = ...") sit above the
        # real header row - located by content, not assumed to be a fixed line number.
        $lines = Get-Content -Path $Path
        $headerIndex = -1
        for ($i = 0; $i -lt $lines.Count; $i++) {
            if ($lines[$i] -match '^"?Protection Group"?,') { $headerIndex = $i; break }
        }
        if ($headerIndex -lt 0) {
            Write-CheckLog -Site 'n/a' -Object $Path -CheckName 'Backup CSV header' -ErrorMessage "Could not locate the 'Protection Group' header row in this file."
            return
        }
        $rows = ($lines[$headerIndex..($lines.Count - 1)] -join "`n") | ConvertFrom-Csv

        $lastKnownSite = $null
        foreach ($row in $rows) {
            $pgName = "$($row.'Protection Group')".Trim()
            if (-not $pgName) { continue }
            $source = "$($row.'Registered Source')".Trim()

            $site = Get-SiteForBackupRow -ProtectionGroup $pgName -Source $source
            if (-not $site) {
                if (-not $lastKnownSite) {
                    Write-CheckLog -Site 'n/a' -Object $pgName -CheckName 'Backup CSV site detection' -ErrorMessage "Could not determine the site for '$pgName' (Source: '$source') and no prior row in this file to fall back to - row skipped."
                    continue
                }
                $site = $lastKnownSite
            }
            $lastKnownSite = $site

            $runsN = 0
            [int]::TryParse("$($row.'Number of Runs')", [ref]$runsN) | Out-Null
            $successN = "$($row.'Last Run Successful Objects')".Trim()
            $errorN = "$($row.'Last Run Error Objects')".Trim()
            $lastRunObjText = if ($successN -or $errorN) {
                "$(if ($successN) { $successN } else { '0' }) / $(if ($errorN) { $errorN } else { '0' }) objects"
            } else { '0 / 0 objects' }
            $dataRead = ConvertTo-DataReadText "$($row.'Bytes Read - GiB')"
            $sla = "$($row.Sla)".Trim()
            $lastStatus = "$($row.'Last Run Status')".Trim()
            $bandwidthRaw = "$($row.'Bandwidth (/sec)')".Trim()
            $bandwidth = if ($bandwidthRaw) { "$bandwidthRaw/sec" } else { '-/sec' }

            # 0 runs this period (every other column blank in the export) is its own neutral
            # status - not counted as Healthy, Warning or Critical, same as a template site with
            # no data at all.
            $status = if ($runsN -eq 0 -or -not $lastStatus) { 'Information' }
                elseif ($lastStatus -ieq 'Error') { 'Critical' }
                elseif ($lastStatus -ieq 'Warning' -or $sla -ieq 'Fail') { 'Warning' }
                elseif ($lastStatus -ieq 'Success' -and $sla -ieq 'Pass') { 'Healthy' }
                else { 'Information' }

            $value = "$runsN|$lastRunObjText|$dataRead|$sla|$lastStatus|$bandwidth"
            New-Finding -Site $site -Area 'Backup' -Group $pgName -Object $source -Item 'Backup Job' -Value $value -Status $status
        }
    }
}

# ============================================================================
# 2b. STORAGE CAPACITY WORKBOOK COLLECTION
#     One sheet, one row per site (not one tab per site like the Backup workbook) - the sheet
#     isn't assumed to be named any particular thing, every worksheet in the workbook is tried
#     until the Site/Total Capacity/Used Capacity/Free Capacity header row is found.
# ============================================================================
# Maps a raw "Site" cell value from the Storage sheet (case/space-insensitive) to this
# dashboard's canonical site name. The real sheet seen so far only has Abha/Six Flags/Tabuk/
# Alhamra rows - AquaArabia and SceneCinema aren't in it yet, so their Storage panel stays "There
# is no data provided yet" until a row for them appears. Add an entry here if a site's real label
# in that sheet turns out different from what's guessed.
$StorageSiteNameMap = @{
    'aquaarabia'  = 'AquaArabia'
    'aqua arabia' = 'AquaArabia'
    'sixflags'    = 'SixFlags'
    'six flags'   = 'SixFlags'
    'scenecinema' = 'SceneCinema'
    'scene cinema'= 'SceneCinema'
    'tabuk'       = 'SEVEN Tabuk'
    'abha'        = 'SEVEN ABHA'
    'alhamra'     = 'SEVEN Alhamra'
}

function ConvertFrom-CapacityNumber {
    param([string]$Text)
    if ($Text -match '([\d.]+)') { return [double]$Matches[1] }
    return $null
}

function Read-StorageWorkbook {
    param($Excel, [string]$Path)
    if (-not $Path -or -not (Test-Path $Path)) {
        if ($Path) { Write-CheckLog -Site 'n/a' -Object $Path -CheckName 'Storage workbook' -ErrorMessage 'File not found.' }
        return
    }
    $wb = $Excel.Workbooks.Open($Path, 0, $true)
    try {
        $hdr = $null; $ws = $null
        foreach ($candidate in $wb.Worksheets) {
            $hdr = Find-HeaderRow -Worksheet $candidate -Labels @('Site', 'Total Capacity', 'Used Capacity', 'Free Capacity')
            if ($hdr) { $ws = $candidate; break }
        }
        if (-not $ws) {
            Write-CheckLog -Site 'n/a' -Object $Path -CheckName 'Storage Capacity by Site table' -ErrorMessage 'Could not locate the Site/Total Capacity/Used Capacity/Free Capacity header on any tab.'
            return
        }
        Invoke-SafeCheck -Site 'n/a' -ObjectName $ws.Name -CheckName 'Storage Capacity by Site table' -Script {
            $cols = $hdr.Columns
            $siteCol = $cols['Site']; $totalCol = $cols['Total Capacity']; $usedCol = $cols['Used Capacity']; $freeCol = $cols['Free Capacity']
            $modelCol = $null
            for ($c = 1; $c -le 20; $c++) {
                $v = $ws.Cells.Item($hdr.Row, $c).Value2
                if ($null -ne $v -and "$v".Trim() -ieq 'Model') { $modelCol = $c; break }
            }

            $used = $ws.UsedRange
            $lastRow = $used.Row + $used.Rows.Count - 1
            for ($r = $hdr.Row + 1; $r -le $lastRow; $r++) {
                $siteVal = $ws.Cells.Item($r, $siteCol).Value2
                if (-not $siteVal -or "$siteVal".Trim() -eq '') { continue }
                $rawSite = "$siteVal".Trim()
                if ($rawSite -ieq 'Grand Total') { continue }

                $key = ($rawSite -replace '\s+', ' ').Trim().ToLower()
                if (-not $StorageSiteNameMap.ContainsKey($key)) {
                    Write-CheckLog -Site $rawSite -Object $Path -CheckName 'Storage site mapping' -ErrorMessage "Site label '$rawSite' on the Storage sheet doesn't match any known site - add it to `$StorageSiteNameMap."
                    continue
                }
                $site = $StorageSiteNameMap[$key]

                $totalText = "$($ws.Cells.Item($r, $totalCol).Value2)".Trim()
                $usedText  = "$($ws.Cells.Item($r, $usedCol).Value2)".Trim()
                $freeText  = "$($ws.Cells.Item($r, $freeCol).Value2)".Trim()
                $model     = if ($modelCol) { "$($ws.Cells.Item($r, $modelCol).Value2)".Trim() } else { '' }

                $totalN = ConvertFrom-CapacityNumber $totalText
                $usedN  = ConvertFrom-CapacityNumber $usedText
                $usedPct = if ($totalN -and $totalN -gt 0 -and $usedN -ne $null) { [Math]::Round(($usedN / $totalN) * 100, 1) } else { $null }

                $status = if ($usedPct -eq $null) { 'Information' }
                    elseif ($usedPct -ge 90) { 'Critical' }
                    elseif ($usedPct -ge 75) { 'Warning' }
                    else { 'Healthy' }

                $value = "$totalText|$usedText|$freeText|$usedPct"
                New-Finding -Site $site -Area 'Storage' -Group $site -Object $model -Item 'Storage Capacity' -Value $value -Status $status
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
# This dashboard's 6 sites, in display order - not derived from any workbook content, since
# Backup and Storage each only have rows for whichever sites currently have data.
$SiteDisplayOrder = @('AquaArabia', 'SixFlags', 'SceneCinema', 'SEVEN Tabuk', 'SEVEN ABHA', 'SEVEN Alhamra')

if ($BackupCsvPaths -and $BackupCsvPaths.Count -gt 0) {
    Write-Host "`n=== Collecting Backup ===" -ForegroundColor Green
    foreach ($csvPath in $BackupCsvPaths) {
        Read-BackupCsvReport -Path $csvPath
    }
} else {
    Write-Host "No -BackupCsvPaths supplied - every site's Backup section shows 'There is no data provided yet'." -ForegroundColor Yellow
}

if ($StorageWorkbookPath -and (Test-Path $StorageWorkbookPath)) {
    $Excel = $null
    try {
        $Excel = New-Object -ComObject Excel.Application
        $Excel.Visible = $false
        $Excel.DisplayAlerts = $false
    } catch {
        throw "Microsoft Excel is not available via COM automation on this machine - cannot read the Storage workbook. $($_.Exception.Message)"
    }
    try {
        Write-Host "`n=== Collecting Storage ===" -ForegroundColor Green
        Read-StorageWorkbook -Excel $Excel -Path $StorageWorkbookPath
    } finally {
        $Excel.Quit()
        [System.Runtime.Interopservices.Marshal]::ReleaseComObject($Excel) | Out-Null
    }
} else {
    if ($StorageWorkbookPath) { Write-CheckLog -Site 'n/a' -Object $StorageWorkbookPath -CheckName 'Storage workbook' -ErrorMessage 'File not found.' }
    Write-Host "No -StorageWorkbookPath supplied - every site's Storage section shows 'There is no data provided yet'." -ForegroundColor Yellow
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

$NoBackupSolutionSites = @{
    'SceneCinema' = 'Backup solution license has expired - this environment currently has no backup solution in place. This is a known risk and has already been communicated to management.'
}

function Get-SiteDashboardSummary {
    param([string]$SiteLabel)
    $SiteFindings = @($Global:AllResults | Where-Object { $_.Site -eq $SiteLabel -and $_.Area -eq 'Backup' })
    if ($SiteFindings.Count -eq 0) {
        if ($NoBackupSolutionSites.ContainsKey($SiteLabel)) {
            # A deliberately different case from "no data yet": there IS no backup solution at
            # this site at all (expired license) - a real, already-escalated risk, not a pending
            # data-collection gap - so it counts as Critical rather than the neutral NoData bucket.
            return [pscustomobject]@{
                Site             = $SiteLabel
                HasData          = $true
                NoBackupSolution = $true
                OverallHealth    = 'Critical'
                CriticalCount    = 0
                WarningCount     = 0
                HealthyCount     = 0
                NoRunsCount      = 0
                TotalGroups      = 0
                BackupRows       = @()
                SummaryText      = $NoBackupSolutionSites[$SiteLabel]
            }
        }
        return [pscustomobject]@{
            Site             = $SiteLabel
            HasData          = $false
            NoBackupSolution = $false
            OverallHealth    = 'NoData'
            CriticalCount    = 0
            WarningCount     = 0
            HealthyCount     = 0
            NoRunsCount      = 0
            TotalGroups      = 0
            BackupRows       = @()
            SummaryText      = ''
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
        Site             = $SiteLabel
        HasData          = $true
        NoBackupSolution = $false
        OverallHealth    = $OverallHealth
        CriticalCount    = $criticalCount
        WarningCount     = $warningCount
        HealthyCount     = $healthyCount
        NoRunsCount      = $noRunsCount
        TotalGroups      = $rows.Count
        BackupRows       = $rows
        SummaryText      = $SummaryText
    }
}

function Get-SiteStorageSummary {
    param([string]$SiteLabel)
    $f = $Global:AllResults | Where-Object { $_.Site -eq $SiteLabel -and $_.Area -eq 'Storage' } | Select-Object -First 1
    if (-not $f) {
        return [pscustomobject]@{ HasData = $false; Status = 'NoData'; TotalCapacity = ''; UsedCapacity = ''; FreeCapacity = ''; UsedPct = $null; Model = '' }
    }
    $p = $f.Value -split '\|'
    $usedPctVal = $null
    if ($p[3]) { try { $usedPctVal = [double]$p[3] } catch { } }
    [pscustomobject]@{
        HasData       = $true
        Status        = $f.Status
        TotalCapacity = $p[0]
        UsedCapacity  = $p[1]
        FreeCapacity  = $p[2]
        UsedPct       = $usedPctVal
        Model         = $f.Object
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
    function Get-PctBarColor {
        param([double]$Pct)
        if ($Pct -ge 90) { return '#c62828' }
        if ($Pct -ge 75) { return '#e6a100' }
        return '#2e7d32'
    }

    # --- Overview page ---
    $overviewCards = ($summaries | ForEach-Object {
        $s = $_
        $slug = ConvertTo-Slug $s.Site
        $color = $healthColor[$s.OverallHealth]
        if ($s.NoBackupSolution) {
@"
      <div class="ov-card" onclick="showPage('$slug')" style="border-top-color:$color">
        <div class="ov-head"><h2>$(ConvertTo-HtmlSafe $s.Site)</h2><span class="badge" style="background:$color">No Backup Solution - Risk Communicated</span></div>
        <p class="risk-text">$(ConvertTo-HtmlSafe $s.SummaryText)</p>
        <span class="ov-link">View Page &rarr;</span>
      </div>
"@
            return
        }
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

        $storage = Get-SiteStorageSummary -SiteLabel $s.Site
        $storagePanel = if (-not $storage.HasData) {
@"
        <div class="panel panel-full" style="border-top-color:#8a8f98">
          <h3><span class="n" style="background:#8a8f98">&#128451;&#65039;</span>Storage</h3>
          <div class="nodata-panel">
            <p class="nodata-text-big">There is no data provided yet</p>
          </div>
        </div>
"@
        } else {
            $storageColor = $healthColor[$storage.Status]
            $pctText = if ($storage.UsedPct -ne $null) { "{0:N1}%" -f $storage.UsedPct } else { 'n/a' }
            $barPct = if ($storage.UsedPct -ne $null) { $storage.UsedPct } else { 0 }
@"
        <div class="panel panel-full" style="border-top-color:$storageColor">
          <h3><span class="n" style="background:$storageColor">&#128451;&#65039;</span>Storage <span style="font-weight:normal;font-size:14px;color:#999">($(ConvertTo-HtmlSafe $storage.Model))</span>$(Get-StatusPillHtml -Status $storage.Status -Text $storage.Status)</h3>
          <div class="storage-card">
            <div class="meter-track"><div class="meter-fill" style="width:$barPct%;background:$(Get-PctBarColor $barPct)"></div></div>
            <div class="storage-pct">$pctText used</div>
            <table class="metrics">
              <tr><td>Total Capacity</td><td>$(ConvertTo-HtmlSafe $storage.TotalCapacity)</td></tr>
              <tr><td>Used Capacity</td><td>$(ConvertTo-HtmlSafe $storage.UsedCapacity)</td></tr>
              <tr><td>Free Capacity</td><td>$(ConvertTo-HtmlSafe $storage.FreeCapacity)</td></tr>
              <tr><td>Model</td><td>$(ConvertTo-HtmlSafe $storage.Model)</td></tr>
            </table>
          </div>
        </div>
"@
        }

        if ($s.NoBackupSolution) {
@"
      <section class="page" id="page-$slug" data-site="$(ConvertTo-HtmlSafe $s.Site)">
        <div class="site-hero" style="border-left-color:$color">
          <h1>$(ConvertTo-HtmlSafe $s.Site)</h1>
          <span class="badge big" style="background:$color">No Backup Solution - Risk Communicated</span>
          <div class="meta">Report generated $(ConvertTo-HtmlSafe $RunDateDisplay)</div>
        </div>
        <div class="panel panel-full risk-panel" style="border-top-color:$color">
          <h3><span class="n" style="background:$color">&#9888;&#65039;</span>Backup</h3>
          <p class="risk-text-big">$(ConvertTo-HtmlSafe $s.SummaryText)</p>
        </div>
$storagePanel
      </section>
"@
            return
        }
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
              <tr><td>Last Run Status</td><td>$(ConvertTo-HtmlSafe $row.LastRunStatus)</td></tr>
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
  .risk-text { color:#c62828; font-weight:600; margin:4px 0 0; }
  .risk-panel { text-align:center; padding:12px 24px 28px; }
  .risk-text-big { font-size:16px; font-weight:600; color:#c62828; max-width:640px; margin:0 auto; line-height:1.6; }
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

  /* Storage - capacity meter bar + table, one per site */
  .storage-card { max-width:420px; }
  .storage-card .meter-track { height:12px; border-radius:6px; overflow:hidden; background:#eef2f5; margin-bottom:8px; }
  .storage-card .meter-fill { height:100%; border-radius:6px; }
  .storage-card .storage-pct { font-size:13px; color:#666; font-weight:600; margin-bottom:12px; }

  /* No-data placeholder (template site, or a site's Storage panel until that data is supplied) */
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
$SiteLabels = $SiteDisplayOrder
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
