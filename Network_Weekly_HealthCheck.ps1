#Requires -Version 5.1
<#
.SYNOPSIS
    Network Weekly Health Check - reads the IDF Room Report and DC RackWise Health Check
    workbooks for AquaArabia and SixFlags and produces a single combined HTML dashboard, in the
    same visual style as VMware_Weekly_HealthCheck.ps1's dashboard.

.DESCRIPTION
    This dashboard covers the physical Network/data-center layer: IDF rooms/cabinets/access
    control AND rack-level physical health (power, PDU, temperature, cabling) plus the devices
    mounted in those racks. Virtualization Infrastructure (VMware - its own existing dashboard)
    and Backup & Storage (backup job/replication status) are separate dashboards, not part of
    this script.

    Reads 4 source workbooks (2 per site):
      - <Site> IDF Room Report: a single "Summary" tab that lists ONLY the issues found (an
        exceptions list, not a full per-switch checklist) - one row per issue, each naming the
        affected zone/location code(s). The two companies word this slightly differently
        (AquaArabia: "Issues" + a free-text "Zone N: loc1, loc2" cell per issue type; SixFlags:
        "No." + "Observation" + a flat comma-separated "Zone - IDF(s)" list per observation) but
        both are read the same way: find the issue-label column, the affected-location column is
        always the one immediately to its right.
      - <Site> DC RackWise Health Check: a "Summary" tab with one row per rack (Health %,
        Status, Rack/Node check counts - note the "Status" text on this tab is manually set by
        the engineer and is NOT reliably tied to Health % in the real data, so this script
        derives each rack's own Healthy/Warning/Critical status from its Health % instead), and
        one tab per rack (A01, B01, P1, S1, ...) with a rack-level checklist (Check Item /
        Status / Comments, read until the first blank row rather than matched against a fixed
        item list, since the exact wording/row count varies slightly by company) followed by a
        per-device checklist anchored on the "RU" column further down the same sheet.

    Every tab is located by searching for its known column headers rather than assuming a
    fixed row/column position, since the real workbooks have inconsistent header rows (a
    banner/logo above the real header, blank spacer rows, typos in item text) - a tab whose
    headers can't be found is skipped and logged, never guessed at.

    Produces ONE combined HTML dashboard (no Word/PDF) covering both sites, matching the
    VMware dashboard's look: rectangular site tabs (sized to align with the KPI tile row below
    them), KPI stat tiles, an IDF Room Issues panel, a Rack Health panel, and an itemized
    Action Plan of every Rack/Node item needing attention.

    This script never embeds evidence photos - that is handled by a separate, optional script
    (Network_Evidence_Photos.ps1) run afterward against the dashboard this script produces, so
    a month with no photos supplied needs no change here.

    NEVER writes/modifies the source workbooks - opened read-only, closed without saving.

.PARAMETER AquaArabiaIdfPath
    Path to AquaArabia's IDF Room Report workbook. Omit to skip AquaArabia's IDF section.
.PARAMETER AquaArabiaRackPath
    Path to AquaArabia's DC RackWise Health Check workbook. Omit to skip AquaArabia's rack section.
.PARAMETER SixFlagsIdfPath
    Path to SixFlags' IDF Room Report workbook. Omit to skip SixFlags' IDF section.
.PARAMETER SixFlagsRackPath
    Path to SixFlags' DC RackWise Health Check workbook. Omit to skip SixFlags' rack section.
.PARAMETER OutputPath
    Folder the dashboard + log are written to.
#>

[CmdletBinding()]
param(
    [string]$AquaArabiaIdfPath = '',
    [string]$AquaArabiaRackPath = '',
    [string]$SixFlagsIdfPath = '',
    [string]$SixFlagsRackPath = '',

    [string]$OutputPath = (Join-Path $PSScriptRoot "Network_HealthCheck_Reports")
)

$ScriptBuild = '2026-10-01-04-network-with-rackwise'
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
    # Area: IDF | Rack | Node. Group: zone/location label (IDF) or rack ID (Rack/Node).
    # RowNum: the IDF Summary row this finding came from (used only to match evidence photos,
    # which are supplied separately and named by row number - see Network_Evidence_Photos.ps1).
    $obj = [pscustomobject]@{
        Site = $Site; Area = $Area; Group = $Group; Object = $Object
        Item = $Item; Value = $Value; Status = $Status; Notes = $Notes; RowNum = $RowNum
    }
    $Global:AllResults.Add($obj) | Out-Null
}

# Maps a rack/node checklist answer to a Status. Real answers seen in the workbooks go beyond
# plain Yes/No: 'Ok' is used interchangeably with 'Yes' on the PDU columns, 'Pending' means the
# device isn't fully set up yet (not a failure, but not confirmed healthy either), and 'N/A'
# means the check doesn't apply to this device (e.g. a passive patch panel). -AlarmStyle is for
# the two inverted-polarity columns ("Any Disk Alarms", "Front/Back Led Alarms") where an empty
# answer or 'None'/'N/A' is the GOOD outcome and anything else indicates a real alarm.
function Get-YesNoStatus {
    param([string]$Value, [switch]$AlarmStyle)
    $v = "$Value".Trim()
    if ($AlarmStyle) {
        if ($v -eq '' -or $v -ieq 'None' -or $v -ieq 'N/A' -or $v -ieq 'No') { return 'Healthy' }
        return 'Critical'
    }
    switch -Regex ($v) {
        '^(Yes|Ok)$' { return 'Healthy' }
        '^No$'       { return 'Critical' }
        '^Pending$'  { return 'Warning' }
        '^N/A$'      { return 'Information' }
        default      { return 'Information' }
    }
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
#    real header, blank spacer rows, typo'd item text) that make a fixed-position read silently
#    wrong instead of simply failing.
# ============================================================================

# Finds the first row (within MaxRow/MaxCol) where every label in $Labels appears as an exact
# (trimmed, case-insensitive) cell match somewhere on that row. Returns $null if not found.
function Find-HeaderRow {
    param($Worksheet, [string[]]$Labels, [int]$MaxRow = 25, [int]$MaxCol = 40)
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

# Finds the first cell (within MaxRow/MaxCol) whose text CONTAINS $Contains (case-insensitive).
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
# 3. DC RACKWISE HEALTH CHECK COLLECTION
# ============================================================================
# Fixed, confirmed-consistent column order of the per-device grid starting 2 columns after "RU"
# ("RU" then "Expected Device" are read separately, not part of this Yes/No/Ok grid).
$NodeGridItems = @('Mac Address / IP','Device Powered On','Device PDU-1','Device PDU-2','Any Disk Alarms','Device Labeling','Fiber/UTP Link Labeling','Temp OK','Link OK','Front/Back Led Alarms')
$NodeGridAlarmItems = @('Any Disk Alarms','Front/Back Led Alarms')

function Read-RackSummaryTab {
    param($Worksheet, [string]$Site)
    $hdr = Find-HeaderRow -Worksheet $Worksheet -Labels @('Rack','Health %','Status') -MaxRow 20 -MaxCol 40
    if (-not $hdr) { throw "Could not locate the Rack/Health %/Status header on the Summary tab." }
    $cols = $hdr.Columns
    $rackCol = $cols['Rack']; $healthCol = $cols['Health %']
    $yesCol = $null; $noCol = $null; $nodeYesCol = $null; $nodeNoCol = $null; $commentsCol = $null
    foreach ($name in @('Rack Checks (Yes)','Rack Checks (No)','Node Yes','Node No','Comments')) {
        for ($c = 1; $c -le 40; $c++) {
            $v = $Worksheet.Cells.Item($hdr.Row, $c).Value2
            if ($null -ne $v -and "$v".Trim() -ieq $name) {
                switch ($name) {
                    'Rack Checks (Yes)' { $yesCol = $c }
                    'Rack Checks (No)'  { $noCol = $c }
                    'Node Yes'          { $nodeYesCol = $c }
                    'Node No'           { $nodeNoCol = $c }
                    'Comments'          { $commentsCol = $c }
                }
                break
            }
        }
    }

    $used = $Worksheet.UsedRange
    $lastRow = $used.Row + $used.Rows.Count - 1
    $racks = @()
    for ($r = $hdr.Row + 1; $r -le $lastRow; $r++) {
        $rackVal = $Worksheet.Cells.Item($r, $rackCol).Value2
        if (-not $rackVal -or "$rackVal".Trim() -eq '') { continue }
        $rackName = "$rackVal".Trim()
        $healthVal = $Worksheet.Cells.Item($r, $healthCol).Value2
        $rackYes = if ($yesCol) { $Worksheet.Cells.Item($r, $yesCol).Value2 } else { $null }
        $rackNo  = if ($noCol) { $Worksheet.Cells.Item($r, $noCol).Value2 } else { $null }
        $nodeYes = if ($nodeYesCol) { $Worksheet.Cells.Item($r, $nodeYesCol).Value2 } else { $null }
        $nodeNo  = if ($nodeNoCol) { $Worksheet.Cells.Item($r, $nodeNoCol).Value2 } else { $null }
        $comments = if ($commentsCol) { ConvertTo-CompactNote "$($Worksheet.Cells.Item($r, $commentsCol).Value2)" } else { '' }

        $healthPct = 0.0
        if ($healthVal) { try { $healthPct = [double]$healthVal * 100 } catch { } }
        # The sheet's own "Status" column is set manually and is not reliably tied to Health % in
        # the real data (racks at 45% have shown "Healthy" while racks at 59% show "Attention") -
        # so the rack's Status here is derived from Health % instead, for a consistent dashboard.
        $status = if ($healthPct -ge 90) { 'Healthy' } elseif ($healthPct -ge 75) { 'Warning' } else { 'Critical' }

        New-Finding -Site $Site -Area 'Rack' -Group $rackName -Object $rackName -Item 'Rack Health' `
            -Value "$([Math]::Round($healthPct,1))|$rackYes|$rackNo|$nodeYes|$nodeNo" -Status $status -Notes $comments
        $racks += $rackName
    }
    return $racks
}

function Read-RackDetailTab {
    param($Worksheet, [string]$Site, [string]$RackName)
    # Rack-level checklist: Column = Check Item, next column = Status, 2 columns over = Comments.
    # Anchored by finding the literal "Check Item" / "Status" header, then read every row until
    # the first blank Check Item (the natural gap before the Node/Device-Wise section below) -
    # the exact item wording/count varies slightly between companies (and has typos), so this is
    # read verbatim rather than matched against a fixed, hardcoded item list.
    $hdr = Find-HeaderRow -Worksheet $Worksheet -Labels @('Check Item','Status') -MaxRow 15 -MaxCol 10
    if ($hdr) {
        $itemCol = $hdr.Columns['Check Item']; $statusCol = $hdr.Columns['Status']
        $commentsCol = $itemCol + 2
        for ($r = $hdr.Row + 1; $r -le ($hdr.Row + 25); $r++) {
            $itemVal = $Worksheet.Cells.Item($r, $itemCol).Value2
            if (-not $itemVal -or "$itemVal".Trim() -eq '') { break }
            $itemName = "$itemVal".Trim()
            $statusVal = $Worksheet.Cells.Item($r, $statusCol).Value2
            if (-not $statusVal -or "$statusVal".Trim() -eq '') { continue }
            $commentsVal = ConvertTo-CompactNote "$($Worksheet.Cells.Item($r, $commentsCol).Value2)"
            New-Finding -Site $Site -Area 'Rack' -Group $RackName -Object $RackName -Item $itemName `
                -Value "$statusVal".Trim() -Status (Get-YesNoStatus "$statusVal") -Notes $commentsVal
        }
    }

    # Node/device grid: anchored on the literal "RU" column header - column order after it is
    # fixed/confirmed, so read by position once that anchor is found. "Expected Device" (the
    # column right after RU) is descriptive metadata folded into the device label, not one of
    # the Yes/No/Ok check columns.
    $ruHdr = Find-HeaderRow -Worksheet $Worksheet -Labels @('RU') -MaxRow 40 -MaxCol 20
    if (-not $ruHdr) { return }
    $ruCol = $ruHdr.Columns['RU']
    $used = $Worksheet.UsedRange
    $lastRow = $used.Row + $used.Rows.Count - 1
    for ($r = $ruHdr.Row + 1; $r -le $lastRow; $r++) {
        $ruVal = $Worksheet.Cells.Item($r, $ruCol).Value2
        if (-not $ruVal -or "$ruVal".Trim() -eq '') { continue }
        $deviceVal = $Worksheet.Cells.Item($r, $ruCol + 1).Value2
        $deviceLabel = "RU$("$ruVal".Trim()) $("$deviceVal".Trim())"
        $commentsVal = ConvertTo-CompactNote "$($Worksheet.Cells.Item($r, $ruCol + 12).Value2)"
        for ($i = 0; $i -lt $NodeGridItems.Count; $i++) {
            $itemName = $NodeGridItems[$i]
            $cellVal = $Worksheet.Cells.Item($r, $ruCol + 2 + $i).Value2
            if ($null -eq $cellVal -or "$cellVal".Trim() -eq '' -or "$cellVal".Trim() -ieq 'N/A') { continue }
            $isAlarm = $itemName -in $NodeGridAlarmItems
            $status = Get-YesNoStatus "$cellVal" -AlarmStyle:$isAlarm
            $notes = if ($status -in 'Critical','Warning') { $commentsVal } else { '' }
            New-Finding -Site $Site -Area 'Node' -Group $RackName -Object $deviceLabel `
                -Item $itemName -Value "$cellVal".Trim() -Status $status -Notes $notes
        }
    }
}

function Read-RackWorkbook {
    param($Excel, [string]$Site, [string]$Path)
    if (-not $Path -or -not (Test-Path $Path)) {
        if ($Path) { Write-CheckLog -Site $Site -Object $Path -CheckName 'RackWise workbook' -ErrorMessage 'File not found.' }
        return
    }
    $wb = $Excel.Workbooks.Open($Path, 0, $true)
    try {
        $summaryWs = $wb.Worksheets | Where-Object { $_.Name.Trim() -ieq 'Summary' } | Select-Object -First 1
        if (-not $summaryWs) {
            Write-CheckLog -Site $Site -Object $Path -CheckName 'RackWise Summary tab' -ErrorMessage "No tab named 'Summary' found."
            return
        }
        $racks = Invoke-SafeCheck -Site $Site -ObjectName 'Summary' -CheckName 'RackWise Summary tab' -Script {
            Read-RackSummaryTab -Worksheet $summaryWs -Site $Site
        }
        if (-not $racks) { return }
        # Not every rack tab is necessarily listed on the Summary tab (e.g. a rack was surveyed
        # on its own tab but never rolled into Summary) - this script can only discover racks
        # that ARE listed there; any Summary gap like that is a workbook data-entry issue, not
        # something this script can detect on its own.
        foreach ($rackName in $racks) {
            $rackWs = $wb.Worksheets | Where-Object { $_.Name.Trim() -ieq $rackName } | Select-Object -First 1
            if (-not $rackWs) {
                Write-CheckLog -Site $Site -Object $rackName -CheckName 'RackWise rack tab' -ErrorMessage "No tab named '$rackName' found - rack-level/node detail skipped (Summary rollup still recorded)."
                continue
            }
            Invoke-SafeCheck -Site $Site -ObjectName $rackName -CheckName 'RackWise rack tab' -Script {
                Read-RackDetailTab -Worksheet $rackWs -Site $Site -RackName $rackName
            }
        }
    } finally {
        $wb.Close($false)
        [System.Runtime.Interopservices.Marshal]::ReleaseComObject($wb) | Out-Null
    }
}

# ============================================================================
# 4. RUN COLLECTION
# ============================================================================
$Sites = @(
    [pscustomobject]@{ Site = 'AquaArabia'; IdfPath = $AquaArabiaIdfPath; RackPath = $AquaArabiaRackPath }
    [pscustomobject]@{ Site = 'SixFlags';   IdfPath = $SixFlagsIdfPath;   RackPath = $SixFlagsRackPath }
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
    Read-RackWorkbook -Excel $Excel -Site $s.Site -Path $s.RackPath
}

$Excel.Quit()
[System.Runtime.Interopservices.Marshal]::ReleaseComObject($Excel) | Out-Null

# ============================================================================
# 5. DASHBOARD GENERATION HELPERS
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

# Computes every per-site aggregate the dashboard needs directly from $Global:AllResults -
# the IDF issues list, per-rack health rollup, and the combined itemized Action Plan.
function Get-SiteDashboardSummary {
    param([string]$SiteLabel)
    $SiteFindings = $Global:AllResults | Where-Object { $_.Site -eq $SiteLabel }
    if (-not $SiteFindings) { return $null }

    $CritCount = @($SiteFindings | Where-Object { $_.Status -eq 'Critical' }).Count
    $WarnCount = @($SiteFindings | Where-Object { $_.Status -eq 'Warning' }).Count

    # --- IDF issues (exceptions list - grouped back up by source row) ---
    $idfFindings = @($SiteFindings | Where-Object { $_.Area -eq 'IDF' })
    $idfIssueRows = @($idfFindings | Group-Object RowNum | ForEach-Object {
        $grp = $_.Group
        $first = $grp[0]
        $locList = @($grp | ForEach-Object {
            if ($_.Object -eq '(unspecified location)') { $null }
            elseif ($_.Group) { "$($_.Group): $($_.Object)" }
            else { $_.Object }
        } | Where-Object { $_ })
        [pscustomobject]@{
            RowNum    = $first.RowNum
            Issue     = $first.Item
            Locations = $locList
            Notes     = $first.Notes
        }
    } | Sort-Object { [int]($_.RowNum -replace '\D','0') })
    $idfIssueCount = $idfIssueRows.Count
    $idfLocationCount = @($idfFindings | Where-Object { $_.Object -ne '(unspecified location)' } | Select-Object -ExpandProperty Object -Unique).Count

    # --- Rack rollup ---
    $rackHealthFindings = @($SiteFindings | Where-Object { $_.Area -eq 'Rack' -and $_.Item -eq 'Rack Health' })
    $rackRowsUnsorted = @(foreach ($rf in $rackHealthFindings) {
        $p = $rf.Value -split '\|'
        $rackYesN = 0; $rackNoN = 0; $nodeYesN = 0; $nodeNoN = 0
        [int]::TryParse("$($p[1])", [ref]$rackYesN) | Out-Null
        [int]::TryParse("$($p[2])", [ref]$rackNoN) | Out-Null
        [int]::TryParse("$($p[3])", [ref]$nodeYesN) | Out-Null
        [int]::TryParse("$($p[4])", [ref]$nodeNoN) | Out-Null
        [pscustomobject]@{
            Rack = $rf.Group; HealthPct = [double]$p[0]
            RackYes = $rackYesN; RackNo = $rackNoN; NodeYes = $nodeYesN; NodeNo = $nodeNoN
            Status = $rf.Status; Comments = $rf.Notes
        }
    })
    $rackRows = @($rackRowsUnsorted | Sort-Object Rack)
    $rackCount = $rackRows.Count
    $rackHealthyCount = @($rackRows | Where-Object { $_.Status -eq 'Healthy' }).Count
    $rackAttentionCount = @($rackRows | Where-Object { $_.Status -ne 'Healthy' }).Count
    $rackCriticalCount = @($rackRows | Where-Object { $_.Status -eq 'Critical' }).Count
    $avgHealthPct = if ($rackCount -gt 0) { ($rackRows | Measure-Object -Property HealthPct -Average).Average } else { $null }

    # A couple of bad racks out of what can be 15-20 per site shouldn't flip the WHOLE site to a
    # red "Critical" badge - that collapsed "Healthy Sites" to 0 even when the overwhelming
    # majority of racks were fine. Critical is reserved for sites where the problem is actually
    # widespread (more than a quarter of racks Critical, or the site-wide average itself is
    # poor); an isolated issue or two still shows as Warning, fully itemized in Action Plan.
    $criticalRackFraction = if ($rackCount -gt 0) { $rackCriticalCount / $rackCount } else { 0 }
    $OverallHealth = if (($avgHealthPct -ne $null -and $avgHealthPct -lt 70) -or $criticalRackFraction -gt 0.25) { 'Critical' }
        elseif ($CritCount -gt 0 -or $WarnCount -gt 0) { 'Warning' }
        else { 'Healthy' }

    # --- Action plan: every Rack/Node item needing attention (IDF issues get their own panel
    # instead, and the per-rack Health rollup is already shown in the Rack Health panel). ---
    $ActionItems = @($SiteFindings | Where-Object { $_.Status -in 'Critical','Warning' -and $_.Area -in 'Rack','Node' -and $_.Item -ne 'Rack Health' } |
        Sort-Object @{Expression = { if ($_.Status -eq 'Critical') { 0 } else { 1 } }} |
        ForEach-Object {
            [pscustomobject]@{
                Severity = if ($_.Status -eq 'Critical') { 'High' } else { 'Medium' }
                Object   = $_.Object
                Item     = $_.Item
                Value    = $_.Value
                Notes    = $_.Notes
                Group    = $_.Group
            }
        })

    $SummaryText = if ($CritCount -gt 0) {
        "$SiteLabel has $CritCount critical and $WarnCount warning item(s) across $idfIssueCount IDF issue(s) and $rackCount racks that require attention."
    } elseif ($WarnCount -gt 0) {
        "$SiteLabel is stable overall, with $WarnCount minor item(s) flagged across $idfIssueCount IDF issue(s) and $rackCount racks - see the panels below."
    } else {
        "$SiteLabel's IDF rooms and $rackCount racks are all reporting healthy with no items flagged."
    }

    [pscustomobject]@{
        Site               = $SiteLabel
        OverallHealth      = $OverallHealth
        HighRisk           = $CritCount
        MediumRisk         = $WarnCount
        IdfIssueCount      = $idfIssueCount
        IdfLocationCount   = $idfLocationCount
        IdfIssueRows       = $idfIssueRows
        RackCount          = $rackCount
        RackHealthyCount   = $rackHealthyCount
        RackAttentionCount = $rackAttentionCount
        AvgHealthPct       = $avgHealthPct
        RackRows           = $rackRows
        ActionItems        = $ActionItems
        SummaryText        = $SummaryText
    }
}

function Write-DashboardHtml {
    param([string[]]$SiteLabels, [string]$OutputPath, [string]$RunDateDisplay)

    $summaries = @($SiteLabels | ForEach-Object { Get-SiteDashboardSummary -SiteLabel $_ } | Where-Object { $_ } | Sort-Object Site)
    if ($summaries.Count -eq 0) { return }

    $healthColor = @{ Healthy = '#2e7d32'; Warning = '#e6a100'; Critical = '#c62828' }
    $healthLabelText = @{ Healthy = 'Healthy - No Issues Detected'; Warning = 'Healthy - Minor Issues Detected'; Critical = 'Attention Required - Critical Issues' }
    $healthCounts = @{
        Healthy  = @($summaries | Where-Object { $_.OverallHealth -ne 'Critical' }).Count
        Warning  = @($summaries | Where-Object { $_.OverallHealth -eq 'Warning' }).Count
        Critical = @($summaries | Where-Object { $_.OverallHealth -eq 'Critical' }).Count
    }
    $totalHigh      = ($summaries | Measure-Object -Property HighRisk -Sum).Sum
    $totalMedium    = ($summaries | Measure-Object -Property MediumRisk -Sum).Sum
    $totalIdfIssues = ($summaries | Measure-Object -Property IdfIssueCount -Sum).Sum
    $totalRacks     = ($summaries | Measure-Object -Property RackCount -Sum).Sum

    $tabPalette = @('#2563EB','#7C3AED','#0D9488','#C026D3','#EA580C','#4F46E5','#DB2777','#0EA5E9')
    $tileBlue = '#1565C0'; $tilePurple = '#6A1B9A'; $tileTeal = '#00897B'; $tileIndigo = '#283593'

    function Get-StatusPillHtml {
        param([string]$Status, [string]$Text)
        $variant = switch ($Status) { 'Healthy' { 'ok' }; 'Warning' { 'warn' }; 'Critical' { 'bad' }; default { 'info' } }
        "<span class=`"status-pill $variant`"><span class=`"dot`"></span>$(ConvertTo-HtmlSafe $Text)</span>"
    }
    function Get-PctBarColor {
        param([double]$Pct)
        if ($Pct -ge 90) { return '#2e7d32' }
        if ($Pct -ge 75) { return '#e6a100' }
        return '#c62828'
    }

    # --- Overview page ---
    $overviewCards = ($summaries | ForEach-Object {
        $s = $_
        $slug = ConvertTo-Slug $s.Site
        $color = $healthColor[$s.OverallHealth]
        $avgHealthText = if ($s.AvgHealthPct -ne $null) { "{0:N1}%" -f $s.AvgHealthPct } else { 'n/a' }
@"
      <div class="ov-card" onclick="showPage('$slug')" style="border-top-color:$color">
        <div class="ov-head"><h2>$(ConvertTo-HtmlSafe $s.Site)</h2><span class="badge" style="background:$color">$(ConvertTo-HtmlSafe $healthLabelText[$s.OverallHealth])</span></div>
        <div class="risk-row"><span class="risk risk-high">High: $($s.HighRisk)</span><span class="risk risk-med">Medium: $($s.MediumRisk)</span></div>
        <table class="metrics">
          <tr><td>IDF Issues Reported</td><td>$($s.IdfIssueCount) ($($s.IdfLocationCount) locations)</td></tr>
          <tr><td>Racks (Healthy / Attention)</td><td>$($s.RackCount) ($($s.RackHealthyCount) / $($s.RackAttentionCount))</td></tr>
          <tr><td>Average Rack Health</td><td>$avgHealthText</td></tr>
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

        # Each card carries a data-idf-row attribute so Network_Evidence_Photos.ps1 (a separate,
        # optional script) can find the right insertion point when photos are supplied for a
        # given month - this script itself never depends on photos existing.
        $idfRows = if ($s.IdfIssueRows.Count -gt 0) {
            ($s.IdfIssueRows | ForEach-Object {
                $chips = if ($_.Locations.Count -gt 0) {
                    (($_.Locations | ForEach-Object { "<span class=`"loc-chip`">$(ConvertTo-HtmlSafe $_)</span>" }) -join '')
                } else { '<span class="loc-chip loc-chip-empty">No location specified</span>' }
@"
          <div class="idf-issue-card" data-idf-row="$(ConvertTo-HtmlSafe $_.RowNum)">
            <div class="issue-no">$(ConvertTo-HtmlSafe $_.RowNum)</div>
            <div class="issue-body">
              <div class="issue-title">$(ConvertTo-HtmlSafe $_.Issue)</div>
              <div class="loc-chips">$chips</div>
$(if ($_.Notes) { "              <div class=`"issue-notes`">$(ConvertTo-HtmlSafe $_.Notes)</div>" })
              <div class="idf-evidence-cell"></div>
            </div>
          </div>
"@
            }) -join "`n"
        } else { '<div class="center-callout">' + (Get-StatusPillHtml -Status 'Healthy' -Text 'No IDF issues reported for this site') + '</div>' }

        $rackCards = if ($s.RackRows.Count -gt 0) {
            ($s.RackRows | ForEach-Object {
                $rackTotal = $_.RackYes + $_.RackNo
                $nodeTotal = $_.NodeYes + $_.NodeNo
                $trackColor = switch ($_.Status) { 'Critical' { '#fdecea' }; 'Warning' { '#fff6e0' }; default { '#e8f5e9' } }
@"
          <div class="rack-card">
            <div class="rack-id"><span>$(ConvertTo-HtmlSafe $_.Rack)</span>$(Get-StatusPillHtml -Status $_.Status -Text $_.Status)</div>
            <div class="meter-track" style="background:$trackColor"><div class="meter-fill" style="width:$($_.HealthPct)%;background:$(Get-PctBarColor $_.HealthPct)"></div></div>
            <div class="rack-health-pct">$("{0:N1}" -f $_.HealthPct)% health</div>
            <div class="rack-stats">
              <span>Rack checks: <strong>$($_.RackYes)/$rackTotal</strong></span>
              <span>Device checks: <strong>$($_.NodeYes)/$nodeTotal</strong></span>
            </div>
$(if ($_.Comments) { "            <div class=`"rack-comment`">$(ConvertTo-HtmlSafe $_.Comments)</div>" })
          </div>
"@
            }) -join "`n"
        } else { '<div class="center-callout">' + (Get-StatusPillHtml -Status 'Information' -Text 'No rack data found for this site') + '</div>' }

        $actionRows = if ($s.ActionItems.Count -gt 0) {
            ($s.ActionItems | ForEach-Object {
                $sevClass = if ($_.Severity -eq 'High') { 'high' } else { 'med' }
                $location = if ($_.Group -and $_.Group -ne $_.Object) { "$($_.Group) / $($_.Object)" } else { "$($_.Object)" }
@"
          <div class="action-card $sevClass">
            <span class="sev-badge $sevClass">$($_.Severity)</span>
            <div class="ac-body">
              <div class="ac-title">$(ConvertTo-HtmlSafe $_.Item): $(ConvertTo-HtmlSafe $_.Value)</div>
              <div class="ac-sub">$(ConvertTo-HtmlSafe $location)$(if ($_.Notes) { " - $(ConvertTo-HtmlSafe $_.Notes)" })</div>
            </div>
          </div>
"@
            }) -join "`n"
        } else { $null }
        $actionPlanBody = if ($actionRows) {
            "<div class=`"action-grid`">`n$actionRows`n</div>"
        } else {
            "<div class=`"center-callout`">$(Get-StatusPillHtml -Status 'Healthy' -Text 'No Rack/Node items flagged for this site')</div>"
        }

        $avgHealthValueText = if ($s.AvgHealthPct -ne $null) { "{0:N1}%" -f $s.AvgHealthPct } else { 'n/a' }

@"
      <section class="page" id="page-$slug" data-site="$(ConvertTo-HtmlSafe $s.Site)">
        <div class="site-hero" style="border-left-color:$color">
          <h1>$(ConvertTo-HtmlSafe $s.Site)</h1>
          <span class="badge big" style="background:$color">$(ConvertTo-HtmlSafe $healthLabelText[$s.OverallHealth])</span>
          <div class="meta">Report generated $(ConvertTo-HtmlSafe $RunDateDisplay)</div>
        </div>

        <div class="risk-row big">
          <span class="risk risk-high">High Risk: $($s.HighRisk)</span>
          <span class="risk risk-med">Medium Risk: $($s.MediumRisk)</span>
        </div>

        <div class="tile-row">
          <div class="tile" style="background:$tileBlue"><span class="num">$($s.IdfIssueCount)</span><span class="label">IDF Issues Reported</span></div>
          <div class="tile" style="background:$tilePurple"><span class="num">$($s.IdfLocationCount)</span><span class="label">Locations Affected</span></div>
          <div class="tile" style="background:$tileTeal"><span class="num">$($s.RackCount)</span><span class="label">Racks</span></div>
          <div class="tile" style="background:#2e7d32"><span class="num">$($s.RackHealthyCount)</span><span class="label">Racks Healthy</span></div>
          <div class="tile" style="background:#c62828"><span class="num">$($s.RackAttentionCount)</span><span class="label">Racks Attention</span></div>
          <div class="tile" style="background:$tileIndigo"><span class="num" style="font-size:18px">$avgHealthValueText</span><span class="label">Avg Rack Health</span></div>
        </div>

        <div class="panel-grid">
          <div class="panel panel-full" style="border-top-color:$tilePurple">
            <h3><span class="n" style="background:$tilePurple">&#127968;</span>IDF Room Issues <span style="font-weight:normal;font-size:14px;color:#999">($($s.IdfIssueCount) reported)</span></h3>
            <div class="idf-issue-list">
$idfRows
            </div>
          </div>

          <div class="panel panel-full" style="border-top-color:$tileTeal">
            <h3><span class="n" style="background:$tileTeal">&#128451;&#65039;</span>Rack Health <span style="font-weight:normal;font-size:14px;color:#999">($($s.RackCount) racks)</span></h3>
            <div class="rack-grid">
$rackCards
            </div>
          </div>

          <div class="panel panel-full" style="border-top-color:#1E3A5F">
            <h3><span class="n" style="background:#1E3A5F">&#128203;</span>Summary</h3>
            <p class="summary-text">$(ConvertTo-HtmlSafe $s.SummaryText)</p>
          </div>
        </div>

        <div class="panel" style="border-top-color:#e6a100">
          <h3><span class="n" style="background:#e6a100">&#9888;&#65039;</span>Action Plan - Rack/Node Items Requiring Attention</h3>
$actionPlanBody
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

  .panel-grid { display:grid; grid-template-columns: repeat(auto-fit, minmax(380px, 1fr)); gap:24px; margin-bottom:24px; }
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
  table.metrics td, table.metrics th { padding:10px 6px; border-bottom:1px solid #eee; vertical-align:top; }
  table.metrics td:first-child { color:#666; width:55%; }
  table.metrics td:last-child:not(:first-child) { text-align:right; font-weight:600; }
  table.metrics.wide th { text-align:center; color:#999; font-size:13px; text-transform:uppercase; padding-bottom:10px; }
  table.metrics.wide th:first-child { text-align:left; }
  table.metrics.wide td { text-align:center; }
  table.metrics.wide td:first-child { text-align:left; color:#333; font-weight:600; width:auto; }
  table.metrics.wide td:last-child:not(:first-child) { text-align:center; font-weight:normal; }
  .panel-full { grid-column: 1 / -1; }
  .summary-text { color:#444; font-size:15px; line-height:1.6; }
  .center-callout { text-align:center; padding:8px 0 18px; }

  /* IDF Room Issues - one card per reported issue */
  .idf-issue-list { display:flex; flex-direction:column; gap:12px; }
  .idf-issue-card { display:flex; gap:16px; padding:16px 18px; border-radius:8px; background:#fff; border-left:5px solid #e6a100; box-shadow:0 1px 3px rgba(0,0,0,0.08); }
  .idf-issue-card .issue-no { flex-shrink:0; width:34px; height:34px; border-radius:50%; background:#e6a100; color:#fff; display:flex; align-items:center; justify-content:center; font-weight:bold; font-size:14px; }
  .idf-issue-card .issue-body { flex:1; min-width:0; }
  .idf-issue-card .issue-title { font-weight:bold; font-size:15px; color:#1a1a1a; margin-bottom:8px; }
  .idf-issue-card .issue-notes { font-size:13px; color:#8a6100; margin-top:6px; font-style:italic; }
  .loc-chips { display:flex; flex-wrap:wrap; gap:6px; }
  .loc-chip { display:inline-block; background:#eef2f5; color:#444; font-size:12px; font-weight:600; padding:4px 11px; border-radius:12px; }
  .loc-chip-empty { background:#f5f5f5; color:#999; font-weight:normal; font-style:italic; }
  .idf-evidence-cell:empty { display:none; }
  .idf-evidence-cell { margin-top:10px; display:flex; gap:8px; flex-wrap:wrap; }
  .idf-evidence-cell img { max-width:90px; max-height:90px; border-radius:4px; box-shadow:0 1px 3px rgba(0,0,0,0.3); cursor:zoom-in; }

  /* Rack Health - one card per rack */
  .rack-grid { display:grid; grid-template-columns: repeat(auto-fill, minmax(210px, 1fr)); gap:14px; }
  .rack-card { background:#fff; border:1px solid #eee; border-radius:8px; padding:16px 18px; box-shadow:0 1px 3px rgba(0,0,0,0.06); }
  .rack-card .rack-id { display:flex; justify-content:space-between; align-items:center; font-weight:bold; font-size:17px; color:#1a1a1a; margin-bottom:10px; }
  .rack-card .meter-track { height:9px; border-radius:5px; overflow:hidden; margin-bottom:6px; }
  .rack-card .meter-fill { height:100%; border-radius:5px; }
  .rack-card .rack-health-pct { font-size:13px; color:#666; font-weight:600; margin-bottom:10px; }
  .rack-card .rack-stats { font-size:12.5px; color:#555; display:flex; flex-direction:column; gap:4px; padding-top:10px; border-top:1px solid #f0f0f0; }
  .rack-card .rack-comment { font-size:12.5px; color:#8a6100; margin-top:10px; font-style:italic; }

  /* Action Plan - one card per item */
  .action-grid { display:flex; flex-direction:column; gap:12px; }
  .action-card { display:flex; gap:14px; padding:16px 18px; border-radius:8px; background:#fff; border:1px solid #eee; border-left:5px solid #c62828; box-shadow:0 1px 3px rgba(0,0,0,0.08); align-items:flex-start; }
  .action-card.med { border-left-color:#e6a100; }
  .action-card .sev-badge { flex-shrink:0; padding:5px 12px; border-radius:6px; font-size:12px; font-weight:bold; white-space:nowrap; background:#fdecea; color:#c62828; }
  .action-card .sev-badge.med { background:#fff6e0; color:#8a6100; }
  .action-card .ac-body { flex:1; min-width:0; }
  .action-card .ac-title { font-weight:bold; font-size:15px; color:#1a1a1a; margin-bottom:4px; }
  .action-card .ac-sub { font-size:13.5px; color:#666; }
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
      <div class="stat" style="background:#e6a100"><span class="num">$($healthCounts.Warning)</span><span class="label">Sites with Warnings</span></div>
      <div class="stat" style="background:#c62828"><span class="num">$($healthCounts.Critical)</span><span class="label">Sites Critical</span></div>
      <div class="stat" style="background:#4F46E5"><span class="num">$totalMedium</span><span class="label">Total Medium Risk Issues</span></div>
      <div class="stat" style="background:#1565C0"><span class="num">$totalIdfIssues</span><span class="label">Total IDF Issues Reported</span></div>
      <div class="stat" style="background:#00897B"><span class="num">$totalRacks</span><span class="label">Total Racks</span></div>
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
# 6. OUTPUT: DASHBOARD + LOG
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
