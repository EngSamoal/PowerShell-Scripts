#Requires -Version 5.1
<#
.SYNOPSIS
    Infrastructure Weekly Health Check - reads the IDF Room Report and DC RackWise Health
    Check workbooks for AquaArabia and SixFlags and produces a single combined HTML dashboard,
    in the same visual style as VMware_Weekly_HealthCheck.ps1's dashboard.

.DESCRIPTION
    Reads 4 source workbooks (2 per site):
      - <Site> IDF Room Report: one "Zone - N" tab per zone, each a per-switch/per-cabinet
        checklist (IDF Room clean, Rodent Trap, Air Conditioning, Cabinet doors, Switch Label,
        Fiber Uplink Labeling, Access Card Working) plus free-text Comments.
      - <Site> DC RackWise Health Check: a "Summary" tab with one row per rack (Health %,
        Status, Rack/Node check counts), and one tab per rack (A01, B01, P1, S1, ...) with a
        14-item rack-level checklist followed by a per-device ("Node-Wise"/"Device-Wise")
        checklist further down the same sheet.

    Every tab is located by searching for its known column headers rather than assuming a
    fixed row/column position, since the real workbooks have inconsistent header rows (a
    title row, blank spacer rows, or a second workbook's extra stray columns) - a tab whose
    headers can't be found is skipped and logged, never guessed at.

    Produces ONE combined HTML dashboard (no Word/PDF) covering both sites, matching the
    VMware dashboard's look: circular site tabs, KPI stat tiles, colored panels, an itemized
    Action Plan of every "No" answer (room/rack/device) with its Comments as context, and a
    Backup & Disaster Recovery panel driven by -BackupInfo exactly like the VMware script
    (this workbook set has no backup data of its own, so that panel reads "Manual/External
    Required" unless you pass it).

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
.PARAMETER BackupInfo
    Same shape as VMware_Weekly_HealthCheck.ps1's -BackupInfo: @{ 'AquaArabia' = @{
    Appliance=...; Schedule=...; Retention=...; Status=...; Notes=... } }.
#>

[CmdletBinding()]
param(
    [string]$AquaArabiaIdfPath = '',
    [string]$AquaArabiaRackPath = '',
    [string]$SixFlagsIdfPath = '',
    [string]$SixFlagsRackPath = '',

    [string]$OutputPath = (Join-Path $PSScriptRoot "Infrastructure_HealthCheck_Reports"),

    [hashtable]$BackupInfo = @{}
)

$ScriptBuild = '2026-10-01-01-initial'
Write-Host "Infrastructure_Weekly_HealthCheck.ps1 - build $ScriptBuild" -ForegroundColor Magenta

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
$Global:SiteRackMap = @{}   # Site label -> ordered list of rack names discovered (for per-site Clusters-style table).

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
    # Area: IDF | Rack | Node. Group: zone name (IDF) or rack ID (Rack/Node).
    # Status: Healthy / Warning / Critical / Information / Unable to Check / Manual/External Required
    $obj = [pscustomobject]@{
        Site = $Site; Area = $Area; Group = $Group; Object = $Object
        Item = $Item; Value = $Value; Status = $Status; Notes = $Notes
    }
    $Global:AllResults.Add($obj) | Out-Null
}

function Get-WorstStatus {
    param([string[]]$Statuses)
    $order = @{ 'Critical' = 5; 'Warning' = 4; 'Unable to Check' = 3; 'Manual/External Required' = 2; 'Healthy' = 1; 'Information' = 0 }
    if (-not $Statuses -or $Statuses.Count -eq 0) { return 'Healthy' }
    return ($Statuses | Sort-Object { $order[$_] } -Descending | Select-Object -First 1)
}

# Maps a plain Yes/No/N/A-style answer to a Status - used for every IDF/Rack/Node checklist item.
function Get-YesNoStatus {
    param([string]$Value)
    switch -Regex ($Value.Trim()) {
        '^Yes$'     { return 'Healthy' }
        '^No$'      { return 'Warning' }
        '^(N/A|Pending)$' { return 'Information' }
        default     { return 'Information' }
    }
}

# Collapses a multi-line cell (real Comments cells contain embedded newlines) into one compact
# line for display as a finding's Notes / Action Plan subtitle.
function ConvertTo-CompactNote {
    param([string]$Text)
    if ([string]::IsNullOrWhiteSpace($Text)) { return '' }
    return ($Text -replace '\r?\n', ' | ').Trim()
}

# ============================================================================
# 1. EXCEL HEADER-ANCHORING HELPERS
#    Every tab is located by searching for its known header text rather than assuming a fixed
#    row/column - the real workbooks have inconsistent header rows (title rows, blank spacer
#    rows, stray leftover columns from a different tab's data-validation list) that make a
#    fixed-position read silently wrong instead of simply failing.
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

# Finds the first cell (within MaxRow/MaxCol) whose text CONTAINS $Contains (case-insensitive) -
# used for the IDF checklist header, whose exact wording drifts slightly between tabs/sites
# ("IDF Room (Clean)" vs "IDF Room" vs "IDF Room(Clean)") but always starts the same 8-column
# checklist block in the same fixed order.
function Find-HeaderCellContains {
    param($Worksheet, [string]$Contains, [int]$MaxRow = 10, [int]$MaxCol = 40)
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
# 2. IDF ROOM REPORT COLLECTION
# ============================================================================
# Fixed, confirmed-consistent order of the 8 checklist columns starting at the "IDF Room"
# column - hardcoded rather than read from the (inconsistently worded) header text itself.
$IdfChecklistItems = @('IDF Room Clean','Rodent Trap','Air Conditioning','Cabinet - Front & Back Door','Switch Label','Fiber Uplink Labeling','Access Card Working')

function Read-IdfWorkbook {
    param($Excel, [string]$Site, [string]$Path)
    if (-not $Path -or -not (Test-Path $Path)) {
        if ($Path) { Write-CheckLog -Site $Site -Object $Path -CheckName 'IDF workbook' -ErrorMessage 'File not found.' }
        return
    }
    $wb = $Excel.Workbooks.Open($Path, 0, $true)
    try {
        foreach ($ws in $wb.Worksheets) {
            if ($ws.Name.Trim() -notmatch '^Zone') { continue }
            $zoneName = $ws.Name.Trim()
            Invoke-SafeCheck -Site $Site -ObjectName $zoneName -CheckName 'IDF zone tab' -Script {
                $idfCell = Find-HeaderCellContains -Worksheet $ws -Contains 'IDF Room'
                $idCell  = Find-HeaderRow -Worksheet $ws -Labels @('Zone','Cabinet No','Building','Switch - Hostname') -MaxRow 15
                if (-not $idfCell -or -not $idCell) {
                    throw "Could not locate the checklist header or the Zone/Cabinet/Building/Switch-Hostname header on this tab."
                }
                $checklistStartCol = $idfCell.Col
                $zoneCol   = $idCell.Columns['Zone']
                $cabCol    = $idCell.Columns['Cabinet No']
                $bldgCol   = $idCell.Columns['Building']
                $switchCol = $idCell.Columns['Switch - Hostname']
                $commentsCol = $checklistStartCol + 7   # 8th checklist column, after the 7 Yes/No items

                $used = $ws.UsedRange
                $lastRow = $used.Row + $used.Rows.Count - 1
                $dataStart = $idCell.Row + 1

                $currentCabinet = ''
                $currentBuilding = ''
                for ($r = $dataStart; $r -le $lastRow; $r++) {
                    $zoneVal = $ws.Cells.Item($r, $zoneCol).Value2
                    $cabVal  = $ws.Cells.Item($r, $cabCol).Value2
                    $bldgVal = $ws.Cells.Item($r, $bldgCol).Value2
                    $switchVal = $ws.Cells.Item($r, $switchCol).Value2
                    if ($cabVal -and "$cabVal".Trim() -ne '') { $currentCabinet = "$cabVal".Trim() }
                    if ($bldgVal -and "$bldgVal".Trim() -ne '') { $currentBuilding = "$bldgVal".Trim() }
                    if (-not $switchVal -or "$switchVal".Trim() -eq '') { continue }   # a pure spacer row
                    $objectLabel = "$($currentBuilding) / $($currentCabinet) / $("$switchVal".Trim())"

                    $commentsRaw = $ws.Cells.Item($r, $commentsCol).Value2
                    $commentsNote = ConvertTo-CompactNote "$commentsRaw"

                    for ($i = 0; $i -lt $IdfChecklistItems.Count; $i++) {
                        $cellVal = $ws.Cells.Item($r, $checklistStartCol + $i).Value2
                        if ($null -eq $cellVal -or "$cellVal".Trim() -eq '') { continue }
                        $status = Get-YesNoStatus "$cellVal"
                        $notes = if ($status -eq 'Warning') { $commentsNote } else { '' }
                        New-Finding -Site $Site -Area 'IDF' -Group $zoneName -Object $objectLabel `
                            -Item $IdfChecklistItems[$i] -Value "$cellVal".Trim() -Status $status -Notes $notes
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
# Fixed, confirmed-consistent column order of the per-device grid starting at the "RU" column.
$NodeGridItems = @('Expected Device','Mac Address / IP','Device Powered On','Device PDU-1','Device PDU-2','Any Disk Alarms','Device Labeling','Fiber/UTP Link Labeling','Temp OK','Link OK','Front/Back Led Alarms')
# Fixed, confirmed-consistent order of the 14-item rack-level checklist (Column A, Status in B).
$RackChecklistItems = @('Rack power OK','PDU -1  status OK','PDU -2  status OK','Devices powered ON','No critical alarms','Temperature within threshold','Airflow clear','Cable management OK','Fiber/UTP Uplinks OK','Fiber/UTP Labeling OK','Front and Back doors locked/secure','Rack Clean (No Foreign objects)','No physical damage')

function Read-RackSummaryTab {
    param($Worksheet, [string]$Site)
    $hdr = Find-HeaderRow -Worksheet $Worksheet -Labels @('Rack','Health %','Status') -MaxRow 20 -MaxCol 40
    if (-not $hdr) { throw "Could not locate the Rack/Health %/Status header on the Summary tab." }
    $cols = $hdr.Columns
    $rackCol = $cols['Rack']; $healthCol = $cols['Health %']; $statusCol = $cols['Status']
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
        $statusVal = $Worksheet.Cells.Item($r, $statusCol).Value2
        $rackYes = if ($yesCol) { $Worksheet.Cells.Item($r, $yesCol).Value2 } else { $null }
        $rackNo  = if ($noCol) { $Worksheet.Cells.Item($r, $noCol).Value2 } else { $null }
        $nodeYes = if ($nodeYesCol) { $Worksheet.Cells.Item($r, $nodeYesCol).Value2 } else { $null }
        $nodeNo  = if ($nodeNoCol) { $Worksheet.Cells.Item($r, $nodeNoCol).Value2 } else { $null }
        $comments = if ($commentsCol) { ConvertTo-CompactNote "$($Worksheet.Cells.Item($r, $commentsCol).Value2)" } else { '' }

        $status = switch ("$statusVal".Trim()) {
            'Healthy'   { 'Healthy' }
            'Attention' { 'Warning' }
            default     { 'Unable to Check' }
        }
        $healthPct = 0.0
        if ($healthVal) { try { $healthPct = [double]$healthVal * 100 } catch { } }

        New-Finding -Site $Site -Area 'Rack' -Group $rackName -Object $rackName -Item 'Rack Health' `
            -Value "$([Math]::Round($healthPct,1))|$rackYes|$rackNo|$nodeYes|$nodeNo" -Status $status -Notes $comments
        $racks += $rackName
    }
    return $racks
}

function Read-RackDetailTab {
    param($Worksheet, [string]$Site, [string]$RackName)
    # Rack-level checklist: Column A = Check Item (hardcoded list), Column B = Status, Column C = Comments.
    # Anchored by finding the literal "Check Item" / "Status" header rather than a fixed row.
    $hdr = Find-HeaderRow -Worksheet $Worksheet -Labels @('Check Item','Status') -MaxRow 15 -MaxCol 10
    if ($hdr) {
        $itemCol = $hdr.Columns['Check Item']; $statusCol = $hdr.Columns['Status']
        $commentsCol = $itemCol + 2
        for ($r = $hdr.Row + 1; $r -le ($hdr.Row + 20); $r++) {
            $itemVal = $Worksheet.Cells.Item($r, $itemCol).Value2
            if (-not $itemVal) { continue }
            $itemName = "$itemVal".Trim()
            if ($itemName -notin $RackChecklistItems) { continue }   # stop once we hit the node-grid section title/header
            $statusVal = $Worksheet.Cells.Item($r, $statusCol).Value2
            if (-not $statusVal -or "$statusVal".Trim() -eq '') { continue }
            $commentsVal = ConvertTo-CompactNote "$($Worksheet.Cells.Item($r, $commentsCol).Value2)"
            New-Finding -Site $Site -Area 'Rack' -Group $RackName -Object $RackName -Item $itemName `
                -Value "$statusVal".Trim() -Status (Get-YesNoStatus "$statusVal") -Notes $commentsVal
        }
    }

    # Node/device grid: anchored on the literal "RU" column header - column order after it is
    # fixed/confirmed, so read by position once that anchor is found.
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
            $cellVal = $Worksheet.Cells.Item($r, $ruCol + 2 + $i).Value2
            if ($null -eq $cellVal -or "$cellVal".Trim() -eq '' -or "$cellVal".Trim() -ieq 'N/A') { continue }
            $status = Get-YesNoStatus "$cellVal"
            $notes = if ($status -eq 'Warning') { $commentsVal } else { '' }
            New-Finding -Site $Site -Area 'Node' -Group $RackName -Object $deviceLabel `
                -Item $NodeGridItems[$i] -Value "$cellVal".Trim() -Status $status -Notes $notes
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
        if (-not $Global:SiteRackMap.ContainsKey($Site)) { $Global:SiteRackMap[$Site] = [System.Collections.Generic.List[string]]::new() }
        foreach ($rackName in $racks) {
            $Global:SiteRackMap[$Site].Add($rackName)
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
    Read-IdfWorkbook -Excel $Excel -Site $s.Site -Path $s.IdfPath
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
# IDF zone pass/fail rollup, per-rack health rollup, and the combined itemized Action Plan.
function Get-SiteDashboardSummary {
    param([string]$SiteLabel)
    $SiteFindings = $Global:AllResults | Where-Object { $_.Site -eq $SiteLabel }
    if (-not $SiteFindings) { return $null }

    $CritCount = @($SiteFindings | Where-Object { $_.Status -eq 'Critical' }).Count
    $WarnCount = @($SiteFindings | Where-Object { $_.Status -eq 'Warning' }).Count
    $OverallHealth = if ($CritCount -gt 0) { 'Critical' } elseif ($WarnCount -gt 0) { 'Warning' } else { 'Healthy' }

    # --- IDF rollup ---
    $idfFindings = @($SiteFindings | Where-Object { $_.Area -eq 'IDF' })
    $idfZones = @($idfFindings | Select-Object -ExpandProperty Group -Unique | Sort-Object)
    $idfZoneRows = @(foreach ($z in $idfZones) {
        $zf = @($idfFindings | Where-Object { $_.Group -eq $z })
        $zFail = @($zf | Where-Object { $_.Status -eq 'Warning' }).Count
        $zInfo = @($zf | Where-Object { $_.Status -eq 'Information' }).Count
        [pscustomobject]@{ Zone = $z; Total = $zf.Count; Fail = $zFail; Pass = $zf.Count - $zFail - $zInfo }
    })
    $idfSwitchCount = @($idfFindings | Select-Object -ExpandProperty Object -Unique).Count

    # --- Rack rollup ---
    $rackHealthFindings = @($SiteFindings | Where-Object { $_.Area -eq 'Rack' -and $_.Item -eq 'Rack Health' })
    $rackRowsUnsorted = @(foreach ($rf in $rackHealthFindings) {
        $p = $rf.Value -split '\|'
        [pscustomobject]@{
            Rack = $rf.Group; HealthPct = [double]$p[0]
            RackYes = $p[1]; RackNo = $p[2]; NodeYes = $p[3]; NodeNo = $p[4]
            Status = $rf.Status; Comments = $rf.Notes
        }
    })
    $rackRows = @($rackRowsUnsorted | Sort-Object Rack)
    $rackCount = $rackRows.Count
    $rackHealthyCount = @($rackRows | Where-Object { $_.Status -eq 'Healthy' }).Count
    $rackAttentionCount = @($rackRows | Where-Object { $_.Status -eq 'Warning' }).Count
    $avgHealthPct = if ($rackCount -gt 0) { ($rackRows | Measure-Object -Property HealthPct -Average).Average } else { $null }

    # --- Action plan: every Warning/Critical IDF/Rack-checklist/Node finding (the per-rack
    # Health rollup itself is excluded here since it's already shown in the Rack Health panel). ---
    $ActionItems = @($SiteFindings | Where-Object { $_.Status -in 'Critical','Warning' -and $_.Item -ne 'Rack Health' } |
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

    $backupSupplied = $BackupInfo.ContainsKey($SiteLabel)
    $backupDetail = if ($backupSupplied) { $BackupInfo[$SiteLabel] } else { $null }
    $backupStatus = if ($backupDetail -and $backupDetail.Status) { $backupDetail.Status } else { 'Manual/External Required' }
    if ($backupStatus -eq 'Critical') { $CritCount++; $OverallHealth = 'Critical' }
    elseif ($backupStatus -eq 'Warning' -and $OverallHealth -ne 'Critical') { $WarnCount++; $OverallHealth = 'Warning' }

    $SummaryText = if ($CritCount -gt 0) {
        "$SiteLabel has $CritCount critical and $WarnCount warning item(s) across $($idfZones.Count) IDF zones and $rackCount racks that require attention."
    } elseif ($WarnCount -gt 0) {
        "$SiteLabel is stable overall, with $WarnCount minor item(s) flagged across $($idfZones.Count) IDF zones and $rackCount racks - see the Action Plan below."
    } else {
        "$SiteLabel's $($idfZones.Count) IDF zones and $rackCount racks are all reporting healthy with no items flagged."
    }

    [pscustomobject]@{
        Site               = $SiteLabel
        OverallHealth      = $OverallHealth
        HighRisk           = $CritCount
        MediumRisk         = $WarnCount
        IdfZoneCount       = $idfZones.Count
        IdfSwitchCount     = $idfSwitchCount
        IdfZoneRows        = $idfZoneRows
        RackCount          = $rackCount
        RackHealthyCount   = $rackHealthyCount
        RackAttentionCount = $rackAttentionCount
        AvgHealthPct       = $avgHealthPct
        RackRows           = $rackRows
        ActionItems        = $ActionItems
        BackupSupplied     = $backupSupplied
        BackupDetail       = $backupDetail
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
    $totalHigh   = ($summaries | Measure-Object -Property HighRisk -Sum).Sum
    $totalMedium = ($summaries | Measure-Object -Property MediumRisk -Sum).Sum
    $totalZones  = ($summaries | Measure-Object -Property IdfZoneCount -Sum).Sum
    $totalRacks  = ($summaries | Measure-Object -Property RackCount -Sum).Sum

    $tabPalette = @('#2563EB','#7C3AED','#0D9488','#C026D3','#EA580C','#4F46E5','#DB2777','#0EA5E9')
    $tileBlue = '#1565C0'; $tilePurple = '#6A1B9A'; $tileTeal = '#00897B'; $tileIndigo = '#283593'; $tileBrown = '#6D4C41'

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
        $backupStatusText = if ($s.BackupSupplied -and $s.BackupDetail.Status -in 'Healthy','Warning','Critical') { $s.BackupDetail.Status } else { 'Manual/External Required' }
@"
      <div class="ov-card" onclick="showPage('$slug')" style="border-top-color:$color">
        <div class="ov-head"><h2>$(ConvertTo-HtmlSafe $s.Site)</h2><span class="badge" style="background:$color">$(ConvertTo-HtmlSafe $healthLabelText[$s.OverallHealth])</span></div>
        <div class="risk-row"><span class="risk risk-high">High: $($s.HighRisk)</span><span class="risk risk-med">Medium: $($s.MediumRisk)</span></div>
        <table class="metrics">
          <tr><td>IDF Zones / Switches</td><td>$($s.IdfZoneCount) / $($s.IdfSwitchCount)</td></tr>
          <tr><td>Racks (Healthy / Attention)</td><td>$($s.RackCount) ($($s.RackHealthyCount) / $($s.RackAttentionCount))</td></tr>
          <tr><td>Average Rack Health</td><td>$avgHealthText</td></tr>
          <tr><td>Backup Status</td><td>$(ConvertTo-HtmlSafe $backupStatusText)</td></tr>
        </table>
        <span class="ov-link">View Full Details &rarr;</span>
      </div>
"@
    }) -join "`n"

    $overviewTabButton = '<button class="tab overview-tab active" id="tab-overview" onclick="showPage(''overview'')"><span class="tab-circle">Overview</span></button>'
    $siteTabButtons = $(for ($i = 0; $i -lt $summaries.Count; $i++) {
        $s = $summaries[$i]
        $slug = ConvertTo-Slug $s.Site
        $tabColor = $tabPalette[$i % $tabPalette.Count]
        "<button class=`"tab`" id=`"tab-$slug`" onclick=`"showPage('$slug')`"><span class=`"tab-circle`" style=`"background:$tabColor`">$(ConvertTo-HtmlSafe $s.Site)</span></button>"
    }) -join "`n    "
    $tabButtons = @($overviewTabButton, $siteTabButtons) -join "`n    "

    # --- Per-site pages ---
    $sitePages = ($summaries | ForEach-Object {
        $s = $_
        $slug = ConvertTo-Slug $s.Site
        $color = $healthColor[$s.OverallHealth]

        $idfRows = if ($s.IdfZoneRows.Count -gt 0) {
            ($s.IdfZoneRows | ForEach-Object {
                $pct = if ($_.Total -gt 0) { ($_.Pass / $_.Total) * 100 } else { 0 }
@"
          <tr><td>$(ConvertTo-HtmlSafe $_.Zone)</td><td>$($_.Total)</td><td>$($_.Pass)</td><td>$($_.Fail)</td>
            <td><div class="mini-bar-wrap"><div class="mini-bar-track"><div class="mini-bar-fill" style="width:$pct%;background:$(Get-PctBarColor $pct)"></div></div><span>$("{0:N0}" -f $pct)%</span></div></td></tr>
"@
            }) -join "`n"
        } else { "<tr><td colspan='5' style='color:#999'>No IDF zone data found</td></tr>" }

        $rackRows = if ($s.RackRows.Count -gt 0) {
            ($s.RackRows | ForEach-Object {
@"
          <tr><td>$(ConvertTo-HtmlSafe $_.Rack)</td>
            <td><div class="mini-bar-wrap"><div class="mini-bar-track"><div class="mini-bar-fill" style="width:$($_.HealthPct)%;background:$(Get-PctBarColor $_.HealthPct)"></div></div><span>$("{0:N1}" -f $_.HealthPct)%</span></div></td>
            <td>$(Get-StatusPillHtml -Status $_.Status -Text $_.Status)</td></tr>
"@
            }) -join "`n"
        } else { "<tr><td colspan='3' style='color:#999'>No rack data found</td></tr>" }

        $actionRows = if ($s.ActionItems.Count -gt 0) {
            ($s.ActionItems | ForEach-Object {
                $sevClass = if ($_.Severity -eq 'High') { 'high' } else { 'med' }
                $subtitle = if ($_.Notes) { "$($_.Group) / $($_.Object) - $($_.Notes)" } else { "$($_.Group) / $($_.Object)" }
@"
        <li><span class="sev $sevClass">$($_.Severity)</span><div class="txt"><strong>$(ConvertTo-HtmlSafe $_.Item): $(ConvertTo-HtmlSafe $_.Value)</strong><span>$(ConvertTo-HtmlSafe $subtitle)</span></div></li>
"@
            }) -join "`n"
        } else { $null }
        $actionPlanBody = if ($actionRows) {
            "<ul class=`"action-list`">`n$actionRows`n</ul>"
        } else {
            "<div class=`"center-callout`">$(Get-StatusPillHtml -Status 'Healthy' -Text 'No Warning or Critical items flagged for this site')</div>"
        }

        $backupBody = if ($s.BackupSupplied -and $s.BackupDetail.Status -in 'Healthy','Warning','Critical') {
            $bi = $s.BackupDetail
            $biStatus = $bi.Status
            $statusText = switch ($biStatus) {
                'Healthy'  { 'Successfully Completed' }
                'Critical' { if ($bi.Notes) { $bi.Notes } else { 'Failed' } }
                'Warning'  { if ($bi.Notes) { $bi.Notes } else { 'Attention Required' } }
            }
            $flagClass = switch ($biStatus) { 'Critical' { 'critical' }; 'Warning' { 'warn' }; default { 'healthy' } }
@"
      <table class="metrics">
        <tr><td>Appliance</td><td>$(ConvertTo-HtmlSafe $bi.Appliance) ($(ConvertTo-HtmlSafe $biStatus))</td></tr>
        <tr><td>Schedule</td><td>$(ConvertTo-HtmlSafe $bi.Schedule)</td></tr>
        <tr><td>Retention</td><td style="white-space:nowrap">$(ConvertTo-HtmlSafe $bi.Retention)</td></tr>
        <tr><td>Status</td><td>$(ConvertTo-HtmlSafe $statusText)</td></tr>
      </table>
      <div class="backup-flag $flagClass">$(ConvertTo-HtmlSafe $biStatus)</div>
"@
        } elseif ($s.BackupSupplied) {
            $bi = $s.BackupDetail
            $reasonText = if ($bi.Notes) { $bi.Notes } else { 'Manual/External Required' }
@"
      <div class="center-callout">
        $(Get-StatusPillHtml -Status 'Manual/External Required' -Text 'Manual/External Required')
        <p style="color:#999;font-size:13px;margin:12px 0 0">$(ConvertTo-HtmlSafe $reasonText)</p>
      </div>
      <div class="backup-flag warn">Pending</div>
"@
        } else {
@"
      <div class="center-callout">
        $(Get-StatusPillHtml -Status 'Manual/External Required' -Text 'Manual/External Required')
        <p style="color:#999;font-size:13px;margin:12px 0 0">Not supplied for this run - pass -BackupInfo to include backup device/solution status here.</p>
      </div>
"@
        }

        $avgHealthValueText = if ($s.AvgHealthPct -ne $null) { "{0:N1}%" -f $s.AvgHealthPct } else { 'n/a' }

@"
      <section class="page" id="page-$slug">
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
          <div class="tile" style="background:$tileBlue"><span class="num">$($s.IdfZoneCount)</span><span class="label">IDF Zones</span></div>
          <div class="tile" style="background:$tilePurple"><span class="num">$($s.IdfSwitchCount)</span><span class="label">Switches / Cabinets</span></div>
          <div class="tile" style="background:$tileTeal"><span class="num">$($s.RackCount)</span><span class="label">Racks</span></div>
          <div class="tile" style="background:#2e7d32"><span class="num">$($s.RackHealthyCount)</span><span class="label">Racks Healthy</span></div>
          <div class="tile" style="background:#c62828"><span class="num">$($s.RackAttentionCount)</span><span class="label">Racks Attention</span></div>
          <div class="tile" style="background:$tileIndigo"><span class="num" style="font-size:18px">$avgHealthValueText</span><span class="label">Avg Rack Health</span></div>
        </div>

        <div class="panel-grid">
          <div class="panel panel-full" style="border-top-color:$tilePurple">
            <h3><span class="n" style="background:$tilePurple">&#127968;</span>IDF Room Health <span style="font-weight:normal;font-size:14px;color:#999">($($s.IdfZoneRows.Count) zones)</span></h3>
            <div class="scroll-box">
              <table class="metrics wide">
                <thead><tr><th>Zone</th><th>Total Checks</th><th>Pass</th><th>Fail</th><th>Pass Rate</th></tr></thead>
                <tbody>
$idfRows
                </tbody>
              </table>
            </div>
          </div>

          <div class="panel panel-full" style="border-top-color:$tileTeal">
            <h3><span class="n" style="background:$tileTeal">&#128451;&#65039;</span>Rack Health <span style="font-weight:normal;font-size:14px;color:#999">($($s.RackCount) racks)</span></h3>
            <div class="scroll-box">
              <table class="metrics wide">
                <thead><tr><th>Rack</th><th>Health %</th><th>Status</th></tr></thead>
                <tbody>
$rackRows
                </tbody>
              </table>
            </div>
          </div>

          <div class="panel" style="border-top-color:$tileBrown">
            <h3><span class="n" style="background:$tileBrown">&#9729;&#65039;</span>Backup &amp; Disaster Recovery</h3>
$backupBody
          </div>

          <div class="panel" style="border-top-color:#1E3A5F">
            <h3><span class="n" style="background:#1E3A5F">&#128203;</span>Summary</h3>
            <p class="summary-text">$(ConvertTo-HtmlSafe $s.SummaryText)</p>
          </div>
        </div>

        <div class="panel" style="border-top-color:#e6a100">
          <h3><span class="n" style="background:#e6a100">&#9888;&#65039;</span>Action Plan - Items Requiring Attention</h3>
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
<title>Infrastructure Weekly Health Check - Dashboard</title>
<style>
  body { font-family: Calibri, Arial, sans-serif; background:#eef1f5; color:#1a1a1a; margin:0; padding:0 32px 32px; font-size:16px; line-height:1.4; }
  h1 { color:#1E3A5F; margin:0; font-size:36px; }
  .subtitle { color:#555; margin:6px 0 20px; font-size:16px; }
  .tabs { display:grid; grid-template-columns: repeat(auto-fit, minmax(170px, 1fr)); gap:20px; padding:24px 0 20px; position:sticky; top:0; background:#eef1f5; z-index:10; border-bottom:1px solid #dfe3e8; margin-bottom:32px; }
  .tab { display:flex; flex-direction:column; align-items:center; cursor:pointer; background:none; border:none; padding:0; font-family:inherit; }
  .tab .tab-circle { width:150px; height:150px; border-radius:50%; display:flex; align-items:center; justify-content:center; color:#fff; font-weight:bold; font-size:17px; text-align:center; line-height:1.25; padding:10px; box-sizing:border-box; box-shadow:0 1px 4px rgba(0,0,0,0.2); border:3px solid transparent; }
  .tab.overview-tab .tab-circle { background:#1E3A5F; }
  .tab.active .tab-circle { border-color:#1a1a1a; box-shadow:0 0 0 3px rgba(0,0,0,0.15), 0 1px 4px rgba(0,0,0,0.2); }
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
  .risk-low  { background:#eef2f5; color:#555; }
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
  .scroll-box { max-height:340px; overflow-y:auto; border:1px solid #f0f0f0; border-radius:6px; }
  .scroll-box table.metrics.wide { font-size:15px; }
  .scroll-box thead th { position:sticky; top:0; background:#fff; }
  .panel-full { grid-column: 1 / -1; }
  .mini-bar-wrap { display:flex; align-items:center; gap:8px; justify-content:center; }
  .mini-bar-track { width:60px; height:8px; background:#eef1f4; border-radius:4px; overflow:hidden; flex-shrink:0; }
  .mini-bar-fill { height:100%; border-radius:4px; }
  .action-list { list-style:none; margin:0; padding:0; }
  .action-list li { display:flex; gap:14px; padding:14px 0; border-bottom:1px solid #eee; align-items:flex-start; }
  .action-list li:last-child { border-bottom:none; }
  .action-list .sev { flex-shrink:0; padding:4px 12px; border-radius:6px; font-size:12px; font-weight:bold; white-space:nowrap; margin-top:2px; }
  .action-list .sev.high { background:#fdecea; color:#c62828; }
  .action-list .sev.med { background:#fff6e0; color:#8a6100; }
  .action-list .txt strong { display:block; font-size:15px; color:#1a1a1a; }
  .action-list .txt span { font-size:14px; color:#666; }
  .summary-text { color:#444; font-size:15px; line-height:1.6; }
  .center-callout { text-align:center; padding:8px 0 18px; }
  .backup-flag { margin:18px -28px -26px -28px; padding:10px 0; text-align:center; font-weight:bold; color:#fff; border-radius:0 0 10px 10px; font-size:14px; letter-spacing:0.5px; }
  .backup-flag.healthy { background:#2e7d32; }
  .backup-flag.critical { background:#c62828; }
  .backup-flag.warn { background:#e6a100; }
</style>
</head>
<body>
  <h1>Infrastructure Weekly Health Check - Dashboard</h1>
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
      <div class="stat" style="background:#1565C0"><span class="num">$totalZones</span><span class="label">Total IDF Zones</span></div>
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

    $dashboardPath = Join-Path $OutputPath "Infrastructure_HealthCheck_Dashboard_$RunDate.html"
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

$LogPath = Join-Path $OutputPath "Infrastructure_Weekly_HealthCheck_$RunDate.log"
$LogLines = @()
$LogLines += "Infrastructure Weekly Health Check run - $($ScriptStart.ToString('u')) - build $ScriptBuild"
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
