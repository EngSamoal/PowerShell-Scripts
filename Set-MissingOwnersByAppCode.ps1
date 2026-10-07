<#
.SYNOPSIS
    Fills missing Business Owner / Technical Owner values in an Excel VM inventory using the owners
    of other VMs that belong to the same application (2nd hyphen-separated part of the VM name).

.DESCRIPTION
    Application code = 2nd hyphen-separated part of "VM Name", compared case-insensitively.
        DX-BC-AP02  ->  BC

    For every application and for each owner field INDEPENDENTLY:
      * Exactly one distinct non-blank owner value among its VMs -> blank cells of that field are filled
        with it ("John Smith" and " john  smith" count as the same value).
      * Two or more distinct values -> CONFLICT. Nothing is filled for that field of that application;
        all values and the VMs holding them are listed for manual review.
      * No value on any VM of the application -> cannot be determined; nothing is filled.
    Existing (non-blank) owner cells are never changed. Rows with a blank or invalid VM name are skipped
    and reported; they are not used as evidence either.

    SAFETY
      * Dry-run (report only) is the DEFAULT. Nothing is written to Excel unless -Apply is given.
      * Only the 3 columns VM Name / Business Owner / Technical Owner are read; only EMPTY owner cells are
        written. All other columns, sheets, colours and formats are left exactly as they are.
      * -Apply UPDATES THE ORIGINAL FILE (-InputPath), safely:
          1. A backup copy is made first (<input folder>\Backup\<name>_backup_<timestamp>.xlsx, SHA256 verified).
          2. The changes are made and saved in a temporary copy - never directly in the original.
          3. The saved copy is re-opened and EVERY cell of EVERY sheet is compared with the original: the
             value and the format (colour/font/border/number format) must be identical, except the filled
             owner cells (new value, same format). Any difference -> the original is NOT touched.
          4. Only then the temporary copy replaces the original (if the file is open in Excel this fails and
             the original stays unchanged - close Excel and run again).
      * Dry-run never writes the workbook; the original's SHA256 hash is checked at the end to prove it.

    REPORTS (always written, dry-run and apply; folder = -ReportFolder):
      <name>_<mode>_<timestamp>_Changes.csv       every proposed/applied change
      <name>_<mode>_<timestamp>_Unresolved.csv    every missing owner that could NOT be filled, with the reason
      <name>_<mode>_<timestamp>_Conflicts.csv     every conflicting value per application/field and its VMs
      <name>_<mode>_<timestamp>_InvalidVMNames.csv
      <name>_<mode>_<timestamp>_Summary.txt       the console summary

    Requires the ImportExcel module (does NOT need Microsoft Excel). The script does not install anything.

.PARAMETER Apply
    Update the original workbook (after a backup). Without it the script only reports what it would do.

.EXAMPLE
    # 1) Dry run (default) - nothing is written to Excel, only the reports:
    .\Set-MissingOwnersByAppCode.ps1

.EXAMPLE
    # 2) Apply - backup, then the missing owners are filled in C:\temp\master_draft.xlsx itself:
    .\Set-MissingOwnersByAppCode.ps1 -Apply

.EXAMPLE
    # 3) Another file (dry run first, then apply):
    .\Set-MissingOwnersByAppCode.ps1 -InputPath 'D:\inv\master.xlsx'
    .\Set-MissingOwnersByAppCode.ps1 -InputPath 'D:\inv\master.xlsx' -Apply

.NOTES
    Script build: see $ScriptBuild below (printed as the first line of output).
    Works on Windows PowerShell 5.1 and PowerShell 7+. Only .xlsx / .xlsm are supported (not legacy .xls).
#>
[CmdletBinding()]
param(
    [string]$InputPath  = 'C:\temp\master_draft.xlsx',
    [string]$WorksheetName,                                   # default: first sheet that has all 3 headers
    [string]$ReportFolder,                                    # default: <input folder>\OwnerFill_Reports
    [string]$BackupFolder,                                    # default: <input folder>\Backup
    [string]$VMNameHeader         = 'VM Name',
    [string]$BusinessOwnerHeader  = 'Business Owner',
    [string]$TechnicalOwnerHeader = 'Technical Owner',
    [ValidateRange(1, 1000)][int]$HeaderSearchRows = 25,      # header row is searched in the first N rows
    [switch]$Apply
)

$ErrorActionPreference = 'Stop'
$ScriptBuild = 'Set-MissingOwnersByAppCode build 2026-10-07.5'
Write-Host $ScriptBuild -ForegroundColor Cyan

# ------------------------------------------------------------------------------------------------
# Helpers
# ------------------------------------------------------------------------------------------------
function ConvertTo-CleanText($Value) {
    # Cell value -> trimmed string with collapsed whitespace (incl. non-breaking spaces). $null -> ''.
    if ($null -eq $Value) { return '' }
    if ($Value -is [double] -and [double]::IsNaN($Value)) { return '' }   # formula never calculated/saved by Excel
    return (([string]$Value) -replace '[\s\u00A0]+', ' ').Trim()
}

function Get-CompareKey([string]$Text) {
    # Key used to decide whether two owner values are "the same person" (case/whitespace-insensitive).
    return (ConvertTo-CleanText $Text).ToUpperInvariant()
}

function Get-AppCode([string]$VMName) {
    # Returns @{ AppCode = 'BC'; Reason = $null } or @{ AppCode = $null; Reason = '<why invalid>' }
    $name = ConvertTo-CleanText $VMName
    if (-not $name) { return @{ AppCode = $null; Reason = 'Blank VM name' } }
    $parts = @($name -split '-')
    if ($parts.Count -lt 2) { return @{ AppCode = $null; Reason = "No '-' separator (need at least 2 parts, e.g. DX-BC-AP02)" } }
    if (-not $parts[0].Trim()) { return @{ AppCode = $null; Reason = 'First part of the name is empty' } }
    $app = $parts[1].Trim()
    if (-not $app) { return @{ AppCode = $null; Reason = 'Application part (2nd part) is empty' } }
    return @{ AppCode = $app.ToUpperInvariant(); Reason = $null }
}

function Find-HeaderLayout($Worksheet, [int]$MaxRows, [string[]]$Headers) {
    # Returns @{ Row = n; Columns = @{ header -> column } } for the first row containing all $Headers, else $null.
    if ($null -eq $Worksheet.Dimension) { return $null }
    $lastRow = [Math]::Min($MaxRows, $Worksheet.Dimension.End.Row)
    $lastCol = $Worksheet.Dimension.End.Column
    for ($r = 1; $r -le $lastRow; $r++) {
        $found = @{}
        for ($c = 1; $c -le $lastCol; $c++) {
            $key = Get-CompareKey $Worksheet.Cells[$r, $c].Value
            if (-not $key) { continue }
            foreach ($h in $Headers) {
                if ($key -eq (Get-CompareKey $h)) {
                    if ($found.ContainsKey($h)) {
                        throw "Worksheet '$($Worksheet.Name)' row ${r}: header '$h' appears more than once (columns $($found[$h]) and $c). Ambiguous - fix the header or rename one column."
                    }
                    $found[$h] = $c
                }
            }
        }
        if ($found.Count -eq $Headers.Count) { return @{ Row = $r; Columns = $found } }
    }
    return $null
}

function Get-ColorSignature($c) {
    if ($null -eq $c) { return '-' }
    return '{0}/{1}/{2}/{3}/{4}' -f $c.Rgb, $c.Theme, $c.Tint, $c.Indexed, $c.Auto
}

function Get-XfSignature($Xf) {
    # The ACTUAL format of a cell style (font, fill colour, borders, number format, alignment).
    # Compared instead of the StyleID number, because saving renumbers the style list (all IDs shift)
    # while the formats themselves stay the same.
    $f = $Xf.Font; $fl = $Xf.Fill; $b = $Xf.Border
    $parts = @(
        "font=$($f.Name)|$($f.Size)|$($f.Bold)|$($f.Italic)|$($f.UnderLineType)|$($f.Strike)|$($f.VerticalAlign)|$(Get-ColorSignature $f.Color)"
        "fill=$($fl.PatternType)|$(Get-ColorSignature $fl.BackgroundColor)|$(Get-ColorSignature $fl.PatternColor)"
        "border=$(foreach ($side in 'Left', 'Right', 'Top', 'Bottom', 'Diagonal') { '{0}:{1}:{2};' -f $side, $b.$side.Style, (Get-ColorSignature $b.$side.Color) })$($b.DiagonalUp)|$($b.DiagonalDown)"
        "numfmt=$($Xf.Numberformat.Format)"
        "align=$($Xf.HorizontalAlignment)|$($Xf.VerticalAlignment)|$($Xf.WrapText)|$($Xf.Indent)|$($Xf.TextRotation)|$($Xf.ShrinkToFit)"
        "protect=$($Xf.Locked)|$($Xf.Hidden)"
    )
    return ($parts -join ' ')
}

function Get-WorkbookSnapshot([string]$Path, [string]$TargetSheet, $TargetCells) {
    # Every existing cell of every sheet -> "value<TAB>formula<TAB>format", so equal snapshots = same data
    # AND same formatting (colours, fonts, borders, number formats, alignment).
    $snap = @{}
    $p = Open-ExcelPackage -Path $Path
    try {
        $xfCache = @{}      # StyleID -> format signature (computed once per style)
        $snap['#defaultFormat'] = Get-XfSignature $p.Workbook.Styles.CellXfs[0]
        foreach ($sheet in $p.Workbook.Worksheets) {
            $snap["#sheet|$($sheet.Name)"] = "$($sheet.Index)"
            if ($null -eq $sheet.Dimension) { continue }
            foreach ($cell in $sheet.Cells[$sheet.Dimension.Address]) {
                $sid = $cell.StyleID
                if (-not $xfCache.ContainsKey($sid)) { $xfCache[$sid] = Get-XfSignature $p.Workbook.Styles.CellXfs[$sid] }
                $snap["$($sheet.Name)|$($cell.Start.Row)|$($cell.Start.Column)"] = "{0}`t{1}`t{2}" -f [string]$cell.Value, $cell.Formula, $xfCache[$sid]
            }
        }
        # Cells to be filled may not exist in the file yet (empty, formatted only by row/column) - record
        # the format Excel shows for them, so the check can prove it is the same after the fill.
        $tws = $p.Workbook.Worksheets[$TargetSheet]
        foreach ($tc in @($TargetCells)) {
            $cell = $tws.Cells[$tc.Row, $tc.Column]
            $sid = $cell.StyleID
            if (-not $xfCache.ContainsKey($sid)) { $xfCache[$sid] = Get-XfSignature $p.Workbook.Styles.CellXfs[$sid] }
            $snap["$TargetSheet|$($tc.Row)|$($tc.Column)"] = "{0}`t{1}`t{2}" -f [string]$cell.Value, $cell.Formula, $xfCache[$sid]
        }
    }
    finally { Close-ExcelPackage $p -NoSave }
    return $snap
}

function ConvertTo-ColumnLetter([int]$Column) {
    $s = ''
    while ($Column -gt 0) { $m = ($Column - 1) % 26; $s = [char](65 + $m) + $s; $Column = [int][Math]::Floor(($Column - 1) / 26) }
    return $s
}

function ConvertFrom-CellRef([string]$Ref) {
    # 'AB12' -> @{ Row = 12; Column = 28 }
    if ($Ref -notmatch '^([A-Za-z]+)(\d+)$') { throw "Unexpected cell reference '$Ref' in the sheet XML." }
    $col = 0; foreach ($ch in $Matches[1].ToUpperInvariant().ToCharArray()) { $col = $col * 26 + ([int]$ch - 64) }
    return @{ Row = [int]$Matches[2]; Column = $col }
}

function Get-XlsxCellStyles([string]$Path, [string]$SheetName, $Cells) {
    # Style index Excel uses to DISPLAY each cell, read straight from the worksheet XML:
    # the cell's own s="" if the cell exists, else the row style (customFormat) else the column style, else 0.
    # Returns @{ 'row|col' = 'styleIndex' }.
    Add-Type -AssemblyName System.IO.Compression, System.IO.Compression.FileSystem
    $zip = [System.IO.Compression.ZipFile]::OpenRead($Path)
    try {
        $info = Open-XlsxSheetXml $zip $SheetName
        $doc = $info.Doc; $nsm = $info.Nsm
        $colStyle = @{}
        foreach ($col in $doc.SelectNodes('/m:worksheet/m:cols/m:col', $nsm)) {
            if ($col.HasAttribute('style')) { for ($i = [int]$col.GetAttribute('min'); $i -le [int]$col.GetAttribute('max'); $i++) { $colStyle[$i] = $col.GetAttribute('style') } }
        }
        $result = @{}
        foreach ($c in $Cells) {
            $ref = (ConvertTo-ColumnLetter $c.Column) + $c.Row
            $style = $null
            $rowNode = $doc.SelectSingleNode("/m:worksheet/m:sheetData/m:row[@r='$($c.Row)']", $nsm)
            if ($rowNode) {
                $cellNode = $rowNode.SelectSingleNode("m:c[@r='$ref']", $nsm)
                if ($cellNode) { $style = if ($cellNode.HasAttribute('s')) { $cellNode.GetAttribute('s') } else { '0' } }
                elseif ($rowNode.GetAttribute('customFormat') -in @('1', 'true') -and $rowNode.HasAttribute('s')) { $style = $rowNode.GetAttribute('s') }
            }
            if ($null -eq $style) { $style = if ($colStyle.ContainsKey([int]$c.Column)) { $colStyle[[int]$c.Column] } else { '0' } }
            $result["$($c.Row)|$($c.Column)"] = $style
        }
        return $result
    }
    finally { $zip.Dispose() }
}

function Open-XlsxSheetXml($Zip, [string]$SheetName) {
    # Finds the worksheet part for $SheetName and loads it. Returns @{ Doc; Nsm; Entry; Path }.
    $rNs = 'http://schemas.openxmlformats.org/officeDocument/2006/relationships'
    $wbEntry = $Zip.GetEntry('xl/workbook.xml')
    $relEntry = $Zip.GetEntry('xl/_rels/workbook.xml.rels')
    if (-not $wbEntry -or -not $relEntry) { throw 'Workbook structure not recognised (xl/workbook.xml missing).' }
    $wb = Read-ZipXml $wbEntry
    $sheetNode = @($wb.GetElementsByTagName('sheet', $wb.DocumentElement.NamespaceURI) | Where-Object { $_.GetAttribute('name') -eq $SheetName })[0]
    if (-not $sheetNode) { throw "Sheet '$SheetName' not found in xl/workbook.xml." }
    $rid = $sheetNode.GetAttribute('id', $rNs)
    if (-not $rid) { $rid = $sheetNode.GetAttribute('id', 'http://purl.oclc.org/ooxml/officeDocument/relationships') }
    $rels = Read-ZipXml $relEntry
    $rel = @($rels.DocumentElement.ChildNodes | Where-Object { $_.LocalName -eq 'Relationship' -and $_.GetAttribute('Id') -eq $rid })[0]
    if (-not $rel) { throw "Relationship '$rid' for sheet '$SheetName' not found." }
    $target = $rel.GetAttribute('Target')
    $sheetPath = if ($target.StartsWith('/')) { $target.TrimStart('/') } else { 'xl/' + $target }
    $sheetEntry = $Zip.GetEntry($sheetPath)
    if (-not $sheetEntry) { throw "Worksheet part '$sheetPath' not found in the file." }
    $doc = Read-ZipXml $sheetEntry
    $nsm = New-Object System.Xml.XmlNamespaceManager($doc.NameTable)
    $nsm.AddNamespace('m', $doc.DocumentElement.NamespaceURI)
    return @{ Doc = $doc; Nsm = $nsm; Entry = $sheetEntry; Path = $sheetPath }
}

function Read-ZipXml($Entry) {
    $doc = New-Object System.Xml.XmlDocument
    $doc.PreserveWhitespace = $true
    $st = $Entry.Open(); try { $doc.Load($st) } finally { $st.Dispose() }
    return $doc
}

function Set-XlsxCellText([string]$Path, [string]$SheetName, $Cells) {
    # Writes text into specific cells by editing ONLY that worksheet's XML inside the .xlsx (zip).
    # Every other part of the file (other sheets, styles, formulas, charts, ...) stays byte-for-byte identical.
    # Each cell keeps its style; a cell that does not exist yet gets the row/column style Excel would use.
    # $Cells: objects with Row, Column, Value.
    Add-Type -AssemblyName System.IO.Compression, System.IO.Compression.FileSystem
    $zip = [System.IO.Compression.ZipFile]::Open($Path, [System.IO.Compression.ZipArchiveMode]::Update)
    try {
        $info = Open-XlsxSheetXml $zip $SheetName
        $sheetEntry = $info.Entry; $sheetPath = $info.Path
        $doc = $info.Doc
        $ns = $doc.DocumentElement.NamespaceURI
        $pfx = $doc.DocumentElement.Prefix
        $nsm = New-Object System.Xml.XmlNamespaceManager($doc.NameTable)
        $nsm.AddNamespace('m', $ns)
        $sheetData = $doc.SelectSingleNode('/m:worksheet/m:sheetData', $nsm)
        if (-not $sheetData) { throw "No <sheetData> in '$sheetPath'." }

        # Column styles (<cols><col min max style>) for cells that do not exist yet.
        $colStyle = @{}
        foreach ($col in $doc.SelectNodes('/m:worksheet/m:cols/m:col', $nsm)) {
            if ($col.HasAttribute('style')) {
                for ($i = [int]$col.GetAttribute('min'); $i -le [int]$col.GetAttribute('max'); $i++) { $colStyle[$i] = $col.GetAttribute('style') }
            }
        }
        $rowsByNum = @{}
        foreach ($rowNode in $sheetData.SelectNodes('m:row', $nsm)) {
            if (-not $rowNode.HasAttribute('r')) { throw "A row in '$sheetPath' has no row number - unsupported file layout, nothing was changed." }
            $rowsByNum[[int]$rowNode.GetAttribute('r')] = $rowNode
        }

        foreach ($c in $Cells) {
            $ref = (ConvertTo-ColumnLetter $c.Column) + $c.Row
            $rowNode = $rowsByNum[[int]$c.Row]
            if (-not $rowNode) { throw "Row $($c.Row) not found in '$sheetPath' (it has a VM name, so it must exist) - nothing was changed." }
            $cellNode = $null; $insertBefore = $null
            foreach ($cn in $rowNode.SelectNodes('m:c', $nsm)) {
                if (-not $cn.HasAttribute('r')) { throw "A cell in row $($c.Row) has no reference - unsupported file layout, nothing was changed." }
                $pos = ConvertFrom-CellRef $cn.GetAttribute('r')
                if ($pos.Column -eq $c.Column) { $cellNode = $cn; break }
                if ($pos.Column -gt $c.Column) { $insertBefore = $cn; break }
            }
            if ($cellNode) {
                if ($cellNode.SelectSingleNode('m:f', $nsm)) { throw "Cell $ref contains a formula - refusing to overwrite it." }
                foreach ($child in @($cellNode.ChildNodes)) { [void]$cellNode.RemoveChild($child) }   # old blank <v>/<is>
            }
            else {
                $cellNode = $doc.CreateElement($pfx, 'c', $ns)
                $cellNode.SetAttribute('r', $ref)
                if ($rowNode.GetAttribute('customFormat') -in @('1', 'true') -and $rowNode.HasAttribute('s')) { $cellNode.SetAttribute('s', $rowNode.GetAttribute('s')) }
                elseif ($colStyle.ContainsKey([int]$c.Column)) { $cellNode.SetAttribute('s', $colStyle[[int]$c.Column]) }
                if ($insertBefore) { [void]$rowNode.InsertBefore($cellNode, $insertBefore) } else { [void]$rowNode.AppendChild($cellNode) }
            }
            $cellNode.SetAttribute('t', 'inlineStr')
            $is = $doc.CreateElement($pfx, 'is', $ns)
            $t = $doc.CreateElement($pfx, 't', $ns)
            $t.InnerText = [string]$c.Value
            $sp = $doc.CreateAttribute('xml', 'space', 'http://www.w3.org/XML/1998/namespace'); $sp.Value = 'preserve'
            [void]$t.Attributes.Append($sp)
            [void]$is.AppendChild($t)
            [void]$cellNode.AppendChild($is)
        }

        # Write the worksheet XML back into the same zip entry (UTF-8 without BOM, no re-indentation).
        $ms = New-Object System.IO.MemoryStream
        $xws = New-Object System.Xml.XmlWriterSettings
        $xws.Encoding = New-Object System.Text.UTF8Encoding($false)
        $xw = [System.Xml.XmlWriter]::Create($ms, $xws)
        $doc.Save($xw); $xw.Dispose()
        $bytes = $ms.ToArray()
        $st = $sheetEntry.Open()
        try { $st.SetLength(0); $st.Write($bytes, 0, $bytes.Length) } finally { $st.Dispose() }
        return $sheetPath
    }
    finally { $zip.Dispose() }
}

function Get-ZipPartHashes([string]$Path) {
    # zip entry name -> SHA256 of its content
    Add-Type -AssemblyName System.IO.Compression, System.IO.Compression.FileSystem
    $h = @{}
    $sha = [System.Security.Cryptography.SHA256]::Create()
    $zip = [System.IO.Compression.ZipFile]::OpenRead($Path)
    try {
        foreach ($e in $zip.Entries) {
            $st = $e.Open(); try { $h[$e.FullName] = [BitConverter]::ToString($sha.ComputeHash($st)) } finally { $st.Dispose() }
        }
    }
    finally { $zip.Dispose(); $sha.Dispose() }
    return $h
}

function Copy-FileVerified([string]$Source, [string]$Destination) {
    # Copy and prove the copy is byte-identical (SHA256). Throws on mismatch.
    Copy-Item -LiteralPath $Source -Destination $Destination -Force
    $a = (Get-FileHash -LiteralPath $Source -Algorithm SHA256).Hash
    $b = (Get-FileHash -LiteralPath $Destination -Algorithm SHA256).Hash
    if ($a -ne $b) { throw "Copy verification failed: '$Destination' does not match '$Source'." }
}

# ------------------------------------------------------------------------------------------------
# Pre-flight checks
# ------------------------------------------------------------------------------------------------
$mode  = if ($Apply) { 'APPLY' } else { 'DRYRUN' }
$stamp = Get-Date -Format 'yyyyMMdd_HHmmss'
$pkg = $null
$tempCopy = $null
$exitCode = 0

try {
    if (-not (Get-Module -ListAvailable -Name ImportExcel)) {
        Write-Host ''
        Write-Host 'The ImportExcel PowerShell module is not installed. Nothing was changed.' -ForegroundColor Red
        Write-Host 'Install it (current user, no admin needed) and run the script again:' -ForegroundColor Yellow
        Write-Host '    Install-Module -Name ImportExcel -Scope CurrentUser' -ForegroundColor White
        Write-Host 'Offline server: on a PC with internet run  Save-Module -Name ImportExcel -Path C:\temp\Modules' -ForegroundColor Gray
        Write-Host 'then copy C:\temp\Modules\ImportExcel to  Documents\WindowsPowerShell\Modules  on this server.' -ForegroundColor Gray
        exit 2
    }
    Import-Module ImportExcel -ErrorAction Stop -WarningAction SilentlyContinue

    if (-not (Test-Path -LiteralPath $InputPath -PathType Leaf)) { throw "Input file not found: $InputPath" }
    $inFull  = (Resolve-Path -LiteralPath $InputPath).ProviderPath
    $ext = [IO.Path]::GetExtension($inFull).ToLowerInvariant()
    if ($ext -notin @('.xlsx', '.xlsm')) { throw "Unsupported file type '$ext'. Only .xlsx / .xlsm are supported (save .xls as .xlsx first)." }

    $inDir = Split-Path -Parent $inFull
    $baseName = [IO.Path]::GetFileNameWithoutExtension($inFull)
    if (-not $ReportFolder) { $ReportFolder = Join-Path $inDir 'OwnerFill_Reports' }
    if (-not $BackupFolder) { $BackupFolder = Join-Path $inDir 'Backup' }
    if (-not (Test-Path -LiteralPath $ReportFolder)) { New-Item -ItemType Directory -Path $ReportFolder -Force | Out-Null }
    $reportPrefix = Join-Path $ReportFolder ('{0}_{1}_{2}' -f $baseName, $mode, $stamp)

    $originalHash = (Get-FileHash -LiteralPath $inFull -Algorithm SHA256).Hash
    Write-Host ("Mode        : {0}" -f $(if ($Apply) { 'APPLY (the original file will be updated after a backup)' } else { 'DRY RUN (report only - no workbook is written)' })) -ForegroundColor $(if ($Apply) { 'Yellow' } else { 'Green' })
    Write-Host "Input       : $inFull"
    Write-Host "Reports     : $ReportFolder"

    # --------------------------------------------------------------------------------------------
    # Backups (apply mode) and working copy
    # --------------------------------------------------------------------------------------------
    if ($Apply) {
        if (-not (Test-Path -LiteralPath $BackupFolder)) { New-Item -ItemType Directory -Path $BackupFolder -Force | Out-Null }
        $backupFile = Join-Path $BackupFolder ('{0}_backup_{1}{2}' -f $baseName, $stamp, $ext)
        Copy-FileVerified $inFull $backupFile
        Write-Host "Backup      : $backupFile (SHA256 verified)" -ForegroundColor Green
    }
    # All work happens in a temporary copy; the original is only replaced at the very end after verification.
    $tempCopy = Join-Path ([IO.Path]::GetTempPath()) ('OwnerFill_{0}_{1}{2}' -f $stamp, $baseName, $ext)
    Copy-FileVerified $inFull $tempCopy
    $workFile = $tempCopy

    try { $pkg = Open-ExcelPackage -Path $workFile }
    catch { throw "Could not open the workbook copy '$workFile' (corrupt, password-protected, or not a real .xlsx?): $($_.Exception.Message)" }

    # --------------------------------------------------------------------------------------------
    # Locate worksheet, header row and the 3 columns
    # --------------------------------------------------------------------------------------------
    $headers = @($VMNameHeader, $BusinessOwnerHeader, $TechnicalOwnerHeader)
    $ws = $null; $layout = $null
    $candidates = @($pkg.Workbook.Worksheets)
    if ($WorksheetName) {
        $candidates = @($candidates | Where-Object { $_.Name -ieq $WorksheetName })
        if ($candidates.Count -eq 0) { throw "Worksheet '$WorksheetName' not found. Sheets: $((@($pkg.Workbook.Worksheets) | ForEach-Object { $_.Name }) -join ', ')" }
    }
    foreach ($s in $candidates) {
        $layout = Find-HeaderLayout $s $HeaderSearchRows $headers
        if ($layout) { $ws = $s; break }
    }
    if (-not $ws) {
        throw "No worksheet has all three headers ($($headers -join ' / ')) in its first $HeaderSearchRows rows. Use -WorksheetName / -VMNameHeader / -BusinessOwnerHeader / -TechnicalOwnerHeader if the names differ."
    }
    $hdrRow  = $layout.Row
    $vmCol   = $layout.Columns[$VMNameHeader]
    $fields  = @(
        @{ Name = 'Business Owner';  Col = $layout.Columns[$BusinessOwnerHeader] },
        @{ Name = 'Technical Owner'; Col = $layout.Columns[$TechnicalOwnerHeader] }
    )
    $lastRow = $ws.Dimension.End.Row
    $lastCol = $ws.Dimension.End.Column
    Write-Host ("Worksheet   : '{0}', header row {1}, VM Name col {2}, Business Owner col {3}, Technical Owner col {4}, last row {5}" -f `
        $ws.Name, $hdrRow, $vmCol, $fields[0].Col, $fields[1].Col, $lastRow)

    # --------------------------------------------------------------------------------------------
    # Read rows
    # --------------------------------------------------------------------------------------------
    $records = New-Object System.Collections.Generic.List[object]
    $invalid = New-Object System.Collections.Generic.List[object]
    $emptyRows = 0
    for ($r = $hdrRow + 1; $r -le $lastRow; $r++) {
        $vmRaw = ConvertTo-CleanText $ws.Cells[$r, $vmCol].Value
        $owners = @{}; $formulas = @{}
        foreach ($f in $fields) {
            $cell = $ws.Cells[$r, $f.Col]
            $owners[$f.Name]   = ConvertTo-CleanText $cell.Value
            $formulas[$f.Name] = [string]$cell.Formula
        }
        if (-not $vmRaw -and -not $owners['Business Owner'] -and -not $owners['Technical Owner']) {
            # Completely empty row (common after the real data, formatting only) -> ignore silently.
            $rowHasData = $false
            for ($c = 1; $c -le $lastCol; $c++) { if (ConvertTo-CleanText $ws.Cells[$r, $c].Value) { $rowHasData = $true; break } }
            if (-not $rowHasData) { $emptyRows++; continue }
        }
        $app = Get-AppCode $vmRaw
        $rec = [pscustomobject]@{
            Row = $r; VMName = $vmRaw; AppCode = $app.AppCode; InvalidReason = $app.Reason
            Owners = $owners; Formulas = $formulas
        }
        $records.Add($rec)
        if ($app.Reason) {
            $invalid.Add([pscustomobject]@{
                ExcelRow = $r; VMName = $(if ($vmRaw) { $vmRaw } else { '(blank)' }); Reason = $app.Reason
                BusinessOwner = $owners['Business Owner']; TechnicalOwner = $owners['Technical Owner']
            })
        }
    }
    $validRecords = @($records | Where-Object { $_.AppCode })
    $apps = @($validRecords | ForEach-Object { $_.AppCode } | Sort-Object -Unique)

    # Duplicate VM names are not an error, but worth knowing about.
    $dupNames = @($validRecords | Group-Object { $_.VMName.ToUpperInvariant() } | Where-Object { $_.Count -gt 1 } |
        ForEach-Object { '{0} (rows {1})' -f $_.Group[0].VMName, (($_.Group | ForEach-Object { $_.Row }) -join ', ') })

    # --------------------------------------------------------------------------------------------
    # Evidence per application and field, then decide per application/field
    # --------------------------------------------------------------------------------------------
    # $resolution[field][app] = @{ Status = Resolved|Conflict|NoEvidence; Value; Source; Values = list }
    $resolution = @{}
    $conflicts  = New-Object System.Collections.Generic.List[object]
    foreach ($f in $fields) {
        $resolution[$f.Name] = @{}
        foreach ($grp in ($validRecords | Group-Object AppCode)) {
            $app = $grp.Name
            # compareKey -> @{ Spellings = { exact text -> count } (case-SENSITIVE, keeps every spelling); VMs = list }
            $byKey = [ordered]@{}
            foreach ($rec in $grp.Group) {
                $val = $rec.Owners[$f.Name]
                if (-not $val) { continue }
                $k = Get-CompareKey $val
                if (-not $byKey.Contains($k)) { $byKey[$k] = @{ Spellings = (New-Object System.Collections.Specialized.OrderedDictionary ([StringComparer]::Ordinal)); VMs = New-Object System.Collections.Generic.List[string] } }
                if (-not $byKey[$k].Spellings.Contains($val)) { $byKey[$k].Spellings[$val] = 0 }
                $byKey[$k].Spellings[$val]++
                $byKey[$k].VMs.Add($rec.VMName)
            }
            $missingVMs = @($grp.Group | Where-Object { -not $_.Owners[$f.Name] } | ForEach-Object { $_.VMName })

            if ($byKey.Count -eq 0) {
                $resolution[$f.Name][$app] = @{ Status = 'NoEvidence' }
            }
            elseif ($byKey.Count -eq 1) {
                $entry = $byKey[0]
                # Same person written with different case/spacing -> use the most common spelling (first seen on a tie).
                $best = $null; $bestCount = -1
                foreach ($s in $entry.Spellings.Keys) { if ($entry.Spellings[$s] -gt $bestCount) { $best = $s; $bestCount = $entry.Spellings[$s] } }
                $src = "Consistent owner on {0} VM(s) of app {1}: {2}" -f $entry.VMs.Count, $app, ($entry.VMs -join ', ')
                if ($entry.Spellings.Count -gt 1) { $src += " (spelling variants: $(@($entry.Spellings.Keys) -join ' | '))" }
                $resolution[$f.Name][$app] = @{ Status = 'Resolved'; Value = $best; Source = $src }
            }
            else {
                $summary = @(foreach ($k in $byKey.Keys) { "'{0}' on {1}" -f @($byKey[$k].Spellings.Keys)[0], ($byKey[$k].VMs -join ', ') }) -join '; '
                $resolution[$f.Name][$app] = @{ Status = 'Conflict'; Source = $summary }
                foreach ($k in $byKey.Keys) {
                    $conflicts.Add([pscustomobject]@{
                        AppCode          = $app
                        Field            = $f.Name
                        DistinctValues   = $byKey.Count
                        OwnerValue       = (@($byKey[$k].Spellings.Keys) -join ' | ')
                        VMCount          = $byKey[$k].VMs.Count
                        VMsWithThisValue = ($byKey[$k].VMs -join ', ')
                        VMsMissingOwner  = ($missingVMs -join ', ')
                    })
                }
            }
        }
    }

    # --------------------------------------------------------------------------------------------
    # Build changes / unresolved list
    # --------------------------------------------------------------------------------------------
    $changes    = New-Object System.Collections.Generic.List[object]
    $unresolved = New-Object System.Collections.Generic.List[object]
    $missingCount = @{ 'Business Owner' = 0; 'Technical Owner' = 0 }
    foreach ($rec in $records) {
        foreach ($f in $fields) {
            if ($rec.Owners[$f.Name]) { continue }                       # has a value -> never touched
            $missingCount[$f.Name]++
            $reason = $null
            if (-not $rec.AppCode) {
                $reason = "Invalid VM name: $($rec.InvalidReason)"
            }
            elseif ($rec.Formulas[$f.Name]) {
                $reason = "Cell contains a formula (=$($rec.Formulas[$f.Name])) that evaluates blank - not overwritten"
            }
            else {
                $res = $resolution[$f.Name][$rec.AppCode]
                switch ($res.Status) {
                    'Resolved' {
                        $changes.Add([pscustomobject]@{
                            'VM Name'       = $rec.VMName
                            'App Code'      = $rec.AppCode
                            'Field'         = $f.Name
                            'Old Value'     = ''
                            'New Value'     = $res.Value
                            'Source/Reason' = $res.Source
                            'Status'        = $(if ($Apply) { 'Applied' } else { 'Proposed' })
                            'Excel Row'     = $rec.Row
                            'Column'        = $f.Col
                        })
                    }
                    'Conflict'   { $reason = "CONFLICT - manual review: $($res.Source)" }
                    'NoEvidence' { $reason = "No VM of app $($rec.AppCode) has a $($f.Name) value" }
                }
            }
            if ($reason) {
                $unresolved.Add([pscustomobject]@{
                    'Excel Row' = $rec.Row
                    'VM Name'   = $(if ($rec.VMName) { $rec.VMName } else { '(blank)' })
                    'App Code'  = $rec.AppCode
                    'Field'     = $f.Name
                    'Reason'    = $reason
                })
            }
        }
    }

    # --------------------------------------------------------------------------------------------
    # Apply (only with -Apply): write the filled cells, save, re-open and verify
    # --------------------------------------------------------------------------------------------
    if ($Apply -and $changes.Count -eq 0) {
        Close-ExcelPackage $pkg -NoSave
        $pkg = $null
        Write-Host 'Nothing to fill - the original file was left unchanged.' -ForegroundColor Yellow
    }
    elseif ($Apply) {
        # The workbook is NOT re-saved by the Excel library (that rewrites the whole file). Instead only the
        # filled cells are written into the worksheet XML of the temp copy; everything else stays byte-identical.
        Close-ExcelPackage $pkg -NoSave
        $pkg = $null
        $targets = @($changes | ForEach-Object { [pscustomobject]@{ Row = $_.'Excel Row'; Column = $_.Column; Value = $_.'New Value' } })
        $beforeSnapshot = Get-WorkbookSnapshot $tempCopy $ws.Name $targets
        $beforeParts = Get-ZipPartHashes $tempCopy
        $beforeStyles = Get-XlsxCellStyles $tempCopy $ws.Name $targets
        try { $sheetPart = Set-XlsxCellText $tempCopy $ws.Name $targets }
        catch { throw "Writing the owners into the temporary copy failed - the original was not changed: $($_.Exception.Message)" }

        # Check 1: every part of the file except that one worksheet is byte-for-byte identical.
        $afterParts = Get-ZipPartHashes $tempCopy
        $problems = New-Object System.Collections.Generic.List[string]
        foreach ($k in $beforeParts.Keys) {
            if ($k -eq $sheetPart) { continue }
            if (-not $afterParts.ContainsKey($k)) { $problems.Add("file part missing after update: $k") }
            elseif ($afterParts[$k] -ne $beforeParts[$k]) { $problems.Add("file part changed: $k") }
        }
        foreach ($k in $afterParts.Keys) { if (-not $beforeParts.ContainsKey($k)) { $problems.Add("unexpected new file part: $k") } }

        # Check 2: every filled cell shows exactly the same format (style) as the empty cell did before.
        $afterStyles = Get-XlsxCellStyles $tempCopy $ws.Name $targets
        foreach ($k in $beforeStyles.Keys) {
            if ($afterStyles[$k] -ne $beforeStyles[$k]) { $problems.Add("cell $k : format changed (style $($beforeStyles[$k]) -> $($afterStyles[$k]))") }
        }

        # Check 3: cell-by-cell - same value + formula + format everywhere, except the filled cells, which
        # must hold the new value with their ORIGINAL format.
        $expected = @{}
        foreach ($ch in $changes) { $expected["$($ws.Name)|$($ch.'Excel Row')|$($ch.Column)"] = $ch.'New Value' }
        $afterSnapshot = Get-WorkbookSnapshot $tempCopy $ws.Name $targets
        $defaultFormat = $beforeSnapshot['#defaultFormat']
        $allKeys = New-Object System.Collections.Generic.HashSet[string]
        foreach ($k in $beforeSnapshot.Keys) { [void]$allKeys.Add($k) }
        foreach ($k in $afterSnapshot.Keys)  { [void]$allKeys.Add($k) }
        foreach ($k in $allKeys) {
            $b = if ($beforeSnapshot.ContainsKey($k)) { $beforeSnapshot[$k] } else { "`t`t$defaultFormat" }
            $a = if ($afterSnapshot.ContainsKey($k))  { $afterSnapshot[$k] }  else { "`t`t$defaultFormat" }
            if ($expected.ContainsKey($k)) {
                # (format of filled cells is checked from the XML in check 2)
                $aParts = $a -split "`t"
                if ($aParts[0] -cne $expected[$k]) { $problems.Add("$k : expected value '$($expected[$k])', found '$($aParts[0])'") }
            }
            elseif ($a -cne $b) {
                $bp = $b -split "`t"; $ap = $a -split "`t"
                if ($ap[0] -cne $bp[0]) { $problems.Add("$k : value changed '$($bp[0])' -> '$($ap[0])'") }
                if ($ap[1] -cne $bp[1]) { $problems.Add("$k : formula changed '$($bp[1])' -> '$($ap[1])'") }
                if ($ap[2] -cne $bp[2]) { $problems.Add("$k : format changed [$($bp[2])] -> [$($ap[2])]") }
            }
        }
        if ($problems.Count -gt 0) {
            throw ("Verification FAILED ({0} difference(s)) - the ORIGINAL FILE WAS NOT CHANGED. First differences:`n  {1}" -f `
                $problems.Count, (($problems | Select-Object -First 10) -join "`n  "))
        }

        # Nobody else may have saved the original while we worked, otherwise their edits would be lost.
        if ((Get-FileHash -LiteralPath $inFull -Algorithm SHA256).Hash -ne $originalHash) {
            throw 'The original file was changed by someone else while the script was running - it was NOT overwritten. Run the script again.'
        }
        try { Copy-Item -LiteralPath $tempCopy -Destination $inFull -Force }
        catch { throw "Could not update '$inFull' (is it open in Excel? close it and run again). The original was not changed: $($_.Exception.Message)" }
        if ((Get-FileHash -LiteralPath $inFull -Algorithm SHA256).Hash -ne (Get-FileHash -LiteralPath $tempCopy -Algorithm SHA256).Hash) {
            Copy-Item -LiteralPath $backupFile -Destination $inFull -Force
            throw "Writing the original did not complete correctly - it was restored from the backup $backupFile"
        }
        Write-Host "Verified    : only the $($changes.Count) empty owner cell(s) were filled; all other cells, formulas, colours, formats and sheets are unchanged." -ForegroundColor Green
        Write-Host "Updated     : $inFull" -ForegroundColor Green
    }
    else {
        Close-ExcelPackage $pkg -NoSave
        $pkg = $null
    }

    # --------------------------------------------------------------------------------------------
    # Reports
    # --------------------------------------------------------------------------------------------
    $boConflictApps = @($conflicts | Where-Object { $_.Field -eq 'Business Owner' }  | ForEach-Object { $_.AppCode } | Sort-Object -Unique)
    $toConflictApps = @($conflicts | Where-Object { $_.Field -eq 'Technical Owner' } | ForEach-Object { $_.AppCode } | Sort-Object -Unique)
    $boFilled = @($changes | Where-Object { $_.Field -eq 'Business Owner' }).Count
    $toFilled = @($changes | Where-Object { $_.Field -eq 'Technical Owner' }).Count
    $boUnres  = @($unresolved | Where-Object { $_.Field -eq 'Business Owner' }).Count
    $toUnres  = @($unresolved | Where-Object { $_.Field -eq 'Technical Owner' }).Count

    $changeCols = 'VM Name', 'App Code', 'Field', 'Old Value', 'New Value', 'Source/Reason', 'Status', 'Excel Row'
    $files = @{
        Changes    = "${reportPrefix}_Changes.csv"
        Unresolved = "${reportPrefix}_Unresolved.csv"
        Conflicts  = "${reportPrefix}_Conflicts.csv"
        Invalid    = "${reportPrefix}_InvalidVMNames.csv"
        Summary    = "${reportPrefix}_Summary.txt"
    }
    # Header-only CSVs when a list is empty, so "no rows" is visibly different from "report missing".
    function Export-Report($Rows, [string[]]$Columns, [string]$Path) {
        # (no @($Rows) here: wrapping a List[object] in @() throws 'Argument types do not match' on PowerShell 7.4)
        if ($Rows.Count -gt 0) { $Rows | Select-Object $Columns | Export-Csv -LiteralPath $Path -NoTypeInformation -Encoding UTF8 }
        else { ('"' + ($Columns -join '","') + '"') | Out-File -LiteralPath $Path -Encoding UTF8 }
    }
    Export-Report $changes    $changeCols $files.Changes
    Export-Report $unresolved @('Excel Row', 'VM Name', 'App Code', 'Field', 'Reason') $files.Unresolved
    Export-Report $conflicts  @('AppCode', 'Field', 'DistinctValues', 'OwnerValue', 'VMCount', 'VMsWithThisValue', 'VMsMissingOwner') $files.Conflicts
    Export-Report $invalid    @('ExcelRow', 'VMName', 'Reason', 'BusinessOwner', 'TechnicalOwner') $files.Invalid

    $verb = if ($Apply) { 'filled' } else { 'identified' }
    $lines = New-Object System.Collections.Generic.List[string]
    $lines.Add('=' * 78)
    $lines.Add("OWNER FILL SUMMARY  -  $mode  -  $(Get-Date -Format 'yyyy-MM-dd HH:mm:ss')  -  $ScriptBuild")
    $lines.Add('=' * 78)
    $lines.Add("Input workbook                      : $inFull  (sheet '$($ws.Name)')")
    if ($Apply) { $lines.Add("Updated workbook (original file)    : $inFull") ; $lines.Add("Backup taken before the update      : $backupFile") }
    $lines.Add("Total VMs processed (data rows)     : $($records.Count)   (+ $emptyRows completely empty row(s) ignored)")
    $lines.Add("Unique applications found           : $($apps.Count)")
    $lines.Add("Missing Business Owners found       : $($missingCount['Business Owner'])")
    $lines.Add("Missing Technical Owners found      : $($missingCount['Technical Owner'])")
    $lines.Add(("{0,-36}: {1}" -f "Business Owners $verb", $boFilled))
    $lines.Add(("{0,-36}: {1}" -f "Technical Owners $verb", $toFilled))
    $lines.Add("Owners that could not be determined : $($boUnres + $toUnres)  (Business: $boUnres, Technical: $toUnres)")
    $lines.Add("Apps with conflicting Business Owner: $($boConflictApps.Count)$(if ($boConflictApps.Count) { '  -> ' + ($boConflictApps -join ', ') })")
    $lines.Add("Apps with conflicting Tech. Owner   : $($toConflictApps.Count)$(if ($toConflictApps.Count) { '  -> ' + ($toConflictApps -join ', ') })")
    $lines.Add("Invalid / blank VM names            : $($invalid.Count)")
    if ($dupNames.Count) { $lines.Add("WARNING duplicate VM names          : $($dupNames.Count) -> $($dupNames -join '; ')") }

    $lines.Add('')
    $lines.Add("---- $(if ($Apply) { 'APPLIED' } else { 'PROPOSED' }) CHANGES ($($changes.Count)) ----")
    if ($changes.Count) {
        $lines.Add(($changes | Select-Object 'VM Name', 'App Code', 'Field', @{ n = 'Old Value'; e = { '(blank)' } }, 'New Value', 'Source/Reason' |
            Format-Table -AutoSize -Wrap | Out-String -Width 4096).TrimEnd())
    } else { $lines.Add('(none)') }

    $lines.Add('')
    $lines.Add("---- CONFLICTS - MANUAL REVIEW ($($boConflictApps.Count + $toConflictApps.Count) app/field combination(s)) ----")
    if ($conflicts.Count) {
        $lines.Add(($conflicts | Select-Object AppCode, Field, OwnerValue, VMsWithThisValue, VMsMissingOwner |
            Format-Table -AutoSize -Wrap | Out-String -Width 4096).TrimEnd())
    } else { $lines.Add('(none)') }

    $lines.Add('')
    $lines.Add("---- MISSING OWNERS NOT DETERMINED ($($unresolved.Count)) ----")
    if ($unresolved.Count) { $lines.Add(($unresolved | Format-Table -AutoSize -Wrap | Out-String -Width 4096).TrimEnd()) } else { $lines.Add('(none)') }

    $lines.Add('')
    $lines.Add("---- INVALID / BLANK VM NAMES ($($invalid.Count)) ----")
    if ($invalid.Count) { $lines.Add(($invalid | Format-Table -AutoSize -Wrap | Out-String -Width 4096).TrimEnd()) } else { $lines.Add('(none)') }

    $lines.Add('')
    $lines.Add('Report files:')
    foreach ($k in 'Changes', 'Conflicts', 'Unresolved', 'Invalid', 'Summary') { $lines.Add("  $($files[$k])") }
    if (-not $Apply) {
        $lines.Add('')
        $lines.Add('DRY RUN - the workbook was NOT changed. Review the reports, then run again with -Apply to update it.')
    }
    $lines | Out-File -LiteralPath $files.Summary -Encoding UTF8
    Write-Host ''
    $lines | ForEach-Object { Write-Host $_ }

    # Dry run (or nothing to fill): final proof that the original file was not touched.
    if ($Apply -and $changes.Count -gt 0) { }
    elseif ((Get-FileHash -LiteralPath $inFull -Algorithm SHA256).Hash -ne $originalHash) {
        Write-Host 'WARNING: the input file changed while the script was running (someone else saved it?). Re-run the script.' -ForegroundColor Red
        $exitCode = 3
    }
    else {
        Write-Host "Original file unchanged (SHA256 verified): $inFull" -ForegroundColor Green
    }
}
catch {
    Write-Host ''
    Write-Host "ERROR: $($_.Exception.Message)" -ForegroundColor Red
    if ($_.InvocationInfo) { Write-Host "       at line $($_.InvocationInfo.ScriptLineNumber): $($_.InvocationInfo.Line.Trim())" -ForegroundColor DarkGray }
    Write-Host 'The original workbook was not modified.' -ForegroundColor Yellow
    $exitCode = 1
}
finally {
    if ($pkg) { try { Close-ExcelPackage $pkg -NoSave } catch { } }
    if ($tempCopy -and (Test-Path -LiteralPath $tempCopy)) { Remove-Item -LiteralPath $tempCopy -Force -ErrorAction SilentlyContinue }
}
exit $exitCode
