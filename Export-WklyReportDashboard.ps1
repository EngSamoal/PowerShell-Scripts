<#
.SYNOPSIS
    Presents the "Wkly_Report" worksheet of the weekly ITSM workbook as an interactive,
    executive HTML dashboard - strictly READ-ONLY.

.DESCRIPTION
    Workflow:  Excel workbook (Wkly_Report)  ->  this script  ->  <workbook name>.html dashboard

    SAFETY - the script never changes anything:
      * The original workbook is never opened. It is copied to a private temp folder and
        only the copy is opened, read-only, in a separate hidden Excel instance.
      * Links are not updated, formulas are calculated exactly as when you open the file in Excel
        (time-based formulas such as NOW()/TODAY() show current values), macros/events are disabled, and the
        copy is closed WITHOUT saving and then deleted.
      * Only the "Wkly_Report" worksheet is read. No other tab is read or changed, and
        nothing is ever written to Wkly_Report itself.
      * Excel windows you already have open are not touched. Only the Excel process this
        script started is closed at the end.
      * The SHA-256 hash of the original file is recorded before and after the run and both
        are shown in the dashboard's "Data check" panel.

    WHAT IS SHOWN:
      * Every page of Wkly_Report is detected automatically on every run, so the dashboard
        follows the workbook when pages are added or removed (10 this week, 9 or 12 next week).
      * Values are shown exactly as Excel displays them (cell .Text / shape text), so numbers,
        percentages, dates, durations, "-" and Excel errors such as #DIV/0! are unchanged.
      * Charts are exported as images by Excel itself, so they match the workbook exactly.
      * Cell formatting, KPI-card shapes, colours, rounded corners and shadows are re-drawn
        in HTML at their original positions.
      * "Excel view" on each page shows Excel's own PDF rendering of that page for comparison.
      * "Data check" lists every Excel error, empty formula result, '####' cell, object that
        could not be reproduced and the page-detection method used.

.PARAMETER WorkbookPath
    The weekly workbook. Default: C:\Temp\Bi_Wkly_Report-Daily Data.xlsx

.PARAMETER OutputDirectory
    Where the dashboard (<workbook name>.html, .pdf and .pptx) is written. Default: C:\Temp\ITSM_Dashboard
    Each run replaces only what this script generated earlier (the .html, assets\, the log
    and the PDFs); any other files in that folder are left alone.

.PARAMETER PageMode
    How report pages are found:
      Auto       (default) Page titles first, then manual page breaks, else one page.
      Titles     Large-font page titles start a new page.
      PageBreaks Manual horizontal page breaks in Wkly_Report define the pages.
      Manual     Use -PageStartRows.

.PARAMETER RenderMode
    Rebuilt (default): each page is rebuilt from what Wkly_Report contains - text, KPI cards, tables,
    colours and fonts in Excel's arrangement, charts redrawn natively from their own data (sharp,
    values on hover). Picture: each page is a high-resolution picture of Excel's rendering (r26 behaviour).

.PARAMETER ChartRender
    Excel (default): every chart is Excel's own picture of it, taken at high resolution, so it looks
    exactly as in Excel (labels, callouts, 3D, smoothed lines, colours). Native: charts are redrawn
    from their data (sharp at any zoom, values on hover, but not identical to Excel).

.PARAMETER PageStartRows
    Only with -PageMode Manual: first row of each page, e.g. 1,40,72,105

.PARAMETER TitleMinFontSize
    Smallest font size (pt) that counts as a page title in Titles/Auto mode. Default 20.

.PARAMETER MinPageHeightPt
    Two title candidates closer than this (points) belong to the same page. Default 250.

.PARAMETER DashboardTitle
    Text in the dashboard's top bar. Default "ITSM Executive Dashboard".

.PARAMETER Classification
    Marking shown on the top bar and every page footer. Default "CONFIDENTIAL".

.PARAMETER SkipExcelView
    No longer used (per-page Excel view PDFs are not produced); kept so older command lines still run.

.PARAMETER SkipSheetPdf
    No longer used (no PDF is produced); kept so older command lines still run.

.PARAMETER ExportDashboardPdf
    No longer used (no PDF is produced); kept so older command lines still run.

.PARAMETER SkipPdf
    Do not write <workbook name>.pdf. By default a PDF is written next to the HTML: one PDF page per
    dashboard page, each exactly that page's size (nothing cut off).

.PARAMETER SkipPptx
    Do not write <workbook name>.pptx. By default a PowerPoint file (16:9) is written next to the HTML:
    one slide per dashboard page, each page fitted whole. PowerPoint itself is not needed.

    The PDF and PowerPoint pages are pictures of the finished dashboard pages, taken by Microsoft Edge
    (or Chrome) in a hidden window from the local HTML file; nothing is sent anywhere.

.PARAMETER OpenWhenDone
    Open the dashboard in the default browser when finished.

.EXAMPLE
    .\Export-WklyReportDashboard.ps1

.EXAMPLE
    .\Export-WklyReportDashboard.ps1 -OutputDirectory 'C:\Temp\ITSM_Dashboard' -OpenWhenDone

.EXAMPLE
    .\Export-WklyReportDashboard.ps1 -PageMode Manual -PageStartRows 1,40,72,105,140

.NOTES
    Requires Windows, Microsoft Excel (2010 or later) and Windows PowerShell 5.1 or PowerShell 7.
    No internet access, modules or third-party software are needed.
#>
[CmdletBinding()]
param(
    [string]$WorkbookPath = 'C:\Temp\Bi_Wkly_Report-Daily Data.xlsx',
    [string]$OutputDirectory = 'C:\Temp\ITSM_Dashboard',
    [ValidateSet('Auto', 'Titles', 'PageBreaks', 'Manual')]
    [string]$PageMode = 'Auto',
    [ValidateSet('Rebuilt', 'Picture')]
    [string]$RenderMode = 'Rebuilt',
    [ValidateSet('Excel', 'Native')]
    [string]$ChartRender = 'Excel',
    [int[]]$PageStartRows,
    [double]$TitleMinFontSize = 17,
    [double]$MinPageHeightPt = 250,
    [string]$DashboardTitle = 'ITSM Executive Dashboard',
    [string]$Classification = 'CONFIDENTIAL',
    [switch]$SkipExcelView,
    [switch]$SkipSheetPdf,
    [switch]$ExportDashboardPdf,
    [switch]$SkipPdf,
    [switch]$SkipPptx,
    [switch]$OpenWhenDone
)

$ScriptBuild = '2026.10.23-r34'
$SheetName = 'Wkly_Report'          # fixed by design - no other worksheet is ever read
$Inv = [System.Globalization.CultureInfo]::InvariantCulture
$ErrRegex = '#(DIV/0!|N/A|VALUE!|REF!|NAME\?|NUM!|NULL!|SPILL!|CALC!|FIELD!|BLOCKED!|CONNECT!|BUSY!|UNKNOWN!|GETTING_DATA)'
$GeneratedMarker = '.generated-by-Export-WklyReportDashboard'

$script:LogLines = New-Object System.Collections.Generic.List[string]
$script:Issues = New-Object System.Collections.Generic.List[object]
$script:ChartIndex = 0
$script:PictureIndex = 0
$script:LoggedShapeTypes = @{}
$script:XlApp = $null
$script:XlSheet = $null
$script:ChartMode = $null
$script:XlMadeVisible = $false
$script:ThemeMinor = 'Calibri'
$script:ThemeMajor = 'Calibri Light'
$script:WorkDir = $null
$script:LogPath = Join-Path (Join-Path ([System.IO.Path]::GetTempPath()) 'ITSM_Dashboard_logs') ('extraction-log ' + (Get-Date -Format 'yyyy-MM-dd HHmmss') + '.txt')
$script:Clip = $null
$script:TestClipboard = $null
$script:PictureScale = 3
$script:PdfScale = 2
$script:PdfReady = $null
$script:PictureMethods = @{}
$script:BitmapScales = New-Object System.Collections.Generic.List[double]
$script:ChartScales = New-Object System.Collections.Generic.List[double]
$script:ChartZoomSet = $false
$script:ChartShots = @{ sharp = 0; export = 0; fails = 0; why = $null }
$script:ItemSeq = 0

# ======================================================================================
#  Small helpers
# ======================================================================================

function Write-Log {
    param([string]$Message, [ValidateSet('INFO', 'WARN', 'ERROR', 'OK')][string]$Level = 'INFO', [switch]$FileOnly)
    $line = '{0} [{1}] {2}' -f (Get-Date -Format 'yyyy-MM-dd HH:mm:ss'), $Level, $Message
    $script:LogLines.Add($line)
    if ($FileOnly) { return }
    $color = switch ($Level) { 'ERROR' { 'Red' } 'WARN' { 'Yellow' } 'OK' { 'Green' } default { 'Gray' } }
    Write-Host $line -ForegroundColor $color
}

function Add-Issue {
    param(
        [ValidateSet('Error', 'Warning', 'Info')][string]$Severity,
        [string]$Location,
        [string]$Message,
        [double]$Y = -1
    )
    $script:Issues.Add([pscustomobject]@{ Severity = $Severity; Location = $Location; Message = $Message; Y = $Y; Page = $null })
    $level = switch ($Severity) { 'Error' { 'WARN' } 'Warning' { 'WARN' } default { 'INFO' } }
    Write-Log ("{0}: {1}" -f $Location, $Message) $level
}

function Format-ErrorDetail {
    # Message plus the script line and call path, so a failure can be located from the log alone.
    param($ErrorRecord)
    $msg = $ErrorRecord.Exception.Message
    $ii = $ErrorRecord.InvocationInfo
    if ($ii -and $ii.ScriptLineNumber) { $msg += (' [line {0}: {1}]' -f $ii.ScriptLineNumber, $ii.Line.Trim()) }
    $stack = [string]$ErrorRecord.ScriptStackTrace
    if ($stack) { $msg += ' [stack: ' + (($stack -split "`r?`n" | Select-Object -First 4) -join ' <- ') + ']' }
    return $msg
}

function Get-SourceHash {
    # SHA-256 of the workbook, read with full sharing so it also works while the file is open in Excel.
    param([string]$Path)
    $fs = $null; $sha = $null
    try {
        $fs = [System.IO.File]::Open($Path, [System.IO.FileMode]::Open, [System.IO.FileAccess]::Read, ([System.IO.FileShare]::ReadWrite -bor [System.IO.FileShare]::Delete))
        $sha = [System.Security.Cryptography.SHA256]::Create()
        $bytes = $sha.ComputeHash($fs)
        return ([System.BitConverter]::ToString($bytes)).Replace('-', '')
    }
    catch { return $null }
    finally { if ($sha) { $sha.Dispose() }; if ($fs) { $fs.Dispose() } }
}

function Copy-SourceShared {
    # Copy the workbook even while Excel has it open (read-only, full sharing).
    param([string]$From, [string]$To)
    try { Copy-Item -LiteralPath $From -Destination $To -Force -ErrorAction Stop; return }
    catch { }
    $in = $null; $out = $null
    try {
        $in = [System.IO.File]::Open($From, [System.IO.FileMode]::Open, [System.IO.FileAccess]::Read, ([System.IO.FileShare]::ReadWrite -bor [System.IO.FileShare]::Delete))
        $out = [System.IO.File]::Create($To)
        $in.CopyTo($out)
    }
    finally { if ($out) { $out.Dispose() }; if ($in) { $in.Dispose() } }
}

function Format-Num {
    param([double]$Value)
    return $Value.ToString('0.###', $Inv)
}

function Test-MsoTrue {
    param($Value)
    if ($null -eq $Value -or $Value -is [System.DBNull]) { return $false }
    if ($Value -is [bool]) { return $Value }
    try { return ([int]$Value -eq -1 -or [int]$Value -eq 1) } catch { return $false }
}

function ConvertTo-CssColor {
    param($Ole, [double]$Alpha = 1)
    if ($null -eq $Ole -or $Ole -is [System.DBNull]) { return $null }
    try { $c = [int64]$Ole } catch { return $null }
    if ($c -lt 0 -or $c -gt 0xFFFFFF) { return $null }
    $r = [int]($c -band 0xFF); $g = [int](($c -shr 8) -band 0xFF); $b = [int](($c -shr 16) -band 0xFF)
    if ($Alpha -lt 0.999) { return ('rgba({0},{1},{2},{3})' -f $r, $g, $b, (Format-Num ([math]::Max(0, $Alpha)))) }
    return ('#{0:x2}{1:x2}{2:x2}' -f $r, $g, $b)
}

function ConvertTo-ColumnLetter {
    param([int]$Col)
    $s = ''
    while ($Col -gt 0) { $m = ($Col - 1) % 26; $s = [char](65 + $m) + $s; $Col = [int][math]::Floor(($Col - 1) / 26) }
    return $s
}

function ConvertTo-HtmlText {
    param([string]$Text)
    if ([string]::IsNullOrEmpty($Text)) { return '' }
    $t = [System.Net.WebUtility]::HtmlEncode($Text)
    $t = $t -replace "`r`n", '<br>'
    $t = $t -replace "[`r`n`v]", '<br>'
    return $t
}

function ConvertTo-Attr {
    param([string]$Text)
    if ($null -eq $Text) { return '' }
    return [System.Net.WebUtility]::HtmlEncode(($Text -replace "[`r`n`v]+", ' '))
}

function Get-CssFont {
    param([string]$Name)
    $fallback = "Calibri,'Segoe UI',Arial,sans-serif"
    if ($Name -and $Name.StartsWith('+mj')) { $Name = $script:ThemeMajor }       # theme heading font
    elseif ($Name -and $Name.StartsWith('+')) { $Name = $script:ThemeMinor }     # theme body font
    if ([string]::IsNullOrWhiteSpace($Name)) { return $fallback }
    return "'" + ($Name -replace "['`"<>;]", '') + "',$fallback"
}

function Get-ItemText {
    param($Item)
    if ($Item.Kind -eq 'cell') { return [string]$Item.Text }
    if ($Item.Kind -eq 'shape' -and $Item.Runs) { return (($Item.Runs | ForEach-Object { $_.Text }) -join '') }
    return ''
}

function Get-ItemMaxFont {
    param($Item)
    if ($Item.Kind -eq 'cell') { return [double]$Item.FontSize }
    $max = 0
    if ($Item.Kind -eq 'shape' -and $Item.Runs) {
        foreach ($r in $Item.Runs) { if ($r.Text -match '\S' -and $r.Size -gt $max) { $max = [double]$r.Size } }
    }
    return $max
}

function Get-RowAtY {
    param([double[]]$RowTops, [int]$LastRow, [double]$Y)
    $lo = 1; $hi = $LastRow
    while ($lo -lt $hi) {
        $mid = [int][math]::Ceiling(($lo + $hi) / 2)
        if ($RowTops[$mid] -le $Y + 0.01) { $lo = $mid } else { $hi = $mid - 1 }
    }
    return $lo
}

# ======================================================================================
#  Pictures: located in the copied .xlsx package (avoids the clipboard entirely)
# ======================================================================================

function Get-PictureMediaMap {
    param([string]$PackagePath, [string]$Sheet, [string]$AssetsDir)
    $map = @{}
    Add-Type -AssemblyName System.IO.Compression.FileSystem
    $zip = [System.IO.Compression.ZipFile]::OpenRead($PackagePath)
    try {
        $readXml = {
            param($path)
            $e = $zip.GetEntry($path)
            if (-not $e) { return $null }
            $sr = New-Object System.IO.StreamReader($e.Open())
            try { [xml]$sr.ReadToEnd() } finally { $sr.Dispose() }
        }
        $resolve = {
            param([string]$baseDir, [string]$target)
            if ($target.StartsWith('/')) { return $target.TrimStart('/') }
            $parts = New-Object System.Collections.Generic.List[string]
            foreach ($p in ($baseDir.TrimEnd('/') + '/' + $target).Split('/')) {
                if ($p -eq '..') { if ($parts.Count) { $parts.RemoveAt($parts.Count - 1) } }
                elseif ($p -and $p -ne '.') { $parts.Add($p) }
            }
            return ($parts -join '/')
        }
        $relsPathFor = { param([string]$part) $d = $part.Substring(0, $part.LastIndexOf('/')); $f = $part.Substring($part.LastIndexOf('/') + 1); "$d/_rels/$f.rels" }

        $rNs = 'http://schemas.openxmlformats.org/officeDocument/2006/relationships'
        $wb = & $readXml 'xl/workbook.xml'
        $sheetNode = @($wb.workbook.sheets.sheet) | Where-Object { $_.name -eq $Sheet } | Select-Object -First 1
        if (-not $sheetNode) { return $map }
        $rid = $sheetNode.GetAttribute('id', $rNs)
        $wbRels = & $readXml 'xl/_rels/workbook.xml.rels'
        $rel = @($wbRels.Relationships.Relationship) | Where-Object { $_.Id -eq $rid } | Select-Object -First 1
        if (-not $rel) { return $map }
        $sheetPart = & $resolve 'xl' $rel.Target
        $sheetRels = & $readXml (& $relsPathFor $sheetPart)
        if (-not $sheetRels) { return $map }
        $drawRel = @($sheetRels.Relationships.Relationship) | Where-Object { $_.Type -like '*/drawing' } | Select-Object -First 1
        if (-not $drawRel) { return $map }
        $drawPart = & $resolve ($sheetPart.Substring(0, $sheetPart.LastIndexOf('/'))) $drawRel.Target
        $drawing = & $readXml $drawPart
        $drawRels = & $readXml (& $relsPathFor $drawPart)
        if (-not $drawing -or -not $drawRels) { return $map }

        $ns = New-Object System.Xml.XmlNamespaceManager($drawing.NameTable)
        $ns.AddNamespace('xdr', 'http://schemas.openxmlformats.org/drawingml/2006/spreadsheetDrawing')
        $ns.AddNamespace('a', 'http://schemas.openxmlformats.org/drawingml/2006/main')
        $drawDir = $drawPart.Substring(0, $drawPart.LastIndexOf('/'))
        $n = 0
        foreach ($pic in $drawing.SelectNodes('//xdr:pic', $ns)) {
            $nv = $pic.SelectSingleNode('xdr:nvPicPr/xdr:cNvPr', $ns)
            $blip = $pic.SelectSingleNode('.//a:blip', $ns)
            if (-not $nv -or -not $blip) { continue }
            $name = $nv.GetAttribute('name')
            $embed = $blip.GetAttribute('embed', $rNs)
            $mrel = @($drawRels.Relationships.Relationship) | Where-Object { $_.Id -eq $embed } | Select-Object -First 1
            if (-not $mrel -or $map.ContainsKey($name)) { continue }
            $mediaPart = & $resolve $drawDir $mrel.Target
            $entry = $zip.GetEntry($mediaPart)
            if (-not $entry) { continue }
            $n++
            $ext = [System.IO.Path]::GetExtension($mediaPart).ToLowerInvariant()
            if ($ext -notin '.png', '.jpg', '.jpeg', '.gif', '.bmp', '.svg', '.webp') { continue }
            $file = 'picture-{0:D3}{1}' -f $n, $ext
            $dest = Join-Path $AssetsDir $file
            $in = $entry.Open(); $out = [System.IO.File]::Create($dest)
            try { $in.CopyTo($out) } finally { $out.Dispose(); $in.Dispose() }
            $map[$name] = "assets/$file"
        }
    }
    finally { $zip.Dispose() }
    return $map
}

# ======================================================================================
#  Excel extraction (read-only)
# ======================================================================================

function Get-SheetGeometry {
    param($Ws, [int]$LastRow, [int]$LastCol)
    $rowTops = New-Object double[] ($LastRow + 2)
    for ($r = 1; $r -le $LastRow + 1; $r++) { $rowTops[$r] = [double]$Ws.Rows.Item($r).Top }
    $colLefts = New-Object double[] ($LastCol + 2)
    for ($c = 1; $c -le $LastCol + 1; $c++) { $colLefts[$c] = [double]$Ws.Columns.Item($c).Left }
    return [pscustomobject]@{ RowTops = $rowTops; ColLefts = $colLefts; LastRow = $LastRow; LastCol = $LastCol }
}

function Get-FillRuns {
    param($Ws, [int]$Row, [int]$C1, [int]$C2, $Out)
    $rng = $Ws.Range($Ws.Cells.Item($Row, $C1), $Ws.Cells.Item($Row, $C2))
    $interior = $rng.Interior
    $pattern = $interior.Pattern
    $color = $interior.Color
    $uniform = ($null -ne $pattern) -and ($pattern -isnot [System.DBNull]) -and ($null -ne $color) -and ($color -isnot [System.DBNull])
    if ($uniform) {
        if ([int]$pattern -ne -4142) { $Out.Add([pscustomobject]@{ C1 = $C1; C2 = $C2; Color = (ConvertTo-CssColor $color) }) }
        return
    }
    if ($C1 -ge $C2) { return }
    $mid = [int][math]::Floor(($C1 + $C2) / 2)
    Get-FillRuns -Ws $Ws -Row $Row -C1 $C1 -C2 $mid -Out $Out
    Get-FillRuns -Ws $Ws -Row $Row -C1 ($mid + 1) -C2 $C2 -Out $Out
}

function Read-FillItems {
    # Background fills of cells (incl. empty ones, e.g. coloured side bars), merged into rectangles.
    param($Ws, $Geo, [int]$R1, [int]$C1, [int]$R2, [int]$C2)
    $items = New-Object System.Collections.Generic.List[object]
    $open = @{}
    for ($r = $R1; $r -le $R2; $r++) {
        $h = $Geo.RowTops[$r + 1] - $Geo.RowTops[$r]
        if ($h -le 0) { continue }
        $runs = New-Object System.Collections.Generic.List[object]
        try { Get-FillRuns -Ws $Ws -Row $r -C1 $C1 -C2 $C2 -Out $runs } catch { continue }
        $merged = New-Object System.Collections.Generic.List[object]
        foreach ($run in $runs) {
            if (-not $run.Color) { continue }
            $last = if ($merged.Count) { $merged[$merged.Count - 1] } else { $null }
            if ($last -and $last.C2 + 1 -eq $run.C1 -and $last.Color -eq $run.Color) { $last.C2 = $run.C2 } else { $merged.Add($run) }
        }
        $next = @{}
        foreach ($run in $merged) {
            $key = '{0}|{1}|{2}' -f $run.C1, $run.C2, $run.Color
            if ($open.ContainsKey($key) -and $open[$key].LastRow -eq $r - 1) {
                $it = $open[$key]; $it.H = $Geo.RowTops[$r + 1] - $it.T; $it.LastRow = $r
            }
            else {
                $L = $Geo.ColLefts[$run.C1]
                $it = [pscustomobject]@{ Kind = 'fill'; L = $L; T = $Geo.RowTops[$r]; W = $Geo.ColLefts[$run.C2 + 1] - $L; H = $h; Z = 0; Bg = $run.Color; LastRow = $r }
                $items.Add($it)
            }
            $next[$key] = $it
        }
        # keep rectangles that may continue after a hidden row
        foreach ($k in @($open.Keys)) { if (-not $next.ContainsKey($k)) { $open.Remove($k) } }
        foreach ($k in $next.Keys) { $open[$k] = $next[$k] }
    }
    return , $items
}

function Read-CellItems {
    param($Ws, $Geo)
    $items = New-Object System.Collections.Generic.List[object]
    $used = $Ws.UsedRange
    $r0 = [int]$used.Row; $c0 = [int]$used.Column
    $nr = [int]$used.Rows.Count; $nc = [int]$used.Columns.Count
    $vals = $used.Value2
    $forms = $used.Formula
    if ($nr -eq 1 -and $nc -eq 1) {
        $v1 = $vals; $f1 = $forms
        $vals = [Array]::CreateInstance([object], @(1, 1), @(1, 1)); $vals.SetValue($v1, 1, 1)
        $forms = [Array]::CreateInstance([object], @(1, 1), @(1, 1)); $forms.SetValue($f1, 1, 1)
    }
    $hAlign = @{ -4131 = 'left'; -4108 = 'center'; -4152 = 'right'; 7 = 'center'; -4130 = 'justify'; -4117 = 'justify'; 5 = 'left' }
    $vAlign = @{ -4160 = 'top'; -4108 = 'middle'; -4107 = 'bottom'; -4130 = 'middle'; -4117 = 'middle' }
    $weights = @{ 1 = 0.5; 2 = 1; -4138 = 2; 4 = 3 }
    $count = 0

    for ($i = 1; $i -le $nr; $i++) {
        if ($i % 50 -eq 0) { Write-Progress -Activity 'Reading Wkly_Report cells' -PercentComplete ([int](100 * $i / $nr)) }
        for ($j = 1; $j -le $nc; $j++) {
            $v = $vals[$i, $j]
            $f = $forms[$i, $j]
            $r = $r0 + $i - 1; $c = $c0 + $j - 1
            $addr = (ConvertTo-ColumnLetter $c) + $r
            $isFormula = ($f -is [string]) -and $f.StartsWith('=')
            if ($null -eq $v -or ($v -is [string] -and $v.Length -eq 0)) {
                if ($isFormula) {
                    $L = $Geo.ColLefts[$c]; $T = $Geo.RowTops[$r]
                    $W = $Geo.ColLefts[$c + 1] - $L; $H = $Geo.RowTops[$r + 1] - $T
                    if ($W -gt 0 -and $H -gt 0) {
                        $items.Add([pscustomobject]@{ Kind = 'blank'; L = $L; T = $T; W = $W; H = $H; Z = 0.5; Address = $addr })
                    }
                }
                continue
            }
            try {
                $cell = $Ws.Cells.Item($r, $c)
                $text = [string]$cell.Text
                if ($text.Length -eq 0) { continue }        # Excel itself displays nothing
                $r2 = $r; $c2 = $c
                if ($cell.MergeCells -eq $true) {
                    $ma = $cell.MergeArea
                    $r2 = [int]$ma.Row + [int]$ma.Rows.Count - 1
                    $c2 = [int]$ma.Column + [int]$ma.Columns.Count - 1
                    if ($r2 -gt $Geo.LastRow) { $r2 = $Geo.LastRow }
                    if ($c2 -gt $Geo.LastCol) { $c2 = $Geo.LastCol }
                }
                $L = $Geo.ColLefts[$c]; $T = $Geo.RowTops[$r]
                $W = $Geo.ColLefts[$c2 + 1] - $L; $H = $Geo.RowTops[$r2 + 1] - $T
                if ($W -le 0 -or $H -le 0) { continue }       # hidden row/column

                $df = $cell.DisplayFormat
                $font = $df.Font
                $ha = [int]$df.HorizontalAlignment
                $align = if ($hAlign.ContainsKey($ha)) { $hAlign[$ha] }
                elseif ($v -is [double]) { 'right' }
                elseif ($v -is [bool] -or $v -is [int]) { 'center' }
                else { 'left' }
                $va = [int]$df.VerticalAlignment
                $vAlignName = if ($vAlign.ContainsKey($va)) { $vAlign[$va] } else { 'bottom' }

                $bg = $null
                $interior = $df.Interior
                if ([int]$interior.Pattern -ne -4142) { $bg = ConvertTo-CssColor $interior.Color }

                $borders = New-Object System.Collections.Generic.List[string]
                foreach ($edge in @(@(7, 'left'), @(8, 'top'), @(9, 'bottom'), @(10, 'right'))) {
                    $b = $df.Borders.Item($edge[0])
                    if ([int]$b.LineStyle -ne -4142) {
                        $bw = if ($weights.ContainsKey([int]$b.Weight)) { $weights[[int]$b.Weight] } else { 1 }
                        $bc = ConvertTo-CssColor $b.Color
                        if (-not $bc) { $bc = '#000' }
                        $borders.Add(('border-{0}:{1}pt solid {2}' -f $edge[1], (Format-Num $bw), $bc))
                    }
                }
                $orient = 0
                try { $o = [int]$df.Orientation; if ($o -ge -90 -and $o -le 90) { $orient = $o } elseif ($o -eq -4171) { $orient = 90 } elseif ($o -eq -4170) { $orient = -90 } } catch { }

                $items.Add([pscustomobject]@{
                        Kind      = 'cell'; L = $L; T = $T; W = $W; H = $H; Z = 1
                        Address   = $addr; Text = $text
                        FontName  = [string]$font.Name; FontSize = [double]$font.Size
                        Bold      = (Test-MsoTrue $font.Bold); Italic = (Test-MsoTrue $font.Italic)
                        Underline = ([int]$font.Underline -ne -4142)
                        Color     = (ConvertTo-CssColor $font.Color)
                        Align     = $align; VAlign = $vAlignName; Wrap = (Test-MsoTrue $df.WrapText)
                        Indent    = [int]$df.IndentLevel; Orientation = $orient
                        Bg        = $bg; Borders = ($borders -join ';')
                    })
                $count++
                if ($text -match "^$ErrRegex$" -or $text -match $ErrRegex) {
                    Add-Issue -Severity Error -Location "Wkly_Report!$addr" -Message ("Excel error value '{0}' is shown exactly as Excel displays it." -f $text) -Y $T
                }
                elseif ($text -match '^#+$') {
                    Add-Issue -Severity Warning -Location "Wkly_Report!$addr" -Message "Excel displays '####' (column too narrow for the value). Shown as displayed; widen the column in the workbook to see the value." -Y $T
                }
            }
            catch {
                Add-Issue -Severity Warning -Location "Wkly_Report!$addr" -Message ("Cell could not be read: {0}" -f $_.Exception.Message) -Y $Geo.RowTops[$r]
            }
        }
    }
    Write-Progress -Activity 'Reading Wkly_Report cells' -Completed
    Write-Log ("Read {0} non-empty cells." -f $count)
    return , $items
}

function Get-ShapeBackground {
    param($Shape)
    try {
        $fill = $Shape.Fill
        if (-not (Test-MsoTrue $fill.Visible)) { return $null }
        $type = [int]$fill.Type
        if ($type -eq 3) {
            try {
                $stops = $fill.GradientStops
                $parts = New-Object System.Collections.Generic.List[string]
                for ($i = 1; $i -le $stops.Count; $i++) {
                    $s = $stops.Item($i)
                    $parts.Add(('{0} {1}%' -f (ConvertTo-CssColor $s.Color.RGB (1 - [double]$s.Transparency)), (Format-Num (100 * [double]$s.Position))))
                }
                $angle = 180
                try { $angle = [double]$fill.GradientAngle + 90 } catch { }
                if ($parts.Count -ge 2) { return ('linear-gradient({0}deg,{1})' -f (Format-Num $angle), ($parts -join ',')) }
            }
            catch { }
        }
        return (ConvertTo-CssColor $fill.ForeColor.RGB (1 - [double]$fill.Transparency))
    }
    catch { return $null }
}

function Get-ChartKind {
    # Excel chart type -> native renderer family; $null = not rebuilt (Excel's chart image is shown).
    param([int]$Type)
    if ($Type -in 51, 52, 53, 54, 55, 56, -4100, 92, 93, 94, 99, 100, 101, 106, 107, 108) { return 'column' }
    if ($Type -in 57, 58, 59, 60, 61, 62, 95, 96, 97, 102, 103, 104, 109, 110, 111) { return 'bar' }
    if ($Type -in 4, 63, 64, 65, 66, 67, -4101, 72, 73, 74, 75) { return 'line' }
    if ($Type -in 1, 76, 77, -4098) { return 'area' }
    if ($Type -in 5, 69, -4102, 70) { return 'pie' }
    if ($Type -in -4120, 80) { return 'doughnut' }
    return $null
}

function Get-OleCss {
    param($Value)
    if ($null -eq $Value -or $Value -is [System.DBNull]) { return $null }
    try { return (ConvertTo-CssColor ([int64][double]$Value)) } catch { return $null }
}

function Get-LegacyFont {
    # Font of a chart element (title, legend, tick labels, data labels) as a small hashtable.
    param($Font)
    $f = [ordered]@{ color = $null; size = $null; bold = $false; italic = $false; name = $null }
    if ($null -eq $Font) { return $f }
    try { $f.color = Get-OleCss $Font.Color } catch { }
    try { $f.size = [double]$Font.Size } catch { }
    try { $f.bold = Test-MsoTrue $Font.Bold } catch { }
    try { $f.italic = Test-MsoTrue $Font.Italic } catch { }
    try { $f.name = [string]$Font.Name } catch { }
    return $f
}

function ConvertTo-NumberList {
    # COM array (or single value) -> list of numbers, $null for blanks / Excel errors.
    param($Values)
    $out = New-Object System.Collections.Generic.List[object]
    if ($null -eq $Values) { return , $out }
    $arr = $Values
    if ($arr -isnot [System.Array]) { $arr = , $arr }
    foreach ($v in $arr) {
        if ($v -is [double] -or $v -is [single] -or $v -is [decimal] -or $v -is [int64]) { $out.Add([double]$v) }
        elseif ($v -is [int] -and [int]$v -gt -2146826300 -and [int]$v -lt -2146826200) { $out.Add($null) }   # Excel error code
        elseif ($v -is [int] -or $v -is [int16] -or $v -is [byte]) { $out.Add([double]$v) }
        else { $out.Add($null) }
    }
    return , $out
}

function ConvertTo-TextList {
    param($Values)
    $out = New-Object System.Collections.Generic.List[string]
    if ($null -eq $Values) { return , $out }
    $arr = $Values
    if ($arr -isnot [System.Array]) { $arr = , $arr }
    foreach ($v in $arr) { $out.Add([string]$v) }
    return , $out
}

function Get-AxisSpec {
    param($Chart, [int]$Type, [int]$Group)
    $a = [ordered]@{ show = $false; min = $null; max = $null; major = $null; format = 'General'; font = $null; grid = $null; reverse = $false }
    try { if (-not (Test-MsoTrue $Chart.HasAxis($Type, $Group))) { return $a } } catch { return $a }
    $ax = $null
    try { $ax = $Chart.Axes($Type, $Group) } catch { return $a }
    $a.show = $true
    try { if ([int]$ax.TickLabelPosition -eq -4142) { $a.show = $false } } catch { }
    if ($Type -eq 2) {
        try { $a.min = [double]$ax.MinimumScale; $a.max = [double]$ax.MaximumScale; $a.major = [double]$ax.MajorUnit } catch { }
    }
    try { $a.format = [string]$ax.TickLabels.NumberFormat } catch { }
    try { $a.font = Get-LegacyFont $ax.TickLabels.Font } catch { }
    try { $a.reverse = Test-MsoTrue $ax.ReversePlotOrder } catch { }
    try {
        if (Test-MsoTrue $ax.HasMajorGridlines) {
            $a.grid = ConvertTo-CssColor $ax.MajorGridlines.Format.Line.ForeColor.RGB (1 - [double]$ax.MajorGridlines.Format.Line.Transparency)
            if (-not $a.grid) { $a.grid = '#d9d9d9' }
        }
    }
    catch { }
    return $a
}

function Get-ChartSpec {
    # Everything needed to redraw an Excel chart natively: type, series values/categories/colours,
    # Excel's own axis scale and number formats, labels, legend, title and background.
    param($Shape)
    $ch = $Shape.Chart
    $spec = [ordered]@{ kind = $null; excelType = 0; title = $null; titleFont = $null; bg = $null; plotBg = $null; rounded = $false; shadow = $false
        font = $null; legend = $null; legendFont = $null; gap = 150; overlap = 0; firstAngle = 0; hole = 50
        x = $null; y = $null; y2 = $null; series = (New-Object System.Collections.Generic.List[object]) }
    try { $spec.excelType = [int]$ch.ChartType } catch { }
    $spec.kind = Get-ChartKind $spec.excelType
    try { if (Test-MsoTrue $ch.HasTitle) { $spec.title = [string]$ch.ChartTitle.Text; $spec.titleFont = Get-LegacyFont $ch.ChartTitle.Font } } catch { }
    try { $spec.bg = Get-ShapeBackground ([pscustomobject]@{ Fill = $ch.ChartArea.Format.Fill }) } catch { }
    try { $spec.plotBg = Get-ShapeBackground ([pscustomobject]@{ Fill = $ch.PlotArea.Format.Fill }) } catch { }
    try { $spec.rounded = Test-MsoTrue $ch.Parent.RoundedCorners } catch { }
    try { $spec.shadow = Test-MsoTrue $ch.Parent.Shadow } catch { }
    try { $spec.font = Get-LegacyFont $ch.ChartArea.Font } catch { }
    try {
        if (Test-MsoTrue $ch.HasLegend) {
            $spec.legend = switch ([int]$ch.Legend.Position) { -4160 { 'top' } -4131 { 'left' } -4152 { 'right' } 2 { 'right' } default { 'bottom' } }
            $spec.legendFont = Get-LegacyFont $ch.Legend.Font
        }
    }
    catch { }
    try { $g = $ch.ChartGroups(1); try { $spec.gap = [double]$g.GapWidth } catch { }; try { $spec.overlap = [double]$g.Overlap } catch { }; try { $spec.firstAngle = [double]$g.FirstSliceAngle } catch { }; try { $spec.hole = [double]$g.DoughnutHoleSize } catch { } } catch { }
    $spec.x = Get-AxisSpec $ch 1 1
    $spec.y = Get-AxisSpec $ch 2 1
    $spec.y2 = Get-AxisSpec $ch 2 2
    $n = 0
    try { $n = [int]$ch.SeriesCollection().Count } catch { }
    for ($i = 1; $i -le $n; $i++) {
        $se = $null
        try { $se = $ch.SeriesCollection($i) } catch { continue }
        $sr = [ordered]@{ name = ''; kind = $null; axis = 1; values = $null; cats = $null; fill = $null; fillOn = $true; line = $null; lineOn = $true; width = 2
            marker = $false; markerColor = $null; points = $null; labels = $null; stacked = $false }
        try { $sr.name = [string]$se.Name } catch { }
        try { $t = [int]$se.ChartType; $sr.kind = Get-ChartKind $t; $sr.stacked = ($t -in 52, 53, 55, 56, 58, 59, 61, 62, 63, 64, 66, 67, 76, 77) } catch { $sr.kind = $spec.kind }
        if (-not $sr.kind) { $spec.kind = $null }
        try { $sr.axis = [int]$se.AxisGroup } catch { }
        try { $sr.values = ConvertTo-NumberList $se.Values } catch { }
        try { $sr.cats = ConvertTo-TextList $se.XValues } catch { }
        try { $sr.fillOn = Test-MsoTrue $se.Format.Fill.Visible; $sr.fill = ConvertTo-CssColor $se.Format.Fill.ForeColor.RGB } catch { }
        try { $sr.lineOn = Test-MsoTrue $se.Format.Line.Visible; $sr.line = ConvertTo-CssColor $se.Format.Line.ForeColor.RGB; $sr.width = [double]$se.Format.Line.Weight } catch { }
        try { $sr.marker = ([int]$se.MarkerStyle -ne -4142); $sr.markerColor = Get-OleCss $se.MarkerBackgroundColor } catch { }
        if ($sr.kind -in 'column', 'bar', 'pie', 'doughnut') {
            $cols = New-Object System.Collections.Generic.List[object]
            $cnt = 0; if ($sr.values) { $cnt = [math]::Min(60, $sr.values.Count) }
            for ($j = 1; $j -le $cnt; $j++) {
                $c = $null
                try { $pf = $se.Points($j).Format.Fill; if (Test-MsoTrue $pf.Visible) { $c = ConvertTo-CssColor $pf.ForeColor.RGB } } catch { }
                $cols.Add($c)
            }
            $sr.points = $cols
        }
        try {
            if (Test-MsoTrue $se.HasDataLabels) {
                $dl = $se.DataLabels()
                $lb = [ordered]@{ value = $true; percent = $false; category = $false; series = $false; format = 'General'; font = $null }
                try { $lb.value = Test-MsoTrue $dl.ShowValue } catch { }
                try { $lb.percent = Test-MsoTrue $dl.ShowPercentage } catch { }
                try { $lb.category = Test-MsoTrue $dl.ShowCategoryName } catch { }
                try { $lb.series = Test-MsoTrue $dl.ShowSeriesName } catch { }
                try { $lb.format = [string]$dl.NumberFormat } catch { }
                try { $lb.font = Get-LegacyFont $dl.Font } catch { }
                $sr.labels = $lb
            }
        }
        catch { }
        $spec.series.Add($sr)
    }
    if ($spec.series.Count -eq 0) { $spec.kind = $null }
    Repair-ChartSpec $spec
    return $spec
}

function Repair-ChartSpec {
    # A combination chart (bars + line) has no single Excel chart type; it is drawn natively when
    # every series is a kind the dashboard can draw.
    param($Spec)
    if ($null -eq $Spec -or $Spec.kind) { return }
    $kinds = New-Object System.Collections.Generic.List[string]
    foreach ($sr in $Spec.series) {
        $k = $sr.kind
        if (-not $k -and [int]$Spec.excelType -in 72, 73, 74, 75) { $k = 'line'; $sr.kind = 'line' }
        if (-not $k) { return }
        $kinds.Add([string]$k)
    }
    if ($kinds.Count -eq 0) { return }
    if ($kinds.Contains('pie') -or $kinds.Contains('doughnut')) { return }
    if ($kinds.Contains('bar')) { $Spec.kind = 'bar' } elseif ($kinds.Contains('column')) { $Spec.kind = 'column' } elseif ($kinds.Contains('area')) { $Spec.kind = 'area' } else { $Spec.kind = 'line' }
}

function Get-RunInfo {
    param($Range)
    $f = $Range.Font
    $color = $null
    try { $color = ConvertTo-CssColor $f.Fill.ForeColor.RGB } catch { }
    $under = $false
    try { $under = ([int]$f.UnderlineStyle -ne 0) } catch { }
    $colorInfo = ''
    try { $fc = $f.Fill.ForeColor; $colorInfo = 'theme {0}, brightness {1}, tint {2}' -f [int]$fc.ObjectThemeColor, (Format-Num ([double]$fc.Brightness)), (Format-Num ([double]$fc.TintAndShade)) } catch { }
    return [pscustomobject]@{
        Text = [string]$Range.Text; Size = [double]$f.Size; Bold = (Test-MsoTrue $f.Bold)
        Italic = (Test-MsoTrue $f.Italic); Underline = $under; Color = $color; Font = [string]$f.Name
        ColorInfo = $colorInfo; ColorAdjustedFrom = $null
    }
}

function Get-LinkedText {
    # Text a linked text box displays when TextFrame2 reports no text: legacy text API first, then
    # (only for links to a cell on Wkly_Report itself) that cell's displayed text.
    param($Shape, [string]$Name, [double]$Y)
    $formula = $null
    try { $formula = [string]$Shape.DrawingObject.Formula } catch { }
    $text = $null
    try { $text = [string]$Shape.TextFrame.Characters().Text } catch { }
    $linkPattern = '^=(?:''?(?<sheet>[^''!]+)''?!)?\$?(?<col>[A-Za-z]{1,3})\$?(?<row>\d+)$'
    if (-not $text -and $formula -match $linkPattern) {
        $sheet = $Matches['sheet']
        if (-not $sheet -or $sheet -eq $SheetName) {
            try { $text = [string]$script:XlSheet.Range($Matches['col'] + $Matches['row']).Text } catch { }
        }
        else {
            Add-Issue -Severity Warning -Location "Text box '$Name'" -Message ("Linked to {0} on another worksheet and Excel returned no text for it; that tab is not read, so the box is shown empty. See Excel view." -f $formula) -Y $Y
        }
    }
    if ($formula) { Write-Log ("  linked text box '{0}' {1} -> '{2}'" -f $Name, $formula, $text) -FileOnly }
    return $text
}

function Read-ShapeText {
    param($Shape, [string]$Name = '', [double]$Y = -1)
    $hasText = $false
    try { $hasText = Test-MsoTrue $Shape.TextFrame2.HasText } catch { }
    $tf = $null; $tr = $null
    try { $tf = $Shape.TextFrame2; $tr = $tf.TextRange } catch { }
    $runs = New-Object System.Collections.Generic.List[object]
    if ($hasText) {
        $cnt = 0
        try { $cnt = [int]$tr.Runs().Count } catch { }
        if ($cnt -gt 0) { for ($i = 1; $i -le $cnt; $i++) { $runs.Add((Get-RunInfo $tr.Runs($i, 1))) } }
        else { $runs.Add((Get-RunInfo $tr)) }
    }
    else {
        $linked = Get-LinkedText -Shape $Shape -Name $Name -Y $Y
        if (-not $linked) { return $null }
        $run = $null
        try { $run = Get-RunInfo $tr } catch { }
        if (-not $run) { $run = [pscustomobject]@{ Text = ''; Size = 11; Bold = $false; Italic = $false; Underline = $false; Color = $null; Font = $null } }
        $run.Text = $linked
        $runs.Add($run)
    }
    if (-not $tf) { return [pscustomobject]@{ Runs = $runs; Aligns = $null; Paras = $null; Anchor = 'top'; Margins = @(7.2, 3.6, 7.2, 3.6); Wrap = $true; Orientation = 1 } }
    $aligns = New-Object System.Collections.Generic.List[string]
    $paras = New-Object System.Collections.Generic.List[object]
    $map = @{ 1 = 'left'; 2 = 'center'; 3 = 'right'; 4 = 'justify'; 5 = 'justify'; 7 = 'justify' }
    try {
        $pc = [int]$tr.Paragraphs().Count
        $baseSize = if ($runs.Count -and $runs[0].Size -gt 0) { [double]$runs[0].Size } else { 11 }
        for ($p = 1; $p -le $pc; $p++) {
            $pf = $tr.Paragraphs($p, 1).ParagraphFormat
            $a = [int]$pf.Alignment
            $align = if ($map.ContainsKey($a)) { $map[$a] } else { 'left' }
            $aligns.Add($align)
            # Excel paragraph spacing: values are points, or lines when the matching LineRule is set
            $before = 0; $after = 0; $lh = $null
            try { $before = [double]$pf.SpaceBefore; if (Test-MsoTrue $pf.LineRuleBefore) { $before = $before * 1.2 * $baseSize } } catch { }
            try { $after = [double]$pf.SpaceAfter; if (Test-MsoTrue $pf.LineRuleAfter) { $after = $after * 1.2 * $baseSize } } catch { }
            try {
                $sw = [double]$pf.SpaceWithin
                if ($sw -gt 0) { $lh = if (Test-MsoTrue $pf.LineRuleWithin) { Format-Num (1.2 * $sw) } else { (Format-Num $sw) + 'pt' } }
            }
            catch { }
            $paras.Add([pscustomobject]@{ Align = $align; Before = [math]::Max(0, $before); After = [math]::Max(0, $after); LineHeight = $lh })
        }
    }
    catch { $aligns.Add('left') }
    $anchor = 'top'
    try { $anchor = switch ([int]$tf.VerticalAnchor) { 3 { 'middle' } 4 { 'bottom' } 5 { 'bottom' } default { 'top' } } } catch { }
    $m = @(7.2, 3.6, 7.2, 3.6)
    try { $m = @([double]$tf.MarginLeft, [double]$tf.MarginTop, [double]$tf.MarginRight, [double]$tf.MarginBottom) } catch { }
    $wrap = $true
    try { $wrap = Test-MsoTrue $tf.WordWrap } catch { }
    $orient = 1
    try { $orient = [int]$tf.Orientation } catch { }
    return [pscustomobject]@{ Runs = $runs; Aligns = $aligns; Paras = $paras; Anchor = $anchor; Margins = $m; Wrap = $wrap; Orientation = $orient }
}

function Export-ChartImage {
    # Returns $true when the image file was written; otherwise $false with the error text of every
    # method tried in $ErrorText. Every Excel call is assigned to $null: COM methods such as
    # Activate() return a value, and anything left uncaptured would become part of the result.
    # Some Excel builds only render a chart for export once its sheet/chart is active, or only
    # while the application window exists, so these are tried in turn on the hidden temp copy
    # (nothing is saved). The first method that works is reused for the remaining charts.
    param($Shape, [string]$Path, [ref]$ErrorText)
    $errors = New-Object System.Collections.Generic.List[string]
    $methods = @('direct', 'sheet-active', 'chart-activate', 'offscreen-window')
    if ($script:ChartMode) { $methods = @($script:ChartMode) + @($methods | Where-Object { $_ -ne $script:ChartMode }) }
    foreach ($method in $methods) {
        try {
            if ($method -eq 'offscreen-window' -and -not $script:XlMadeVisible) {
                $null = ($script:XlApp.Visible = $true)
                try { $script:XlApp.WindowState = -4143; $script:XlApp.Left = -30000; $script:XlApp.Top = -30000 } catch { }
                $script:XlMadeVisible = $true
                Write-Log 'Charts only export with a visible Excel window on this PC; the private Excel window (temp copy) was moved off-screen.' 'WARN'
            }
            if ($method -ne 'direct') { $null = $script:XlSheet.Activate() }
            if ($method -eq 'chart-activate' -or $method -eq 'offscreen-window') { try { $null = $Shape.Chart.Parent.Activate() } catch { } }
            if (Test-Path -LiteralPath $Path) { $null = Remove-Item -LiteralPath $Path -Force }
            $ret = $Shape.Chart.Export($Path, 'PNG', $false)
            if ((Test-Path -LiteralPath $Path) -and (Get-Item -LiteralPath $Path).Length -gt 0) {
                if ($script:ChartMode -ne $method) { $script:ChartMode = $method; Write-Log "Chart export method in use: $method" }
                $ErrorText.Value = $null
                return $true
            }
            $errors.Add(('{0}: Export returned {1} but wrote no file' -f $method, $ret))
        }
        catch { $errors.Add(('{0}: {1}' -f $method, $_.Exception.Message)) }
    }
    $ErrorText.Value = ($errors -join ' | ')
    return $false
}

function Export-ChartSharp {
    # Excel's own screen picture of the chart, taken with the hidden temp copy zoomed to 300%, so the
    # chart looks exactly as in Excel and stays sharp in the dashboard. $false (with the reason) when it
    # cannot be taken or is not sharper than Chart.Export; the caller then uses Chart.Export.
    param($Shape, [string]$Path, [double]$WidthPt, [double]$HeightPt, [ref]$ErrorText)
    try {
        if ([System.Threading.Thread]::CurrentThread.GetApartmentState() -ne 'STA') {
            $ErrorText.Value = 'clipboard access needs Windows PowerShell or PowerShell ISE (STA thread)'
            return $false
        }
        Add-Type -AssemblyName System.Windows.Forms, System.Drawing
        $null = Initialize-EmfHelper
        Save-ClipboardOnce
        # zoom is raised for this one picture only and put back straight away, so every other object
        # (linked pictures in particular) is read exactly as before
        $win = $null; $was = $null
        try {
            $null = $script:XlSheet.Activate()
            $win = $script:XlApp.ActiveWindow
            $was = $win.Zoom
            $win.Zoom = 300
            if (-not $script:ChartZoomSet) { $script:ChartZoomSet = $true; Write-Log ("Chart pictures: zoom of the temporary copy raised from {0} to 300% while each chart's picture is taken." -f $was) }
        }
        catch { if (-not $script:ChartZoomSet) { $script:ChartZoomSet = $true; Write-Log ("Chart pictures: zoom could not be raised ({0})." -f $_.Exception.Message) 'WARN' } }
        $err = $null
        try { $img = Get-ClipboardPicture -Target $Shape -Appearance 1 -Format 2 -ErrorText ([ref]$err) }   # xlScreen, xlBitmap
        finally { if ($win -and $null -ne $was) { try { $win.Zoom = $was } catch { } } }
        if (-not $img) { $ErrorText.Value = [string]$err; return $false }
        $sx = $img.Width / ($WidthPt * 96 / 72); $sy = $img.Height / ($HeightPt * 96 / 72)
        if ($sx -lt 1.3 -or [math]::Abs($sx - $sy) / $sx -gt 0.04) {
            $img.Dispose()
            $ErrorText.Value = ('picture came out at {0}x by {1}x the chart size' -f [math]::Round($sx, 2), [math]::Round($sy, 2))
            return $false
        }
        # a hidden Excel window can hand back an empty picture: one colour everywhere means blank
        $blank = $true
        try {
            $first = $img.GetPixel(0, 0).ToArgb()
            for ($gy = 1; $gy -le 9 -and $blank; $gy++) {
                for ($gx = 1; $gx -le 9; $gx++) {
                    if ($img.GetPixel([int]($img.Width * $gx / 10), [int]($img.Height * $gy / 10)).ToArgb() -ne $first) { $blank = $false; break }
                }
            }
        }
        catch { $blank = $false }
        if ($blank) { $img.Dispose(); $ErrorText.Value = 'picture was blank'; return $false }
        $script:ChartScales.Add([math]::Round($sx, 2))
        $img.Save($Path, ('System.Drawing.Imaging.ImageFormat' -as [type])::Png); $img.Dispose()
        $ErrorText.Value = $null
        return $true
    }
    catch { $ErrorText.Value = $_.Exception.Message; return $false }
}

function Get-ClipboardPicture {
    # Asks Excel for a picture of $Target (a range or shape) and reads it from the clipboard.
    # Another program can hold the clipboard for a moment, so every step waits and retries; the
    # clipboard change counter makes sure a previous picture is never taken for the new one.
    param($Target, [int]$Appearance, [int]$Format, [ref]$ErrorText)
    $cb = 'System.Windows.Forms.Clipboard' -as [type]
    $helper = 'ItsmDash.EmfClip4' -as [type]
    if ($script:TestClipboard) { $cb = $script:TestClipboard; $helper = $script:TestClipboard }   # test hook
    $last = 'no picture on the clipboard'
    for ($attempt = 1; $attempt -le 5; $attempt++) {
        $before = [uint32]0
        if ($helper) { try { $before = $helper::Seq() } catch { } }
        try { $null = $Target.CopyPicture($Appearance, $Format) }
        catch { $last = 'Excel copy: ' + $_.Exception.Message; Start-Sleep -Milliseconds (400 * $attempt); continue }
        if ($helper) {
            $waited = 0
            while ($waited -lt 30) {
                $now = $before
                try { $now = $helper::Seq() } catch { }
                if ($now -ne $before) { break }
                Start-Sleep -Milliseconds 100
                $waited = $waited + 1
            }
            if ($waited -ge 30) { $last = 'Excel did not put a new picture on the clipboard'; Start-Sleep -Milliseconds (400 * $attempt); continue }
        }
        if ($Format -ne 2) { $ErrorText.Value = $null; return 'copied' }     # vector: read by the helper
        $img = $null
        $tries = 0
        while ($tries -lt 20 -and $null -eq $img) {
            try { $img = $cb::GetImage() } catch { $last = 'clipboard read: ' + $_.Exception.Message }
            if ($null -eq $img) { Start-Sleep -Milliseconds 150 }
            $tries = $tries + 1
        }
        if ($img) { $ErrorText.Value = $null; return $img }
        Start-Sleep -Milliseconds (400 * $attempt)
    }
    $ErrorText.Value = $last
    return $null
}

function Save-ClipboardOnce {
    # Remember the user's clipboard text/picture once, before the first picture is copied.
    if ($null -ne $script:Clip) { return }
    $cb = 'System.Windows.Forms.Clipboard' -as [type]
    $saved = [pscustomobject]@{ Text = $null; Image = $null }
    try { if ($cb::ContainsText()) { $saved.Text = $cb::GetText() } } catch { }
    try { if (-not $saved.Text -and $cb::ContainsImage()) { $saved.Image = $cb::GetImage() } } catch { }
    $script:Clip = $saved
}

function Export-ShapeViaClipboard {
    # Camera/linked pictures and some objects have no image stored in the file. Excel renders them
    # with CopyPicture; the clipboard's text/picture is saved before the first use and restored at the end.
    param($Shape, [string]$Path, [ref]$ErrorText)
    try {
        if ([System.Threading.Thread]::CurrentThread.GetApartmentState() -ne 'STA') {
            $ErrorText.Value = 'clipboard access needs Windows PowerShell or PowerShell ISE (STA thread)'
            return $false
        }
        Add-Type -AssemblyName System.Windows.Forms, System.Drawing
        $null = Initialize-EmfHelper
        $png = ('System.Drawing.Imaging.ImageFormat' -as [type])::Png
        Save-ClipboardOnce
        $errs = New-Object System.Collections.Generic.List[string]
        foreach ($mode in @(1, 2)) {                     # xlScreen, then xlPrinter (both as bitmap)
            $err = $null
            $img = Get-ClipboardPicture -Target $Shape -Appearance $mode -Format 2 -ErrorText ([ref]$err)
            if ($img) {
                $img.Save($Path, $png); $img.Dispose()
                $ErrorText.Value = $null
                return $true
            }
            $errs.Add(('mode {0}: {1}' -f $mode, $err))
        }
        $ErrorText.Value = ($errs -join ' | ')
        return $false
    }
    catch { $ErrorText.Value = $_.Exception.Message; return $false }
}

function Initialize-EmfHelper {
    # Reads Excel's vector picture (EMF) from the clipboard and saves it as a sharp PNG.
    if ('ItsmDash.EmfClip4' -as [type]) { return $true }
    $code = @'
using System;
using System.Runtime.InteropServices;
using System.Drawing;
using System.Drawing.Imaging;
namespace ItsmDash {
  public static class EmfClip4 {
    [StructLayout(LayoutKind.Sequential)] struct RECT { public int Left, Top, Right, Bottom; }
    [DllImport("user32.dll")] static extern bool OpenClipboard(IntPtr h);
    [DllImport("user32.dll")] static extern bool CloseClipboard();
    [DllImport("user32.dll")] static extern IntPtr GetClipboardData(uint f);
    [DllImport("user32.dll")] static extern bool IsClipboardFormatAvailable(uint f);
    [DllImport("user32.dll")] static extern uint GetClipboardSequenceNumber();
    [DllImport("kernel32.dll")] static extern void Sleep(uint ms);
    public static uint Seq() { return GetClipboardSequenceNumber(); }
    [DllImport("gdi32.dll")] static extern IntPtr CopyEnhMetaFile(IntPtr h, IntPtr file);
    [DllImport("gdi32.dll")] static extern bool PlayEnhMetaFile(IntPtr hdc, IntPtr hemf, ref RECT r);
    [DllImport("gdi32.dll")] static extern bool DeleteEnhMetaFile(IntPtr h);
    public static string LastMethod = "";
    public static bool UseGdi = false;
    // Native GDI playback keeps Excel's exact per-letter positions (GDI+ re-lays out text and squeezes it).
    public static bool SavePng(string path, int width, int height) {
      LastMethod = "";
      if (!IsClipboardFormatAvailable(14)) return false;
      IntPtr copy = IntPtr.Zero;
      bool opened = false;
      for (int i = 0; i < 20 && !opened; i++) { opened = OpenClipboard(IntPtr.Zero); if (!opened) Sleep(150); }
      if (!opened) return false;
      try { IntPtr h = GetClipboardData(14); if (h != IntPtr.Zero) copy = CopyEnhMetaFile(h, IntPtr.Zero); }
      finally { CloseClipboard(); }
      if (copy == IntPtr.Zero) return false;
      bool played = false;
      if (UseGdi) try {
        using (Bitmap bmp = new Bitmap(width, height, PixelFormat.Format24bppRgb)) {
          using (Graphics g = Graphics.FromImage(bmp)) {
            g.Clear(Color.White);
            IntPtr hdc = g.GetHdc();
            try { RECT r = new RECT(); r.Left = 0; r.Top = 0; r.Right = width; r.Bottom = height; played = PlayEnhMetaFile(hdc, copy, ref r); }
            finally { g.ReleaseHdc(hdc); }
          }
          if (played) { bmp.Save(path, ImageFormat.Png); LastMethod = "GDI"; }
        }
      }
      catch { played = false; }
      if (played) { DeleteEnhMetaFile(copy); return true; }
      LastMethod = "GDI+";
      using (Metafile mf = new Metafile(copy, true))
      using (Bitmap bmp = new Bitmap(width, height))
      using (Graphics g = Graphics.FromImage(bmp)) {
        g.Clear(Color.White);
        g.SmoothingMode = System.Drawing.Drawing2D.SmoothingMode.HighQuality;
        g.InterpolationMode = System.Drawing.Drawing2D.InterpolationMode.HighQualityBicubic;
        g.TextRenderingHint = System.Drawing.Text.TextRenderingHint.AntiAliasGridFit;
        g.DrawImage(mf, new Rectangle(0, 0, width, height));
        bmp.Save(path, ImageFormat.Png);
      }
      return true;
    }
  }
}
'@
    try { Add-Type -TypeDefinition $code -ReferencedAssemblies System.Drawing -ErrorAction Stop; return $true }
    catch { Write-Log ("Vector page pictures unavailable ({0}); using bitmap pictures." -f $_.Exception.Message) 'WARN'; return $false }
}

function Export-RangePicture {
    # Excel's own picture of a cell range including every chart, shape and text box on it.
    # 1) print-quality vector picture rendered at 3x (sharp text, exact letter positions);
    # 2) Excel's screen image (exactly as on screen, but only at the monitor's resolution);
    # 3) screen-layout vector picture. Uses the clipboard (restored later).
    param($Range, [string]$Path, [double]$WidthPt, [double]$HeightPt, [ref]$ErrorText, [bool]$PrintQuality = $true)
    $errs = New-Object System.Collections.Generic.List[string]
    if (-not $script:TestClipboard) {
        if ([System.Threading.Thread]::CurrentThread.GetApartmentState() -ne 'STA') {
            $ErrorText.Value = 'clipboard access needs Windows PowerShell or PowerShell ISE (STA thread)'
            return $false
        }
        Add-Type -AssemblyName System.Windows.Forms, System.Drawing
        Save-ClipboardOnce
    }
    $wPx = [int][math]::Ceiling($WidthPt * 96 / 72 * $script:PictureScale)
    $hPx = [int][math]::Ceiling($HeightPt * 96 / 72 * $script:PictureScale)
    $emfType = 'ItsmDash.EmfClip4' -as [type]
    if ($script:TestClipboard) { $emfType = $script:TestClipboard }   # test hook
    $attempts = 0
    while ($attempts -lt 2) {
        $attempts = $attempts + 1
        # 1) print-quality vector picture (xlPrinter): letter positions at printer precision, no gridlines
        if ($PrintQuality -and $emfType) {
            try {
                $err = $null
                $copied = Get-ClipboardPicture -Target $Range -Appearance 2 -Format -4147 -ErrorText ([ref]$err)   # xlPrinter, xlPicture
                if ($copied -and $emfType::SavePng($Path, $wPx, $hPx)) {
                    $script:PictureMethods['print-quality picture'] = 1 + [int]$script:PictureMethods['print-quality picture']
                    $ErrorText.Value = $null
                    return $true
                }
                $why = 'picture could not be read'
                if ($err) { $why = [string]$err }
                $errs.Add('print-quality: ' + $why)
            }
            catch { $errs.Add('print-quality: ' + $_.Exception.Message) }
        }
        # 2) Excel's own screen image of the range (identical to what Excel shows; resolution follows
        #    the hidden window's zoom, which Add-PagePictures raises)
        try {
            $err = $null
            $img = Get-ClipboardPicture -Target $Range -Appearance 1 -Format 2 -ErrorText ([ref]$err)   # xlScreen, xlBitmap
            if ($img) {
                $oneX = $WidthPt * 96 / 72
                if ($oneX -gt 0) { $script:BitmapScales.Add([math]::Round($img.Width / $oneX, 2)) }
                $img.Save($Path, ('System.Drawing.Imaging.ImageFormat' -as [type])::Png); $img.Dispose()
                $script:PictureMethods['Excel screen image'] = 1 + [int]$script:PictureMethods['Excel screen image']
                $ErrorText.Value = $null
                return $true
            }
            $errs.Add('screen image: ' + $err)
        }
        catch { $errs.Add('screen image: ' + $_.Exception.Message) }
        # 3) fallback: screen-layout vector picture drawn with GDI+
        if ($emfType) {
            try {
                $err = $null
                $copied = Get-ClipboardPicture -Target $Range -Appearance 1 -Format -4147 -ErrorText ([ref]$err)   # xlScreen, xlPicture (vector)
                if ($copied -and $emfType::SavePng($Path, $wPx, $hPx)) {
                    $how = 'vector ' + [string]$emfType::LastMethod
                    $script:PictureMethods[$how] = 1 + [int]$script:PictureMethods[$how]
                    $ErrorText.Value = $null
                    return $true
                }
                $why = 'picture could not be read'
                if ($err) { $why = [string]$err }
                $errs.Add('vector: ' + $why)
            }
            catch { $errs.Add('vector: ' + $_.Exception.Message) }
        }
        # Some Excel builds only draw pictures while the application window exists: retry once off-screen.
        if ($script:XlMadeVisible) { break }
        try {
            $null = ($script:XlApp.Visible = $true)
            try { $script:XlApp.WindowState = -4143; $script:XlApp.Left = -30000; $script:XlApp.Top = -30000 } catch { }
            $script:XlMadeVisible = $true
            $null = $script:XlSheet.Activate()
        }
        catch { break }
    }
    $ErrorText.Value = ($errs -join ' | ')
    return $false
}

function Initialize-PdfRenderer {
    # Windows' built-in PDF renderer (Windows.Data.Pdf, Windows 10/11). Returns $true when usable.
    if ($script:PdfReady -ne $null) { return $script:PdfReady }
    $script:PdfReady = $false
    try {
        Add-Type -AssemblyName System.Runtime.WindowsRuntime
        $null = [Windows.Storage.StorageFile, Windows.Storage, ContentType = WindowsRuntime]
        $null = [Windows.Data.Pdf.PdfDocument, Windows.Data.Pdf, ContentType = WindowsRuntime]
        $null = [Windows.Data.Pdf.PdfPageRenderOptions, Windows.Data.Pdf, ContentType = WindowsRuntime]
        $null = [Windows.Storage.Streams.InMemoryRandomAccessStream, Windows.Storage.Streams, ContentType = WindowsRuntime]
        $methods = [System.WindowsRuntimeSystemExtensions].GetMethods()
        $script:AsTaskOp = $methods | Where-Object { $_.Name -eq 'AsTask' -and $_.GetParameters().Count -eq 1 -and $_.GetParameters()[0].ParameterType.Name -eq 'IAsyncOperation`1' } | Select-Object -First 1
        $script:AsTaskAction = $methods | Where-Object { $_.Name -eq 'AsTask' -and $_.GetParameters().Count -eq 1 -and $_.GetParameters()[0].ParameterType.Name -eq 'IAsyncAction' } | Select-Object -First 1
        if ($script:AsTaskOp -and $script:AsTaskAction) { $script:PdfReady = $true }
        else { Write-Log 'Windows PDF renderer not available; pages use Excel''s screen picture.' 'WARN' }
    }
    catch { Write-Log ("Windows PDF renderer not available ({0}); pages use Excel's screen picture." -f $_.Exception.Message) 'WARN' }
    return $script:PdfReady
}

function Wait-WinRt {
    param($Operation, [Type]$ResultType)
    if ($ResultType) {
        $task = $script:AsTaskOp.MakeGenericMethod($ResultType).Invoke($null, @($Operation))
        $null = $task.Wait(60000)
        return $task.Result
    }
    $task = $script:AsTaskAction.Invoke($null, @($Operation))
    $null = $task.Wait(60000)
}

function Set-PdfPageSetup {
    # One page per export, no margins/headers, centred - on the temporary copy only (never saved).
    param($Ws)
    try { $script:XlApp.PrintCommunication = $false } catch { }
    $ps = $Ws.PageSetup
    foreach ($prop in 'LeftMargin', 'RightMargin', 'TopMargin', 'BottomMargin', 'HeaderMargin', 'FooterMargin') { try { $ps.$prop = 0 } catch { } }
    foreach ($prop in 'LeftHeader', 'CenterHeader', 'RightHeader', 'LeftFooter', 'CenterFooter', 'RightFooter') { try { $ps.$prop = '' } catch { } }
    try { $ps.Zoom = $false; $ps.FitToPagesWide = 1; $ps.FitToPagesTall = 1 } catch { }
    try { $ps.CenterHorizontally = $true; $ps.CenterVertically = $true } catch { }
    try { $ps.PrintGridlines = $false; $ps.PrintHeadings = $false; $ps.BlackAndWhite = $false; $ps.Draft = $false } catch { }
    try { $script:XlApp.PrintCommunication = $true } catch { }
}

function Export-PagePdfPicture {
    # Excel exports the range to PDF (real, sharp text); Windows renders it to a high-resolution PNG,
    # cropped to the page content. Returns $true on success; $ErrorText says why not.
    param($Ws, $Range, [string]$PdfPath, [string]$PngPath, [double]$WidthPt, [double]$HeightPt, [ref]$ErrorText)
    try {
        try {
            $script:XlApp.PrintCommunication = $false
            if ($WidthPt -ge $HeightPt) { $Ws.PageSetup.Orientation = 2 } else { $Ws.PageSetup.Orientation = 1 }
            $script:XlApp.PrintCommunication = $true
        }
        catch { }
        if (Test-Path -LiteralPath $PdfPath) { Remove-Item -LiteralPath $PdfPath -Force }
        $null = $Range.ExportAsFixedFormat(0, $PdfPath, 0, $false, $true)
        if (-not (Test-Path -LiteralPath $PdfPath)) { $ErrorText.Value = 'Excel did not write the PDF'; return $false }
        $file = Wait-WinRt ([Windows.Storage.StorageFile]::GetFileFromPathAsync($PdfPath)) ([Windows.Storage.StorageFile])
        $doc = Wait-WinRt ([Windows.Data.Pdf.PdfDocument]::LoadFromFileAsync($file)) ([Windows.Data.Pdf.PdfDocument])
        if ([int]$doc.PageCount -ne 1) { $ErrorText.Value = ('Excel split the page into {0} PDF pages' -f $doc.PageCount); return $false }
        $page = $doc.GetPage(0)
        # render the whole PDF page (width only; Windows keeps the paper's proportions)
        $targetW = [double][math]::Round($WidthPt * 96 / 72 * $script:PdfScale)
        $opts = New-Object Windows.Data.Pdf.PdfPageRenderOptions
        $opts.DestinationWidth = [uint32][math]::Round($targetW * 1.4)
        $mem = New-Object Windows.Storage.Streams.InMemoryRandomAccessStream
        Wait-WinRt ($page.RenderToStreamAsync($mem, $opts)) $null
        try { $page.Dispose() } catch { }
        $in = [System.IO.WindowsRuntimeStreamExtensions]::AsStreamForRead($mem.GetInputStreamAt(0))
        $full = $null; $crop = $null
        try {
            Add-Type -AssemblyName System.Drawing
            $full = New-Object System.Drawing.Bitmap -ArgumentList $in
            $pw = [double]$full.Width; $ph = [double]$full.Height
            if ($pw -lt 10 -or $ph -lt 10) { $ErrorText.Value = 'rendered picture is empty'; return $false }
            # content is scaled to fit the paper and centred: full size in the limiting direction
            $aspect = $HeightPt / $WidthPt
            if ($ph / $pw -ge $aspect) { $cw = $pw; $ch = [math]::Min($ph, $pw * $aspect * 1.03) }
            else { $ch = $ph; $cw = [math]::Min($pw, $ph / $aspect * 1.03) }
            $rect = New-Object System.Drawing.Rectangle -ArgumentList ([int][math]::Floor(($pw - $cw) / 2)), ([int][math]::Floor(($ph - $ch) / 2)), ([int][math]::Floor($cw)), ([int][math]::Floor($ch))
            $crop = $full.Clone($rect, $full.PixelFormat)
            $crop.Save($PngPath, [System.Drawing.Imaging.ImageFormat]::Png)
        }
        finally { if ($crop) { $crop.Dispose() }; if ($full) { $full.Dispose() }; $in.Dispose() }
        if ((Get-Item -LiteralPath $PngPath).Length -lt 2000) { $ErrorText.Value = 'rendered picture is empty'; return $false }
        $ErrorText.Value = $null
        return $true
    }
    catch { $ErrorText.Value = $_.Exception.Message; return $false }
}

function Get-OverhangingObjects {
    # Objects (charts, shapes, pictures) that start on one page but run a little past its last row.
    # Their page is extended to show them whole, and they are hidden while the next page's picture
    # is taken. Objects taller than a page or overhanging by more than half a page (e.g. side bars,
    # page backgrounds) are left alone.
    param($Pages, $Ws)
    $list = New-Object System.Collections.Generic.List[object]
    $n = 0
    try { $n = [int]$Ws.Shapes.Count } catch { return , $list }
    for ($k = 1; $k -le $n; $k++) {
        $s = $null
        try { $s = $Ws.Shapes.Item($k) } catch { continue }
        if ($null -eq $s) { continue }
        $top = 0.0; $bottom = 0.0; $name = ''
        try {
            if (-not (Test-MsoTrue $s.Visible)) { continue }
            $top = [double]$s.Top; $bottom = $top + [double]$s.Height; $name = [string]$s.Name
        }
        catch { continue }
        $owner = $null
        foreach ($p in $Pages) { if ($top -ge $p.StartY -and $top -lt $p.EndY) { $owner = $p; break } }
        if ($null -eq $owner) { continue }
        $pageHeight = $owner.EndY - $owner.StartY
        $over = $bottom - $owner.EndY
        if ($over -le 1 -or $over -gt 0.5 * $pageHeight -or ($bottom - $top) -gt $pageHeight) { continue }
        $list.Add([pscustomobject]@{ Shape = $s; Name = $name; Top = $top; Bottom = $bottom; Page = $owner.Index })
    }
    return , $list
}

function Add-PagePictures {
    param($Pages, $Ws, $Geo, [string]$AssetsDir)
    $null = Initialize-EmfHelper
    # The pictures are taken "as shown on screen". The hidden copy's window shows cell gridlines and the
    # page-break lines Excel switches on after paginating, and both come out as dotted lines between
    # shapes. Hide them on this temporary copy only (it is never saved).
    $gridWas = 'unknown'; $breaksWas = 'unknown'
    try { $null = $Ws.Parent.Activate(); $null = $Ws.Activate() } catch { }
    try { $win = $script:XlApp.ActiveWindow; $gridWas = [string]$win.DisplayGridlines; $win.DisplayGridlines = $false } catch { }
    try { $breaksWas = [string]$Ws.DisplayPageBreaks; $Ws.DisplayPageBreaks = $false } catch { }
    $zoomWas = 'unknown'
    try { $win = $script:XlApp.ActiveWindow; $zoomWas = [string]$win.Zoom; $win.Zoom = 200 } catch { }
    Write-Log ("Page pictures: gridlines were {0}, page-break lines were {1}, zoom was {2}; gridlines/page breaks hidden and zoom set to 200% for the pictures (temporary copy only)." -f $gridWas, $breaksWas, $zoomWas)

    # keep charts/objects that run slightly past the end of their page on that page only
    $overhang = New-Object System.Collections.Generic.List[object]
    try { $overhang = Get-OverhangingObjects -Pages $Pages -Ws $Ws } catch { }
    foreach ($o in $overhang) {
        foreach ($p in $Pages) {
            if ($p.Index -ne $o.Page) { continue }
            if ($o.Bottom -gt $p.EndY) { $p.EndY = $o.Bottom; $p.EndRow = Get-RowAtY $Geo.RowTops $Geo.LastRow ($p.EndY - 0.5) }
        }
        Write-Log ("'{0}' runs past the end of page {1}: page {1} extended to show it whole; it is hidden from the next page's picture." -f $o.Name, $o.Page)
    }

    # Pages are taken from Excel's own PDF export (sharp text) when Windows can render PDFs;
    # otherwise, or if it fails for a page, Excel's screen picture is used exactly as in r16.
    # A selected chart would make Excel export only that chart, so a cell is selected first.
    $printOk = Initialize-PdfRenderer
    if ($printOk) {
        try { $null = $Ws.Activate(); $null = $Ws.Range('A1').Select() } catch { }
        try { Set-PdfPageSetup -Ws $Ws } catch { Write-Log ("Page setup for PDF pages failed: {0}" -f $_.Exception.Message) 'WARN' }
    }
    $pdfDir = Join-Path ([System.IO.Path]::GetTempPath()) ('ITSMDash_pdf_' + [guid]::NewGuid().ToString('N'))
    if ($printOk) { New-Item -ItemType Directory -Path $pdfDir -Force | Out-Null }
    $noPrint = New-Object System.Collections.Generic.List[object]
    if ($printOk) {
        try {
            $n = [int]$Ws.Shapes.Count
            for ($k = 1; $k -le $n; $k++) {
                try {
                    $s = $Ws.Shapes.Item($k)
                    if ((Test-MsoTrue $s.Visible) -and ($s.DrawingObject.PrintObject -eq $false)) {
                        $noPrint.Add([pscustomobject]@{ Name = [string]$s.Name; Top = [double]$s.Top; Bottom = [double]$s.Top + [double]$s.Height })
                    }
                }
                catch { }
            }
        }
        catch { }
    }

    $ok = 0
    foreach ($p in $Pages) {
        $usePrint = $printOk
        foreach ($np in $noPrint) {
            if ($np.Bottom -gt $p.StartY -and $np.Top -lt $p.EndY) {
                $usePrint = $false
                Write-Log ("Page {0}: '{1}' is set not to print (it would be missing from the PDF), so this page uses Excel's screen picture." -f $p.Index, $np.Name)
                break
            }
        }
        $hidden = New-Object System.Collections.Generic.List[object]
        foreach ($o in $overhang) {
            if ($o.Page -ne $p.Index -and $o.Bottom -gt $p.StartY -and $o.Top -lt $p.EndY) {
                try { $o.Shape.Visible = 0; $hidden.Add($o) } catch { }
            }
        }
        try {
            $c1 = 1
            for ($c = 1; $c -le $Geo.LastCol; $c++) { if ($Geo.ColLefts[$c] -le $p.OffsetX + 0.5) { $c1 = $c } }
            $right = $p.OffsetX + $p.Width - 0.5
            $c2 = $c1
            for ($c = $c1; $c -le $Geo.LastCol; $c++) { if ($Geo.ColLefts[$c] -lt $right) { $c2 = $c } }
            $L = $Geo.ColLefts[$c1]; $T = $Geo.RowTops[$p.StartRow]
            $W = $Geo.ColLefts[$c2 + 1] - $L; $H = $Geo.RowTops[$p.EndRow + 1] - $T
            $addr = (ConvertTo-ColumnLetter $c1) + [string]$p.StartRow + ':' + (ConvertTo-ColumnLetter $c2) + [string]$p.EndRow
            $file = 'page-{0:D2}.png' -f $p.Index
            $err = $null
            $done = $false
            if ($usePrint) {
                $done = (Export-PagePdfPicture -Ws $Ws -Range $Ws.Range($addr) -PdfPath (Join-Path $pdfDir ('page-{0:D2}.pdf' -f $p.Index)) -PngPath (Join-Path $AssetsDir $file) -WidthPt $W -HeightPt $H -ErrorText ([ref]$err)) -eq $true
                if ($done) { $script:PictureMethods['Excel PDF'] = 1 + [int]$script:PictureMethods['Excel PDF'] }
                else { Write-Log ("Page {0}: PDF picture not used ({1}); using Excel's screen picture." -f $p.Index, $err) 'WARN' }
            }
            if (-not $done) { $done = (Export-RangePicture -Range $Ws.Range($addr) -Path (Join-Path $AssetsDir $file) -WidthPt $W -HeightPt $H -ErrorText ([ref]$err) -PrintQuality $false) -eq $true }
            if (-not $done) {
                Write-Log ("Page {0}: picture not captured yet ({1}); trying again in 3 seconds." -f $p.Index, $err) 'WARN'
                Start-Sleep -Seconds 3
                $done = (Export-RangePicture -Range $Ws.Range($addr) -Path (Join-Path $AssetsDir $file) -WidthPt $W -HeightPt $H -ErrorText ([ref]$err) -PrintQuality $false) -eq $true
            }
            if ($done) {
                $p.Picture = [pscustomobject]@{ Kind = 'image'; L = $L; T = $T; W = $W; H = $H; Z = 0; Rot = 0; Src = "assets/$file"; Alt = ('Wkly_Report ' + $addr); IsChart = $false }
                $ok = $ok + 1
            }
            else {
                Add-Issue -Severity Warning -Location ("Page {0}" -f $p.Index) -Message ("Excel's picture of {0} could not be captured [{1}]; the re-drawn view is shown for this page." -f $addr, $err) -Y $p.StartY
            }
        }
        catch { Add-Issue -Severity Warning -Location ("Page {0}" -f $p.Index) -Message ("Page picture failed; the re-drawn view is shown: {0}" -f (Format-ErrorDetail $_)) -Y $p.StartY }
        finally { foreach ($o in $hidden) { try { $o.Shape.Visible = -1 } catch { } } }
    }
    if ($printOk) { Remove-Item -LiteralPath $pdfDir -Recurse -Force -ErrorAction SilentlyContinue }
    $how = ($script:PictureMethods.Keys | ForEach-Object { '{0}: {1}' -f $_, $script:PictureMethods[$_] }) -join ', '
    $res = ''
    if ($script:BitmapScales.Count) { $res = '; screen image resolution ' + (($script:BitmapScales | ForEach-Object { Format-Num $_ }) -join ', ') + 'x' }
    Write-Log ("Excel page pictures: {0} of {1} pages captured ({2}{3})." -f $ok, $Pages.Count, $how, $res) 'OK'
    return $ok
}

function Restore-Clipboard {
    if ($null -eq $script:Clip) { return }
    $cb = 'System.Windows.Forms.Clipboard' -as [type]
    $done = $false; $last = ''
    for ($i = 1; $i -le 10 -and -not $done; $i++) {
        try {
            if ($script:Clip.Text) { $cb::SetText($script:Clip.Text) }
            elseif ($script:Clip.Image) { $cb::SetImage($script:Clip.Image) }
            $done = $true
        }
        catch { $last = $_.Exception.Message; Start-Sleep -Milliseconds 200 }
    }
    if ($done) { Write-Log 'Clipboard restored to what it held before the run.' }
    else { Write-Log ("Clipboard could not be restored: {0}" -f $last) 'WARN' }
    $script:Clip = $null
}

function Add-ObjectPicture {
    # Adds Excel's own picture of an object, or a labelled placeholder plus a Data check entry.
    param($Shape, [string]$Name, [string]$What, $Items, [string]$AssetsDir, [double]$L, [double]$T, [double]$W, [double]$H, [double]$Z, [double]$Rot)
    $script:PictureIndex++
    $file = 'object-{0:D3}.png' -f $script:PictureIndex
    $err = $null
    $ok = (Export-ShapeViaClipboard -Shape $Shape -Path (Join-Path $AssetsDir $file) -ErrorText ([ref]$err)) -eq $true
    if (-not $ok) {
        # second try at 100% zoom with the worksheet active (how Excel shows it normally)
        $zoomWas = $null; $win = $null
        try { $null = $script:XlSheet.Activate(); $win = $script:XlApp.ActiveWindow; $zoomWas = $win.Zoom; $win.Zoom = 100 } catch { }
        $err2 = $null
        $ok = (Export-ShapeViaClipboard -Shape $Shape -Path (Join-Path $AssetsDir $file) -ErrorText ([ref]$err2)) -eq $true
        if ($win -and $null -ne $zoomWas) { try { $win.Zoom = $zoomWas } catch { } }
        if ($ok) { Write-Log ("{0} '{1}': first capture failed [{2}]; captured on the second try at 100% zoom." -f $What, $Name, $err) 'WARN' }
        else { $err = '{0} | retry at 100% zoom: {1}' -f $err, $err2 }
    }
    if ($ok) {
        $Items.Add([pscustomobject]@{ Kind = 'image'; L = $L; T = $T; W = $W; H = $H; Z = $Z; Rot = $Rot; Src = "assets/$file"; Alt = $Name; IsChart = $false })
        Write-Log ("{0} '{1}' captured as Excel's own picture." -f $What, $Name)
        return
    }
    $Items.Add([pscustomobject]@{ Kind = 'missing'; L = $L; T = $T; W = $W; H = $H; Z = $Z; Rot = $Rot; Message = "$What '$Name' not reproduced - see Excel view" })
    Add-Issue -Severity Warning -Location "$What '$Name'" -Message ("Could not be reproduced [{0}]. Placeholder shown; see Excel view." -f $err) -Y $T
}

function Read-OneShape {
    param($Shape, [double]$Z, $Items, [string]$AssetsDir, [hashtable]$MediaMap)
    $name = '(unnamed)'
    try { $name = [string]$Shape.Name } catch { }
    try {
        if (-not (Test-MsoTrue $Shape.Visible)) { return }
        $L = [double]$Shape.Left; $T = [double]$Shape.Top; $W = [double]$Shape.Width; $H = [double]$Shape.Height
        $type = [int]$Shape.Type
        $rot = 0; try { $rot = [double]$Shape.Rotation } catch { }

        if ($type -eq 6) {
            $gi = $Shape.GroupItems
            for ($i = 1; $i -le $gi.Count; $i++) { Read-OneShape -Shape $gi.Item($i) -Z ($Z + $i * 0.001) -Items $Items -AssetsDir $AssetsDir -MediaMap $MediaMap }
            return
        }

        $hasChart = $false
        try { $hasChart = Test-MsoTrue $Shape.HasChart } catch { }
        if ($type -eq 3 -or $hasChart) {
            $script:ChartIndex++
            $file = 'chart-{0:D3}.png' -f $script:ChartIndex
            $full = Join-Path $AssetsDir $file
            $title = $name
            try { if (Test-MsoTrue $Shape.Chart.HasTitle) { $title = [string]$Shape.Chart.ChartTitle.Text } } catch { }
            $exportError = $null
            $exported = $false
            if ($ChartRender -eq 'Excel' -and $script:ChartShots.fails -lt 2) {
                $sharpError = $null
                $exported = (Export-ChartSharp -Shape $Shape -Path $full -WidthPt $W -HeightPt $H -ErrorText ([ref]$sharpError)) -eq $true
                if ($exported) { $script:ChartShots.sharp++ }
                else {
                    if ($script:ChartShots.sharp -eq 0) { $script:ChartShots.fails++ }   # stops trying after two failures in a row at the start
                    if (-not $script:ChartShots.why) { $script:ChartShots.why = "'$name': $sharpError" }
                }
            }
            if (-not $exported) {
                $exported = (Export-ChartImage -Shape $Shape -Path $full -ErrorText ([ref]$exportError)) -eq $true
                if ($exported) { $script:ChartShots.export++ }
            }
            $spec = $null
            try { $spec = Get-ChartSpec $Shape } catch { Write-Log ("Chart '{0}': data could not be read ({1}); Excel's chart image is used." -f $name, $_.Exception.Message) 'WARN' }
            if ($ChartRender -eq 'Native' -and $spec -and -not $spec.kind) { Write-Log ("Chart '{0}': chart type {1} is not rebuilt natively; Excel's chart image is used." -f $name, $spec.excelType) }
            if ($exported) {
                $Items.Add([pscustomobject]@{ Kind = 'image'; L = $L; T = $T; W = $W; H = $H; Z = $Z; Rot = $rot; Src = "assets/$file"; Alt = $title; IsChart = $true; Spec = $spec })
                if ((Get-Item -LiteralPath $full).Length -lt 1500) {
                    Add-Issue -Severity Warning -Location "Chart '$name'" -Message 'Exported chart image is very small and may be blank - compare with Excel view.' -Y $T
                }
            }
            elseif ($spec -and $spec.kind) {
                $Items.Add([pscustomobject]@{ Kind = 'image'; L = $L; T = $T; W = $W; H = $H; Z = $Z; Rot = $rot; Src = $null; Alt = $title; IsChart = $true; Spec = $spec })
            }
            else {
                $Items.Add([pscustomobject]@{ Kind = 'missing'; L = $L; T = $T; W = $W; H = $H; Z = $Z; Rot = $rot; Message = "Chart '$title' could not be exported - see Excel view" })
                Add-Issue -Severity Error -Location "Chart '$name'" -Message ("Excel could not export this chart as an image [{0}]. A labelled placeholder is shown instead; see Excel view." -f $exportError) -Y $T
            }
            return
        }

        if ($type -in 11, 13, 28) {
            if ($MediaMap.ContainsKey($name)) {
                $Items.Add([pscustomobject]@{ Kind = 'image'; L = $L; T = $T; W = $W; H = $H; Z = $Z; Rot = $rot; Src = $MediaMap[$name]; Alt = $name; IsChart = $false })
                try {
                    $pf = $Shape.PictureFormat
                    if (([double]$pf.CropLeft + [double]$pf.CropTop + [double]$pf.CropRight + [double]$pf.CropBottom) -gt 0.5) {
                        Add-Issue -Severity Info -Location "Picture '$name'" -Message 'Picture is cropped in Excel; the dashboard shows the full picture.' -Y $T
                    }
                }
                catch { }
            }
            else {
                # linked/camera picture (or .xlsb/.xls): nothing stored in the file, ask Excel for its picture
                Add-ObjectPicture -Shape $Shape -Name $name -What 'Picture' -Items $Items -AssetsDir $AssetsDir -L $L -T $T -W $W -H $H -Z $Z -Rot $rot
            }
            return
        }

        $isConnector = $false
        try { $isConnector = Test-MsoTrue $Shape.Connector } catch { }
        if ($type -eq 9 -or $isConnector) {
            $color = '#000'; $weight = 0.75
            try { $color = ConvertTo-CssColor $Shape.Line.ForeColor.RGB (1 - [double]$Shape.Line.Transparency); $weight = [double]$Shape.Line.Weight } catch { }
            $fh = $false; $fv = $false
            try { $fh = Test-MsoTrue $Shape.HorizontalFlip; $fv = Test-MsoTrue $Shape.VerticalFlip } catch { }
            $Items.Add([pscustomobject]@{ Kind = 'line'; L = $L; T = $T; W = $W; H = $H; Z = $Z; Rot = $rot; Color = $color; Weight = $weight; FlipH = $fh; FlipV = $fv })
            return
        }

        if ($type -in 1, 5, 15, 17, 14, 20) {
            $bg = Get-ShapeBackground $Shape
            $border = $null
            try { if (Test-MsoTrue $Shape.Line.Visible) { $border = '{0}pt solid {1}' -f (Format-Num ([double]$Shape.Line.Weight)), (ConvertTo-CssColor $Shape.Line.ForeColor.RGB (1 - [double]$Shape.Line.Transparency)) } } catch { }
            $radius = '0'
            $ast = 1
            try { $ast = [int]$Shape.AutoShapeType } catch { }
            switch ($ast) {
                5 {
                    $adj = 0.16667
                    try { $adj = [double]$Shape.Adjustments.Item(1) } catch { }
                    $radius = (Format-Num ($adj * [math]::Min($W, $H))) + 'pt'
                }
                9 { $radius = '50%' }
                1 { }
                default {
                    if (($bg -or $border) -and -not $script:LoggedShapeTypes.ContainsKey($ast)) {
                        $script:LoggedShapeTypes[$ast] = $true
                        Add-Issue -Severity Info -Location "Shape '$name'" -Message ("Shape outline type {0} is drawn as a rectangle (text and colours are kept)." -f $ast) -Y $T
                    }
                }
            }
            $shadow = $null
            try {
                $sh = $Shape.Shadow
                if (Test-MsoTrue $sh.Visible) {
                    $shadow = '{0}pt {1}pt {2}pt {3}' -f (Format-Num ([double]$sh.OffsetX)), (Format-Num ([double]$sh.OffsetY)), (Format-Num ([double]$sh.Blur)), (ConvertTo-CssColor $sh.ForeColor.RGB (1 - [double]$sh.Transparency))
                }
            }
            catch { }
            $text = Read-ShapeText -Shape $Shape -Name $name -Y $T
            if (-not $bg -and -not $border -and -not $text) { return }   # invisible helper shape
            $formula = $null; $autoSize = $null
            try { $formula = [string]$Shape.DrawingObject.Formula } catch { }
            try { $autoSize = [int]$Shape.TextFrame2.AutoSize } catch { }
            $obj = [pscustomobject]@{
                Kind = 'shape'; L = $L; T = $T; W = $W; H = $H; Z = $Z; Rot = $rot; Name = $name; Formula = $formula; AutoSize = $autoSize; ShapeType = $ast
                Bg = $bg; Border = $border; Radius = $radius; Shadow = $shadow
                Runs = $null; Aligns = $null; Paras = $null; Anchor = 'top'; Margins = $null; Wrap = $true; TextOrientation = 1
            }
            if ($text) {
                # assigned directly so single-item lists are not unrolled into scalars
                $obj.Runs = $text.Runs; $obj.Aligns = $text.Aligns; $obj.Paras = $text.Paras; $obj.Anchor = $text.Anchor
                $obj.Margins = $text.Margins; $obj.Wrap = $text.Wrap; $obj.TextOrientation = $text.Orientation
            }
            $Items.Add($obj)
            if ($text) {
                $joined = ($text.Runs | ForEach-Object { $_.Text }) -join ''
                foreach ($m in [regex]::Matches($joined, $ErrRegex)) {
                    Add-Issue -Severity Error -Location "Shape '$name'" -Message ("Excel error value '{0}' in text box is shown exactly as Excel displays it." -f $m.Value) -Y $T
                }
            }
            return
        }

        $typeNames = @{ 2 = 'callout'; 7 = 'embedded OLE object'; 8 = 'form control'; 10 = 'linked OLE object'; 12 = 'ActiveX control'; 21 = 'ink'; 24 = 'SmartArt'; 25 = 'slicer'; 26 = 'web video'; 29 = '3D model' }
        $tn = if ($typeNames.ContainsKey($type)) { $typeNames[$type] } else { "type $type" }
        Add-ObjectPicture -Shape $Shape -Name $name -What $tn -Items $Items -AssetsDir $AssetsDir -L $L -T $T -W $W -H $H -Z $Z -Rot 0
    }
    catch {
        Add-Issue -Severity Warning -Location "Object '$name'" -Message ("Could not be read: {0}" -f $_.Exception.Message)
    }
}

function Read-ShapeItems {
    # Each object is read on its own: a failure is recorded and skipped, never fatal.
    param($Ws, [string]$AssetsDir, [hashtable]$MediaMap, $Items)
    $shapes = $Ws.Shapes
    $n = [int]$shapes.Count
    $failedInRow = 0
    for ($k = 1; $k -le $n; $k++) {
        Write-Progress -Activity 'Reading Wkly_Report shapes and charts' -PercentComplete ([int](100 * $k / [math]::Max(1, $n)))
        $s = $null
        try { $s = $shapes.Item($k) } catch { }
        if ($null -eq $s) {
            # Excel can hand back an empty collection reference after a chart was activated; fetch it again once.
            try { $shapes = $Ws.Shapes; $s = $shapes.Item($k) }
            catch { Add-Issue -Severity Warning -Location "Object #$k" -Message ("Could not be read and was skipped: {0}" -f (Format-ErrorDetail $_)) }
        }
        if ($null -eq $s) {
            $failedInRow++
            if ($failedInRow -ge 10) {
                Add-Issue -Severity Error -Location 'Wkly_Report objects' -Message ("Excel stopped answering after object {0} of {1}; the remaining objects are missing from the dashboard (see Excel view)." -f ($k - 10), $n)
                break
            }
            continue
        }
        $failedInRow = 0
        $zp = $k
        try { $zp = [int]$s.ZOrderPosition } catch { }
        try { Read-OneShape -Shape $s -Z (10 + $zp) -Items $Items -AssetsDir $AssetsDir -MediaMap $MediaMap }
        catch { Add-Issue -Severity Warning -Location "Object #$k" -Message ("Could not be read and was skipped: {0}" -f (Format-ErrorDetail $_)) }
    }
    Write-Progress -Activity 'Reading Wkly_Report shapes and charts' -Completed
    Write-Log ("Read {0} shapes ({1} charts exported)." -f $n, $script:ChartIndex)
    if ($ChartRender -eq 'Excel' -and $script:ChartIndex -gt 0) {
        $res = ''; if ($script:ChartScales.Count) { $res = ' at ' + (Format-Num (($script:ChartScales | Measure-Object -Minimum).Minimum)) + 'x resolution or better' }
        Write-Log ("Charts shown as Excel's own picture: {0} sharp screen picture(s){1}, {2} via Chart.Export (normal resolution)." -f $script:ChartShots.sharp, $res, $script:ChartShots.export)
        if ($script:ChartShots.export -gt 0 -and $script:ChartShots.why) { Write-Log ("Sharp chart picture not available, first reason: {0}" -f $script:ChartShots.why) 'WARN' }
    }
}

# ======================================================================================
#  Page detection (pure logic - no Excel calls)
# ======================================================================================

function Get-CssLuminance {
    param([string]$Css)
    if (-not $Css) { return 1.0 }
    if ($Css -match '#([0-9a-fA-F]{2})([0-9a-fA-F]{2})([0-9a-fA-F]{2})') { $r = [Convert]::ToInt32($Matches[1], 16); $g = [Convert]::ToInt32($Matches[2], 16); $b = [Convert]::ToInt32($Matches[3], 16) }
    elseif ($Css -match 'rgba?\((\d+),(\d+),(\d+)') { $r = [int]$Matches[1]; $g = [int]$Matches[2]; $b = [int]$Matches[3] }
    else { return 1.0 }
    return (0.2126 * $r + 0.7152 * $g + 0.0722 * $b) / 255
}

function Get-DarkBand {
    # A page-wide dark band/panel (navy header, dark cover) whose top is just above the text.
    param($It, $Items, [double]$ContentWidth)
    $px = $It.L + [math]::Min(10, $It.W / 2); $py = $It.T + [math]::Min(6, $It.H / 2)
    foreach ($c in $Items) {
        if (($c.Kind -ne 'shape' -and $c.Kind -ne 'fill') -or -not $c.Bg) { continue }
        if ($c.W -lt 0.45 * $ContentWidth) { continue }
        if ($c.L -gt $px -or ($c.L + $c.W) -lt $px -or $c.T -gt $py -or ($c.T + $c.H) -lt $py) { continue }
        if ($It.T - $c.T -gt 80) { continue }
        if ((Get-CssLuminance $c.Bg) -lt 0.35) { return $c }
    }
    return $null
}

function Repair-InvisibleText {
    # Text drawn in (almost) the same colour as what is behind it cannot be read. Excel shows it, so the
    # colour Excel reported is incomplete (e.g. a theme colour with brightness). Make it readable and say so.
    param($Items)
    $surfaces = @($Items | Where-Object { $_.Bg -and $_.Kind -in 'fill', 'cell', 'shape' })
    $count = 0
    foreach ($it in $Items) {
        if ($it.Kind -ne 'shape' -or -not $it.Runs) { continue }
        $bg = $it.Bg
        if (-not $bg) {
            $cx = $it.L + $it.W / 2; $cy = $it.T + $it.H / 2; $bestZ = -1
            foreach ($s in $surfaces) {
                if ($s -eq $it -or $s.Z -ge $it.Z -or $s.Z -le $bestZ) { continue }
                if ($cx -lt $s.L -or $cx -gt $s.L + $s.W -or $cy -lt $s.T -or $cy -gt $s.T + $s.H) { continue }
                $bg = $s.Bg; $bestZ = $s.Z
            }
        }
        if (-not $bg) { $bg = '#ffffff' }
        $bl = Get-CssLuminance $bg
        foreach ($r in $it.Runs) {
            if (-not $r.Color -or ([string]$r.Text) -notmatch '\S') { continue }
            $tl = Get-CssLuminance $r.Color
            $ratio = ([math]::Max($tl, $bl) + 0.05) / ([math]::Min($tl, $bl) + 0.05)
            if ($ratio -ge 1.6) { continue }
            $r | Add-Member -NotePropertyName ColorAdjustedFrom -NotePropertyValue $r.Color -Force
            $r.Color = if ($bl -gt 0.5) { '#1b2333' } else { '#ffffff' }
            $count++
            $t = ([string]$r.Text).Trim(); if ($t.Length -gt 40) { $t = $t.Substring(0, 37) + '...' }
            Add-Issue -Severity Info -Location "Text box '$($it.Name)'" -Message ("Text '{0}' had colour {1} on background {2} (unreadable; Excel reported {3}). Shown in {4} so it can be read - the text itself is unchanged." -f $t, $r.ColorAdjustedFrom, $bg, $r.ColorInfo, $r.Color) -Y $it.T
        }
    }
    if ($count) { Write-Log ("{0} text run(s) had the same colour as their background and were made readable (listed in Data check)." -f $count) 'WARN' }
}

function Find-OverlappingText {
    param($Items)
    $texts = @($Items | Where-Object { $_.Kind -eq 'shape' -and $_.Runs -and (($_.Runs | ForEach-Object { $_.Text }) -join '') -match '\S' })
    $overlaps = 0
    for ($i = 0; $i -lt $texts.Count; $i++) {
        for ($j = $i + 1; $j -lt $texts.Count; $j++) {
            $a = $texts[$i]; $b = $texts[$j]
            $ox = [math]::Min($a.L + $a.W, $b.L + $b.W) - [math]::Max($a.L, $b.L)
            $oy = [math]::Min($a.T + $a.H, $b.T + $b.H) - [math]::Max($a.T, $b.T)
            if ($ox -le 0 -or $oy -le 0) { continue }
            $small = [math]::Min($a.W * $a.H, $b.W * $b.H)
            if ($small -le 0 -or ($ox * $oy) / $small -lt 0.2) { continue }
            $script:Issues.Add([pscustomobject]@{ Severity = 'Info'; Location = ("Text boxes '{0}' / '{1}'" -f $a.Name, $b.Name); Message = 'These text boxes overlap in Excel. If their text collides in the dashboard, turn on Inspect, click each one and send a photo of the panel.'; Y = [math]::Min($a.T, $b.T); Page = $null })
            $overlaps++
        }
    }
    if ($overlaps) { Write-Log ("{0} pairs of overlapping text boxes found (listed in Data check)." -f $overlaps) }
}

function Get-TitleCandidates {
    # A page title is short text on the left of the sheet that is either very large (>= MinFont)
    # or >= 16pt and sitting at the top of a page-wide dark header band.
    param($Items, [double]$ContentWidth, [double]$MinFont)
    $list = New-Object System.Collections.Generic.List[object]
    foreach ($it in $Items) {
        if ($it.Kind -ne 'cell' -and $it.Kind -ne 'shape') { continue }
        $size = Get-ItemMaxFont $it
        if ($size -lt [math]::Min(16, $MinFont)) { continue }
        $t = ((Get-ItemText $it) -replace '\s+', ' ').Trim()
        if ($t.Length -eq 0 -or $t.Length -gt 120) { continue }
        if ([regex]::Matches($t, '\p{L}').Count -lt 4) { continue }
        if ($t -match $ErrRegex) { continue }
        if ($it.L -gt 0.35 * $ContentWidth) { continue }
        $reason = $null
        if ($size -ge $MinFont) { $reason = 'large font' }
        elseif (Get-DarkBand -It $it -Items $Items -ContentWidth $ContentWidth) { $reason = 'header band' }
        if (-not $reason) { continue }
        $list.Add([pscustomobject]@{ Text = $t; Top = [double]$it.T; Item = $it; Size = $size; Reason = $reason })
    }
    return @($list | Sort-Object Top, { $_.Item.L })
}

function Get-PageBoundary {
    # The page starts at the top of the coloured band/panel that holds its title.
    param($Candidate, $Items, [double]$MinY)
    $px = $Candidate.Item.L + [math]::Min(10, $Candidate.Item.W / 2)
    $py = $Candidate.Top + [math]::Min(6, $Candidate.Item.H / 2)
    $best = $null
    foreach ($it in $Items) {
        if ($it.Kind -ne 'shape' -and $it.Kind -ne 'fill') { continue }
        if (-not $it.Bg -or $it.W -lt 40) { continue }
        if ($it.L -gt $px -or ($it.L + $it.W) -lt $px -or $it.T -gt $py -or ($it.T + $it.H) -lt $py) { continue }
        if ($it.T -lt $MinY -or $it.T -lt $Candidate.Top - 150) { continue }
        if ($null -eq $best -or $it.T -lt $best) { $best = [double]$it.T }
    }
    if ($null -eq $best) { $best = [math]::Max($MinY, $Candidate.Top - 4) }
    return $best
}

function Get-PageTitle {
    param($Items, [double]$StartY, [double]$EndY, [int]$Index)
    $limit = $StartY + ($EndY - $StartY) * 0.4
    $best = $null; $bestSize = 0
    foreach ($it in $Items) {
        if ($it.Kind -ne 'cell' -and $it.Kind -ne 'shape') { continue }
        if ($it.T -lt $StartY -or $it.T -gt $limit) { continue }
        $t = ((Get-ItemText $it) -split "[`r`n`v]" | Where-Object { $_.Trim() } | Select-Object -First 1)
        if (-not $t -or [regex]::Matches($t, '\p{L}').Count -lt 4) { continue }
        $s = Get-ItemMaxFont $it
        if ($s -gt $bestSize -or ($s -eq $bestSize -and $best -and $it.T -lt $best.T)) { $best = $it; $bestSize = $s; $best | Add-Member -NotePropertyName TitleText -NotePropertyValue $t.Trim() -Force }
    }
    if ($best) { $t = $best.TitleText; if ($t.Length -gt 80) { $t = $t.Substring(0, 77) + '...' }; return $t }
    return "Page $Index"
}

function Find-ReportPages {
    param(
        $Items, $Geo, [double]$ContentWidth, [double]$ContentBottom,
        [string]$Mode, [int[]]$StartRows, [int[]]$BreakRows,
        [double]$MinFont, [double]$MinHeight
    )
    $starts = New-Object System.Collections.Generic.List[object]   # @{Y; Title}
    $method = $null

    if ($Mode -eq 'Manual') {
        if (-not $StartRows -or $StartRows.Count -eq 0) { throw '-PageMode Manual needs -PageStartRows, e.g. -PageStartRows 1,40,72' }
        $rows = @($StartRows | Where-Object { $_ -ge 1 -and $_ -le $Geo.LastRow } | Sort-Object -Unique)
        if ($rows[0] -ne 1) { $rows = @(1) + $rows }
        foreach ($r in $rows) { $starts.Add([pscustomobject]@{ Y = $Geo.RowTops[$r]; Title = $null }) }
        $method = 'Manual page start rows: ' + ($rows -join ', ')
    }

    if (-not $method -and ($Mode -eq 'Auto' -or $Mode -eq 'Titles')) {
        $cands = Get-TitleCandidates -Items $Items -ContentWidth $ContentWidth -MinFont $MinFont
        $cur = [pscustomobject]@{ Y = 0.0; Title = $null; TitleTop = $null }
        $starts.Add($cur)
        foreach ($c in $cands) {
            $row = Get-RowAtY $Geo.RowTops $Geo.LastRow $c.Top
            $short = if ($c.Text.Length -gt 50) { $c.Text.Substring(0, 47) + '...' } else { $c.Text }
            if ($null -eq $cur.Title -and ($c.Top - $cur.Y) -lt $MinHeight) { $cur.Title = $c.Text; $cur.TitleTop = $c.Top; Write-Log ("  title row {0} ({1}pt, {2}): '{3}' -> title of current page" -f $row, $c.Size, $c.Reason, $short); continue }
            $ref = if ($null -ne $cur.TitleTop) { $cur.TitleTop } else { $cur.Y }
            if (($c.Top - $ref) -lt $MinHeight) { Write-Log ("  title row {0} ({1}pt, {2}): '{3}' -> ignored, within {4}pt of the previous title" -f $row, $c.Size, $c.Reason, $short, $MinHeight); continue }
            Write-Log ("  title row {0} ({1}pt, {2}): '{3}' -> NEW PAGE" -f $row, $c.Size, $c.Reason, $short)
            $b = Get-PageBoundary -Candidate $c -Items $Items -MinY ($ref + 1)
            $cur = [pscustomobject]@{ Y = $b; Title = $c.Text; TitleTop = $c.Top }
            $starts.Add($cur)
        }
        if ($starts.Count -ge 2 -or $Mode -eq 'Titles') {
            $method = 'Page titles (font >= {0}pt): {1} page(s)' -f $MinFont, $starts.Count
        }
        else { $starts.Clear() }
    }

    if (-not $method -and ($Mode -eq 'Auto' -or $Mode -eq 'PageBreaks')) {
        $rows = @(1) + @($BreakRows | Where-Object { $_ -gt 1 -and $_ -le $Geo.LastRow } | Sort-Object -Unique)
        if ($rows.Count -ge 2 -or $Mode -eq 'PageBreaks') {
            foreach ($r in $rows) { $starts.Add([pscustomobject]@{ Y = $Geo.RowTops[$r]; Title = $null }) }
            $method = 'Manual page breaks in Wkly_Report: ' + ($rows -join ', ')
        }
    }

    if (-not $method) {
        $starts.Add([pscustomobject]@{ Y = 0.0; Title = $null })
        $method = 'No page titles or page breaks found - whole sheet shown as one page'
        Add-Issue -Severity Warning -Location 'Page detection' -Message 'Pages could not be detected; the whole worksheet is shown as one page. Use -PageMode Manual -PageStartRows to split it.'
    }

    $pages = New-Object System.Collections.Generic.List[object]
    $sorted = @($starts | Sort-Object Y)
    for ($i = 0; $i -lt $sorted.Count; $i++) {
        $sy = [double]$sorted[$i].Y
        $ey = if ($i + 1 -lt $sorted.Count) { [double]$sorted[$i + 1].Y } else { $ContentBottom }
        if ($ey - $sy -lt 1) { continue }
        $pageItems = @($Items | Where-Object { $_.T -lt $ey - 1 -and ($_.T + [math]::Max($_.H, 0.5)) -gt $sy + 1 } | Sort-Object Z)
        $content = @($pageItems | Where-Object { $_.Kind -in 'cell', 'shape', 'image', 'missing' })
        if ($content.Count -eq 0) { continue }
        # content extent of this page; left edge includes coloured side bars
        $minL = $ContentWidth; $maxR = 0
        foreach ($it in $content) { if ($it.L -lt $minL) { $minL = $it.L }; if ($it.L + $it.W -gt $maxR) { $maxR = $it.L + $it.W } }
        foreach ($it in $pageItems) { if ($it.Kind -eq 'fill' -and $it.L -lt $minL -and $it.L + $it.W -gt $minL - 40) { $minL = $it.L } }
        $minL = [math]::Max(0, $minL - 4); $maxR = [math]::Min($ContentWidth, $maxR + 4)
        if ($maxR - $minL -lt 50) { $minL = 0; $maxR = $ContentWidth }
        $idx = $pages.Count + 1
        $title = $sorted[$i].Title
        if (-not $title) { $title = Get-PageTitle -Items $content -StartY $sy -EndY $ey -Index $idx }
        if ($title.Length -gt 80) { $title = $title.Substring(0, 77) + '...' }
        $pages.Add([pscustomobject]@{
                Index    = $idx; StartY = $sy; EndY = $ey; Title = $title; OffsetX = $minL; Width = $maxR - $minL
                StartRow = (Get-RowAtY $Geo.RowTops $Geo.LastRow $sy)
                EndRow   = (Get-RowAtY $Geo.RowTops $Geo.LastRow ($ey - 0.5))
                Items    = $pageItems; ExcelView = $null; IssueCount = 0; Picture = $null
            })
    }
    Write-Log ("Page detection - {0}. Pages: {1}" -f $method, $pages.Count) 'OK'
    foreach ($p in $pages) { Write-Log ("  page {0}: rows {1}-{2}, {3} x {4} pt, '{5}'" -f $p.Index, $p.StartRow, $p.EndRow, [int]$p.Width, [int]($p.EndY - $p.StartY), $p.Title) }
    return [pscustomobject]@{ Pages = $pages; Method = $method }
}

function Find-LabelValue {
    param($Items, [string]$Label)
    $textItems = @($Items | Where-Object { $_.Kind -in 'cell', 'shape' } | Sort-Object T, L)
    $rx = '^\s*' + [regex]::Escape($Label) + '\s*:?\s*(.*)$'
    foreach ($it in $textItems) {
        $lines = @((Get-ItemText $it) -split "[`r`n`v]+" | Where-Object { $_.Trim() })
        foreach ($line in $lines) {
            if ($line -notmatch $rx) { continue }
            $rest = $Matches[1].Trim()
            if ($rest) { return $rest }
            if ($lines.Count -gt 1) { continue }
            $cy = $it.T + $it.H / 2
            $right = @($textItems | Where-Object {
                    $_ -ne $it -and $_.L -ge ($it.L + $it.W * 0.5) -and $_.L -le ($it.L + $it.W + 500) -and
                    [math]::Abs(($_.T + $_.H / 2) - $cy) -le [math]::Max(6, $it.H / 2) } | Sort-Object L)
            $v = (($right | ForEach-Object { ((Get-ItemText $_) -replace '\s+', ' ').Trim() }) -join ' ').Trim()
            if ($v) { return $v }
        }
    }
    return $null
}

# ======================================================================================
#  Page clean-up before drawing (pure logic - no Excel calls)
#  - every object belongs to exactly one page (the one it mostly sits on)
#  - helper cells outside the page panels are left out
#  - values placed over gaps in a sentence (separate small text boxes) are merged into
#    that sentence, so text can never overlap when the browser's line breaks differ
#  - each page gets its own width (no blank strip because another page is wider)
# ======================================================================================

function Test-OpaqueBg {
    param([string]$Bg)
    if (-not $Bg) { return $false }
    if ($Bg -match '^rgba\(\s*\d+\s*,\s*\d+\s*,\s*\d+\s*,\s*0(?:\.0+)?\s*\)$') { return $false }
    return $true
}

function Get-TextWidthPt {
    # Calibri advance widths (em) - only used to find which gap a value box fills.
    param([string]$Text, [double]$Size)
    if ($null -eq $script:CalibriEm) {
        # ordinal keys: 'a' and 'A' have different widths (a plain @{} would merge them); space is handled below
        $t = New-Object System.Collections.Hashtable ([System.StringComparer]::Ordinal)
        $pairs = 'a .479 b .525 c .423 d .525 e .498 f .305 g .471 h .525 i .229 j .239 k .455 l .229 m .799 n .525 o .527 p .525 q .525 r .349 s .391 t .335 u .525 v .452 w .715 x .433 y .453 z .395 ' +
            'A .579 B .544 C .533 D .615 E .488 F .459 G .631 H .623 I .252 J .319 K .520 L .420 M .855 N .646 O .662 P .517 Q .673 R .543 S .459 T .487 U .642 V .567 W .890 X .519 Y .487 Z .468 ' +
            ', .250 . .252 : .268 ; .268 | .460 / .386 ( .303 ) .303 % .715 & .682 + .498 - .306 _ .498 = .498 # .498 ~ .498 ! .326 ? .463 '' .221 " .401'
        $tok = $pairs.Split(' ', [System.StringSplitOptions]::RemoveEmptyEntries)
        for ($k = 0; $k + 1 -lt $tok.Count; $k += 2) { $t[[string]$tok[$k]] = [double]::Parse($tok[$k + 1], $Inv) }
        $script:CalibriEm = $t
    }
    $w = 0.0
    foreach ($ch in $Text.ToCharArray()) {
        $c = [int]$ch
        if ($ch -eq ' ') { $f = 0.226 }
        elseif ($c -ge 48 -and $c -le 57) { $f = 0.507 }
        elseif ($script:CalibriEm.ContainsKey([string]$ch)) { $f = $script:CalibriEm[[string]$ch] }
        elseif ($c -eq 0x2013) { $f = 0.5 }
        elseif ($c -eq 0x2014) { $f = 1.0 }
        elseif ($c -eq 0xB7) { $f = 0.268 }
        elseif ($c -ge 0xD800 -and $c -le 0xDFFF) { $f = 0.5 }
        elseif ($c -gt 0x2000) { $f = 1.0 }
        else { $f = 0.5 }
        $w += $f * $Size
    }
    return $w
}

function Remove-ItemRef {
    # List.Remove by reference (never by value equality)
    param($List, $Obj)
    for ($k = 0; $k -lt $List.Count; $k++) { if ([object]::ReferenceEquals($List[$k], $Obj)) { $List.RemoveAt($k); return $true } }
    return $false
}

function Get-ItemId {
    param($It)
    if ($null -eq $It.PSObject.Properties['_id']) { $script:ItemSeq++; $It | Add-Member -NotePropertyName _id -NotePropertyValue $script:ItemSeq -Force }
    return [string]$It._id
}

function Get-RunsText {
    param($It)
    if (-not $It.Runs) { return '' }
    $s = ''
    foreach ($r in $It.Runs) { $s += [string]$r.Text }
    return $s
}

function Get-TextLayout {
    # Where each gap (3+ spaces) and each paragraph end of a text box falls on the sheet,
    # by simulating Excel's word wrap with approximate character widths.
    param($It)
    $text = Get-RunsText $It
    $size = 11.0
    foreach ($r in $It.Runs) { if ([string]$r.Text -match '\S' -and [double]$r.Size -gt 0) { $size = [double]$r.Size; break } }
    $m = $It.Margins
    if (-not $m -or $m.Count -lt 4) { $m = @(7.2, 3.6, 7.2, 3.6) }
    $avail = [math]::Max(10.0, [double]$It.W - [double]$m[0] - [double]$m[2])
    $lineH = 1.22 * $size
    $slots = New-Object System.Collections.Generic.List[object]
    $line = 0; $x = 0.0; $pi = 0
    $lineStartSlot = 0
    $lineWidths = New-Object System.Collections.Generic.List[double]
    $i = 0; $n = $text.Length
    $paraAlign = 'left'
    $alignOf = { param($k) if ($It.Aligns -and $k -lt $It.Aligns.Count) { return [string]$It.Aligns[$k] } return 'left' }
    $paraAlign = & $alignOf 0
    $lineAlign = New-Object System.Collections.Generic.List[string]
    $finishLine = { $lineWidths.Add($x); $lineAlign.Add($paraAlign) }
    $noWrap = ($It.Wrap -eq $false)
    $paraStart = $true
    $paraLine = 0
    while ($i -le $n) {
        if ($paraStart) {
            $paraStart = $false
            $paraLine = $line
            $lead = ''
            if ($i -lt $n) { $lead = [string]$text[$i] }
            if ($lead -eq ' ' -or $lead -eq '/' -or $lead -eq '·') {
                $slots.Add([pscustomobject]@{ Kind = 'start'; Line = $line; X = 0.0; Start = $i; End = $i; Fill = (New-Object System.Collections.Generic.List[object]) })
            }
        }
        if ($i -eq $n -or $text[$i] -eq "`r") {
            $slots.Add([pscustomobject]@{ Kind = 'end'; Line = $line; X = $x; Start = $i; End = $i; ParaLine = $paraLine; Fill = (New-Object System.Collections.Generic.List[object]) })
            & $finishLine
            if ($i -eq $n) { break }
            $line++; $x = 0.0; $pi++; $paraAlign = & $alignOf $pi; $i++; $paraStart = $true; continue
        }
        if ($text[$i] -eq "`n" -or $text[$i] -eq "`v") { & $finishLine; $line++; $x = 0.0; $i++; continue }
        if ($text[$i] -eq ' ') {
            $j = $i; while ($j -lt $n -and $text[$j] -eq ' ') { $j++ }
            $runW = Get-TextWidthPt (' ' * ($j - $i)) $size
            if ($j - $i -ge 3) {
                $slots.Add([pscustomobject]@{ Kind = 'gap'; Line = $line; X = $x + $runW / 2; Start = $i; End = $j; Fill = (New-Object System.Collections.Generic.List[object]) })
            }
            $x += $runW; $i = $j; continue
        }
        $j = $i; while ($j -lt $n -and $text[$j] -ne ' ' -and $text[$j] -ne "`r" -and $text[$j] -ne "`n" -and $text[$j] -ne "`v") { $j++ }
        $ww = Get-TextWidthPt $text.Substring($i, $j - $i) $size
        if (-not $noWrap -and $x -gt 0 -and $x + $ww -gt $avail * 1.04) { & $finishLine; $line++; $x = 0.0 }
        $x += $ww; $i = $j
    }
    # paragraph alignment shifts each line; absolute sheet positions
    $nLines = [math]::Max(1, $lineWidths.Count)
    $top = [double]$It.T + [double]$m[1]
    if ($It.Anchor -eq 'middle') { $top = [double]$It.T + ([double]$It.H - $nLines * $lineH) / 2 }
    elseif ($It.Anchor -eq 'bottom') { $top = [double]$It.T + [double]$It.H - [double]$m[3] - $nLines * $lineH }
    foreach ($s in $slots) {
        $off = 0.0
        if ($s.Line -lt $lineWidths.Count) {
            $lw = $lineWidths[$s.Line]
            if ($lineAlign[$s.Line] -eq 'center') { $off = ($avail - $lw) / 2 } elseif ($lineAlign[$s.Line] -eq 'right') { $off = $avail - $lw }
        }
        $s | Add-Member -NotePropertyName AbsX -NotePropertyValue ([double]$It.L + [double]$m[0] + $off + $s.X)
        $s | Add-Member -NotePropertyName AbsY -NotePropertyValue ($top + ($s.Line + 0.5) * $lineH)
        if ($s.Kind -eq 'end') { $s | Add-Member -NotePropertyName ParaY -NotePropertyValue ($top + ($s.ParaLine + 0.5) * $lineH) }
    }
    return [pscustomobject]@{ Text = $text; Size = $size; LineH = $lineH; Slots = $slots; Lines = $nLines }
}

function Get-OverlayAnchor {
    # Sheet position of the visible text of a small value box.
    param($It)
    $raw = Get-RunsText $It
    $t = $raw.Trim()
    $size = 11.0
    foreach ($r in $It.Runs) { if ([string]$r.Text -match '\S' -and [double]$r.Size -gt 0) { $size = [double]$r.Size; break } }
    $m = $It.Margins
    if (-not $m -or $m.Count -lt 4) { $m = @(7.2, 3.6, 7.2, 3.6) }
    $tw = Get-TextWidthPt $t $size
    $lead = Get-TextWidthPt (' ' * ($raw.Length - $raw.TrimStart().Length)) $size
    $align = 'left'
    if ($It.Aligns -and $It.Aligns.Count) { $align = [string]$It.Aligns[0] }
    $x = [double]$It.L + [double]$m[0] + $lead + $tw / 2
    if ($align -eq 'center') { $x = [double]$It.L + [double]$It.W / 2 }
    elseif ($align -eq 'right') { $x = [double]$It.L + [double]$It.W - [double]$m[2] - $tw / 2 }
    $y = [double]$It.T + [double]$It.H / 2
    return [pscustomobject]@{ X = $x; Y = $y; Text = $t; Size = $size; Width = $tw }
}

function Get-ContainerOf {
    # Smallest filled card/panel under a point (so values are only joined within the same card).
    param($Items, [double]$X, [double]$Y)
    $best = $null; $bestA = 1e18
    foreach ($c in $Items) {
        if ($c.Kind -ne 'fill' -and $c.Kind -ne 'shape') { continue }
        if (-not (Test-OpaqueBg $c.Bg)) { continue }
        if ($X -lt $c.L -or $X -gt $c.L + $c.W -or $Y -lt $c.T -or $Y -gt $c.T + $c.H) { continue }
        $a = [double]$c.W * [double]$c.H
        if ($a -lt $bestA) { $bestA = $a; $best = $c }
    }
    if ($null -eq $best) { return '' }
    return (Get-ItemId $best)
}

function Test-SlotHost {
    # A sentence that has room for values Excel places over it: a run of 3+ spaces, or an ending such
    # as "Report Date:", "1 January –", "Completion at", "(".
    param([string]$Text)
    if ($Text -match '\S {3,}|^ {3,}\S|\( {2,}\)') { return $true }
    $t = $Text.TrimEnd()
    if ($t -match '[:–—\-|(/=]$' -or $t -match '\s(at|to|of|by|is|from)$') { return $true }
    return $false
}

function Get-SlotCost {
    # How well a value box (anchor A) fits a slot of a sentence (layout Lay, box H); 1e12 = cannot fit.
    param($A, $Sl, $Lay, $H)
    $dx = $A.X - $Sl.AbsX
    $dl = [math]::Abs(($A.Y - $Sl.AbsY) / $Lay.LineH)
    if ($Sl.Kind -eq 'end') {
        $right = [double]$H.L + [double]$H.W
        if ($A.X -ge $right - 4) {
            # placed just past the sentence's box: appended to that paragraph, whatever the wrap estimate says
            $dlp = [math]::Abs(($A.Y - $Sl.ParaY) / $Lay.LineH)
            $alt = 20 + ($A.X - $right) + 10 * [math]::Min($dl, $dlp)
            if ($A.X - $right -le 130 -and ($dl -le 1.6 -or $dlp -le 1.6)) { return $alt }
        }
        if ($dl -gt 1.6 -or $dx -lt -15 -or $dx -gt 130) { return 1e12 }
        return 20 + [math]::Max(0, $dx) + 10 * $dl
    }
    if ($dl -gt 1.6 -or [math]::Abs($dx) -gt [math]::Max(45.0, 4 * $Lay.Size)) { return 1e12 }
    $pen = 0
    # spaces that start a continuation paragraph are indentation rather than a value slot
    if ($Sl.Kind -eq 'gap' -and $Sl.Start -gt 0 -and "`r`n`v".IndexOf($Lay.Text[$Sl.Start - 1]) -ge 0) { $pen = 30 }
    return [math]::Abs($dx) + 10 * $dl + $pen
}

function Merge-OverlayText {
    # Values that Excel shows inside a sentence by placing a separate small text box over a gap
    # ("Report Date:  [10/22/2026]", "—  [67.66%]  of total") are written into that sentence.
    # Values are matched to the sentence's gaps in reading order; positions only confirm the match.
    param($Items)
    $plain = New-Object System.Collections.Generic.List[object]
    foreach ($it in $Items) {
        if ($it.Kind -ne 'shape' -or -not $it.Runs) { continue }
        if ((Test-OpaqueBg $it.Bg) -or $it.Border) { continue }
        $plain.Add($it)
    }
    $hosts = New-Object System.Collections.Generic.List[object]
    $overlays = New-Object System.Collections.Generic.List[object]
    foreach ($it in $plain) {
        $t = Get-RunsText $it
        $trim = $t.Trim()
        if ($trim.Length -gt 0 -and $trim.Length -le 32 -and $trim -notmatch '   ' -and $trim -notmatch "[`r`n`v]") { $overlays.Add($it) }
        if ($trim.Length -ge 4 -and (Test-SlotHost $t)) { $hosts.Add($it) }
    }
    if ($hosts.Count -eq 0 -or $overlays.Count -eq 0) { return 0 }
    $layouts = @{}; $boxOf = @{}
    foreach ($h in $hosts) {
        $hid = Get-ItemId $h
        try { $layouts[$hid] = Get-TextLayout $h } catch { continue }
        $boxOf[$hid] = Get-ContainerOf $Items ([double]$h.L + [math]::Min(12.0, [double]$h.W / 2)) ([double]$h.T + [math]::Min(8.0, [double]$h.H / 2))
    }
    # each value box goes to the one sentence it sits in (same card, inside or just right of it)
    $byHost = @{}
    foreach ($o in $overlays) {
        $a = Get-OverlayAnchor $o
        $obox = Get-ContainerOf $Items $a.X $a.Y
        $best = $null; $bestScore = 1e9
        foreach ($h in $hosts) {
            $hid = Get-ItemId $h
            if ([object]::ReferenceEquals($h, $o) -or -not $layouts.ContainsKey($hid)) { continue }
            if ($boxOf[$hid] -ne $obox) { continue }
            $lay = $layouts[$hid]
            if ($a.Y -lt $h.T - 4 -or $a.Y -gt $h.T + $h.H + 4) { continue }
            if ($a.X -lt $h.L - 8 -or $a.X -gt $h.L + $h.W + 70) { continue }
            $dist = 0.0
            if ($a.X -gt $h.L + $h.W) { $dist = $a.X - ($h.L + $h.W) }
            $fit = 1e12
            foreach ($sl in $lay.Slots) { $c = Get-SlotCost $a $sl $lay $h; if ($c -lt $fit) { $fit = $c } }
            if ($fit -ge 1e12) { continue }
            $score = $fit + $dist
            if ($score -lt $bestScore) { $bestScore = $score; $best = $h }
        }
        if ($null -eq $best) { continue }
        $bid = Get-ItemId $best
        if (-not $byHost.ContainsKey($bid)) { $byHost[$bid] = New-Object System.Collections.Generic.List[object] }
        $byHost[$bid].Add([pscustomobject]@{ O = $o; A = $a })
    }
    $mergedO = New-Object System.Collections.Generic.List[object]
    foreach ($h in $hosts) {
        $hid = Get-ItemId $h
        if (-not $byHost.ContainsKey($hid)) { continue }
        $lay = $layouts[$hid]
        $top = $h.T
        # values in reading order (same visual line -> left to right)
        $byY = @($byHost[$hid] | Sort-Object @{ Expression = { $_.A.Y } })
        $band = 0; $prevY = -1e9
        foreach ($e in $byY) { if ($e.A.Y - $prevY -gt 0.75 * $lay.LineH) { $band++ }; $prevY = $e.A.Y; $e | Add-Member -NotePropertyName Band -NotePropertyValue $band -Force }
        $ovs = @($byY | Sort-Object Band, @{ Expression = { $_.A.X } })
        $slots = @($lay.Slots | Sort-Object Start, @{ Expression = { if ($_.Kind -eq 'start') { 0 } elseif ($_.Kind -eq 'gap') { 1 } else { 2 } } })
        $nO = $ovs.Count; $nS = $slots.Count
        if ($nS -eq 0) { continue }
        # monotonic alignment of values to slots (dynamic programming): minimal total distance
        $INF = 1e12
        $dp = New-Object 'double[,]' ($nO + 1), ($nS + 1)
        $bk = New-Object 'int[,]' ($nO + 1), ($nS + 1)
        for ($i = 0; $i -le $nO; $i++) { for ($j = 0; $j -le $nS; $j++) { $dp[$i, $j] = $INF } }
        $dp[0, 0] = 0
        $costOf = {
            param($i, $j)
            $c = Get-SlotCost $ovs[$i].A $slots[$j] $lay $h
            if ($c -ge 1e12) { return $INF }
            return $c
        }
        for ($i = 0; $i -le $nO; $i++) {
            for ($j = 0; $j -le $nS; $j++) {
                $cur = $dp[$i, $j]
                if ($cur -ge $INF) { continue }
                if ($j -lt $nS -and $cur -lt $dp[$i, ($j + 1)]) { $dp[$i, ($j + 1)] = $cur; $bk[$i, ($j + 1)] = 1 }            # slot left empty
                if ($i -lt $nO -and $cur + 200 -lt $dp[($i + 1), $j]) { $dp[($i + 1), $j] = $cur + 200; $bk[($i + 1), $j] = 2 } # value not joined
                if ($i -lt $nO -and $j -lt $nS) {
                    $c = & $costOf $i $j
                    if ($c -lt $INF) {
                        if ($slots[$j].Kind -eq 'end') {
                            if ($cur + $c -lt $dp[($i + 1), $j]) { $dp[($i + 1), $j] = $cur + $c; $bk[($i + 1), $j] = 4 }       # several values after the text end
                        }
                        elseif ($cur + $c -lt $dp[($i + 1), ($j + 1)]) { $dp[($i + 1), ($j + 1)] = $cur + $c; $bk[($i + 1), ($j + 1)] = 3 }
                    }
                }
            }
        }
        $i = $nO; $j = $nS
        $pairs = New-Object System.Collections.Generic.List[object]
        while ($i -gt 0 -or $j -gt 0) {
            $b = $bk[$i, $j]
            if ($b -eq 1) { $j-- }
            elseif ($b -eq 2) { $i-- }
            elseif ($b -eq 3) { $pairs.Add(@(($i - 1), ($j - 1))); $i--; $j-- }
            elseif ($b -eq 4) { $pairs.Add(@(($i - 1), $j)); $i-- }
            else { break }
        }
        if ($pairs.Count -eq 0) { continue }
        foreach ($pr in $pairs) { $slots[$pr[1]].Fill.Add($ovs[$pr[0]]); $mergedO.Add($ovs[$pr[0]].O) }
        # rebuild the runs: host text with the values written into their slots
        $owner = New-Object System.Collections.Generic.List[object]
        foreach ($r in $h.Runs) { foreach ($ch in ([string]$r.Text).ToCharArray()) { $owner.Add($r) } }
        $at = @{}
        foreach ($sl in $slots) { $k = [string]$sl.Start; if (-not $at.ContainsKey($k)) { $at[$k] = New-Object System.Collections.Generic.List[object] }; $at[$k].Add($sl) }
        $out = New-Object System.Collections.Generic.List[object]
        $emit = {
            param($Run, [string]$Txt)
            if ($Txt.Length -eq 0) { return }
            if ($out.Count -gt 0 -and [object]::ReferenceEquals($out[$out.Count - 1].Src, $Run)) { $out[$out.Count - 1].Text += $Txt; return }
            $out.Add([pscustomobject]@{ Src = $Run; Text = $Txt })
        }
        $valueRun = {
            param($f)
            $vr = $f.O.Runs[0]
            foreach ($orr in $f.O.Runs) { if ([string]$orr.Text -match '\S') { $vr = $orr; break } }
            return $vr
        }
        $text = $lay.Text
        $i = 0
        $hasFill = $false
        while ($i -le $text.Length) {
            $k = [string]$i
            $skipTo = -1
            if ($at.ContainsKey($k)) {
                foreach ($sl in $at[$k]) {
                    if ($sl.Kind -eq 'start' -and $sl.Fill.Count) {
                        foreach ($f in $sl.Fill) { & $emit (& $valueRun $f) ((Get-RunsText $f.O).Trim() + ' ') }
                    }
                    elseif ($sl.Kind -eq 'end' -and $sl.Fill.Count) {
                        foreach ($f in $sl.Fill) {
                            $prev = ''; if ($out.Count) { $prev = $out[$out.Count - 1].Text }
                            $sep = ' '; if ($prev.EndsWith(' ') -or $prev.Length -eq 0) { $sep = '' }
                            & $emit (& $valueRun $f) ($sep + (Get-RunsText $f.O).Trim())
                        }
                    }
                    elseif ($sl.Kind -eq 'gap') {
                        $prevCh = ''; if ($i -gt 0) { $prevCh = [string]$text[$i - 1] }
                        $lead = ($i -eq 0 -or $prevCh -eq "`r" -or $prevCh -eq "`n" -or $prevCh -eq "`v")
                        if ($sl.Fill.Count) {
                            $f = $sl.Fill[0]
                            $nextCh = ''; if ($sl.End -lt $text.Length) { $nextCh = [string]$text[$sl.End] }
                            if (-not $lead -and $prevCh -ne '(') { & $emit $owner[$i] ' ' }
                            & $emit (& $valueRun $f) (Get-RunsText $f.O).Trim()
                            if (',.;:)%'.IndexOf($nextCh) -lt 0 -or $nextCh -eq '') { & $emit $owner[$sl.End - 1] ' ' }
                        }
                        elseif (-not $lead) { & $emit $owner[$i] ' ' }
                        $skipTo = $sl.End
                    }
                }
            }
            if ($i -eq $text.Length) { break }
            if ($skipTo -ge 0) { $i = $skipTo; continue }
            & $emit $owner[$i] ([string]$text[$i])
            $i++
        }
        $newRuns = New-Object System.Collections.Generic.List[object]
        foreach ($seg in $out) { $nr = $seg.Src.PSObject.Copy(); $nr.Text = $seg.Text; $newRuns.Add($nr) }
        $h.Runs = $newRuns
        # the sentence's box now also covers where the values stood, so it never wraps because of them
        $l0 = [double]$h.L; $r0 = [double]$h.L + [double]$h.W
        foreach ($pr in $pairs) { $ob = $ovs[$pr[0]].O; if ($ob.L -lt $l0) { $l0 = [double]$ob.L }; if ($ob.L + $ob.W -gt $r0) { $r0 = [double]$ob.L + [double]$ob.W } }
        $h.L = $l0; $h.W = $r0 - $l0
        $h | Add-Member -NotePropertyName MergedValues -NotePropertyValue $true -Force
    }
    foreach ($o in $mergedO) { [void](Remove-ItemRef $Items $o) }
    return $mergedO.Count
}

function Optimize-ReportPages {
    # Runs once after page detection, before any drawing. Never changes text or numbers -
    # it only decides which page an object is shown on and joins split sentences.
    param($Pages)
    $all = New-Object System.Collections.Generic.List[object]
    $seen = @{}
    foreach ($p in $Pages) { foreach ($it in $p.Items) { $id = Get-ItemId $it; if (-not $seen.ContainsKey($id)) { $seen[$id] = $true; $all.Add($it) } } }
    $lists = @{}
    foreach ($p in $Pages) { $lists[$p.Index] = New-Object System.Collections.Generic.List[object] }
    foreach ($it in $all) {
        $best = $null; $bestOv = -1.0
        $t0 = [double]$it.T; $t1 = [double]$it.T + [math]::Max([double]$it.H, 0.5)
        if ($it.Kind -eq 'shape' -and -not $it.Runs -and -not (Test-OpaqueBg $it.Bg)) { $c = ($t0 + $t1) / 2; $t0 = $c - 0.25; $t1 = $c + 0.25 }
        foreach ($p in $Pages) {
            $ov = [math]::Min($t1, [double]$p.EndY) - [math]::Max($t0, [double]$p.StartY)
            if ($ov -gt $bestOv) { $bestOv = $ov; $best = $p }
        }
        if ($best -and $bestOv -gt 0) { $lists[$best.Index].Add($it) }
    }
    $dropped = 0; $mergedTotal = 0; $helper = New-Object System.Collections.Generic.List[string]
    foreach ($p in $Pages) {
        $items = $lists[$p.Index]
        # empty text boxes (a single space) carry nothing to show
        foreach ($it in @($items | Where-Object { $_.Kind -eq 'shape' -and $_.Runs -and -not (Test-OpaqueBg $_.Bg) -and -not $_.Border -and (Get-RunsText $_).Trim().Length -eq 0 })) { [void](Remove-ItemRef $items $it); $dropped++ }
        # the same value box stacked several times at one spot: keep the top one
        $keep = @{}
        foreach ($it in ($items | Where-Object { $_.Kind -eq 'shape' -and $_.Runs } | Sort-Object Z -Descending)) {
            $key = '{0}|{1}|{2}' -f (Get-RunsText $it).Trim(), [math]::Round([double]$it.L / 3), [math]::Round([double]$it.T / 3)
            if ($keep.ContainsKey($key)) { [void](Remove-ItemRef $items $it); $dropped++ } else { $keep[$key] = $true }
        }
        # helper cells beside the page panels (calculation area) are not part of the page design
        $panelL = 1e9; $panelR = -1e9
        $extR = 0.0
        foreach ($it in $items) { if ($it.L + $it.W -gt $extR) { $extR = $it.L + $it.W } }
        foreach ($it in $items) {
            if ($it.Kind -ne 'shape' -or -not (Test-OpaqueBg $it.Bg)) { continue }
            if ($it.W -lt 0.4 * $extR -or $it.H -lt 100) { continue }
            if ($it.L -lt $panelL) { $panelL = $it.L }
            if ($it.L + $it.W -gt $panelR) { $panelR = $it.L + $it.W }
        }
        $kept = New-Object System.Collections.Generic.List[object]
        if ($panelR -gt 0) {
            foreach ($it in @($items | Where-Object { $_.Kind -in 'cell', 'fill', 'blank' -and ($_.L -ge $panelR + 2 -or $_.L + $_.W -le $panelL - 2) })) {
                [void](Remove-ItemRef $items $it)
                if ($it.Kind -eq 'cell') { $helper.Add([string]$it.Address); $kept.Add($it) }
            }
        }
        $p | Add-Member -NotePropertyName HelperCells -NotePropertyValue $kept -Force
        try { $mergedTotal += Merge-OverlayText -Items $items } catch { Add-Issue -Severity Warning -Location ('Page {0}' -f $p.Index) -Message ('Values inside sentences could not be joined: {0}' -f $_.Exception.Message) }
        # this page's own width
        # transparent decorations (no fill, no text) may reach past the panels; they never set the page size
        $minL = 1e9; $maxR = 0.0
        foreach ($it in $items) {
            if ($it.Kind -in 'blank') { continue }
            if ($it.Kind -eq 'shape' -and -not $it.Runs -and -not (Test-OpaqueBg $it.Bg)) { continue }
            $l = [double]$it.L; $r = [double]$it.L + [double]$it.W
            # a text box without fill or outline only shows its text: it counts as far as the text reaches
            if ($it.Kind -eq 'shape' -and $it.Runs -and -not (Test-OpaqueBg $it.Bg) -and -not $it.Border) {
                try {
                    $lineW = 0.0; $textW = 0.0
                    foreach ($run in $it.Runs) {
                        $parts = ([string]$run.Text) -split "[`r`n`v]"
                        for ($q = 0; $q -lt $parts.Count; $q++) {
                            if ($q -gt 0) { $lineW = 0.0 }
                            $lineW += Get-TextWidthPt $parts[$q] ([double]$run.Size)
                            if ($lineW -gt $textW) { $textW = $lineW }
                        }
                    }
                    $m = $it.Margins; if (-not $m) { $m = @(7.2, 3.6, 7.2, 3.6) }
                    $need = [math]::Min([double]$it.W, $textW * 1.05 + [double]$m[0] + [double]$m[2])
                    $al = 'left'; if ($it.Aligns -and $it.Aligns.Count) { $al = [string]$it.Aligns[0] }
                    if ($al -eq 'center') { $cx = $l + $it.W / 2; $l = $cx - $need / 2; $r = $cx + $need / 2 }
                    elseif ($al -eq 'right') { $l = $r - $need }
                    else { $r = $l + $need }
                } catch { }
            }
            if ($l -lt $minL) { $minL = $l }
            if ($r -gt $maxR) { $maxR = $r }
        }
        if ($maxR -gt $minL) { $p.OffsetX = [math]::Max(0, $minL - 4); $p.Width = $maxR + 4 - $p.OffsetX }
        # full-width bands (header bar, footer bar) reach both page edges, as in Excel's printed page
        foreach ($it in $items) {
            if (-not (($it.Kind -eq 'shape' -and (Test-OpaqueBg $it.Bg)) -or $it.Kind -eq 'fill')) { continue }
            if ($it.W -lt 0.85 * $p.Width) { continue }
            $r = [double]$it.L + [double]$it.W
            if ($it.L - $p.OffsetX -le 0.1 * $p.Width) { $it.L = [double]$p.OffsetX; $it.W = $r - $it.L }
            if ($p.OffsetX + $p.Width - $r -le 0.15 * $p.Width) { $it.W = [double]$p.OffsetX + [double]$p.Width - [double]$it.L }
        }
        # page height follows the content, so empty rows at the end of the page range add no space;
        # cell fills count only near the content (a panel's edge), never as a large empty area
        $minT = [double]$p.StartY; $maxB = 0.0
        foreach ($it in $items) {
            if ($it.Kind -in 'blank', 'fill') { continue }
            if ($it.Kind -eq 'shape' -and -not $it.Runs -and -not (Test-OpaqueBg $it.Bg)) { continue }
            if ($it.T -lt $minT) { $minT = [double]$it.T }; if ($it.T + $it.H -gt $maxB) { $maxB = [double]$it.T + [double]$it.H }
        }
        if ($maxB -le $minT) { $maxB = [double]$p.EndY }
        $contentB = $maxB
        foreach ($it in $items) {
            if ($it.Kind -ne 'fill') { continue }
            $b = [double]$it.T + [double]$it.H
            if ($b -gt $maxB -and $b -le $contentB + 40 -and $it.T -lt $contentB) { $maxB = $b }
        }
        $p.StartY = $minT; $p.EndY = $maxB + 4
        $sorted = New-Object System.Collections.Generic.List[object]
        foreach ($it in ($items | Sort-Object Z)) { $sorted.Add($it) }
        $p.Items = $sorted
    }
    Write-Log ("Page clean-up: {0} value(s) joined into their sentences, {1} empty/duplicate text box(es) skipped, each object shown on one page only." -f $mergedTotal, $dropped)
    if ($helper.Count) { Write-Log ("Helper cells beside the page panels are not shown: {0}" -f ($helper -join ', ')) }
}

# ======================================================================================
#  HTML rendering (pure logic - no Excel calls)
# ======================================================================================

function Get-BoxStyle {
    param($It, [double]$OffsetX, [double]$OffsetY)
    $s = 'left:{0}pt;top:{1}pt;width:{2}pt;height:{3}pt' -f (Format-Num ($It.L - $OffsetX)), (Format-Num ($It.T - $OffsetY)), (Format-Num ([math]::Max($It.W, 0))), (Format-Num ([math]::Max($It.H, 0)))
    if ($It.Rot -and [math]::Abs([double]$It.Rot) -gt 0.01) { $s += ';transform:rotate({0}deg)' -f (Format-Num $It.Rot) }
    return $s
}

function ConvertTo-RunStyle {
    param($Run)
    $p = New-Object System.Collections.Generic.List[string]
    if ($Run.Size -gt 0) { $p.Add('font-size:{0}pt' -f (Format-Num $Run.Size)) }
    if ($Run.Bold) { $p.Add('font-weight:700') }
    if ($Run.Italic) { $p.Add('font-style:italic') }
    if ($Run.Underline) { $p.Add('text-decoration:underline') }
    if ($Run.Color) { $p.Add('color:' + $Run.Color) }
    if ($Run.Font) { $p.Add('font-family:' + (Get-CssFont $Run.Font)) }
    return ($p -join ';')
}

function Get-InspectText {
    # Everything Excel reported for a text box, shown by the dashboard's Inspect mode.
    # Written with plain string concatenation only (Windows PowerShell 5.1 safe).
    param($It)
    $nl = [string][char]10
    $fill = 'none'; if ($It.Bg) { $fill = [string]$It.Bg }
    $outline = 'none'; if ($It.Border) { $outline = [string]$It.Border }
    $out = [string]$It.Name + ' | at ' + [string][int]$It.L + ',' + [string][int]$It.T + ' pt | size ' + [string][int]$It.W + ' x ' + [string][int]$It.H + ' pt | fill ' + $fill + ' | outline ' + $outline + ' | shape type ' + [string]$It.ShapeType
    if ($It.Runs) {
        $m = $It.Margins
        $margins = ''
        if ($m -and $m.Count -ge 4) { $margins = 'L' + [string]$m[0] + ' T' + [string]$m[1] + ' R' + [string]$m[2] + ' B' + [string]$m[3] }
        $link = 'none'; if ($It.Formula) { $link = [string]$It.Formula }
        $out += $nl + 'text anchor ' + [string]$It.Anchor + ' | margins ' + $margins + ' | wrap ' + [string]$It.Wrap + ' | autosize ' + [string]$It.AutoSize + ' | link ' + $link
        $k = 0
        foreach ($pp in $It.Paras) {
            if ($null -eq $pp) { continue }
            $k = $k + 1
            $out += $nl + 'paragraph ' + [string]$k + ': align ' + [string]$pp.Align + ', before ' + (Format-Num $pp.Before) + 'pt, after ' + (Format-Num $pp.After) + 'pt, line ' + [string]$pp.LineHeight
        }
        foreach ($r in $It.Runs) {
            $t = [string]$r.Text
            $t = $t.Replace([string][char]13, '[para]').Replace([string][char]11, '[line]').Replace([string][char]10, '[line]')
            if ($t.Length -gt 60) { $t = $t.Substring(0, 57) + '...' }
            $style = [string]$r.Font + ' ' + [string]$r.Size + 'pt'
            if ($r.Bold) { $style += ' bold' }
            if ($r.Italic) { $style += ' italic' }
            $colour = [string]$r.Color
            if ($r.ColorAdjustedFrom) { $colour += ' (Excel colour ' + [string]$r.ColorAdjustedFrom + ', adjusted to be readable)' }
            $out += $nl + '"' + $t + '" | ' + $style + ' | colour ' + $colour + ' | ' + [string]$r.ColorInfo
        }
    }
    return $out
}

function ConvertTo-ItemHtml {
    param($It, [double]$OffsetX, [double]$OffsetY)
    $box = Get-BoxStyle $It $OffsetX $OffsetY
    switch ($It.Kind) {
        'fill' { return '<div class="it" style="{0};background:{1}"></div>' -f $box, $It.Bg }
        'blank' { return '<div class="it blankcell" title="Wkly_Report!{0} - formula returns an empty value (shown as blank, not zero)"></div>' -f $It.Address }
        'cell' {
            $jc = switch ($It.Align) { 'center' { 'center' } 'right' { 'flex-end' } default { 'flex-start' } }
            $ai = switch ($It.VAlign) { 'top' { 'flex-start' } 'middle' { 'center' } default { 'flex-end' } }
            $st = '{0};justify-content:{1};align-items:{2};font-family:{3};font-size:{4}pt' -f $box, $jc, $ai, (Get-CssFont $It.FontName), (Format-Num $It.FontSize)
            if ($It.Bold) { $st += ';font-weight:700' }
            if ($It.Italic) { $st += ';font-style:italic' }
            if ($It.Underline) { $st += ';text-decoration:underline' }
            if ($It.Color) { $st += ';color:' + $It.Color }
            if ($It.Bg) { $st += ';background:' + $It.Bg }
            if ($It.Borders) { $st += ';' + $It.Borders }
            if ($It.Indent -gt 0) { $st += ';padding-left:{0}pt' -f (Format-Num (1.5 + 9 * $It.Indent)) }
            $spanStyle = "text-align:$($It.Align)"
            if ($It.Orientation -ne 0) { $spanStyle += ';transform:rotate({0}deg)' -f (Format-Num (-1 * $It.Orientation)) }
            $cls = if ($It.Wrap) { 'it cell wrap' } else { 'it cell' }
            return '<div class="{0}" title="Wkly_Report!{1}" style="{2}"><span style="{3}">{4}</span></div>' -f $cls, $It.Address, $st, $spanStyle, (ConvertTo-HtmlText $It.Text)
        }
        'shape' {
            $st = $box
            if ($It.Bg) { $st += ';background:' + $It.Bg }
            if ($It.Border) { $st += ';border:' + $It.Border }
            if ($It.Radius -and $It.Radius -ne '0') { $st += ';border-radius:' + $It.Radius }
            if ($It.Shadow) { $st += ';box-shadow:' + $It.Shadow }
            $inner = ''
            if ($It.Runs) {
                $m = $It.Margins
                $jc = switch ($It.Anchor) { 'middle' { 'center' } 'bottom' { 'flex-end' } default { 'flex-start' } }
                $tx = 'padding:{0}pt {1}pt {2}pt {3}pt;justify-content:{4}' -f (Format-Num $m[1]), (Format-Num $m[2]), (Format-Num $m[3]), (Format-Num $m[0]), $jc
                if ($It.TextOrientation -eq 2 -or $It.TextOrientation -eq 6) { $tx += ';writing-mode:vertical-rl;transform:rotate(180deg)' }
                elseif ($It.TextOrientation -eq 3 -or $It.TextOrientation -eq 5) { $tx += ';writing-mode:vertical-rl' }
                # base size = smallest run size, so the browser's default 16px never sets a taller
                # minimum line height than Excel uses (small-font text otherwise spills out of its box)
                $minSize = 0
                foreach ($run in $It.Runs) { if ($run.Size -gt 0 -and ($minSize -eq 0 -or $run.Size -lt $minSize)) { $minSize = [double]$run.Size } }
                if ($minSize -gt 0) { $tx += ';font-size:{0}pt' -f (Format-Num $minSize) }
                $sb = New-Object System.Text.StringBuilder
                $pi = 0
                $alignOf = {
                    param($k)
                    if ($It.Paras -and $k -lt $It.Paras.Count) {
                        $pp = $It.Paras[$k]
                        $st2 = 'text-align:' + $pp.Align
                        if ($pp.Before -gt 0) { $st2 += ';margin-top:' + (Format-Num $pp.Before) + 'pt' }
                        if ($pp.After -gt 0) { $st2 += ';margin-bottom:' + (Format-Num $pp.After) + 'pt' }
                        if ($pp.LineHeight) { $st2 += ';line-height:' + $pp.LineHeight }
                        return $st2
                    }
                    if ($It.Aligns -and $k -lt $It.Aligns.Count) { return 'text-align:' + $It.Aligns[$k] }
                    return 'text-align:left'
                }
                [void]$sb.Append('<div class="pa" style="' + (& $alignOf $pi) + '">')
                foreach ($run in $It.Runs) {
                    $segments = ([string]$run.Text) -split "`r"
                    for ($s = 0; $s -lt $segments.Count; $s++) {
                        if ($s -gt 0) {
                            $pi++
                            [void]$sb.Append('</div><div class="pa" style="' + (& $alignOf $pi) + '">')
                        }
                        if ($segments[$s].Length) { [void]$sb.Append('<span style="' + (ConvertTo-RunStyle $run) + '">' + (ConvertTo-HtmlText $segments[$s]) + '</span>') }
                        elseif ($s -gt 0 -or $segments.Count -eq 1) { [void]$sb.Append('<span style="' + (ConvertTo-RunStyle $run) + '">&#8203;</span>') }
                    }
                }
                [void]$sb.Append('</div>')
                $inner = '<div class="tx" style="{0}">{1}</div>' -f $tx, $sb.ToString()
            }
            $cls = if ($It.Wrap) { 'it shp' } else { 'it shp nowrap' }
            if ($It.Runs -and (Get-RunsText $It).Trim() -notmatch '\s') { $cls = 'it shp nowrap' }
            $inspect = ''
            try { $inspect = ([System.Net.WebUtility]::HtmlEncode([string](Get-InspectText $It))).Replace([string][char]10, '&#10;') }
            catch { $inspect = [System.Net.WebUtility]::HtmlEncode([string]$It.Name) + ' (details unavailable: ' + [System.Net.WebUtility]::HtmlEncode($_.Exception.Message) + ')' }
            return '<div class="{0}" data-i="{1}" style="{2}">{3}</div>' -f $cls, $inspect, $st, $inner
        }
        'line' {
            $w = [math]::Max($It.W, 0.01); $h = [math]::Max($It.H, 0.01)
            $x1 = if ($It.FlipH) { $w } else { 0 }; $x2 = $w - $x1
            $y1 = if ($It.FlipV) { $h } else { 0 }; $y2 = $h - $y1
            $st = 'left:{0}pt;top:{1}pt;width:{2}pt;height:{3}pt;overflow:visible' -f (Format-Num ($It.L - $OffsetX)), (Format-Num ($It.T - $OffsetY)), (Format-Num $w), (Format-Num $h)
            if ($It.Rot -and [math]::Abs([double]$It.Rot) -gt 0.01) { $st += ';transform:rotate({0}deg)' -f (Format-Num $It.Rot) }
            return '<svg class="it" style="{0}" viewBox="0 0 {1} {2}" preserveAspectRatio="none"><line x1="{3}" y1="{4}" x2="{5}" y2="{6}" stroke="{7}" stroke-width="{8}"/></svg>' -f $st, (Format-Num $w), (Format-Num $h), (Format-Num $x1), (Format-Num $y1), (Format-Num $x2), (Format-Num $y2), $It.Color, (Format-Num $It.Weight)
        }
        'image' {
            if ($It.IsChart -and $It.Spec -and $It.Spec.kind -and ($ChartRender -eq 'Native' -or -not $It.Src)) {
                $json = $It.Spec | ConvertTo-Json -Depth 8 -Compress
                $fb = ''
                if ($It.Src) { $fb = '<img src="{0}" alt="{1}" style="width:100%;height:100%">' -f $It.Src, (ConvertTo-Attr $It.Alt) }
                return ('<div class="it rchart" style="{0}" data-w="{1}" data-h="{2}" data-chart="{3}">{4}</div>' -f $box, (Format-Num ([double]$It.W * 4 / 3)), (Format-Num ([double]$It.H * 4 / 3)), [System.Net.WebUtility]::HtmlEncode($json), $fb)
            }
            $cls = if ($It.IsChart) { 'it img chart' } else { 'it img' }
            return '<img class="{0}" src="{1}" alt="{2}" title="{2}" style="{3}">' -f $cls, $It.Src, (ConvertTo-Attr $It.Alt), $box
        }
        'missing' { return '<div class="it miss" style="{0}">&#9888; {1}</div>' -f $box, (ConvertTo-HtmlText $It.Message) }
    }
    return ''
}

# ======================================================================================
#  Rebuilt pages: Excel page content laid out as flowing HTML (XY-cut), charts redrawn natively
# ======================================================================================

function Get-NodeArea { param($N) return [math]::Max(0.0, [double]$N.W) * [math]::Max(0.0, [double]$N.H) }

function Test-Contains {
    # $Outer contains most (>= 80%) of $Inner's area
    param($Outer, $Inner)
    $ix = [math]::Min($Outer.L + $Outer.W, $Inner.L + $Inner.W) - [math]::Max($Outer.L, $Inner.L)
    $iy = [math]::Min($Outer.T + $Outer.H, $Inner.T + $Inner.H) - [math]::Max($Outer.T, $Inner.T)
    if ($ix -le 0 -or $iy -le 0) { return $false }
    $a = Get-NodeArea $Inner
    if ($a -le 0) { return $false }
    return (($ix * $iy) / $a) -ge 0.8
}

function New-PageTree {
    # Builds containers (filled/bordered shapes and cell fills) with their children; thin bars become
    # decorations drawn at their exact position; returns the top-level nodes of the page.
    param($Items)
    $nodes = New-Object System.Collections.Generic.List[object]
    foreach ($it in $Items) {
        if ($it.Kind -in 'blank', 'line') { continue }
        if ($it.W -le 0.5 -or $it.H -le 0.5) { continue }
        $hasText = ($it.Kind -eq 'cell') -or ($it.Kind -eq 'shape' -and $it.Runs -and ((($it.Runs | ForEach-Object { $_.Text }) -join '') -match '\S'))
        $hasLook = ($it.Kind -eq 'fill') -or ($it.Kind -eq 'shape' -and ($it.Bg -or $it.Border))
        $thin = ($it.W -le 12 -or $it.H -le 9) -and -not $hasText -and $it.Kind -in 'shape', 'fill'
        $isBox = $hasLook -and -not $thin -and -not $hasText -and (Get-NodeArea $it) -ge 1500
        if ($it.Kind -eq 'shape' -and $hasLook -and $hasText -and (Get-NodeArea $it) -ge 6000) { $isBox = $true }   # card with its own text
        $nodes.Add([pscustomobject]@{ Item = $it; L = [double]$it.L; T = [double]$it.T; W = [double]$it.W; H = [double]$it.H; Z = [double]$it.Z
                IsBox = $isBox; Thin = $thin; Children = (New-Object System.Collections.Generic.List[object]); Decor = (New-Object System.Collections.Generic.List[object]); Parent = $null })
    }
    $boxes = @($nodes | Where-Object { $_.IsBox } | Sort-Object { Get-NodeArea $_ })
    foreach ($n in $nodes) {
        foreach ($b in $boxes) {
            if ($b -eq $n) { continue }
            if ((Get-NodeArea $b) -le (Get-NodeArea $n)) { continue }
            if ($n.Item.Kind -ne 'cell' -and $b.Item.Kind -ne 'fill' -and $n.Z -lt $b.Z) { continue }   # behind the box
            if (Test-Contains $b $n) { $n.Parent = $b; break }
        }
    }
    $top = New-Object System.Collections.Generic.List[object]
    foreach ($n in $nodes) {
        if ($n.Thin) {
            # accent bar on a card edge: attach to the smallest box whose top/left edge it touches
            $edgeBox = $null
            foreach ($b in $boxes) {
                $ox = [math]::Min($n.L + $n.W, $b.L + $b.W) - [math]::Max($n.L, $b.L)
                $oy = [math]::Min($n.T + $n.H, $b.T + $b.H) - [math]::Max($n.T, $b.T)
                $topEdge = $n.W -ge $n.H -and $ox -ge 0.6 * [math]::Min($n.W, $b.W) -and [math]::Abs($n.T - $b.T) -le 9
                $leftEdge = $n.H -gt $n.W -and $oy -ge 0.6 * [math]::Min($n.H, $b.H) -and [math]::Abs($n.L - $b.L) -le 9
                if ($topEdge -or $leftEdge) { $edgeBox = $b; break }
            }
            if ($edgeBox) {
                $n.T = [math]::Max($n.T, $edgeBox.T); $n.L = [math]::Max($n.L, $edgeBox.L)
                if ($n.W -ge $n.H) { $n.W = [math]::Max($n.W, 0.6 * $edgeBox.W) } else { $n.H = [math]::Max($n.H, 0.6 * $edgeBox.H) }
                $edgeBox.Decor.Add($n); continue
            }
            if ($n.Parent) { $n.Parent.Children.Add($n) } else { $top.Add($n) }
            continue
        }
        if ($n.Parent) { $n.Parent.Children.Add($n) } else { $top.Add($n) }
    }
    return , $top
}

function Split-Nodes {
    # Splits nodes into groups along one axis wherever there is a clean gap (XY-cut).
    param($Nodes, [string]$Axis)
    $sorted = if ($Axis -eq 'y') { @($Nodes | Sort-Object T, L) } else { @($Nodes | Sort-Object L, T) }
    $groups = New-Object System.Collections.Generic.List[object]
    $cur = New-Object System.Collections.Generic.List[object]
    $end = -1e9
    foreach ($n in $sorted) {
        $start = if ($Axis -eq 'y') { $n.T } else { $n.L }
        $stop = if ($Axis -eq 'y') { $n.T + $n.H } else { $n.L + $n.W }
        if ($cur.Count -gt 0 -and $start -ge $end - 2) { $groups.Add($cur); $cur = New-Object System.Collections.Generic.List[object]; $end = -1e9 }
        $cur.Add($n)
        if ($stop -gt $end) { $end = $stop }
    }
    if ($cur.Count -gt 0) { $groups.Add($cur) }
    return , $groups
}

function Get-Bounds {
    param($Nodes)
    $l = 1e9; $t = 1e9; $r = -1e9; $b = -1e9
    foreach ($n in $Nodes) {
        if ($n.L -lt $l) { $l = $n.L }; if ($n.T -lt $t) { $t = $n.T }
        if ($n.L + $n.W -gt $r) { $r = $n.L + $n.W }; if ($n.T + $n.H -gt $b) { $b = $n.T + $n.H }
    }
    return [pscustomobject]@{ L = $l; T = $t; W = $r - $l; H = $b - $t }
}

function ConvertTo-FlowHtml {
    # Grid-normalised layout of $Nodes (XY-cut): Excel's order and proportions are kept, but rows fill
    # the full width, side-by-side boxes get equal heights and all gaps are small and uniform.
    param($Nodes, [int]$Depth = 0)
    if (-not $Nodes -or $Nodes.Count -eq 0) { return '' }
    $sb = New-Object System.Text.StringBuilder
    if ($Nodes.Count -eq 1) { return (ConvertTo-NodeHtml $Nodes[0] $Depth) }
    $rows = Split-Nodes $Nodes 'y'
    if ($rows.Count -gt 1 -and $Depth -lt 40) {
        [void]$sb.Append('<div class="gs">')
        foreach ($r in $rows) { [void]$sb.Append('<div class="gsi">'); [void]$sb.Append((ConvertTo-FlowHtml $r ($Depth + 1))); [void]$sb.Append('</div>') }
        [void]$sb.Append('</div>')
        return $sb.ToString()
    }
    $cols = Split-Nodes $Nodes 'x'
    if ($cols.Count -gt 1 -and $Depth -lt 40) {
        $total = 0.0
        foreach ($c in $cols) { $total += (Get-Bounds $c).W }
        [void]$sb.Append('<div class="gr">')
        foreach ($c in $cols) {
            $bd = Get-Bounds $c
            # thin columns (side bars, separators) keep their width; others share the row by Excel proportions
            if ($bd.W -le 14) { [void]$sb.Append(('<div class="gc gthin" style="flex:0 0 {0}pt">' -f (Format-Num $bd.W))) }
            else { [void]$sb.Append(('<div class="gc" style="flex:{0} 1 0">' -f (Format-Num ([math]::Round($bd.W, 1))))) }
            [void]$sb.Append((ConvertTo-FlowHtml $c ($Depth + 1)))
            [void]$sb.Append('</div>')
        }
        [void]$sb.Append('</div>')
        return $sb.ToString()
    }
    # objects that genuinely overlap (e.g. a value box placed inside a sentence): keep Excel's relative positions
    $bd = Get-Bounds $Nodes
    [void]$sb.Append(('<div class="fa" style="width:{0}pt;height:{1}pt">' -f (Format-Num $bd.W), (Format-Num $bd.H)))
    foreach ($n in @($Nodes | Sort-Object Z)) {
        [void]$sb.Append(('<div class="fabs" style="left:{0}pt;top:{1}pt;width:{2}pt">' -f (Format-Num ($n.L - $bd.L)), (Format-Num ($n.T - $bd.T)), (Format-Num $n.W)))
        [void]$sb.Append((ConvertTo-NodeHtml $n $Depth -Fixed))
        [void]$sb.Append('</div>')
    }
    [void]$sb.Append('</div>')
    return $sb.ToString()
}

function Get-BoxCss {
    param($It)
    $st = ''
    if ($It.Bg) { $st += ';background:' + $It.Bg }
    if ($It.Kind -eq 'shape') {
        if ($It.Border) { $st += ';border:' + $It.Border }
        if ($It.Radius -and $It.Radius -ne '0') { $st += ';border-radius:' + $It.Radius }
        if ($It.Shadow) { $st += ';box-shadow:' + $It.Shadow }
    }
    elseif ($It.Kind -eq 'cell' -and $It.Borders) { $st += ';' + $It.Borders }
    return $st
}

function ConvertTo-NodeHtml {
    param($N, [int]$Depth, [switch]$Fixed)
    $it = $N.Item
    if ($N.IsBox) {
        $sb = New-Object System.Text.StringBuilder
        $hasKids = $N.Children.Count -gt 0
        # inner margin: Excel's own inset, kept small and even
        $pad = 6.0
        if ($hasKids) {
            $kb = Get-Bounds $N.Children
            $pad = [math]::Min(12.0, [math]::Max(4.0, [math]::Min($kb.L - $N.L, $kb.T - $N.T)))
        }
        $style = (Get-BoxCss $it)
        if (-not $hasKids) { $style += ';min-height:{0}pt' -f (Format-Num $N.H) }
        if ($Fixed) { $style += ';width:{0}pt' -f (Format-Num $N.W) }
        $style += ';padding:{0}pt' -f (Format-Num $pad)
        # accent bars attach to the card edge they touch; other thin decorations keep their position
        foreach ($d in $N.Decor) {
            $decBg = 'transparent'
            if ($d.Item.Bg) { $decBg = $d.Item.Bg }
            if ($d.W -ge 0.6 * $N.W -and ($d.T - $N.T) -le 8) { $style += ';border-top:{0}pt solid {1}' -f (Format-Num ([math]::Max(3, $d.H))), $decBg; continue }
            if ($d.H -ge 0.6 * $N.H -and ($d.L - $N.L) -le 8) { $style += ';border-left:{0}pt solid {1}' -f (Format-Num ([math]::Max(3, $d.W))), $decBg; continue }
            if ($d.W -ge 0.6 * $N.W -and ($N.T + $N.H - $d.T - $d.H) -le 8) { $style += ';border-bottom:{0}pt solid {1}' -f (Format-Num ([math]::Max(2, $d.H))), $decBg; continue }
        }
        [void]$sb.Append(('<div class="rbox" style="{0}">' -f $style.TrimStart(';')))
        if ($it.Kind -eq 'shape' -and $it.Runs) {
            $own = [pscustomobject]@{ Kind = 'shape'; L = 0; T = 0; W = $N.W; H = $N.H; Z = 0; Rot = 0; Name = $it.Name; Bg = $null; Border = $null; Radius = '0'; Shadow = $null
                Runs = $it.Runs; Aligns = $it.Aligns; Paras = $it.Paras; Anchor = $it.Anchor; Margins = $it.Margins; Wrap = $it.Wrap; TextOrientation = $it.TextOrientation; Formula = $it.Formula; AutoSize = $it.AutoSize; ShapeType = $it.ShapeType }
            try { [void]$sb.Append(('<div class="rown">{0}</div>' -f (ConvertTo-ItemHtml $own 0 0))) } catch { }
        }
        if ($hasKids) { [void]$sb.Append((ConvertTo-FlowHtml $N.Children ($Depth + 1))) }
        [void]$sb.Append('</div>')
        return $sb.ToString()
    }
    if ($it.Kind -eq 'image' -and $it.IsChart -and $it.Spec -and $it.Spec.kind) {
        $json = $it.Spec | ConvertTo-Json -Depth 8 -Compress
        $fallback = ''
        if ($it.Src) { $fallback = '<img src="{0}" alt="{1}" style="width:100%;height:100%">' -f $it.Src, (ConvertTo-Attr $it.Alt) }
        $w = 'width:100%'
        if ($Fixed) { $w = 'width:{0}pt' -f (Format-Num $N.W) }
        return ('<div class="rchart" data-w="{0}" data-h="{1}" data-chart="{2}" style="{3};aspect-ratio:{4} / {5}">{6}</div>' -f (Format-Num ($N.W * 4 / 3)), (Format-Num ($N.H * 4 / 3)), [System.Net.WebUtility]::HtmlEncode($json), $w, (Format-Num $N.W), (Format-Num $N.H), $fallback)
    }
    if ($it.Kind -eq 'image') {
        $w = 'width:100%'
        if ($Fixed) { $w = 'width:{0}pt' -f (Format-Num $N.W) }
        return ('<div class="rimg" style="{0};aspect-ratio:{1} / {2}"><img src="{3}" alt="{4}"></div>' -f $w, (Format-Num $N.W), (Format-Num $N.H), $it.Src, (ConvertTo-Attr $it.Alt))
    }
    # text and small shapes: drawn like the Excel object; plain text flows with its content height,
    # shapes with a fill (pills, labels) keep their Excel size
    $copy = $it.PSObject.Copy()
    $html = ''
    try { $html = ConvertTo-ItemHtml $copy $it.L $it.T } catch { }
    $solid = ($it.Kind -eq 'shape' -and ($it.Bg -or $it.Border)) -or ($it.Kind -eq 'cell' -and ($it.Bg -or $it.Borders)) -or $it.Kind -in 'missing', 'fill' -or $Fixed
    if ($N.Thin) {
        $barBg = 'transparent'
        if ($it.Bg) { $barBg = $it.Bg }
        return ('<div class="rleaf rbar" style="flex:1 1 auto;min-width:{0}pt;min-height:{1}pt;background:{2}"></div>' -f (Format-Num ([math]::Min($N.W, 12))), (Format-Num ([math]::Min($N.H, 12))), $barBg)
    }
    if ($solid) { return ('<div class="rleaf rsolid" style="width:{0}pt;height:{1}pt">{2}</div>' -f (Format-Num $N.W), (Format-Num $N.H), $html) }
    return ('<div class="rleaf rtext">{0}</div>' -f $html)
}

function Save-PageInventory {
    # What was read from each page (positions, text, colours, chart data) - attach this file when
    # something in the dashboard differs from Excel, so the exact structure can be checked.
    param($Pages, [string]$Path)
    try {
        $list = New-Object System.Collections.Generic.List[object]
        foreach ($p in $Pages) {
            $objs = New-Object System.Collections.Generic.List[object]
            foreach ($it in $p.Items) {
                $text = $null
                if ($it.Kind -eq 'cell') { $text = [string]$it.Text }
                elseif ($it.Kind -eq 'shape' -and $it.Runs) { $text = (($it.Runs | ForEach-Object { $_.Text }) -join '') }
                $o = [ordered]@{ kind = $it.Kind; name = [string]$it.Name; address = [string]$it.Address; L = [math]::Round($it.L, 1); T = [math]::Round($it.T, 1); W = [math]::Round($it.W, 1); H = [math]::Round($it.H, 1); Z = $it.Z
                    bg = $it.Bg; text = $text; link = $it.Formula }
                if ($it.Kind -eq 'shape' -and $it.Runs) { $o['runs'] = @($it.Runs | ForEach-Object { [ordered]@{ text = $_.Text; size = $_.Size; bold = $_.Bold; italic = $_.Italic; color = $_.Color; font = $_.Font } }) }
                if ($it.Kind -eq 'shape') {
                    $o['border'] = $it.Border; $o['radius'] = $it.Radius; $o['shadow'] = $it.Shadow; $o['rot'] = $it.Rot
                    $o['anchor'] = $it.Anchor; $o['wrap'] = $it.Wrap; $o['orient'] = $it.TextOrientation
                    if ($it.Margins) { $o['margins'] = @($it.Margins) }
                    if ($it.Aligns) { $al = New-Object System.Collections.Generic.List[string]; foreach ($x in $it.Aligns) { $al.Add([string]$x) }; $o['aligns'] = $al.ToArray() }
                }
                if ($it.Kind -eq 'cell') {
                    $o['font'] = [ordered]@{ name = $it.FontName; size = $it.FontSize; bold = $it.Bold; italic = $it.Italic; color = $it.Color }
                    $o['align'] = $it.Align; $o['valign'] = $it.VAlign; $o['wrapText'] = $it.Wrap; $o['borders'] = $it.Borders
                }
                if ($it.Kind -eq 'image') { $o['isChart'] = [bool]$it.IsChart; $o['hasPicture'] = [bool]$it.Src }
                if ($it.Spec) { $o['chart'] = $it.Spec }
                $objs.Add($o)
            }
            $list.Add([ordered]@{ page = $p.Index; title = $p.Title; rows = ('{0}-{1}' -f $p.StartRow, $p.EndRow); objects = $objs })
        }
        $dir = Split-Path -Parent $Path
        if (-not (Test-Path -LiteralPath $dir)) { New-Item -ItemType Directory -Path $dir -Force | Out-Null }
        [System.IO.File]::WriteAllText($Path, ($list | ConvertTo-Json -Depth 10), (New-Object System.Text.UTF8Encoding($false)))
        Write-Log ("Page inventory (what was read from each page): {0}" -f $Path)
    }
    catch { Write-Log ("Page inventory could not be written: {0}" -f $_.Exception.Message) 'WARN' }
}

function ConvertTo-RebuiltPage {
    param($Page)
    $items = @($Page.Items | Where-Object { $_.Kind -ne 'blank' })
    $top = New-PageTree $items
    return (ConvertTo-FlowHtml $top 0)
}

function New-DashboardSite {
    # Writes ONE self-contained .html to $OutDir; pictures are read from $WorkDir\assets and embedded.
    param($Pages, [string]$OutDir, $Meta, $Issues, [string]$WorkDir)
    if (-not $WorkDir) { $WorkDir = $OutDir }

    foreach ($iss in $Issues) {
        if ($iss.Y -lt 0) { continue }
        foreach ($p in $Pages) { if ($iss.Y -ge $p.StartY - 1 -and $iss.Y -lt $p.EndY) { $iss.Page = $p.Index; break } }
    }
    foreach ($p in $Pages) { $p.IssueCount = @($Issues | Where-Object { $_.Page -eq $p.Index -and $_.Severity -ne 'Info' }).Count }
    $issueTotal = @($Issues | Where-Object { $_.Severity -ne 'Info' }).Count

    $enc = { param($t) ConvertTo-HtmlText $t }
    $sb = New-Object System.Text.StringBuilder
    $add = { param($s) [void]$sb.Append($s) }
    $pt2px = 4.0 / 3.0
    $printW = 790.0; $printH = 520.0     # usable A4-landscape area in points

    & $add "<!doctype html>`n<html lang=`"en`">`n<head>`n<meta charset=`"utf-8`">`n<meta name=`"viewport`" content=`"width=device-width,initial-scale=1`">`n"
    & $add ('<title>{0}</title>' -f (& $enc $DashboardTitle))
    & $add "`n<link rel=`"stylesheet`" href=`"assets/dashboard.css`">`n</head>`n<body>`n"

    # ---- top bar
    & $add '<header class="topbar"><button class="btn menu" id="btnMenu" title="Show/hide the page list">&#9776;</button><div class="brand"><div class="mark">ITSM</div><span class="lbl">EXECUTIVE DASHBOARD</span></div>'
    & $add ('<div class="tb-title"><b>{0}</b><span>Source: {1} &middot; Wkly_Report (read-only snapshot)</span></div>' -f (& $enc $DashboardTitle), (& $enc $Meta.FileName))
    if ($Meta.ReportingPeriod) { & $add ('<span class="chip" title="Reporting Period (from Wkly_Report)">{0}</span>' -f (& $enc $Meta.ReportingPeriod)) }
    if ($Meta.ReportDate) { & $add ('<span class="chip" title="Report Date (from Wkly_Report)">Report Date: {0}</span>' -f (& $enc $Meta.ReportDate)) }
    & $add '<div class="zoom" title="Zoom (keys + and -). Fit = whole page width"><button class="btn" id="zOut">&minus;</button><span id="zVal">Fit</span><button class="btn" id="zIn">+</button><button class="btn" id="zFit">Fit</button></div>'
    & $add '<button class="btn" id="btnPresent" title="Presentation mode (P)">&#x26F6; Present</button>'
    & $add '<button class="btn" onclick="window.print()">Print / PDF</button>'
    if ($Classification) { & $add ('<span class="conf">{0}</span>' -f (& $enc $Classification)) }
    & $add "</header>`n<div class=`"layout`">`n"

    # ---- sidebar
    & $add '<aside class="sidebar"><h4>REPORT PAGES</h4><nav><ol id="nav">'
    foreach ($p in $Pages) {
        $flag = ''
        & $add ('<li><a href="#p{0}"><span class="n">{0}</span><span class="t">{1}</span>{2}</a></li>' -f $p.Index, (& $enc $p.Title), $flag)
    }
    & $add '</ol></nav><div class="src">'
    & $add ('<b>Workbook:</b> {0}<br><b>Worksheet:</b> Wkly_Report<br><b>Workbook saved:</b> {1}<br><b>Snapshot taken:</b> {2}<br><b>Pages detected:</b> {3}<br><b>Generator build:</b> {4}' -f (& $enc $Meta.FileName), (& $enc $Meta.Saved), (& $enc $Meta.Snapshot), $Pages.Count, $ScriptBuild)
    & $add "</div></aside>`n<main id=`"main`">`n"

    # ---- pages
    $n = $Pages.Count
    foreach ($p in $Pages) {
        $hPt = $p.EndY - $p.StartY
        $wPt = $p.Width
        $pz = [math]::Min(1.0, [math]::Min($printW / $wPt, $printH / $hPt))
        $pageClass = 'page'; if ($p.Picture) { $pageClass = 'page haspic' }
        & $add ('<section class="{2}" id="p{0}" data-title="{1}">' -f $p.Index, (ConvertTo-Attr $p.Title), $pageClass)
        & $add ('<div class="pbar"><div class="pt"><span class="pn">{0:D2}</span><h2>{1}</h2><span class="rows">rows {2}&ndash;{3}</span></div><div class="tools">' -f $p.Index, (& $enc $p.Title), $p.StartRow, $p.EndRow)
        & $add ('<span class="pnum">{0} / {1}</span></div></div>' -f $p.Index, $n)
        & $add ('<div class="pbody"><div class="cvwrap" style="--w:{0}pt;--h:{1}pt;--pz:{2}">' -f (Format-Num $wPt), (Format-Num $hPt), (Format-Num $pz))
        & $add ('<div class="cv" data-w="{0}" data-h="{1}" style="width:{2}pt;height:{3}pt">' -f (Format-Num ($wPt * $pt2px)), (Format-Num ($hPt * $pt2px)), (Format-Num $wPt), (Format-Num $hPt))
        if ($p.Picture) {
            $picHtml = ''
            try { $picHtml = ConvertTo-ItemHtml $p.Picture $p.OffsetX $p.StartY } catch { }
            & $add ('<div class="lay-pic">' + $picHtml + '</div>')
        }
        & $add '<div class="lay-html">'
        # pages shown as Excel's picture don't need the hidden re-drawn layer (keeps the single file small)
        $layerItems = $p.Items
        if ($p.Picture) { $layerItems = $null }
        foreach ($it in $layerItems) {
            $html = $null
            try { $html = ConvertTo-ItemHtml $it $p.OffsetX $p.StartY }
            catch {
                $what = [string]$it.Name; if (-not $what) { $what = [string]$it.Address }
                Add-Issue -Severity Warning -Location ("{0} '{1}'" -f $it.Kind, $what) -Message ("Could not be drawn and was skipped: {0}" -f (Format-ErrorDetail $_))
            }
            if ($html) { & $add $html; & $add "`n" }
        }
        & $add '</div></div></div></div>'
        & $add ('<div class="pf"><span>{0}</span><span>Wkly_Report &middot; page {1} of {2} &middot; rows {3}&ndash;{4}</span><span class="c">{5}</span></div></section>' -f (& $enc $DashboardTitle), $p.Index, $n, $p.StartRow, $p.EndRow, (& $enc $Classification))
        & $add "`n"
    }
    & $add "</main>`n</div>`n"

    # ---- presentation controls
    & $add '<div class="present-hud"><span id="hudPos"></span><button class="btn" id="hudPrev">&#9664;</button><button class="btn" id="hudNext">&#9654;</button><button class="btn" id="hudExit">Exit (Esc)</button></div>'
    & $add "<script src=`"assets/dashboard.js`"></script>`n</body>`n</html>`n"

    $index = Join-Path $OutDir ($Meta.BaseName + '.html')
    # an index.html left by builds r1-r5 of this script is removed so only the current dashboard remains
    $old = Join-Path $OutDir 'index.html'
    if ((Test-Path -LiteralPath $old) -and (Select-String -LiteralPath $old -Pattern 'assets/dashboard.js' -SimpleMatch -Quiet)) { Remove-Item -LiteralPath $old -Force }
    # Self-contained: styling, script and every picture are embedded, so the single .html file works
    # when sent by email or Teams without the assets folder.
    $html = $sb.ToString()
    $html = $html.Replace('<link rel="stylesheet" href="assets/dashboard.css">', "<style>`n" + $DashboardCss + "`n</style>")
    $html = $html.Replace('<script src="assets/dashboard.js"></script>', "<script>`n" + $DashboardJs + "`n</script>")
    $embedded = @{}
    foreach ($m in [regex]::Matches($html, 'src="(assets/[^"]+)"')) {
        $rel = $m.Groups[1].Value
        if ($embedded.ContainsKey($rel)) { continue }
        $file = Join-Path $WorkDir ($rel.Replace('/', [string][System.IO.Path]::DirectorySeparatorChar))
        if (-not (Test-Path -LiteralPath $file)) { Write-Log ("Picture {0} is missing and could not be embedded." -f $rel) 'WARN'; continue }
        $ext = [System.IO.Path]::GetExtension($file).ToLowerInvariant().TrimStart('.')
        $mime = 'image/png'
        if ($ext -eq 'jpg' -or $ext -eq 'jpeg') { $mime = 'image/jpeg' } elseif ($ext -eq 'gif') { $mime = 'image/gif' } elseif ($ext -eq 'svg') { $mime = 'image/svg+xml' } elseif ($ext -eq 'bmp') { $mime = 'image/bmp' } elseif ($ext -eq 'webp') { $mime = 'image/webp' }
        $embedded[$rel] = 'data:' + $mime + ';base64,' + [System.Convert]::ToBase64String([System.IO.File]::ReadAllBytes($file))
    }
    foreach ($rel in $embedded.Keys) { $html = $html.Replace('src="' + $rel + '"', 'src="' + $embedded[$rel] + '"') }
    # anything still pointing outside the file is a defect: report it instead of shipping a broken page
    $external = [regex]::Matches($html, '(?:src|href)="assets/[^"]*"').Count
    if ($external -gt 0) { Add-Issue -Severity Error -Location 'Dashboard file' -Message ("{0} reference(s) to the assets folder could not be embedded." -f $external) }
    if ([regex]::IsMatch($html, '(?:src|href)="https?://')) { Add-Issue -Severity Warning -Location 'Dashboard file' -Message 'The dashboard references an internet address.' }
    [System.IO.File]::WriteAllText($index, $html, (New-Object System.Text.UTF8Encoding($false)))
    $mb = [math]::Round((Get-Item -LiteralPath $index).Length / 1MB, 1)
    Write-Log ("Dashboard is a single self-contained file ({0} MB, {1} pictures embedded) - it can be sent by email or Teams on its own." -f $mb, $embedded.Count)
    if ($mb -gt 20) { Write-Log 'The dashboard file is larger than 20 MB; some email systems reject attachments that big - share it via Teams/OneDrive instead.' 'WARN' }
    return $index
}

# ======================================================================================
#  Dashboard stylesheet and script (written to assets\ on every run)
# ======================================================================================

$DashboardCss = @'
:root{--navy:#0b1640;--navy-2:#16266a;--gold:#c9a227;--gold-2:#e6c75c;--red:#c8372d;--green:#2f8f5b;--ink:#1b2333;--muted:#5d6678;--line:#dde1ea;--bg:#eceff4}
*{box-sizing:border-box}
html{scroll-behavior:smooth}
body{margin:0;font-family:"Segoe UI",Calibri,Arial,sans-serif;color:var(--ink);background:var(--bg);-webkit-print-color-adjust:exact;print-color-adjust:exact}
.topbar{position:sticky;top:0;z-index:30;min-height:62px;background:var(--navy);color:#fff;display:flex;align-items:center;gap:12px;padding:6px 18px 6px 0;border-bottom:3px solid var(--gold)}
.menu{margin-left:12px;font-size:16px;padding:5px 10px}
.brand{width:250px;flex:none;display:flex;align-items:center;gap:10px;padding-left:8px;font-weight:700;letter-spacing:.6px;font-size:14px}
.brand .mark{width:34px;height:34px;border-radius:7px;background:linear-gradient(135deg,var(--gold-2),#8c6d10);display:grid;place-items:center;color:var(--navy);font-size:12px;font-weight:800}
.tb-title{flex:1;min-width:0}
.tb-title b{display:block;font-size:15px;white-space:nowrap;overflow:hidden;text-overflow:ellipsis}
.tb-title span{display:block;font-size:12px;color:#c3cbe4;white-space:nowrap;overflow:hidden;text-overflow:ellipsis}
.topbar .btn,.topbar .chip,.topbar .conf{flex:none}
@media (max-width:1750px){.brand{width:auto}.brand .lbl{display:none}.chip{max-width:220px}}
.chip{border:1px solid rgba(201,162,39,.65);color:var(--gold-2);border-radius:999px;padding:4px 12px;font-size:12px;white-space:nowrap;max-width:340px;overflow:hidden;text-overflow:ellipsis}
.conf{background:var(--red);color:#fff;border-radius:4px;padding:4px 10px;font-size:11px;font-weight:700;letter-spacing:1.5px}
.btn{background:rgba(255,255,255,.08);border:1px solid rgba(255,255,255,.22);color:#fff;border-radius:6px;padding:7px 12px;font-size:13px;cursor:pointer;font-family:inherit;white-space:nowrap}
.btn:hover{background:rgba(255,255,255,.18)}
.btn[disabled]{opacity:.45;cursor:not-allowed}
.btn.warn{border-color:#f0b4ad}
.btn .cnt{background:var(--red);border-radius:999px;padding:0 6px;margin-left:6px;font-size:11px}
.layout{display:flex}
.sidebar{position:fixed;left:0;top:65px;z-index:45;height:calc(100vh - 65px);width:300px;background:#fff;border-right:1px solid var(--line);box-shadow:8px 0 24px rgba(0,0,0,.15);display:flex;flex-direction:column;transform:translateX(-105%);transition:transform .2s ease}
body.sb .sidebar{transform:none}
.zoom{display:flex;align-items:center;gap:4px}
.zoom .btn{padding:6px 10px}
.zoom span{font-size:12px;color:#c3cbe4;min-width:44px;text-align:center}
.pbody{overflow-x:auto}
.sidebar h4{margin:16px 22px 4px;font-size:11px;letter-spacing:1.6px;color:var(--muted)}
nav{overflow:auto}
nav ol{list-style:none;margin:0;padding:6px 12px}
nav a{display:flex;gap:10px;align-items:center;padding:8px 10px;border-radius:8px;color:var(--ink);text-decoration:none;font-size:13.5px;margin-bottom:2px}
nav a:hover{background:var(--bg)}
nav a .n{width:26px;height:26px;flex:none;border-radius:6px;background:var(--bg);display:grid;place-items:center;font-size:12px;font-weight:700;color:var(--muted)}
nav a .t{overflow:hidden;text-overflow:ellipsis;white-space:nowrap}
nav a.active{background:var(--navy);color:#fff}
nav a.active .n{background:var(--gold);color:var(--navy)}
nav a .flag{margin-left:auto;font-size:11px;color:#b42318;font-weight:700;white-space:nowrap}
nav a.active .flag{color:#ffb4ab}
.src{margin-top:auto;border-top:1px solid var(--line);padding:14px 22px;font-size:12px;color:var(--muted);line-height:1.65}
.src b{color:var(--ink);font-weight:600}
main{flex:1;min-width:0;padding:12px 12px}
.page{background:#fff;border-radius:10px;box-shadow:0 1px 3px rgba(16,24,40,.08),0 10px 28px rgba(16,24,40,.07);margin:0 0 16px;overflow:hidden;scroll-margin-top:84px;display:flex;flex-direction:column}
.pbar{display:flex;justify-content:space-between;align-items:center;gap:12px;padding:6px 14px;background:linear-gradient(90deg,var(--navy),var(--navy-2));border-left:6px solid var(--gold);color:#fff}
.pbar .pt{display:flex;align-items:center;gap:12px;min-width:0}
.pbar .pn{background:var(--gold);color:var(--navy);font-weight:800;border-radius:6px;padding:3px 8px;font-size:13px}
.pbar h2{margin:0;font-size:15px;white-space:nowrap;overflow:hidden;text-overflow:ellipsis}
.pbar .rows{font-size:11px;color:#aab3cf;white-space:nowrap}
.pbar .tools{display:flex;align-items:center;gap:10px}
.pbar .badge{background:#fdecea;color:#b42318;border-radius:999px;padding:3px 10px;font-size:12px;font-weight:700;white-space:nowrap}
.pnum{font-size:12px;color:#aab3cf;white-space:nowrap}
.xbtn{font-size:12px;padding:5px 10px}
.pbody{padding:6px;background:#f6f7fa;flex:1}
.cvwrap{position:relative;overflow:hidden;margin:0;border-radius:4px}
.cv{position:absolute;left:0;top:0;transform-origin:0 0;background:#fff;overflow:hidden}
.it{position:absolute;box-sizing:border-box;margin:0}
.cell{display:flex;padding:0 1.5pt;line-height:1.2;overflow:visible}
.cell>span{white-space:pre}
.cell.wrap>span{white-space:pre-wrap;overflow-wrap:anywhere;width:100%}
.shp{overflow:visible}
.shp .tx{position:absolute;inset:0;display:flex;flex-direction:column;overflow:visible}
.shp .pa{white-space:pre-wrap;line-height:1.2;overflow-wrap:anywhere}
.shp.nowrap .pa{white-space:pre}
.img{object-fit:fill}
.miss{border:1.5px dashed #c8372d;background:rgba(200,55,45,.07);color:#8a1c14;font:600 9pt "Segoe UI",sans-serif;display:flex;align-items:center;justify-content:center;text-align:center;padding:4pt}
.xerr{outline:2px dashed #e5484d;outline-offset:1px;border-radius:2px;cursor:help}
body.no-markers .xerr{outline:none}
.blankcell{border:1.5px dashed rgba(110,120,140,.75);background:repeating-linear-gradient(45deg,transparent 0 4px,rgba(110,120,140,.12) 4px 8px);cursor:help}
body.no-markers .blankcell{display:none}
.xlview{display:none;height:78vh}
.xlview iframe{width:100%;height:100%;border:0;background:#fff}
.page.xl .cvwrap{display:none}
.page.xl .xlview{display:block}
.pf{display:flex;justify-content:space-between;gap:12px;padding:8px 18px;background:var(--navy);color:#aab3cf;font-size:11px;letter-spacing:.3px}
.pf .c{color:var(--gold-2);letter-spacing:1.5px}
.drawer{position:fixed;top:0;right:0;height:100vh;width:560px;max-width:100%;background:#fff;box-shadow:-10px 0 30px rgba(0,0,0,.18);z-index:50;transform:translateX(105%);transition:transform .25s ease;display:flex;flex-direction:column}
.drawer.open{transform:none}
.drawer header{background:var(--navy);color:#fff;padding:14px 20px;display:flex;justify-content:space-between;align-items:center;border-bottom:3px solid var(--gold)}
.drawer .body{padding:14px 20px;overflow:auto;font-size:13px;line-height:1.5}
.drawer .lead{margin-top:0;color:var(--muted)}
.drawer h5{margin:16px 0 6px;font-size:12px;letter-spacing:1.5px;color:var(--muted)}
.drawer ul{padding-left:18px}
.drawer li{margin-bottom:6px}
.drawer code{background:#f1f3f8;padding:0 4px;border-radius:3px;font-size:11px;word-break:break-all}
.ok{color:var(--green);font-weight:700}.er{color:#b42318;font-weight:700}.wn{color:#a15c00;font-weight:700}
.iss{border-collapse:collapse;width:100%;font-size:12px}
.iss th{text-align:left;background:#f1f3f8;padding:6px}
.iss td{border-bottom:1px solid var(--line);padding:6px;vertical-align:top}
.iss tr.er td:first-child{border-left:3px solid #b42318}
.iss tr.wn td:first-child{border-left:3px solid #e5892a}
.iss tr.in td:first-child{border-left:3px solid #9aa3b2}
.iss tr td{font-weight:400;color:var(--ink)}
#inspectBox{display:none;position:fixed;left:12px;right:12px;bottom:12px;z-index:60;background:#0b1640;color:#fff;border:2px solid #c9a227;border-radius:8px;padding:10px 14px;max-height:45vh;overflow:auto}
#inspectBox pre{margin:0;white-space:pre-wrap;font:13px/1.45 Consolas,"Cascadia Mono",monospace}
#inspectBox .btn{float:right}
body.inspect .shp{cursor:crosshair}
body.inspect .shp:hover{outline:2px solid #e5484d;outline-offset:-1px}
.lay-pic,.lay-html{position:absolute;inset:0}
.cv.rb{position:absolute;left:0;top:0;transform-origin:0 0;background:#fff;overflow:visible}
.cv.rb{padding:8pt}
.gs{display:flex;flex-direction:column;gap:8pt}
.gsi{min-width:0}
.gr{display:flex;align-items:stretch;gap:10pt}
.gc{min-width:0;display:flex;flex-direction:column;gap:8pt}
.gc>.rbox,.gc>.gs,.gc>.rchart{flex:1 1 auto}
.gthin>*{flex:1 1 auto}
.fa{position:relative;max-width:100%}
.fabs{position:absolute}
.rbox{position:relative;box-sizing:border-box;display:flex;flex-direction:column;gap:8pt;min-width:0}
.rbox>.gs,.rbox>.gr{flex:1 1 auto}
.rown{position:relative;min-height:1em}
.rown>.it{position:relative!important;left:0!important;top:0!important;width:100%!important;height:auto!important}
.rown .tx{position:relative!important;inset:auto!important}
.rleaf{position:relative;box-sizing:border-box;max-width:100%}
.rleaf.rsolid{align-self:flex-start}
.rleaf.rtext>.it{position:relative!important;left:0!important;top:0!important;width:100%!important;height:auto!important;background:none}
.rleaf.rtext .tx{position:relative!important;inset:auto!important}
.rleaf.rtext .cell{display:flex}
.rimg img{width:100%;height:100%;object-fit:contain;display:block}
.rchart{position:relative;overflow:hidden}
.it.rchart{position:absolute}
.rchart svg{display:block;width:100%;height:100%}
.rchart text{pointer-events:none}
.rchart [data-tip]{pointer-events:all;cursor:default}
.rchart .ctip{position:absolute;pointer-events:none;background:rgba(11,22,64,.92);color:#fff;font:12px "Segoe UI",sans-serif;padding:4px 8px;border-radius:4px;white-space:nowrap;display:none;z-index:5}
.page.haspic .lay-html{display:none}
body.redraw .page.haspic .lay-html{display:block}
body.redraw .page.haspic .lay-pic{display:none}
.present-hud{display:none}
/* ?shot=N: one page alone at its exact size (used to make the PDF and PowerPoint files) */
body.shot{background:#fff;overflow:hidden}
body.shot .topbar,body.shot .sidebar,body.shot .present-hud,body.shot .pbar,body.shot .pf{display:none!important}
body.shot main{padding:0}
body.shot main>*:not(.shot-on){display:none!important}
body.shot .page.shot-on{display:block;margin:0;border-radius:0;box-shadow:none}
body.shot .pbody{padding:0;overflow:hidden;background:#fff}
body.shot .cvwrap{border-radius:0}
body.present .topbar,body.present .sidebar,body.present .xbtn,body.present .badge{display:none}
body.present main{padding:0}
body.present .page{display:none;margin:0;border-radius:0;min-height:100vh;box-shadow:none}
body.present .page.current{display:flex}
body.present .pbody{padding:6px;display:flex;align-items:flex-start;justify-content:center}
body.present .present-hud{display:flex;position:fixed;right:18px;bottom:40px;z-index:40;gap:8px;align-items:center;background:rgba(11,22,64,.85);color:#fff;border-radius:999px;padding:6px 8px 6px 16px;font-size:12px}
body.present .present-hud .btn{padding:4px 10px}
@media (max-width:980px){.sidebar,.brand{display:none}.tb-title{padding-left:16px}.chip{display:none}}
@media print{
  @page{size:A4 landscape;margin:8mm}
  body{background:#fff}
  .topbar,.sidebar,.drawer,.present-hud,.tools{display:none!important}
  main{padding:0}
  .page{box-shadow:none;border-radius:0;margin:0;break-after:page;page-break-after:always;display:flex!important}
  .page:last-child{break-after:auto;page-break-after:auto}
  .pbar{padding:4px 10px}
  .pbody{padding:4px;background:#fff}
  .xlview{display:none!important}
  .page.xl .cvwrap,.cvwrap{display:block!important;width:calc(var(--w) * var(--pz))!important;height:calc(var(--h) * var(--pz))!important;box-shadow:none}
  .cv{transform:scale(var(--pz))!important}
  .xerr{outline:none}.blankcell{display:none}
}
'@

$DashboardJs = @'
(function () {
  'use strict';
  var body = document.body;
  var pages = [].slice.call(document.querySelectorAll('.page'));
  var shotM = /[?&]shot=(\d+)/.exec(location.search), shot = !!shotM;
  if (shot) {
    body.classList.add('shot');
    var sp0 = document.getElementById('p' + shotM[1]);
    if (sp0) {
      sp0.classList.add('shot-on');
      // check mark just right of the page: proves the picture shows this page (it is cropped away)
      var cv0 = sp0.querySelector('.cv'), mk = document.createElement('div');
      mk.style.cssText = 'position:absolute;top:20px;width:10px;height:10px;background:rgb(1,2,3);left:' + (parseFloat(cv0.dataset.w) + 20) + 'px';
      body.appendChild(mk);
    }
  }
  var ERR = /#(DIV\/0!|N\/A|VALUE!|REF!|NAME\?|NUM!|NULL!|SPILL!|CALC!|FIELD!|BLOCKED!|CONNECT!|BUSY!|UNKNOWN!|GETTING_DATA)/g;

  // Excel errors are shown exactly as displayed; only an outline is added around them.
  pages.forEach(function (p) {
    var cv = p.querySelector('.cv'); if (!cv) return;
    var w = document.createTreeWalker(cv, NodeFilter.SHOW_TEXT, null), hits = [];
    while (w.nextNode()) { ERR.lastIndex = 0; if (ERR.test(w.currentNode.nodeValue)) hits.push(w.currentNode); }
    hits.forEach(function (t) {
      var s = t.nodeValue, frag = document.createDocumentFragment(), last = 0, m;
      ERR.lastIndex = 0;
      while ((m = ERR.exec(s))) {
        if (m.index > last) frag.appendChild(document.createTextNode(s.slice(last, m.index)));
        var sp = document.createElement('span'); sp.className = 'xerr';
        sp.title = 'Excel error value - shown exactly as displayed in Wkly_Report (not converted to 0)';
        sp.textContent = m[0]; frag.appendChild(sp); last = m.index + m[0].length;
      }
      if (last < s.length) frag.appendChild(document.createTextNode(s.slice(last)));
      t.parentNode.replaceChild(frag, t);
    });
  });

  var zoom = 1;
  try { var zs = localStorage.getItem('itsmdash.zoom'); zoom = parseFloat(zs) || 1; } catch (e) { }
  // Sharp 1:1: scale so one pixel of Excel's page picture is one screen pixel (no resampling blur)
  function sharpScale() {
    var img = document.querySelector('.lay-pic img');
    if (!img || !img.naturalWidth || !img.offsetWidth) return 0;
    return img.naturalWidth / ((window.devicePixelRatio || 1) * img.offsetWidth);
  }
  function setZoom(z) {
    if (z === 'sharp') {
      zoom = 'sharp';
      document.getElementById('zVal').textContent = '1:1';
      try { localStorage.setItem('itsmdash.zoom', 'sharp'); } catch (e) { }
      fit();
      return;
    }
    if (zoom === 'sharp' || !(z > 0)) z = 1;
    zoom = Math.max(0.5, Math.min(3, Math.round(z * 100) / 100));
    document.getElementById('zVal').textContent = zoom === 1 ? 'Fit' : Math.round(zoom * 100) + '%';
    try { localStorage.setItem('itsmdash.zoom', String(zoom)); } catch (e) { }
    fit();
  }

  // ---- native charts redrawn from Excel's chart data (no external libraries) ----
  function decimalsOf(nf) { var m = /0\.(0+)/.exec(nf || ''); return m ? m[1].length : 0; }
  function fmt(v, nf) {
    if (v === null || v === undefined || isNaN(v)) return '';
    nf = nf || 'General';
    var sec = nf.split(';');
    if (v === 0 && sec.length > 2) { var z = sec[2].replace(/_.|\*.|["\\\s]/g, ''); if (/^[-–]$/.test(z)) return '-'; }
    if (v < 0 && sec.length > 1) nf = sec[1]; else nf = sec[0];
    if (!nf.replace(/["\s]/g, '')) return '';   // ;;; = hidden in Excel
    if (/general/i.test(nf)) { var r = Math.round(v * 100) / 100; return String(r); }
    if (/\[?h\]?:mm/i.test(nf)) {
      var t = Math.round(Math.abs(v) * 86400), h = Math.floor(t / 3600), m = Math.floor(t % 3600 / 60), s2 = t % 60;
      return h + ':' + (m < 10 ? '0' : '') + m + (/ss/i.test(nf) ? ':' + (s2 < 10 ? '0' : '') + s2 : '');
    }
    var d = decimalsOf(nf), pct = nf.indexOf('%') >= 0, x = Math.abs(pct ? v * 100 : v);
    var str = x.toFixed(d);
    if (nf.indexOf(',') >= 0) { var parts = str.split('.'); parts[0] = parts[0].replace(/\B(?=(\d{3})+(?!\d))/g, ','); str = parts.join('.'); }
    if (v < 0 && sec.length < 2) str = '-' + str;
    return str + (pct ? '%' : '');
  }
  function esc(t) { return String(t).replace(/&/g, '&amp;').replace(/</g, '&lt;').replace(/>/g, '&gt;'); }
  function px(pt) { return pt * 4 / 3; }
  function fontAttr(f, defSize, defColor) {
    f = f || {};
    var sz = px(f.size || defSize), col = f.color || defColor;
    return 'font-size="' + sz.toFixed(1) + '" fill="' + col + '" font-family="' + esc(f.name || 'Calibri') + ',Calibri,Segoe UI,sans-serif"' +
      (f.bold ? ' font-weight="700"' : '') + (f.italic ? ' font-style="italic"' : '');
  }
  function niceAxis(min, max) {
    if (min > 0) min = 0;
    if (max <= min) max = min + 1;
    var span = max - min, step = Math.pow(10, Math.floor(Math.log10(span / 5))), f = span / 5 / step;
    step *= f <= 1 ? 1 : f <= 2 ? 2 : f <= 5 ? 5 : 10;
    return { min: Math.floor(min / step) * step, max: Math.ceil(max / step) * step, major: step };
  }
  function axisOf(ax, vals) {
    if (ax && ax.max !== null && ax.max !== undefined && ax.major > 0 && ax.max > ax.min) return { min: ax.min, max: ax.max, major: ax.major };
    var lo = Math.min.apply(null, vals.concat([0])), hi = Math.max.apply(null, vals.concat([0]));
    return niceAxis(lo, hi);
  }
  function wrapText(t, maxChars) {
    var words = String(t).split(/\s+/), lines = [], cur = '';
    words.forEach(function (w) { if ((cur + ' ' + w).trim().length > maxChars && cur) { lines.push(cur); cur = w; } else cur = (cur + ' ' + w).trim(); });
    if (cur) lines.push(cur);
    return lines.slice(0, 3);
  }
  // text is measured with the real font, so labels are shortened or skipped instead of overlapping
  var mctx = null;
  function tw(t, f, defSize) {
    f = f || {};
    try {
      if (!mctx) mctx = document.createElement('canvas').getContext('2d');
      mctx.font = (f.italic ? 'italic ' : '') + (f.bold ? '700 ' : '') + px(f.size || defSize).toFixed(1) + 'px ' + (f.name ? '"' + f.name + '",' : '') + 'Calibri,Carlito,"Segoe UI",sans-serif';
      return mctx.measureText(String(t)).width;
    } catch (e) { return String(t).length * px(f.size || defSize) * 0.55; }
  }
  function clip(t, maxW, f, defSize) {
    t = String(t);
    if (maxW <= 0) return '';
    if (tw(t, f, defSize) <= maxW) return t;
    while (t.length > 1 && tw(t + '\u2026', f, defSize) > maxW) t = t.slice(0, -1);
    return t.replace(/\s+$/, '') + '\u2026';
  }
  function wrapPx(t, maxW, f, defSize, maxLines) {
    var words = String(t).split(/\s+/), lines = [], cur = '';
    words.forEach(function (w) {
      var nx = cur ? cur + ' ' + w : w;
      if (cur && tw(nx, f, defSize) > maxW) { lines.push(cur); cur = w; } else cur = nx;
    });
    if (cur) lines.push(cur);
    if (lines.length > maxLines) { lines = lines.slice(0, maxLines); lines[maxLines - 1] += '\u2026'; }
    return lines.map(function (ln) { return clip(ln, maxW, f, defSize); });
  }
  function scaled(f, k) { var g = {}; f = f || {}; for (var key in f) g[key] = f[key]; g.size = (f.size || 9) * k; return g; }
  // places a label only where it does not overlap one already placed
  function placer(W, H) {
    var boxes = [];
    return function (x, y, text, f, anchor, defSize) {
      if (!text) return '';
      var w = tw(text, f, defSize), h = px((f && f.size) || defSize);
      var x0 = anchor === 'middle' ? x - w / 2 : anchor === 'end' ? x - w : x;
      if (x0 < 1) { x0 = 1; } if (x0 + w > W - 1) { x0 = W - 1 - w; }
      var b = { l: x0, r: x0 + w, t: y - h * 0.8, b: y + h * 0.2 };
      for (var i = 0; i < boxes.length; i++) { var q = boxes[i]; if (b.l < q.r + 1 && b.r > q.l - 1 && b.t < q.b && b.b > q.t) return ''; }
      boxes.push(b);
      return '<text x="' + x0.toFixed(1) + '" y="' + y.toFixed(1) + '" ' + fontAttr(f, defSize, chartTxt) + '>' + esc(text) + '</text>';
    };
  }
  var chartTxt = '#404040';
  function seriesColor(sr, i) {
    var pal = ['#4472c4', '#ed7d31', '#a5a5a5', '#ffc000', '#5b9bd5', '#70ad47'];
    if (sr.kind === 'line' || sr.kind === 'area') return sr.lineOn !== false && sr.line ? sr.line : (sr.fill || pal[i % 6]);
    return sr.fillOn !== false && sr.fill ? sr.fill : (sr.line || pal[i % 6]);
  }
  function renderChart(el) {
    var sp = JSON.parse(el.getAttribute('data-chart'));
    var W = parseFloat(el.dataset.w), H = parseFloat(el.dataset.h);
    var txt = (sp.font && sp.font.color) || '#404040';
    chartTxt = txt;
    var put = placer(W, H);
    var o = [], tips = [];
    var top = 8, bottom = H - 8, left = 10, right = W - 10;
    if (sp.bg) { el.style.background = sp.bg; }
    if (sp.rounded) el.style.borderRadius = '14px';
    if (sp.shadow) el.style.boxShadow = '0 3px 8px rgba(0,0,0,.35)';
    if (sp.title) {
      var tf = sp.titleFont || {}, tsz = px(tf.size || 14);
      wrapText(sp.title, Math.max(10, Math.floor(W / (tsz * 0.55)))).forEach(function (ln, k) {
        o.push('<text x="' + W / 2 + '" y="' + (top + tsz * (k + 1)) + '" text-anchor="middle" ' + fontAttr(tf, 14, txt) + '>' + esc(ln) + '</text>');
        top += 0;
      });
      top += tsz * wrapText(sp.title, Math.max(10, Math.floor(W / (tsz * 0.55)))).length + 8;
    }
    var series = sp.series || [];
    var cats = [];
    series.forEach(function (s) { if (s.cats && s.cats.length > cats.length) cats = s.cats; });
    var pie = sp.kind === 'pie' || sp.kind === 'doughnut';
    // legend
    var legend = [];
    if (sp.legend) {
      if (pie && series[0]) (series[0].cats || []).forEach(function (c, i) { legend.push({ name: c, color: (series[0].points && series[0].points[i]) || seriesColor(series[0], i), line: false }); });
      else series.forEach(function (s, i) { legend.push({ name: s.name, color: seriesColor(s, i), line: s.kind === 'line' }); });
    }
    var lf = sp.legendFont || sp.font || {}, lsz = px(lf.size || 9);
    if (legend.length && (sp.legend === 'bottom' || sp.legend === 'top')) {
      // legend entries wrap onto more rows when they do not fit on one
      var avail = W - 20, rows = [[]], rw = [0];
      legend.forEach(function (e) {
        var nm = clip(e.name, avail - lsz * 2.2, lf, 9), w = tw(nm, lf, 9) + lsz * 2.2;
        if (rw[rw.length - 1] + w > avail && rows[rows.length - 1].length) { rows.push([]); rw.push(0); }
        rows[rows.length - 1].push({ e: e, nm: nm, w: w }); rw[rw.length - 1] += w;
      });
      var rowH = lsz * 1.45, ly0 = sp.legend === 'bottom' ? bottom - lsz * 0.3 - rowH * (rows.length - 1) : top + lsz;
      rows.forEach(function (row, ri) {
        var x0 = Math.max(10, (W - rw[ri]) / 2), ly = ly0 + ri * rowH;
        row.forEach(function (it) {
          var e = it.e;
          o.push(e.line ? '<line x1="' + x0 + '" y1="' + (ly - lsz * 0.35) + '" x2="' + (x0 + lsz * 1.4) + '" y2="' + (ly - lsz * 0.35) + '" stroke="' + e.color + '" stroke-width="3"/>'
            : '<rect x="' + x0 + '" y="' + (ly - lsz * 0.8) + '" width="' + lsz * 0.8 + '" height="' + lsz * 0.8 + '" fill="' + e.color + '"/>');
          o.push('<text x="' + (x0 + lsz * 1.6) + '" y="' + ly + '" ' + fontAttr(lf, 9, txt) + '>' + esc(it.nm) + '</text>');
          x0 += it.w;
        });
      });
      if (sp.legend === 'bottom') bottom -= rowH * rows.length + lsz * 0.4; else top += rowH * rows.length + lsz * 0.4;
    } else if (legend.length) {
      var ry = top + 10, lw = Math.min(W * 0.32, Math.max.apply(null, legend.map(function (e) { return tw(e.name, lf, 9); })) + lsz * 2);
      var rx = sp.legend === 'left' ? left : right - lw;
      legend.forEach(function (e) {
        if (ry > bottom) return;
        o.push('<rect x="' + rx + '" y="' + (ry - lsz * 0.8) + '" width="' + lsz * 0.8 + '" height="' + lsz * 0.8 + '" fill="' + e.color + '"/>');
        o.push('<text x="' + (rx + lsz * 1.4) + '" y="' + ry + '" ' + fontAttr(lf, 9, txt) + '>' + esc(clip(e.name, lw - lsz * 1.4, lf, 9)) + '</text>');
        ry += lsz * 1.5;
      });
      if (sp.legend === 'left') left += lw + 6; else right -= lw + 6;
    }
    if (pie) {
      var s0 = series[0], vals = (s0.values || []).map(function (v) { return v > 0 ? v : 0; });
      var tot = vals.reduce(function (a, b) { return a + b; }, 0);
      var cx = (left + right) / 2, cy = (top + bottom) / 2, R = Math.max(10, Math.min(right - left, bottom - top) / 2 - 18);
      var r0 = sp.kind === 'doughnut' ? R * (sp.hole || 50) / 100 : 0;
      var a = (sp.firstAngle || 0) * Math.PI / 180 - Math.PI / 2;
      var lb = s0.labels, pieLabels = [];
      vals.forEach(function (v, i) {
        if (!tot) return;
        var a2 = a + v / tot * 2 * Math.PI, col = (s0.points && s0.points[i]) || seriesColor(s0, i), big = a2 - a > Math.PI ? 1 : 0;
        var P = function (r, ang) { return (cx + r * Math.cos(ang)).toFixed(2) + ',' + (cy + r * Math.sin(ang)).toFixed(2); };
        var d = v / tot >= 0.9999 ? (r0 ? 'M' + P(R, 0) + ' A' + R + ',' + R + ' 0 1 1 ' + P(R, Math.PI) + ' A' + R + ',' + R + ' 0 1 1 ' + P(R, 0) + ' M' + P(r0, 0) + ' A' + r0 + ',' + r0 + ' 0 1 0 ' + P(r0, Math.PI) + ' A' + r0 + ',' + r0 + ' 0 1 0 ' + P(r0, 0) + 'Z'
          : 'M' + P(R, 0) + ' A' + R + ',' + R + ' 0 1 1 ' + P(R, Math.PI) + ' A' + R + ',' + R + ' 0 1 1 ' + P(R, 0) + 'Z')
          : 'M' + P(R, a) + ' A' + R + ',' + R + ' 0 ' + big + ' 1 ' + P(R, a2) + (r0 ? ' L' + P(r0, a2) + ' A' + r0 + ',' + r0 + ' 0 ' + big + ' 0 ' + P(r0, a) + 'Z' : ' L' + cx + ',' + cy + 'Z');
        var cat = (s0.cats || [])[i] || '';
        var vf = lb && lb.format && lb.format.indexOf('%') < 0 && !/general/i.test(lb.format) ? lb.format : '#,##0';
        tips.push(cat + ': ' + fmt(s0.values[i], vf) + ' (' + (v / tot * 100).toFixed(2) + '%)');
        o.push('<path d="' + d + '" fill="' + col + '" stroke="rgba(0,0,0,.25)" stroke-width="1" data-tip="' + (tips.length - 1) + '"/>');
        if (lb && v > 0 && v / tot >= 0.02) {
          var parts = [];
          if (lb.category) parts.push(cat);
          if (lb.value) parts.push(fmt(s0.values[i], lb.format));
          if (lb.percent) parts.push(fmt(v / tot, lb.format && lb.format.indexOf('%') >= 0 ? lb.format : '0%'));
          var mid = (a + a2) / 2, lr = r0 ? (R + r0) / 2 : (v / tot < 0.08 ? R + 14 : R * 0.65);
          pieLabels.push([cx + lr * Math.cos(mid), cy + lr * Math.sin(mid) + 4, parts.join(', ')]);
        }
        a = a2;
      });
      pieLabels.forEach(function (q) { o.push(put(q[0], q[1], q[2], lb.font, 'middle', 9)); });
      el.innerHTML = '<svg viewBox="0 0 ' + W + ' ' + H + '" preserveAspectRatio="none">' + o.join('') + '</svg><div class="ctip"></div>';
      bindTips(el, tips);
      return;
    }
    // cartesian charts
    var horiz = series.some(function (s) { return s.kind === 'bar'; });
    var prim = [], sec = [];
    series.forEach(function (s) {
      var stackedSum = null;
      (s.axis === 2 ? sec : prim).push.apply(s.axis === 2 ? sec : prim, (s.values || []).filter(function (v) { return v !== null; }));
    });
    // stacked totals
    var stacked = series.some(function (s) { return s.stacked; });
    if (stacked) { prim = cats.map(function (c, i) { return series.filter(function (s) { return s.axis !== 2; }).reduce(function (a, s) { return a + Math.max(0, (s.values || [])[i] || 0); }, 0); }); }
    var A1 = axisOf(sp.y, prim.length ? prim : [0]), A2 = sec.length ? axisOf(sp.y2, sec) : null;
    var yf = sp.y || {}, y2f = sp.y2 || {}, xf = sp.x || {};
    var tsz = px((yf.font && yf.font.size) || 9), csz = px((xf.font && xf.font.size) || 9);
    var n = Math.max(1, cats.length || (series[0] && series[0].values ? series[0].values.length : 1));
    var catW = 0;
    if (horiz) {
      catW = Math.min((right - left) * 0.42, Math.max.apply(null, cats.map(function (c) { return tw(c, xf.font, 9); }).concat([1])) + 10);
      // value labels at the bar ends need room inside the chart
      var lblW = 0;
      series.forEach(function (s) { if ((s.kind === 'bar' || s.kind === 'column') && s.labels && s.labels.value) (s.values || []).forEach(function (v) { if (v !== null) lblW = Math.max(lblW, tw(fmt(v, s.labels.format), s.labels.font, 9)); }); });
      if (lblW) right -= lblW + 6;
      if (xf.show !== false) left += catW;
      if (yf.show !== false) top += tsz + 6;
    } else {
      if (yf.show !== false) left += Math.max.apply(null, [fmt(A1.max, yf.format), fmt(A1.min, yf.format)].map(function (t) { return t.length; })) * tsz * 0.6 + 8;
      if (A2 && y2f.show !== false) right -= Math.max.apply(null, [fmt(A2.max, y2f.format)].map(function (t) { return t.length; })) * tsz * 0.6 + 8;
      var maxLines = Math.max.apply(null, cats.map(function (c) { return wrapPx(c, (right - left) / n - 4, xf.font, 9, 2).length; }).concat([1]));
      if (series.some(function (s) { return s.labels && s.labels.value; })) top += tsz + 2;
      if (xf.show !== false) bottom -= csz * 1.25 * maxLines + 6;
    }
    var plotL = left, plotR = right, plotT = top + 6, plotB = bottom;
    if (sp.plotBg) o.push('<rect x="' + plotL + '" y="' + plotT + '" width="' + (plotR - plotL) + '" height="' + (plotB - plotT) + '" fill="' + sp.plotBg + '"/>');
    var V = function (v, ax) { ax = ax || A1; var f = (v - ax.min) / (ax.max - ax.min); return horiz ? plotL + f * (plotR - plotL) : plotB - f * (plotB - plotT); };
    // gridlines + value axis
    for (var g = A1.min; g <= A1.max + A1.major / 1000; g += A1.major) {
      var gp = V(g);
      if (yf.grid) o.push(horiz ? '<line x1="' + gp + '" y1="' + plotT + '" x2="' + gp + '" y2="' + plotB + '" stroke="' + yf.grid + '" stroke-width="1"/>' : '<line x1="' + plotL + '" y1="' + gp + '" x2="' + plotR + '" y2="' + gp + '" stroke="' + yf.grid + '" stroke-width="1"/>');
      if (yf.show !== false) o.push(horiz ? '<text x="' + gp + '" y="' + (plotT - 6) + '" text-anchor="middle" ' + fontAttr(yf.font, 9, txt) + '>' + esc(fmt(g, yf.format)) + '</text>'
        : '<text x="' + (plotL - 6) + '" y="' + (gp + tsz * 0.35) + '" text-anchor="end" ' + fontAttr(yf.font, 9, txt) + '>' + esc(fmt(g, yf.format)) + '</text>');
    }
    if (A2 && y2f.show !== false) for (var g2 = A2.min; g2 <= A2.max + A2.major / 1000; g2 += A2.major) {
      o.push('<text x="' + (plotR + 6) + '" y="' + (V(g2, A2) + tsz * 0.35) + '" ' + fontAttr(y2f.font, 9, txt) + '>' + esc(fmt(g2, y2f.format)) + '</text>');
    }
    // category axis
    var band = (horiz ? plotB - plotT : plotR - plotL) / n;
    var order = function (i) { return horiz ? (xf.reverse ? i : n - 1 - i) : (xf.reverse ? n - 1 - i : i); };
    var C = function (i) { return (horiz ? plotT : plotL) + band * (order(i) + 0.5); };
    // category names: shortened with an ellipsis to their own space (full name in the tooltip),
    // smaller when the rows are tighter than the font, and only every k-th one when still crowded
    var cf = xf.font, cfs = csz;
    if (horiz && band < csz * 1.15) { var k0 = Math.max(0.75, band / (csz * 1.15)); cf = scaled(xf.font, k0); cfs = csz * k0; }
    var every = horiz ? Math.max(1, Math.ceil(cfs * 1.1 / band)) : 1;
    if (xf.show !== false) cats.forEach(function (c, i) {
      if (i % every) return;
      if (horiz) o.push('<text x="' + (plotL - 6) + '" y="' + (C(i) + cfs * 0.35) + '" text-anchor="end" ' + fontAttr(cf, 9, txt) + '><title>' + esc(c) + '</title>' + esc(clip(c, catW - 8, cf, 9)) + '</text>');
      else wrapPx(c, band - 4, xf.font, 9, 2).forEach(function (ln, k) {
        o.push('<text x="' + C(i) + '" y="' + (plotB + csz * (k + 1.1)) + '" text-anchor="middle" ' + fontAttr(xf.font, 9, txt) + '><title>' + esc(c) + '</title>' + esc(ln) + '</text>');
      });
    });
    o.push(horiz ? '<line x1="' + V(0) + '" y1="' + plotT + '" x2="' + V(0) + '" y2="' + plotB + '" stroke="#8c8c8c"/>' : '<line x1="' + plotL + '" y1="' + V(Math.max(A1.min, 0)) + '" x2="' + plotR + '" y2="' + V(Math.max(A1.min, 0)) + '" stroke="#8c8c8c"/>');
    // bars
    var barSeries = series.filter(function (s) { return s.kind === 'column' || s.kind === 'bar'; });
    var gap = (sp.gap === undefined ? 150 : sp.gap) / 100, groupW = band / (1 + gap);
    var nb = stacked ? 1 : Math.max(1, barSeries.length), ov = (sp.overlap || 0) / 100;
    var bw = groupW / (nb - (nb - 1) * ov);
    var stackBase = cats.map(function () { return 0; });
    var labelsOut = [];
    barSeries.forEach(function (s, si) {
      (s.values || []).forEach(function (v, i) {
        if (v === null) return;
        var col = (s.points && s.points[i]) || seriesColor(s, series.indexOf(s));
        var base = stacked ? stackBase[i] : 0, end = base + v;
        if (stacked) stackBase[i] = end;
        var c0 = C(i) - groupW / 2 + (stacked ? 0 : si * bw * (1 - ov));
        var p0 = V(Math.max(A1.min, Math.min(A1.max, base))), p1 = V(Math.max(A1.min, Math.min(A1.max, end)));
        var outline = s.fillOn === false;
        var attrs = outline ? 'fill="none" stroke="' + (s.line || col) + '" stroke-width="1.5"' : 'fill="' + col + '"';
        tips.push((cats[i] || '') + ' - ' + (s.name || '') + ': ' + fmt(v, (s.labels && s.labels.format) || yf.format));
        if (horiz) o.push('<rect x="' + Math.min(p0, p1) + '" y="' + c0 + '" width="' + Math.abs(p1 - p0) + '" height="' + bw + '" ' + attrs + ' data-tip="' + (tips.length - 1) + '"/>');
        else o.push('<rect x="' + c0 + '" y="' + Math.min(p0, p1) + '" width="' + bw + '" height="' + Math.abs(p1 - p0) + '" ' + attrs + ' data-tip="' + (tips.length - 1) + '"/>');
        if (s.labels) {
          var t = [];
          if (s.labels.value) t.push(fmt(v, s.labels.format));
          if (s.labels.category) t.push(cats[i] || '');
          if (s.labels.series) t.push(s.name || '');
          if (t.length) labelsOut.push(horiz ? [Math.max(p0, p1) + 4, c0 + bw / 2 + tsz * 0.35, t.join(', '), s.labels.font, 'start']
            : [c0 + bw / 2, Math.min(p0, p1) - 4, t.join(', '), s.labels.font, 'middle']);
        }
      });
    });
    // lines / areas
    series.forEach(function (s, si) {
      if (s.kind !== 'line' && s.kind !== 'area') return;
      var ax = s.axis === 2 && A2 ? A2 : A1, col = seriesColor(s, si), pts = [];
      (s.values || []).forEach(function (v, i) { if (v !== null) pts.push([C(i), V(Math.max(ax.min, Math.min(ax.max, v)), ax), v, i]); });
      if (!pts.length) return;
      var d = 'M' + pts.map(function (q) { return q[0].toFixed(2) + ',' + q[1].toFixed(2); }).join(' L');
      if (s.kind === 'area') o.push('<path d="' + d + ' L' + pts[pts.length - 1][0] + ',' + V(Math.max(ax.min, 0), ax) + ' L' + pts[0][0] + ',' + V(Math.max(ax.min, 0), ax) + 'Z" fill="' + (s.fill || col) + '" opacity=".6"/>');
      if (s.lineOn !== false) o.push('<path d="' + d + '" fill="none" stroke="' + col + '" stroke-width="' + Math.max(1, px(s.width || 1.5)) + '" stroke-linejoin="round" stroke-linecap="round"/>');
      pts.forEach(function (q) {
        tips.push((cats[q[3]] || '') + ' - ' + (s.name || '') + ': ' + fmt(q[2], (s.labels && s.labels.format) || (s.axis === 2 ? y2f.format : yf.format)));
        o.push('<circle cx="' + q[0] + '" cy="' + q[1] + '" r="' + (s.marker ? 4 : 6) + '" fill="' + (s.marker ? (s.markerColor || col) : 'transparent') + '" data-tip="' + (tips.length - 1) + '"/>');
        if (s.labels && s.labels.value) labelsOut.push([q[0], q[1] - 6, fmt(q[2], s.labels.format), s.labels.font, 'middle']);
      });
    });
    // value labels never overlap each other: a label that would is left to the tooltip
    labelsOut.forEach(function (q) { o.push(put(q[0], q[1], q[2], q[3], q[4], 9)); });
    el.innerHTML = '<svg viewBox="0 0 ' + W + ' ' + H + '" preserveAspectRatio="none">' + o.join('') + '</svg><div class="ctip"></div>';
    bindTips(el, tips);
  }
  function bindTips(el, tips) {
    var tip = el.querySelector('.ctip');
    el.addEventListener('mousemove', function (e) {
      var t = e.target.getAttribute && e.target.getAttribute('data-tip');
      if (t === null || t === undefined) { tip.style.display = 'none'; return; }
      var r = el.getBoundingClientRect(), sc = r.width / el.offsetWidth || 1;
      tip.textContent = tips[+t]; tip.style.display = 'block';
      tip.style.left = ((e.clientX - r.left) / sc + 12) + 'px'; tip.style.top = ((e.clientY - r.top) / sc + 12) + 'px';
    });
    el.addEventListener('mouseleave', function () { tip.style.display = 'none'; });
  }
  [].slice.call(document.querySelectorAll('.rchart[data-chart]')).forEach(function (el) {
    try { renderChart(el); }
    catch (e) { el.setAttribute('data-error', String(e)); }   // Excel's chart image stays as fallback
  });

  function fit() {
    var present = body.classList.contains('present');
    pages.forEach(function (p) {
      var wrap = p.querySelector('.cvwrap'), cv = p.querySelector('.cv');
      if (!wrap || !cv) return;
      var W = parseFloat(cv.dataset.w), H = parseFloat(cv.dataset.h);
      if (cv.dataset.h === 'auto') {
        H = cv.offsetHeight;
        var wPt = W * 0.75, hPt = H * 0.75;
        wrap.style.setProperty('--w', wPt + 'pt'); wrap.style.setProperty('--h', hPt + 'pt');
        wrap.style.setProperty('--pz', Math.min(1, 790 / wPt, 520 / hPt));
      }
      var avail = p.querySelector('.pbody').clientWidth - 12;
      if (present) avail = window.innerWidth - 12;
      var s = avail / W * (present || zoom === 'sharp' ? 1 : zoom);
      if (shot) s = 1;
      if (!present && zoom === 'sharp') { var ss = sharpScale(); if (ss > 0) s = ss; }
      if (present) {
        var ah = window.innerHeight - p.querySelector('.pbar').offsetHeight - p.querySelector('.pf').offsetHeight - 14;
        s = Math.min(s, ah / H);
      }
      if (!(s > 0)) return;
      cv.style.transform = 'scale(' + s + ')';
      wrap.style.width = (W * s) + 'px';
      wrap.style.height = (H * s) + 'px';
    });
  }
  // wrapped text that needs a little more room than its Excel box is shrunk slightly (at most 15%)
  // so it does not run over the next object; the wording is never changed
  function fitText() {
    [].slice.call(document.querySelectorAll('.shp > .tx')).forEach(function (tx) {
      try {
        if (tx.parentNode.classList.contains('nowrap')) return;
        var els = [tx].concat([].slice.call(tx.querySelectorAll('span')));
        els.forEach(function (el) {
          if (!el.getAttribute('data-fs')) el.setAttribute('data-fs', parseFloat(getComputedStyle(el).fontSize));
          el.style.fontSize = el.getAttribute('data-fs') + 'px';
        });
        for (var f = 0.97; f >= 0.849 && tx.scrollHeight > tx.clientHeight + 1; f -= 0.03) {
          els.forEach(function (el) { el.style.fontSize = (parseFloat(el.getAttribute('data-fs')) * f) + 'px'; });
        }
      } catch (e) { }
    });
  }
  fitText();
  window.addEventListener('resize', fit);
  fit();
  if (document.fonts && document.fonts.ready) document.fonts.ready.then(function () { fitText(); fit(); });

  // page list is a pop-out, so pages always use the whole window width
  document.getElementById('btnMenu').addEventListener('click', function () { body.classList.toggle('sb'); });
  [].slice.call(document.querySelectorAll('#nav a')).forEach(function (a) { a.addEventListener('click', function () { body.classList.remove('sb'); }); });
  document.getElementById('zIn').addEventListener('click', function () { setZoom(zoom + 0.25); });
  document.getElementById('zOut').addEventListener('click', function () { setZoom(zoom - 0.25); });
  document.getElementById('zFit').addEventListener('click', function () { setZoom(1); });
  window.addEventListener('load', fit);
  setZoom(zoom);

  // navigation highlight
  var links = [].slice.call(document.querySelectorAll('#nav a'));
  if ('IntersectionObserver' in window) {
    var io = new IntersectionObserver(function (es) {
      es.forEach(function (e) {
        if (e.isIntersecting) links.forEach(function (l) { l.classList.toggle('active', l.getAttribute('href') === '#' + e.target.id); });
      });
    }, { rootMargin: '-40% 0px -55% 0px' });
    pages.forEach(function (p) { io.observe(p); });
  }

  // presentation mode
  var cur = 0;
  function show(i) {
    cur = Math.max(0, Math.min(pages.length - 1, i));
    pages.forEach(function (p, k) { p.classList.toggle('current', k === cur); });
    document.getElementById('hudPos').textContent = (cur + 1) + ' / ' + pages.length;
    fit();
  }
  function startPresent() {
    body.classList.add('present');
    var a = -1; links.forEach(function (l, k) { if (l.classList.contains('active')) a = k; });
    show(a < 0 ? 0 : a);
    if (document.documentElement.requestFullscreen) document.documentElement.requestFullscreen().catch(function () { });
  }
  function stopPresent() {
    body.classList.remove('present');
    pages.forEach(function (p) { p.classList.remove('current'); });
    if (document.fullscreenElement) document.exitFullscreen();
    fit(); pages[cur].scrollIntoView();
  }
  document.getElementById('btnPresent').addEventListener('click', startPresent);
  document.getElementById('hudPrev').addEventListener('click', function () { show(cur - 1); });
  document.getElementById('hudNext').addEventListener('click', function () { show(cur + 1); });
  document.getElementById('hudExit').addEventListener('click', stopPresent);
  document.addEventListener('fullscreenchange', function () { if (!document.fullscreenElement && body.classList.contains('present')) stopPresent(); else fit(); });
  document.addEventListener('keydown', function (e) {
    if (!body.classList.contains('present')) {
      if ((e.key === '+' || e.key === '=') && !e.ctrlKey) { setZoom(zoom + 0.25); return; }
      if (e.key === '-' && !e.ctrlKey) { setZoom(zoom - 0.25); return; }
      if ((e.key === 'p' || e.key === 'P') && !e.ctrlKey && !e.metaKey && !e.altKey) { e.preventDefault(); startPresent(); }
      return;
    }
    if (['ArrowRight', 'PageDown', ' '].indexOf(e.key) >= 0) { e.preventDefault(); show(cur + 1); }
    if (['ArrowLeft', 'PageUp'].indexOf(e.key) >= 0) { e.preventDefault(); show(cur - 1); }
    if (e.key === 'Home') show(0);
    if (e.key === 'End') show(pages.length - 1);
    if (e.key === 'Escape') stopPresent();
  });
})();
'@

# ======================================================================================
#  Edge headless PDF of the dashboard (optional, local only)
# ======================================================================================

# ======================================================================================
#  Main
# ======================================================================================

# ======================================================================================
#  PDF and PowerPoint copies: one picture per dashboard page, taken from the finished HTML
# ======================================================================================

function Find-HeadlessBrowser {
    # Microsoft Edge (part of Windows), then Google Chrome. Only used to open the local HTML file and
    # take a picture of each page; no web address is ever opened. Returns every one found.
    $candidates = New-Object System.Collections.Generic.List[string]
    foreach ($exe in @('Microsoft\Edge\Application\msedge.exe', 'Google\Chrome\Application\chrome.exe')) {
        foreach ($base in @(${env:ProgramFiles(x86)}, $env:ProgramFiles, $env:LOCALAPPDATA)) {
            if ($base) { $candidates.Add((Join-Path $base $exe)) }
        }
        foreach ($hive in @('HKLM:', 'HKCU:')) {
            try { $v = (Get-ItemProperty -LiteralPath ("$hive\SOFTWARE\Microsoft\Windows\CurrentVersion\App Paths\" + (Split-Path $exe -Leaf)) -ErrorAction Stop).'(default)'; if ($v) { $candidates.Add([string]$v) } } catch { }
        }
    }
    $found = New-Object System.Collections.Generic.List[string]
    foreach ($c in $candidates) {
        if (-not $c -or -not (Test-Path -LiteralPath $c -PathType Leaf)) { continue }
        $full = (Resolve-Path -LiteralPath $c).ProviderPath
        if (-not ($found | Where-Object { $_ -ieq $full })) { $found.Add($full) }
    }
    return , $found
}

function Get-HeadlessPolicyNote {
    # Organisations can switch off the hidden (headless) mode of Edge/Chrome by policy; report it.
    param([string]$Browser)
    $key = if ($Browser -match 'msedge') { 'Microsoft\Edge' } else { 'Google\Chrome' }
    foreach ($hive in @('HKLM:', 'HKCU:')) {
        try {
            $v = (Get-ItemProperty -LiteralPath "$hive\SOFTWARE\Policies\$key" -ErrorAction Stop).HeadlessModeEnabled
            if ($null -ne $v) {
                $what = ''
                if ([int]$v -eq 0) { $what = ' (hidden mode switched off by the organisation)' }
                return ('policy {0}\SOFTWARE\Policies\{1}\HeadlessModeEnabled = {2}{3}' -f $hive, $key, $v, $what)
            }
        }
        catch { }
    }
    return 'no HeadlessModeEnabled policy set'
}

function Stop-ShotBrowsers {
    # Closes browser copies this script started that are still running: only processes whose command
    # line contains this run's own private folder; the user's own Edge/Chrome windows are never touched.
    param([string]$ShotDir)
    try {
        $mine = @(Get-CimInstance -ClassName Win32_Process -ErrorAction Stop | Where-Object { $_.CommandLine -and $_.CommandLine.IndexOf($ShotDir, [System.StringComparison]::OrdinalIgnoreCase) -ge 0 })
        foreach ($pr in $mine) { try { Stop-Process -Id $pr.ProcessId -Force -ErrorAction SilentlyContinue } catch { } }
        if ($mine.Count) { Write-Log ("{0} browser process(es) started by this script for the page pictures were closed." -f $mine.Count) }
    }
    catch { }
}

function Invoke-HeadlessShot {
    # Runs the browser once; returns $null when the picture was written, else the reason, including the
    # browser's exit code and its own error output.
    param([string]$Browser, [string]$Arguments, [string]$Raw)
    $psi = New-Object System.Diagnostics.ProcessStartInfo
    $psi.FileName = $Browser
    $psi.Arguments = $Arguments
    $psi.UseShellExecute = $false
    $psi.CreateNoWindow = $true
    $psi.RedirectStandardError = $true
    $psi.RedirectStandardOutput = $true
    $proc = [System.Diagnostics.Process]::Start($psi)
    $errTask = $proc.StandardError.ReadToEndAsync()
    $outTask = $proc.StandardOutput.ReadToEndAsync()
    if (-not $proc.WaitForExit(120000)) {
        try { $proc.Kill() } catch { }   # only the browser this script started
        return 'browser did not finish in 2 minutes'
    }
    # When PowerShell runs as Administrator, Edge hands the job to a non-administrator copy of itself
    # and exits at once; that copy writes the picture a few seconds later. Wait for it (up to 60 s)
    # until the file exists and has stopped growing.
    $last = -1
    for ($t = 0; $t -lt 120; $t++) {
        $len = -1
        if (Test-Path -LiteralPath $Raw) { try { $len = (Get-Item -LiteralPath $Raw).Length } catch { } }
        if ($len -gt 0 -and $len -eq $last) { return $null }
        $last = $len
        Start-Sleep -Milliseconds 500
    }
    $text = ''
    try { if ($errTask.Wait(3000)) { $text += $errTask.Result } } catch { }
    try { if ($outTask.Wait(1000)) { $text += $outTask.Result } } catch { }
    $text = (($text -split "`r?`n") | Where-Object { $_.Trim() } | Select-Object -Last 6) -join ' / '
    if ($text.Length -gt 600) { $text = $text.Substring($text.Length - 600) }
    if (-not $text) { $text = 'no message from the browser' }
    return ('no picture written (exit code {0}; browser said: {1})' -f $proc.ExitCode, $text)
}

function Save-CroppedPng {
    # Keeps the top-left $W x $H pixels (the page itself) of the browser's picture.
    # The dashboard's check mark at ($MarkX, $MarkY) must be present, so a browser start page or an
    # error page can never end up in the PDF / PowerPoint.
    param([string]$Source, [string]$Target, [int]$W, [int]$H, [int]$MarkX, [int]$MarkY)
    Add-Type -AssemblyName System.Drawing
    $img = [System.Drawing.Bitmap]::FromFile($Source)
    try {
        if ($img.Width -lt $W -or $img.Height -lt $H) { throw ('picture is {0}x{1}, smaller than the page {2}x{3}' -f $img.Width, $img.Height, $W, $H) }
        $c = $img.GetPixel([math]::Min($MarkX, $img.Width - 1), [math]::Min($MarkY, $img.Height - 1))
        if ([math]::Abs($c.R - 1) + [math]::Abs($c.G - 2) + [math]::Abs($c.B - 3) -gt 9) { throw 'the browser did not show the dashboard page (check mark missing)' }
        $crop = $img.Clone((New-Object System.Drawing.Rectangle 0, 0, $W, $H), [System.Drawing.Imaging.PixelFormat]::Format24bppRgb)
        try { $crop.Save($Target, [System.Drawing.Imaging.ImageFormat]::Png) } finally { $crop.Dispose() }
    }
    finally { $img.Dispose() }
}

function Export-PageShots {
    # Opens the finished dashboard in a hidden browser window (?shot=N shows page N alone at its exact
    # size) and saves one picture per page at twice the screen resolution. The window is made larger
    # than the page and the picture cropped to the page, so nothing at the edges can be cut off.
    # Page 1 decides which browser/mode works; when none does, the reasons are reported and the rest skipped.
    param([string]$HtmlPath, $Pages, [string]$ShotDir, [ref]$ErrorText)
    $shots = New-Object System.Collections.Generic.List[object]
    $browsers = Find-HeadlessBrowser
    if ($browsers.Count -eq 0) { $ErrorText.Value = 'Microsoft Edge / Google Chrome was not found on this PC'; return , $shots }
    New-Item -ItemType Directory -Path $ShotDir -Force | Out-Null
    $url = (New-Object System.Uri($HtmlPath)).AbsoluteUri
    $scale = 2
    $errs = New-Object System.Collections.Generic.List[string]
    $working = $null
    foreach ($p in $Pages) {
        $wPx = [int][math]::Ceiling([double]$p.Width * 4 / 3)
        $hPx = [int][math]::Ceiling(([double]$p.EndY - [double]$p.StartY) * 4 / 3)
        $raw = Join-Path $ShotDir ('raw-{0:D2}.png' -f $p.Index)
        $png = Join-Path $ShotDir ('page-{0:D2}.png' -f $p.Index)
        $combos = New-Object System.Collections.Generic.List[object]
        if ($working) { $combos.Add($working) }
        else { foreach ($br in $browsers) { foreach ($mode in @('--headless=new', '--headless')) { $combos.Add([pscustomobject]@{ Browser = $br; Mode = $mode }) } } }
        $ok = $false
        foreach ($cb in $combos) {
            if (Test-Path -LiteralPath $raw) { Remove-Item -LiteralPath $raw -Force }
            # a fresh profile folder per attempt: a copy of the browser still finishing the previous
            # page can never block this one ("profile in use", exit code 21)
            $profileDir = Join-Path $ShotDir ('browser-profile-' + [guid]::NewGuid().ToString('N').Substring(0, 8))
            $argList = @($cb.Mode, '--do-not-de-elevate', '--disable-gpu', '--no-first-run', '--no-default-browser-check', '--hide-scrollbars', '--disable-extensions',
                '--disable-sync', '--disable-background-networking', '--disable-component-update', '--disable-default-apps', '--mute-audio',
                '--enable-logging=stderr', '--v=0',
                # no network at all for this hidden browser: the dashboard is a local file with no
                # external references, and the browser's own background calls (updates etc.) go nowhere
                '--proxy-server=http://127.0.0.1:9', '--no-pings', '--disable-domain-reliability', '--disable-client-side-phishing-detection',
                ('"--user-data-dir={0}"' -f $profileDir), ('--force-device-scale-factor={0}' -f $scale), ('--window-size={0},{1}' -f ($wPx + 200), ($hPx + 200)),
                '--virtual-time-budget=5000', ('"--screenshot={0}"' -f $raw), ('"{0}?shot={1}"' -f $url, $p.Index))
            $who = '{0} {1}' -f (Split-Path $cb.Browser -Leaf), $cb.Mode
            try {
                $why = Invoke-HeadlessShot -Browser $cb.Browser -Arguments ($argList -join ' ') -Raw $raw
                if ($why) { $errs.Add(('page {0} ({1}): {2}' -f $p.Index, $who, $why)); continue }
                $null = Save-CroppedPng -Source $raw -Target $png -W ($wPx * $scale) -H ($hPx * $scale) -MarkX (($wPx + 25) * $scale) -MarkY (25 * $scale)
                if (-not $working) { $working = $cb; Write-Log ("PDF/PowerPoint pages: pictures taken with {0} ({1}, hidden, local file only)." -f $cb.Browser, $cb.Mode) }
                $ok = $true; break
            }
            catch { $errs.Add(('page {0} ({1}): {2}' -f $p.Index, $who, $_.Exception.Message)) }
        }
        if ($ok) {
            $shots.Add([pscustomobject]@{ Index = $p.Index; Title = [string]$p.Title; Path = $png; WPx = $wPx * $scale; HPx = $hPx * $scale; WPt = $wPx * 0.75; HPt = $hPx * 0.75 })
        }
        elseif (-not $working) {
            foreach ($br in $browsers) { $errs.Add(('{0}: {1}' -f (Split-Path $br -Leaf), (Get-HeadlessPolicyNote $br))) }
            break   # no browser works on this PC: the other pages would fail the same way
        }
    }
    Stop-ShotBrowsers -ShotDir $ShotDir
    Get-ChildItem -LiteralPath $ShotDir -Directory -Filter 'browser-profile-*' -ErrorAction SilentlyContinue | Remove-Item -Recurse -Force -ErrorAction SilentlyContinue
    $ErrorText.Value = $null
    if ($errs.Count) { $ErrorText.Value = ($errs -join ' | ') }
    return , $shots
}

function ConvertTo-JpegFile {
    # PNG page picture -> high-quality JPEG (keeps the PDF small; PDF readers show JPEG directly).
    param([string]$PngPath, [string]$JpegPath, [int]$Quality = 92)
    Add-Type -AssemblyName System.Drawing
    $img = [System.Drawing.Image]::FromFile($PngPath)
    try {
        $codec = [System.Drawing.Imaging.ImageCodecInfo]::GetImageEncoders() | Where-Object { $_.MimeType -eq 'image/jpeg' } | Select-Object -First 1
        $ep = New-Object System.Drawing.Imaging.EncoderParameters 1
        $ep.Param[0] = New-Object System.Drawing.Imaging.EncoderParameter ([System.Drawing.Imaging.Encoder]::Quality), ([int64]$Quality)
        $img.Save($JpegPath, $codec, $ep)
    }
    finally { $img.Dispose() }
}

function Write-ImagePdf {
    # A PDF with one page per dashboard page; every PDF page has exactly that page's size, so the
    # whole page is shown with nothing cut off. $Shots items: Jpeg, WPx, HPx, WPt, HPt, Title.
    param($Shots, [string]$Path, [string]$Title)
    $inv = [System.Globalization.CultureInfo]::InvariantCulture
    $fs = [System.IO.File]::Create($Path)
    try {
        $offsets = New-Object System.Collections.Generic.List[long]
        $ascii = [System.Text.Encoding]::ASCII
        $put = { param([string]$t) $b = $ascii.GetBytes($t); $fs.Write($b, 0, $b.Length) }
        $obj = { param([int]$n) while ($offsets.Count -lt $n) { $offsets.Add(0) }; $offsets[$n - 1] = $fs.Position; & $put ("{0} 0 obj`n" -f $n) }
        & $put "%PDF-1.4`n%"
        $fs.Write([byte[]](0xE2, 0xE3, 0xCF, 0xD3, 0x0A), 0, 5)
        $n = $Shots.Count
        # objects: 1 catalog, 2 pages, 3 info, then 3 per page (page, content, image)
        $kids = (0..($n - 1) | ForEach-Object { '{0} 0 R' -f (4 + 3 * $_) }) -join ' '
        & $obj 1; & $put "<< /Type /Catalog /Pages 2 0 R >>`nendobj`n"
        & $obj 2; & $put ("<< /Type /Pages /Kids [{0}] /Count {1} >>`nendobj`n" -f $kids, $n)
        $safeTitle = ($Title -replace '[^\x20-\x7E]', '' -replace '([\\()])', '\$1')
        & $obj 3; & $put ("<< /Title ({0}) /Producer (Export-WklyReportDashboard) >>`nendobj`n" -f $safeTitle)
        for ($i = 0; $i -lt $n; $i++) {
            $s = $Shots[$i]
            $pg = 4 + 3 * $i
            $w = [string]::Format($inv, '{0:0.##}', [double]$s.WPt); $h = [string]::Format($inv, '{0:0.##}', [double]$s.HPt)
            $content = "q $w 0 0 $h 0 0 cm /Im0 Do Q"
            $jpeg = [System.IO.File]::ReadAllBytes($s.Jpeg)
            & $obj $pg; & $put ("<< /Type /Page /Parent 2 0 R /MediaBox [0 0 {0} {1}] /Resources << /XObject << /Im0 {2} 0 R >> >> /Contents {3} 0 R >>`nendobj`n" -f $w, $h, ($pg + 2), ($pg + 1))
            & $obj ($pg + 1); & $put ("<< /Length {0} >>`nstream`n{1}`nendstream`nendobj`n" -f $content.Length, $content)
            & $obj ($pg + 2); & $put ("<< /Type /XObject /Subtype /Image /Width {0} /Height {1} /ColorSpace /DeviceRGB /BitsPerComponent 8 /Filter /DCTDecode /Length {2} >>`nstream`n" -f $s.WPx, $s.HPx, $jpeg.Length)
            $fs.Write($jpeg, 0, $jpeg.Length)
            & $put "`nendstream`nendobj`n"
        }
        $xref = $fs.Position
        & $put ("xref`n0 {0}`n0000000000 65535 f `n" -f ($offsets.Count + 1))
        foreach ($o in $offsets) { & $put ("{0:D10} 00000 n `n" -f $o) }
        & $put ("trailer`n<< /Size {0} /Root 1 0 R /Info 3 0 R >>`nstartxref`n{1}`n%%EOF`n" -f ($offsets.Count + 1), $xref)
    }
    finally { $fs.Dispose() }
}

function Write-ImagePptx {
    # A PowerPoint file (16:9) with one slide per dashboard page. Each page picture is fitted whole
    # into its slide (never cropped) and centred. Written directly as Open XML: PowerPoint is not needed.
    param($Shots, [string]$Path, [string]$Title)
    Add-Type -AssemblyName System.IO.Compression, System.IO.Compression.FileSystem
    $SW = 12192000; $SH = 6858000                     # 13.333 x 7.5 in
    $nsA = 'http://schemas.openxmlformats.org/drawingml/2006/main'
    $nsR = 'http://schemas.openxmlformats.org/officeDocument/2006/relationships'
    $nsP = 'http://schemas.openxmlformats.org/presentationml/2006/main'
    $relNs = 'http://schemas.openxmlformats.org/package/2006/relationships'
    $rt = 'http://schemas.openxmlformats.org/officeDocument/2006/relationships'
    $hdr = '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>' + "`r`n"
    $esc = { param($t) [System.Security.SecurityElement]::Escape([string]$t) }
    $tree = '<p:nvGrpSpPr><p:cNvPr id="1" name=""/><p:cNvGrpSpPr/><p:nvPr/></p:nvGrpSpPr><p:grpSpPr><a:xfrm><a:off x="0" y="0"/><a:ext cx="0" cy="0"/><a:chOff x="0" y="0"/><a:chExt cx="0" cy="0"/></a:xfrm></p:grpSpPr>'
    $n = $Shots.Count
    $parts = [ordered]@{}
    $ct = '<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types"><Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/><Default Extension="xml" ContentType="application/xml"/><Default Extension="png" ContentType="image/png"/>' +
        '<Override PartName="/ppt/presentation.xml" ContentType="application/vnd.openxmlformats-officedocument.presentationml.presentation.main+xml"/>' +
        '<Override PartName="/ppt/slideMasters/slideMaster1.xml" ContentType="application/vnd.openxmlformats-officedocument.presentationml.slideMaster+xml"/>' +
        '<Override PartName="/ppt/slideLayouts/slideLayout1.xml" ContentType="application/vnd.openxmlformats-officedocument.presentationml.slideLayout+xml"/>' +
        '<Override PartName="/ppt/theme/theme1.xml" ContentType="application/vnd.openxmlformats-officedocument.theme+xml"/>' +
        '<Override PartName="/ppt/presProps.xml" ContentType="application/vnd.openxmlformats-officedocument.presentationml.presProps+xml"/>' +
        '<Override PartName="/ppt/viewProps.xml" ContentType="application/vnd.openxmlformats-officedocument.presentationml.viewProps+xml"/>' +
        '<Override PartName="/ppt/tableStyles.xml" ContentType="application/vnd.openxmlformats-officedocument.presentationml.tableStyles+xml"/>' +
        '<Override PartName="/docProps/core.xml" ContentType="application/vnd.openxmlformats-package.core-properties+xml"/>' +
        '<Override PartName="/docProps/app.xml" ContentType="application/vnd.openxmlformats-officedocument.extended-properties+xml"/>'
    for ($i = 1; $i -le $n; $i++) { $ct += '<Override PartName="/ppt/slides/slide{0}.xml" ContentType="application/vnd.openxmlformats-officedocument.presentationml.slide+xml"/>' -f $i }
    $parts['[Content_Types].xml'] = $ct + '</Types>'
    $parts['_rels/.rels'] = "<Relationships xmlns=`"$relNs`"><Relationship Id=`"rId1`" Type=`"$rt/officeDocument`" Target=`"ppt/presentation.xml`"/><Relationship Id=`"rId2`" Type=`"http://schemas.openxmlformats.org/package/2006/relationships/metadata/core-properties`" Target=`"docProps/core.xml`"/><Relationship Id=`"rId3`" Type=`"$rt/extended-properties`" Target=`"docProps/app.xml`"/></Relationships>"
    $now = (Get-Date).ToUniversalTime().ToString('yyyy-MM-ddTHH:mm:ssZ')
    $parts['docProps/core.xml'] = '<cp:coreProperties xmlns:cp="http://schemas.openxmlformats.org/package/2006/metadata/core-properties" xmlns:dc="http://purl.org/dc/elements/1.1/" xmlns:dcterms="http://purl.org/dc/terms/" xmlns:xsi="http://www.w3.org/2001/XMLSchema-instance"><dc:title>' + (& $esc $Title) + '</dc:title><dc:creator>Export-WklyReportDashboard</dc:creator><dcterms:created xsi:type="dcterms:W3CDTF">' + $now + '</dcterms:created><dcterms:modified xsi:type="dcterms:W3CDTF">' + $now + '</dcterms:modified></cp:coreProperties>'
    $parts['docProps/app.xml'] = '<Properties xmlns="http://schemas.openxmlformats.org/officeDocument/2006/extended-properties"><Application>Microsoft Office PowerPoint</Application><Slides>' + $n + '</Slides></Properties>'
    $sldIds = ''; $presRels = "<Relationship Id=`"rId1`" Type=`"$rt/slideMaster`" Target=`"slideMasters/slideMaster1.xml`"/><Relationship Id=`"rId2`" Type=`"$rt/theme`" Target=`"theme/theme1.xml`"/><Relationship Id=`"rId3`" Type=`"$rt/presProps`" Target=`"presProps.xml`"/><Relationship Id=`"rId4`" Type=`"$rt/viewProps`" Target=`"viewProps.xml`"/><Relationship Id=`"rId5`" Type=`"$rt/tableStyles`" Target=`"tableStyles.xml`"/>"
    for ($i = 1; $i -le $n; $i++) {
        $sldIds += '<p:sldId id="{0}" r:id="rId{1}"/>' -f (255 + $i), (9 + $i)
        $presRels += "<Relationship Id=`"rId$(9 + $i)`" Type=`"$rt/slide`" Target=`"slides/slide$i.xml`"/>"
    }
    $parts['ppt/presentation.xml'] = "<p:presentation xmlns:a=`"$nsA`" xmlns:r=`"$nsR`" xmlns:p=`"$nsP`" saveSubsetFonts=`"1`"><p:sldMasterIdLst><p:sldMasterId id=`"2147483648`" r:id=`"rId1`"/></p:sldMasterIdLst><p:sldIdLst>$sldIds</p:sldIdLst><p:sldSz cx=`"$SW`" cy=`"$SH`"/><p:notesSz cx=`"6858000`" cy=`"9144000`"/></p:presentation>"
    $parts['ppt/_rels/presentation.xml.rels'] = "<Relationships xmlns=`"$relNs`">$presRels</Relationships>"
    $parts['ppt/presProps.xml'] = "<p:presentationPr xmlns:a=`"$nsA`" xmlns:r=`"$nsR`" xmlns:p=`"$nsP`"/>"
    $parts['ppt/viewProps.xml'] = "<p:viewPr xmlns:a=`"$nsA`" xmlns:r=`"$nsR`" xmlns:p=`"$nsP`"><p:normalViewPr><p:restoredLeft sz=`"15620`"/><p:restoredTop sz=`"94660`"/></p:normalViewPr><p:gridSpacing cx=`"76200`" cy=`"76200`"/></p:viewPr>"
    $parts['ppt/tableStyles.xml'] = "<a:tblStyleLst xmlns:a=`"$nsA`" def=`"{5C22544A-7EE6-4342-B048-85BDC9FD1C3A}`"/>"
    $parts['ppt/slideMasters/slideMaster1.xml'] = "<p:sldMaster xmlns:a=`"$nsA`" xmlns:r=`"$nsR`" xmlns:p=`"$nsP`"><p:cSld><p:bg><p:bgPr><a:solidFill><a:srgbClr val=`"FFFFFF`"/></a:solidFill><a:effectLst/></p:bgPr></p:bg><p:spTree>$tree</p:spTree></p:cSld><p:clrMap bg1=`"lt1`" tx1=`"dk1`" bg2=`"lt2`" tx2=`"dk2`" accent1=`"accent1`" accent2=`"accent2`" accent3=`"accent3`" accent4=`"accent4`" accent5=`"accent5`" accent6=`"accent6`" hlink=`"hlink`" folHlink=`"folHlink`"/><p:sldLayoutIdLst><p:sldLayoutId id=`"2147483649`" r:id=`"rId1`"/></p:sldLayoutIdLst></p:sldMaster>"
    $parts['ppt/slideMasters/_rels/slideMaster1.xml.rels'] = "<Relationships xmlns=`"$relNs`"><Relationship Id=`"rId1`" Type=`"$rt/slideLayout`" Target=`"../slideLayouts/slideLayout1.xml`"/><Relationship Id=`"rId2`" Type=`"$rt/theme`" Target=`"../theme/theme1.xml`"/></Relationships>"
    $parts['ppt/slideLayouts/slideLayout1.xml'] = "<p:sldLayout xmlns:a=`"$nsA`" xmlns:r=`"$nsR`" xmlns:p=`"$nsP`" type=`"blank`" preserve=`"1`"><p:cSld name=`"Blank`"><p:spTree>$tree</p:spTree></p:cSld><p:clrMapOvr><a:masterClrMapping/></p:clrMapOvr></p:sldLayout>"
    $parts['ppt/slideLayouts/_rels/slideLayout1.xml.rels'] = "<Relationships xmlns=`"$relNs`"><Relationship Id=`"rId1`" Type=`"$rt/slideMaster`" Target=`"../slideMasters/slideMaster1.xml`"/></Relationships>"
    $fill3 = '<a:solidFill><a:schemeClr val="phClr"/></a:solidFill>'
    $line = '<a:ln w="9525"><a:solidFill><a:schemeClr val="phClr"/></a:solidFill></a:ln>'
    $clr = { param($nm, $v) "<a:$nm><a:srgbClr val=`"$v`"/></a:$nm>" }
    $parts['ppt/theme/theme1.xml'] = "<a:theme xmlns:a=`"$nsA`" name=`"Office Theme`"><a:themeElements><a:clrScheme name=`"Office`"><a:dk1><a:sysClr val=`"windowText`" lastClr=`"000000`"/></a:dk1><a:lt1><a:sysClr val=`"window`" lastClr=`"FFFFFF`"/></a:lt1>" +
        (& $clr 'dk2' '44546A') + (& $clr 'lt2' 'E7E6E6') + (& $clr 'accent1' '4472C4') + (& $clr 'accent2' 'ED7D31') + (& $clr 'accent3' 'A5A5A5') + (& $clr 'accent4' 'FFC000') + (& $clr 'accent5' '5B9BD5') + (& $clr 'accent6' '70AD47') + (& $clr 'hlink' '0563C1') + (& $clr 'folHlink' '954F72') +
        "</a:clrScheme><a:fontScheme name=`"Office`"><a:majorFont><a:latin typeface=`"Calibri Light`"/><a:ea typeface=`"`"/><a:cs typeface=`"`"/></a:majorFont><a:minorFont><a:latin typeface=`"Calibri`"/><a:ea typeface=`"`"/><a:cs typeface=`"`"/></a:minorFont></a:fontScheme>" +
        "<a:fmtScheme name=`"Office`"><a:fillStyleLst>$fill3$fill3$fill3</a:fillStyleLst><a:lnStyleLst>$line$line$line</a:lnStyleLst><a:effectStyleLst><a:effectStyle><a:effectLst/></a:effectStyle><a:effectStyle><a:effectLst/></a:effectStyle><a:effectStyle><a:effectLst/></a:effectStyle></a:effectStyleLst><a:bgFillStyleLst>$fill3$fill3$fill3</a:bgFillStyleLst></a:fmtScheme></a:themeElements></a:theme>"
    for ($i = 1; $i -le $n; $i++) {
        $s = $Shots[$i - 1]
        $wE = [double]$s.WPt * 12700; $hE = [double]$s.HPt * 12700
        $k = [math]::Min($SW / $wE, $SH / $hE)
        $cx = [int64][math]::Floor($wE * $k); $cy = [int64][math]::Floor($hE * $k)
        $x = [int64](($SW - $cx) / 2); $y = [int64](($SH - $cy) / 2)
        $parts["ppt/slides/slide$i.xml"] = "<p:sld xmlns:a=`"$nsA`" xmlns:r=`"$nsR`" xmlns:p=`"$nsP`"><p:cSld><p:spTree>$tree<p:pic><p:nvPicPr><p:cNvPr id=`"2`" name=`"Page $($s.Index)`" descr=`"$(& $esc $s.Title)`"/><p:cNvPicPr><a:picLocks noChangeAspect=`"1`"/></p:cNvPicPr><p:nvPr/></p:nvPicPr><p:blipFill><a:blip r:embed=`"rId2`"/><a:stretch><a:fillRect/></a:stretch></p:blipFill><p:spPr><a:xfrm><a:off x=`"$x`" y=`"$y`"/><a:ext cx=`"$cx`" cy=`"$cy`"/></a:xfrm><a:prstGeom prst=`"rect`"><a:avLst/></a:prstGeom></p:spPr></p:pic></p:spTree></p:cSld><p:clrMapOvr><a:masterClrMapping/></p:clrMapOvr></p:sld>"
        $parts["ppt/slides/_rels/slide$i.xml.rels"] = "<Relationships xmlns=`"$relNs`"><Relationship Id=`"rId1`" Type=`"$rt/slideLayout`" Target=`"../slideLayouts/slideLayout1.xml`"/><Relationship Id=`"rId2`" Type=`"$rt/image`" Target=`"../media/image$i.png`"/></Relationships>"
    }
    if (Test-Path -LiteralPath $Path) { Remove-Item -LiteralPath $Path -Force }
    $utf8 = New-Object System.Text.UTF8Encoding($false)
    $zip = [System.IO.Compression.ZipFile]::Open($Path, [System.IO.Compression.ZipArchiveMode]::Create)
    try {
        foreach ($name in $parts.Keys) {
            $e = $zip.CreateEntry($name)
            $writer = New-Object System.IO.StreamWriter($e.Open(), $utf8)
            try { $writer.Write($hdr + $parts[$name]) } finally { $writer.Dispose() }
        }
        for ($i = 1; $i -le $n; $i++) {
            $e = $zip.CreateEntry("ppt/media/image$i.png", [System.IO.Compression.CompressionLevel]::NoCompression)
            $st = $e.Open()
            try { $b = [System.IO.File]::ReadAllBytes($Shots[$i - 1].Path); $st.Write($b, 0, $b.Length) } finally { $st.Dispose() }
        }
    }
    finally { $zip.Dispose() }
}

function Export-PdfAndPptx {
    # Writes <workbook name>.pdf and .pptx next to the HTML, from one picture per finished page.
    # Both are written only when every page was captured, so neither can silently miss a page.
    param([string]$HtmlPath, $Pages, [string]$OutDir, [string]$BaseName, [string]$Title)
    $shotDir = Join-Path $script:WorkDir 'page-shots'
    $err = $null
    $shots = Export-PageShots -HtmlPath $HtmlPath -Pages $Pages -ShotDir $shotDir -ErrorText ([ref]$err)
    if ($shots.Count -ne $Pages.Count) {
        Add-Issue -Severity Warning -Location 'PDF / PowerPoint' -Message ("Not written: pictures were taken of {0} of {1} page(s) [{2}]. The HTML dashboard is complete." -f $shots.Count, $Pages.Count, $err)
        Write-Log ("PDF and PowerPoint files not written: pictures were taken of {0} of {1} page(s) [{2}]." -f $shots.Count, $Pages.Count, $err) 'WARN'
        return
    }
    if ($err) { Write-Log ("Page pictures needed a second try: {0}" -f $err) }
    $targets = @()
    if (-not $SkipPdf) { $targets += 'pdf' }
    if (-not $SkipPptx) { $targets += 'pptx' }
    foreach ($kind in $targets) {
        $final = Join-Path $OutDir ($BaseName + '.' + $kind)
        $tmp = Join-Path $script:WorkDir ('out.' + $kind)
        try {
            if ($kind -eq 'pdf') {
                foreach ($s in $shots) {
                    $jpg = [System.IO.Path]::ChangeExtension($s.Path, '.jpg')
                    ConvertTo-JpegFile -PngPath $s.Path -JpegPath $jpg
                    $s | Add-Member -NotePropertyName Jpeg -NotePropertyValue $jpg -Force
                }
                Write-ImagePdf -Shots $shots -Path $tmp -Title $Title
            }
            else { Write-ImagePptx -Shots $shots -Path $tmp -Title $Title }
            Copy-Item -LiteralPath $tmp -Destination $final -Force
            $mb = [math]::Round((Get-Item -LiteralPath $final).Length / 1MB, 1)
            $what = if ($kind -eq 'pdf') { 'PDF: {0} page(s), each sized to its dashboard page' -f $shots.Count } else { 'PowerPoint: {0} slide(s) (16:9), each page fitted whole' -f $shots.Count }
            Write-Log ("{0} ({1} MB) written: {2}" -f $what, $mb, $final) 'OK'
        }
        catch {
            $why = $_.Exception.Message
            if ($why -match 'being used by another process') { $why = 'the file is open (close it in PowerPoint / the PDF viewer and run again)' }
            Add-Issue -Severity Warning -Location ('{0} file' -f $kind.ToUpperInvariant()) -Message ("Not written: {0}" -f $why)
            Write-Log ("{0} file not written: {1}" -f $kind.ToUpperInvariant(), $why) 'WARN'
        }
    }
}

function Invoke-Main {
    Write-Log "Export-WklyReportDashboard build $ScriptBuild" 'OK'

    if (-not (Test-Path -LiteralPath $WorkbookPath -PathType Leaf)) { throw "Workbook not found: $WorkbookPath" }
    $source = (Resolve-Path -LiteralPath $WorkbookPath).ProviderPath
    $ext = [System.IO.Path]::GetExtension($source).ToLowerInvariant()
    if ($ext -notin '.xlsx', '.xlsm', '.xlsb', '.xls') { throw "Unsupported file type '$ext'. Expected an Excel workbook." }
    $sourceInfo = Get-Item -LiteralPath $source

    # ---- output folder: only our own previous output is replaced
    if (-not (Test-Path -LiteralPath $OutputDirectory)) { New-Item -ItemType Directory -Path $OutputDirectory -Force | Out-Null }
    $outDir = (Resolve-Path -LiteralPath $OutputDirectory).ProviderPath
    if ($outDir.TrimEnd('\') -ieq $sourceInfo.DirectoryName.TrimEnd('\')) { throw 'Choose an OutputDirectory different from the workbook folder.' }
    # an assets folder left in the output folder by builds up to r25 is removed (only if this script made it)
    $oldAssets = Join-Path $outDir 'assets'
    if ((Test-Path -LiteralPath $oldAssets) -and (Test-Path -LiteralPath (Join-Path $oldAssets $GeneratedMarker))) {
        Remove-Item -LiteralPath $oldAssets -Recurse -Force -ErrorAction SilentlyContinue
    }
    if ($SkipSheetPdf -or $ExportDashboardPdf -or $SkipExcelView) { Write-Log 'Note: -SkipSheetPdf / -ExportDashboardPdf / -SkipExcelView no longer do anything (the PDF is controlled by -SkipPdf).' }
    # working files (pictures, charts) go to a private temp folder that is deleted at the end
    $script:WorkDir = Join-Path ([System.IO.Path]::GetTempPath()) ('ITSMDash_work_' + [guid]::NewGuid().ToString('N'))
    $assets = Join-Path $script:WorkDir 'assets'
    New-Item -ItemType Directory -Path $assets -Force | Out-Null

    # ---- hash + private temp copy (the original is never opened)
    $hashBefore = Get-SourceHash $source
    if (-not $hashBefore) { Write-Log 'The original file''s fingerprint could not be read (file locked); the unchanged-check will be reported as not verified.' 'WARN' }
    Write-Log "Source: $source (SHA-256 $hashBefore)"
    $tempDir = Join-Path $env:TEMP ('ITSMDash_' + [guid]::NewGuid().ToString('N'))
    New-Item -ItemType Directory -Path $tempDir -Force | Out-Null
    $tempCopy = Join-Path $tempDir ('snapshot' + $ext)
    try { Copy-SourceShared -From $source -To $tempCopy }
    catch { throw "Could not copy the workbook (is it locked by another program?): $($_.Exception.Message)" }

    $mediaMap = @{}
    if ($ext -in '.xlsx', '.xlsm') {
        try { $mediaMap = Get-PictureMediaMap -PackagePath $tempCopy -Sheet $SheetName -AssetsDir $assets }
        catch { Add-Issue -Severity Warning -Location 'Pictures' -Message ("Pictures could not be read from the workbook package: {0}" -f $_.Exception.Message) }
    }

    $excel = $null; $wb = $null; $blank = $null; $excelPid = 0
    $items = New-Object System.Collections.Generic.List[object]
    $result = $null; $geo = $null; $contentWidth = 0; $cellCount = 0; $shapeCount = 0
    $excelViewNote = 'not produced'
    $pictureNote = 'not captured'
    try {
        try { $excel = New-Object -ComObject Excel.Application }
        catch { throw 'Microsoft Excel is not available on this computer (COM class Excel.Application could not be created).' }
        try {
            if (-not ('ItsmDash.Win32' -as [type])) {
                Add-Type -Namespace ItsmDash -Name Win32 -MemberDefinition '[DllImport("user32.dll")] public static extern uint GetWindowThreadProcessId(IntPtr hWnd, out uint lpdwProcessId);'
            }
            [uint32]$tmpPid = 0
            [void][ItsmDash.Win32]::GetWindowThreadProcessId([IntPtr][int64]$excel.Hwnd, [ref]$tmpPid)
            $excelPid = [int]$tmpPid
        }
        catch { }
        $excel.Visible = $false
        $excel.DisplayAlerts = $false
        $excel.AskToUpdateLinks = $false
        $excel.EnableEvents = $false
        try { $excel.AutomationSecurity = 3 } catch { }                 # msoAutomationSecurityForceDisable
        $blank = $excel.Workbooks.Add()
        # Calculate like a normal Excel open, so time-based formulas show the same values the user sees.
        # (In Manual mode Excel 365 shows out-of-date values struck through - "Format Stale Values".)
        try { $excel.Calculation = -4105 } catch { }                    # xlCalculationAutomatic
        try { $excel.FormatStaleValues = $false } catch { }             # Excel 365 only; ignored elsewhere

        $m = [Type]::Missing
        Write-Log 'Opening the temporary copy read-only (links not updated)...'
        $wb = $excel.Workbooks.Open($tempCopy, 0, $true, $m, $m, $m, $true, $m, $m, $false, $false, $m, $false)
        try { $null = $excel.Calculate(); Write-Log 'Formulas calculated as in a normal Excel open.' }
        catch { Write-Log ("Formulas could not be recalculated ({0}); values are as last saved." -f $_.Exception.Message) 'WARN' }

        try {
            $script:ThemeMinor = [string]$wb.Theme.ThemeFontScheme.MinorFont.Item(1).Name
            $script:ThemeMajor = [string]$wb.Theme.ThemeFontScheme.MajorFont.Item(1).Name
            Write-Log ("Theme fonts: body '{0}', headings '{1}'" -f $script:ThemeMinor, $script:ThemeMajor)
        }
        catch { }
        $ws = $null
        $script:XlApp = $excel
        try { $ws = $wb.Worksheets.Item($SheetName); $script:XlSheet = $ws }
        catch {
            $names = @(); foreach ($s in $wb.Worksheets) { $names += $s.Name }
            throw "Worksheet '$SheetName' was not found. Worksheets in the file: $($names -join ', ')"
        }

        $used = $ws.UsedRange
        $lastRow = [int]$used.Row + [int]$used.Rows.Count - 1
        $lastCol = [int]$used.Column + [int]$used.Columns.Count - 1
        for ($k = 1; $k -le [int]$ws.Shapes.Count; $k++) {
            try {
                $br = $ws.Shapes.Item($k).BottomRightCell
                if ([int]$br.Row -gt $lastRow) { $lastRow = [int]$br.Row }
                if ([int]$br.Column -gt $lastCol) { $lastCol = [int]$br.Column }
            }
            catch { }
        }
        Write-Log ("Wkly_Report content area: A1:{0}{1}" -f (ConvertTo-ColumnLetter $lastCol), $lastRow)
        $geo = Get-SheetGeometry -Ws $ws -LastRow $lastRow -LastCol $lastCol

        Write-Log 'Reading cell background fills...'
        try { foreach ($it in (Read-FillItems -Ws $ws -Geo $geo -R1 ([int]$used.Row) -C1 ([int]$used.Column) -R2 ([int]$used.Row + [int]$used.Rows.Count - 1) -C2 ([int]$used.Column + [int]$used.Columns.Count - 1))) { $items.Add($it) } }
        catch { Add-Issue -Severity Warning -Location 'Cell fills' -Message ("Background colours could not be read: {0}" -f (Format-ErrorDetail $_)) }
        Write-Log 'Reading cells (displayed text and formatting)...'
        $cellItems = @()
        try { $cellItems = Read-CellItems -Ws $ws -Geo $geo }
        catch { Add-Issue -Severity Error -Location 'Cells' -Message ("Cell values could not be read: {0}" -f (Format-ErrorDetail $_)) }
        foreach ($it in $cellItems) { $items.Add($it) }
        $cellCount = @($cellItems | Where-Object { $_.Kind -eq 'cell' }).Count
        foreach ($b in @($cellItems | Where-Object { $_.Kind -eq 'blank' })) {
            Add-Issue -Severity Info -Location "Wkly_Report!$($b.Address)" -Message 'Formula returns an empty value - shown as blank (dashed marker), not as zero.' -Y $b.T
        }
        Write-Log 'Reading shapes, text boxes, charts and pictures...'
        $shapeItems = New-Object System.Collections.Generic.List[object]
        try { Read-ShapeItems -Ws $ws -AssetsDir $assets -MediaMap $mediaMap -Items $shapeItems }
        catch { Add-Issue -Severity Error -Location 'Wkly_Report objects' -Message ("Reading objects stopped early; what was read is shown: {0}" -f (Format-ErrorDetail $_)) }
        foreach ($it in $shapeItems) { $items.Add($it) }
        $shapeCount = @($shapeItems | Where-Object { $_.Kind -eq 'shape' }).Count

        try {
            $fcs = $used.FormatConditions
            for ($i = 1; $i -le [int]$fcs.Count; $i++) {
                $fc = $fcs.Item($i)
                $t = [int]$fc.Type
                if ($t -eq 4 -or $t -eq 6) {
                    $what = if ($t -eq 4) { 'Data bars' } else { 'Icon sets' }
                    Add-Issue -Severity Warning -Location ("Wkly_Report!{0}" -f $fc.AppliesTo.Address($false, $false)) -Message "$what (conditional formatting) are not drawn in the dashboard - see Excel view." -Y ([double]$fc.AppliesTo.Top)
                }
            }
        }
        catch { }
        try { if ([int]$used.SparklineGroups.Count -gt 0) { Add-Issue -Severity Warning -Location 'Wkly_Report' -Message 'Sparklines are not drawn in the dashboard - see Excel view.' } } catch { }

        $contentWidth = $geo.ColLefts[$lastCol + 1]
        $contentBottom = $geo.RowTops[$lastRow + 1]
        foreach ($it in $items) {
            if ($it.L + $it.W -gt $contentWidth) { $contentWidth = $it.L + $it.W }
            if ($it.T + $it.H -gt $contentBottom) { $contentBottom = $it.T + $it.H }
        }

        $breakRows = @()
        if ($PageMode -in 'Auto', 'PageBreaks') {
            try {
                $hpb = $ws.HPageBreaks
                for ($i = 1; $i -le [int]$hpb.Count; $i++) {
                    $pb = $hpb.Item($i)
                    if ([int]$pb.Type -eq -4135) { $breakRows += [int]$pb.Location.Row }
                }
            }
            catch { Write-Log 'Page breaks could not be read (no printer configured?).' 'WARN' }
        }

        try { Find-OverlappingText -Items $items } catch { }
        $result = Find-ReportPages -Items $items -Geo $geo -ContentWidth $contentWidth -ContentBottom $contentBottom `
            -Mode $PageMode -StartRows $PageStartRows -BreakRows $breakRows -MinFont $TitleMinFontSize -MinHeight $MinPageHeightPt
        Optimize-ReportPages -Pages $result.Pages

        $lastColLetter = ConvertTo-ColumnLetter ([math]::Max(1, $lastCol))
        $pictureNote = 'not captured'
        try {
            $pc = 0
            if ($RenderMode -eq 'Picture') { $pc = Add-PagePictures -Pages $result.Pages -Ws $ws -Geo $geo -AssetsDir $assets }
            elseif ($ChartRender -eq 'Excel') { Write-Log 'Render mode Rebuilt: every object is drawn at its exact Excel position; charts are Excel''s own pictures of them (-ChartRender Excel).' }
            else { Write-Log 'Render mode Rebuilt: every object is drawn at its exact Excel position; charts are redrawn from their data (-ChartRender Native).' }
            $pictureNote = '{0} of {1} pages shown as Excel''s own picture' -f $pc, $result.Pages.Count
        }
        catch { Add-Issue -Severity Warning -Location 'Page pictures' -Message ("Page pictures failed; the re-drawn view is shown: {0}" -f (Format-ErrorDetail $_)) }
    }
    finally {
        Restore-Clipboard
        Write-Log 'Closing the temporary copy without saving and closing the private Excel instance...'
        if ($wb) { try { $null = $wb.Close($false) } catch { } }
        if ($blank) { try { $null = $blank.Close($false) } catch { } }
        if ($excel) { try { $null = $excel.Quit() } catch { } }
        foreach ($o in @($ws, $used, $wb, $blank, $excel)) {
            if ($null -ne $o -and [System.Runtime.InteropServices.Marshal]::IsComObject($o)) { try { [void][System.Runtime.InteropServices.Marshal]::ReleaseComObject($o) } catch { } }
        }
        $ws = $null; $used = $null; $wb = $null; $blank = $null; $excel = $null; $script:XlApp = $null; $script:XlSheet = $null
        [GC]::Collect(); [GC]::WaitForPendingFinalizers(); [GC]::Collect(); [GC]::WaitForPendingFinalizers()
        if ($excelPid -gt 0) {
            $proc = Get-Process -Id $excelPid -ErrorAction SilentlyContinue
            if ($proc -and -not $proc.WaitForExit(15000)) {
                Write-Log "The Excel instance started by this script (PID $excelPid) did not exit; stopping that process only." 'WARN'
                Stop-Process -Id $excelPid -Force -ErrorAction SilentlyContinue
            }
        }
        Remove-Item -LiteralPath $tempDir -Recurse -Force -ErrorAction SilentlyContinue
    }

    $hashAfter = Get-SourceHash $source
    if (-not $hashBefore -or -not $hashAfter) { Write-Log 'Original workbook fingerprint could not be compared (file locked); this script only reads it.' 'WARN' }
    elseif ($hashAfter -eq $hashBefore) { Write-Log 'Original workbook unchanged (SHA-256 identical before and after).' 'OK' }
    else { Add-Issue -Severity Error -Location 'Source workbook' -Message 'The original file changed during the run (saved by someone else?). This script never writes to it. Re-run to get a consistent snapshot.' }

    $meta = [pscustomobject]@{
        FileName        = $sourceInfo.Name
        BaseName        = (($sourceInfo.BaseName -replace '[\\/:*?"<>|]', '_').Trim())
        SourcePath      = $source
        Saved           = $sourceInfo.LastWriteTime.ToString('dd/MM/yyyy HH:mm')
        Snapshot        = (Get-Date).ToString('dd/MM/yyyy HH:mm')
        HashBefore      = $hashBefore
        HashAfter       = $hashAfter
        PageMethod      = $result.Method
        CellCount       = $cellCount
        ShapeCount      = $shapeCount
        ChartCount      = $script:ChartIndex
        PictureCount    = $mediaMap.Count
        ExcelViewNote   = $excelViewNote
        PictureNote     = $pictureNote
        ReportingPeriod = (Find-LabelValue -Items $items -Label 'Reporting Period')
        ReportDate      = (Find-LabelValue -Items $items -Label 'Report Date')
    }
    Save-PageInventory -Pages $result.Pages -Path ($script:LogPath -replace 'extraction-log', 'page-inventory' -replace '\.txt$', '.json')
    $index = New-DashboardSite -Pages $result.Pages -OutDir $outDir -Meta $meta -Issues $script:Issues -WorkDir $script:WorkDir
    # validation of the single output file
    $validLevel = 'OK'
    $written = Get-Content -LiteralPath $index -Raw -Encoding UTF8
    $sections = [regex]::Matches($written, '<section class="page').Count
    $pictures = [regex]::Matches($written, 'class="lay-pic"><img').Count
    $externalRefs = [regex]::Matches($written, '(?:src|href)="(?!data:|#)[^"]+"').Count
    if ($sections -ne $result.Pages.Count -or $externalRefs -ne 0) { $validLevel = 'WARN' }
    $nativeCharts = [regex]::Matches($written, 'class="rchart"').Count
    Write-Log ("Validation: {0} page(s) in the file (expected {1}), {2} page(s) as Excel picture, {4} chart(s) redrawn from data, {3} external reference(s) (expected 0)." -f $sections, $result.Pages.Count, $pictures, $externalRefs, $nativeCharts) $validLevel
    Write-Log "Dashboard written: $index" 'OK'
    if (-not ($SkipPdf -and $SkipPptx)) {
        try { Export-PdfAndPptx -HtmlPath $index -Pages $result.Pages -OutDir $outDir -BaseName $meta.BaseName -Title ('{0} - {1}' -f $DashboardTitle, $meta.FileName) }
        catch { Write-Log ("PDF / PowerPoint files not written: {0}" -f (Format-ErrorDetail $_)) 'WARN' }
    }


    $errs = @($script:Issues | Where-Object { $_.Severity -eq 'Error' }).Count
    $warns = @($script:Issues | Where-Object { $_.Severity -eq 'Warning' }).Count
    Write-Log ("Done: {0} page(s), {1} chart(s), {2} Excel error value(s)/error(s), {3} warning(s). Log: {4}" -f $result.Pages.Count, $script:ChartIndex, $errs, $warns, $script:LogPath) 'OK'
    if ($OpenWhenDone) { Start-Process -FilePath $index }
    return $outDir
}

# Run only when executed (dot-sourcing just loads the functions, e.g. for testing).
if ($MyInvocation.InvocationName -ne '.') {
    $outForLog = $null
    try { $outForLog = Invoke-Main }
    catch {
        Write-Log (Format-ErrorDetail $_) 'ERROR'
        Write-Log 'No dashboard was produced. The original workbook was not modified.' 'ERROR'
        $exitCode = 1
    }
    finally {
        if ($script:WorkDir -and (Test-Path -LiteralPath $script:WorkDir)) { Remove-Item -LiteralPath $script:WorkDir -Recurse -Force -ErrorAction SilentlyContinue }
        # the log is kept outside the output folder so the HTML stays the only output
        try {
            $logDir = Split-Path -Parent $script:LogPath
            if (-not (Test-Path -LiteralPath $logDir)) { New-Item -ItemType Directory -Path $logDir -Force | Out-Null }
            [System.IO.File]::WriteAllLines($script:LogPath, $script:LogLines, (New-Object System.Text.UTF8Encoding($false)))
            Write-Host ("Log file: {0}" -f $script:LogPath) -ForegroundColor Gray
        }
        catch { }
    }
    if ($exitCode) { exit $exitCode }
}
