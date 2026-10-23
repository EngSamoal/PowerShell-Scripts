<#
.SYNOPSIS
    Presents the "Wkly_Report" worksheet of the weekly ITSM workbook as an executive HTML dashboard
    in the "Style C - Briefing" style (charcoal / green / gold / stone) - strictly READ-ONLY.
    Also writes a PDF and a PowerPoint copy of the dashboard.

.DESCRIPTION
    Workflow:  Excel workbook (Wkly_Report)  ->  this script  ->  <workbook name>.html / .pdf / .pptx

    This is a separate script from Export-WklyReportDashboard.ps1 (Style A) and
    Export-WklyReportDashboardB.ps1 (Design B). It reads the workbook in exactly the same safe way, but
    presents every page in the approved Style C "Briefing" layout.

    SAFETY - the script never changes anything:
      * The original workbook is never opened. It is copied to a private temp folder and only the copy
        is opened, read-only, in a separate hidden Excel instance (links not updated, macros disabled).
        The copy is closed WITHOUT saving and deleted.
      * Only the "Wkly_Report" worksheet is read. No other tab is read or changed.
      * Excel windows you already have open are not touched. Only the Excel process this script
        started is closed at the end.
      * The SHA-256 hash of the original file is compared before and after the run (see the log).
      * No internet: the PDF/PowerPoint pictures are taken by Edge (or Chrome) from the local file,
        with no network access.

    WHAT IS SHOWN (the approved sample, filled from the workbook):
      * One page at a time with a page dock, progress bar, arrow keys, zoom (- / + / Fit), All pages,
        Light / Dark, Full screen and Print.
      * Every page: a solid colour panel (charcoal / green / gold) with the page title, the page's own
        subtitle from Excel as its key message, and the page's KPI cards as big numbers; the charts,
        tables and text on the right in the sample's chart forms (stacked area, waffle, bullet bars,
        lollipop ranking, proportion bars, progress rings, ranked list, risk list, timeline).
      * Every number, chart, table and sentence is taken from Wkly_Report. Values are shown exactly as
        Excel displays them; nothing is recalculated. No message, target or comparison is invented: a
        target line is drawn only where the workbook states one. Anything that cannot be found is
        shown as "Not found in Wkly_Report".
      * A workbook chart that the sample has no block for is still shown, in the best sample form.
      * A page that is added to the workbook later is shown in the same style.

.PARAMETER WorkbookPath
    The weekly workbook. Default: C:\Temp\Bi_Wkly_Report-Daily Data.xlsx

.PARAMETER OutputDirectory
    Where <workbook name>.html, .pdf and .pptx are written. Default: C:\Temp\ITSM_Dashboard_StyleC
    (a different folder from the Style A and Design B scripts, so they never overwrite each other).

.PARAMETER PageMode
    How report pages are found: Auto (default; page titles, then manual page breaks), Titles,
    PageBreaks, Manual (use -PageStartRows).

.PARAMETER PageStartRows
    Only with -PageMode Manual: first row of each page, e.g. 1,40,72,105

.PARAMETER TitleMinFontSize
    Smallest font size (pt) that counts as a page title. Default 17 (the Executive Summary title is 17pt).

.PARAMETER MinPageHeightPt
    Two title candidates closer than this (points) belong to the same page. Default 250.

.PARAMETER DashboardTitle
    Text in the dashboard's top bar. Default "IT Service Management Briefing".

.PARAMETER Classification
    Marking shown on the top bar. Default "CONFIDENTIAL".

.PARAMETER OpenWhenDone
    Open the dashboard in the default browser when finished.

.PARAMETER SkipPdf
    Do not write <workbook name>.pdf. By default a PDF is written next to the HTML: one PDF page per
    dashboard page as designed (1920 x 1080; a longer page is kept whole), each exactly that page's size.

.PARAMETER SkipPptx
    Do not write <workbook name>.pptx. By default a PowerPoint file (16:9) is written next to the HTML:
    one slide per dashboard page (a 1920 x 1080 page fills its slide exactly). PowerPoint itself is not needed.

    The PDF and PowerPoint pages are pictures of the finished dashboard pages, taken by Microsoft Edge
    (or Chrome) in a hidden window from the local HTML file, with no network access.

.EXAMPLE
    .\Export-WklyReportDashboardC.ps1 -OpenWhenDone

.NOTES
    Requires Windows, Microsoft Excel (2010 or later) and Windows PowerShell 5.1 or PowerShell 7.
    No internet access, modules or third-party software are needed.
#>
[CmdletBinding()]
param(
    [string]$WorkbookPath = 'C:\Temp\Bi_Wkly_Report-Daily Data.xlsx',
    [string]$OutputDirectory = 'C:\Temp\ITSM_Dashboard_StyleC',
    [ValidateSet('Auto', 'Titles', 'PageBreaks', 'Manual')]
    [string]$PageMode = 'Auto',
    [int[]]$PageStartRows,
    [double]$TitleMinFontSize = 17,
    [double]$MinPageHeightPt = 250,
    [string]$DashboardTitle = 'IT Service Management Briefing',
    [string]$Classification = 'CONFIDENTIAL',
    [switch]$SkipPdf,
    [switch]$SkipPptx,
    [switch]$OpenWhenDone
)

$ScriptBuild = '2026.10.23-C1'
$SheetName = 'Wkly_Report'          # fixed by design - no other worksheet is ever read
$Inv = [System.Globalization.CultureInfo]::InvariantCulture
$ErrRegex = '#(DIV/0!|N/A|VALUE!|REF!|NAME\?|NUM!|NULL!|SPILL!|CALC!|FIELD!|BLOCKED!|CONNECT!|BUSY!|UNKNOWN!|GETTING_DATA)'

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
$script:LogPath = Join-Path (Join-Path ([System.IO.Path]::GetTempPath()) 'ITSM_Dashboard_logs') ('styleC-log ' + (Get-Date -Format 'yyyy-MM-dd HHmmss') + '.txt')
$script:Clip = $null
$script:TestClipboard = $null
$script:PictureScale = 3
$script:PdfScale = 2
$script:PdfReady = $null
$script:PictureMethods = @{}
$script:BitmapScales = New-Object System.Collections.Generic.List[double]
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
            $exported = (Export-ChartImage -Shape $Shape -Path $full -ErrorText ([ref]$exportError)) -eq $true
            $spec = $null
            try { $spec = Get-ChartSpec $Shape } catch { Write-Log ("Chart '{0}': data could not be read ({1}); Excel's chart image is used." -f $name, $_.Exception.Message) 'WARN' }
            if ($spec -and -not $spec.kind) { Write-Log ("Chart '{0}': chart type {1} is not rebuilt natively; Excel's chart image is used." -f $name, $spec.excelType) }
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
            if ($it.L -lt $minL) { $minL = [double]$it.L }
            if ($it.L + $it.W -gt $maxR) { $maxR = [double]$it.L + [double]$it.W }
        }
        if ($maxR -gt $minL) { $p.OffsetX = [math]::Max(0, $minL - 4); $p.Width = $maxR + 4 - $p.OffsetX }
        $minT = [double]$p.StartY; $maxB = [double]$p.EndY
        foreach ($it in $items) {
            if ($it.Kind -eq 'shape' -and -not $it.Runs -and -not (Test-OpaqueBg $it.Bg)) { continue }
            if ($it.T -lt $minT) { $minT = [double]$it.T }; if ($it.T + $it.H -gt $maxB) { $maxB = [double]$it.T + [double]$it.H }
        }
        $p.StartY = $minT; $p.EndY = $maxB
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
            if ($It.IsChart -and $It.Spec -and $It.Spec.kind) {
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


# ======================================================================================
#  Workbook content finders shared with Design B (pure logic - no Excel calls, no HTML)
# ======================================================================================

function ConvertTo-BArray {
    # A List (chart values, categories, colours) as a plain array - never @($list), which fails in
    # Windows PowerShell 5.1 with "Argument types do not match".
    param($X)
    if ($null -eq $X) { return , (New-Object object[] 0) }
    if ($X -is [string]) { return , ([object[]]@($X)) }
    if ($X -is [System.Collections.IList]) {
        $a = New-Object object[] $X.Count
        for ($k = 0; $k -lt $X.Count; $k++) { $a[$k] = $X[$k] }
        return , $a
    }
    return , ([object[]]@($X))
}

function Get-BTexts {
    # Text objects of a page with one-line text, largest font size and centre point.
    param($Page)
    $list = New-Object System.Collections.Generic.List[object]
    foreach ($it in $Page.Items) {
        $t = ''
        if ($it.Kind -eq 'cell') { $t = [string]$it.Text }
        elseif ($it.Kind -eq 'shape' -and $it.Runs) { $t = Get-RunsText $it }
        else { continue }
        $one = (($t -replace "[`r`n`v]+", ' ') -replace '\s{2,}', ' ').Trim()
        if (-not $one) { continue }
        $size = Get-ItemMaxFont $it
        $bold = $false
        if ($it.Kind -eq 'cell') { $bold = [bool]$it.Bold } else { foreach ($r in $it.Runs) { if ($r.Bold -and ([string]$r.Text).Trim()) { $bold = $true } } }
        $list.Add([pscustomobject]@{ Item = $it; Text = $one; Raw = $t; Size = [double]$size; Bold = $bold
                L = [double]$it.L; T = [double]$it.T; W = [double]$it.W; H = [double]$it.H; CX = [double]$it.L + [double]$it.W / 2; CY = [double]$it.T + [double]$it.H / 2 })
    }
    return , $list
}

function Get-BCards {
    # Filled shapes that look like cards (not page panels, not thin bars), smallest first.
    param($Page, [double]$MinArea = 1500, [double]$MaxArea = 120000)
    $list = New-Object System.Collections.Generic.List[object]
    foreach ($it in $Page.Items) {
        if ($it.Kind -ne 'shape' -or -not (Test-OpaqueBg $it.Bg)) { continue }
        $a = [double]$it.W * [double]$it.H
        if ($a -lt $MinArea -or $a -gt $MaxArea -or $it.W -lt 20 -or $it.H -lt 14) { continue }
        $list.Add($it)
    }
    return , $list
}

function Test-BInside {
    param($T, $Box, [double]$Pad = 2)
    return ($T.CX -ge $Box.L - $Pad -and $T.CX -le $Box.L + $Box.W + $Pad -and $T.CY -ge $Box.T - $Pad -and $T.CY -le $Box.T + $Box.H + $Pad)
}

function Test-BValue {
    param([string]$Text)
    $t = $Text.Trim()
    if ($t.Length -eq 0 -or $t.Length -gt 24) { return $false }
    if ($t -match '^[A-Z]\d{1,2}$') { return $false }          # badges such as E1 / R2
    if ($t -match $ErrRegex) { return $true }
    if ($t -notmatch '\d') { return $false }
    return ([regex]::Matches($t, '\p{L}').Count -le 4)
}

function Get-BTile {
    # A KPI card: its largest figure, the label and the details, in reading order.
    param($Texts, $Card)
    $in = @($Texts | Where-Object { Test-BInside $_ $Card })
    if ($in.Count -lt 2) { return $null }
    $vals = @($in | Where-Object { Test-BValue $_.Text } | Sort-Object Size -Descending)
    if ($vals.Count -eq 0) { return $null }
    $v = $vals[0]
    $others = @($in | Where-Object { -not [object]::ReferenceEquals($_, $v) } | Sort-Object T, L)
    if ($others.Count -eq 0) { return $null }
    $maxOther = ($others | Measure-Object -Property Size -Maximum).Maximum
    if ($v.Size -lt $maxOther) { return $null }
    # label: the largest non-figure text; details: everything else in reading order
    $labelT = @($others | Where-Object { -not (Test-BValue $_.Text) -and [regex]::Matches($_.Text, '\p{L}').Count -ge 2 } | Sort-Object @{ Expression = { $_.Size }; Descending = $true }, T)
    $label = ''; $labelObj = $null
    if ($labelT.Count) { $labelObj = $labelT[0]; $label = $labelObj.Text }
    $metaParts = New-Object System.Collections.Generic.List[string]
    $rows = @($others | Where-Object { -not [object]::ReferenceEquals($_, $labelObj) } | Sort-Object @{ Expression = { [math]::Round($_.CY / 6) } }, L)
    foreach ($o in $rows) { $metaParts.Add($o.Text) }
    return [pscustomobject]@{ Label = $label; Value = $v.Text; Meta = (($metaParts -join ' ') -replace '\s*/\s*', ' / ').Trim(); T = [double]$Card.T; L = [double]$Card.L; Card = $Card; Texts = $in }
}

function Get-BQuarterTile {
    # The "Q1 / Q2 / Q3 Volume" card: volumes in order and the % changes between them.
    param($Texts, $Card)
    $in = @($Texts | Where-Object { Test-BInside $_ $Card })
    $qs = @($in | Where-Object { $_.Text -match '^Q\d\s+Volume$' } | Sort-Object L)
    if ($qs.Count -lt 2) { return $null }
    $vols = New-Object System.Collections.Generic.List[string]
    foreach ($q in $qs) {
        $above = @($in | Where-Object { $_.CY -lt $q.T + 2 -and [math]::Abs($_.CX - $q.CX) -lt 30 -and $_.Text -match '\d' -and $_.Text -notmatch '%' } | Sort-Object @{ Expression = { [math]::Abs($_.CX - $q.CX) } })
        if ($above.Count) { $vols.Add($above[0].Text) } else { $vols.Add('—') }
    }
    $pcts = @($in | Where-Object { $_.Text -match '^-?\d+(\.\d+)?%$' } | Sort-Object L)
    $note = @($in | Where-Object { $_.Text -match '[A-Za-z]{4}' -and $_.Text -notmatch '^Q\d\s+Volume$' } | Sort-Object T)
    return [pscustomobject]@{ Kind = 'quarter'; Label = 'Quarterly volume'; Volumes = $vols; Pcts = @($pcts | ForEach-Object { $_.Text }); Note = (@($note | ForEach-Object { $_.Text }) -join ' '); T = [double]$Card.T; L = [double]$Card.L; Card = $Card }
}

function Get-BTiles {
    # KPI cards of a page (or of a region of it) in Excel's order.
    param($Page, $Texts, [double]$Top = -1e9, [double]$Bottom = 1e9)
    $tiles = New-Object System.Collections.Generic.List[object]
    $used = New-Object System.Collections.Generic.List[object]
    foreach ($c in ((Get-BCards $Page 3000 60000) | Sort-Object @{ Expression = { [math]::Round([double]$_.T / 25) } }, L)) {
        if ($c.T -lt $Top -or $c.T -gt $Bottom) { continue }
        $inside = $false
        foreach ($u in $used) { if ($c.L -ge $u.L - 1 -and $c.L + $c.W -le $u.L + $u.W + 1 -and $c.T -ge $u.T - 1 -and $c.T + $c.H -le $u.T + $u.H + 1) { $inside = $true } }
        if ($inside) { continue }
        $q = Get-BQuarterTile $Texts $c
        if ($q) { $tiles.Add($q); $used.Add($c); continue }
        $t = Get-BTile $Texts $c
        if ($t) { $tiles.Add($t); $used.Add($c) }
    }
    return , $tiles
}

function Get-BChartItem {
    param($Page, [string]$Pattern)
    foreach ($it in ($Page.Items | Sort-Object T, L)) {
        if ($it.Kind -ne 'image' -or -not $it.IsChart -or -not $it.Spec) { continue }
        if ([string]$it.Spec.title -match $Pattern) { return $it }
    }
    return $null
}

function Get-BPage {
    param($Pages, [string]$Pattern)
    foreach ($p in $Pages) { if ([string]$p.Title -match $Pattern) { return $p } }
    return $null
}

function Get-BSubtitle {
    # The line under the page title in Excel's header band.
    param($Page, $Texts)
    $title = $null
    foreach ($t in ($Texts | Sort-Object T)) { if ($t.Text -eq ([string]$Page.Title).Trim() -or $t.Text.StartsWith(([string]$Page.Title).Trim())) { $title = $t; break } }
    if ($null -eq $title) { return '' }
    $sub = @($Texts | Where-Object { $_.T -ge $title.T + $title.H * 0.6 -and $_.T -le $title.T + $title.H + 30 -and $_.L -lt $title.L + 300 -and $_.Size -lt $title.Size } | Sort-Object T, L)
    if ($sub.Count) { return $sub[0].Text }
    return ''
}

function ConvertTo-BNumberText {
    # Excel's value with its own number format applied (labels in Excel use the same format).
    param($V, [string]$Format)
    if ($null -eq $V) { return '—' }
    $f = [string]$Format
    if ($f -match '%') {
        $d = 0; if ($f -match '0\.(0+)') { $d = $Matches[1].Length }
        return ([double]$V * 100).ToString('N' + $d, $Inv) + '%'
    }
    $d = 0; if ($f -match '0\.(0+)') { $d = $Matches[1].Length }
    return ([double]$V).ToString('N' + $d, $Inv)
}

function Get-BSeries {
    param($Spec, [string]$NamePattern, [string]$Kind)
    foreach ($s in $Spec.series) {
        if ($NamePattern -and [string]$s.name -notmatch $NamePattern) { continue }
        if ($Kind -and [string]$s.kind -ne $Kind) { continue }
        return $s
    }
    return $null
}

function Get-BMonthLabel {
    param([string]$Cat)
    $names = @('Jan', 'Feb', 'Mar', 'Apr', 'May', 'Jun', 'Jul', 'Aug', 'Sep', 'Oct', 'Nov', 'Dec')
    if ($Cat -match '^\s*(?:Q\d\s+)?(\d{1,2})\s*$') { $m = [int]$Matches[1]; if ($m -ge 1 -and $m -le 12) { return $names[$m - 1] } }
    return $Cat.Trim()
}

function Get-BCardTexts {
    # Text inside each white card of a row (achievements, insight cards): reading order per card.
    param($Page, $Texts, [double]$Top, [double]$Bottom, [double]$MinW = 120)
    $out = New-Object System.Collections.Generic.List[object]
    foreach ($c in ((Get-BCards $Page 2000 80000) | Where-Object { $_.T -ge $Top -and $_.T -le $Bottom -and $_.W -ge $MinW } | Sort-Object L)) {
        $in = @($Texts | Where-Object { Test-BInside $_ $c } | Sort-Object T, L)
        if ($in.Count -eq 0) { continue }
        $out.Add([pscustomobject]@{ Card = $c; Texts = $in })
    }
    return , $out
}

function Get-BShapeTable {
    # A table drawn with shapes: a dark header bar with column titles, then one bar per row.
    param($Page, $Texts, $HeaderPattern)
    $hdr = @($Texts | Where-Object { $_.Text -match $HeaderPattern } | Sort-Object T)
    if ($hdr.Count -eq 0) { return $null }
    $h0 = $hdr[0]
    $heads = @($Texts | Where-Object { [math]::Abs($_.CY - $h0.CY) -le 4 } | Sort-Object L)
    $bars = @($Page.Items | Where-Object { $_.Kind -eq 'shape' -and (Test-OpaqueBg $_.Bg) -and $_.T -gt $h0.T + 4 -and $_.W -ge 0.8 * (($heads[-1].L + $heads[-1].W) - $heads[0].L) } | Sort-Object T)
    $rows = New-Object System.Collections.Generic.List[object]
    foreach ($b in $bars) {
        $in = @($Texts | Where-Object { Test-BInside $_ $b 1 })
        if ($in.Count -eq 0) { continue }
        $colTexts = @{}
        foreach ($t in ($in | Sort-Object T, L)) {
            $ci = 0
            for ($k = 0; $k -lt $heads.Count; $k++) { if ($t.L -ge $heads[$k].L - 12) { $ci = $k } }
            if (-not $colTexts.ContainsKey($ci)) { $colTexts[$ci] = New-Object System.Collections.Generic.List[object] }
            $colTexts[$ci].Add($t)
        }
        # a real row has an ID badge or text in at least three columns (page footers do not)
        $hasId = @($in | Where-Object { $_.Text -match '^[A-Z]\d{1,2}$' }).Count -gt 0
        if (-not $hasId -and $colTexts.Count -lt 3) { continue }
        $rows.Add([pscustomobject]@{ Bar = $b; Cols = $colTexts })
    }
    return [pscustomobject]@{ Heads = $heads; Rows = $rows }
}

function Get-BPicture {
    param($Page)
    foreach ($it in $Page.Items) { if ($it.Kind -eq 'image' -and -not $it.IsChart -and $it.Src) { return $it } }
    return $null
}

function Get-StatusClass {
    # Status words get a coloured pill; the word itself is never changed.
    param([string]$Text)
    $t = $Text.Trim().ToLowerInvariant()
    if ($t -in 'met', 'resolved', 'closed', 'complete', 'completed', 'done', 'achieved', 'on track', 'low', 'green', 'pass', 'passed', 'compliant') { return 'ok' }
    if ($t -in 'open', 'breached', 'not met', 'high', 'critical', 'red', 'fail', 'failed', 'overdue', 'missed', 'non-compliant') { return 'bd' }
    if ($t -in 'med', 'medium', 'moderate', 'monitoring', 'mitigating', 'in progress', 'amber', 'at risk', 'pending', 'on hold') { return 'wn' }
    if ($t -in 'planned', 'planning', 'not started', 'n/a', 'new') { return 'nt' }
    return $null
}

function Get-BCellValue {
    # A value from the helper table beside the pages (row label + column header), e.g. SAMI_AEC / Total.
    param($Pages, [string]$Row, [string]$Col)
    $cells = New-Object System.Collections.Generic.List[object]
    foreach ($p in $Pages) {
        foreach ($c in $p.Items) { if ($c.Kind -eq 'cell') { $cells.Add($c) } }
        if ($p.PSObject.Properties['HelperCells']) { foreach ($c in $p.HelperCells) { $cells.Add($c) } }
    }
    $r = $null; $h = $null
    foreach ($c in $cells) { if (([string]$c.Text).Trim() -eq $Row) { $r = $c; break } }
    if ($null -eq $r) { return $null }
    foreach ($c in ($cells | Sort-Object T -Descending)) { if (([string]$c.Text).Trim() -eq $Col -and $c.T -lt $r.T) { $h = $c; break } }
    if ($null -eq $h) { return $null }
    foreach ($c in $cells) { if ([math]::Abs($c.T - $r.T) -lt 1 -and [math]::Abs($c.L - $h.L) -lt 1) { return ([string]$c.Text).Trim() } }
    return $null
}

function Find-BTile {
    param($Tiles, [string]$LabelPattern)
    foreach ($t in $Tiles) { if ([string]$t.Label -match $LabelPattern) { return $t } }
    return $null
}

function Add-BUsedChart {
    param($Used, $It)
    if ($It) { $Used.Add($It) }
}


# ======================================================================================
#  Style C "Briefing": the approved sample layout, filled from Wkly_Report (pure logic - no Excel calls)
#  Every number, chart, table and sentence comes from the workbook. The key message of a page is the
#  page's own subtitle from Excel; nothing is written for it, and no target is drawn unless the
#  workbook states one. Anything that cannot be found is shown as "Not found in Wkly_Report".
# ======================================================================================

function ConvertTo-CEnc { param([string]$T) return (ConvertTo-HtmlText $T) }

function New-CChart {
    param($Data)
    return ('<div class="chart" data-c="{0}"></div>' -f [System.Net.WebUtility]::HtmlEncode(($Data | ConvertTo-Json -Depth 6 -Compress)))
}

function New-CBlock {
    param([string]$Title, [string]$Sub, [string]$Body)
    $h = '<div class="blk">'
    if ($Title) { $h += '<h3>' + (ConvertTo-CEnc (($Title -replace "[`r`n]+", ' ').Trim())) + '</h3>' }
    if ($Sub) { $h += '<div class="sub">' + (ConvertTo-CEnc $Sub) + '</div>' }
    if (-not $Body) { $Body = '<div class="miss">Not found in Wkly_Report</div>' }
    return $h + $Body + '</div>'
}

function Get-CTone {
    # Palette slot for a category name (status / priority words keep their meaning).
    param([string]$Name, [int]$Index)
    $n = $Name.ToLowerInvariant()
    if ($n -match 'critical|breach|fail|not met|overdue') { return 'bad' }
    if ($n -match '\blow\b|closed|\bmet\b|within|complete|fulfil') { return 's1' }
    if ($n -match 'high|resolved') { return 's2' }
    if ($n -match 'moderate|medium|progress') { return 's3' }
    if ($n -match 'hold|pending') { return 's4' }
    if ($n -match '\bnew\b|cancel') { return 's5' }
    return (@('s1', 's2', 's3', 's4', 's5'))[$Index % 5]
}

function Get-CHexTone {
    # Palette slot for an Excel colour: red stays red (problems), yellow/orange -> gold, green -> green, grey -> charcoal.
    param([string]$Hex, [int]$Index)
    if ($Hex -match '^#([0-9a-fA-F]{2})([0-9a-fA-F]{2})([0-9a-fA-F]{2})$') {
        $r = [Convert]::ToInt32($Matches[1], 16) / 255.0; $g = [Convert]::ToInt32($Matches[2], 16) / 255.0; $b = [Convert]::ToInt32($Matches[3], 16) / 255.0
        $mx = [math]::Max($r, [math]::Max($g, $b)); $mn = [math]::Min($r, [math]::Min($g, $b)); $d = $mx - $mn
        if ($d -lt 0.12) { return 's3' }
        $h = 0.0
        if ($mx -eq $r) { $h = 60 * ((($g - $b) / $d) % 6) } elseif ($mx -eq $g) { $h = 60 * (($b - $r) / $d + 2) } else { $h = 60 * (($r - $g) / $d + 4) }
        if ($h -lt 0) { $h += 360 }
        if ($h -lt 14 -or $h -ge 340) { return 'bad' }
        if ($h -lt 70) { return 's2' }
        if ($h -lt 170) { return 's1' }
    }
    return (@('s1', 's2', 's3', 's4'))[$Index % 4]
}

function Test-CPctSeries {
    param($S)
    return (([string]$S.name -match '%|percent|share') -or ($S.labels -and [string]$S.labels.format -match '%'))
}

function Get-CPctFormat {
    param($S)
    if ($S.labels -and [string]$S.labels.format -match '%') { return [string]$S.labels.format }
    return '0.00%'
}

function Test-CMonths {
    param($Cats)
    $hit = 0; $n = 0
    foreach ($c in $Cats) { $n++; $t = ([string]$c).Trim(); if ((Get-BMonthLabel $t) -ne $t -or $t -match '^(Jan|Feb|Mar|Apr|May|Jun|Jul|Aug|Sep|Oct|Nov|Dec)') { $hit++ } }
    return ($n -gt 0 -and $hit -ge [math]::Ceiling($n * 0.6))
}

function Get-CBars {
    # Count series of a chart (columns / bars / lines that are not percentages).
    param($Spec)
    $l = New-Object System.Collections.Generic.List[object]
    foreach ($s in $Spec.series) { if ($s.kind -in 'column', 'bar', 'line', 'area' -and -not (Test-CPctSeries $s)) { $l.Add($s) } }
    return , $l
}

function Get-CKeep {
    # Category positions to show: trailing empty categories (blank name, no values) are left out.
    param($Series, [int]$N, $Cats)
    $keep = New-Object System.Collections.Generic.List[int]
    $last = -1
    for ($i = 0; $i -lt $N; $i++) { foreach ($s in $Series) { if ($null -ne (ConvertTo-BArray $s.values)[$i]) { $last = $i } } }
    if ($last -ge 0) { $N = $last + 1 }      # months still to come (no value in any series) are not shown
    for ($i = 0; $i -lt $N; $i++) {
        $c = ''; if ($i -lt $Cats.Count) { $c = [string]$Cats[$i] }
        $empty = $true
        foreach ($s in $Series) { $v = (ConvertTo-BArray $s.values)[$i]; if ($null -ne $v -and [double]$v -ne 0) { $empty = $false } }
        if (-not $c.Trim() -and $empty) { continue }
        $keep.Add($i)
    }
    return , $keep
}

# ---------------------------------------------------------------- workbook chart -> sample chart

function ConvertTo-CArea {
    # Monthly volumes as the sample's stacked area (several series) or area (one series).
    param($Spec, [int]$H = 250)
    $bars = Get-CBars $Spec
    if ($bars.Count -eq 0) { return $null }
    $cats0 = (ConvertTo-BArray $bars[0].cats)
    $n = 0; foreach ($s in $bars) { $c = (ConvertTo-BArray $s.values).Count; if ($c -gt $n) { $n = $c } }
    $keep = Get-CKeep $bars $n $cats0
    if ($keep.Count -lt 2) { return $null }
    $cats = New-Object System.Collections.Generic.List[string]
    foreach ($i in $keep) { $c = ''; if ($i -lt $cats0.Count) { $c = [string]$cats0[$i] }; $cats.Add((Get-BMonthLabel $c)) }
    $series = New-Object System.Collections.Generic.List[object]
    foreach ($s in $bars) {
        $vals = New-Object System.Collections.Generic.List[object]
        $sv = (ConvertTo-BArray $s.values)
        foreach ($i in $keep) { $v = $sv[$i]; if ($null -eq $v) { return $null }; $vals.Add([double]$v) }   # a gap cannot be stacked: shown as columns instead
        $series.Add([ordered]@{ n = [string]$s.name; v = $vals })
    }
    return [ordered]@{ t = 'area'; cats = $cats; series = $series; h = $H }
}

function ConvertTo-CCols {
    # Columns: one series with per-bar Excel colours and Excel's % labels, or several series grouped;
    # a line series named "Open" becomes the "n open" note above each group.
    param($Spec, [int]$H = 300)
    $bars = New-Object System.Collections.Generic.List[object]
    foreach ($s in $Spec.series) { if ($s.kind -in 'column', 'bar' -and -not (Test-CPctSeries $s)) { $bars.Add($s) } }
    if ($bars.Count -eq 0) { $bars = Get-CBars $Spec }
    if ($bars.Count -eq 0) { return $null }
    $cats0 = (ConvertTo-BArray $bars[0].cats)
    $n = 0; foreach ($s in $bars) { $c = (ConvertTo-BArray $s.values).Count; if ($c -gt $n) { $n = $c } }
    $keep = Get-CKeep $bars $n $cats0
    if ($keep.Count -eq 0) { return $null }
    $cats = New-Object System.Collections.Generic.List[string]; $full = New-Object System.Collections.Generic.List[string]
    $weeks = $true
    foreach ($i in $keep) { $c = ''; if ($i -lt $cats0.Count) { $c = ([string]$cats0[$i]).Trim() }; if ($c -notmatch '^Week\s*\d+') { $weeks = $false } }
    foreach ($i in $keep) {
        $c = ''; if ($i -lt $cats0.Count) { $c = ([string]$cats0[$i]).Trim() }
        $lab = $c
        if ($weeks -and $c -match '^Week\s*(\d+)') { $lab = 'W' + $Matches[1] } elseif (Test-CMonths $cats0) { $lab = Get-BMonthLabel $c }
        $cats.Add($lab); $full.Add($c)
    }
    if ($bars.Count -eq 1) {
        $s = $bars[0]; $sv = (ConvertTo-BArray $s.values); $pts = (ConvertTo-BArray $s.points)
        $vals = New-Object System.Collections.Generic.List[object]; $cols = New-Object System.Collections.Generic.List[string]; $cc = New-Object System.Collections.Generic.List[string]
        $distinct = @($pts | Where-Object { $_ } | Sort-Object -Unique)
        $k = 0
        foreach ($i in $keep) {
            if ($null -eq $sv[$i]) { $k++; continue }
            $vals.Add([double]$sv[$i]); $cc.Add($cats[$k])
            if ($distinct.Count -ge 2) { $cols.Add((Get-CHexTone ([string]$pts[$i]) 0)) } else { $cols.Add('s1') }
            $k++
        }
        $o = [ordered]@{ t = 'cols'; cats = $cc; v = $vals; colors = $cols; h = $H }
        $pct = $null; foreach ($p in $Spec.series) { if (Test-CPctSeries $p) { $pct = $p; break } }
        if ($pct) {
            $lab = New-Object System.Collections.Generic.List[string]; $pv = (ConvertTo-BArray $pct.values); $fmt = Get-CPctFormat $pct
            foreach ($i in $keep) { if ($null -eq $sv[$i]) { continue }; $lab.Add(([double]$sv[$i]).ToString('N0', $Inv) + ' · ' + (ConvertTo-BNumberText $pv[$i] $fmt)) }
            $o['lab'] = $lab
        }
        return $o
    }
    $series = New-Object System.Collections.Generic.List[object]
    foreach ($s in $bars) {
        $vals = New-Object System.Collections.Generic.List[object]; $sv = (ConvertTo-BArray $s.values)
        foreach ($i in $keep) { if ($null -eq $sv[$i]) { $vals.Add($null) } else { $vals.Add([double]$sv[$i]) } }
        $series.Add([ordered]@{ n = [string]$s.name; v = $vals })
    }
    $o = [ordered]@{ t = 'cols'; cats = $cats; series = $series; h = $H }
    if ($weeks) { $o['tipx'] = $full }
    $open = $null; foreach ($s in $Spec.series) { if ($s.kind -eq 'line' -and [string]$s.name -match '^\s*open\s*$') { $open = $s } }
    if ($open) {
        $gl = New-Object System.Collections.Generic.List[string]; $ov = (ConvertTo-BArray $open.values)
        foreach ($i in $keep) { if ($null -eq $ov[$i]) { $gl.Add('') } else { $gl.Add(('{0} open' -f ([double]$ov[$i]).ToString('N0', $Inv))) } }
        $o['glab'] = $gl
    }
    return $o
}

function ConvertTo-CLolli {
    # Ranked lollipop: counts, with the share as Excel formats it when the chart has a % series.
    param($Spec, [string]$Color = 's1')
    $main = $null; $pct = $null
    foreach ($s in $Spec.series) { if (Test-CPctSeries $s) { if (-not $pct) { $pct = $s } } elseif (-not $main) { $main = $s } }
    if (-not $main) { return $null }
    $mc = (ConvertTo-BArray $main.cats); $mv = (ConvertTo-BArray $main.values)
    $rows = New-Object System.Collections.Generic.List[object]
    for ($i = 0; $i -lt $mv.Count; $i++) {
        if ($null -eq $mv[$i]) { continue }
        $c = ''; if ($i -lt $mc.Count) { $c = ([string]$mc[$i]).Trim() }
        if (-not $c -and [double]$mv[$i] -eq 0) { continue }
        $st = ''; if ($pct) { $st = ConvertTo-BNumberText ((ConvertTo-BArray $pct.values)[$i]) (Get-CPctFormat $pct) }
        $rows.Add([pscustomobject]@{ C = $c; V = [double]$mv[$i]; S = $st })
    }
    if ($rows.Count -eq 0) { return $null }
    # Excel draws the first bar of a bar chart at the bottom; the ranking reads largest first
    $sorted = $rows | Sort-Object -Property V -Descending
    $cats = New-Object System.Collections.Generic.List[string]; $vals = New-Object System.Collections.Generic.List[double]; $sts = New-Object System.Collections.Generic.List[string]
    foreach ($r in $sorted) { $cats.Add($r.C); $vals.Add($r.V); $sts.Add($r.S) }
    $o = [ordered]@{ t = 'lolli'; color = $Color; cats = $cats; v = $vals }
    if ($pct) { $o['st'] = $sts } else { $o['share'] = $true }
    return $o
}

function Get-CPieParts {
    param($Spec)
    $s0 = $Spec.series[0]
    if (-not $s0) { return $null }
    $vals = (ConvertTo-BArray $s0.values); $cats = (ConvertTo-BArray $s0.cats)
    $parts = New-Object System.Collections.Generic.List[object]
    $sum = 0.0
    for ($i = 0; $i -lt $vals.Count; $i++) {
        if ($null -eq $vals[$i]) { continue }
        $name = ''; if ($i -lt $cats.Count) { $name = ([string]$cats[$i]).Trim() }
        $parts.Add([pscustomobject]@{ C = $name; V = [double]$vals[$i]; I = $i }); $sum += [double]$vals[$i]
    }
    $fmt = '0.00%'; if ($s0.labels -and [string]$s0.labels.format -match '%') { $fmt = [string]$s0.labels.format }
    return [pscustomobject]@{ Parts = $parts; Sum = $sum; IsShare = ($sum -gt 0.98 -and $sum -lt 1.02); Fmt = $fmt }
}

function ConvertTo-CWaffle {
    param($Spec)
    $pp = Get-CPieParts $Spec
    if ($null -eq $pp -or $pp.Parts.Count -eq 0 -or $pp.Sum -le 0) { return $null }
    $cats = New-Object System.Collections.Generic.List[string]; $v = New-Object System.Collections.Generic.List[double]; $cols = New-Object System.Collections.Generic.List[string]; $vs = New-Object System.Collections.Generic.List[string]
    $k = 0
    foreach ($p in $pp.Parts) {
        $name = $p.C -replace '^Group\d+\s+', ''
        $cats.Add($name); $v.Add($p.V); $cols.Add((Get-CTone $name $k)); $k++
        if ($pp.IsShare) { $vs.Add((ConvertTo-BNumberText $p.V $pp.Fmt)) }
    }
    $o = [ordered]@{ t = 'waffle'; cats = $cats; v = $v; colors = $cols }
    if ($pp.IsShare) { $o['vs'] = $vs }
    return $o
}

function ConvertTo-CRing {
    # A two-part SLA doughnut (met / breached) as the sample's progress ring; no target line is drawn
    # because the chart does not hold one.
    param($Spec)
    $pp = Get-CPieParts $Spec
    if ($null -eq $pp -or $pp.Parts.Count -lt 1 -or $pp.Sum -le 0) { return $null }
    $main = $null
    foreach ($p in $pp.Parts) { if ($p.C -match '(?i)\bmet\b|within|achiev|complian') { $main = $p; break } }
    if ($null -eq $main) { foreach ($p in $pp.Parts) { if ($null -eq $main -or $p.V -gt $main.V) { $main = $p } } }
    $pct = $main.V / $pp.Sum * 100
    $s = ''
    if ($pp.IsShare) { $s = ConvertTo-BNumberText $main.V $pp.Fmt } else { $s = $pct.ToString('N2', $Inv) + '%' }
    $note = New-Object System.Collections.Generic.List[string]
    $note.Add($main.C)
    foreach ($p in $pp.Parts) {
        if ([object]::ReferenceEquals($p, $main)) { continue }
        $t = ''; if ($pp.IsShare) { $t = ConvertTo-BNumberText $p.V $pp.Fmt } else { $t = $p.V.ToString('N0', $Inv) }
        $note.Add(('{0} {1}' -f $p.C, $t))
    }
    return [ordered]@{ t = 'ring'; v = [math]::Round($pct, 2); s = $s; note = ($note -join ' · '); target = $null }
}

function ConvertTo-CProp {
    # Share bar(s): a pie as one bar; a column chart with several series (e.g. incidents vs service
    # requests per priority) as one bar per series.
    param($Spec)
    $kind = [string]$Spec.kind
    if ($kind -in 'pie', 'doughnut') {
        $pp = Get-CPieParts $Spec
        if ($null -eq $pp -or $pp.Parts.Count -eq 0 -or $pp.Sum -le 0) { return $null }
        $cats = New-Object System.Collections.Generic.List[string]; $v = New-Object System.Collections.Generic.List[double]; $cols = New-Object System.Collections.Generic.List[string]
        $k = 0; foreach ($p in $pp.Parts) { $cats.Add($p.C); $v.Add($p.V); $cols.Add((Get-CTone $p.C $k)); $k++ }
        return [ordered]@{ t = 'prop'; cats = $cats; v = $v; colors = $cols }
    }
    $bars = Get-CBars $Spec
    if ($bars.Count -eq 0) { return $null }
    $cats0 = (ConvertTo-BArray $bars[0].cats)
    $n = 0; foreach ($s in $bars) { $c = (ConvertTo-BArray $s.values).Count; if ($c -gt $n) { $n = $c } }
    $keep = Get-CKeep $bars $n $cats0
    $cats = New-Object System.Collections.Generic.List[string]; $cols = New-Object System.Collections.Generic.List[string]
    $k = 0; foreach ($i in $keep) { $c = ''; if ($i -lt $cats0.Count) { $c = ([string]$cats0[$i]).Trim() }; $cats.Add($c); $cols.Add((Get-CTone $c $k)); $k++ }
    $rows = New-Object System.Collections.Generic.List[object]
    foreach ($s in $bars) {
        $sv = (ConvertTo-BArray $s.values); $vals = New-Object System.Collections.Generic.List[double]
        foreach ($i in $keep) { if ($null -eq $sv[$i]) { $vals.Add(0) } else { $vals.Add([double]$sv[$i]) } }
        $rows.Add([ordered]@{ n = [string]$s.name; v = $vals })
    }
    return [ordered]@{ t = 'prop'; rows = $rows; cats = $cats; colors = $cols }
}

function ConvertTo-CRank {
    # The sample's ranked list (rank, name, bar, value) from a bar chart, largest first.
    param($Spec)
    $d = ConvertTo-CLolli $Spec
    if ($null -eq $d) { return $null }
    $mx = 0.0; foreach ($v in $d.v) { if ($v -gt $mx) { $mx = $v } }
    $sb = New-Object System.Text.StringBuilder
    [void]$sb.Append('<ol class="rank">')
    for ($i = 0; $i -lt $d.cats.Count; $i++) {
        $w = 0; if ($mx -gt 0) { $w = [math]::Round($d.v[$i] / $mx * 100, 1) }
        [void]$sb.Append(('<li><span class="k">{0:D2}</span><span>{1}</span><span class="t"><i style="width:{2}%"></i></span><span class="v">{3}</span></li>' -f ($i + 1), (ConvertTo-CEnc $d.cats[$i]), (Format-Num $w), $d.v[$i].ToString('N0', $Inv)))
    }
    [void]$sb.Append('</ol>')
    return $sb.ToString()
}

function Get-CChartHtml {
    # A workbook chart in the sample's style ($As), or the best sample form for its type ('auto');
    # Excel's own picture of the chart if it cannot be redrawn.
    param($It, [string]$As = 'auto', [hashtable]$Opt = @{})
    if ($null -eq $It) { return '' }
    $sp = $It.Spec
    $html = $null
    try {
        if ($As -eq 'auto') {
            $kind = [string]$sp.kind
            $cats = @(); if ($sp.series.Count) { $cats = (ConvertTo-BArray $sp.series[0].cats) }
            $bars = Get-CBars $sp
            if ($kind -in 'pie', 'doughnut') {
                $names = ($cats | ForEach-Object { [string]$_ }) -join '|'
                if ($cats.Count -le 3 -and $names -match '(?i)\bmet\b|within|breach') { $As = 'ring' }
                elseif ($names -match '(?i)\bP\d\b|critical|\bhigh\b|moderate|\blow\b') { $As = 'prop' }
                else { $As = 'waffle' }
            }
            elseif ($kind -eq 'bar') { $As = 'lolli' }
            else {
                $len = 0; foreach ($c in $cats) { $len += ([string]$c).Length }
                $avg = 0; if ($cats.Count) { $avg = $len / $cats.Count }
                if ((Test-CMonths $cats) -and $bars.Count -ge 1) { $As = 'area' }
                elseif ($bars.Count -ge 2) { $As = 'cols' }
                elseif ($cats.Count -gt 8 -or $avg -gt 12) { $As = 'lolli' }
                else { $As = 'cols' }
            }
        }
        $d = $null
        switch ($As) {
            'area' { $h = 250; if ($Opt.H) { $h = $Opt.H }; $d = ConvertTo-CArea $sp $h; if ($null -eq $d) { $d = ConvertTo-CCols $sp $h } }
            'cols' { $h = 300; if ($Opt.H) { $h = $Opt.H }; $d = ConvertTo-CCols $sp $h }
            'lolli' { $c = 's1'; if ($Opt.Color) { $c = $Opt.Color }; $d = ConvertTo-CLolli $sp $c }
            'waffle' { $d = ConvertTo-CWaffle $sp; if ($d -and $Opt.Size) { $d['size'] = $Opt.Size } }
            'ring' { $d = ConvertTo-CRing $sp }
            'prop' { $d = ConvertTo-CProp $sp }
            'rank' { $html = ConvertTo-CRank $sp }
        }
        if ($d) { $html = New-CChart $d }
    }
    catch { Add-Issue -Severity Warning -Location ("Chart '{0}'" -f $sp.title) -Message ('Could not be redrawn: {0}' -f $_.Exception.Message) }
    if ($html) { return $html }
    if ($It.Src) { return ('<div class="pic"><img src="{0}" alt="{1}"></div>' -f $It.Src, (ConvertTo-Attr ([string]$sp.title))) }
    return ''
}

function Get-CChartBlock {
    param($It, [string]$As = 'auto', [string]$Sub = '', [hashtable]$Opt = @{})
    if ($null -eq $It) { return '' }
    return (New-CBlock ([string]$It.Spec.title) $Sub (Get-CChartHtml $It $As $Opt))
}

# ---------------------------------------------------------------- panel, lists and tables

function ConvertTo-CBig {
    # A KPI card of the workbook as one of the panel's big numbers (label, value, details).
    param($Tile)
    if ($null -eq $Tile) { return $null }
    if ($Tile.PSObject.Properties['Kind'] -and $Tile.Kind -eq 'quarter') {
        $vols = New-Object System.Collections.Generic.List[string]; foreach ($x in $Tile.Volumes) { $vols.Add([string]$x) }
        $st = (@($Tile.Pcts | ForEach-Object { [string]$_ }) -join ' · ')
        if ($Tile.Note) { $st = ($st + ' · ' + $Tile.Note).Trim([char[]]@([char]32, [char]0xB7)) }
        return [pscustomobject]@{ V = ($vols -join ' → '); L = [string]$Tile.Label; S = $st }
    }
    return [pscustomobject]@{ V = [string]$Tile.Value; L = [string]$Tile.Label; S = [string]$Tile.Meta }
}

function New-CPanel {
    param([int]$No, [int]$Total, [string]$Title, [string]$Msg, $Bigs)
    $sb = New-Object System.Text.StringBuilder
    [void]$sb.Append(('<div class="panel"><div class="no">PAGE <b>{0:D2}</b> / {1}</div><h2>{2}</h2><div class="rule"></div>' -f $No, $Total, (ConvertTo-CEnc $Title)))
    if ($Msg) { [void]$sb.Append('<div class="msg">' + (ConvertTo-CEnc $Msg) + '</div>') }
    $cnt = 0; foreach ($b in $Bigs) { if ($b) { $cnt++ } }
    if ($cnt -gt 0) {
        $cls = 'big'; if ($cnt -ge 6) { $cls = 'big m6' } elseif ($cnt -ge 4) { $cls = 'big m4' }
        [void]$sb.Append('<div class="' + $cls + '">')
        foreach ($b in $Bigs) {
            if (-not $b) { continue }
            $vc = 'v'
            if ($b.V.Length -gt 12) { $vc += ' long' }
            if ($b.V -match $ErrRegex) { $vc += ' xerr' }
            $st = ''; if ($b.S) { $st = '<span class="st">' + (ConvertTo-CEnc $b.S) + '</span>' }
            [void]$sb.Append(('<div class="n"><div class="{0}">{1}</div><div class="l">{2}{3}</div></div>' -f $vc, (ConvertTo-CEnc $b.V), (ConvertTo-CEnc $b.L), $st))
        }
        [void]$sb.Append('</div>')
    }
    [void]$sb.Append('</div>')
    return $sb.ToString()
}

function Get-CTagClass {
    param([string]$Text)
    $c = Get-StatusClass $Text
    switch ($c) { 'ok' { return 'tag st-good' } 'bd' { return 'tag st-bad' } 'wn' { return 'tag st-warn' } default { return 'tag' } }
}

function Split-CNote {
    # A card sentence as a short head and the rest (all workbook text; nothing is reworded).
    param([string]$Head, [string]$Body)
    if (-not $Body -and $Head.Length -gt 46) {
        if ($Head -match '^(.{12,}?)\s+[—–]\s+(.+)$') { return @($Matches[1], $Matches[2]) }
        if ($Head -match '^(.{12,}?):\s+(.+)$') { return @($Matches[1], $Matches[2]) }
        if ($Head -match '^(.{12,}?)\s+\((.+)\)\s*$') { return @($Matches[1], $Matches[2]) }
    }
    return @($Head, $Body)
}

function New-CFacts {
    param($Items, [switch]$Two)
    $sb = New-Object System.Text.StringBuilder
    if ($Two) { [void]$sb.Append('<ul class="facts two">') } else { [void]$sb.Append('<ul class="facts">') }
    foreach ($n in $Items) {
        $hb = Split-CNote ([string]$n.Head) ([string]$n.Body)
        $tag = ''; if ($n.PSObject.Properties['Tag'] -and $n.Tag) { $tag = ('<span class="{0}">{1}</span>' -f (Get-CTagClass $n.Tag), (ConvertTo-CEnc $n.Tag.ToUpperInvariant())) }
        [void]$sb.Append(('<li><div><b>{0}{1}</b><span>{2}</span></div></li>' -f (ConvertTo-CEnc $hb[0]), $tag, (ConvertTo-CEnc $hb[1])))
    }
    [void]$sb.Append('</ul>')
    return $sb.ToString()
}

function New-CTimeline {
    param($Items)
    $sb = New-Object System.Text.StringBuilder
    [void]$sb.Append('<ul class="tl">')
    foreach ($n in $Items) {
        $w = ''; if ($n.W) { $w = '<div class="w">' + (ConvertTo-CEnc $n.W) + '</div>' }
        [void]$sb.Append(('<li>{0}<b>{1}</b><span>{2}</span></li>' -f $w, (ConvertTo-CEnc $n.Head), (ConvertTo-CEnc $n.Body)))
    }
    [void]$sb.Append('</ul>')
    return $sb.ToString()
}

function New-CTable {
    param($Heads, $Rows)
    $sb = New-Object System.Text.StringBuilder
    [void]$sb.Append('<table><thead><tr>')
    foreach ($h in $Heads) { [void]$sb.Append('<th>' + (ConvertTo-CEnc $h) + '</th>') }
    [void]$sb.Append('</tr></thead><tbody>')
    foreach ($r in $Rows) {
        [void]$sb.Append('<tr>')
        foreach ($v in $r) {
            $cls = Get-StatusClass ([string]$v)
            if ($cls) { [void]$sb.Append('<td><span class="' + (Get-CTagClass $v) + '">' + (ConvertTo-CEnc ([string]$v).Trim()) + '</span></td>') }
            else { [void]$sb.Append('<td>' + (ConvertTo-CEnc $v) + '</td>') }
        }
        [void]$sb.Append('</tr>')
    }
    [void]$sb.Append('</tbody></table>')
    return $sb.ToString()
}

function Get-CNotesFromCards {
    # Text cards of a row (achievements, insights) as head + body, in reading order.
    param($Page, $Texts, [double]$Top, [double]$Bottom)
    $notes = New-Object System.Collections.Generic.List[object]
    foreach ($c in (Get-BCardTexts $Page $Texts $Top $Bottom 150)) {
        $parts = @($c.Texts | Sort-Object T, L)
        if ($parts.Count -eq 0) { continue }
        if ($parts.Count -eq 1) { $notes.Add([pscustomobject]@{ Head = $parts[0].Text; Body = '' }); continue }
        $notes.Add([pscustomobject]@{ Head = $parts[0].Text; Body = (@($parts | Select-Object -Skip 1 | ForEach-Object { $_.Text }) -join ' ') })
    }
    return , $notes
}

# ---------------------------------------------------------------- pages

function Build-CCover {
    param($Page, $Pages, $Meta)
    $tx = Get-BTexts $Page
    $title = $null
    foreach ($t in ($tx | Sort-Object @{ Expression = { $_.Size }; Descending = $true })) { $title = $t; break }
    $eyebrow = @($tx | Where-Object { $title -and $_.T -lt $title.T -and $_.Item.Kind -ne 'cell' } | Sort-Object T | ForEach-Object { $_.Text }) -join ' · '
    $kv = New-Object System.Collections.Generic.List[object]
    $tag = New-Object System.Collections.Generic.List[string]
    foreach ($t in ($tx | Where-Object { $title -and $_.T -gt $title.T + 2 -and $_.Item.Kind -ne 'cell' } | Sort-Object T, L)) {
        if ($t.Text -match '^\s*([^:]{3,30}?)\s*:\s*(.+)$') { $kv.Add(@($Matches[1], $Matches[2])); continue }
        if ($t.Text -match '^(CONFIDENTIAL|RESTRICTED|INTERNAL)') { continue }
        if ($kv.Count -eq 0) { $tag.Add($t.Text) }
    }
    # the three big numbers of the cover (same sources as Design B): total tickets, resolution SLA, average resolution time
    $exec = Get-BPage $Pages 'Executive'
    $execTiles = New-Object System.Collections.Generic.List[object]
    if ($exec) { $execTiles = Get-BTiles $exec (Get-BTexts $exec) }
    $total = Get-BCellValue $Pages 'SAMI_AEC' 'Total'
    $bigs = New-Object System.Collections.Generic.List[object]
    if ($total) { $bigs.Add([pscustomobject]@{ V = $total; L = 'Total tickets (YTD)' }) }
    $s2 = Find-BTile $execTiles 'Resolution SLA'; if ($s2) { $bigs.Add([pscustomobject]@{ V = [string]$s2.Value; L = [string]$s2.Label }) }
    $s3 = Find-BTile $execTiles 'MTTR Resolution'; if ($s3) { $bigs.Add([pscustomobject]@{ V = [string]$s3.Value; L = [string]$s3.Label }) }
    if ($bigs.Count -eq 0) { foreach ($t in $execTiles) { if ($bigs.Count -lt 3) { $b = ConvertTo-CBig $t; if ($b) { $bigs.Add($b) } } } }
    $sb = New-Object System.Text.StringBuilder
    [void]$sb.Append('<div class="panel">')
    if ($eyebrow) { [void]$sb.Append('<div class="eyebrow">' + (ConvertTo-CEnc $eyebrow.ToUpperInvariant()) + '</div>') }
    $tt = [string]$Page.Title; if ($title) { $tt = $title.Text }
    [void]$sb.Append('<h2>' + (ConvertTo-CEnc $tt) + '</h2>')
    if ($tag.Count) { [void]$sb.Append('<div class="msg">' + (ConvertTo-CEnc ($tag -join ' ')) + '</div>') }
    if ($kv.Count) {
        [void]$sb.Append('<div class="meta">')
        foreach ($p in $kv) { [void]$sb.Append('<div><span>' + (ConvertTo-CEnc $p[0]) + '</span>' + (ConvertTo-CEnc $p[1]) + '</div>') }
        [void]$sb.Append('</div>')
    }
    if ($bigs.Count) {
        [void]$sb.Append('<div class="big">')
        foreach ($b in $bigs) {
            $vc = 'v'; if ([string]$b.V -match $ErrRegex) { $vc += ' xerr' }
            [void]$sb.Append(('<div class="n"><div class="{0}">{1}</div><div class="l">{2}</div></div>' -f $vc, (ConvertTo-CEnc $b.V), (ConvertTo-CEnc $b.L)))
        }
        [void]$sb.Append('</div>')
    }
    [void]$sb.Append('</div>')
    return $sb.ToString()
}

function Get-CTileBigs {
    param($Tiles)
    $l = New-Object System.Collections.Generic.List[object]
    foreach ($t in $Tiles) { $b = ConvertTo-CBig $t; if ($b) { $l.Add($b) } }
    return , $l
}

function Build-CExec {
    param($Page, $Used)
    $tx = Get-BTexts $Page
    $bigs = Get-CTileBigs (Get-BTiles $Page $tx)
    $mon = Get-BChartItem $Page 'Monthly'; $stat = Get-BChartItem $Page 'Status'
    $wsr = Get-BChartItem $Page 'Weekly Service'; $winc = Get-BChartItem $Page 'Weekly Incident'
    foreach ($c in @($mon, $stat, $wsr, $winc)) { Add-BUsedChart $Used $c }
    $sb = New-Object System.Text.StringBuilder
    [void]$sb.Append('<div class="row r21">' + (Get-CChartBlock $mon 'area' 'Incidents and service requests, stacked' @{ H = 215 }) + (Get-CChartBlock $stat 'waffle' 'Each square = 1% of all incidents' @{ Size = 200 }) + '</div>')
    [void]$sb.Append('<div class="row r2">' + (Get-CChartBlock $winc 'cols' 'Created vs closed per week' @{ H = 215 }) + (Get-CChartBlock $wsr 'cols' 'Created vs resolved / closed per week' @{ H = 215 }) + '</div>')
    $ach = @($tx | Where-Object { $_.Text -match '^ACHIEVEMENTS' } | Select-Object -First 1)
    if ($ach.Count) {
        $notes = Get-CNotesFromCards $Page $tx $ach[0].T ($ach[0].T + 140)
        if ($notes.Count) { [void]$sb.Append((New-CBlock ((Get-Culture).TextInfo.ToTitleCase($ach[0].Text.ToLowerInvariant())) '' (New-CFacts $notes -Two))) }
    }
    return [pscustomobject]@{ Bigs = $bigs; Content = $sb.ToString() }
}

function Build-CIncident {
    param($Page, $Used)
    $tx = Get-BTexts $Page
    $bigs = Get-CTileBigs (Get-BTiles $Page $tx)
    $pri = Get-BChartItem $Page 'Priority'; $cls = Get-BChartItem $Page 'Classification'
    Add-BUsedChart $Used $pri; Add-BUsedChart $Used $cls
    $c = (Get-CChartBlock $pri 'prop' 'Share of all incidents') + (Get-CChartBlock $cls 'lolli' 'Incidents and share of total')
    return [pscustomobject]@{ Bigs = $bigs; Content = $c }
}

function Build-CServiceRequests {
    param($Page, $Used)
    $tx = Get-BTexts $Page
    $bigs = Get-CTileBigs (Get-BTiles $Page $tx)
    $mon = Get-BChartItem $Page 'Monthly'; $cat = Get-BChartItem $Page 'category'
    Add-BUsedChart $Used $mon; Add-BUsedChart $Used $cat
    $c = (Get-CChartBlock $mon 'area' 'Service requests per month' @{ H = 230 }) + (Get-CChartBlock $cat 'lolli' 'Requests and share of all requests' @{ Color = 's2' })
    return [pscustomobject]@{ Bigs = $bigs; Content = $c }
}

function Build-CSla {
    param($Page, $Used)
    $tx = Get-BTexts $Page
    $bigs = Get-CTileBigs (Get-BTiles $Page $tx -Bottom ($Page.StartY + 150))
    $res = Get-BChartItem $Page 'Resolution SLA'; $rsp = Get-BChartItem $Page 'Response SLA'
    Add-BUsedChart $Used $res; Add-BUsedChart $Used $rsp
    $sb = New-Object System.Text.StringBuilder
    [void]$sb.Append('<div class="row r2">' + (Get-CChartBlock $res 'ring' 'Share within SLA') + (Get-CChartBlock $rsp 'ring' 'Share within SLA') + '</div>')
    # MTTR: share of each SLA used, exactly as the workbook writes it (e.g. "16.96 / 30 min SLA (57%)")
    $mt = @($tx | Where-Object { $_.Text -eq 'MTTR' } | Select-Object -First 1)
    if ($mt.Count) {
        $zone = @($tx | Where-Object { $_.L -ge $mt[0].L - 20 -and $_.T -ge $mt[0].T -and $_.T -le $mt[0].T + 300 } | Sort-Object T, L)
        $rows = New-Object System.Collections.Generic.List[object]
        foreach ($part in @('RESPONSE', 'RESOLUTION')) {
            $lab = @($zone | Where-Object { $_.Text -eq $part } | Select-Object -First 1)
            if (-not $lab.Count) { continue }
            $line = @($zone | Where-Object { $_.T -gt $lab[0].T -and $_.T -lt $lab[0].T + 100 -and $_.Text -match 'min SLA' } | Select-Object -First 1)
            $big = @($zone | Where-Object { $_.T -gt $lab[0].T -and $_.T -lt $lab[0].T + 40 -and (Test-BValue $_.Text) } | Sort-Object @{ Expression = { $_.Size }; Descending = $true } | Select-Object -First 1)
            $pct = ''; $txt = ''
            if ($line.Count) { $txt = $line[0].Text; if ($txt -match '\(\s*([\d.]+%)\s*\)') { $pct = $Matches[1] } }
            if (-not $pct) { continue }
            $name = (Get-Culture).TextInfo.ToTitleCase($part.ToLowerInvariant())
            $lineTxt = ($txt -replace '\s*\(\s*[\d.]+%\s*\)', '').Trim()
            if ($big.Count -and -not $lineTxt.StartsWith($big[0].Text)) { $name += ' · ' + $big[0].Text }
            if ($lineTxt) { $name += ' · ' + $lineTxt }
            $v = [double]::Parse(($pct.TrimEnd('%')), $Inv)
            $rows.Add([ordered]@{ n = $name; v = $v; t = $null; s = $pct })
        }
        $sr = @($zone | Where-Object { $_.Text -match 'SR SLA' } | Select-Object -First 1)
        $srTxt = ''
        if ($sr.Count) {
            $after = @($tx | Where-Object { $_.T -gt $sr[0].T - 1 -and $_.T -lt $sr[0].T + 40 -and $_.L -ge $sr[0].L - 30 -and -not [object]::ReferenceEquals($_, $sr[0]) } | Sort-Object @{ Expression = { [math]::Round($_.CY / 6) } }, L | ForEach-Object { $_.Text })
            $srTxt = (($sr[0].Text -replace '^[^\p{L}]+', '') + ' ' + (($after -join ' ') -replace '\s*/\s*', ' / ')).Trim()
        }
        if ($rows.Count) {
            $body = New-CChart ([ordered]@{ t = 'bullet'; rows = $rows; min = 0; max = 100; unit = '%'; color = 's2' })
            if ($srTxt) { $body += '<div class="note">' + (ConvertTo-CEnc $srTxt) + '</div>' }
            [void]$sb.Append('<div class="hr"></div>' + (New-CBlock 'MTTR vs SLA budget' 'Average time against the SLA budget' $body))
        }
    }
    $ins = @($tx | Where-Object { $_.Text -match 'INSIGHTS' } | Select-Object -First 1)
    if ($ins.Count) {
        $notes = Get-CNotesFromCards $Page $tx $ins[0].T ($ins[0].T + 60)
        $head = ($ins[0].Text -replace '^.*?[—–-]\s*', '')
        if ($notes.Count) { [void]$sb.Append((New-CBlock $head '' (New-CFacts $notes))) }
    }
    return [pscustomobject]@{ Bigs = $bigs; Content = $sb.ToString() }
}

function Build-CPriority {
    param($Page, $Used)
    $tx = Get-BTexts $Page
    $bigs = Get-CTileBigs (Get-BTiles $Page $tx)
    $pd = Get-BChartItem $Page 'Priority'; $top = Get-BChartItem $Page 'Incident.*Categor'; $src = Get-BChartItem $Page 'Service Request.*categ'
    Add-BUsedChart $Used $pd; Add-BUsedChart $Used $top; Add-BUsedChart $Used $src
    $c = Get-CChartBlock $pd 'prop' 'Share of each record type'
    if ($src) { $c += '<div class="row r2">' + (Get-CChartBlock $top 'lolli' 'Count and share of total') + (Get-CChartBlock $src 'lolli' 'Count and share of total' @{ Color = 's2' }) + '</div>' }
    else { $c += Get-CChartBlock $top 'lolli' 'Count and share of total' }
    return [pscustomobject]@{ Bigs = $bigs; Content = $c }
}

function Build-CAgents {
    param($Page, $Used)
    $tx = Get-BTexts $Page
    $bigs = Get-CTileBigs (Get-BTiles $Page $tx)
    $ag = Get-BChartItem $Page 'Agent'
    Add-BUsedChart $Used $ag
    $pic = Get-BPicture $Page
    $picHtml = ''
    if ($pic) { $picHtml = ('<div class="pic"><img src="{0}" alt="Assignment groups"></div><div class="note">This table is a linked picture in Wkly_Report, so it is shown as Excel''s picture of it.</div>' -f $pic.Src) }
    $sb = New-Object System.Text.StringBuilder
    $left = ''; if ($ag) { $left = Get-CChartBlock $ag 'rank' 'Tickets handled' }
    $right = ''; if ($pic) { $right = New-CBlock 'Assignment groups' 'Incidents and average MTTR' $picHtml }
    [void]$sb.Append('<div class="row r2">' + $left + $right + '</div>')
    $enh = @($tx | Where-Object { $_.Text -match '^ENHANCEMENTS' } | Select-Object -First 1)
    if ($enh.Count) {
        $sub = @($tx | Where-Object { $_.T -gt $enh[0].T -and $_.T -lt $enh[0].T + 30 -and $_.Text -ne $enh[0].Text } | Select-Object -First 1 | ForEach-Object { $_.Text })
        $items = New-Object System.Collections.Generic.List[object]
        foreach ($c in ((Get-BCards $Page 8000 80000) | Where-Object { $_.T -gt $enh[0].T -and $_.T -lt $enh[0].T + 120 } | Sort-Object L)) {
            $in = @($tx | Where-Object { Test-BInside $_ $c } | Sort-Object T, L)
            $id = (@($in | Where-Object { $_.Text -match '^[A-Z]\d{1,2}$' } | Select-Object -First 1 | ForEach-Object { $_.Text }) -join '')
            $st = (@($in | Where-Object { $_.Size -le 7.5 -and $_.Text.Length -le 16 } | Select-Object -First 1 | ForEach-Object { $_.Text }) -join '')
            $rest = @($in | Where-Object { $_.Text -notmatch '^[A-Z]\d{1,2}$' -and -not ($_.Size -le 7.5 -and $_.Text.Length -le 16) })
            $ttl = @($rest | Sort-Object @{ Expression = { $_.Size }; Descending = $true }, T | Select-Object -First 1)
            $body = @($rest | Where-Object { $ttl.Count -eq 0 -or -not [object]::ReferenceEquals($_, $ttl[0]) } | ForEach-Object { $_.Text }) -join ' '
            $tt = ''; if ($ttl.Count) { $tt = $ttl[0].Text }
            if ($id) { $tt = $id + ' · ' + $tt }
            $items.Add([pscustomobject]@{ Head = $tt; Body = $body; Tag = $st })
        }
        $title = (Get-Culture).TextInfo.ToTitleCase($enh[0].Text.ToLowerInvariant())
        if ($items.Count) { [void]$sb.Append((New-CBlock $title ($sub -join '') (New-CFacts $items))) }
    }
    return [pscustomobject]@{ Bigs = $bigs; Content = $sb.ToString() }
}

function Build-CResolution {
    param($Page, $Used)
    $tx = Get-BTexts $Page
    $ia = Get-BChartItem $Page 'Incident.*Aging'; $sa = Get-BChartItem $Page 'Service Request Aging'
    Add-BUsedChart $Used $ia; Add-BUsedChart $Used $sa
    # the "Key Insights" cards are the panel's big numbers
    $bigs = New-Object System.Collections.Generic.List[object]
    $ki = @($tx | Where-Object { $_.Text -match '^Key Insights' } | Select-Object -First 1)
    if ($ki.Count) {
        foreach ($c in ((Get-BCards $Page 8000 80000) | Where-Object { $_.T -gt $ki[0].T -and $_.T -lt $ki[0].T + 60 } | Sort-Object L)) {
            $in = @($tx | Where-Object { Test-BInside $_ $c } | Sort-Object T, L)
            $v = @($in | Where-Object { Test-BValue $_.Text } | Select-Object -First 1)
            if (-not $v.Count) { continue }
            $others = @($in | Where-Object { -not [object]::ReferenceEquals($_, $v[0]) } | Sort-Object @{ Expression = { $_.Size } }, T)
            $desc = ''; $unit = ''
            if ($others.Count) { $desc = $others[0].Text; $unit = (@($others | Select-Object -Skip 1 | ForEach-Object { $_.Text }) -join ' ') }
            $bigs.Add([pscustomobject]@{ V = $v[0].Text; L = $desc; S = $unit })
        }
    }
    if ($bigs.Count -eq 0) { $bigs = Get-CTileBigs (Get-BTiles $Page $tx) }
    $c = '<div class="row r2">' + (Get-CChartBlock $ia 'cols' 'Share of incidents by time to resolve' @{ H = 330 }) + (Get-CChartBlock $sa 'cols' 'Share of requests by time to fulfil' @{ H = 330 }) + '</div>'
    return [pscustomobject]@{ Bigs = $bigs; Content = $c }
}

function Build-CRisks {
    param($Page, $Used)
    $tx = Get-BTexts $Page
    $bigs = Get-CTileBigs (Get-BTiles $Page $tx)
    $tb = Get-BShapeTable $Page $tx '^Risk$'
    if ($null -eq $tb -or $tb.Rows.Count -eq 0) { return [pscustomobject]@{ Bigs = $bigs; Content = '' } }
    $heads = @($tb.Heads | ForEach-Object { $_.Text })
    $sb = New-Object System.Text.StringBuilder
    [void]$sb.Append('<ul class="risks">')
    $matrixItems = New-Object System.Collections.Generic.List[object]
    $impCol = -1; $likCol = -1
    for ($k = 2; $k -lt $heads.Count; $k++) { if ($heads[$k] -match '(?i)impact|severity') { $impCol = $k }; if ($heads[$k] -match '(?i)likelihood|probability') { $likCol = $k } }
    foreach ($r in $tb.Rows) {
        $first = New-Object System.Collections.Generic.List[object]
        foreach ($k in 0, 1) { if ($r.Cols.ContainsKey($k)) { foreach ($t in $r.Cols[$k]) { $first.Add($t) } } }
        $id = (@($first | Where-Object { $_.Text -match '^[A-Z]\d{1,2}$' } | Select-Object -First 1 | ForEach-Object { $_.Text }) -join '')
        $txts = @($first | Where-Object { $_.Text -notmatch '^[A-Z]\d{1,2}$' } | Sort-Object T, L)
        $ttl = ''; $desc = ''
        if ($txts.Count) { $ttl = $txts[0].Text; $desc = (@($txts | Select-Object -Skip 1 | ForEach-Object { $_.Raw -replace "^\s*-\s*", '' -replace "[`r`n`v]+\s*-?\s*", '; ' }) -join ' ').Trim() }
        $tags = ''
        $extra = New-Object System.Collections.Generic.List[string]
        for ($k = 2; $k -lt $heads.Count; $k++) {
            $cell = ''
            if ($r.Cols.ContainsKey($k)) { $cell = (@($r.Cols[$k] | ForEach-Object { $_.Text }) -join ' ') }
            if (-not $cell) { continue }
            if (Get-StatusClass $cell) { $tags += ('<span class="{0}">{1}</span> ' -f (Get-CTagClass $cell), (ConvertTo-CEnc ($heads[$k] + ': ' + $cell).ToUpperInvariant())) }
            else { $extra.Add(($heads[$k] + ': ' + $cell)) }
        }
        $span = $desc
        if ($extra.Count) { $span = (($span + ' · ' + ($extra -join ' · ')).Trim([char[]]@([char]32, [char]0xB7))) }
        [void]$sb.Append(('<li><span class="id">{0}</span><div><b>{1}</b><span>{2}</span></div><span>{3}</span></li>' -f (ConvertTo-CEnc $id), (ConvertTo-CEnc $ttl), (ConvertTo-CEnc $span), $tags.Trim()))
        if ($impCol -ge 0 -and $likCol -ge 0 -and $r.Cols.ContainsKey($impCol) -and $r.Cols.ContainsKey($likCol)) {
            $matrixItems.Add([pscustomobject]@{ Id = $id; I = (@($r.Cols[$impCol] | ForEach-Object { $_.Text }) -join ' '); P = (@($r.Cols[$likCol] | ForEach-Object { $_.Text }) -join ' ') })
        }
    }
    [void]$sb.Append('</ul>')
    $list = New-CBlock 'Mitigations' (($heads | Select-Object -Skip 1) -join ' · ') $sb.ToString()
    # the impact x likelihood matrix only when the workbook has both columns (never guessed)
    if ($matrixItems.Count) {
        $lvl = { param($t) $x = $t.ToLowerInvariant(); if ($x -match 'high|critical|major') { 2 } elseif ($x -match 'med|moderate') { 1 } else { 0 } }
        $cells = @{}
        foreach ($m in $matrixItems) { $key = '{0}{1}' -f (& $lvl $m.I), (& $lvl $m.P); if (-not $cells.ContainsKey($key)) { $cells[$key] = '' }; $cells[$key] += '<span class="id">' + (ConvertTo-CEnc $m.Id) + '</span>' }
        $mx = '<div class="matrix"><div class="ylab">Impact &rarr;</div>'
        foreach ($i in 2, 1, 0) { foreach ($p in 0, 1, 2) { $cl = 'c'; if ($i + $p -ge 3) { $cl = 'c h' } elseif ($i + $p -eq 2) { $cl = 'c m' }; $mx += '<div class="' + $cl + '">' + [string]$cells[('{0}{1}' -f $i, $p)] + '</div>' } }
        $mx += '<div class="xl"><span>Low</span><span>Likelihood &rarr;</span><span>High</span></div></div>'
        return [pscustomobject]@{ Bigs = $bigs; Content = ('<div class="row r12">' + (New-CBlock 'Risk matrix' 'Impact × likelihood' $mx) + $list + '</div>') }
    }
    return [pscustomobject]@{ Bigs = $bigs; Content = $list }
}

function Build-CSummary {
    param($Page, $Used)
    $tx = Get-BTexts $Page
    $bigs = Get-CTileBigs (Get-BTiles $Page $tx)
    $rec = @($tx | Where-Object { $_.Text -match '^Recommendations' } | Select-Object -First 1)
    $split = 470.0; if ($rec.Count) { $split = $rec[0].L - 5 }
    $log = @($tx | Where-Object { $_.Text -match 'UPDATE LOG' } | Select-Object -First 1)
    $bottom = 1e9; if ($log.Count) { $bottom = $log[0].T - 2 }
    $top = $Page.StartY + 60
    if ($rec.Count) { $top = $rec[0].T - 5 }
    $left = @($tx | Where-Object { $_.L -lt $split -and $_.T -gt $top - 10 -and $_.T -lt $bottom } | Sort-Object T, L)
    $notes = New-Object System.Collections.Generic.List[object]
    $cur = $null
    foreach ($t in $left) {
        if ($t.Size -ge 10) { $cur = [pscustomobject]@{ Head = $t.Text; Body = '' }; $notes.Add($cur); continue }
        if ($cur) { $cur.Body = ($cur.Body + ' ' + $t.Text).Trim() }
    }
    $right = @($tx | Where-Object { $_.L -ge $split -and $_.T -gt $top + 5 -and $_.T -lt $bottom } | Sort-Object T, L)
    $recs = New-Object System.Collections.Generic.List[object]
    $cur = $null; $num = ''
    foreach ($t in $right) {
        if ($t.Text -match '^\d{2}$') { $num = $t.Text; continue }
        if ($t.Size -ge 10) { $cur = [pscustomobject]@{ W = $num; Head = $t.Text; Body = '' }; $recs.Add($cur); $num = ''; continue }
        if ($cur) { $cur.Body = ($cur.Body + ' ' + $t.Text).Trim() }
    }
    $rt = 'Recommendations'; if ($rec.Count) { $rt = $rec[0].Text }
    $sb = New-Object System.Text.StringBuilder
    [void]$sb.Append('<div class="row r2">' + (New-CBlock 'Summary' '' (New-CFacts $notes)) + (New-CBlock $rt '' (New-CTimeline $recs)) + '</div>')
    $tb = Get-BShapeTable $Page $tx '^Week$'
    if ($tb -and $tb.Rows.Count) {
        $title = 'Weekly update log'; if ($log.Count) { $title = (Get-Culture).TextInfo.ToTitleCase((($log[0].Text -replace '^[^\p{L}]+', '')).ToLowerInvariant()) }
        $heads = New-Object System.Collections.Generic.List[string]; foreach ($h in $tb.Heads) { $heads.Add($h.Text) }
        $rows = New-Object System.Collections.Generic.List[object]
        $skipped = 0
        foreach ($r in $tb.Rows) {
            $vals = New-Object System.Collections.Generic.List[string]
            for ($k = 0; $k -lt $tb.Heads.Count; $k++) { $v = ''; if ($r.Cols.ContainsKey($k)) { $v = (@($r.Cols[$k] | ForEach-Object { $_.Text }) -join ' ') }; $vals.Add($v) }
            $empty = $true
            for ($k = 1; $k -lt $vals.Count; $k++) { if ($vals[$k] -and $vals[$k] -ne '[]') { $empty = $false } }
            if ($empty) { $skipped++; continue }
            $rows.Add($vals)
        }
        if ($rows.Count) { [void]$sb.Append((New-CBlock $title '' (New-CTable $heads $rows))) }
        if ($skipped) { Write-Log ("Style C: {0} empty row(s) of the weekly update log (only '[]' placeholders) are not shown." -f $skipped) }
    }
    return [pscustomobject]@{ Bigs = $bigs; Content = $sb.ToString() }
}

function Build-CGeneric {
    # A page added to the workbook later: KPI cards in the panel, charts and remaining text on the right.
    param($Page, $Used)
    $tx = Get-BTexts $Page
    $tiles = Get-BTiles $Page $tx
    $bigs = Get-CTileBigs $tiles
    $inTiles = @{}
    foreach ($t in $tiles) { if ($t.PSObject.Properties['Texts']) { foreach ($x in $t.Texts) { $inTiles[(Get-ItemId $x.Item)] = $true } } }
    $notes = New-Object System.Collections.Generic.List[object]
    foreach ($x in ($tx | Sort-Object T, L)) {
        if ($inTiles.ContainsKey((Get-ItemId $x.Item)) -or $x.Size -ge 17) { continue }
        if ($x.Text -eq ([string]$Page.Title).Trim()) { continue }
        $notes.Add([pscustomobject]@{ Head = $x.Text; Body = '' })
    }
    $c = ''; if ($notes.Count) { $c = New-CBlock '' '' (New-CFacts $notes) }
    return [pscustomobject]@{ Bigs = $bigs; Content = $c }
}

function Get-CLeftoverCharts {
    # Charts of the page that no block of the sample shows - added so nothing is left out.
    param($Page, $Used)
    $blocks = New-Object System.Collections.Generic.List[string]
    foreach ($it in ($Page.Items | Where-Object { $_.Kind -eq 'image' -and $_.IsChart -and $_.Spec } | Sort-Object T, L)) {
        $skip = $false
        foreach ($u in $Used) { if ([object]::ReferenceEquals($u, $it)) { $skip = $true } }
        if ($skip) { continue }
        $blocks.Add((Get-CChartBlock $it 'auto' ''))
    }
    $sb = New-Object System.Text.StringBuilder
    for ($k = 0; $k -lt $blocks.Count; $k += 2) {
        if ($k + 1 -lt $blocks.Count) { [void]$sb.Append('<div class="row r2">' + $blocks[$k] + $blocks[$k + 1] + '</div>') } else { [void]$sb.Append($blocks[$k]) }
    }
    return $sb.ToString()
}

function New-StyleCSite {
    # Writes ONE self-contained .html to $OutDir; pictures are read from $WorkDir\assets and embedded.
    param($Pages, [string]$OutDir, $Meta, [string]$WorkDir)
    $sections = New-Object System.Text.StringBuilder
    $total = $Pages.Count
    $n = 0
    foreach ($p in $Pages) {
        $n++
        $used = New-Object System.Collections.Generic.List[object]
        $tx = Get-BTexts $p
        $kind = 'generic'
        $t = [string]$p.Title
        if ($p.Index -eq 1 -and -not ($p.Items | Where-Object { $_.Kind -eq 'image' -and $_.IsChart })) { $kind = 'cover' }
        elseif ($t -match 'Executive') { $kind = 'exec' }
        elseif ($t -match 'Incident Summary') { $kind = 'incident' }
        elseif ($t -match 'Service Request') { $kind = 'sr' }
        elseif ($t -match 'SLA') { $kind = 'sla' }
        elseif ($t -match 'Priority') { $kind = 'priority' }
        elseif ($t -match 'Agent') { $kind = 'agents' }
        elseif ($t -match 'Resolution Time') { $kind = 'resolution' }
        elseif ($t -match 'Risk') { $kind = 'risks' }
        elseif ($t -match 'Summary') { $kind = 'summary' }
        $color = (@('green', 'gold', 'char'))[($n - 2 + 3) % 3]
        if ($kind -eq 'cover') {
            $body = ''
            try { $body = Build-CCover $p $Pages $Meta }
            catch {
                Add-Issue -Severity Warning -Location ('Page {0}' -f $p.Index) -Message ('Style C cover could not be built: {0}' -f (Format-ErrorDetail $_))
                $body = '<div class="panel"><h2>' + (ConvertTo-CEnc $t) + '</h2></div>'
            }
            [void]$sections.Append(('<section class="pg cover" data-c="char" data-t="Cover" id="p{0}">{1}</section>' -f $n, $body))
            [void]$sections.Append("`n")
            Write-Log ("  page {0} '{1}': Style C cover" -f $p.Index, $t)
            continue
        }
        $r = $null
        try {
            switch ($kind) {
                'exec' { $r = Build-CExec $p $used }
                'incident' { $r = Build-CIncident $p $used }
                'sr' { $r = Build-CServiceRequests $p $used }
                'sla' { $r = Build-CSla $p $used }
                'priority' { $r = Build-CPriority $p $used }
                'agents' { $r = Build-CAgents $p $used }
                'resolution' { $r = Build-CResolution $p $used }
                'risks' { $r = Build-CRisks $p $used }
                'summary' { $r = Build-CSummary $p $used }
                default { $r = Build-CGeneric $p $used }
            }
            $r.Content += Get-CLeftoverCharts $p $used
        }
        catch {
            Add-Issue -Severity Warning -Location ('Page {0}' -f $p.Index) -Message ('Style C page could not be built: {0}' -f (Format-ErrorDetail $_))
            $r = [pscustomobject]@{ Bigs = (New-Object System.Collections.Generic.List[object]); Content = '<div class="miss">This page could not be built - see the log.</div>' }
        }
        if (-not $r.Content) { $r.Content = '<div class="miss">Not found in Wkly_Report</div>' }
        $msg = Get-BSubtitle $p $tx
        $panel = New-CPanel $n $total $t $msg $r.Bigs
        [void]$sections.Append(('<section class="pg" data-c="{0}" data-t="{1}" id="p{2}">{3}<div class="content">{4}</div></section>' -f $color, (ConvertTo-Attr $t), $n, $panel, $r.Content))
        [void]$sections.Append("`n")
        Write-Log ("  page {0} '{1}': Style C {2} page" -f $p.Index, $t, $kind)
    }
    $html = $StyleCShell
    $html = $html.Replace('{{TITLE}}', (ConvertTo-CEnc $DashboardTitle))
    $sub = ''
    if ($Meta.ReportingPeriod) { $sub = [string]$Meta.ReportingPeriod }
    if ($Meta.ReportDate) { if ($sub) { $sub += ' · ' }; $sub += 'Report date ' + [string]$Meta.ReportDate }
    $html = $html.Replace('{{SUBTITLE}}', (ConvertTo-CEnc $sub))
    $html = $html.Replace('{{CONF}}', (ConvertTo-CEnc $Classification))
    $html = $html.Replace('{{CSS}}', $StyleCCss)
    $html = $html.Replace('{{SECTIONS}}', $sections.ToString())
    $html = $html.Replace('{{JS}}', $StyleCJs)
    # every picture is embedded so the single .html works when sent by email or Teams
    $embedded = @{}
    foreach ($m in [regex]::Matches($html, 'src="(assets/[^"]+)"')) {
        $rel = $m.Groups[1].Value
        if ($embedded.ContainsKey($rel)) { continue }
        $file = Join-Path $WorkDir ($rel.Replace('/', [string][System.IO.Path]::DirectorySeparatorChar))
        if (-not (Test-Path -LiteralPath $file)) { Write-Log ("Picture {0} is missing and could not be embedded." -f $rel) 'WARN'; continue }
        $ext = [System.IO.Path]::GetExtension($file).ToLowerInvariant().TrimStart('.')
        $mime = 'image/png'
        if ($ext -eq 'jpg' -or $ext -eq 'jpeg') { $mime = 'image/jpeg' } elseif ($ext -eq 'gif') { $mime = 'image/gif' }
        $embedded[$rel] = 'data:' + $mime + ';base64,' + [System.Convert]::ToBase64String([System.IO.File]::ReadAllBytes($file))
    }
    foreach ($rel in $embedded.Keys) { $html = $html.Replace('src="' + $rel + '"', 'src="' + $embedded[$rel] + '"') }
    $index = Join-Path $OutDir ($Meta.BaseName + '.html')
    [System.IO.File]::WriteAllText($index, $html, (New-Object System.Text.UTF8Encoding($false)))
    $mb = [math]::Round((Get-Item -LiteralPath $index).Length / 1MB, 1)
    Write-Log ("Dashboard is a single self-contained file ({0} MB, {1} pictures embedded)." -f $mb, $embedded.Count)
    return $index
}

# ======================================================================================
#  Style C page shell, stylesheet and script (the approved sample, embedded in the single .html)
# ======================================================================================

$StyleCShell = @'
<!doctype html>
<html lang="en" data-theme="light">
<head>
<meta charset="utf-8">
<meta name="viewport" content="width=device-width, initial-scale=1">
<title>{{TITLE}}</title>
<style>{{CSS}}</style>
</head>
<body>
<header class="bar">
  <span class="dot"></span>
  <h1>{{TITLE}}<span>{{SUBTITLE}}</span></h1>
  <div class="zoom"><button class="btn" id="zOut" title="Zoom out (-)">&minus;</button><span id="zVal">100%</span><button class="btn" id="zIn" title="Zoom in (+)">+</button><button class="btn" id="zFit" title="Fit the page to the screen">Fit</button></div>
  <button class="btn" id="all" title="Show every page one after another">All pages</button>
  <button class="btn" id="theme">Light / Dark</button>
  <button class="btn" id="full" title="Full screen (F)">Full screen</button>
  <button class="btn" onclick="window.print()">Print</button>
  <span class="conf">{{CONF}}</span>
</header>

<div class="stage" id="stage">
{{SECTIONS}}
</div>
<div class="prog"><i id="progbar"></i></div>
<nav class="dock">
  <button class="btn" id="prev" title="Previous page">&#9664;</button>
  <div class="pages" id="pages"></div>
  <span class="pos" id="pos"></span>
  <button class="btn" id="next" title="Next page">&#9654;</button>
</nav>
<div class="tip" id="tip"></div>
<script>{{JS}}</script>
</body>
</html>
'@

$StyleCCss = @'
/* ===== Style C "Briefing" - hall theme colours (charcoal, green, gold, stone), solid colours only ===== */
:root{
  color-scheme:light;
  --bg:#f3f0e8;--ink:#1f2227;--ink-2:#4b4f57;--ink-3:#757982;--rule:#d9d4c7;--track:#e4dfd3;
  --bar:#262a31;--bar-ink:#f3f0e8;--bar-ink-2:#a9acb3;--bar-btn:#343943;--bar-btn-h:#424854;
  --p-char:#262a31;--p-green:#13805a;--p-gold:#c99a2e;--p-ink:#f3f0e8;--p-ink-2:#c9cbd0;--p-gold-ink:#1f2227;--p-gold-ink-2:#3d3a2f;
  --s1:#13805a;--s2:#a87b0e;--s3:#3e434c;--gold:#c99a2e;
  --good:#13805a;--warn:#a87b0e;--bad:#a5371f;--chip:#e9e4d8;--mid:#cfc6b0;--s4:#8a8f98;--s5:#c8bfa8;
  --tip-bg:#262a31;--tip-ink:#f3f0e8;
}
:root[data-theme="dark"]{
  color-scheme:dark;
  --bg:#1c1e22;--ink:#eeeae1;--ink-2:#c3c1ba;--ink-3:#8f9199;--rule:#33363d;--track:#2c2f35;
  --bar:#141518;--bar-ink:#eeeae1;--bar-ink-2:#8f9199;--bar-btn:#2a2d33;--bar-btn-h:#363a41;
  --p-char:#2e323a;--p-green:#17704f;--p-gold:#b08a24;--p-ink:#eeeae1;--p-ink-2:#b9bcc3;--p-gold-ink:#16171a;--p-gold-ink-2:#2e2a1e;
  --s1:#2a9a6c;--s2:#b08a24;--s3:#8a9099;--gold:#c99a2e;
  --good:#3cae7e;--warn:#d0a33a;--bad:#c94236;--chip:#2a2d33;--mid:#454a53;--s4:#5d636d;--s5:#7a7262;
  --tip-bg:#eeeae1;--tip-ink:#16171a;
}
*{box-sizing:border-box}
html,body{height:100%}
body{margin:0;background:var(--bg);color:var(--ink);font:18px/1.45 "Segoe UI",system-ui,-apple-system,Arial,sans-serif;-webkit-print-color-adjust:exact;print-color-adjust:exact;display:flex;flex-direction:column}
.sample{background:var(--gold);color:#1f2227;text-align:center;font-size:13px;font-weight:600;padding:4px 12px}

/* slim top bar */
.bar{flex:none;display:flex;align-items:center;gap:12px;padding:10px 24px;background:var(--bar);color:var(--bar-ink)}
.bar .dot{width:12px;height:12px;background:var(--gold)}
.bar h1{margin:0;font-size:19px;font-weight:600;letter-spacing:.2px;flex:1;min-width:0;white-space:nowrap;overflow:hidden;text-overflow:ellipsis}
.bar h1 span{font-weight:400;color:var(--bar-ink-2);margin-left:10px;font-size:16px}
.btn{background:var(--bar-btn);color:var(--bar-ink);border:0;padding:7px 12px;font:600 14px "Segoe UI",sans-serif;cursor:pointer;letter-spacing:.3px}
.btn:hover{background:var(--bar-btn-h)}
.zoom{display:flex;align-items:center;gap:3px}
.zoom .btn{min-width:34px;padding:7px 10px}
.zoom span{min-width:52px;text-align:center;font-size:14px;color:var(--bar-ink-2);font-variant-numeric:tabular-nums}
.conf{color:var(--gold);font-size:13px;font-weight:700;letter-spacing:2.5px;border:2px solid var(--gold);padding:4px 9px}

/* stage: one page at a time */
.stage{flex:1;min-height:0;position:relative}
.pg{position:absolute;inset:0;display:none;grid-template-columns:minmax(380px,31%) 1fr}
.pg.cur{display:grid}

/* left colour panel */
.panel{background:var(--pc);color:var(--pi);padding:38px 40px 30px;display:flex;flex-direction:column;min-height:0;overflow:auto}
.pg[data-c="char"]{--pc:var(--p-char);--pi:var(--p-ink);--pi2:var(--p-ink-2);--acc:var(--gold)}
.pg[data-c="green"]{--pc:var(--p-green);--pi:var(--p-ink);--pi2:var(--p-ink-2);--acc:var(--gold)}
.pg[data-c="gold"]{--pc:var(--p-gold);--pi:var(--p-gold-ink);--pi2:var(--p-gold-ink-2);--acc:var(--p-gold-ink)}
.panel .no{font-size:15px;letter-spacing:3px;font-weight:600;color:var(--pi2)}
.panel .no b{color:var(--acc);font-weight:700}
.panel h2{margin:10px 0 0;font-size:40px;font-weight:600;line-height:1.08;letter-spacing:-.6px}
.panel .msg{margin:18px 0 0;font-size:22px;font-weight:300;line-height:1.35;color:var(--pi)}
.panel .rule{width:56px;height:5px;background:var(--acc);margin:22px 0 4px}
.big{margin-top:auto;padding-top:20px}
.big .n{padding:14px 0;border-top:1px solid var(--pi2)}
.big .n:first-child{border-top:0}
.big .v{font-size:52px;font-weight:300;line-height:1.05;letter-spacing:-1px;font-variant-numeric:tabular-nums}
.big .v small{font-size:22px;margin-left:4px}
.big .l{font-size:14px;letter-spacing:1.6px;text-transform:uppercase;color:var(--pi2);margin-top:4px}
.big .l .st{color:var(--pi);font-weight:700;letter-spacing:.5px;text-transform:none;margin-left:6px}

/* right content */
.content{padding:30px 36px 26px;overflow:auto;min-width:0;display:flex;flex-direction:column;gap:26px}
.row{display:grid;gap:34px}.r2{grid-template-columns:1fr 1fr}.r21{grid-template-columns:1.5fr 1fr}.r12{grid-template-columns:1fr 1.5fr}
.blk{min-width:0}
.blk h3{margin:0;font-size:14px;font-weight:700;letter-spacing:1.8px;text-transform:uppercase;color:var(--ink-2)}
.blk h3::after{content:"";display:block;width:34px;height:4px;background:var(--gold);margin:8px 0 2px}
.blk .sub{font-size:15px;color:var(--ink-3);margin:6px 0 12px}
.hr{height:1px;background:var(--rule)}
.legend{display:flex;flex-wrap:wrap;gap:4px 18px;font-size:15px;color:var(--ink-2);margin-top:8px}
.legend i{display:inline-block;width:12px;height:12px;margin-right:6px;vertical-align:-1px}
.chart{position:relative;width:100%}
.chart svg{display:block;width:100%;overflow:visible}
.ax{fill:var(--ink-3);font-size:14px}.lbl{fill:var(--ink);font-size:14px;font-weight:600}.gl{stroke:var(--rule)}
.hit{fill:transparent}
.tip{position:fixed;pointer-events:none;background:var(--tip-bg);color:var(--tip-ink);font-size:15px;padding:6px 10px;display:none;z-index:50;white-space:nowrap}
.st-good{color:var(--good)}.st-warn{color:var(--warn)}.st-bad{color:var(--bad)}
.tag{display:inline-block;font-size:13px;font-weight:700;letter-spacing:.6px;padding:2px 8px;background:var(--chip)}

/* ranked list */
.rank{margin:0;padding:0;list-style:none}
.rank li{display:grid;grid-template-columns:34px 1fr 1fr 70px;align-items:center;gap:12px;padding:7px 0;border-bottom:1px solid var(--rule);font-size:16px}
.rank .k{font-size:14px;color:var(--ink-3);font-variant-numeric:tabular-nums}
.rank .t{height:10px;background:var(--track)}.rank .t i{display:block;height:100%;background:var(--s1)}
.rank .v{text-align:right;font-weight:600;font-variant-numeric:tabular-nums}
table{width:100%;border-collapse:collapse;font-size:16px}
th{text-align:left;font-size:12.5px;letter-spacing:1.4px;text-transform:uppercase;color:var(--ink-3);padding:6px 8px;border-bottom:2px solid var(--ink)}
td{padding:9px 8px;border-bottom:1px solid var(--rule)}
td.n,th.n{text-align:right;font-variant-numeric:tabular-nums}

/* risk matrix */
.matrix{display:grid;grid-template-columns:auto repeat(3,1fr);grid-template-rows:repeat(3,96px) auto;gap:4px;font-size:14px}
.matrix .ylab{writing-mode:vertical-rl;transform:rotate(180deg);text-align:center;color:var(--ink-3);letter-spacing:1.4px;text-transform:uppercase;font-size:12.5px;grid-row:1/4}
.matrix .c{background:var(--track);display:flex;flex-wrap:wrap;align-content:center;justify-content:center;gap:6px;padding:6px}
.matrix .c.m{background:var(--mid)}.matrix .c.h{background:var(--s3)}
.matrix .c.h .id{background:var(--gold);color:#1f2227}
.matrix .xl{grid-column:2/5;display:flex;justify-content:space-between;color:var(--ink-3);letter-spacing:1.4px;text-transform:uppercase;font-size:12.5px;padding-top:4px}
.id{display:inline-grid;place-items:center;min-width:36px;height:30px;background:var(--ink);color:var(--bg);font-weight:700;font-size:14px}
.risks{margin:0;padding:0;list-style:none}
.risks li{display:grid;grid-template-columns:44px 1fr auto;gap:12px;padding:10px 0;border-bottom:1px solid var(--rule);align-items:start}
.risks b{display:block;font-size:17px}.risks span{color:var(--ink-2);font-size:15px}.risks .id{color:var(--bg)}

/* timeline */
.tl{margin:0;padding:0 0 0 26px;list-style:none;border-left:4px solid var(--gold)}
.tl li{position:relative;padding:0 0 20px 16px}
.tl li::before{content:"";position:absolute;left:-38px;top:4px;width:20px;height:20px;background:var(--bg);border:4px solid var(--gold)}
.tl .w{font-size:13px;letter-spacing:1.6px;text-transform:uppercase;color:var(--ink-3)}
.tl b{display:block;font-size:19px;margin:2px 0}.tl span{color:var(--ink-2);font-size:16px}
.facts{margin:0;padding:0;list-style:none;counter-reset:f}
.facts li{display:grid;grid-template-columns:auto 1fr;gap:16px;padding:12px 0;border-bottom:1px solid var(--rule)}
.facts li::before{counter-increment:f;content:counter(f,decimal-leading-zero);font-size:34px;font-weight:300;color:var(--s2);line-height:1}
.facts b{display:block;font-size:18px}.facts span{color:var(--ink-2);font-size:16px}

/* cover */
.pg.cover{grid-template-columns:1fr}
.cover .panel{padding:70px 90px;justify-content:center}
.cover .eyebrow{color:var(--gold);letter-spacing:5px;font-weight:700;font-size:16px}
.cover h2{font-size:96px;font-weight:600;letter-spacing:-2px;margin:18px 0 10px;line-height:1}
.cover .msg{font-size:28px;max-width:1100px}
.cover .meta{display:flex;flex-wrap:wrap;gap:12px 46px;margin-top:40px;font-size:18px}
.cover .meta div span{display:block;color:var(--pi2);font-size:13px;letter-spacing:1.6px;text-transform:uppercase;margin-bottom:2px}
.cover .big{display:grid;grid-template-columns:repeat(3,minmax(0,1fr));gap:0 40px;margin-top:56px;padding-top:0}
.cover .big .n{border-top:5px solid var(--gold);padding-top:16px}
.cover .big .v{font-size:76px}

/* dock */
.dock{flex:none;display:flex;align-items:center;gap:10px;padding:8px 24px 10px;background:var(--bar);color:var(--bar-ink)}
.dock .pages{display:flex;gap:4px;flex:1;min-width:0;overflow-x:auto;scrollbar-width:none}
.dock .pages::-webkit-scrollbar{display:none}
.dock .pages button{flex:none;background:var(--bar-btn);color:var(--bar-ink-2);border:0;padding:6px 11px;font:600 14px "Segoe UI",sans-serif;cursor:pointer;white-space:nowrap}
.dock .pages button span{font-weight:400;margin-left:6px}
.dock .pages button.on{background:var(--gold);color:#1f2227}
.dock .pos{font-size:14px;color:var(--bar-ink-2);min-width:62px;text-align:center;font-variant-numeric:tabular-nums}
.prog{flex:none;height:4px;background:var(--bar-btn)}.prog i{display:block;height:100%;background:var(--gold);transition:width .2s}

/* all-pages view, small screens and print: every page one after another */
body.all .stage,body.narrow .stage{position:static}
body.all .pg,body.narrow .pg{position:relative;display:grid;min-height:calc(100vh - 120px);border-bottom:8px solid var(--bar)}
body.all .dock .pages button.on{background:var(--bar-btn);color:var(--bar-ink-2)}
body.all,body.narrow{display:block}
body.all .bar,body.narrow .bar{position:sticky;top:0;z-index:5}
body.all .dock,body.all .prog,body.narrow .dock,body.narrow .prog{display:none}
@media (max-width:980px){
  .pg,.pg.cover{grid-template-columns:1fr}
  .panel{padding:28px 22px}.content{padding:22px 18px;overflow:visible}
  .r2,.r21,.r12{grid-template-columns:1fr}
  .cover .panel{padding:40px 22px}.cover h2{font-size:40px;overflow-wrap:anywhere}.bar{flex-wrap:wrap}.cover .big{grid-template-columns:1fr}.cover .big .v{font-size:58px}
  .bar h1 span,.conf{display:none}
}
@media print{
  @page{size:A4 landscape;margin:0}
  body{display:block}.bar,.dock,.prog,.sample{display:none!important}
  .stage{position:static}.pg{position:relative;display:grid!important;height:100vh;break-after:page;border:0}
  .content{overflow:visible}
}

/* ---- additions for workbook data (same design language) ---- */
.big.m4 .v{font-size:42px}.big.m4 .n{padding:10px 0}
.big.m6{display:grid;grid-template-columns:1fr 1fr;gap:0 22px}.big.m6 .n{padding:9px 0;border-top:1px solid var(--pi2)}.big.m6 .n:nth-child(-n+2){border-top:0}.big.m6 .v{font-size:34px}
.big .v.long{font-size:30px;letter-spacing:-.3px}
.xerr{outline:2px dashed var(--bad);outline-offset:3px}
.pic img{display:block;max-width:100%;height:auto;border:1px solid var(--rule)}
.note{font-size:14px;color:var(--ink-3);margin-top:8px}
.miss{color:var(--ink-3);font-size:15px;border:2px dashed var(--rule);padding:14px}
.facts .tag{margin-left:8px;vertical-align:2px}
.facts.two{display:grid;grid-template-columns:1fr 1fr;column-gap:34px}
/* ?shot=N: page N alone, 1920 px wide, at least 1080 px high (pictures for the PDF and PowerPoint files) */
body.shot{display:block;height:auto}
body.shot .bar,body.shot .dock,body.shot .prog{display:none!important}
body.shot .stage{position:static;zoom:1!important}
body.shot .pg{display:none!important}
body.shot .pg.cur{display:grid!important;position:relative;min-height:1080px;height:auto}
body.shot .content,body.shot .panel{overflow:visible}
'@

$StyleCJs = @'
(function () {
  var secs = [].slice.call(document.querySelectorAll('.pg'));
  var body = document.body, cur = 0;
  // ---- navigation: one page at a time, dock with page buttons
  var pagesEl = document.getElementById('pages');
  secs.forEach(function (s, i) {
    var b = document.createElement('button');
    b.innerHTML = String(i + 1).padStart(2, '0') + '<span>' + s.getAttribute('data-t') + '</span>';
    b.onclick = function () { go(i); };
    pagesEl.appendChild(b);
  });
  var btns = [].slice.call(pagesEl.children);
  function go(i) {
    cur = Math.max(0, Math.min(secs.length - 1, i));
    if (body.classList.contains('all') || body.classList.contains('narrow')) { secs[cur].scrollIntoView(); return; }
    secs.forEach(function (s, k) { s.classList.toggle('cur', k === cur); });
    btns.forEach(function (b, k) { b.classList.toggle('on', k === cur); });
    btns[cur].scrollIntoView({ block: 'nearest', inline: 'nearest' });
    document.getElementById('pos').textContent = (cur + 1) + ' / ' + secs.length;
    document.getElementById('progbar').style.width = ((cur + 1) / secs.length * 100) + '%';
    if (!body.classList.contains('shot')) { try { history.replaceState(null, '', '#' + secs[cur].id); } catch (e) { } }
    drawAll();
  }
  document.getElementById('prev').onclick = function () { go(cur - 1); };
  document.getElementById('next').onclick = function () { go(cur + 1); };
  document.addEventListener('keydown', function (e) {
    if (['ArrowRight', 'PageDown'].indexOf(e.key) >= 0) { e.preventDefault(); go(cur + 1); }
    if (['ArrowLeft', 'PageUp'].indexOf(e.key) >= 0) { e.preventDefault(); go(cur - 1); }
    if (e.key === 'Home') go(0);
    if (e.key === 'End') go(secs.length - 1);
    if (e.key === 'f' || e.key === 'F') full();
  });
  var sx = null;
  document.getElementById('stage').addEventListener('touchstart', function (e) { sx = e.touches[0].clientX; }, { passive: true });
  document.getElementById('stage').addEventListener('touchend', function (e) { if (sx === null) return; var d = e.changedTouches[0].clientX - sx; if (Math.abs(d) > 60) go(cur + (d < 0 ? 1 : -1)); sx = null; });
  function full() { if (!document.fullscreenElement) { if (document.documentElement.requestFullscreen) document.documentElement.requestFullscreen().catch(function () { }); } else document.exitFullscreen(); }
  document.getElementById('full').onclick = full;
  document.getElementById('all').onclick = function () {
    body.classList.toggle('all');
    this.textContent = body.classList.contains('all') ? 'One page' : 'All pages';
    fitStage(); drawAll(); if (!body.classList.contains('all')) go(cur);
  };
  document.getElementById('theme').onclick = function () {
    var r = document.documentElement; r.dataset.theme = r.dataset.theme === 'dark' ? 'light' : 'dark'; drawAll();
  };
  function narrowCheck() { body.classList.toggle('narrow', window.innerWidth <= 980); fitStage(); }
  // pages are designed for 1920 x 1080; on smaller screens the page is scaled down (never below 72%) so it fits without scrolling
  // zoom: the automatic fit times the viewer's own zoom (- / + / Fit, remembered on this PC)
  var userZoom = 1;
  try { userZoom = parseFloat(localStorage.getItem('itsmC.zoom')) || 1; } catch (e) { }
  function fitStage() {
    var st = document.getElementById('stage'), fit = 1;
    if (body.classList.contains('shot')) { st.style.zoom = ''; return; }
    if (!body.classList.contains('narrow') && !body.classList.contains('all')) {
      st.style.zoom = '';
      var avail = window.innerHeight - st.getBoundingClientRect().top - document.querySelector('.dock').offsetHeight - 4;
      fit = Math.max(0.72, Math.min(1, window.innerWidth / 1920, avail / 925));
    }
    var z = fit * userZoom;
    st.style.zoom = Math.abs(z - 1) > 0.005 ? z : '';
    document.getElementById('zVal').textContent = Math.round(userZoom * 100) + '%';
  }
  function setZoom(z) {
    userZoom = Math.max(0.5, Math.min(2, Math.round(z * 10) / 10));
    try { localStorage.setItem('itsmC.zoom', String(userZoom)); } catch (e) { }
    fitStage(); drawAll();
  }
  document.getElementById('zIn').onclick = function () { setZoom(userZoom + 0.1); };
  document.getElementById('zOut').onclick = function () { setZoom(userZoom - 0.1); };
  document.getElementById('zFit').onclick = function () { setZoom(1); };
  document.addEventListener('keydown', function (e) {
    if (e.ctrlKey || e.metaKey || e.altKey) return;
    if (e.key === '+' || e.key === '=') setZoom(userZoom + 0.1);
    if (e.key === '-' || e.key === '_') setZoom(userZoom - 0.1);
  });

  // ---- charts (SVG, solid colours, direct labels, hover values)
  function cv(n) { return getComputedStyle(document.documentElement).getPropertyValue('--' + n).trim(); }
  function nf(v) { return Number(v).toLocaleString('en-US'); }
  function pc(v, t) { return (t ? v / t * 100 : 0).toFixed(1) + '%'; }
  function esc(t) { return String(t).replace(/&/g, '&amp;').replace(/</g, '&lt;').replace(/"/g, '&quot;'); }
  var mctx = document.createElement('canvas').getContext('2d');
  function tw(t, sz, w) { mctx.font = (w || 400) + ' ' + sz + 'px "Segoe UI",system-ui,sans-serif'; return mctx.measureText(String(t)).width; }
  function clip(t, max, sz) { t = String(t); if (tw(t, sz) <= max) return t; while (t.length > 1 && tw(t + '…', sz) > max) t = t.slice(0, -1); return t + '…'; }
  function nice(m) { var p = Math.pow(10, Math.floor(Math.log10(m || 1))), f = m / p; return (f <= 1 ? 1 : f <= 2 ? 2 : f <= 5 ? 5 : 10) * p; }
  function svg(W, H, g) { return '<svg viewBox="0 0 ' + W + ' ' + H + '" height="' + H + '">' + g + '</svg>'; }
  function legend(items) { return '<div class="legend">' + items.map(function (it) { return '<span><i style="background:' + cv(it[1]) + '"></i>' + it[0] + '</span>'; }).join('') + '</div>'; }
  var tip = document.getElementById('tip');
  document.addEventListener('mousemove', function (e) {
    var t = e.target && e.target.getAttribute && e.target.getAttribute('data-tip');
    if (!t) { tip.style.display = 'none'; return; }
    tip.innerHTML = t; tip.style.display = 'block';
    tip.style.left = Math.min(e.clientX + 14, window.innerWidth - tip.offsetWidth - 8) + 'px'; tip.style.top = (e.clientY + 16) + 'px';
  });

  function area(el, o) {   // stacked area, solid fills, 2px surface gap between layers
    var W = el.clientWidth, H = o.h || 240, L = 52, R = 18, T = 24, B = 30, n = o.cats.length, ser = o.series;
    var tot = o.cats.map(function (c, i) { return ser.reduce(function (a, s) { return a + s.v[i]; }, 0); });
    var mx = Math.max.apply(null, tot), step = nice(mx / 4), top = Math.ceil(mx * 1.12 / step) * step;
    var x = function (i) { return L + (W - L - R) * i / (n - 1); }, y = function (v) { return T + (H - T - B) * (1 - v / top); };
    var g = '';
    for (var t = 0; t <= top + 1e-9; t += step) g += '<line class="gl" x1="' + L + '" x2="' + (W - R) + '" y1="' + y(t) + '" y2="' + y(t) + '"/><text class="ax" x="' + (L - 8) + '" y="' + (y(t) + 5) + '" text-anchor="end">' + nf(t) + '</text>';
    var base = o.cats.map(function () { return 0; }), cols = ['s1', 's2', 's3'];
    ser.forEach(function (s, k) {
      var topPts = s.v.map(function (v, i) { return [x(i), y(base[i] + v)]; }), botPts = base.map(function (b, i) { return [x(i), y(b)]; }).reverse();
      g += '<polygon fill="' + cv(cols[k]) + '" points="' + topPts.concat(botPts).map(function (p) { return p[0].toFixed(1) + ',' + p[1].toFixed(1); }).join(' ') + '"/>';
      if (k < ser.length - 1) g += '<polyline fill="none" stroke="' + cv('bg') + '" stroke-width="2" points="' + topPts.map(function (p) { return p[0].toFixed(1) + ',' + p[1].toFixed(1); }).join(' ') + '"/>';
      s.v.forEach(function (v, i) { base[i] += v; });
    });
    var hi = tot.indexOf(mx);
    o.cats.forEach(function (c, i) {
      g += '<text class="ax" x="' + x(i).toFixed(1) + '" y="' + (H - 8) + '" text-anchor="' + (i === 0 ? 'start' : i === n - 1 ? 'end' : 'middle') + '">' + esc(c) + '</text>';
      g += '<rect class="hit" x="' + (x(i) - (W - L - R) / (n - 1) / 2).toFixed(1) + '" y="' + T + '" width="' + ((W - L - R) / (n - 1)).toFixed(1) + '" height="' + (H - T - B) + '" data-tip="' + esc('<b>' + c + '</b> · ' + ser.map(function (s) { return s.n + ' ' + nf(s.v[i]); }).join(' · ') + (ser.length > 1 ? ' · total ' + nf(tot[i]) : '')) + '"/>';
      if (i === hi || i === n - 1) g += '<circle cx="' + x(i).toFixed(1) + '" cy="' + y(tot[i]).toFixed(1) + '" r="5" fill="' + cv('ink') + '"/><text class="lbl" x="' + x(i).toFixed(1) + '" y="' + (y(tot[i]) - 12).toFixed(1) + '" text-anchor="' + (i === n - 1 ? 'end' : 'middle') + '">' + (i === hi ? 'Peak ' : '') + nf(tot[i]) + '</text>';
    });
    el.innerHTML = svg(W, H, g) + (ser.length > 1 ? legend(ser.map(function (s, k) { return [s.n, cols[k]]; })) : '');
  }

  function waffle(el, o) {   // 10 x 10 squares, each = 1%
    var sum = o.v.reduce(function (a, b) { return a + b; }, 0), cells = [];
    var raw = o.v.map(function (v) { return v / sum * 100; }), n = raw.map(Math.floor), left = 100 - n.reduce(function (a, b) { return a + b; }, 0);
    raw.map(function (r, i) { return [r - Math.floor(r), i]; }).sort(function (a, b) { return b[0] - a[0]; }).slice(0, left).forEach(function (p) { n[p[1]]++; });
    n.forEach(function (c, i) { for (var k = 0; k < c; k++) cells.push(i); });
    var W = Math.min(el.clientWidth, o.size || 300), gap = 4, s = (W - gap * 9) / 10, g = '';
    cells.forEach(function (ci, k) {
      var r = Math.floor(k / 10), c = k % 10;
      g += '<rect x="' + (c * (s + gap)).toFixed(1) + '" y="' + (r * (s + gap)).toFixed(1) + '" width="' + s.toFixed(1) + '" height="' + s.toFixed(1) + '" fill="' + cv(o.colors[ci]) + '" data-tip="' + esc(o.cats[ci] + ': <b>' + (o.vs ? o.vs[ci] : nf(o.v[ci])) + '</b> · ' + pc(o.v[ci], sum)) + '"/>';
    });
    el.innerHTML = '<div style="display:flex;gap:24px;align-items:center;flex-wrap:wrap"><div style="width:' + W + 'px">' + svg(W, W, g) + '</div><div class="legend" style="flex-direction:column;gap:8px;margin:0">' +
      o.cats.map(function (c, i) { return '<span><i style="background:' + cv(o.colors[i]) + '"></i>' + esc(c) + ' · <b>' + (o.vs ? esc(o.vs[i]) : nf(o.v[i])) + '</b> (' + pc(o.v[i], sum) + ')</span>'; }).join('') + '</div></div>';
  }

  function bullet(el, o) {   // actual bar against a target line
    var W = el.clientWidth, rowH = 46, lw = Math.min(Math.max.apply(null, o.rows.map(function (r) { return tw(r.n, 15); })) + 16, W * 0.42), vw = 70;
    var H = rowH * o.rows.length, x = function (v) { return lw + (W - lw - vw) * (v - o.min) / (o.max - o.min); }, g = '';
    o.rows.forEach(function (r, i) {
      var yy = i * rowH + 8, hasT = r.t !== null && r.t !== undefined, low = o.flag && hasT && r.v < r.t, c = low ? cv('bad') : cv(o.color || 's1'), shown = r.s || (r.v + (o.unit || ''));
      g += '<text x="0" y="' + (yy + 19) + '" style="fill:var(--ink-2);font-size:15px">' + esc(clip(r.n, lw - 12, 15)) + '</text>';
      g += '<rect x="' + lw + '" y="' + (yy + 4) + '" width="' + (W - lw - vw) + '" height="22" fill="' + cv('track') + '"/>';
      g += '<rect x="' + lw + '" y="' + (yy + 4) + '" width="' + Math.max(2, x(r.v) - lw).toFixed(1) + '" height="22" fill="' + c + '" data-tip="' + esc(r.n + ': <b>' + shown + '</b>' + (hasT ? ' · target ' + r.t + (o.unit || '') : '')) + '"/>';
      if (hasT && r.t < o.max) g += '<rect x="' + (x(r.t) - 1.5).toFixed(1) + '" y="' + yy + '" width="3" height="30" fill="' + cv('ink') + '"/>';
      g += '<text class="lbl" x="' + (W - vw + 10) + '" y="' + (yy + 21) + '" style="font-size:16px">' + esc(shown) + (low ? ' ▼' : '') + '</text>';
    });
    el.innerHTML = svg(W, H, g);
  }

  function lolli(el, o) {   // ranked lollipop: thin stem, round end, value label
    var W = el.clientWidth, rowH = 31, n = o.cats.length, H = rowH * n + 4, sum = o.v.reduce(function (a, b) { return a + b; }, 0), mx = Math.max.apply(null, o.v);
    var lw = Math.min(Math.max.apply(null, o.cats.map(function (c) { return tw(c, 15); })) + 22, W * 0.42), vw = tw('0,000 · 00.0%', 14, 600) + 20, g = '';
    var c = cv(o.color || 's1');
    o.cats.forEach(function (k, i) {
      var v = o.v[i], yy = i * rowH + rowH / 2 + 2, xe = lw + (W - lw - vw) * v / mx;
      g += '<text x="' + (lw - 12) + '" y="' + (yy + 5) + '" text-anchor="end" style="fill:var(--ink-2);font-size:15px">' + esc(clip(k, lw - 14, 15)) + '</text>';
      g += '<line x1="' + lw + '" x2="' + xe.toFixed(1) + '" y1="' + yy + '" y2="' + yy + '" stroke="' + c + '" stroke-width="3"/><circle cx="' + xe.toFixed(1) + '" cy="' + yy + '" r="7" fill="' + c + '"/>';
      var sh = o.st ? (o.st[i] ? ' · ' + o.st[i] : '') : (o.share ? ' · ' + pc(v, sum) : '');
      g += '<text class="lbl" x="' + (xe + 14).toFixed(1) + '" y="' + (yy + 5) + '">' + nf(v) + esc(sh) + '</text>';
      g += '<rect class="hit" x="0" y="' + (yy - rowH / 2) + '" width="' + W + '" height="' + rowH + '" data-tip="' + esc(k + ': <b>' + nf(v) + '</b>' + sh) + '"/>';
    });
    el.innerHTML = svg(W, H, g);
  }

  function prop(el, o) {   // 100% proportion bars with labels for each part
    var rows = o.rows || [{ n: '', v: o.v }], W = el.clientWidth, lw = o.rows ? Math.max.apply(null, rows.map(function (r) { return tw(r.n, 15, 600); })) + 18 : 0, bh = 44, H = rows.length * (bh + 40), g = '';
    rows.forEach(function (r, ri) {
      var sum = r.v.reduce(function (a, b) { return a + b; }, 0), x = lw, yy = ri * (bh + 40) + 4, lastLabelEnd = -1e9;
      if (o.rows) g += '<text x="0" y="' + (yy + bh / 2 + 5) + '" style="fill:var(--ink);font-size:15px;font-weight:600">' + esc(r.n) + '</text>';
      r.v.forEach(function (v, i) {
        var w = (W - lw) * v / sum, wd = Math.max(w - 2, 1.5);
        g += '<rect x="' + x.toFixed(1) + '" y="' + yy + '" width="' + wd.toFixed(1) + '" height="' + bh + '" fill="' + cv(o.colors[i]) + '" data-tip="' + esc((r.n ? r.n + ' · ' : '') + o.cats[i] + ': <b>' + nf(v) + '</b> · ' + pc(v, sum)) + '"/>';
        var txt = o.cats[i] + ' ' + pc(v, sum), tx = Math.max(x, lastLabelEnd + 12);
        if (tx + tw(txt, 14, 600) <= W) { g += '<text class="lbl" x="' + tx.toFixed(1) + '" y="' + (yy + bh + 20) + '" style="font-weight:400;fill:var(--ink-2)"><tspan style="font-weight:700;fill:var(--ink)">' + esc(pc(v, sum)) + '</tspan> ' + esc(o.cats[i]) + '</text>'; lastLabelEnd = tx + tw(txt, 14, 600); }
        x += w;
      });
    });
    el.innerHTML = svg(W, H, g) + legend(o.cats.map(function (c, i) { return [c, o.colors[i]]; }));
  }

  function ring(el, o) {   // single progress ring with target tick
    var W = el.clientWidth, D = Math.min(W, 250), R = D / 2 - 6, r = R - 26, cx = W / 2, cy = D / 2, H = D + 34;
    function P(rr, f) { var a = -Math.PI / 2 + f * 2 * Math.PI; return (cx + rr * Math.cos(a)).toFixed(2) + ',' + (cy + rr * Math.sin(a)).toFixed(2); }
    var f = o.v / 100, big = f > 0.5 ? 1 : 0, g = '';
    g += '<circle cx="' + cx + '" cy="' + cy + '" r="' + ((R + r) / 2) + '" fill="none" stroke="' + cv('track') + '" stroke-width="' + (R - r) + '"/>';
    g += '<path d="M' + P(R, 0) + ' A' + R + ',' + R + ' 0 ' + big + ' 1 ' + P(R, f) + ' L' + P(r, f) + ' A' + r + ',' + r + ' 0 ' + big + ' 0 ' + P(r, 0) + 'Z" fill="' + cv('s1') + '" data-tip="' + esc('<b>' + (o.s || (o.v + '%')) + '</b>' + (o.note ? ' · ' + o.note : '')) + '"/>';
    var tf = o.target / 100; if (o.target !== null && o.target !== undefined) g += '<line x1="' + P(r - 6, tf).split(',')[0] + '" y1="' + P(r - 6, tf).split(',')[1] + '" x2="' + P(R + 6, tf).split(',')[0] + '" y2="' + P(R + 6, tf).split(',')[1] + '" stroke="' + cv('ink') + '" stroke-width="3"/>';
    g += '<text x="' + cx + '" y="' + (cy + 12) + '" text-anchor="middle" style="font-size:' + Math.round(r * 0.5) + 'px;font-weight:300;fill:var(--ink)">' + esc(o.s || (o.v + '%')) + '</text>';
    g += '<text x="' + cx + '" y="' + (D + 26) + '" text-anchor="middle" class="ax" style="font-size:15px">' + esc(o.note || '') + (o.target !== null && o.target !== undefined ? ' · target ' + o.target + '% (tick)' : '') + '</text>';
    el.innerHTML = svg(W, H, g);
  }

  function cols(el, o) {   // columns with share labels (or Excel's own labels); several series = grouped columns
    if (o.series) return gcols(el, o);
    var W = el.clientWidth, H = o.h || 300, T = 30, B = 32, n = o.cats.length, sum = o.v.reduce(function (a, b) { return a + b; }, 0), mx = Math.max.apply(null, o.v);
    var band = W / n, bw = band * 0.62, y = function (v) { return T + (H - T - B) * (1 - v / mx); }, g = '';
    o.cats.forEach(function (c, i) {
      var v = o.v[i], x = band * i + (band - bw) / 2;
      g += '<rect x="' + x.toFixed(1) + '" y="' + y(v).toFixed(1) + '" width="' + bw.toFixed(1) + '" height="' + (H - B - y(v)).toFixed(1) + '" fill="' + cv(o.colors[i]) + '" data-tip="' + esc(c + ': <b>' + nf(v) + '</b> · ' + pc(v, sum)) + '"/>';
      g += '<text class="lbl" x="' + (x + bw / 2).toFixed(1) + '" y="' + (y(v) - 8).toFixed(1) + '" text-anchor="middle" style="font-size:16px">' + esc(o.lab ? (o.lab[i] || '') : pc(v, sum)) + '</text>';
      g += '<text class="ax" x="' + (x + bw / 2).toFixed(1) + '" y="' + (H - 8) + '" text-anchor="middle">' + esc(c) + '</text>';
    });
    g += '<line x1="0" x2="' + W + '" y1="' + (H - B) + '" y2="' + (H - B) + '" stroke="' + cv('ink-3') + '"/>';
    el.innerHTML = svg(W, H, g);
  }

  function gcols(el, o) {
    var W = el.clientWidth, H = o.h || 280, L = 50, T = 34, B = 32, n = o.cats.length, ser = o.series, cl = o.colors || ['s1', 's2', 's3', 's4'];
    var mx = 0; ser.forEach(function (s) { s.v.forEach(function (v) { if (v > mx) mx = v; }); });
    var step = nice(mx / 4), top = Math.ceil(mx * 1.15 / step) * step || 1, band = (W - L) / n, gw = band * 0.7, bw = (gw - 3 * (ser.length - 1)) / ser.length;
    var y = function (v) { return T + (H - T - B) * (1 - v / top); }, g = '';
    for (var t = 0; t <= top + 1e-9; t += step) g += '<line class="gl" x1="' + L + '" x2="' + W + '" y1="' + y(t) + '" y2="' + y(t) + '"/><text class="ax" x="' + (L - 8) + '" y="' + (y(t) + 5) + '" text-anchor="end">' + nf(t) + '</text>';
    o.cats.forEach(function (c, i) {
      var x0 = L + band * i + (band - gw) / 2, hi = 0;
      ser.forEach(function (s, k) {
        var v = s.v[i]; if (v === null || v === undefined) return;
        var x = x0 + k * (bw + 3); hi = Math.max(hi, v);
        g += '<rect x="' + x.toFixed(1) + '" y="' + y(v).toFixed(1) + '" width="' + bw.toFixed(1) + '" height="' + Math.max(0, H - B - y(v)).toFixed(1) + '" fill="' + cv(cl[k % cl.length]) + '" data-tip="' + esc((o.tipx ? o.tipx[i] : c) + ' · ' + s.n + ': <b>' + nf(v) + '</b>' + (o.glab && o.glab[i] ? ' · ' + o.glab[i] : '')) + '"/>';
        if (bw >= 20) g += '<text class="lbl" x="' + (x + bw / 2).toFixed(1) + '" y="' + (y(v) - 6).toFixed(1) + '" text-anchor="middle" style="font-size:13px">' + nf(v) + '</text>';
      });
      if (o.glab && o.glab[i]) g += '<text class="ax" x="' + (x0 + gw / 2).toFixed(1) + '" y="' + (y(hi) - 22).toFixed(1) + '" text-anchor="middle" style="font-size:13px">' + esc(o.glab[i]) + '</text>';
      g += '<text class="ax" x="' + (L + band * i + band / 2).toFixed(1) + '" y="' + (H - 8) + '" text-anchor="middle">' + esc(clip(c, band - 4, 14)) + '</text>';
    });
    g += '<line x1="' + L + '" x2="' + W + '" y1="' + (H - B) + '" y2="' + (H - B) + '" stroke="' + cv('ink-3') + '"/>';
    el.innerHTML = svg(W, H, g) + legend(ser.map(function (s, k) { return [esc(s.n), cl[k % cl.length]]; }));
  }

  var kinds = { area: area, waffle: waffle, bullet: bullet, lolli: lolli, prop: prop, ring: ring, cols: cols };
  function drawAll() {
    [].forEach.call(document.querySelectorAll('.chart[data-c]'), function (el) {
      if (!el.offsetParent || !el.clientWidth) return;   // hidden pages are drawn when shown
      try { var o = JSON.parse(el.getAttribute('data-c')); kinds[o.t](el, o); }
      catch (e) { el.innerHTML = '<div class="st-bad">Chart could not be drawn</div>'; }
    });
  }
  var rt; window.addEventListener('resize', function () { clearTimeout(rt); rt = setTimeout(function () { narrowCheck(); drawAll(); }, 120); });
  var shotM = /[?&]shot=(\d+)/.exec(location.search);
  if (shotM) body.classList.add('shot');
  narrowCheck();
  var start = 0, h = location.hash.replace('#', '');
  secs.forEach(function (s, i) { if (s.id === h) start = i; });
  if (shotM) start = Math.max(0, Math.min(secs.length - 1, parseInt(shotM[1], 10) - 1));
  go(start);
  if (document.fonts && document.fonts.ready) document.fonts.ready.then(drawAll);
  // shot mode: a check mark 20px below the page tells the script where the page ends (it is cropped away)
  if (shotM) {
    var mark = function () {
      drawAll();
      var o = document.getElementById('shotmark'); if (o) o.parentNode.removeChild(o);
      var r = secs[cur].getBoundingClientRect(), m = document.createElement('div'); m.id = 'shotmark';
      m.style.cssText = 'position:absolute;left:20px;width:10px;height:10px;background:rgb(1,2,3);top:' + (Math.ceil(r.bottom + window.scrollY) + 20) + 'px';
      body.appendChild(m);
    };
    mark(); window.addEventListener('load', mark); setTimeout(mark, 1500); if (document.fonts && document.fonts.ready) document.fonts.ready.then(mark);
  }
})();
'@

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

function Save-MarkedCrop {
    # A page's height is found from the dashboard's check mark (pages longer than 1080 px are kept whole),
    # drawn $Gap pixels below the page: everything above it is the page. Returns the page height in pixels.
    # No mark means the browser did not show the page in full; nothing is cropped then.
    param([string]$Source, [string]$Target, [int]$W, [int]$MarkX, [int]$Gap, [int]$MinH)
    Add-Type -AssemblyName System.Drawing
    $img = [System.Drawing.Bitmap]::FromFile($Source)
    try {
        if ($img.Width -lt $W) { throw ('picture is {0} px wide, narrower than the page ({1} px)' -f $img.Width, $W) }
        $markY = -1
        for ($y = 0; $y -lt $img.Height; $y++) {
            $c = $img.GetPixel($MarkX, $y)
            if ([math]::Abs($c.R - 1) + [math]::Abs($c.G - 2) + [math]::Abs($c.B - 3) -le 9) { $markY = $y; break }
        }
        if ($markY -lt 0) { throw 'check mark not found' }
        $h = $markY - $Gap
        if ($h -lt $MinH) { throw ('page came out only {0} px high' -f $h) }
        $crop = $img.Clone((New-Object System.Drawing.Rectangle 0, 0, $W, $h), [System.Drawing.Imaging.PixelFormat]::Format24bppRgb)
        try { $crop.Save($Target, [System.Drawing.Imaging.ImageFormat]::Png) } finally { $crop.Dispose() }
        return $h
    }
    finally { $img.Dispose() }
}

function Export-PageShotsB {
    # Opens the finished dashboard in a hidden browser window (?shot=N shows page N alone as designed,
    # 1920 px wide and at least 1080 px high) and saves one picture per page at twice
    # that resolution. A page taller than the window is taken again with a taller window, so nothing
    # is cut off. Page 1 decides which browser/mode works; when none does, the rest is skipped.
    param([string]$HtmlPath, $Titles, [string]$ShotDir, [ref]$ErrorText)
    $shots = New-Object System.Collections.Generic.List[object]
    $browsers = Find-HeadlessBrowser
    if ($browsers.Count -eq 0) { $ErrorText.Value = 'Microsoft Edge / Google Chrome was not found on this PC'; return , $shots }
    New-Item -ItemType Directory -Path $ShotDir -Force | Out-Null
    $url = (New-Object System.Uri($HtmlPath)).AbsoluteUri
    $scale = 2; $wCss = 1920
    $errs = New-Object System.Collections.Generic.List[string]
    $working = $null
    for ($i = 1; $i -le $Titles.Count; $i++) {
        $raw = Join-Path $ShotDir ('raw-{0:D2}.png' -f $i)
        $png = Join-Path $ShotDir ('page-{0:D2}.png' -f $i)
        $combos = New-Object System.Collections.Generic.List[object]
        if ($working) { $combos.Add($working) }
        else { foreach ($br in $browsers) { foreach ($mode in @('--headless=new', '--headless')) { $combos.Add([pscustomobject]@{ Browser = $br; Mode = $mode }) } } }
        $hPx = 0
        foreach ($cb in $combos) {
            $who = '{0} {1}' -f (Split-Path $cb.Browser -Leaf), $cb.Mode
            foreach ($winH in @(2400, 8200)) {
                if (Test-Path -LiteralPath $raw) { Remove-Item -LiteralPath $raw -Force }
                $profileDir = Join-Path $ShotDir ('browser-profile-' + [guid]::NewGuid().ToString('N').Substring(0, 8))
                $argList = @($cb.Mode, '--do-not-de-elevate', '--disable-gpu', '--no-first-run', '--no-default-browser-check', '--hide-scrollbars', '--disable-extensions',
                    '--disable-sync', '--disable-background-networking', '--disable-component-update', '--disable-default-apps', '--mute-audio',
                    '--enable-logging=stderr', '--v=0',
                    # no network at all for this hidden browser (local file, no external references)
                    '--proxy-server=http://127.0.0.1:9', '--no-pings', '--disable-domain-reliability', '--disable-client-side-phishing-detection',
                    ('"--user-data-dir={0}"' -f $profileDir), ('--force-device-scale-factor={0}' -f $scale), ('--window-size={0},{1}' -f $wCss, $winH),
                    '--virtual-time-budget=5000', ('"--screenshot={0}"' -f $raw), ('"{0}?shot={1}"' -f $url, $i))
                try {
                    $why = Invoke-HeadlessShot -Browser $cb.Browser -Arguments ($argList -join ' ') -Raw $raw
                    if ($why) { $errs.Add(('page {0} ({1}): {2}' -f $i, $who, $why)); break }
                    $hPx = [int](Save-MarkedCrop -Source $raw -Target $png -W ($wCss * $scale) -MarkX (25 * $scale) -Gap (20 * $scale) -MinH (500 * $scale))
                    if (-not $working) { $working = $cb; Write-Log ("PDF/PowerPoint pages: pictures taken with {0} ({1}, hidden, local file only, no network)." -f $cb.Browser, $cb.Mode) }
                    break
                }
                catch {
                    $errs.Add(('page {0} ({1}, window {2} px high): {3}' -f $i, $who, $winH, $_.Exception.Message))
                    if ($_.Exception.Message -ne 'check mark not found') { break }
                }
            }
            if ($hPx -gt 0) { break }
        }
        if ($hPx -gt 0) {
            $shots.Add([pscustomobject]@{ Index = $i; Title = [string]$Titles[$i - 1]; Path = $png; WPx = $wCss * $scale; HPx = $hPx; WPt = $wCss * 0.75; HPt = $hPx / $scale * 0.75 })
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

function Export-PdfAndPptx {
    # Writes <workbook name>.pdf and .pptx next to the HTML, from one picture per finished page.
    # Both are written only when every page was captured, so neither can silently miss a page.
    param([string]$HtmlPath, $Titles, [string]$OutDir, [string]$BaseName, [string]$Title)
    $shotDir = Join-Path $script:WorkDir 'page-shots'
    $err = $null
    $shots = Export-PageShotsB -HtmlPath $HtmlPath -Titles $Titles -ShotDir $shotDir -ErrorText ([ref]$err)
    if ($shots.Count -ne $Titles.Count) {
        Add-Issue -Severity Warning -Location 'PDF / PowerPoint' -Message ("Not written: pictures were taken of {0} of {1} page(s) [{2}]. The HTML dashboard is complete." -f $shots.Count, $Titles.Count, $err)
        Write-Log ("PDF and PowerPoint files not written: pictures were taken of {0} of {1} page(s) [{2}]." -f $shots.Count, $Titles.Count, $err) 'WARN'
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
    Write-Log "Export-WklyReportDashboardC (Style C) build $ScriptBuild" 'OK'

    if (-not (Test-Path -LiteralPath $WorkbookPath -PathType Leaf)) { throw "Workbook not found: $WorkbookPath" }
    $source = (Resolve-Path -LiteralPath $WorkbookPath).ProviderPath
    $ext = [System.IO.Path]::GetExtension($source).ToLowerInvariant()
    if ($ext -notin '.xlsx', '.xlsm', '.xlsb', '.xls') { throw "Unsupported file type '$ext'. Expected an Excel workbook." }
    $sourceInfo = Get-Item -LiteralPath $source

    # ---- output folder: only our own previous output is replaced
    if (-not (Test-Path -LiteralPath $OutputDirectory)) { New-Item -ItemType Directory -Path $OutputDirectory -Force | Out-Null }
    $outDir = (Resolve-Path -LiteralPath $OutputDirectory).ProviderPath
    if ($outDir.TrimEnd('\') -ieq $sourceInfo.DirectoryName.TrimEnd('\')) { throw 'Choose an OutputDirectory different from the workbook folder.' }
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
        ReportingPeriod = (Find-LabelValue -Items $items -Label 'Reporting Period')
        ReportDate      = (Find-LabelValue -Items $items -Label 'Report Date')
        PreparedBy      = (Find-LabelValue -Items $items -Label 'Prepared By')
    }
    Save-PageInventory -Pages $result.Pages -Path ($script:LogPath -replace 'extraction-log', 'page-inventory' -replace '\.txt$', '.json')
    Write-Log 'Arranging pages in the Style C layout:'
    $index = New-StyleCSite -Pages $result.Pages -OutDir $outDir -Meta $meta -WorkDir $script:WorkDir
    # validation of the single output file
    $validLevel = 'OK'
    $written = Get-Content -LiteralPath $index -Raw -Encoding UTF8
    $sections = [regex]::Matches($written, '<section class="pg').Count
    $externalRefs = [regex]::Matches($written, '(?:src|href)="(?!data:|#)[^"]+"').Count
    if ($sections -ne $result.Pages.Count -or $externalRefs -ne 0) { $validLevel = 'WARN' }
    $nativeCharts = [regex]::Matches($written, 'class="chart" data-c=').Count
    Write-Log ("Validation: {0} page(s) in the file (expected {1}), {2} chart(s) redrawn from data, {3} external reference(s) (expected 0)." -f $sections, $result.Pages.Count, $nativeCharts, $externalRefs) $validLevel
    Write-Log "Dashboard written: $index" 'OK'
    if (-not ($SkipPdf -and $SkipPptx)) {
        try {
            $titles = New-Object System.Collections.Generic.List[string]
            for ($k = 0; $k -lt $sections; $k++) { $t = ''; if ($k -lt $result.Pages.Count) { $t = [string]$result.Pages[$k].Title }; $titles.Add($t) }
            Export-PdfAndPptx -HtmlPath $index -Titles $titles -OutDir $outDir -BaseName $meta.BaseName -Title ('{0} - {1}' -f $DashboardTitle, $meta.FileName)
        }
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
