<#
.SYNOPSIS
    Presents the "Wkly_Report" worksheet of the weekly ITSM workbook as an executive HTML dashboard
    in the "Design B" style (charcoal / green / gold) - strictly READ-ONLY.

.DESCRIPTION
    Workflow:  Excel workbook (Wkly_Report)  ->  this script  ->  <workbook name>.html dashboard

    This is a separate script from Export-WklyReportDashboard.ps1. It reads the workbook in exactly the
    same safe way, but presents every page in the Design B layout instead of Excel's own layout.

    SAFETY - the script never changes anything:
      * The original workbook is never opened. It is copied to a private temp folder and only the copy
        is opened, read-only, in a separate hidden Excel instance (links not updated, macros disabled).
        The copy is closed WITHOUT saving and deleted.
      * Only the "Wkly_Report" worksheet is read. No other tab is read or changed.
      * Excel windows you already have open are not touched. Only the Excel process this script
        started is closed at the end.
      * The SHA-256 hash of the original file is compared before and after the run (see the log).

    WHAT IS SHOWN:
      * The approved Design B sample layout: full-screen cover, one panel per report page with a
        coloured header bar, KPI tiles, charts in the sample's style, tables and text cards.
      * Every number, chart, table and sentence is taken from Wkly_Report: KPI cards are found by their
        label, charts by their title, tables and cards by their position on the page. Values are shown
        exactly as Excel displays them; nothing is recalculated. Anything that cannot be found is shown
        as "Not found in Wkly_Report" - never invented.
      * Values that Excel places over gaps in a sentence (separate text boxes) are joined into that
        sentence. A workbook chart that the sample has no card for is still shown, as an extra card.
      * A page that is added to the workbook later is shown in the same style (its KPI cards, charts and text).

.PARAMETER WorkbookPath
    The weekly workbook. Default: C:\Temp\Bi_Wkly_Report-Daily Data.xlsx

.PARAMETER OutputDirectory
    Where <workbook name>.html is written. Default: C:\Temp\ITSM_Dashboard_DesignB
    (a different folder from Export-WklyReportDashboard.ps1, so the two never overwrite each other).

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
    Text in the dashboard's top bar. Default "IT Service Management".

.PARAMETER Classification
    Marking shown on the top bar, the cover and every page footer. Default "CONFIDENTIAL".

.PARAMETER OpenWhenDone
    Open the dashboard in the default browser when finished.

.EXAMPLE
    .\Export-WklyReportDashboardB.ps1 -OpenWhenDone

.NOTES
    Requires Windows, Microsoft Excel (2010 or later) and Windows PowerShell 5.1 or PowerShell 7.
    No internet access, modules or third-party software are needed.
#>
[CmdletBinding()]
param(
    [string]$WorkbookPath = 'C:\Temp\Bi_Wkly_Report-Daily Data.xlsx',
    [string]$OutputDirectory = 'C:\Temp\ITSM_Dashboard_DesignB',
    [ValidateSet('Auto', 'Titles', 'PageBreaks', 'Manual')]
    [string]$PageMode = 'Auto',
    [int[]]$PageStartRows,
    [double]$TitleMinFontSize = 17,
    [double]$MinPageHeightPt = 250,
    [string]$DashboardTitle = 'IT Service Management',
    [string]$Classification = 'CONFIDENTIAL',
    [switch]$OpenWhenDone
)

$ScriptBuild = '2026.10.22-B2'
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
$script:LogPath = Join-Path (Join-Path ([System.IO.Path]::GetTempPath()) 'ITSM_Dashboard_logs') ('designB-log ' + (Get-Date -Format 'yyyy-MM-dd HHmmss') + '.txt')
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
        $t = @{}
        $pairs = ' .226 a .479 b .525 c .423 d .525 e .498 f .305 g .471 h .525 i .229 j .239 k .455 l .229 m .799 n .525 o .527 p .525 q .525 r .349 s .391 t .335 u .525 v .452 w .715 x .433 y .453 z .395 ' +
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
#  Design B: the approved sample layout, filled from Wkly_Report (pure logic - no Excel calls)
#  Every number, chart, table and sentence comes from the workbook. Fixed wording is limited to
#  the sample's card headings; anything that cannot be found is shown as "Not found in Wkly_Report".
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

function Get-BColorFor {
    # Dashboard colour for an Excel colour (same rule as the sample: reds stay red for problems).
    param([string]$Hex, [int]$Index)
    $order = @('--s1', '--s3', '--s4', '--c4', '--c5', '--c6')
    if ($Hex -match '^#([0-9a-fA-F]{2})([0-9a-fA-F]{2})([0-9a-fA-F]{2})$') {
        $r = [Convert]::ToInt32($Matches[1], 16) / 255.0; $g = [Convert]::ToInt32($Matches[2], 16) / 255.0; $b = [Convert]::ToInt32($Matches[3], 16) / 255.0
        $mx = [math]::Max($r, [math]::Max($g, $b)); $mn = [math]::Min($r, [math]::Min($g, $b)); $d = $mx - $mn
        if ($d -lt 0.12) { return '--s4' }
        $h = 0.0
        if ($mx -eq $r) { $h = 60 * ((($g - $b) / $d) % 6) } elseif ($mx -eq $g) { $h = 60 * (($b - $r) / $d + 2) } else { $h = 60 * (($r - $g) / $d + 4) }
        if ($h -lt 0) { $h += 360 }
        if ($h -lt 14 -or $h -ge 340) { return '--s8' }
        if ($h -lt 70) { return '--s3' }
        if ($h -lt 170) { return '--s1' }
    }
    return $order[$Index % $order.Count]
}

function Get-BKeywordColor {
    param([string]$Name, [int]$Index)
    $n = $Name.ToLowerInvariant()
    if ($n -match 'critical|breach|fail') { return '--s8' }
    if ($n -match 'high|resolved') { return '--s3' }
    if ($n -match 'moderate|medium|cancel') { return '--s4' }
    if ($n -match 'low|closed|met|within') { return '--s1' }
    if ($n -match 'progress') { return '--c4' }
    if ($n -match 'new') { return '--c5' }
    if ($n -match 'hold') { return '--c6' }
    return (@('--s1', '--s3', '--s4', '--c4', '--c5', '--c6'))[$Index % 6]
}

# ---------------------------------------------------------------- chart descriptors for the page script

function New-BChart {
    param([string]$Type, $Data)
    $Data['type'] = $Type
    return ('<div class="chart" data-c="{0}"></div>' -f [System.Net.WebUtility]::HtmlEncode(($Data | ConvertTo-Json -Depth 6 -Compress)))
}

function Get-BColumns {
    # Columns: bar series of an Excel chart; a line series named Open becomes the "n open" labels,
    # a % line becomes the share labels; per-bar Excel colours are kept (mapped to the palette).
    param($Spec, [switch]$Stack, [switch]$MonthNames, [switch]$WeekNames)
    $bars = @($Spec.series | Where-Object { $_.kind -in 'column', 'bar' })
    if ($bars.Count -eq 0) { return $null }
    $cats = New-Object System.Collections.Generic.List[string]
    $full = New-Object System.Collections.Generic.List[string]
    $n = 0; foreach ($s in $bars) { if ($s.values -and (ConvertTo-BArray $s.values).Count -gt $n) { $n = (ConvertTo-BArray $s.values).Count } }
    $catSrc = (ConvertTo-BArray $bars[0].cats)
    $keep = New-Object System.Collections.Generic.List[int]
    for ($i = 0; $i -lt $n; $i++) {
        $c = ''; if ($i -lt $catSrc.Count) { $c = [string]$catSrc[$i] }
        $allZero = $true; foreach ($s in $bars) { $v = (ConvertTo-BArray $s.values)[$i]; if ($null -ne $v -and [double]$v -ne 0) { $allZero = $false } }
        if (-not $c.Trim() -and $allZero) { continue }          # empty trailing months
        $keep.Add($i)
        $lab = $c.Trim()
        if ($MonthNames) { $lab = Get-BMonthLabel $c }
        if ($WeekNames -and $c -match '^\s*Week\s*(\d+)') { $lab = 'W' + $Matches[1] }
        $cats.Add($lab); $full.Add($c.Trim())
    }
    $series = New-Object System.Collections.Generic.List[object]
    $k = 0
    foreach ($s in $bars) {
        $vals = New-Object System.Collections.Generic.List[object]
        foreach ($i in $keep) { $v = (ConvertTo-BArray $s.values)[$i]; if ($null -eq $v) { $vals.Add($null) } else { $vals.Add([double]$v) } }
        $col = (@('--s1', '--s3', '--s4', '--c4'))[$k % 4]
        $entry = [ordered]@{ n = [string]$s.name; c = $col; v = $vals }
        $pts = (ConvertTo-BArray $s.points)
        $distinct = @($pts | Where-Object { $_ } | Sort-Object -Unique)
        if ($distinct.Count -ge 2) {
            $cols = New-Object System.Collections.Generic.List[string]
            foreach ($i in $keep) { $cols.Add((Get-BColorFor ([string]$pts[$i]) $k)) }
            $entry['colors'] = $cols
        }
        $series.Add($entry); $k++
    }
    $o = [ordered]@{}
    if ($Stack) { $o['stack'] = 1 }
    $open = @($Spec.series | Where-Object { $_.kind -eq 'line' -and [string]$_.name -match '^open$' })
    $pct = @($Spec.series | Where-Object { $_.kind -eq 'line' -and ([string]$_.name -match '%|percent' -or ($_.labels -and [string]$_.labels.format -match '%')) })
    if ($open.Count) {
        $labs = New-Object System.Collections.Generic.List[string]
        foreach ($i in $keep) { $v = (ConvertTo-BArray $open[0].values)[$i]; if ($null -eq $v) { $labs.Add('') } else { $labs.Add(('{0} open' -f [double]$v)) } }
        $o['labels'] = $labs
    }
    elseif ($pct.Count) {
        $labs = New-Object System.Collections.Generic.List[string]
        $fmt = '0.00%'; if ($pct[0].labels -and [string]$pct[0].labels.format -match '%') { $fmt = [string]$pct[0].labels.format }
        foreach ($i in $keep) { $labs.Add((ConvertTo-BNumberText ((ConvertTo-BArray $pct[0].values)[$i]) $fmt)) }
        $o['labels'] = $labs
    }
    $tipx = New-Object System.Collections.Generic.List[string]
    foreach ($f in $full) { $tipx.Add($f) }
    if ($WeekNames) { $o['tipx'] = $tipx }
    return [ordered]@{ cats = $cats; series = $series; o = $o }
}

function Get-BHBars {
    # Horizontal bars with "value · share" labels (share = the chart's % series, as formatted in Excel).
    param($Spec, [string]$Color = '--s1')
    $main = $null; $share = $null
    foreach ($s in $Spec.series) {
        $isPct = ([string]$s.name -match '%|percent') -or ($s.labels -and [string]$s.labels.format -match '%')
        if ($isPct) { if (-not $share) { $share = $s } } elseif (-not $main) { $main = $s }
    }
    if (-not $main) { return $null }
    $cats = New-Object System.Collections.Generic.List[string]; $vals = New-Object System.Collections.Generic.List[double]; $sh = New-Object System.Collections.Generic.List[string]
    $mc = (ConvertTo-BArray $main.cats); $mv = (ConvertTo-BArray $main.values)
    for ($i = 0; $i -lt $mv.Count; $i++) {
        if ($null -eq $mv[$i]) { continue }
        $cats.Add(([string]$mc[$i]).Trim()); $vals.Add([double]$mv[$i])
        if ($share) { $fmt = '0.00%'; if ($share.labels -and [string]$share.labels.format -match '%') { $fmt = [string]$share.labels.format }; $sh.Add((ConvertTo-BNumberText ((ConvertTo-BArray $share.values)[$i]) $fmt)) }
    }
    $o = [ordered]@{ c = $Color }
    if ($share) { $o['share'] = $sh }
    $lw = 140; foreach ($c in $cats) { $w = 9 * $c.Length + 20; if ($w -gt $lw) { $lw = $w } }
    $o['lw'] = [math]::Min(330, $lw)
    return [ordered]@{ cats = $cats; vals = $vals; o = $o }
}

function Get-BDonut {
    # Doughnut: counts are shown as counts, shares (0-1 values) as percentages; centre text given.
    param($Spec, [string]$Center, [string]$Sub, [switch]$StripGroup)
    $s0 = $Spec.series[0]
    if (-not $s0) { return $null }
    $vals = (ConvertTo-BArray $s0.values); $cats = (ConvertTo-BArray $s0.cats)
    $sum = 0.0; foreach ($v in $vals) { if ($null -ne $v) { $sum += [double]$v } }
    $isShare = ($sum -gt 0.98 -and $sum -lt 1.02)
    $parts = New-Object System.Collections.Generic.List[object]
    $best = $null
    for ($i = 0; $i -lt $vals.Count; $i++) {
        if ($null -eq $vals[$i]) { continue }
        $name = ([string]$cats[$i]).Trim()
        if ($StripGroup) { $name = $name -replace '^Group\d+\s+', '' }
        $v = [double]$vals[$i]
        $parts.Add(@($name, $v, (Get-BKeywordColor $name $i)))
        if ($null -eq $best -or $v -gt $best[1]) { $best = @($name, $v) }
    }
    $o = [ordered]@{ pct = $isShare }
    if ($Center) { $o['center'] = $Center } elseif ($best) {
        if ($isShare) { $o['center'] = ([double]$best[1] * 100).ToString('N2', $Inv) + '%' } else { $o['center'] = ([double]$best[1]).ToString('N0', $Inv) }
    }
    if ($Sub) { $o['sub'] = $Sub } elseif ($best) { $o['sub'] = ([string]$best[0]).ToLowerInvariant() }
    return [ordered]@{ parts = $parts; o = $o }
}

function Get-BLine {
    param($Spec, [string]$Color = '--s3')
    $s = $Spec.series[0]
    if (-not $s) { return $null }
    $cats = New-Object System.Collections.Generic.List[string]; $vals = New-Object System.Collections.Generic.List[double]
    $c = (ConvertTo-BArray $s.cats); $v = (ConvertTo-BArray $s.values)
    for ($i = 0; $i -lt $v.Count; $i++) { if ($null -eq $v[$i]) { continue }; $cats.Add((Get-BMonthLabel ([string]$c[$i]))); $vals.Add([double]$v[$i]) }
    return [ordered]@{ cats = $cats; vals = $vals; o = [ordered]@{ c = $Color } }
}

# ---------------------------------------------------------------- HTML pieces

function ConvertTo-BEnc { param([string]$T) return (ConvertTo-HtmlText $T) }

function ConvertTo-BTileHtml {
    param($Tile, [switch]$Big)
    if ($null -eq $Tile) { return '<div class="card stat"><div class="lab">&nbsp;</div><div class="val miss-v">—</div><div class="meta">Not found in Wkly_Report</div></div>' }
    if ($Tile.PSObject.Properties['Kind'] -and $Tile.Kind -eq 'quarter') {
        $val = ($Tile.Volumes | ForEach-Object { $_ }) -join ' → '
        $pills = ''
        for ($i = 0; $i -lt $Tile.Pcts.Count; $i++) {
            $p = [string]$Tile.Pcts[$i]
            $cls = 'wn'; $arrow = '▲'
            if ($p.StartsWith('-')) { $cls = 'bd'; $arrow = '▼' }
            $pills += ('<span class="pill {0}">{1} {2} Q{3}→Q{4}</span>' -f $cls, $arrow, (ConvertTo-BEnc $p.TrimStart('-')), ($i + 1), ($i + 2))
        }
        if ($Tile.Note) { $pills += ('<span>{0}</span>' -f (ConvertTo-BEnc $Tile.Note)) }
        return ('<div class="card stat"><div class="lab">{0}</div><div class="val" style="font-size:26px">{1}</div><div class="meta">{2}</div></div>' -f (ConvertTo-BEnc $Tile.Label), (ConvertTo-BEnc $val), $pills)
    }
    $meta = ConvertTo-BEnc $Tile.Meta
    if ($Tile.Meta -match '^(.*?)\s*\bMet\b\s*(.*)$') { $meta = '<span class="pill ok">✓ Met</span>' + (ConvertTo-BEnc (($Matches[1] + ' ' + $Matches[2]).Trim())) }
    elseif ($Tile.Label -match 'Breach') { $meta = '<span class="pill bd">! Breached</span>' + (ConvertTo-BEnc $Tile.Meta) }
    $bar = ''
    if ($Tile.Value -match '^(\d+(?:\.\d+)?)%$' -and $Tile.Label -match 'Rate') { $bar = ('<div class="bar"><i style="width:{0}%"></i></div>' -f $Matches[1]) }
    return ('<div class="card stat"><div class="lab">{0}</div><div class="val">{1}</div>{2}<div class="meta">{3}</div></div>' -f (ConvertTo-BEnc $Tile.Label), (ConvertTo-BEnc $Tile.Value), $bar, $meta)
}

function ConvertTo-BGrid {
    param($Tiles, [string]$Style = '')
    if ($Tiles.Count -eq 0) { return '' }
    $sb = New-Object System.Text.StringBuilder
    [void]$sb.Append(('<div class="grid" style="grid-template-columns:repeat({0},minmax(0,1fr));{1}">' -f [math]::Min(7, $Tiles.Count), $Style))
    foreach ($t in $Tiles) { [void]$sb.Append((ConvertTo-BTileHtml $t)) }
    [void]$sb.Append('</div>')
    return $sb.ToString()
}

function ConvertTo-BChartCard {
    param([string]$Title, [string]$Sub, [string]$Body)
    $h = ''
    if ($Title) { $h += '<h3>' + (ConvertTo-BEnc $Title) + '</h3>' }
    if ($Sub) { $h += '<div class="sub">' + (ConvertTo-BEnc $Sub) + '</div>' }
    if (-not $Body) { $Body = '<div class="miss">Not found in Wkly_Report</div>' }
    return ('<div class="card">{0}{1}</div>' -f $h, $Body)
}

function Get-BChartHtml {
    # A workbook chart in the sample's style; Excel's picture if it cannot be redrawn.
    param($It, [string]$As, [hashtable]$Opt)
    if ($null -eq $It) { return '' }
    $sp = $It.Spec
    $d = $null
    try {
        switch ($As) {
            'columns' { $d = Get-BColumns $sp -Stack:([bool]$Opt.Stack) -MonthNames:([bool]$Opt.Months) -WeekNames:([bool]$Opt.Weeks); if ($d -and $Opt.H) { $d.o['h'] = $Opt.H } }
            'hbars' { $c = '--s1'; if ($Opt.Color) { $c = $Opt.Color }; $d = Get-BHBars $sp $c }
            'donut' { $d = Get-BDonut $sp $Opt.Center $Opt.Sub -StripGroup:([bool]$Opt.Strip) }
            'line' { $c = '--s3'; if ($Opt.Color) { $c = $Opt.Color }; $d = Get-BLine $sp $c }
        }
    }
    catch { Add-Issue -Severity Warning -Location ("Chart '{0}'" -f $sp.title) -Message ('Could not be redrawn: {0}' -f $_.Exception.Message) }
    if ($d) { return (New-BChart $As $d) }
    if ($It.Src) { return ('<div class="bimg"><img src="{0}" alt="{1}"></div>' -f $It.Src, (ConvertTo-Attr ([string]$sp.title))) }
    return ''
}

function ConvertTo-BNotes {
    # Text cards (achievements, insights, summary): leading emoji in the icon box, a short head in
    # bold and the rest of the sentence below it - all workbook text.
    param($Items)
    $sb = New-Object System.Text.StringBuilder
    foreach ($n in $Items) {
        $head = [string]$n.Head; $body = [string]$n.Body
        $icon = '•'
        if ($head -match '^\s*([^\p{L}\p{N}\s]{1,4})\s*(.*)$') { $icon = $Matches[1]; $head = $Matches[2] }
        if (-not $body -and $head.Length -gt 46) {
            if ($head -match '^(.{12,}?)\s+[—–]\s+(.+)$') { $head = $Matches[1]; $body = $Matches[2] }
            elseif ($head -match '^(.{12,}?):\s+(.+)$') { $head = $Matches[1]; $body = $Matches[2] }
            elseif ($head -match '^(.{12,}?)\s+\((.+)\)\s*$') { $head = $Matches[1]; $body = $Matches[2] }
        }
        $cls = 'nt'
        if ($icon -match '✅|✔|🏆|✓') { $cls = 'ok' } elseif ($icon -match '⚠|❗|!') { $cls = 'wn' } elseif ($icon -match '⚡|📈|🚀') { $cls = 'wn' }
        [void]$sb.Append(('<div class="n"><span class="ic {0}">{1}</span><div><b>{2}</b><span>{3}</span></div></div>' -f $cls, (ConvertTo-BEnc $icon), (ConvertTo-BEnc $head), (ConvertTo-BEnc $body)))
    }
    return $sb.ToString()
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

function ConvertTo-BText {
    param([string]$Text)
    $cls = Get-StatusClass $Text
    if ($cls) { return ('<span class="pill {0}">{1}</span>' -f $cls, (ConvertTo-HtmlText $Text.Trim())) }
    return (ConvertTo-HtmlText $Text)
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

function Get-BLeftoverCharts {
    # Charts of the page that no card of the sample shows - added so nothing is left out.
    param($Page, $Used)
    $sb = New-Object System.Text.StringBuilder
    foreach ($it in ($Page.Items | Where-Object { $_.Kind -eq 'image' -and $_.IsChart -and $_.Spec } | Sort-Object T, L)) {
        $skip = $false
        foreach ($u in $Used) { if ([object]::ReferenceEquals($u, $it)) { $skip = $true } }
        if ($skip) { continue }
        $kind = [string]$it.Spec.kind
        $as = 'columns'
        if ($kind -eq 'bar') { $as = 'hbars' } elseif ($kind -in 'pie', 'doughnut') { $as = 'donut' } elseif ($kind -eq 'line') { $as = 'line' }
        if ($as -eq 'columns') {
            # long category names read better as horizontal bars
            $len = 0; $cnt = 0
            foreach ($c in (ConvertTo-BArray $it.Spec.series[0].cats)) { $len += ([string]$c).Length; $cnt++ }
            if ($cnt -gt 0 -and ($len / $cnt) -gt 14) { $as = 'hbars' }
        }
        [void]$sb.Append('<div class="grid g1" style="margin-top:16px">' + (ConvertTo-BChartCard ([string]$it.Spec.title) '' (Get-BChartHtml $it $as @{})) + '</div>')
    }
    return $sb.ToString()
}

function Build-BCover {
    param($Page, $Pages, $Meta)
    $tx = Get-BTexts $Page
    $title = $null
    foreach ($t in ($tx | Sort-Object @{ Expression = { $_.Size }; Descending = $true })) { $title = $t; break }
    $eyebrow = @($tx | Where-Object { $title -and $_.T -lt $title.T -and $_.Item.Kind -ne 'cell' } | Sort-Object T | ForEach-Object { $_.Text }) -join ' · '
    $kv = New-Object System.Collections.Generic.List[object]
    $tag = New-Object System.Collections.Generic.List[string]
    $conf = $Classification
    foreach ($t in ($tx | Where-Object { $title -and $_.T -gt $title.T + 2 -and $_.Item.Kind -ne 'cell' } | Sort-Object T, L)) {
        if ($t.Text -match '^\s*([^:]{3,30}?)\s*:\s*(.+)$') { $kv.Add(@($Matches[1], $Matches[2])); continue }
        if ($t.Text -match '^(CONFIDENTIAL|RESTRICTED|INTERNAL)') { $conf = $t.Text; continue }
        if ($kv.Count -eq 0) { $tag.Add($t.Text) }
    }
    $exec = Get-BPage $Pages 'Executive'
    $execTiles = New-Object System.Collections.Generic.List[object]
    if ($exec) { $execTiles = Get-BTiles $exec (Get-BTexts $exec) }
    $sla = Get-BPage $Pages 'SLA'
    # stat 1: total tickets (helper table SAMI_AEC / Total), 2: resolution SLA, 3: average resolution time
    $total = Get-BCellValue $Pages 'SAMI_AEC' 'Total'
    $inc = Get-BCellValue $Pages 'SAMI_AEC' 'Inc'
    $srs = Get-BCellValue $Pages 'SAMI_AEC' 'Total SR'
    $s1 = $null
    if ($total) { $m = ''; if ($inc -and $srs) { $m = ('{0} incidents · {1} service requests' -f $inc, $srs) }; $s1 = [pscustomobject]@{ Label = 'Total tickets (YTD)'; Value = $total; Meta = $m } }
    $s2 = Find-BTile $execTiles 'Resolution SLA'
    $s3 = Find-BTile $execTiles 'MTTR Resolution'
    if ($s3 -and $sla) {
        $st = Get-BTexts $sla
        $resLab = @($st | Where-Object { $_.Text -eq 'RESOLUTION' } | Select-Object -First 1)
        if ($resLab.Count) { $line = @($st | Where-Object { $_.T -gt $resLab[0].T -and $_.T -lt $resLab[0].T + 90 -and $_.Text -match 'min SLA' } | Select-Object -First 1); if ($line.Count) { $s3 = [pscustomobject]@{ Label = 'Average resolution time'; Value = $s3.Value; Meta = $line[0].Text } } }
    }
    if ($s2) { $s2 = [pscustomobject]@{ Label = $s2.Label; Value = $s2.Value; Meta = $s2.Meta } }
    $sb = New-Object System.Text.StringBuilder
    [void]$sb.Append('<div class="cover"><div class="card hero">')
    if ($eyebrow) { [void]$sb.Append('<div class="eyebrow">' + (ConvertTo-BEnc $eyebrow) + '</div>') }
    $tt = [string]$Page.Title; if ($title) { $tt = $title.Text }
    [void]$sb.Append('<h1>' + (ConvertTo-BEnc $tt) + '</h1>')
    foreach ($t in $tag) { [void]$sb.Append('<div class="tg">' + (ConvertTo-BEnc $t) + '</div>') }
    if ($kv.Count) {
        [void]$sb.Append('<dl class="kv">')
        foreach ($p in $kv) { [void]$sb.Append('<dt>' + (ConvertTo-BEnc $p[0]) + '</dt><dd>' + (ConvertTo-BEnc $p[1]) + '</dd>') }
        [void]$sb.Append('</dl>')
    }
    if ($conf) { [void]$sb.Append('<span class="conf-big">' + (ConvertTo-BEnc $conf) + '</span>') }
    [void]$sb.Append('</div><div class="grid" style="grid-template-rows:repeat(3,1fr)">')
    foreach ($s in @($s1, $s2, $s3)) { [void]$sb.Append((ConvertTo-BTileHtml $s)) }
    [void]$sb.Append('</div></div>')
    return $sb.ToString()
}

function Build-BExec {
    param($Page, $Used)
    $tx = Get-BTexts $Page
    $tiles = Get-BTiles $Page $tx
    $sb = New-Object System.Text.StringBuilder
    [void]$sb.Append((ConvertTo-BGrid $tiles))
    $mon = Get-BChartItem $Page 'Monthly'; $stat = Get-BChartItem $Page 'Status'
    $wsr = Get-BChartItem $Page 'Weekly Service'; $winc = Get-BChartItem $Page 'Weekly Incident'
    foreach ($c in @($mon, $stat, $wsr, $winc)) { Add-BUsedChart $Used $c }
    [void]$sb.Append('<div class="grid g12" style="margin-top:16px">')
    [void]$sb.Append((ConvertTo-BChartCard ([string]$mon.Spec.title) 'Incidents and service requests, stacked = total' (Get-BChartHtml $mon 'columns' @{ Stack = $true; Months = $true; H = 340 })))
    [void]$sb.Append((ConvertTo-BChartCard ([string]$stat.Spec.title) 'Share of all incidents' (Get-BChartHtml $stat 'donut' @{ Strip = $true })))
    [void]$sb.Append('</div><div class="grid g2" style="margin-top:16px">')
    [void]$sb.Append((ConvertTo-BChartCard ((([string]$wsr.Spec.title) -replace "[`r`n]+", ' ')) 'Created vs resolved / closed · open shown as label' (Get-BChartHtml $wsr 'columns' @{ Weeks = $true })))
    [void]$sb.Append((ConvertTo-BChartCard ([string]$winc.Spec.title) 'Created vs closed · open shown as label' (Get-BChartHtml $winc 'columns' @{ Weeks = $true })))
    [void]$sb.Append('</div>')
    $ach = @($tx | Where-Object { $_.Text -match '^ACHIEVEMENTS' } | Select-Object -First 1)
    if ($ach.Count) {
        $notes = New-Object System.Collections.Generic.List[object]
        foreach ($c in (Get-BCardTexts $Page $tx $ach[0].T ($ach[0].T + 140) 150)) { $notes.Add([pscustomobject]@{ Head = (@($c.Texts | ForEach-Object { $_.Text }) -join ' '); Body = '' }) }
        if ($notes.Count) { [void]$sb.Append('<div class="card notes" style="margin-top:16px"><h3>Achievements</h3><div class="grid g4" style="gap:0 24px">' + (ConvertTo-BNotes $notes) + '</div></div>') }
    }
    return $sb.ToString()
}

function Build-BIncident {
    param($Page, $Used)
    $tx = Get-BTexts $Page
    $tiles = Get-BTiles $Page $tx
    $total = Find-BTile $tiles '^Total Incidents'
    $pri = Get-BChartItem $Page 'Priority'; $cls = Get-BChartItem $Page 'Classification'
    Add-BUsedChart $Used $pri; Add-BUsedChart $Used $cls
    $center = $null; if ($total) { $center = $total.Value }
    $sb = New-Object System.Text.StringBuilder
    [void]$sb.Append((ConvertTo-BGrid $tiles))
    [void]$sb.Append('<div class="grid g21" style="margin-top:16px">')
    [void]$sb.Append((ConvertTo-BChartCard ([string]$pri.Spec.title) 'Incidents' (Get-BChartHtml $pri 'donut' @{ Center = $center; Sub = 'incidents' })))
    [void]$sb.Append((ConvertTo-BChartCard ([string]$cls.Spec.title) 'Incidents, share of total' (Get-BChartHtml $cls 'hbars' @{})))
    [void]$sb.Append('</div>')
    return $sb.ToString()
}

function Build-BServiceRequests {
    param($Page, $Used)
    $tx = Get-BTexts $Page
    $tiles = Get-BTiles $Page $tx
    $mon = Get-BChartItem $Page 'Monthly'; $cat = Get-BChartItem $Page 'category'
    Add-BUsedChart $Used $mon; Add-BUsedChart $Used $cat
    $sb = New-Object System.Text.StringBuilder
    [void]$sb.Append((ConvertTo-BGrid $tiles))
    [void]$sb.Append('<div class="grid g2" style="margin-top:16px">')
    [void]$sb.Append((ConvertTo-BChartCard ([string]$mon.Spec.title) 'Service requests per month' (Get-BChartHtml $mon 'line' @{ Color = '--s3' })))
    [void]$sb.Append((ConvertTo-BChartCard ([string]$cat.Spec.title) 'Service requests, share of total' (Get-BChartHtml $cat 'hbars' @{ Color = '--s3' })))
    [void]$sb.Append('</div>')
    return $sb.ToString()
}

function Build-BSla {
    param($Page, $Used)
    $tx = Get-BTexts $Page
    $tiles = Get-BTiles $Page $tx -Bottom ($Page.StartY + 150)
    $res = Get-BChartItem $Page 'Resolution SLA'; $rsp = Get-BChartItem $Page 'Response SLA'
    Add-BUsedChart $Used $res; Add-BUsedChart $Used $rsp
    $sb = New-Object System.Text.StringBuilder
    [void]$sb.Append((ConvertTo-BGrid $tiles))
    [void]$sb.Append('<div class="grid g3" style="margin-top:16px">')
    [void]$sb.Append((ConvertTo-BChartCard ([string]$res.Spec.title) 'Incidents within SLA' (Get-BChartHtml $res 'donut' @{ Sub = 'within SLA' })))
    [void]$sb.Append((ConvertTo-BChartCard ([string]$rsp.Spec.title) 'Incidents within SLA' (Get-BChartHtml $rsp 'donut' @{ Sub = 'within SLA' })))
    # MTTR panel: response / resolution use of the SLA budget, as written in the workbook
    $mt = @($tx | Where-Object { $_.Text -eq 'MTTR' } | Select-Object -First 1)
    $body = ''
    if ($mt.Count) {
        $zone = @($tx | Where-Object { $_.L -ge $mt[0].L - 20 -and $_.T -ge $mt[0].T -and $_.T -le $mt[0].T + 300 } | Sort-Object T, L)
        foreach ($part in @('RESPONSE', 'RESOLUTION')) {
            $lab = @($zone | Where-Object { $_.Text -eq $part } | Select-Object -First 1)
            if (-not $lab.Count) { continue }
            $line = @($zone | Where-Object { $_.T -gt $lab[0].T -and $_.T -lt $lab[0].T + 100 -and $_.Text -match 'min SLA' } | Select-Object -First 1)
            $big = @($zone | Where-Object { $_.T -gt $lab[0].T -and $_.T -lt $lab[0].T + 40 -and (Test-BValue $_.Text) } | Sort-Object @{ Expression = { $_.Size }; Descending = $true } | Select-Object -First 1)
            $pct = ''; $txt = ''
            if ($line.Count) { $txt = $line[0].Text; if ($txt -match '\(\s*([\d.]+%)\s*\)') { $pct = $Matches[1] } }
            $name = (Get-Culture).TextInfo.ToTitleCase($part.ToLowerInvariant())
            $lt = $name
            $lineTxt = ($txt -replace '\s*\(\s*[\d.]+%\s*\)', '').Trim()
            if ($big.Count -and -not $lineTxt.StartsWith($big[0].Text)) { $lt += ' · ' + $big[0].Text }
            if ($lineTxt) { $lt += ' · ' + $lineTxt }
            $col = '--s3'; if ($part -eq 'RESOLUTION') { $col = '--s1' }
            $w = 0; if ($pct -match '^([\d.]+)%$') { $w = [double]::Parse($Matches[1], $Inv) }
            $body += ('<div class="stat" style="margin-top:12px"><div class="lab">{0}</div><div class="val" style="font-size:29px">{1}</div><div class="bar"><i style="width:{2}%;background:var({3})"></i></div></div>' -f (ConvertTo-BEnc $lt), (ConvertTo-BEnc $pct), (Format-Num ([math]::Min(100, $w))), $col)
        }
        $sr = @($zone | Where-Object { $_.Text -match 'SR SLA' } | Select-Object -First 1)
        if ($sr.Count) {
            $after = @($tx | Where-Object { $_.T -gt $sr[0].T - 1 -and $_.T -lt $sr[0].T + 40 -and $_.L -ge $sr[0].L - 30 -and -not [object]::ReferenceEquals($_, $sr[0]) } | Sort-Object @{ Expression = { [math]::Round($_.CY / 6) } }, L | ForEach-Object { $_.Text })
            $body += ('<div class="meta" style="margin-top:16px;font-size:16px;color:var(--ink-3)">{0} {1}</div>' -f (ConvertTo-BEnc ($sr[0].Text -replace '^[^\p{L}]+', '')), (ConvertTo-BEnc (($after -join ' ') -replace '\s*/\s*', ' / ')))
        }
    }
    [void]$sb.Append((ConvertTo-BChartCard 'MTTR vs SLA budget' 'How much of each SLA is used' $body))
    [void]$sb.Append('</div>')
    $ins = @($tx | Where-Object { $_.Text -match 'INSIGHTS' } | Select-Object -First 1)
    if ($ins.Count) {
        $notes = New-Object System.Collections.Generic.List[object]
        foreach ($c in (Get-BCardTexts $Page $tx $ins[0].T ($ins[0].T + 60) 150)) {
            $parts = @($c.Texts | Sort-Object T, L)
            $notes.Add([pscustomobject]@{ Head = $parts[0].Text; Body = (@($parts | Select-Object -Skip 1 | ForEach-Object { $_.Text }) -join ' ') })
        }
        $head = ($ins[0].Text -replace '^.*?[—–-]\s*', '')
        if ($notes.Count) { [void]$sb.Append(('<div class="card notes" style="margin-top:16px"><h3>{0}</h3><div class="grid g4" style="gap:0 24px">{1}</div></div>' -f (ConvertTo-BEnc $head), (ConvertTo-BNotes $notes))) }
    }
    return $sb.ToString()
}

function Build-BPriority {
    param($Page, $Used)
    $pd = Get-BChartItem $Page 'Priority'; $top = Get-BChartItem $Page 'Incident.*Categor'
    Add-BUsedChart $Used $pd; Add-BUsedChart $Used $top
    $sb = New-Object System.Text.StringBuilder
    [void]$sb.Append('<div class="grid g2">')
    [void]$sb.Append((ConvertTo-BChartCard ([string]$pd.Spec.title) 'Incidents vs service requests' (Get-BChartHtml $pd 'columns' @{ H = 340 })))
    [void]$sb.Append((ConvertTo-BChartCard ([string]$top.Spec.title) 'Count and share' (Get-BChartHtml $top 'hbars' @{})))
    [void]$sb.Append('</div>')
    return $sb.ToString()
}

function Build-BAgents {
    param($Page, $Used)
    $tx = Get-BTexts $Page
    $ag = Get-BChartItem $Page 'Agent'
    Add-BUsedChart $Used $ag
    $sb = New-Object System.Text.StringBuilder
    [void]$sb.Append('<div class="grid g2">')
    [void]$sb.Append((ConvertTo-BChartCard ([string]$ag.Spec.title) 'Tickets handled' (Get-BChartHtml $ag 'hbars' @{})))
    $pic = Get-BPicture $Page
    $picHtml = ''
    if ($pic) { $picHtml = ('<div class="bimg"><img src="{0}" alt="Assignment groups"></div><div class="note">This table is a linked picture in Wkly_Report, so it is shown as Excel''s picture of it.</div>' -f $pic.Src) }
    [void]$sb.Append((ConvertTo-BChartCard 'Assignment groups' 'Incidents and average MTTR' $picHtml))
    [void]$sb.Append('</div>')
    $enh = @($tx | Where-Object { $_.Text -match '^ENHANCEMENTS' } | Select-Object -First 1)
    if ($enh.Count) {
        $sub = @($tx | Where-Object { $_.T -gt $enh[0].T -and $_.T -lt $enh[0].T + 30 -and $_.Text -ne $enh[0].Text } | Select-Object -First 1 | ForEach-Object { $_.Text })
        $cards = New-Object System.Text.StringBuilder
        foreach ($c in ((Get-BCards $Page 8000 80000) | Where-Object { $_.T -gt $enh[0].T -and $_.T -lt $enh[0].T + 120 } | Sort-Object L)) {
            $in = @($tx | Where-Object { Test-BInside $_ $c } | Sort-Object T, L)
            $id = @($in | Where-Object { $_.Text -match '^[A-Z]\d{1,2}$' } | Select-Object -First 1 | ForEach-Object { $_.Text })
            $st = @($in | Where-Object { $_.Size -le 7.5 -and $_.Text.Length -le 16 } | Select-Object -First 1 | ForEach-Object { $_.Text })
            $rest = @($in | Where-Object { $_.Text -notmatch '^[A-Z]\d{1,2}$' -and -not ($_.Size -le 7.5 -and $_.Text.Length -le 16) })
            $ttl = @($rest | Sort-Object @{ Expression = { $_.Size }; Descending = $true }, T | Select-Object -First 1)
            $body = @($rest | Where-Object { $ttl.Count -eq 0 -or -not [object]::ReferenceEquals($_, $ttl[0]) } | ForEach-Object { $_.Text }) -join ' '
            $cls = 'nt'; $s0 = [string]($st -join ''); if ($s0 -match 'progress') { $cls = 'wn' } elseif ($s0 -match 'complete|done') { $cls = 'ok' }
            $tt = ''; if ($ttl.Count) { $tt = $ttl[0].Text }
            [void]$cards.Append(('<div class="card"><div style="display:flex;justify-content:space-between"><span class="id">{0}</span><span class="pill {1}">{2}</span></div><h4>{3}</h4><p>{4}</p></div>' -f (ConvertTo-BEnc ($id -join '')), $cls, (ConvertTo-BEnc $s0), (ConvertTo-BEnc $tt), (ConvertTo-BEnc $body)))
        }
        $title = (Get-Culture).TextInfo.ToTitleCase($enh[0].Text.ToLowerInvariant())
        [void]$sb.Append(('<div class="card" style="margin-top:16px"><h3>{0}</h3><div class="sub">{1}</div><div class="enh">{2}</div></div>' -f (ConvertTo-BEnc $title), (ConvertTo-BEnc ($sub -join '')), $cards.ToString()))
    }
    return $sb.ToString()
}

function Build-BResolution {
    param($Page, $Used)
    $tx = Get-BTexts $Page
    $ia = Get-BChartItem $Page 'Incident.*Aging'; $sa = Get-BChartItem $Page 'Service Request Aging'
    Add-BUsedChart $Used $ia; Add-BUsedChart $Used $sa
    $sb = New-Object System.Text.StringBuilder
    [void]$sb.Append('<div class="grid g2">')
    [void]$sb.Append((ConvertTo-BChartCard ([string]$ia.Spec.title) 'Incidents by resolution time · share as label' (Get-BChartHtml $ia 'columns' @{})))
    [void]$sb.Append((ConvertTo-BChartCard ([string]$sa.Spec.title) 'SRs by resolution time · share as label' (Get-BChartHtml $sa 'columns' @{})))
    [void]$sb.Append('</div>')
    $ki = @($tx | Where-Object { $_.Text -match '^Key Insights' } | Select-Object -First 1)
    if ($ki.Count) {
        $tiles = New-Object System.Collections.Generic.List[object]
        foreach ($c in ((Get-BCards $Page 8000 80000) | Where-Object { $_.T -gt $ki[0].T -and $_.T -lt $ki[0].T + 60 } | Sort-Object L)) {
            $in = @($tx | Where-Object { Test-BInside $_ $c } | Sort-Object T, L)
            $v = @($in | Where-Object { Test-BValue $_.Text } | Select-Object -First 1)
            if (-not $v.Count) { continue }
            $others = @($in | Where-Object { -not [object]::ReferenceEquals($_, $v[0]) } | Sort-Object @{ Expression = { $_.Size } }, T)
            $desc = ''; $unit = ''
            if ($others.Count) { $desc = $others[0].Text; $unit = (@($others | Select-Object -Skip 1 | ForEach-Object { $_.Text }) -join ' ') }
            $tiles.Add([pscustomobject]@{ Label = $desc; Value = $v[0].Text; Meta = $unit })
        }
        if ($tiles.Count) { [void]$sb.Append('<h3 style="margin:22px 0 10px">' + (ConvertTo-BEnc $ki[0].Text) + '</h3>' + (ConvertTo-BGrid $tiles)) }
    }
    return $sb.ToString()
}

function Build-BRisks {
    param($Page, $Used)
    $tx = Get-BTexts $Page
    $tb = Get-BShapeTable $Page $tx '^Risk$'
    if ($null -eq $tb -or $tb.Rows.Count -eq 0) { return '' }
    $heads = @($tb.Heads | ForEach-Object { $_.Text })
    $sb = New-Object System.Text.StringBuilder
    [void]$sb.Append('<div class="card"><table><thead><tr><th>ID</th>')
    $hdr = @($heads[0])
    for ($k = 2; $k -lt $heads.Count; $k++) { $hdr += $heads[$k] }
    foreach ($h in $hdr) { [void]$sb.Append('<th>' + (ConvertTo-BEnc $h) + '</th>') }
    [void]$sb.Append('</tr></thead><tbody>')
    foreach ($r in $tb.Rows) {
        $first = New-Object System.Collections.Generic.List[object]
        foreach ($k in 0, 1) { if ($r.Cols.ContainsKey($k)) { foreach ($t in $r.Cols[$k]) { $first.Add($t) } } }
        $id = @($first | Where-Object { $_.Text -match '^[A-Z]\d{1,2}$' } | Select-Object -First 1 | ForEach-Object { $_.Text })
        $txts = @($first | Where-Object { $_.Text -notmatch '^[A-Z]\d{1,2}$' } | Sort-Object T, L)
        $ttl = ''; $desc = ''
        if ($txts.Count) { $ttl = $txts[0].Text; $desc = (@($txts | Select-Object -Skip 1 | ForEach-Object { $_.Raw -replace "^\s*-\s*", '' -replace "[`r`n`v]+\s*-?\s*", '; ' }) -join ' ').Trim() }
        [void]$sb.Append('<tr><td><b>' + (ConvertTo-BEnc ($id -join '')) + '</b></td><td><b>' + (ConvertTo-BEnc $ttl) + '</b><br><span style="color:var(--ink-2)">' + (ConvertTo-BEnc $desc) + '</span></td>')
        for ($k = 2; $k -lt $heads.Count; $k++) {
            $cell = ''
            if ($r.Cols.ContainsKey($k)) { $cell = (@($r.Cols[$k] | ForEach-Object { $_.Text }) -join ' ') }
            if ($k -lt $heads.Count - 1) { [void]$sb.Append('<td>' + (ConvertTo-BText $cell) + '</td>') }
            else { [void]$sb.Append('<td style="color:var(--ink-2)">' + (ConvertTo-BEnc $cell) + '</td>') }
        }
        [void]$sb.Append('</tr>')
    }
    [void]$sb.Append('</tbody></table></div>')
    return $sb.ToString()
}

function Build-BSummary {
    param($Page, $Used)
    $tx = Get-BTexts $Page
    $rec = @($tx | Where-Object { $_.Text -match '^Recommendations' } | Select-Object -First 1)
    $split = 470.0; if ($rec.Count) { $split = $rec[0].L - 5 }
    $log = @($tx | Where-Object { $_.Text -match 'UPDATE LOG' } | Select-Object -First 1)
    $bottom = 1e9; if ($log.Count) { $bottom = $log[0].T - 2 }
    $top = $Page.StartY + 60
    if ($rec.Count) { $top = $rec[0].T - 5 }
    # summary (left): bigger heading followed by its sentence
    $left = @($tx | Where-Object { $_.L -lt $split -and $_.T -gt $top - 10 -and $_.T -lt $bottom } | Sort-Object T, L)
    $notes = New-Object System.Collections.Generic.List[object]
    $cur = $null
    foreach ($t in $left) {
        if ($t.Size -ge 10) { $cur = [pscustomobject]@{ Head = $t.Text; Body = '' }; $notes.Add($cur); continue }
        if ($cur) { $cur.Body = ($cur.Body + ' ' + $t.Text).Trim() }
    }
    # recommendations (right): number, title, sentence
    $right = @($tx | Where-Object { $_.L -ge $split -and $_.T -gt $top + 5 -and $_.T -lt $bottom } | Sort-Object T, L)
    $recs = New-Object System.Collections.Generic.List[object]
    $cur = $null
    foreach ($t in $right) {
        if ($t.Text -match '^\d{2}$') { continue }
        if ($t.Size -ge 10) { $cur = [pscustomobject]@{ Head = $t.Text; Body = '' }; $recs.Add($cur); continue }
        if ($cur) { $cur.Body = ($cur.Body + ' ' + $t.Text).Trim() }
    }
    $sb = New-Object System.Text.StringBuilder
    [void]$sb.Append('<div class="grid g2"><div class="card notes"><h3>Summary</h3>' + (ConvertTo-BNotes $notes) + '</div>')
    $rt = 'Recommendations'; if ($rec.Count) { $rt = $rec[0].Text }
    [void]$sb.Append('<div class="card"><h3>' + (ConvertTo-BEnc $rt) + '</h3><ol class="reco">')
    foreach ($r in $recs) { [void]$sb.Append('<li><div><b>' + (ConvertTo-BEnc $r.Head) + '</b><br><span style="color:var(--ink-2);font-size:17px">' + (ConvertTo-BEnc $r.Body) + '</span></div></li>') }
    [void]$sb.Append('</ol></div></div>')
    $tb = Get-BShapeTable $Page $tx '^Week$'
    if ($tb -and $tb.Rows.Count) {
        $title = 'Weekly update log'; if ($log.Count) { $title = (Get-Culture).TextInfo.ToTitleCase((($log[0].Text -replace '^[^\p{L}]+', '')).ToLowerInvariant()) }
        [void]$sb.Append('<div class="card" style="margin-top:16px"><h3>' + (ConvertTo-BEnc $title) + '</h3><table><thead><tr>')
        foreach ($h in $tb.Heads) { [void]$sb.Append('<th>' + (ConvertTo-BEnc $h.Text) + '</th>') }
        [void]$sb.Append('</tr></thead><tbody>')
        $skipped = 0
        foreach ($r in $tb.Rows) {
            $vals = New-Object System.Collections.Generic.List[string]
            for ($k = 0; $k -lt $tb.Heads.Count; $k++) { $v = ''; if ($r.Cols.ContainsKey($k)) { $v = (@($r.Cols[$k] | ForEach-Object { $_.Text }) -join ' ') }; $vals.Add($v) }
            $empty = $true
            for ($k = 1; $k -lt $vals.Count; $k++) { if ($vals[$k] -and $vals[$k] -ne '[]') { $empty = $false } }
            if ($empty) { $skipped++; continue }
            [void]$sb.Append('<tr>')
            foreach ($v in $vals) { [void]$sb.Append('<td>' + (ConvertTo-BEnc $v) + '</td>') }
            [void]$sb.Append('</tr>')
        }
        [void]$sb.Append('</tbody></table></div>')
        if ($skipped) { Write-Log ("Design B: {0} empty row(s) of the weekly update log (only '[]' placeholders) are not shown." -f $skipped) }
    }
    return $sb.ToString()
}

function Build-BGeneric {
    # A page added to the workbook later: its KPI cards, charts and remaining text, in the same style.
    param($Page, $Used)
    $tx = Get-BTexts $Page
    $tiles = Get-BTiles $Page $tx
    $sb = New-Object System.Text.StringBuilder
    [void]$sb.Append((ConvertTo-BGrid $tiles))
    $inTiles = @{}
    foreach ($t in $tiles) { if ($t.PSObject.Properties['Texts']) { foreach ($x in $t.Texts) { $inTiles[(Get-ItemId $x.Item)] = $true } } }
    $restList = New-Object System.Collections.Generic.List[object]
    foreach ($x in ($tx | Sort-Object T, L)) { if (-not $inTiles.ContainsKey((Get-ItemId $x.Item)) -and $x.Size -lt 17) { $restList.Add($x) } }
    $rest = @($restList | ForEach-Object { $_ })
    if ($rest.Count) {
        $notes = New-Object System.Collections.Generic.List[object]
        foreach ($t in $rest) { $notes.Add([pscustomobject]@{ Head = $t.Text; Body = '' }) }
        [void]$sb.Append('<div class="card notes" style="margin-top:16px">' + (ConvertTo-BNotes $notes) + '</div>')
    }
    return $sb.ToString()
}

function New-DesignBSite {
    # Writes ONE self-contained .html to $OutDir; pictures are read from $WorkDir\assets and embedded.
    param($Pages, [string]$OutDir, $Meta, [string]$WorkDir)
    $sections = New-Object System.Text.StringBuilder
    $n = 0
    foreach ($p in $Pages) {
        $n++
        $used = New-Object System.Collections.Generic.List[object]
        $tx = Get-BTexts $p
        $body = ''
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
        try {
            switch ($kind) {
                'cover' { $body = Build-BCover $p $Pages $Meta }
                'exec' { $body = Build-BExec $p $used }
                'incident' { $body = Build-BIncident $p $used }
                'sr' { $body = Build-BServiceRequests $p $used }
                'sla' { $body = Build-BSla $p $used }
                'priority' { $body = Build-BPriority $p $used }
                'agents' { $body = Build-BAgents $p $used }
                'resolution' { $body = Build-BResolution $p $used }
                'risks' { $body = Build-BRisks $p $used }
                'summary' { $body = Build-BSummary $p $used }
                default { $body = Build-BGeneric $p $used }
            }
            if ($kind -ne 'cover') { $body += Get-BLeftoverCharts $p $used }
        }
        catch {
            Add-Issue -Severity Warning -Location ('Page {0}' -f $p.Index) -Message ('Design B page could not be built: {0}' -f (Format-ErrorDetail $_))
            $body = '<div class="card"><div class="miss">This page could not be built - see the log.</div></div>'
        }
        Write-Log ("  page {0} '{1}': Design B {2} page" -f $p.Index, $t, $kind)
        if ($kind -eq 'cover') {
            [void]$sections.Append(('<section id="s{0}" data-t="{1}">{2}</section>' -f $n, (ConvertTo-Attr 'Cover'), $body))
            continue
        }
        $sub = Get-BSubtitle $p $tx
        [void]$sections.Append(('<section id="s{0}" data-t="{1}"><div class="sh"><span class="no">{0:D2}</span><h2>{2}</h2><p>{3}</p></div>{4}</section>' -f $n, (ConvertTo-Attr $t), (ConvertTo-BEnc $t), (ConvertTo-BEnc $sub), $body))
        [void]$sections.Append("`n")
    }
    $html = $DesignBShell
    $html = $html.Replace('{{TITLE}}', (ConvertTo-BEnc $DashboardTitle))
    $html = $html.Replace('{{CSS}}', $DesignBCss)
    $chips = ''
    if ($Meta.ReportingPeriod) { $chips += '<span class="chip">Period <b>' + (ConvertTo-BEnc $Meta.ReportingPeriod) + '</b></span>' }
    if ($Meta.ReportDate) { $chips += '<span class="chip">Report date <b>' + (ConvertTo-BEnc $Meta.ReportDate) + '</b></span>' }
    $html = $html.Replace('{{CHIPS}}', $chips)
    $html = $html.Replace('{{CONF}}', (ConvertTo-BEnc $Classification))
    $html = $html.Replace('{{FILE}}', (ConvertTo-BEnc $Meta.FileName))
    $foot = ''
    if ($Meta.ReportingPeriod) { $foot = (ConvertTo-BEnc $Meta.ReportingPeriod) + '<br>' }
    $foot += 'Source: Wkly_Report &middot; read-only<br>Build ' + $ScriptBuild
    $html = $html.Replace('{{RAILFOOT}}', $foot)
    $html = $html.Replace('{{SECTIONS}}', $sections.ToString())
    $html = $html.Replace('{{JS}}', $DesignBJs)
    # every picture is embedded so the single .html works when sent by email or Teams
    $embedded = @{}
    foreach ($m in [regex]::Matches($html, 'src="(assets/[^"]+)"')) {
        $rel = $m.Groups[1].Value
        if ($embedded.ContainsKey($rel)) { continue }
        $file = Join-Path $WorkDir ($rel.Replace('/', [string][System.IO.Path]::DirectorySeparatorChar))
        if (-not (Test-Path -LiteralPath $file)) { Write-Log ("Picture {0} is missing and could not be embedded." -f $rel) 'WARN'; continue }
        $ext = [System.IO.Path]::GetExtension($file).ToLowerInvariant().TrimStart('.')
        $mime = 'image/png'
        if ($ext -eq 'jpg' -or $ext -eq 'jpeg') { $mime = 'image/jpeg' } elseif ($ext -eq 'gif') { $mime = 'image/gif' } elseif ($ext -eq 'emf') { $mime = 'image/emf' }
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
#  Design B page shell, stylesheet and script (the approved sample, embedded in the single .html)
# ======================================================================================

$DesignBShell = @'
<!doctype html>
<html lang="en">
<head>
<meta charset="utf-8">
<meta name="viewport" content="width=device-width,initial-scale=1">
<title>{{TITLE}}</title>
<style>{{CSS}}</style>
</head>
<body data-conf="{{CONF}}">

<div class="scrim" id="scrim"></div>
<aside class="rail" id="rail">
  <button class="x" id="mClose" title="Close (Esc)">&times;</button>
  <div class="brand"><div class="logo"><span class="dot">IT</span>{{TITLE}}</div><small>{{FILE}} · Wkly_Report</small></div>
  <nav id="nav"></nav>
  <div class="foot">{{RAILFOOT}}</div>
</aside>

<main>
<div class="top">
  <button class="menu-btn" id="mOpen" title="Pages"><span class="bars">&#9776;</span>Pages <small id="mPos">1 / 10</small></button>
  <h1>{{TITLE}}</h1>
  {{CHIPS}}
  <div class="zoom" title="Zoom (keys + and -)"><button class="btn" id="zOut">&minus;</button><span id="zVal">100%</span><button class="btn" id="zIn">+</button><button class="btn" id="zFit">Fit</button></div>
  <button class="btn" id="theme" title="Switch light / dark">Light / Dark</button>
  <button class="btn" id="present">Present</button>
  <button class="btn" onclick="window.print()">Print</button>
  <span class="tag-conf">{{CONF}}</span>
</div>
<div id="content">

{{SECTIONS}}
</div></main>
<div class="hud"><span id="hpos"></span><button class="btn" id="hp">&#9664;</button><button class="btn" id="hn">&#9654;</button><button class="btn" id="hx">Exit</button></div>
<script>{{JS}}</script>
</body>
</html>

'@

$DesignBCss = @'

:root{
  color-scheme:light;
  --bg:#eeebe3;--surface:#ffffff;--surface-2:#f7f5ef;--line:#e1dccf;--ink:#0b0b0b;--ink-2:#52514e;--ink-3:#7c7b76;
  --s1:#1e7d5a;--s2:#d39b1a;--s3:#d39b1a;--s4:#8a8f98;--c4:#4a4f58;--c5:#a86b0c;--c6:#6fbf9a;--s5:#e87ba4;--s6:#008300;--s7:#6250d6;--s8:#e34948;
  --good:#1a7f4b;--good-bg:#e3f4ea;--warn:#a66300;--warn-bg:#fbf0db;--bad:#c23b3a;--bad-bg:#fbe5e4;--info-bg:#ece9e0;
  --accent:#1e7d5a;--grid:#ebe7de;--shadow:0 1px 2px rgba(40,32,20,.08),0 6px 18px rgba(40,32,20,.07);
  --rail:#262a31;--rail-ink:#ecebe7;--rail-ink-2:#b0ada5;--rail-on:rgba(255,255,255,.16);--page-bg:#fdfcf9;--top-bg:#262a31;--top-ink:#ffffff;--top-ink-2:#c9c5bc;--top-btn:#3b414b;--top-btn-h:#4a515c;--top-btn-ink:#ffffff;--top-chip:#343943;--menu-bg:#f2b33d;--menu-ink:#262a31;--conf-bg:#c8333a;--conf-ink:#ffffff;--top-line:transparent;
}
@media (prefers-color-scheme:dark){:root:where(:not([data-theme="light"])){
  color-scheme:dark;--bg:#121211;--surface:#1a1a19;--surface-2:#222220;--line:#2e2e2b;--ink:#f0efec;--ink-2:#c3c2b7;--ink-3:#8f8e86;
  --s1:#2f9d72;--s2:#b08018;--s3:#b08018;--s4:#80868f;--c4:#a3a8b1;--c5:#c98a2e;--c6:#7fcfa9;--s5:#d55181;--s6:#008300;--s7:#9085e9;--s8:#e66767;
  --good:#4cc38a;--good-bg:#173226;--warn:#e0a640;--warn-bg:#33270f;--bad:#f07777;--bad-bg:#3a1b1b;--info-bg:#22242c;--accent:#2f9d72;--grid:#272725;--shadow:none;--rail:#1b1d22;--top-bg:color-mix(in srgb,var(--bg) 88%,transparent);--top-ink:var(--ink);--top-ink-2:var(--ink-3);--top-btn:var(--surface);--top-btn-h:var(--surface-2);--top-btn-ink:var(--ink);--top-chip:var(--surface);--menu-bg:var(--rail);--menu-ink:#fff;--conf-bg:var(--bad-bg);--conf-ink:var(--bad);--top-line:var(--line);--rail-ink:#d7e0f2;--rail-ink-2:#8c9cbd;--rail-on:rgba(255,255,255,.12);--page-bg:var(--surface)}}
:root[data-theme="dark"]{
  color-scheme:dark;--bg:#121211;--surface:#1a1a19;--surface-2:#222220;--line:#2e2e2b;--ink:#f0efec;--ink-2:#c3c2b7;--ink-3:#8f8e86;
  --s1:#2f9d72;--s2:#b08018;--s3:#b08018;--s4:#80868f;--c4:#a3a8b1;--c5:#c98a2e;--c6:#7fcfa9;--s5:#d55181;--s6:#008300;--s7:#9085e9;--s8:#e66767;
  --good:#4cc38a;--good-bg:#173226;--warn:#e0a640;--warn-bg:#33270f;--bad:#f07777;--bad-bg:#3a1b1b;--info-bg:#22242c;--accent:#2f9d72;--grid:#272725;--shadow:none;--rail:#1b1d22;--top-bg:color-mix(in srgb,var(--bg) 88%,transparent);--top-ink:var(--ink);--top-ink-2:var(--ink-3);--top-btn:var(--surface);--top-btn-h:var(--surface-2);--top-btn-ink:var(--ink);--top-chip:var(--surface);--menu-bg:var(--rail);--menu-ink:#fff;--conf-bg:var(--bad-bg);--conf-ink:var(--bad);--top-line:var(--line);--rail-ink:#d7e0f2;--rail-ink-2:#8c9cbd;--rail-on:rgba(255,255,255,.12);--page-bg:var(--surface)}
*{box-sizing:border-box}
html{scroll-behavior:smooth}
body{margin:0;background:var(--bg);color:var(--ink);font:18px/1.45 "Segoe UI",system-ui,-apple-system,Arial,sans-serif;-webkit-print-color-adjust:exact;print-color-adjust:exact}
.sample{background:var(--warn-bg);color:var(--warn);text-align:center;font-size:12.5px;padding:7px 16px;border-bottom:1px solid var(--line)}
/* rail */
.rail{position:fixed;inset:0 auto 0 0;width:280px;background:var(--rail);color:var(--rail-ink);display:flex;flex-direction:column;z-index:40;transform:translateX(-100%);transition:transform .22s ease;box-shadow:none}
body.menu .rail{transform:none;box-shadow:8px 0 32px rgba(20,18,15,.35)}
.scrim{position:fixed;inset:0;background:rgba(20,18,15,.35);z-index:35;opacity:0;pointer-events:none;transition:opacity .2s}
body.menu .scrim{opacity:1;pointer-events:auto}
.rail .x{position:absolute;top:16px;right:14px;background:rgba(255,255,255,.12);border:0;color:#fff;width:30px;height:30px;border-radius:8px;font-size:16px;cursor:pointer}
.menu-btn{display:flex;align-items:center;gap:8px;padding:8px 13px;font-size:15px;font-weight:700;background:var(--menu-bg);color:var(--menu-ink);border:0;border-radius:10px;cursor:pointer}
.menu-btn .bars{font-size:18px;line-height:1}
.menu-btn small{font-weight:500;opacity:.85;font-size:12.5px}
.brand{padding:22px 22px 18px;border-bottom:1px solid rgba(255,255,255,.12)}
.brand .logo{display:flex;align-items:center;gap:10px;font-weight:700;font-size:15px}
.brand .dot{width:28px;height:28px;border-radius:8px;background:#f2b33d;color:#262a31!important;display:grid;place-items:center;color:#fff;font-size:13px;font-weight:800}
.brand small{display:block;color:var(--rail-ink-2);font-size:12px;margin-top:6px}
.rail nav{padding:12px 10px;overflow:auto;flex:1}
.rail a{display:flex;align-items:center;gap:11px;padding:9px 12px;border-radius:9px;color:var(--rail-ink);text-decoration:none;font-size:13.5px}
.rail a:hover{background:var(--rail-on);color:#fff}
.rail a.on{background:var(--rail-on);color:#fff;font-weight:600;box-shadow:inset 3px 0 0 #f2b33d}
.rail a .ix{font-variant-numeric:tabular-nums;color:var(--rail-ink-2);width:18px;font-size:12px}
.rail a.on .ix{color:#f2b33d}
.rail .foot{padding:14px 22px;border-top:1px solid rgba(255,255,255,.12);font-size:11.5px;color:var(--rail-ink-2)}
/* main */
main{margin-left:0;padding:0 32px 40px}
.top{position:sticky;top:0;z-index:10;background:var(--top-bg);backdrop-filter:blur(8px);display:flex;align-items:center;gap:12px;padding:14px 32px;margin:0 -32px 8px;border-bottom:1px solid var(--top-line);color:var(--top-ink)}
.top h1{font-size:23px;margin:0;flex:1;letter-spacing:-.2px}
.top h1 span{color:var(--top-ink-2);font-weight:400}
.chip{border:1px solid var(--top-line);background:var(--top-chip);border-radius:999px;padding:5px 12px;font-size:16px;color:var(--top-ink-2);white-space:nowrap}
.chip b{color:var(--top-ink)}
.btn{border:1px solid var(--line);background:var(--surface);color:var(--ink);border-radius:8px;padding:7px 12px;font:inherit;font-size:17px;cursor:pointer}
.btn:hover{background:var(--surface-2)}
.tag-conf{font-size:14px;letter-spacing:1.2px;font-weight:700;color:var(--bad);background:var(--bad-bg);padding:5px 9px;border-radius:6px}
section{scroll-margin-top:76px;padding-top:28px}
.sh{display:flex;align-items:baseline;gap:12px;margin-bottom:14px}
.sh .no{font-size:16px;font-weight:700;color:var(--accent);background:var(--info-bg);border-radius:6px;padding:2px 8px;font-variant-numeric:tabular-nums}
.sh h2{margin:0;font-size:26px;letter-spacing:-.3px}
.sh p{margin:0;color:var(--ink-3);font-size:17px}
.grid{display:grid;gap:16px}
.g6{grid-template-columns:repeat(6,1fr)}.g5{grid-template-columns:repeat(5,1fr)}.g4{grid-template-columns:repeat(4,1fr)}.g3{grid-template-columns:repeat(3,1fr)}.g2{grid-template-columns:1fr 1fr}.g12{grid-template-columns:1.6fr 1fr}.g21{grid-template-columns:1fr 1.6fr}
.card{background:var(--surface);border:1px solid var(--line);border-radius:14px;box-shadow:var(--shadow);padding:16px 18px;min-width:0}
.card h3{margin:0 0 2px;font-size:18px}
.card .sub{color:var(--ink-3);font-size:16px;margin-bottom:10px}
/* stat tiles */
.stat .lab{color:var(--ink-2);font-size:16px}
.stat .val{font-size:36px;font-weight:650;letter-spacing:-.6px;font-variant-numeric:tabular-nums;margin:4px 0 2px}
.stat .meta{font-size:16px;color:var(--ink-3);display:flex;gap:8px;align-items:center;flex-wrap:wrap}
.pill{display:inline-flex;align-items:center;gap:4px;border-radius:999px;padding:1px 8px;font-size:15px;font-weight:600}
.ok{color:var(--good);background:var(--good-bg)}.wn{color:var(--warn);background:var(--warn-bg)}.bd{color:var(--bad);background:var(--bad-bg)}.nt{color:var(--ink-2);background:var(--surface-2)}
.bar{height:6px;border-radius:3px;background:var(--surface-2);overflow:hidden;margin-top:8px}.bar i{display:block;height:100%;border-radius:3px;background:var(--accent)}
/* charts */
.chart{position:relative}
.chart svg{display:block;max-width:100%;overflow:visible}
.chart .ax{fill:var(--ink-3);font-size:14px}
.chart .gl{stroke:var(--grid)}
.chart .base{stroke:var(--line)}
.chart .lbl{fill:var(--ink-2);font-size:14px;font-variant-numeric:tabular-nums}
.chart .hit{fill:transparent}
.legend{display:flex;flex-wrap:wrap;gap:4px 14px;font-size:16px;color:var(--ink-2);margin-top:6px}
.legend i{display:inline-block;width:10px;height:10px;border-radius:3px;margin-right:6px;vertical-align:-1px}
.tip{position:absolute;pointer-events:none;background:var(--ink);color:var(--surface);font-size:16px;padding:6px 9px;border-radius:7px;white-space:nowrap;display:none;z-index:5;box-shadow:0 4px 14px rgba(0,0,0,.2)}
.tip b{font-variant-numeric:tabular-nums}
/* tables */
table{border-collapse:collapse;width:100%;font-size:17px}
th{text-align:left;font-weight:600;color:var(--ink-3);font-size:15px;letter-spacing:.3px;text-transform:uppercase;padding:8px 10px;border-bottom:1px solid var(--line)}
td{padding:9px 10px;border-bottom:1px solid var(--line);vertical-align:top}
td.n,th.n{text-align:right;font-variant-numeric:tabular-nums}
tr.total td{font-weight:650;border-bottom:0}
.inbar{position:relative}
.inbar span{position:absolute;inset:6px auto 6px 0;background:var(--info-bg);border-radius:4px;z-index:0}
.inbar b{position:relative;font-weight:500}
/* cover */
.cover{display:grid;grid-template-columns:1.4fr 1fr;gap:16px}
.cover .hero{padding:34px 34px 30px;background:linear-gradient(135deg,var(--info-bg),var(--surface) 70%)}
.cover .eyebrow{font-size:15px;letter-spacing:1.6px;text-transform:uppercase;color:var(--accent);font-weight:700}
.cover h1{font-size:44px;letter-spacing:-.8px;margin:12px 0 6px}
.cover .tg{color:var(--ink-2);margin-bottom:26px}
.kv{display:grid;grid-template-columns:max-content 1fr;gap:8px 18px;font-size:18px}
.kv dt{color:var(--ink-3)}.kv dd{margin:0;font-weight:600}
.notes .n{display:flex;gap:12px;padding:12px 0;border-bottom:1px solid var(--line)}
.notes .n:last-child{border-bottom:0}
.notes .ic{flex:none;width:28px;height:28px;border-radius:8px;display:grid;place-items:center;font-size:18px}
.notes b{display:block;font-size:18px}
.notes span{color:var(--ink-2);font-size:17px}
.reco{counter-reset:r}
.reco li{list-style:none;display:flex;gap:12px;padding:11px 0;border-bottom:1px solid var(--line)}
.reco li::before{counter-increment:r;content:counter(r,decimal-leading-zero);flex:none;width:30px;height:30px;border-radius:8px;background:var(--info-bg);color:var(--accent);font-weight:700;display:grid;place-items:center;font-size:16px}
.reco{padding:0;margin:0}
.enh{display:grid;grid-template-columns:repeat(5,1fr);gap:12px}
.enh .card{padding:14px}
.enh .id{font-size:15px;font-weight:700;color:var(--ink-3)}
.enh h4{margin:6px 0 6px;font-size:18px}
.enh p{margin:8px 0 0;color:var(--ink-2);font-size:16px}
.foot{margin-top:36px;color:var(--ink-3);font-size:16px;display:flex;justify-content:space-between;border-top:1px solid var(--line);padding-top:14px}
/* pages */
section{--sec:var(--s1)}
section:nth-of-type(3n+2){--sec:#262a31}section:nth-of-type(3n+3){--sec:#13805a}section:nth-of-type(3n+4){--sec:#a86b0c}
.page{background:var(--page-bg);border:1px solid var(--line);border-radius:18px;box-shadow:var(--shadow);overflow:hidden;margin-top:26px}
.ph{display:flex;align-items:center;gap:14px;padding:16px 24px;color:#fff;background:var(--sec)}
.ph .no{background:rgba(255,255,255,.2);border-radius:8px;padding:4px 10px;font-weight:700;font-size:17px;font-variant-numeric:tabular-nums}
.ph h2{margin:0;font-size:27px;letter-spacing:-.3px}
.ph p{margin:0 0 0 auto;font-size:17px;opacity:.85}
.pbd{padding:20px 24px 22px}
.pfoot{display:flex;justify-content:space-between;padding:10px 24px;border-top:1px solid var(--line);font-size:15px;color:var(--ink-3);background:var(--surface-2)}
.pfoot b{color:var(--sec)}
section{padding-top:0}
.pbd .card.stat{border-top:3px solid var(--sec);background:linear-gradient(180deg,color-mix(in srgb,var(--sec) 7%,var(--surface)),var(--surface) 70%)}
.pbd .card h3{color:color-mix(in srgb,var(--sec) 70%,var(--ink))}
:root[data-theme="dark"] .pbd .card h3{color:var(--ink)}
@media (prefers-color-scheme:dark){:root:not([data-theme="light"]) .pbd .card h3{color:var(--ink)}}
/* top bar buttons */
.top .btn{background:var(--top-btn);color:var(--top-btn-ink);border-color:var(--top-line)}.top .btn:hover{background:var(--top-btn-h)}
.top .tag-conf{background:var(--conf-bg);color:var(--conf-ink)}
/* zoom */
.zoom{display:flex;align-items:center;gap:4px}
.zoom .btn{padding:6px 10px;min-width:34px}
.zoom span{font-size:16px;color:var(--top-ink-2);min-width:46px;text-align:center;font-variant-numeric:tabular-nums}
/* big cover */
#s1{height:var(--coverH,calc(100vh - 120px));min-height:520px;display:flex;flex-direction:column}
#s1 .cover{flex:1;min-height:0;grid-template-columns:1.55fr 1fr;gap:20px}
#s1 .hero{display:flex;flex-direction:column;justify-content:center;padding:4vh 4.5vw;overflow:hidden;background:#262a31;color:#fff;border:0;border-radius:22px}
#s1 .eyebrow{color:#f2b33d;font-size:clamp(13px,min(1.2vw,2vh),20px);letter-spacing:3px}
#s1 h1{font-size:clamp(38px,min(5.2vw,9vh),96px);line-height:1.02;margin:2.2vh 0 1.4vh;letter-spacing:-1.5px}
#s1 .tg{color:#ecebe7;font-size:clamp(15px,min(1.5vw,2.5vh),26px);margin-bottom:4vh}
#s1 .kv{font-size:clamp(15px,min(1.4vw,2.4vh),26px);gap:1.4vh 28px}
#s1 .kv dt{color:#b8b4aa}#s1 .kv dd{color:#fff}
#s1 .stat{display:flex;flex-direction:column;justify-content:center;padding:2vh 2vw;min-height:0;border-radius:22px;border-top:0;border-left:6px solid var(--s3)}
#s1 .stat:nth-child(1){border-left-color:#262a31}#s1 .stat:nth-child(2){border-left-color:var(--s1)}#s1 .stat:nth-child(3){border-left-color:var(--s3)}
#s1 .stat .lab{font-size:clamp(14px,min(1.25vw,2.1vh),22px)}
#s1 .stat .val{font-size:clamp(32px,min(4.2vw,7vh),76px);letter-spacing:-1.5px}
#s1 .stat .meta{font-size:clamp(12px,min(1vw,1.8vh),18px)}
#s1 .conf-big{margin-top:4vh;display:inline-block;align-self:flex-start;border:2px solid #f2b33d;color:#f2b33d;border-radius:8px;padding:6px 14px;letter-spacing:4px;font-weight:700;font-size:clamp(12px,1vw,16px)}
/* presentation */
body.present .rail,body.present .scrim,body.present .top,body.present .sample{display:none}
body.present main{margin:0;padding:24px 40px}
body.present section{display:none;min-height:calc(100vh - 48px)}
body.present section.cur{display:block}
.hud{display:none}
body.present .hud{display:flex;position:fixed;right:18px;bottom:18px;gap:6px;align-items:center;background:var(--surface);border:1px solid var(--line);border-radius:999px;padding:5px 6px 5px 14px;font-size:12px;box-shadow:var(--shadow);z-index:30}
@media (max-width:1280px){.g6{grid-template-columns:repeat(3,1fr)}.g5{grid-template-columns:repeat(3,1fr)}.enh{grid-template-columns:repeat(3,1fr)}}
@media (max-width:1600px){.top .chip{display:none}}
@media (max-width:980px){.rail{display:none}main{margin:0;padding:0 16px 30px}.g4,.g3,.g2,.g12,.g21,.cover{grid-template-columns:1fr}.g6,.g5,.enh{grid-template-columns:1fr 1fr}.top .chip{display:none}}
@media print{@page{size:A4 landscape;margin:10mm}.rail,.top,.sample,.hud{display:none!important}main{margin:0;padding:0}section{break-after:page;padding-top:0}.card{box-shadow:none}}
.g1{grid-template-columns:1fr}
.miss{border:1.5px dashed var(--bad);background:var(--bad-bg);color:var(--bad);border-radius:10px;padding:14px;font-weight:600;margin-top:8px}
.val.miss-v{color:var(--ink-3)}
.bimg img{display:block;max-width:100%;height:auto;margin:8px auto 0;border-radius:8px}
.note{font-size:14px;color:var(--ink-3);margin-top:8px}
.xerr{outline:2px dashed var(--bad);outline-offset:1px;border-radius:2px;cursor:help}
td .pill{white-space:nowrap}

'@

$DesignBJs = @'

(function(){
'use strict';
var css=function(v){return getComputedStyle(document.documentElement).getPropertyValue(v).trim();};
var nf=function(v,d){return v===null||v===undefined?'—':Number(v).toLocaleString('en-US',{minimumFractionDigits:d||0,maximumFractionDigits:d||0});};
function esc(t){return String(t).replace(/&/g,'&amp;').replace(/</g,'&lt;');}
function nice(max){var p=Math.pow(10,Math.floor(Math.log10(max||1))),f=max/p;return (f<=1?1:f<=2?2:f<=2.5?2.5:f<=5?5:10)*p;}
function legend(items){return '<div class="legend">'+items.map(function(i){return '<span><i style="background:var('+i[1]+')"></i>'+esc(i[0])+'</span>';}).join('')+'</div>';}
function tipOn(box){var t=document.createElement('div');t.className='tip';box.appendChild(t);if(box._tipBound){return;}box._tipBound=1;
  box.addEventListener('mousemove',function(e){var t=box.querySelector('.tip');var el=e.target.closest('[data-tip]');if(!el){t.style.display='none';return;}
    var r=box.getBoundingClientRect();t.innerHTML=el.getAttribute('data-tip');t.style.display='block';
    var x=e.clientX-r.left+12,y=e.clientY-r.top+12;if(x+t.offsetWidth>r.width)x=e.clientX-r.left-t.offsetWidth-12;t.style.left=x+'px';t.style.top=y+'px';});
  box.addEventListener('mouseleave',function(){var t=box.querySelector('.tip');if(t)t.style.display='none';});}
// vertical bars (optionally stacked), single axis
function columns(el,cats,series,o){o=o||{};var W=Math.max(300,el.clientWidth),H=o.h||320,m={l:62,r:10,t:20,b:o.rot?86:38};
  var n=cats.length,cw=(W-m.l-m.r)/n,stack=o.stack;
  var tot=cats.map(function(_,i){return stack?series.reduce(function(a,s){return a+(s.v[i]||0);},0):Math.max.apply(null,series.map(function(s){return s.v[i]||0;}));});
  var max=nice(Math.max.apply(null,tot)*1.08),Y=function(v){return m.t+(H-m.t-m.b)*(1-v/max);},g='';
  for(var k=0;k<=4;k++){var v=max*k/4,y=Y(v);g+='<line class="gl" x1="'+m.l+'" x2="'+(W-m.r)+'" y1="'+y+'" y2="'+y+'"/><text class="ax" x="'+(m.l-8)+'" y="'+(y+5)+'" text-anchor="end">'+(o.pct?nf(v,0)+'%':nf(v))+'</text>';}
  var bw=stack?cw*0.56:cw*0.72/series.length;
  cats.forEach(function(c,i){var base=0,x0=m.l+cw*i+(cw-(stack?bw:bw*series.length))/2;
    series.forEach(function(s,j){var v=s.v[i];if(v===null||v===undefined)return;var x=stack?x0:x0+j*bw,y1=Y(base+v),y0=Y(base);
      var h=Math.max(0,y0-y1-(stack&&j>0?2:0)),col=s.colors?s.colors[i]:s.c;
      g+='<rect x="'+(x+(stack?0:1))+'" y="'+y1+'" width="'+(bw-(stack?0:2))+'" height="'+h+'" rx="'+(j===series.length-1||!stack?3:0)+'" fill="var('+col+')"/>';
      if(stack)base+=v;});
    var lab=o.labels?o.labels[i]:null;if(lab)g+='<text class="lbl" x="'+(x0+(stack?bw:bw*series.length)/2)+'" y="'+(Y(tot[i])-6)+'" text-anchor="middle">'+lab+'</text>';
    var tip='<b>'+esc(c)+'</b><br>'+series.map(function(s){return esc(s.n)+': <b>'+(s.v[i]===null?'—':(o.pct?nf(s.v[i],2)+'%':nf(s.v[i])))+'</b>';}).join('<br>')+(o.tipx?'<br>'+o.tipx[i]:'');
    g+='<rect class="hit" x="'+(m.l+cw*i)+'" y="'+m.t+'" width="'+cw+'" height="'+(H-m.t-m.b)+'" data-tip="'+esc(tip).replace(/"/g,'&quot;')+'"/>';
    g+=o.rot?'<text class="ax" transform="translate('+(m.l+cw*i+cw/2+4)+','+(H-m.b+12)+') rotate(-35)" text-anchor="end">'+esc(c)+'</text>':'<text class="ax" x="'+(m.l+cw*i+cw/2)+'" y="'+(H-m.b+22)+'" text-anchor="middle">'+esc(c)+'</text>';});
  g+='<line class="base" x1="'+m.l+'" x2="'+(W-m.r)+'" y1="'+Y(0)+'" y2="'+Y(0)+'"/>';
  el.innerHTML='<svg width="'+W+'" height="'+H+'" viewBox="0 0 '+W+' '+H+'">'+g+'</svg>'+(series.length>1?legend(series.map(function(s){return [s.n,s.c];})):'');tipOn(el);}
// horizontal bars with value + share labels
function hbars(el,cats,vals,o){o=o||{};var W=Math.max(300,el.clientWidth),rh=o.rh||36,lw=Math.min(o.lw||270,Math.round(W*0.42)),H=cats.length*rh+8,max=Math.max.apply(null,vals)*(o.share?1.4:1.18),g='';
  var maxCh=Math.max(6,Math.floor((lw-16)/8.3)),short=function(c){c=String(c);return c.length>maxCh?c.slice(0,maxCh-1)+'…':c;};
  cats.forEach(function(c,i){var y=4+i*rh,w=(W-lw-10)*vals[i]/max,share=o.share?o.share[i]:null;
    g+='<text class="ax" x="'+(lw-10)+'" y="'+(y+rh/2+4)+'" text-anchor="end" style="font-size:16px;fill:var(--ink-2)">'+esc(short(c))+'</text>';
    g+='<rect x="'+lw+'" y="'+(y+7)+'" width="'+Math.max(2,w)+'" height="'+(rh-14)+'" rx="3" fill="var('+(o.c||'--s1')+')"/>';
    g+='<text class="lbl" x="'+(lw+w+6)+'" y="'+(y+rh/2+4)+'">'+nf(vals[i])+(share?' · '+share:'')+'</text>';
    g+='<rect class="hit" x="0" y="'+y+'" width="'+W+'" height="'+rh+'" data-tip="<b>'+esc(c)+'</b><br>'+nf(vals[i])+(share?' ('+share+')':'')+'"/>';});
  el.innerHTML='<svg width="'+W+'" height="'+H+'" viewBox="0 0 '+W+' '+H+'">'+g+'</svg>';tipOn(el);}
// donut with centre figure
function donut(el,parts,o){o=o||{};var S=220,c=S/2,R=92,r=64,tot=parts.reduce(function(a,p){return a+p[1];},0),a=-Math.PI/2,g='';
  parts.forEach(function(p){var f=p[1]/tot,a2=a+f*2*Math.PI-(f<1?0.012:0),big=a2-a>Math.PI?1:0,P=function(rr,an){return (c+rr*Math.cos(an)).toFixed(2)+','+(c+rr*Math.sin(an)).toFixed(2);};
    var d=f>=0.9999?'M'+P(R,0)+'A'+R+','+R+' 0 1 1 '+P(R,Math.PI)+'A'+R+','+R+' 0 1 1 '+P(R,0)+'M'+P(r,0)+'A'+r+','+r+' 0 1 0 '+P(r,Math.PI)+'A'+r+','+r+' 0 1 0 '+P(r,0)+'Z':
      'M'+P(R,a)+'A'+R+','+R+' 0 '+big+' 1 '+P(R,a2)+'L'+P(r,a2)+'A'+r+','+r+' 0 '+big+' 0 '+P(r,a)+'Z';
    g+='<path d="'+d+'" fill="var('+p[2]+')" fill-rule="evenodd" data-tip="<b>'+esc(p[0])+'</b><br>'+(o.pct?'':nf(p[1])+' · ')+(f*100).toFixed(2)+'%"/>';a=a2+(f<1?0.012:0);});
  g+='<text x="'+c+'" y="'+(c-2)+'" text-anchor="middle" style="font-size:26px;font-weight:650;fill:var(--ink)">'+(o.center||'')+'</text><text x="'+c+'" y="'+(c+18)+'" text-anchor="middle" class="ax">'+(o.sub||'')+'</text>';
  el.innerHTML='<svg viewBox="0 0 '+S+' '+S+'" style="max-width:320px;margin:0 auto">'+g+'</svg>'+legend(parts.map(function(p){return [p[0]+' · '+(o.pct?(p[1]*100).toFixed(2)+'%':nf(p[1])),p[2]];}));tipOn(el);}
// single line with markers
function line(el,cats,vals,o){o=o||{};var W=Math.max(300,el.clientWidth),H=300,m={l:62,r:12,t:20,b:38},max=nice(Math.max.apply(null,vals)*1.1),n=cats.length,g='';
  var X=function(i){return m.l+(W-m.l-m.r)*(i+.5)/n;},Y=function(v){return m.t+(H-m.t-m.b)*(1-v/max);};
  for(var k=0;k<=4;k++){var v=max*k/4;g+='<line class="gl" x1="'+m.l+'" x2="'+(W-m.r)+'" y1="'+Y(v)+'" y2="'+Y(v)+'"/><text class="ax" x="'+(m.l-8)+'" y="'+(Y(v)+4)+'" text-anchor="end">'+nf(v)+'</text>';}
  g+='<path d="M'+vals.map(function(v,i){return X(i)+','+Y(v);}).join('L')+'" fill="none" stroke="var('+(o.c||'--s1')+')" stroke-width="2" stroke-linejoin="round"/>';
  vals.forEach(function(v,i){g+='<circle cx="'+X(i)+'" cy="'+Y(v)+'" r="4" fill="var(--surface)" stroke="var('+(o.c||'--s1')+')" stroke-width="2"/><rect class="hit" x="'+(X(i)-(W-m.l-m.r)/n/2)+'" y="'+m.t+'" width="'+((W-m.l-m.r)/n)+'" height="'+(H-m.t-m.b)+'" data-tip="<b>'+esc(cats[i])+'</b><br>'+nf(v)+'"/><text class="ax" x="'+X(i)+'" y="'+(H-m.b+22)+'" text-anchor="middle">'+esc(cats[i])+'</text>';});
  el.innerHTML='<svg width="'+W+'" height="'+H+'" viewBox="0 0 '+W+' '+H+'">'+g+'</svg>';tipOn(el);}

function drawAll(){
[].slice.call(document.querySelectorAll('.chart[data-c]')).forEach(function(el){
  try{var d=JSON.parse(el.getAttribute('data-c'));
    if(d.type==='columns')columns(el,d.cats,d.series,d.o);
    else if(d.type==='hbars')hbars(el,d.cats,d.vals,d.o);
    else if(d.type==='donut')donut(el,d.parts,d.o);
    else if(d.type==='line')line(el,d.cats,d.vals,d.o);
  }catch(e){el.innerHTML='<div class="miss">Chart could not be drawn: '+esc(String(e))+'</div>';}
});
}
drawAll();var rz;window.addEventListener('resize',function(){clearTimeout(rz);rz=setTimeout(drawAll,150);});

(function(){var ERR=/#(DIV\/0!|N\/A|VALUE!|REF!|NAME\?|NUM!|NULL!|SPILL!|CALC!|FIELD!|BLOCKED!|CONNECT!|BUSY!|UNKNOWN!|GETTING_DATA)/g;
  var w=document.createTreeWalker(document.getElementById('content'),NodeFilter.SHOW_TEXT,null),hits=[];
  while(w.nextNode()){ERR.lastIndex=0;if(ERR.test(w.currentNode.nodeValue))hits.push(w.currentNode);}
  hits.forEach(function(t){var s=t.nodeValue,f=document.createDocumentFragment(),last=0,m;ERR.lastIndex=0;
    while((m=ERR.exec(s))){if(m.index>last)f.appendChild(document.createTextNode(s.slice(last,m.index)));var sp=document.createElement('span');sp.className='xerr';sp.title='Excel error value - shown exactly as in Wkly_Report';sp.textContent=m[0];f.appendChild(sp);last=m.index+m[0].length;}
    if(last<s.length)f.appendChild(document.createTextNode(s.slice(last)));t.parentNode.replaceChild(f,t);});})();
// pages: coloured header bar, body, footer
var all=[].slice.call(document.querySelectorAll('section')),N=all.length;
all.forEach(function(sec,i){
  var sh=sec.querySelector('.sh'),body=document.createElement('div');body.className='pbd';
  if(sh){var no=sh.querySelector('.no'),h=sh.querySelector('h2'),pp=sh.querySelector('p');var head=document.createElement('div');head.className='ph';
    head.innerHTML='<span class="no">'+no.textContent+'</span><h2>'+h.innerHTML+'</h2>'+(pp?'<p>'+pp.innerHTML+'</p>':'');sh.remove();
    while(sec.firstChild)body.appendChild(sec.firstChild);sec.classList.add('page');sec.appendChild(head);sec.appendChild(body);
    var f=document.createElement('div');f.className='pfoot';f.innerHTML='<span>'+document.title+' · '+(document.body.getAttribute('data-conf')||'')+'</span><span>Page <b>'+(i+1)+'</b> of '+N+'</span>';sec.appendChild(f);}
});
// cover fills the window below the top bar
function fitCover(){var t=document.querySelector('.top'),c=document.getElementById('s1');if(!t||!c)return;
  document.documentElement.style.setProperty('--coverH',Math.max(520,(window.innerHeight-t.getBoundingClientRect().bottom-18)/zoom)+'px');}
window.addEventListener('resize',fitCover);
// zoom
var zoom=1;try{zoom=parseFloat(localStorage.getItem('itsmB.zoom'))||1;}catch(e){}
function setZoom(z){zoom=Math.max(0.6,Math.min(2,Math.round(z*100)/100));document.getElementById('content').style.zoom=zoom;
  document.getElementById('zVal').textContent=Math.round(zoom*100)+'%';try{localStorage.setItem('itsmB.zoom',String(zoom));}catch(e){}
  fitCover();clearTimeout(rz);rz=setTimeout(drawAll,60);}
document.getElementById('zIn').onclick=function(){setZoom(zoom+0.1);};document.getElementById('zOut').onclick=function(){setZoom(zoom-0.1);};document.getElementById('zFit').onclick=function(){setZoom(1);};
document.addEventListener('keydown',function(e){if(document.body.classList.contains('present')||e.ctrlKey||e.metaKey)return;if(e.key==='+'||e.key==='=')setZoom(zoom+0.1);if(e.key==='-')setZoom(zoom-0.1);});
setZoom(zoom);
// navigation
var secs=[].slice.call(document.querySelectorAll('section')),nav=document.getElementById('nav');
nav.innerHTML=secs.map(function(s,i){return '<a href="#'+s.id+'"><span class="ix">'+String(i+1).padStart(2,'0')+'</span>'+s.dataset.t+'</a>';}).join('');
var links=[].slice.call(nav.querySelectorAll('a'));
// pages menu (hidden until opened)
function menu(o){document.body.classList.toggle('menu',o);}
document.getElementById('mOpen').onclick=function(){menu(true);};document.getElementById('mClose').onclick=function(){menu(false);};document.getElementById('scrim').onclick=function(){menu(false);};
links.forEach(function(l){l.addEventListener('click',function(){menu(false);});});
document.addEventListener('keydown',function(e){if(e.key==='Escape')menu(false);});
function track(){var k=0,y=window.innerHeight*0.4;secs.forEach(function(s,i){if(s.getBoundingClientRect().top<=y)k=i;});
  links.forEach(function(l,i){l.classList.toggle('on',i===k);});document.getElementById('mPos').textContent=(k+1)+' / '+secs.length;}
window.addEventListener('scroll',track,{passive:true});window.addEventListener('resize',track);track();
// theme
document.getElementById('theme').onclick=function(){var r=document.documentElement,d=r.dataset.theme==='dark'||(!r.dataset.theme&&matchMedia('(prefers-color-scheme:dark)').matches);r.dataset.theme=d?'light':'dark';};
// present
var cur=0;function show(i){cur=Math.max(0,Math.min(secs.length-1,i));secs.forEach(function(s,k){s.classList.toggle('cur',k===cur);});document.getElementById('hpos').textContent=(cur+1)+' / '+secs.length;}
function start(){document.body.classList.add('present');show(Math.max(0,links.findIndex(function(l){return l.classList.contains('on');})));if(document.documentElement.requestFullscreen)document.documentElement.requestFullscreen().catch(function(){});}
function stop(){document.body.classList.remove('present');if(document.fullscreenElement)document.exitFullscreen();secs[cur].scrollIntoView();}
document.getElementById('present').onclick=start;document.getElementById('hp').onclick=function(){show(cur-1);};document.getElementById('hn').onclick=function(){show(cur+1);};document.getElementById('hx').onclick=stop;
document.addEventListener('keydown',function(e){if(!document.body.classList.contains('present'))return;if(['ArrowRight','PageDown',' '].indexOf(e.key)>=0){e.preventDefault();show(cur+1);}if(['ArrowLeft','PageUp'].indexOf(e.key)>=0)show(cur-1);if(e.key==='Escape')stop();});
document.addEventListener('fullscreenchange',function(){if(!document.fullscreenElement&&document.body.classList.contains('present'))stop();});
})();

'@

# ======================================================================================
#  Main
# ======================================================================================

function Invoke-Main {
    Write-Log "Export-WklyReportDashboardB (Design B) build $ScriptBuild" 'OK'

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
    Write-Log 'Arranging pages in the Design B layout:'
    $index = New-DesignBSite -Pages $result.Pages -OutDir $outDir -Meta $meta -WorkDir $script:WorkDir
    # validation of the single output file
    $validLevel = 'OK'
    $written = Get-Content -LiteralPath $index -Raw -Encoding UTF8
    $sections = [regex]::Matches($written, '<section id="s\d+"').Count
    $externalRefs = [regex]::Matches($written, '(?:src|href)="(?!data:|#)[^"]+"').Count
    if ($sections -ne $result.Pages.Count -or $externalRefs -ne 0) { $validLevel = 'WARN' }
    $nativeCharts = [regex]::Matches($written, 'class="chart" data-c=').Count
    Write-Log ("Validation: {0} page(s) in the file (expected {1}), {2} chart(s) redrawn from data, {3} external reference(s) (expected 0)." -f $sections, $result.Pages.Count, $nativeCharts, $externalRefs) $validLevel
    Write-Log "Dashboard written: $index" 'OK'


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
