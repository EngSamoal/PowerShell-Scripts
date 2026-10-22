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
      * Every page of Wkly_Report is detected on every run (pages may be added or removed weekly).
      * Page 1 becomes the full-screen cover (title, the text around it, its "Label: value" details and
        any KPI cards on it). Every other page gets a coloured header bar (title + the text in Excel's
        header band) and its content as cards, kept in Excel's arrangement:
          - a card whose largest text is a figure  -> KPI tile (label, figure, details)
          - a chart                                -> redrawn natively from the chart's own data, with
                                                      Excel's axis scale and number formats
          - a block of cells                        -> table
          - other text                              -> text card
      * Values are shown exactly as Excel displays them (cell .Text / text-box text); nothing is
        recalculated. Excel errors such as #DIV/0! are shown as-is and outlined. Blank chart values
        stay gaps, never 0.
      * Colours follow the dashboard palette: Excel's chart colours are mapped by hue (reds stay red,
        greens -> green, yellows/oranges -> gold, greys -> grey, others -> the next free palette colour).
        Status words (Met, Breached, Open, High, Low, ...) get a coloured pill; the word is unchanged.
      * A chart type that is not redrawn natively is shown as Excel's own picture of that chart.

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
    Smallest font size (pt) that counts as a page title. Default 20.

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
    [double]$TitleMinFontSize = 20,
    [double]$MinPageHeightPt = 250,
    [string]$DashboardTitle = 'IT Service Management',
    [string]$Classification = 'CONFIDENTIAL',
    [switch]$OpenWhenDone
)

$ScriptBuild = '2026.10.22-B1'
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
    if ($Type -in 4, 63, 64, 65, 66, 67, -4101) { return 'line' }
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
    return $spec
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
    # All pages share one horizontal crop so they have the same scale and line up with each other.
    if ($pages.Count) {
        $gL = ($pages | Measure-Object -Property OffsetX -Minimum).Minimum
        $gR = ($pages | ForEach-Object { $_.OffsetX + $_.Width } | Measure-Object -Maximum).Maximum
        foreach ($p in $pages) { $p.OffsetX = $gL; $p.Width = $gR - $gL }
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
                if ($it.Kind -eq 'shape' -and $it.Runs) { $o['runs'] = @($it.Runs | ForEach-Object { [ordered]@{ text = $_.Text; size = $_.Size; bold = $_.Bold; color = $_.Color; font = $_.Font } }) }
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
#  Design B: each page's content becomes cards (KPI tiles, charts, tables, text) - pure logic
# ======================================================================================

function Get-TextPieces {
    # Paragraphs of a cell or text box with their largest font size and boldness, in reading order.
    param($It)
    $out = New-Object System.Collections.Generic.List[object]
    if ($It.Kind -eq 'cell') {
        $t = ([string]$It.Text).Trim()
        if ($t) { $out.Add([pscustomobject]@{ Text = $t; Size = [double]$It.FontSize; Bold = [bool]$It.Bold; T = [double]$It.T; L = [double]$It.L; Color = $It.Color }) }
        return , $out
    }
    if ($It.Kind -ne 'shape' -or -not $It.Runs) { return , $out }
    $buf = ''; $size = 0.0; $bold = $false; $color = $null; $k = 0
    foreach ($r in $It.Runs) {
        $segs = ([string]$r.Text) -split "`r"
        for ($s = 0; $s -lt $segs.Count; $s++) {
            if ($s -gt 0) {
                $t = ($buf -replace "[`v`n]+", ' ').Trim()
                if ($t) { $out.Add([pscustomobject]@{ Text = $t; Size = $size; Bold = $bold; T = [double]$It.T + $k * 0.01; L = [double]$It.L; Color = $color }); $k++ }
                $buf = ''; $size = 0.0; $bold = $false; $color = $null
            }
            $seg = $segs[$s]
            if ($seg -match '\S') {
                if ([double]$r.Size -gt $size) { $size = [double]$r.Size }
                if ($r.Bold) { $bold = $true }
                if (-not $color) { $color = $r.Color }
            }
            $buf += $seg
        }
    }
    $t = ($buf -replace "[`v`n]+", ' ').Trim()
    if ($t) { $out.Add([pscustomobject]@{ Text = $t; Size = $size; Bold = $bold; T = [double]$It.T + $k * 0.01; L = [double]$It.L; Color = $color }) }
    return , $out
}

function Test-ValueLike {
    # Short text that reads as a figure: 16,818 / 99.60% / 2:23:48 / 16.96 min / #DIV/0! / -
    param([string]$Text)
    $t = $Text.Trim()
    if ($t.Length -eq 0 -or $t.Length -gt 32) { return $false }
    if ($t -match $ErrRegex) { return $true }
    if ($t -match '^[-–—]$') { return $true }
    if ($t -notmatch '\d') { return $false }
    return ([regex]::Matches($t, '\p{L}').Count -le 12)
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
    # Text exactly as Excel shows it; status words wrapped in a pill.
    param([string]$Text)
    $cls = Get-StatusClass $Text
    if ($cls) { return ('<span class="pill {0}">{1}</span>' -f $cls, (ConvertTo-HtmlText $Text.Trim())) }
    return (ConvertTo-HtmlText $Text)
}

function New-Block {
    param([string]$Type, [double]$L, [double]$T, [double]$W, [double]$H)
    return [pscustomobject]@{ Type = $Type; L = $L; T = $T; W = $W; H = $H; Title = $null; Sub = $null
        Spec = $null; Src = $null; Alt = $null; Label = $null; Value = $null; Meta = $null; Pieces = $null; Cells = $null }
}

function Get-RectOf {
    param($List)
    $l = 1e9; $t = 1e9; $r = -1e9; $b = -1e9
    foreach ($x in $List) {
        if ($x.L -lt $l) { $l = $x.L }; if ($x.T -lt $t) { $t = $x.T }
        if ($x.L + $x.W -gt $r) { $r = $x.L + $x.W }; if ($x.T + $x.H -gt $b) { $b = $x.T + $x.H }
    }
    return [pscustomobject]@{ L = $l; T = $t; W = $r - $l; H = $b - $t }
}

function ConvertTo-KpiBlock {
    # A card whose largest text is a figure becomes a KPI tile: label above, figure, details below.
    param($Pieces, $Rect)
    if ($Pieces.Count -lt 2 -or $Pieces.Count -gt 7) { return $null }
    $best = $null
    foreach ($p in $Pieces) { if (Test-ValueLike $p.Text) { if ($null -eq $best -or $p.Size -gt $best.Size) { $best = $p } } }
    if ($null -eq $best) { return $null }
    $others = 0.0
    foreach ($p in $Pieces) { if ($p -ne $best -and $p.Size -gt $others) { $others = $p.Size } }
    if ($best.Size -lt 1.2 * $others) { return $null }
    $b = New-Block 'kpi' $Rect.L $Rect.T $Rect.W $Rect.H
    $before = New-Object System.Collections.Generic.List[string]
    $after = New-Object System.Collections.Generic.List[string]
    $seen = $false
    foreach ($p in $Pieces) {
        if ($p -eq $best) { $seen = $true; continue }
        if ($seen) { $after.Add($p.Text) } else { $before.Add($p.Text) }
    }
    if ($before.Count -eq 0 -and $after.Count -gt 0) { $before.Add($after[0]); $after.RemoveAt(0) }
    $b.Label = ($before -join ' ')
    $b.Value = $best.Text
    $b.Meta = $after
    return $b
}

function Test-TableShape {
    param($Cells)
    if ($Cells.Count -lt 4) { return $false }
    $rows = @{}; $cols = @{}
    foreach ($c in $Cells) { $rows[[string][math]::Round($c.T)] = 1; $cols[[string][math]::Round($c.L)] = 1 }
    return ($rows.Count -ge 2 -and $cols.Count -ge 2)
}

function ConvertTo-TableBlock {
    param($Cells, $Rect)
    $b = New-Block 'table' $Rect.L $Rect.T $Rect.W $Rect.H
    $b.Cells = $Cells
    return $b
}

function Get-LeafItems {
    # Everything inside a card (nested cards included), each object once.
    param($Node, $Out)
    $it = $Node.Item
    if ($it.Kind -eq 'shape' -and $it.Runs) { $Out.Add($it) }
    elseif ($it.Kind -in 'cell', 'image', 'missing') { $Out.Add($it) }
    foreach ($c in $Node.Children) { Get-LeafItems -Node $c -Out $Out }
}

function Add-CardBlocks {
    # One Excel card (filled/bordered shape or coloured cell area) with no cards inside it.
    param($Node, $Blocks)
    $leaves = New-Object System.Collections.Generic.List[object]
    Get-LeafItems -Node $Node -Out $leaves
    $charts = New-Object System.Collections.Generic.List[object]
    $pics = New-Object System.Collections.Generic.List[object]
    $cells = New-Object System.Collections.Generic.List[object]
    $texts = New-Object System.Collections.Generic.List[object]
    foreach ($x in $leaves) {
        if ($x.Kind -eq 'image' -and $x.IsChart) { $charts.Add($x) }
        elseif ($x.Kind -in 'image', 'missing') { $pics.Add($x) }
        elseif ($x.Kind -eq 'cell') { $cells.Add($x) }
        else { $texts.Add($x) }
    }
    $pieces = New-Object System.Collections.Generic.List[object]
    foreach ($x in $texts) { foreach ($p in (Get-TextPieces $x)) { $pieces.Add($p) } }
    $rect = [pscustomobject]@{ L = $Node.L; T = $Node.T; W = $Node.W; H = $Node.H }

    if ($charts.Count -gt 0 -or $pics.Count -gt 0 -or (Test-TableShape $cells)) {
        $first = $true
        $sorted = @($pieces | Sort-Object T, L)
        $visuals = New-Object System.Collections.Generic.List[object]
        foreach ($x in $charts) { $visuals.Add($x) }
        foreach ($x in $pics) { $visuals.Add($x) }
        foreach ($x in ($visuals | Sort-Object T, L)) {
            $kind = 'chart'; if (-not $x.IsChart) { $kind = 'image' }
            $r = $rect; if (($charts.Count + $pics.Count) -gt 1) { $r = $x }
            $b = New-Block $kind $r.L $r.T $r.W $r.H
            $b.Spec = $x.Spec; $b.Src = $x.Src; $b.Alt = $x.Alt
            if ($x.Kind -eq 'missing') { $b.Type = 'missing'; $b.Alt = $x.Message }
            if ($first -and $sorted.Count) {
                $b.Title = $sorted[0].Text
                if ($sorted.Count -gt 1) { $b.Sub = (@($sorted | Select-Object -Skip 1 | ForEach-Object { $_.Text }) -join ' · ') }
                $first = $false
            }
            $Blocks.Add($b)
        }
        if (Test-TableShape $cells) {
            $tb = ConvertTo-TableBlock -Cells $cells -Rect (Get-RectOf $cells)
            if ($first -and $sorted.Count) { $tb.Title = $sorted[0].Text; if ($sorted.Count -gt 1) { $tb.Sub = (@($sorted | Select-Object -Skip 1 | ForEach-Object { $_.Text }) -join ' · ') } }
            if ($charts.Count -eq 0 -and $pics.Count -eq 0) { $tb.L = $rect.L; $tb.T = $rect.T; $tb.W = $rect.W; $tb.H = $rect.H }
            $Blocks.Add($tb)
        }
        return
    }
    foreach ($x in $cells) { foreach ($p in (Get-TextPieces $x)) { $pieces.Add($p) } }
    if ($pieces.Count -eq 0) { return }
    $ordered = New-Object System.Collections.Generic.List[object]
    foreach ($p in ($pieces | Sort-Object T, L)) { $ordered.Add($p) }
    $kpi = ConvertTo-KpiBlock -Pieces $ordered -Rect $rect
    if ($kpi) { $Blocks.Add($kpi); return }
    $tb = New-Block 'text' $rect.L $rect.T $rect.W $rect.H
    $tb.Pieces = $ordered
    $Blocks.Add($tb)
}

function Add-NodeBlocks {
    param($Nodes, $Blocks, $Loose, [double]$PageArea)
    foreach ($n in $Nodes) {
        $it = $n.Item
        if ($n.IsBox) {
            $hasInnerBox = $false
            foreach ($c in $n.Children) { if ($c.IsBox) { $hasInnerBox = $true; break } }
            if ($hasInnerBox -or (Get-NodeArea $n) -gt 0.45 * $PageArea) {
                # a panel that holds other cards: its own text is loose, its cards are handled one by one
                if ($it.Kind -eq 'shape' -and $it.Runs) { $Loose.Add($it) }
                Add-NodeBlocks -Nodes $n.Children -Blocks $Blocks -Loose $Loose -PageArea $PageArea
                continue
            }
            Add-CardBlocks -Node $n -Blocks $Blocks
            continue
        }
        if ($it.Kind -eq 'image' -and $it.IsChart) {
            $b = New-Block 'chart' $n.L $n.T $n.W $n.H; $b.Spec = $it.Spec; $b.Src = $it.Src; $b.Alt = $it.Alt; $Blocks.Add($b); continue
        }
        if ($it.Kind -eq 'image') { $b = New-Block 'image' $n.L $n.T $n.W $n.H; $b.Src = $it.Src; $b.Alt = $it.Alt; $Blocks.Add($b); continue }
        if ($it.Kind -eq 'missing') { $b = New-Block 'missing' $n.L $n.T $n.W $n.H; $b.Alt = $it.Message; $Blocks.Add($b); continue }
        if ($it.Kind -eq 'cell' -or ($it.Kind -eq 'shape' -and $it.Runs)) { $Loose.Add($it) }
        if ($n.Children.Count) { Add-NodeBlocks -Nodes $n.Children -Blocks $Blocks -Loose $Loose -PageArea $PageArea }
    }
}

function Add-LooseBlocks {
    # Text boxes and cells that are not inside a card: grouped by distance into tables, KPI tiles or text.
    param($Loose, $Blocks)
    $n = $Loose.Count
    if ($n -eq 0) { return }
    $grp = New-Object 'int[]' $n
    for ($i = 0; $i -lt $n; $i++) { $grp[$i] = $i }
    $find = { param($k) while ($grp[$k] -ne $k) { $k = $grp[$k] }; return $k }
    $gap = 10.0
    for ($i = 0; $i -lt $n; $i++) {
        for ($j = $i + 1; $j -lt $n; $j++) {
            $a = $Loose[$i]; $c = $Loose[$j]
            $ox = [math]::Min($a.L + $a.W, $c.L + $c.W) - [math]::Max($a.L, $c.L)
            $oy = [math]::Min($a.T + $a.H, $c.T + $c.H) - [math]::Max($a.T, $c.T)
            if ($ox -ge -$gap -and $oy -ge -$gap) {
                $ri = & $find $i; $rj = & $find $j
                if ($ri -ne $rj) { $grp[$rj] = $ri }
            }
        }
    }
    $groups = @{}
    for ($i = 0; $i -lt $n; $i++) {
        $r = & $find $i
        if (-not $groups.ContainsKey($r)) { $groups[$r] = New-Object System.Collections.Generic.List[object] }
        $groups[$r].Add($Loose[$i])
    }
    foreach ($key in $groups.Keys) {
        $members = $groups[$key]
        $rect = Get-RectOf $members
        $cells = New-Object System.Collections.Generic.List[object]
        $others = New-Object System.Collections.Generic.List[object]
        foreach ($m in $members) { if ($m.Kind -eq 'cell') { $cells.Add($m) } else { $others.Add($m) } }
        $pieces = New-Object System.Collections.Generic.List[object]
        if (Test-TableShape $cells) {
            $tb = ConvertTo-TableBlock -Cells $cells -Rect $rect
            foreach ($m in $others) { foreach ($p in (Get-TextPieces $m)) { $pieces.Add($p) } }
            $sorted = @($pieces | Sort-Object T, L)
            if ($sorted.Count) { $tb.Title = $sorted[0].Text; if ($sorted.Count -gt 1) { $tb.Sub = (@($sorted | Select-Object -Skip 1 | ForEach-Object { $_.Text }) -join ' · ') } }
            $Blocks.Add($tb)
            continue
        }
        foreach ($m in $members) { foreach ($p in (Get-TextPieces $m)) { $pieces.Add($p) } }
        if ($pieces.Count -eq 0) { continue }
        $ordered = New-Object System.Collections.Generic.List[object]
        foreach ($p in ($pieces | Sort-Object T, L)) { $ordered.Add($p) }
        $kpi = ConvertTo-KpiBlock -Pieces $ordered -Rect $rect
        if ($kpi) { $Blocks.Add($kpi); continue }
        $tx = New-Block 'text' $rect.L $rect.T $rect.W $rect.H
        $tx.Pieces = $ordered
        $Blocks.Add($tx)
    }
}

function Merge-CaptionBlocks {
    # A short text standing directly above a chart or table is that chart's / table's title.
    param($Blocks)
    $drop = New-Object System.Collections.Generic.List[object]
    foreach ($t in $Blocks) {
        if ($t.Type -ne 'text' -or $t.Pieces.Count -gt 2) { continue }
        if ($t.Pieces[0].Text.Length -gt 90) { continue }
        $target = $null; $bestGap = 1e9
        foreach ($b in $Blocks) {
            if ($b.Type -notin 'chart', 'table', 'image' -or $b.Title) { continue }
            $gapY = $b.T - ($t.T + $t.H)
            if ($gapY -lt -6 -or $gapY -gt 40) { continue }
            $ox = [math]::Min($t.L + $t.W, $b.L + $b.W) - [math]::Max($t.L, $b.L)
            if ($ox -lt 0.4 * [math]::Min($t.W, $b.W)) { continue }
            if ($gapY -lt $bestGap) { $bestGap = $gapY; $target = $b }
        }
        if ($null -eq $target) { continue }
        $target.Title = $t.Pieces[0].Text
        if ($t.Pieces.Count -gt 1) { $target.Sub = $t.Pieces[1].Text }
        $h = $target.T + $target.H - $t.T
        $l = [math]::Min($t.L, $target.L); $r = [math]::Max($t.L + $t.W, $target.L + $target.W)
        $target.T = $t.T; $target.H = $h; $target.L = $l; $target.W = $r - $l
        $drop.Add($t)
    }
    foreach ($d in $drop) { [void]$Blocks.Remove($d) }
}

function Get-PageHeader {
    # The page title and the text beside it in the page's header band; everything above
    # $Bottom belongs to the header and is shown in the coloured header bar.
    param($Page)
    $titleItem = $null
    $want = ([string]$Page.Title).Trim()
    foreach ($it in $Page.Items) {
        if ($it.Kind -ne 'cell' -and $it.Kind -ne 'shape') { continue }
        $first = ((Get-ItemText $it) -split "[`r`n`v]" | Where-Object { $_.Trim() } | Select-Object -First 1)
        if (-not $first) { continue }
        $first = $first.Trim()
        if ($first -eq $want -or ($want.EndsWith('...') -and $first.StartsWith($want.Substring(0, $want.Length - 3)))) {
            if ($null -eq $titleItem -or $it.T -lt $titleItem.T) { $titleItem = $it }
        }
    }
    $res = [pscustomobject]@{ Title = $want; Sub = $null; Bottom = [double]$Page.StartY - 1; TitleItem = $titleItem }
    if ($null -eq $titleItem) { return $res }
    $bottom = $titleItem.T + $titleItem.H
    $band = $null
    try { $band = Get-DarkBand -It $titleItem -Items $Page.Items -ContentWidth ([double]$Page.OffsetX + [double]$Page.Width) } catch { }
    if ($band) { $bottom = [math]::Max($bottom, $band.T + $band.H) }
    $res.Bottom = $bottom
    $subs = New-Object System.Collections.Generic.List[string]
    foreach ($it in ($Page.Items | Sort-Object T, L)) {
        if ($it -eq $titleItem) {
            # further paragraphs of the title box are the subtitle
            $k = 0
            foreach ($p in (Get-TextPieces $it)) { if ($k -gt 0) { $subs.Add($p.Text) }; $k++ }
            continue
        }
        if ($it.Kind -ne 'cell' -and -not ($it.Kind -eq 'shape' -and $it.Runs)) { continue }
        if ($it.T + $it.H -gt $bottom + 2) { continue }
        foreach ($p in (Get-TextPieces $it)) { $subs.Add($p.Text) }
    }
    if ($subs.Count) { $res.Sub = ($subs -join ' · ') }
    return $res
}

function Get-PageBlocks {
    param($Page, [double]$HeaderBottom)
    $items = New-Object System.Collections.Generic.List[object]
    foreach ($it in $Page.Items) {
        if ($it.Kind -in 'blank', 'line') { continue }
        if ($it.T + $it.H -le $HeaderBottom + 2) { continue }
        $items.Add($it)
    }
    $blocks = New-Object System.Collections.Generic.List[object]
    $loose = New-Object System.Collections.Generic.List[object]
    if ($items.Count -eq 0) { return , $blocks }
    $area = [math]::Max(1.0, [double]$Page.Width * ([double]$Page.EndY - [double]$Page.StartY))
    $top = New-PageTree $items
    Add-NodeBlocks -Nodes $top -Blocks $blocks -Loose $loose -PageArea $area
    Add-LooseBlocks -Loose $loose -Blocks $blocks
    Merge-CaptionBlocks -Blocks $blocks
    return , $blocks
}

# ---------------------------------------------------------------- HTML of one card

function ConvertTo-TableHtml {
    param($Cells)
    $rowKeys = New-Object System.Collections.Generic.List[double]
    $colKeys = New-Object System.Collections.Generic.List[double]
    foreach ($c in $Cells) {
        $rt = [math]::Round([double]$c.T, 1); $cl = [math]::Round([double]$c.L, 1)
        if (-not $rowKeys.Contains($rt)) { $rowKeys.Add($rt) }
        if (-not $colKeys.Contains($cl)) { $colKeys.Add($cl) }
    }
    $rowKeys.Sort(); $colKeys.Sort()
    $grid = @{}
    foreach ($c in $Cells) { $grid[('{0}|{1}' -f $rowKeys.IndexOf([math]::Round([double]$c.T, 1)), $colKeys.IndexOf([math]::Round([double]$c.L, 1)))] = $c }
    # numeric columns are right-aligned
    $numCol = @{}
    for ($ci = 0; $ci -lt $colKeys.Count; $ci++) {
        $num = 0; $all = 0
        for ($ri = 1; $ri -lt $rowKeys.Count; $ri++) {
            $c = $grid[('{0}|{1}' -f $ri, $ci)]
            if ($null -eq $c) { continue }
            $all++
            if (([string]$c.Text).Trim() -match '^[-+(]?[\d.,:%\s]+\)?\s*%?$|^#') { $num++ }
        }
        $numCol[$ci] = ($all -gt 0 -and $num -ge 0.6 * $all)
    }
    # first row is a header when it is bold or filled (Excel's table header)
    $header = $true
    for ($ci = 0; $ci -lt $colKeys.Count; $ci++) {
        $c = $grid[('0|{0}' -f $ci)]
        if ($null -ne $c -and -not $c.Bold -and -not $c.Bg) { $header = $false }
    }
    if ($rowKeys.Count -lt 2) { $header = $false }
    $sb = New-Object System.Text.StringBuilder
    [void]$sb.Append('<div class="tblwrap"><table>')
    for ($ri = 0; $ri -lt $rowKeys.Count; $ri++) {
        $isHead = ($ri -eq 0 -and $header)
        $rowBold = $true
        for ($ci = 0; $ci -lt $colKeys.Count; $ci++) { $c = $grid[('{0}|{1}' -f $ri, $ci)]; if ($null -ne $c -and -not $c.Bold) { $rowBold = $false } }
        $trCls = ''
        if (-not $isHead -and $rowBold -and $ri -eq $rowKeys.Count - 1) { $trCls = ' class="total"' }
        [void]$sb.Append('<tr' + $trCls + '>')
        for ($ci = 0; $ci -lt $colKeys.Count; $ci++) {
            $c = $grid[('{0}|{1}' -f $ri, $ci)]
            $tag = 'td'; if ($isHead) { $tag = 'th' }
            $cls = ''; if ($numCol[$ci]) { $cls = ' class="n"' }
            if ($null -eq $c) { [void]$sb.Append('<' + $tag + $cls + '></' + $tag + '>'); continue }
            $title = ' title="Wkly_Report!' + $c.Address + '"'
            $txt = ConvertTo-BText ([string]$c.Text)
            if ($isHead) { $txt = ConvertTo-HtmlText ([string]$c.Text) }
            [void]$sb.Append('<' + $tag + $cls + $title + '>' + $txt + '</' + $tag + '>')
        }
        [void]$sb.Append('</tr>')
    }
    [void]$sb.Append('</table></div>')
    return $sb.ToString()
}

function ConvertTo-BlockHtml {
    param($B)
    $head = ''
    if ($B.Title) { $head += '<h3>' + (ConvertTo-HtmlText $B.Title) + '</h3>' }
    if ($B.Sub) { $head += '<div class="sub">' + (ConvertTo-HtmlText $B.Sub) + '</div>' }
    switch ($B.Type) {
        'kpi' {
            $meta = ''
            foreach ($m in $B.Meta) { $meta += '<span>' + (ConvertTo-BText $m) + '</span>' }
            return ('<div class="card stat"><div class="lab">{0}</div><div class="val">{1}</div><div class="meta">{2}</div></div>' -f (ConvertTo-HtmlText $B.Label), (ConvertTo-HtmlText $B.Value), $meta)
        }
        'chart' {
            $spec = $B.Spec
            if ($spec -and $spec.kind) {
                if (-not $B.Title -and $spec.title) { $head = '<h3>' + (ConvertTo-HtmlText ([string]$spec.title)) + '</h3>' + $head }
                elseif ($B.Title -and $spec.title -and ([string]$spec.title).Trim() -ne ([string]$B.Title).Trim()) { $head += '<div class="sub">' + (ConvertTo-HtmlText ([string]$spec.title)) + '</div>' }
                $json = $spec | ConvertTo-Json -Depth 8 -Compress
                $fallback = ''
                if ($B.Src) { $fallback = '<img class="fb" src="{0}" alt="{1}">' -f $B.Src, (ConvertTo-Attr $B.Alt) }
                return ('<div class="card cc">{0}<div class="bchart" data-chart="{1}">{2}</div></div>' -f $head, [System.Net.WebUtility]::HtmlEncode($json), $fallback)
            }
            if (-not $B.Title -and $B.Alt) { $head = '<h3>' + (ConvertTo-HtmlText $B.Alt) + '</h3>' + $head }
            if ($B.Src) { return ('<div class="card cc">{0}<div class="bimg"><img src="{1}" alt="{2}"></div><div class="note">Chart type not redrawn; Excel''s own picture of this chart is shown.</div></div>' -f $head, $B.Src, (ConvertTo-Attr $B.Alt)) }
            return ('<div class="card cc">{0}<div class="miss">&#9888; This chart could not be read from Wkly_Report.</div></div>' -f $head)
        }
        'image' {
            if (-not $B.Src) { return '' }
            return ('<div class="card cc">{0}<div class="bimg"><img src="{1}" alt="{2}"></div></div>' -f $head, $B.Src, (ConvertTo-Attr $B.Alt))
        }
        'missing' { return ('<div class="card cc">{0}<div class="miss">&#9888; {1}</div></div>' -f $head, (ConvertTo-HtmlText $B.Alt)) }
        'table' { return ('<div class="card tc">{0}{1}</div>' -f $head, (ConvertTo-TableHtml $B.Cells)) }
        'text' {
            $ps = $B.Pieces
            $sb = New-Object System.Text.StringBuilder
            $start = 0
            $maxOther = 0.0
            for ($i = 1; $i -lt $ps.Count; $i++) { if ($ps[$i].Size -gt $maxOther) { $maxOther = $ps[$i].Size } }
            if ($ps.Count -ge 2 -and ($ps[0].Bold -or $ps[0].Size -gt $maxOther + 0.5) -and $ps[0].Text.Length -le 90) {
                [void]$sb.Append('<h3>' + (ConvertTo-HtmlText $ps[0].Text) + '</h3>'); $start = 1
            }
            [void]$sb.Append('<div class="lines">')
            for ($i = $start; $i -lt $ps.Count; $i++) {
                $p = $ps[$i]
                $cls = 'ln'; if ($p.Bold) { $cls = 'ln b' }
                [void]$sb.Append('<p class="' + $cls + '">' + (ConvertTo-BText $p.Text) + '</p>')
            }
            [void]$sb.Append('</div>')
            return ('<div class="card txt">{0}</div>' -f $sb.ToString())
        }
    }
    return ''
}

function ConvertTo-BlockLayout {
    # Cards keep Excel's arrangement: rows of cards side by side (widths in Excel's proportions),
    # cards that sit under each other in Excel stay stacked in the same column.
    param($Blocks)
    $sb = New-Object System.Text.StringBuilder
    $sorted = @($Blocks | Sort-Object T, L)
    $bands = New-Object System.Collections.Generic.List[object]
    $cur = $null; $curBottom = -1e9
    foreach ($b in $sorted) {
        if ($null -ne $cur -and $b.T -lt $curBottom - 8) { $cur.Add($b); $curBottom = [math]::Max($curBottom, $b.T + $b.H); continue }
        $cur = New-Object System.Collections.Generic.List[object]; $cur.Add($b); $bands.Add($cur); $curBottom = $b.T + $b.H
    }
    foreach ($band in $bands) {
        # columns: cards whose horizontal ranges overlap are stacked
        $cols = New-Object System.Collections.Generic.List[object]
        foreach ($b in ($band | Sort-Object L, T)) {
            $placed = $false
            foreach ($c in $cols) {
                $ox = [math]::Min($c.R, $b.L + $b.W) - [math]::Max($c.L, $b.L)
                if ($ox -gt 0.3 * [math]::Min($c.R - $c.L, $b.W)) { $c.Items.Add($b); $c.L = [math]::Min($c.L, $b.L); $c.R = [math]::Max($c.R, $b.L + $b.W); $placed = $true; break }
            }
            if (-not $placed) {
                $nc = [pscustomobject]@{ L = $b.L; R = $b.L + $b.W; Items = (New-Object System.Collections.Generic.List[object]) }
                $nc.Items.Add($b); $cols.Add($nc)
            }
        }
        $allKpi = $true
        foreach ($b in $band) { if ($b.Type -ne 'kpi') { $allKpi = $false } }
        if ($allKpi -and $cols.Count -ge 2) {
            $n = $cols.Count
            $tpl = 'repeat({0},minmax(0,1fr))' -f $n
            if ($n -gt 6) { $tpl = 'repeat(auto-fit,minmax(240px,1fr))' }
            [void]$sb.Append('<div class="row kpis" style="grid-template-columns:' + $tpl + '">')
            foreach ($c in $cols) { foreach ($b in ($c.Items | Sort-Object T)) { [void]$sb.Append((ConvertTo-BlockHtml $b)) } }
            [void]$sb.Append('</div>')
            continue
        }
        $parts = New-Object System.Collections.Generic.List[string]
        foreach ($c in $cols) { $parts.Add(('minmax(0,{0}fr)' -f (Format-Num ([math]::Max(1.0, [math]::Round(($c.R - $c.L) / 10.0, 1)))))) }
        [void]$sb.Append('<div class="row" style="grid-template-columns:' + ($parts -join ' ') + '">')
        foreach ($c in $cols) {
            [void]$sb.Append('<div class="col">')
            foreach ($b in ($c.Items | Sort-Object T, L)) {
                $html = ''
                try { $html = ConvertTo-BlockHtml $b }
                catch { Add-Issue -Severity Warning -Location ('Card "{0}"' -f $b.Title) -Message ('Could not be drawn: {0}' -f $_.Exception.Message) }
                [void]$sb.Append($html)
            }
            [void]$sb.Append('</div>')
        }
        [void]$sb.Append('</div>')
    }
    return $sb.ToString()
}

function ConvertTo-CoverHtml {
    # Page 1 as the full-screen cover: title, the text around it and its label/value details on the
    # left; KPI tiles found on the page on the right.
    param($Page, $Blocks, $Meta)
    $kpis = New-Object System.Collections.Generic.List[object]
    $pieces = New-Object System.Collections.Generic.List[object]
    $rest = New-Object System.Collections.Generic.List[object]
    foreach ($b in $Blocks) {
        if ($b.Type -eq 'kpi') { $kpis.Add($b) }
        elseif ($b.Type -eq 'text') { foreach ($p in $b.Pieces) { $pieces.Add($p) } }
        else { $rest.Add($b) }
    }
    $title = ([string]$Page.Title).Trim()
    $eyebrow = New-Object System.Collections.Generic.List[string]
    $tag = New-Object System.Collections.Generic.List[string]
    $kv = New-Object System.Collections.Generic.List[object]
    $seenTitle = $false
    $ordered = @($pieces | Sort-Object T, L)
    for ($i = 0; $i -lt $ordered.Count; $i++) {
        $t = $ordered[$i].Text
        if (-not $seenTitle -and $t -eq $title) { $seenTitle = $true; continue }
        if ($t -match '^\s*([^:]{2,40}?)\s*:\s*(.+)$') { $kv.Add(@($Matches[1].Trim(), $Matches[2].Trim())); continue }
        if ($t -match '^\s*([^:]{2,40}?)\s*:\s*$' -and $i + 1 -lt $ordered.Count) { $kv.Add(@($Matches[1].Trim(), $ordered[$i + 1].Text)); $i++; continue }
        if ($seenTitle) { $tag.Add($t) } else { $eyebrow.Add($t) }
    }
    $labels = @{}
    foreach ($pair in $kv) { $labels[$pair[0].ToLowerInvariant()] = $true }
    foreach ($known in @(@('Reporting Period', $Meta.ReportingPeriod), @('Report Date', $Meta.ReportDate), @('Prepared By', $Meta.PreparedBy))) {
        if ($known[1] -and -not $labels.ContainsKey($known[0].ToLowerInvariant())) { $kv.Add(@($known[0], $known[1])) }
    }
    $sb = New-Object System.Text.StringBuilder
    $cls = 'cover'; if ($kpis.Count -eq 0) { $cls = 'cover solo' }
    [void]$sb.Append('<div class="' + $cls + '"><div class="card hero">')
    if ($eyebrow.Count) { [void]$sb.Append('<div class="eyebrow">' + (ConvertTo-HtmlText ($eyebrow -join ' · ')) + '</div>') }
    [void]$sb.Append('<h1>' + (ConvertTo-HtmlText $title) + '</h1>')
    foreach ($t in $tag) { [void]$sb.Append('<div class="tg">' + (ConvertTo-HtmlText $t) + '</div>') }
    if ($kv.Count) {
        [void]$sb.Append('<dl class="kv">')
        foreach ($pair in $kv) { [void]$sb.Append('<dt>' + (ConvertTo-HtmlText $pair[0]) + '</dt><dd>' + (ConvertTo-HtmlText $pair[1]) + '</dd>') }
        [void]$sb.Append('</dl>')
    }
    if ($Classification) { [void]$sb.Append('<span class="conf-big">' + (ConvertTo-HtmlText $Classification) + '</span>') }
    [void]$sb.Append('</div>')
    if ($kpis.Count) {
        [void]$sb.Append('<div class="cstats">')
        foreach ($b in ($kpis | Sort-Object T, L)) { [void]$sb.Append((ConvertTo-BlockHtml $b)) }
        [void]$sb.Append('</div>')
    }
    [void]$sb.Append('</div>')
    if ($rest.Count) { [void]$sb.Append('<div class="cover-more">' + (ConvertTo-BlockLayout $rest) + '</div>') }
    return $sb.ToString()
}

function New-DesignBSite {
    # Writes ONE self-contained .html to $OutDir; pictures are read from $WorkDir\assets and embedded.
    param($Pages, [string]$OutDir, $Meta, [string]$WorkDir)
    $sb = New-Object System.Text.StringBuilder
    $n = $Pages.Count
    $enc = { param($t) ConvertTo-HtmlText $t }
    [void]$sb.Append("<!doctype html>`n<html lang=`"en`">`n<head>`n<meta charset=`"utf-8`">`n<meta name=`"viewport`" content=`"width=device-width,initial-scale=1`">`n")
    [void]$sb.Append('<title>' + (& $enc $DashboardTitle) + '</title>' + "`n<style>`n" + $DesignBCss + "`n</style>`n</head>`n<body>`n")

    # ---- page menu (hidden until opened)
    [void]$sb.Append('<div class="scrim" id="scrim"></div><aside class="rail" id="rail"><button class="x" id="mClose" title="Close (Esc)">&times;</button>')
    [void]$sb.Append('<div class="brand"><div class="logo"><span class="dot">IT</span>' + (& $enc $DashboardTitle) + '</div><small>' + (& $enc $Meta.FileName) + ' &middot; Wkly_Report</small></div><nav id="nav">')
    foreach ($p in $Pages) { [void]$sb.Append(('<a href="#s{0}"><span class="ix">{0:D2}</span>{1}</a>' -f $p.Index, (& $enc $p.Title))) }
    [void]$sb.Append('</nav><div class="foot">')
    if ($Meta.ReportingPeriod) { [void]$sb.Append((& $enc $Meta.ReportingPeriod) + '<br>') }
    [void]$sb.Append('Source: Wkly_Report &middot; read-only<br>Snapshot ' + (& $enc $Meta.Snapshot) + ' &middot; build ' + $ScriptBuild + '</div></aside>')

    # ---- top bar
    [void]$sb.Append("`n<main>`n<div class=`"top`">")
    [void]$sb.Append(('<button class="menu-btn" id="mOpen" title="Pages"><span class="bars">&#9776;</span>Pages <small id="mPos">1 / {0}</small></button>' -f $n))
    [void]$sb.Append('<h1>' + (& $enc $DashboardTitle) + '</h1>')
    if ($Meta.ReportingPeriod) { [void]$sb.Append('<span class="chip">Period <b>' + (& $enc $Meta.ReportingPeriod) + '</b></span>') }
    if ($Meta.ReportDate) { [void]$sb.Append('<span class="chip">Report date <b>' + (& $enc $Meta.ReportDate) + '</b></span>') }
    [void]$sb.Append('<div class="zoom" title="Zoom (keys + and -)"><button class="btn" id="zOut">&minus;</button><span id="zVal">100%</span><button class="btn" id="zIn">+</button><button class="btn" id="zFit">Fit</button></div>')
    [void]$sb.Append('<button class="btn" id="theme" title="Switch light / dark">Light / Dark</button><button class="btn" id="present" title="Presentation mode (P)">Present</button><button class="btn" onclick="window.print()">Print</button>')
    if ($Classification) { [void]$sb.Append('<span class="tag-conf">' + (& $enc $Classification) + '</span>') }
    [void]$sb.Append("</div>`n<div id=`"content`">`n")

    # ---- pages
    $colorIx = 0
    foreach ($p in $Pages) {
        $hdr = $null; $blocks = $null
        try {
            $hdr = Get-PageHeader $p
            $hb = $hdr.Bottom
            if ($p.Index -eq 1) { $hb = [double]$p.StartY - 1 }
            $blocks = Get-PageBlocks -Page $p -HeaderBottom $hb
        }
        catch { Add-Issue -Severity Warning -Location ('Page {0}' -f $p.Index) -Message ('Page content could not be arranged: {0}' -f (Format-ErrorDetail $_)) }
        if ($null -eq $blocks) { $blocks = New-Object System.Collections.Generic.List[object] }
        $isCover = $false
        if ($p.Index -eq 1) {
            $isCover = $true
            foreach ($b in $blocks) { if ($b.Type -in 'chart', 'table') { $isCover = $false } }
        }
        $counts = @{}
        foreach ($b in $blocks) { if ($counts.ContainsKey($b.Type)) { $counts[$b.Type]++ } else { $counts[$b.Type] = 1 } }
        $summary = (@($counts.Keys | Sort-Object | ForEach-Object { '{0} {1}' -f $counts[$_], $_ }) -join ', ')
        $coverNote = ''
        if ($isCover) { $coverNote = ' (cover)' }
        Write-Log ("  page {0} '{1}': {2}{3}" -f $p.Index, $p.Title, $summary, $coverNote)
        if ($isCover) {
            [void]$sb.Append(('<section class="cvr" id="s{0}" data-t="{1}">' -f $p.Index, (ConvertTo-Attr $p.Title)))
            $html = ''
            try { $html = ConvertTo-CoverHtml -Page $p -Blocks $blocks -Meta $Meta }
            catch { Add-Issue -Severity Warning -Location 'Cover' -Message ('Cover could not be drawn: {0}' -f (Format-ErrorDetail $_)) }
            [void]$sb.Append($html + "</section>`n")
            continue
        }
        $c = @('c0', 'c1', 'c2')[$colorIx % 3]; $colorIx++
        [void]$sb.Append(('<section class="page {2}" id="s{0}" data-t="{1}">' -f $p.Index, (ConvertTo-Attr $p.Title), $c))
        $sub = ''
        if ($hdr -and $hdr.Sub) { $sub = '<p>' + (& $enc $hdr.Sub) + '</p>' }
        [void]$sb.Append(('<div class="ph"><span class="no">{0:D2}</span><h2>{1}</h2>{2}</div><div class="pbd">' -f $p.Index, (& $enc $p.Title), $sub))
        $layout = ''
        try { $layout = ConvertTo-BlockLayout $blocks }
        catch { Add-Issue -Severity Warning -Location ('Page {0}' -f $p.Index) -Message ('Page could not be drawn: {0}' -f (Format-ErrorDetail $_)) }
        if (-not $layout) { $layout = '<div class="card txt"><div class="lines"><p class="ln">No content was found on this page of Wkly_Report.</p></div></div>' }
        [void]$sb.Append($layout)
        [void]$sb.Append(('</div><div class="pfoot"><span>{0} &middot; Wkly_Report rows {1}&ndash;{2}</span><span>Page <b>{3}</b> of {4}</span></div></section>' -f (& $enc $Classification), $p.StartRow, $p.EndRow, $p.Index, $n))
        [void]$sb.Append("`n")
    }
    [void]$sb.Append("</div>`n</main>`n")
    [void]$sb.Append('<div class="hud"><span id="hpos"></span><button class="btn" id="hp">&#9664;</button><button class="btn" id="hn">&#9654;</button><button class="btn" id="hx">Exit (Esc)</button></div>')
    [void]$sb.Append("`n<script>`n" + $DesignBJs + "`n</script>`n</body>`n</html>`n")

    $html = $sb.ToString()
    # every picture is embedded so the single .html works when sent by email or Teams
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
    $external = [regex]::Matches($html, '(?:src|href)="assets/[^"]*"').Count
    if ($external -gt 0) { Add-Issue -Severity Error -Location 'Dashboard file' -Message ("{0} reference(s) to pictures could not be embedded." -f $external) }

    $index = Join-Path $OutDir ($Meta.BaseName + '.html')
    [System.IO.File]::WriteAllText($index, $html, (New-Object System.Text.UTF8Encoding($false)))
    $mb = [math]::Round((Get-Item -LiteralPath $index).Length / 1MB, 1)
    Write-Log ("Dashboard is a single self-contained file ({0} MB, {1} pictures embedded)." -f $mb, $embedded.Count)
    return $index
}

# ======================================================================================
#  Design B stylesheet and script (embedded in the single .html)
# ======================================================================================

$DesignBCss = @'
:root{
  color-scheme:light;
  --bg:#eeebe3;--surface:#ffffff;--surface-2:#f7f5ef;--line:#e1dccf;--ink:#0b0b0b;--ink-2:#52514e;--ink-3:#7c7b76;
  --s1:#1e7d5a;--s3:#d39b1a;--s4:#8a8f98;--s8:#e34948;--c4:#4a4f58;--c5:#a86b0c;--c6:#6fbf9a;
  --good:#1a7f4b;--good-bg:#e3f4ea;--warn:#a66300;--warn-bg:#fbf0db;--bad:#c23b3a;--bad-bg:#fbe5e4;--info-bg:#ece9e0;
  --accent:#1e7d5a;--grid:#ebe7de;--shadow:0 1px 2px rgba(40,32,20,.08),0 6px 18px rgba(40,32,20,.07);
  --rail:#262a31;--rail-ink:#ecebe7;--rail-ink-2:#b0ada5;--rail-on:rgba(255,255,255,.16);--page-bg:#fdfcf9;
  --top-bg:#262a31;--top-ink:#ffffff;--top-ink-2:#c9c5bc;--top-btn:#3b414b;--top-btn-h:#4a515c;--top-btn-ink:#ffffff;--top-chip:#343943;
  --menu-bg:#f2b33d;--menu-ink:#262a31;--conf-bg:#c8333a;--conf-ink:#ffffff;--top-line:transparent;
  --h0:#262a31;--h1:#13805a;--h2:#a86b0c;--cover:#262a31;
}
@media (prefers-color-scheme:dark){:root:where(:not([data-theme="light"])){
  color-scheme:dark;--bg:#121211;--surface:#1a1a19;--surface-2:#222220;--line:#2e2e2b;--ink:#f0efec;--ink-2:#c3c2b7;--ink-3:#8f8e86;
  --s1:#2f9d72;--s3:#b08018;--s4:#80868f;--s8:#e66767;--c4:#a3a8b1;--c5:#c98a2e;--c6:#7fcfa9;
  --good:#4cc38a;--good-bg:#173226;--warn:#e0a640;--warn-bg:#33270f;--bad:#f07777;--bad-bg:#3a1b1b;--info-bg:#22242c;--accent:#2f9d72;--grid:#272725;--shadow:none;
  --rail:#1b1d22;--rail-ink:#d7dbe4;--rail-ink-2:#8f96a3;--rail-on:rgba(255,255,255,.12);--page-bg:var(--surface);
  --top-bg:color-mix(in srgb,var(--bg) 88%,transparent);--top-ink:var(--ink);--top-ink-2:var(--ink-3);--top-btn:var(--surface);--top-btn-h:var(--surface-2);--top-btn-ink:var(--ink);--top-chip:var(--surface);
  --menu-bg:var(--rail);--menu-ink:#fff;--conf-bg:var(--bad-bg);--conf-ink:var(--bad);--top-line:var(--line);--h0:#30343c;--cover:#1f2228}}
:root[data-theme="dark"]{
  color-scheme:dark;--bg:#121211;--surface:#1a1a19;--surface-2:#222220;--line:#2e2e2b;--ink:#f0efec;--ink-2:#c3c2b7;--ink-3:#8f8e86;
  --s1:#2f9d72;--s3:#b08018;--s4:#80868f;--s8:#e66767;--c4:#a3a8b1;--c5:#c98a2e;--c6:#7fcfa9;
  --good:#4cc38a;--good-bg:#173226;--warn:#e0a640;--warn-bg:#33270f;--bad:#f07777;--bad-bg:#3a1b1b;--info-bg:#22242c;--accent:#2f9d72;--grid:#272725;--shadow:none;
  --rail:#1b1d22;--rail-ink:#d7dbe4;--rail-ink-2:#8f96a3;--rail-on:rgba(255,255,255,.12);--page-bg:var(--surface);
  --top-bg:color-mix(in srgb,var(--bg) 88%,transparent);--top-ink:var(--ink);--top-ink-2:var(--ink-3);--top-btn:var(--surface);--top-btn-h:var(--surface-2);--top-btn-ink:var(--ink);--top-chip:var(--surface);
  --menu-bg:var(--rail);--menu-ink:#fff;--conf-bg:var(--bad-bg);--conf-ink:var(--bad);--top-line:var(--line);--h0:#30343c;--cover:#1f2228}
*{box-sizing:border-box}
html{scroll-behavior:smooth}
body{margin:0;background:var(--bg);color:var(--ink);font:18px/1.45 "Segoe UI",system-ui,-apple-system,Arial,sans-serif;-webkit-print-color-adjust:exact;print-color-adjust:exact}
/* page menu */
.rail{position:fixed;inset:0 auto 0 0;width:300px;background:var(--rail);color:var(--rail-ink);display:flex;flex-direction:column;z-index:40;transform:translateX(-100%);transition:transform .22s ease}
body.menu .rail{transform:none;box-shadow:8px 0 32px rgba(20,18,15,.35)}
.scrim{position:fixed;inset:0;background:rgba(20,18,15,.35);z-index:35;opacity:0;pointer-events:none;transition:opacity .2s}
body.menu .scrim{opacity:1;pointer-events:auto}
.rail .x{position:absolute;top:16px;right:14px;background:rgba(255,255,255,.12);border:0;color:#fff;width:30px;height:30px;border-radius:8px;font-size:16px;cursor:pointer}
.brand{padding:22px 50px 18px 22px;border-bottom:1px solid rgba(255,255,255,.12)}
.brand .logo{display:flex;align-items:center;gap:10px;font-weight:700;font-size:15px}
.brand .dot{width:28px;height:28px;flex:none;border-radius:8px;background:#f2b33d;color:#262a31;display:grid;place-items:center;font-size:13px;font-weight:800}
.brand small{display:block;color:var(--rail-ink-2);font-size:12px;margin-top:6px;overflow-wrap:anywhere}
.rail nav{padding:12px 10px;overflow:auto;flex:1}
.rail a{display:flex;align-items:center;gap:11px;padding:9px 12px;border-radius:9px;color:var(--rail-ink);text-decoration:none;font-size:14px}
.rail a:hover{background:var(--rail-on);color:#fff}
.rail a.on{background:var(--rail-on);color:#fff;font-weight:600;box-shadow:inset 3px 0 0 #f2b33d}
.rail a .ix{font-variant-numeric:tabular-nums;color:var(--rail-ink-2);width:20px;flex:none;font-size:12px}
.rail a.on .ix{color:#f2b33d}
.rail .foot{padding:14px 22px;border-top:1px solid rgba(255,255,255,.12);font-size:11.5px;color:var(--rail-ink-2)}
/* top bar */
main{padding:0 32px 40px}
.top{position:sticky;top:0;z-index:10;background:var(--top-bg);backdrop-filter:blur(8px);display:flex;align-items:center;gap:12px;padding:14px 32px;margin:0 -32px 8px;border-bottom:1px solid var(--top-line);color:var(--top-ink)}
.top h1{font-size:23px;margin:0;flex:1;letter-spacing:-.2px;min-width:0;white-space:nowrap;overflow:hidden;text-overflow:ellipsis}
.menu-btn{display:flex;align-items:center;gap:8px;padding:8px 13px;font-size:15px;font-weight:700;background:var(--menu-bg);color:var(--menu-ink);border:0;border-radius:10px;cursor:pointer;flex:none}
.menu-btn .bars{font-size:18px;line-height:1}
.menu-btn small{font-weight:500;opacity:.85;font-size:12.5px}
.chip{border:1px solid var(--top-line);background:var(--top-chip);border-radius:999px;padding:5px 12px;font-size:16px;color:var(--top-ink-2);white-space:nowrap;max-width:420px;overflow:hidden;text-overflow:ellipsis}
.chip b{color:var(--top-ink)}
.btn{border:1px solid var(--line);background:var(--surface);color:var(--ink);border-radius:8px;padding:7px 12px;font:inherit;font-size:17px;cursor:pointer;white-space:nowrap}
.btn:hover{background:var(--surface-2)}
.top .btn{background:var(--top-btn);color:var(--top-btn-ink);border-color:var(--top-line)}.top .btn:hover{background:var(--top-btn-h)}
.tag-conf{font-size:14px;letter-spacing:1.2px;font-weight:700;padding:5px 9px;border-radius:6px;background:var(--conf-bg);color:var(--conf-ink);white-space:nowrap}
.zoom{display:flex;align-items:center;gap:4px}
.zoom .btn{padding:6px 10px;min-width:34px}
.zoom span{font-size:16px;color:var(--top-ink-2);min-width:46px;text-align:center;font-variant-numeric:tabular-nums}
/* pages */
section{scroll-margin-top:80px}
.page{--sec:var(--h0);background:var(--page-bg);border:1px solid var(--line);border-radius:18px;box-shadow:var(--shadow);overflow:hidden;margin-top:26px}
.page.c1{--sec:var(--h1)}.page.c2{--sec:var(--h2)}
.ph{display:flex;align-items:center;gap:14px;padding:16px 24px;color:#fff;background:var(--sec)}
.ph .no{background:rgba(255,255,255,.2);border-radius:8px;padding:4px 10px;font-weight:700;font-size:17px;font-variant-numeric:tabular-nums}
.ph h2{margin:0;font-size:27px;letter-spacing:-.3px}
.ph p{margin:0 0 0 auto;font-size:17px;opacity:.88;text-align:right}
.pbd{padding:20px 24px 22px;display:flex;flex-direction:column;gap:16px}
.pfoot{display:flex;justify-content:space-between;padding:10px 24px;border-top:1px solid var(--line);font-size:15px;color:var(--ink-3);background:var(--surface-2)}
.pfoot b{color:var(--ink)}
.row{display:grid;gap:16px;align-items:stretch}
.col{display:flex;flex-direction:column;gap:16px;min-width:0}
.col>.card{flex:1 1 auto}
.card{background:var(--surface);border:1px solid var(--line);border-radius:14px;box-shadow:var(--shadow);padding:16px 18px;min-width:0}
.card h3{margin:0 0 2px;font-size:19px;color:color-mix(in srgb,var(--sec,var(--ink)) 70%,var(--ink))}
:root[data-theme="dark"] .card h3{color:var(--ink)}
@media (prefers-color-scheme:dark){:root:not([data-theme="light"]) .card h3{color:var(--ink)}}
.card .sub{color:var(--ink-3);font-size:16px;margin-bottom:10px}
/* KPI tiles */
.stat{border-top:3px solid var(--sec,var(--accent));background:linear-gradient(180deg,color-mix(in srgb,var(--sec,var(--accent)) 7%,var(--surface)),var(--surface) 70%)}
.stat .lab{color:var(--ink-2);font-size:17px}
.stat .val{font-size:38px;font-weight:650;letter-spacing:-.6px;font-variant-numeric:tabular-nums;margin:4px 0 2px;overflow-wrap:anywhere}
.stat .meta{font-size:16px;color:var(--ink-3);display:flex;gap:4px 10px;align-items:center;flex-wrap:wrap}
.pill{display:inline-flex;align-items:center;border-radius:999px;padding:1px 9px;font-size:15px;font-weight:600}
.ok{color:var(--good);background:var(--good-bg)}.wn{color:var(--warn);background:var(--warn-bg)}.bd{color:var(--bad);background:var(--bad-bg)}.nt{color:var(--ink-2);background:var(--surface-2)}
/* text cards */
.txt .lines{display:flex;flex-direction:column;gap:6px;margin-top:6px}
.txt .ln{margin:0;color:var(--ink-2);font-size:17px;white-space:pre-wrap}
.txt .ln.b{color:var(--ink);font-weight:600}
/* tables */
.tblwrap{overflow-x:auto}
table{border-collapse:collapse;width:100%;font-size:17px}
th{text-align:left;font-weight:600;color:var(--ink-3);font-size:15px;letter-spacing:.3px;text-transform:uppercase;padding:8px 10px;border-bottom:1px solid var(--line)}
td{padding:9px 10px;border-bottom:1px solid var(--line);vertical-align:top}
td.n,th.n{text-align:right;font-variant-numeric:tabular-nums;white-space:nowrap}
tr.total td{font-weight:650;border-bottom:0}
/* charts */
.bchart{position:relative;margin-top:6px}
.bchart svg{display:block;width:100%;overflow:visible}
.bchart .fb{display:none}
.bchart.err .fb{display:block;width:100%}
.bchart .ax{fill:var(--ink-3);font-size:14px}
.bchart .gl{stroke:var(--grid)}
.bchart .base{stroke:var(--line)}
.bchart .lbl{fill:var(--ink-2);font-size:14px;font-variant-numeric:tabular-nums;paint-order:stroke;stroke:var(--surface);stroke-width:4px;stroke-linejoin:round}
.bchart .hit{fill:transparent}
.legend{display:flex;flex-wrap:wrap;gap:4px 16px;font-size:16px;color:var(--ink-2);margin-top:8px}
.legend i{display:inline-block;width:11px;height:11px;border-radius:3px;margin-right:6px;vertical-align:-1px}
.legend i.ln{height:3px;width:16px;vertical-align:4px}
.tip{position:absolute;pointer-events:none;background:var(--ink);color:var(--surface);font-size:16px;padding:6px 9px;border-radius:7px;white-space:nowrap;display:none;z-index:5;box-shadow:0 4px 14px rgba(0,0,0,.2)}
.bimg img{display:block;max-width:100%;height:auto;margin:6px auto 0}
.note{font-size:14px;color:var(--ink-3);margin-top:6px}
.miss{border:1.5px dashed var(--bad);background:var(--bad-bg);color:var(--bad);border-radius:10px;padding:14px;font-weight:600}
.xerr{outline:2px dashed var(--bad);outline-offset:1px;border-radius:2px;cursor:help}
/* cover */
.cvr{height:var(--coverH,calc(100vh - 110px));min-height:520px;display:flex;flex-direction:column;margin-top:10px}
.cover{flex:1;min-height:0;display:grid;grid-template-columns:1.55fr 1fr;gap:20px}
.cover.solo{grid-template-columns:1fr}
.hero{display:flex;flex-direction:column;justify-content:center;padding:4vh 4.5vw;overflow:hidden;background:var(--cover);color:#fff;border:0;border-radius:22px}
.hero .eyebrow{color:#f2b33d;font-size:clamp(13px,min(1.2vw,2vh),20px);letter-spacing:3px;text-transform:uppercase;font-weight:700}
.hero h1{font-size:clamp(38px,min(5.2vw,9vh),96px);line-height:1.02;margin:2.2vh 0 1.4vh;letter-spacing:-1.5px}
.hero .tg{color:#ecebe7;font-size:clamp(15px,min(1.5vw,2.5vh),26px);margin-bottom:1vh}
.hero .kv{display:grid;grid-template-columns:max-content 1fr;font-size:clamp(15px,min(1.4vw,2.4vh),26px);gap:1.4vh 28px;margin-top:3vh}
.hero .kv dt{color:#b8b4aa}.hero .kv dd{color:#fff;margin:0;font-weight:600}
.hero .conf-big{margin-top:4vh;display:inline-block;align-self:flex-start;border:2px solid #f2b33d;color:#f2b33d;border-radius:8px;padding:6px 14px;letter-spacing:4px;font-weight:700;font-size:clamp(12px,1vw,16px)}
.cstats{display:flex;flex-direction:column;gap:20px;min-height:0}
.cstats .stat{flex:1;display:flex;flex-direction:column;justify-content:center;padding:2vh 2vw;border-radius:22px;border-top:0;border-left:6px solid var(--s3);background:var(--surface)}
.cstats .stat:nth-child(3n+1){border-left-color:var(--h0)}.cstats .stat:nth-child(3n+2){border-left-color:var(--s1)}
.cstats .stat .lab{font-size:clamp(14px,min(1.25vw,2.1vh),22px)}
.cstats .stat .val{font-size:clamp(32px,min(4.2vw,7vh),76px);letter-spacing:-1.5px}
.cstats .stat .meta{font-size:clamp(12px,min(1vw,1.8vh),18px)}
.cover-more{margin-top:20px;display:flex;flex-direction:column;gap:16px}
/* presentation */
.hud{display:none}
body.present .rail,body.present .scrim,body.present .top{display:none}
body.present main{padding:20px 32px}
body.present section{display:none;margin-top:0}
body.present section.cur{display:block}
body.present .hud{display:flex;position:fixed;right:18px;bottom:18px;gap:6px;align-items:center;background:var(--surface);border:1px solid var(--line);border-radius:999px;padding:5px 6px 5px 14px;font-size:14px;box-shadow:var(--shadow);z-index:30}
@media (max-width:1600px){.top .chip{display:none}}
@media (max-width:980px){main{padding:0 16px 30px}.top{margin:0 -16px 8px;padding:12px 16px}.row{grid-template-columns:1fr!important}.cover{grid-template-columns:1fr}.cvr{height:auto}}
@media print{@page{size:A4 landscape;margin:8mm}body{background:#fff}.rail,.scrim,.top,.hud{display:none!important}main{padding:0}section{break-after:page}.page{margin:0;box-shadow:none}.card{box-shadow:none}.cvr{height:180mm}#content{zoom:1!important}}
'@

$DesignBJs = @'
(function () {
  'use strict';
  var body = document.body, root = document.documentElement;
  var secs = [].slice.call(document.querySelectorAll('#content > section'));
  var ERR = /#(DIV\/0!|N\/A|VALUE!|REF!|NAME\?|NUM!|NULL!|SPILL!|CALC!|FIELD!|BLOCKED!|CONNECT!|BUSY!|UNKNOWN!|GETTING_DATA)/g;
  // Excel error values are shown exactly as displayed, with an outline so they are noticed.
  (function () {
    var w = document.createTreeWalker(document.getElementById('content'), NodeFilter.SHOW_TEXT, null), hits = [];
    while (w.nextNode()) { ERR.lastIndex = 0; if (ERR.test(w.currentNode.nodeValue)) hits.push(w.currentNode); }
    hits.forEach(function (t) {
      var s = t.nodeValue, frag = document.createDocumentFragment(), last = 0, m;
      ERR.lastIndex = 0;
      while ((m = ERR.exec(s))) {
        if (m.index > last) frag.appendChild(document.createTextNode(s.slice(last, m.index)));
        var sp = document.createElement('span'); sp.className = 'xerr'; sp.title = 'Excel error value - shown exactly as in Wkly_Report';
        sp.textContent = m[0]; frag.appendChild(sp); last = m.index + m[0].length;
      }
      if (last < s.length) frag.appendChild(document.createTextNode(s.slice(last)));
      t.parentNode.replaceChild(frag, t);
    });
  })();

  // ---------- number formats (Excel's format codes, so labels read as in Excel)
  function decimalsOf(nf) { var m = /0\.(0+)/.exec(nf || ''); return m ? m[1].length : 0; }
  function fmt(v, nf) {
    if (v === null || v === undefined || isNaN(v)) return '';
    nf = nf || 'General';
    var sec = nf.split(';');
    if (v === 0 && sec.length > 2) { var z = sec[2].replace(/_.|\*.|["\\\s]/g, ''); if (/^[-–]$/.test(z)) return '-'; }
    if (v < 0 && sec.length > 1) nf = sec[1]; else nf = sec[0];
    if (/general/i.test(nf)) { var r = Math.round(v * 100) / 100; return Math.abs(r) >= 1000 ? r.toLocaleString('en-US') : String(r); }
    if (/\[?h\]?:mm/i.test(nf)) {
      var t = Math.round(Math.abs(v) * 86400), h = Math.floor(t / 3600), mi = Math.floor(t % 3600 / 60), s2 = t % 60;
      return h + ':' + (mi < 10 ? '0' : '') + mi + (/ss/i.test(nf) ? ':' + (s2 < 10 ? '0' : '') + s2 : '');
    }
    var d = decimalsOf(nf), pct = nf.indexOf('%') >= 0, x = Math.abs(pct ? v * 100 : v);
    var str = x.toFixed(d);
    if (nf.indexOf(',') >= 0) { var parts = str.split('.'); parts[0] = parts[0].replace(/\B(?=(\d{3})+(?!\d))/g, ','); str = parts.join('.'); }
    if (v < 0 && sec.length < 2) str = '-' + str;
    return str + (pct ? '%' : '');
  }
  function esc(t) { return String(t === null || t === undefined ? '' : t).replace(/&/g, '&amp;').replace(/</g, '&lt;').replace(/>/g, '&gt;').replace(/"/g, '&quot;'); }
  function niceAxis(min, max) {
    if (min > 0) min = 0;
    if (max <= min) max = min + 1;
    var span = max - min, step = Math.pow(10, Math.floor(Math.log10(span / 4))), f = span / 4 / step;
    step *= f <= 1 ? 1 : f <= 2 ? 2 : f <= 2.5 ? 2.5 : f <= 5 ? 5 : 10;
    return { min: Math.floor(min / step) * step, max: Math.ceil(max * 1.04 / step) * step, major: step };
  }
  function axisOf(ax, vals) {
    if (ax && ax.max !== null && ax.max !== undefined && ax.major > 0 && ax.max > ax.min && (ax.max - ax.min) / ax.major <= 12) return { min: ax.min, max: ax.max, major: ax.major };
    var lo = Math.min.apply(null, vals.concat([0])), hi = Math.max.apply(null, vals.concat([0]));
    return niceAxis(lo, hi);
  }

  // ---------- colours: Excel's colours are mapped onto the dashboard palette by hue
  function hue(hex) {
    var m = /^#?([0-9a-f]{2})([0-9a-f]{2})([0-9a-f]{2})/i.exec(hex || ''); if (!m) return null;
    var r = parseInt(m[1], 16) / 255, g = parseInt(m[2], 16) / 255, b = parseInt(m[3], 16) / 255;
    var mx = Math.max(r, g, b), mn = Math.min(r, g, b), l = (mx + mn) / 2, d = mx - mn, h = 0, s = 0;
    if (d > 0) { s = d / (1 - Math.abs(2 * l - 1)); h = mx === r ? ((g - b) / d) % 6 : mx === g ? (b - r) / d + 2 : (r - g) / d + 4; h *= 60; if (h < 0) h += 360; }
    return { h: h, s: s, l: l };
  }
  var ORDER = ['--s1', '--s3', '--s4', '--c4', '--c5', '--c6'];
  function palette() {
    var used = {}, cache = {};
    return function (excel, key) {
      var k = excel ? excel.toLowerCase() : '#' + key;
      if (cache[k]) return cache[k];
      var pref = null, c = hue(excel);
      if (c) {
        if (c.s < 0.15 || c.l > 0.93) pref = '--s4';
        else if ((c.h < 14 || c.h >= 340) && c.s >= 0.35) pref = '--s8';
        else if (c.h >= 14 && c.h < 70) pref = '--s3';
        else if (c.h >= 70 && c.h < 170) pref = '--s1';
      }
      var pick = null;
      if (pref === '--s8') pick = pref;                       // red stays red: it marks problems
      else if (pref && !used[pref]) pick = pref;
      else for (var i = 0; i < ORDER.length; i++) if (!used[ORDER[i]]) { pick = ORDER[i]; break; }
      if (!pick) pick = '--s4';
      used[pick] = true; cache[k] = 'var(' + pick + ')';
      return cache[k];
    };
  }
  function excelColor(s) { return (s.kind === 'line' || s.kind === 'area') ? (s.lineOn !== false && s.line ? s.line : s.fill) : (s.fillOn !== false && s.fill ? s.fill : s.line); }
  function distinct(arr) { var o = {}, n = 0; (arr || []).forEach(function (c) { if (c && !o[c]) { o[c] = 1; n++; } }); return n; }

  function tipOn(el, tips) {
    var tip = el.querySelector('.tip');
    el.onmousemove = function (e) {
      var t = e.target.getAttribute && e.target.getAttribute('data-tip');
      if (t === null || t === undefined) { tip.style.display = 'none'; return; }
      var r = el.getBoundingClientRect(), z = parseFloat(getComputedStyle(document.getElementById('content')).zoom) || 1;
      tip.innerHTML = tips[+t]; tip.style.display = 'block';
      var x = (e.clientX - r.left) / z + 14, y = (e.clientY - r.top) / z + 14;
      if (x + tip.offsetWidth > el.clientWidth) x = Math.max(0, x - tip.offsetWidth - 28);
      tip.style.left = x + 'px'; tip.style.top = y + 'px';
    };
    el.onmouseleave = function () { tip.style.display = 'none'; };
  }
  function legendHtml(items) {
    return '<div class="legend">' + items.map(function (i) { return '<span><i class="' + (i.line ? 'ln' : '') + '" style="background:' + i.c + '"></i>' + esc(i.n) + '</span>'; }).join('') + '</div>';
  }

  // ---------- chart drawing (Excel's data, scale and number formats; dashboard look)
  function draw(el) {
    var sp = el._spec || (el._spec = JSON.parse(el.getAttribute('data-chart')));
    var fb = el.querySelector('.fb'); var fbHtml = fb ? fb.outerHTML : '';
    var W = Math.max(280, el.clientWidth), o = [], tips = [], col = palette();
    var series = sp.series || [], cats = [];
    series.forEach(function (s) { if (s.cats && s.cats.length > cats.length) cats = s.cats; });
    var pie = sp.kind === 'pie' || sp.kind === 'doughnut';
    if (pie) {
      var s0 = series[0], vals = (s0.values || []).map(function (v) { return v > 0 ? v : 0; }), tot = vals.reduce(function (a, b) { return a + b; }, 0);
      var S = Math.min(W, 340), c = S / 2, R = c - 6, r0 = R * 0.62, a = (sp.firstAngle || 0) * Math.PI / 180 - Math.PI / 2, leg = [];
      var many = distinct(s0.points) >= 2;
      vals.forEach(function (v, i) {
        var cc = col(many ? s0.points[i] : null, i), cat = (s0.cats || [])[i] || ('#' + (i + 1));
        var lf = s0.labels && s0.labels.format && s0.labels.format.indexOf('%') < 0 ? s0.labels.format : '#,##0';
        var pctTxt = s0.labels && s0.labels.percent && tot ? ' (' + fmt(v / tot, s0.labels.format && s0.labels.format.indexOf('%') >= 0 ? s0.labels.format : '0%') + ')' : '';
        leg.push({ n: cat + ' · ' + fmt(s0.values[i], lf) + pctTxt, c: cc });
        if (!tot || v <= 0) return;
        var f = v / tot, a2 = a + f * 2 * Math.PI - (f < 1 ? 0.012 : 0), big = a2 - a > Math.PI ? 1 : 0;
        var P = function (rr, an) { return (c + rr * Math.cos(an)).toFixed(2) + ',' + (c + rr * Math.sin(an)).toFixed(2); };
        var d = f >= 0.9999 ? 'M' + P(R, 0) + 'A' + R + ',' + R + ' 0 1 1 ' + P(R, Math.PI) + 'A' + R + ',' + R + ' 0 1 1 ' + P(R, 0) + 'M' + P(r0, 0) + 'A' + r0 + ',' + r0 + ' 0 1 0 ' + P(r0, Math.PI) + 'A' + r0 + ',' + r0 + ' 0 1 0 ' + P(r0, 0) + 'Z'
          : 'M' + P(R, a) + 'A' + R + ',' + R + ' 0 ' + big + ' 1 ' + P(R, a2) + 'L' + P(r0, a2) + 'A' + r0 + ',' + r0 + ' 0 ' + big + ' 0 ' + P(r0, a) + 'Z';
        tips.push('<b>' + esc(cat) + '</b><br>' + esc(fmt(s0.values[i], lf)) + ' · ' + (f * 100).toFixed(2) + '%');
        o.push('<path d="' + d + '" fill="' + cc + '" fill-rule="evenodd" data-tip="' + (tips.length - 1) + '"/>');
        a = a2 + (f < 1 ? 0.012 : 0);
      });
      el.innerHTML = '<svg viewBox="0 0 ' + S + ' ' + S + '" style="max-width:' + S + 'px;margin:0 auto">' + o.join('') + '</svg>' + legendHtml(leg) + '<div class="tip"></div>' + fbHtml;
      tipOn(el, tips); return;
    }
    var horiz = series.some(function (s) { return s.kind === 'bar'; });
    var stacked = series.some(function (s) { return s.stacked; });
    var n = Math.max(1, cats.length || (series[0] && series[0].values ? series[0].values.length : 1));
    var prim = [], sec = [];
    series.forEach(function (s) { (s.axis === 2 ? sec : prim).push.apply(s.axis === 2 ? sec : prim, (s.values || []).filter(function (v) { return v !== null; })); });
    if (stacked) prim = cats.map(function (c, i) { return series.filter(function (s) { return s.axis !== 2; }).reduce(function (a, s) { return a + Math.max(0, (s.values || [])[i] || 0); }, 0); });
    var A1 = axisOf(sp.y, prim.length ? prim : [0]), A2 = sec.length ? axisOf(sp.y2, sec) : null;
    var yf = (sp.y || {}).format, y2f = (sp.y2 || {}).format;
    var perPoint = series.map(function (s) { return (s.kind === 'column' || s.kind === 'bar') && distinct(s.points) >= 2; });
    // bars coloured point by point (e.g. green / amber / red by age) claim their colours first
    series.forEach(function (s, i) { if (perPoint[i]) s.points.forEach(function (c, j) { if (c) col(c, 'p' + j); }); });
    var colors = series.map(function (s, i) {
      if (perPoint[i]) { for (var j = 0; j < s.points.length; j++) if (s.points[j]) return col(s.points[j], 'p' + j); }
      return col(excelColor(s), i);
    });
    var anyLabels = series.some(function (s) { return s.labels; });
    var H, m;
    if (horiz) {
      var maxCat = Math.max.apply(null, cats.map(function (c) { return String(c).length; }).concat([2]));
      m = { l: Math.min(W * 0.42, maxCat * 8.6 + 14), r: 120, t: 6, b: 6 };
      H = n * 38 * Math.max(1, stacked ? 1 : Math.min(2, series.length)) + m.t + m.b;
    } else {
      var tickW = Math.max.apply(null, [fmt(A1.max, yf), fmt(A1.min, yf)].map(function (t) { return t.length; })) * 8.6 + 14;
      var band0 = (W - tickW - 10) / n, longest = Math.max.apply(null, cats.map(function (c) { return String(c).length; }).concat([1]));
      var rot = longest * 8 > band0 * 0.95;
      m = { l: tickW, r: A2 ? Math.max.apply(null, [fmt(A2.max, y2f)].map(function (t) { return t.length; })) * 8.6 + 14 : 10, t: anyLabels ? 26 : 14, b: rot ? Math.min(140, longest * 6.2 + 20) : 34 };
      H = 300 + m.b - 34;
    }
    var pl = m.l, pr = W - m.r, pt = m.t, pb = H - m.b;
    var V = function (v, ax) { ax = ax || A1; var f = (v - ax.min) / (ax.max - ax.min); return horiz ? pl + f * (pr - pl) : pb - f * (pb - pt); };
    if (!horiz) for (var g = A1.min; g <= A1.max + A1.major / 1000; g += A1.major) {
      var gy = V(g);
      o.push('<line class="gl" x1="' + pl + '" x2="' + pr + '" y1="' + gy + '" y2="' + gy + '"/><text class="ax" x="' + (pl - 8) + '" y="' + (gy + 5) + '" text-anchor="end">' + esc(fmt(g, yf)) + '</text>');
    }
    if (A2 && !horiz) for (var g2 = A2.min; g2 <= A2.max + A2.major / 1000; g2 += A2.major) o.push('<text class="ax" x="' + (pr + 8) + '" y="' + (V(g2, A2) + 5) + '">' + esc(fmt(g2, y2f)) + '</text>');
    var band = (horiz ? pb - pt : pr - pl) / n;
    var xr = (sp.x || {}).reverse;
    var order = function (i) { return horiz ? (xr ? i : n - 1 - i) : (xr ? n - 1 - i : i); };   // Excel draws bar categories bottom-up unless reversed
    var C = function (i) { return (horiz ? pt : pl) + band * (order(i) + 0.5); };
    cats.forEach(function (c, i) {
      if (horiz) o.push('<text class="ax" x="' + (pl - 10) + '" y="' + (C(i) + 5) + '" text-anchor="end" style="font-size:16px;fill:var(--ink-2)">' + esc(c) + '</text>');
      else if (rot) o.push('<text class="ax" transform="translate(' + (C(i) + 5) + ',' + (pb + 14) + ') rotate(-35)" text-anchor="end">' + esc(c) + '</text>');
      else o.push('<text class="ax" x="' + C(i) + '" y="' + (pb + 22) + '" text-anchor="middle">' + esc(c) + '</text>');
    });
    var bars = series.filter(function (s) { return s.kind === 'column' || s.kind === 'bar'; });
    var nb = stacked ? 1 : Math.max(1, bars.length), groupW = band * (stacked ? 0.56 : Math.min(0.78, 0.36 * nb + 0.2)), bw = groupW / nb;
    var stackBase = cats.map(function () { return 0; }), labels = [], hits = [];
    // when a line also carries value labels, bar labels go inside the foot of the bar so the two never collide
    var boxes = [];   // [centre x, baseline y, width] of bar labels drawn above bars
    var lineLabels = series.some(function (s) { return (s.kind === 'line' || s.kind === 'area') && s.labels && s.labels.value; });
    bars.forEach(function (s, bi) {
      var si = series.indexOf(s);
      (s.values || []).forEach(function (v, i) {
        if (v === null || v === undefined) return;
        var cc = perPoint[si] ? col(s.points[i], 'p' + i) : colors[si];
        var base = stacked ? stackBase[i] : 0, end = base + v;
        if (stacked) stackBase[i] = end;
        var c0 = C(i) - groupW / 2 + (stacked ? 0 : bi * bw);
        var p0 = V(Math.max(A1.min, Math.min(A1.max, base))), p1 = V(Math.max(A1.min, Math.min(A1.max, end)));
        var gapPx = stacked && bi > 0 ? 2 : 0;
        if (horiz) o.push('<rect x="' + (Math.min(p0, p1) + gapPx) + '" y="' + (c0 + 1) + '" width="' + Math.max(0, Math.abs(p1 - p0) - gapPx) + '" height="' + Math.max(1, bw - 2) + '" rx="3" fill="' + cc + '"/>');
        else o.push('<rect x="' + (c0 + 1) + '" y="' + Math.min(p0, p1) + '" width="' + Math.max(1, bw - 2) + '" height="' + Math.max(0, Math.abs(p1 - p0) - gapPx) + '" rx="3" fill="' + cc + '"/>');
        if (s.labels) {
          var t = [];
          if (s.labels.value) t.push(fmt(v, s.labels.format && !/general/i.test(s.labels.format) ? s.labels.format : yf));
          if (s.labels.category) t.push(cats[i] || '');
          if (s.labels.series) t.push(s.name || '');
          if (t.length) labels.push(horiz ? '<text class="lbl" x="' + (Math.max(p0, p1) + 8) + '" y="' + (c0 + bw / 2 + 5) + '">' + esc(t.join(' · ')) + '</text>'
            : (lineLabels && Math.abs(p1 - p0) > 26 ? '<text class="lbl" x="' + (c0 + bw / 2) + '" y="' + (Math.max(p0, p1) - 10) + '" text-anchor="middle" style="fill:#fff;font-weight:600;stroke:none">' + esc(t.join(' · ')) + '</text>'
              : (boxes.push([c0 + bw / 2, Math.min(p0, p1) - 7, t.join(' · ').length * 8.2]), '<text class="lbl" x="' + (c0 + bw / 2) + '" y="' + (Math.min(p0, p1) - 7) + '" text-anchor="middle">' + esc(t.join(' · ')) + '</text>')));
        }
      });
    });
    series.forEach(function (s, si) {
      if (s.kind !== 'line' && s.kind !== 'area') return;
      var ax = s.axis === 2 && A2 ? A2 : A1, cc = colors[si], pts = [];
      (s.values || []).forEach(function (v, i) { if (v !== null && v !== undefined) pts.push([C(i), V(Math.max(ax.min, Math.min(ax.max, v)), ax), v, i]); });
      if (!pts.length) return;
      var d = 'M' + pts.map(function (q) { return q[0].toFixed(1) + ',' + q[1].toFixed(1); }).join('L');
      if (s.kind === 'area') o.push('<path d="' + d + 'L' + pts[pts.length - 1][0] + ',' + V(Math.max(ax.min, 0), ax) + 'L' + pts[0][0] + ',' + V(Math.max(ax.min, 0), ax) + 'Z" fill="' + cc + '" opacity=".25"/>');
      o.push('<path d="' + d + '" fill="none" stroke="' + cc + '" stroke-width="2.5" stroke-linejoin="round" stroke-linecap="round"/>');
      pts.forEach(function (q) {
        o.push('<circle cx="' + q[0] + '" cy="' + q[1] + '" r="4.5" fill="var(--surface)" stroke="' + cc + '" stroke-width="2"/>');
        if (s.labels && s.labels.value) {
          var lt = fmt(q[2], s.labels.format && !/general/i.test(s.labels.format) ? s.labels.format : (s.axis === 2 ? y2f : yf)), ly = q[1] - 10, lw = lt.length * 8.2;
          // lift the line's label clear of a bar label at the same spot
          for (var k = 0; k < 4; k++) {
            var hit = boxes.some(function (bx) { return Math.abs(bx[0] - q[0]) < (bx[2] + lw) / 2 + 2 && Math.abs(bx[1] - ly) < 17; });
            if (!hit) break; ly -= 18;
          }
          labels.push('<text class="lbl" x="' + q[0] + '" y="' + ly + '" text-anchor="middle">' + esc(lt) + '</text>');
        }
      });
    });
    o.push(horiz ? '<line class="base" x1="' + V(Math.max(A1.min, 0)) + '" x2="' + V(Math.max(A1.min, 0)) + '" y1="' + pt + '" y2="' + pb + '"/>' : '<line class="base" x1="' + pl + '" x2="' + pr + '" y1="' + V(Math.max(A1.min, 0)) + '" y2="' + V(Math.max(A1.min, 0)) + '"/>');
    // one hover target per category: every series' value for it
    cats.forEach(function (c, i) {
      var lines = series.map(function (s) { var v = (s.values || [])[i]; var f = (s.labels && s.labels.format && !/general/i.test(s.labels.format)) ? s.labels.format : (s.axis === 2 ? y2f : yf); return (series.length > 1 ? esc(s.name || '') + ': ' : '') + '<b>' + (v === null || v === undefined ? '—' : esc(fmt(v, f))) + '</b>'; });
      tips.push('<b>' + esc(c) + '</b><br>' + lines.join('<br>'));
      if (horiz) hits.push('<rect class="hit" x="0" y="' + (C(i) - band / 2) + '" width="' + W + '" height="' + band + '" data-tip="' + (tips.length - 1) + '"/>');
      else hits.push('<rect class="hit" x="' + (C(i) - band / 2) + '" y="' + pt + '" width="' + band + '" height="' + (pb - pt) + '" data-tip="' + (tips.length - 1) + '"/>');
    });
    var leg = series.length > 1 ? legendHtml(series.map(function (s, i) { return { n: s.name || ('Series ' + (i + 1)), c: colors[i], line: s.kind === 'line' }; })) : '';
    el.innerHTML = '<svg width="' + W + '" height="' + H + '" viewBox="0 0 ' + W + ' ' + H + '">' + o.concat(labels).concat(hits).join('') + '</svg>' + leg + '<div class="tip"></div>' + fbHtml;
    tipOn(el, tips);
  }
  var charts = [].slice.call(document.querySelectorAll('.bchart[data-chart]'));
  function drawAll() { charts.forEach(function (el) { try { draw(el); el.classList.remove('err'); } catch (e) { el.classList.add('err'); el.setAttribute('data-error', String(e)); } }); }
  var rz; window.addEventListener('resize', function () { clearTimeout(rz); rz = setTimeout(function () { fitCover(); drawAll(); }, 150); });

  // ---------- cover fills the window under the top bar
  var zoom = 1;
  function fitCover() {
    var t = document.querySelector('.top'); if (!t) return;
    var avail = body.classList.contains('present') ? window.innerHeight - 40 : window.innerHeight - t.getBoundingClientRect().height - 26;
    root.style.setProperty('--coverH', Math.max(520, avail / zoom) + 'px');
  }
  // ---------- zoom
  try { zoom = parseFloat(localStorage.getItem('itsmB.zoom')) || 1; } catch (e) { }
  function setZoom(z) {
    zoom = Math.max(0.6, Math.min(2, Math.round(z * 100) / 100));
    document.getElementById('content').style.zoom = zoom;
    document.getElementById('zVal').textContent = Math.round(zoom * 100) + '%';
    try { localStorage.setItem('itsmB.zoom', String(zoom)); } catch (e) { }
    fitCover(); clearTimeout(rz); rz = setTimeout(drawAll, 60);
  }
  document.getElementById('zIn').onclick = function () { setZoom(zoom + 0.1); };
  document.getElementById('zOut').onclick = function () { setZoom(zoom - 0.1); };
  document.getElementById('zFit').onclick = function () { setZoom(1); };
  // ---------- page menu
  var nav = document.getElementById('nav'), links = [].slice.call(nav.querySelectorAll('a'));
  function menu(o) { body.classList.toggle('menu', o); }
  document.getElementById('mOpen').onclick = function () { menu(true); };
  document.getElementById('mClose').onclick = function () { menu(false); };
  document.getElementById('scrim').onclick = function () { menu(false); };
  links.forEach(function (l) { l.addEventListener('click', function () { menu(false); }); });
  function track() {
    var k = 0, y = window.innerHeight * 0.4;
    secs.forEach(function (s, i) { if (s.getBoundingClientRect().top <= y) k = i; });
    links.forEach(function (l, i) { l.classList.toggle('on', i === k); });
    document.getElementById('mPos').textContent = (k + 1) + ' / ' + secs.length;
    return k;
  }
  window.addEventListener('scroll', track, { passive: true });
  // ---------- theme
  try { var th = localStorage.getItem('itsmB.theme'); if (th) root.dataset.theme = th; } catch (e) { }
  document.getElementById('theme').onclick = function () {
    var d = root.dataset.theme === 'dark' || (!root.dataset.theme && matchMedia('(prefers-color-scheme:dark)').matches);
    root.dataset.theme = d ? 'light' : 'dark';
    try { localStorage.setItem('itsmB.theme', root.dataset.theme); } catch (e) { }
    drawAll();
  };
  // ---------- presentation: one page at a time
  var cur = 0;
  function show(i) {
    cur = Math.max(0, Math.min(secs.length - 1, i));
    secs.forEach(function (s, k) { s.classList.toggle('cur', k === cur); });
    document.getElementById('hpos').textContent = (cur + 1) + ' / ' + secs.length;
    fitCover(); drawAll();
  }
  function start() {
    body.classList.add('present'); show(track());
    if (root.requestFullscreen) root.requestFullscreen().catch(function () { });
  }
  function stop() {
    body.classList.remove('present'); secs.forEach(function (s) { s.classList.remove('cur'); });
    if (document.fullscreenElement) document.exitFullscreen();
    fitCover(); drawAll(); secs[cur].scrollIntoView();
  }
  document.getElementById('present').onclick = start;
  document.getElementById('hp').onclick = function () { show(cur - 1); };
  document.getElementById('hn').onclick = function () { show(cur + 1); };
  document.getElementById('hx').onclick = stop;
  document.addEventListener('fullscreenchange', function () { if (!document.fullscreenElement && body.classList.contains('present')) stop(); });
  document.addEventListener('keydown', function (e) {
    if (e.key === 'Escape') menu(false);
    if (body.classList.contains('present')) {
      if (['ArrowRight', 'PageDown', ' '].indexOf(e.key) >= 0) { e.preventDefault(); show(cur + 1); }
      if (['ArrowLeft', 'PageUp'].indexOf(e.key) >= 0) { e.preventDefault(); show(cur - 1); }
      if (e.key === 'Home') show(0);
      if (e.key === 'End') show(secs.length - 1);
      if (e.key === 'Escape') stop();
      return;
    }
    if (e.ctrlKey || e.metaKey || e.altKey) return;
    if (e.key === '+' || e.key === '=') setZoom(zoom + 0.1);
    if (e.key === '-') setZoom(zoom - 0.1);
    if (e.key === 'p' || e.key === 'P') { e.preventDefault(); start(); }
  });
  window.addEventListener('beforeprint', function () { document.getElementById('content').style.zoom = 1; });
  window.addEventListener('afterprint', function () { document.getElementById('content').style.zoom = zoom; });
  setZoom(zoom); track(); drawAll();
  if (document.fonts && document.fonts.ready) document.fonts.ready.then(drawAll);
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
    $sections = [regex]::Matches($written, '<section class="(?:page|cvr)').Count
    $externalRefs = [regex]::Matches($written, '(?:src|href)="(?!data:|#)[^"]+"').Count
    if ($sections -ne $result.Pages.Count -or $externalRefs -ne 0) { $validLevel = 'WARN' }
    $nativeCharts = [regex]::Matches($written, 'class="bchart"').Count
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
