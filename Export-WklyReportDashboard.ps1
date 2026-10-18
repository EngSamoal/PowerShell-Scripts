<#
.SYNOPSIS
    Presents the "Wkly_Report" worksheet of the weekly ITSM workbook as an interactive,
    executive HTML dashboard - strictly READ-ONLY.

.DESCRIPTION
    Workflow:  Excel workbook (Wkly_Report)  ->  this script  ->  index.html dashboard

    SAFETY - the script never changes anything:
      * The original workbook is never opened. It is copied to a private temp folder and
        only the copy is opened, read-only, in a separate hidden Excel instance.
      * Links are not updated, calculation is set to Manual before the copy is opened (so
        Excel shows the values cached at the last save), macros/events are disabled, and the
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
    Where index.html and the assets folder are written. Default: C:\Temp\ITSM_Dashboard
    Each run replaces only what this script generated earlier (index.html, assets\, the log
    and the PDFs); any other files in that folder are left alone.

.PARAMETER PageMode
    How report pages are found:
      Auto       (default) Page titles first, then manual page breaks, else one page.
      Titles     Large-font page titles start a new page.
      PageBreaks Manual horizontal page breaks in Wkly_Report define the pages.
      Manual     Use -PageStartRows.

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
    Do not export the per-page "Excel view" PDFs (faster).

.PARAMETER SkipSheetPdf
    Do not export the full "Wkly_Report (Excel export).pdf".

.PARAMETER ExportDashboardPdf
    Also print the dashboard to "ITSM_Dashboard.pdf" with Microsoft Edge (headless, local).

.PARAMETER OpenWhenDone
    Open index.html in the default browser when finished.

.EXAMPLE
    .\Export-WklyReportDashboard.ps1

.EXAMPLE
    .\Export-WklyReportDashboard.ps1 -OutputDirectory 'C:\Temp\ITSM_Dashboard' -ExportDashboardPdf -OpenWhenDone

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
    [int[]]$PageStartRows,
    [double]$TitleMinFontSize = 20,
    [double]$MinPageHeightPt = 250,
    [string]$DashboardTitle = 'ITSM Executive Dashboard',
    [string]$Classification = 'CONFIDENTIAL',
    [switch]$SkipExcelView,
    [switch]$SkipSheetPdf,
    [switch]$ExportDashboardPdf,
    [switch]$OpenWhenDone
)

$ScriptBuild = '2026.10.18-r1'
$SheetName = 'Wkly_Report'          # fixed by design - no other worksheet is ever read
$Inv = [System.Globalization.CultureInfo]::InvariantCulture
$ErrRegex = '#(DIV/0!|N/A|VALUE!|REF!|NAME\?|NUM!|NULL!|SPILL!|CALC!|FIELD!|BLOCKED!|CONNECT!|BUSY!|UNKNOWN!|GETTING_DATA)'
$GeneratedMarker = '.generated-by-Export-WklyReportDashboard'

$script:LogLines = New-Object System.Collections.Generic.List[string]
$script:Issues = New-Object System.Collections.Generic.List[object]
$script:ChartIndex = 0
$script:LoggedShapeTypes = @{}

# ======================================================================================
#  Small helpers
# ======================================================================================

function Write-Log {
    param([string]$Message, [ValidateSet('INFO', 'WARN', 'ERROR', 'OK')][string]$Level = 'INFO')
    $line = '{0} [{1}] {2}' -f (Get-Date -Format 'yyyy-MM-dd HH:mm:ss'), $Level, $Message
    $script:LogLines.Add($line)
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
    $fallback = "'Segoe UI',Calibri,Arial,sans-serif"
    if ([string]::IsNullOrWhiteSpace($Name) -or $Name.StartsWith('+')) { return $fallback }
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
                $valign = if ($vAlign.ContainsKey($va)) { $vAlign[$va] } else { 'bottom' }

                $bg = $null
                $interior = $df.Interior
                if ([int]$interior.Pattern -ne -4142) { $bg = ConvertTo-CssColor $interior.Color }

                $borders = New-Object System.Collections.Generic.List[string]
                foreach ($edge in @(@(7, 'left'), @(8, 'top'), @(9, 'bottom'), @(10, 'right'))) {
                    $b = $df.Borders.Item($edge[0])
                    if ([int]$b.LineStyle -ne -4142) {
                        $w = if ($weights.ContainsKey([int]$b.Weight)) { $weights[[int]$b.Weight] } else { 1 }
                        $bc = ConvertTo-CssColor $b.Color
                        if (-not $bc) { $bc = '#000' }
                        $borders.Add(('border-{0}:{1}pt solid {2}' -f $edge[1], (Format-Num $w), $bc))
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
                        Align     = $align; VAlign = $valign; Wrap = (Test-MsoTrue $df.WrapText)
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

function Get-RunInfo {
    param($Range)
    $f = $Range.Font
    $color = $null
    try { $color = ConvertTo-CssColor $f.Fill.ForeColor.RGB } catch { }
    $under = $false
    try { $under = ([int]$f.UnderlineStyle -ne 0) } catch { }
    return [pscustomobject]@{
        Text = [string]$Range.Text; Size = [double]$f.Size; Bold = (Test-MsoTrue $f.Bold)
        Italic = (Test-MsoTrue $f.Italic); Underline = $under; Color = $color; Font = [string]$f.Name
    }
}

function Read-ShapeText {
    param($Shape)
    try { if (-not (Test-MsoTrue $Shape.TextFrame2.HasText)) { return $null } } catch { return $null }
    $tf = $Shape.TextFrame2
    $tr = $tf.TextRange
    $runs = New-Object System.Collections.Generic.List[object]
    $cnt = 0
    try { $cnt = [int]$tr.Runs().Count } catch { }
    if ($cnt -gt 0) { for ($i = 1; $i -le $cnt; $i++) { $runs.Add((Get-RunInfo $tr.Runs($i, 1))) } }
    else { $runs.Add((Get-RunInfo $tr)) }
    $aligns = New-Object System.Collections.Generic.List[string]
    $map = @{ 1 = 'left'; 2 = 'center'; 3 = 'right'; 4 = 'justify'; 5 = 'justify'; 7 = 'justify' }
    try {
        $pc = [int]$tr.Paragraphs().Count
        for ($p = 1; $p -le $pc; $p++) {
            $a = [int]$tr.Paragraphs($p, 1).ParagraphFormat.Alignment
            $aligns.Add($(if ($map.ContainsKey($a)) { $map[$a] } else { 'left' }))
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
    return [pscustomobject]@{ Runs = $runs; Aligns = $aligns; Anchor = $anchor; Margins = $m; Wrap = $wrap; Orientation = $orient }
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
            try { [void]$Shape.Chart.Export($full, 'PNG', $false) } catch { }
            if ((Test-Path -LiteralPath $full) -and (Get-Item -LiteralPath $full).Length -gt 0) {
                $Items.Add([pscustomobject]@{ Kind = 'image'; L = $L; T = $T; W = $W; H = $H; Z = $Z; Rot = $rot; Src = "assets/$file"; Alt = $title; IsChart = $true })
                if ((Get-Item -LiteralPath $full).Length -lt 1500) {
                    Add-Issue -Severity Warning -Location "Chart '$name'" -Message 'Exported chart image is very small and may be blank - compare with Excel view.' -Y $T
                }
            }
            else {
                $Items.Add([pscustomobject]@{ Kind = 'missing'; L = $L; T = $T; W = $W; H = $H; Z = $Z; Rot = $rot; Message = "Chart '$title' could not be exported - see Excel view" })
                Add-Issue -Severity Error -Location "Chart '$name'" -Message 'Excel could not export this chart as an image. A labelled placeholder is shown instead; see Excel view.' -Y $T
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
                $Items.Add([pscustomobject]@{ Kind = 'missing'; L = $L; T = $T; W = $W; H = $H; Z = $Z; Rot = $rot; Message = "Picture '$name' not reproduced - see Excel view" })
                Add-Issue -Severity Warning -Location "Picture '$name'" -Message 'Picture could not be extracted from the workbook package (linked/camera picture or .xlsb/.xls file). Placeholder shown; see Excel view.' -Y $T
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
            $text = Read-ShapeText $Shape
            if (-not $bg -and -not $border -and -not $text) { return }   # invisible helper shape
            $obj = [pscustomobject]@{
                Kind = 'shape'; L = $L; T = $T; W = $W; H = $H; Z = $Z; Rot = $rot; Name = $name
                Bg = $bg; Border = $border; Radius = $radius; Shadow = $shadow
                Runs = $null; Aligns = $null; Anchor = 'top'; Margins = $null; Wrap = $true; TextOrientation = 1
            }
            if ($text) {
                # assigned directly so single-item lists are not unrolled into scalars
                $obj.Runs = $text.Runs; $obj.Aligns = $text.Aligns; $obj.Anchor = $text.Anchor
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
        $Items.Add([pscustomobject]@{ Kind = 'missing'; L = $L; T = $T; W = $W; H = $H; Z = $Z; Rot = 0; Message = "$tn '$name' not reproduced - see Excel view" })
        Add-Issue -Severity Warning -Location "Object '$name'" -Message ("{0} cannot be reproduced in HTML. Placeholder shown; see Excel view." -f $tn) -Y $T
    }
    catch {
        Add-Issue -Severity Warning -Location "Object '$name'" -Message ("Could not be read: {0}" -f $_.Exception.Message)
    }
}

function Read-ShapeItems {
    param($Ws, [string]$AssetsDir, [hashtable]$MediaMap)
    $items = New-Object System.Collections.Generic.List[object]
    $n = [int]$Ws.Shapes.Count
    for ($k = 1; $k -le $n; $k++) {
        Write-Progress -Activity 'Reading Wkly_Report shapes and charts' -PercentComplete ([int](100 * $k / [math]::Max(1, $n)))
        $s = $Ws.Shapes.Item($k)
        $zp = $k
        try { $zp = [int]$s.ZOrderPosition } catch { }
        Read-OneShape -Shape $s -Z (10 + $zp) -Items $items -AssetsDir $AssetsDir -MediaMap $MediaMap
    }
    Write-Progress -Activity 'Reading Wkly_Report shapes and charts' -Completed
    Write-Log ("Read {0} shapes ({1} charts exported)." -f $n, $script:ChartIndex)
    return , $items
}

# ======================================================================================
#  Page detection (pure logic - no Excel calls)
# ======================================================================================

function Get-TitleCandidates {
    param($Items, [double]$ContentWidth, [double]$MinFont)
    $list = New-Object System.Collections.Generic.List[object]
    foreach ($it in $Items) {
        if ($it.Kind -ne 'cell' -and $it.Kind -ne 'shape') { continue }
        $size = Get-ItemMaxFont $it
        if ($size -lt $MinFont) { continue }
        $t = ((Get-ItemText $it) -replace '\s+', ' ').Trim()
        if ($t.Length -eq 0 -or $t.Length -gt 120) { continue }
        if ([regex]::Matches($t, '\p{L}').Count -lt 4) { continue }
        if ($t -match $ErrRegex) { continue }
        if ($it.L -gt 0.35 * $ContentWidth) { continue }
        $list.Add([pscustomobject]@{ Text = $t; Top = [double]$it.T; Item = $it })
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
            if ($null -eq $cur.Title -and ($c.Top - $cur.Y) -lt $MinHeight) { $cur.Title = $c.Text; $cur.TitleTop = $c.Top; continue }
            $ref = if ($null -ne $cur.TitleTop) { $cur.TitleTop } else { $cur.Y }
            if (($c.Top - $ref) -lt $MinHeight) { continue }
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
        $idx = $pages.Count + 1
        $title = $sorted[$i].Title
        if (-not $title) { $title = Get-PageTitle -Items $content -StartY $sy -EndY $ey -Index $idx }
        if ($title.Length -gt 80) { $title = $title.Substring(0, 77) + '...' }
        $pages.Add([pscustomobject]@{
                Index    = $idx; StartY = $sy; EndY = $ey; Title = $title
                StartRow = (Get-RowAtY $Geo.RowTops $Geo.LastRow $sy)
                EndRow   = (Get-RowAtY $Geo.RowTops $Geo.LastRow ($ey - 0.5))
                Items    = $pageItems; ExcelView = $null; IssueCount = 0
            })
    }
    Write-Log ("Page detection - {0}. Pages: {1}" -f $method, $pages.Count) 'OK'
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
    param($It, [double]$OffsetY)
    $s = 'left:{0}pt;top:{1}pt;width:{2}pt;height:{3}pt' -f (Format-Num $It.L), (Format-Num ($It.T - $OffsetY)), (Format-Num ([math]::Max($It.W, 0))), (Format-Num ([math]::Max($It.H, 0)))
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

function ConvertTo-ItemHtml {
    param($It, [double]$OffsetY)
    $box = Get-BoxStyle $It $OffsetY
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
                $alignOf = { param($k) if ($It.Aligns -and $k -lt $It.Aligns.Count) { $It.Aligns[$k] } else { 'left' } }
                [void]$sb.Append('<div class="pa" style="text-align:' + (& $alignOf $pi) + '">')
                foreach ($run in $It.Runs) {
                    $segments = ([string]$run.Text) -split "`r"
                    for ($s = 0; $s -lt $segments.Count; $s++) {
                        if ($s -gt 0) {
                            $pi++
                            [void]$sb.Append('</div><div class="pa" style="text-align:' + (& $alignOf $pi) + '">')
                        }
                        if ($segments[$s].Length) { [void]$sb.Append('<span style="' + (ConvertTo-RunStyle $run) + '">' + (ConvertTo-HtmlText $segments[$s]) + '</span>') }
                        elseif ($s -gt 0 -or $segments.Count -eq 1) { [void]$sb.Append('<span style="' + (ConvertTo-RunStyle $run) + '">&#8203;</span>') }
                    }
                }
                [void]$sb.Append('</div>')
                $inner = '<div class="tx" style="{0}">{1}</div>' -f $tx, $sb.ToString()
            }
            $cls = if ($It.Wrap) { 'it shp' } else { 'it shp nowrap' }
            return '<div class="{0}" data-name="{1}" style="{2}">{3}</div>' -f $cls, (ConvertTo-Attr $It.Name), $st, $inner
        }
        'line' {
            $w = [math]::Max($It.W, 0.01); $h = [math]::Max($It.H, 0.01)
            $x1 = if ($It.FlipH) { $w } else { 0 }; $x2 = $w - $x1
            $y1 = if ($It.FlipV) { $h } else { 0 }; $y2 = $h - $y1
            $st = 'left:{0}pt;top:{1}pt;width:{2}pt;height:{3}pt;overflow:visible' -f (Format-Num $It.L), (Format-Num ($It.T - $OffsetY)), (Format-Num $w), (Format-Num $h)
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

function New-DashboardSite {
    param($Pages, [string]$OutDir, $Meta, $Issues, [double]$ContentWidth)

    $assets = Join-Path $OutDir 'assets'
    [System.IO.File]::WriteAllText((Join-Path $assets 'dashboard.css'), $DashboardCss, (New-Object System.Text.UTF8Encoding($false)))
    [System.IO.File]::WriteAllText((Join-Path $assets 'dashboard.js'), $DashboardJs, (New-Object System.Text.UTF8Encoding($false)))

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
    & $add '<header class="topbar"><div class="brand"><div class="mark">ITSM</div>EXECUTIVE DASHBOARD</div>'
    & $add ('<div class="tb-title"><b>{0}</b><span>Source: {1} &middot; Wkly_Report (read-only snapshot)</span></div>' -f (& $enc $DashboardTitle), (& $enc $Meta.FileName))
    if ($Meta.ReportingPeriod) { & $add ('<span class="chip" title="Reporting Period (from Wkly_Report)">{0}</span>' -f (& $enc $Meta.ReportingPeriod)) }
    if ($Meta.ReportDate) { & $add ('<span class="chip" title="Report Date (from Wkly_Report)">Report Date: {0}</span>' -f (& $enc $Meta.ReportDate)) }
    & $add ('<button class="btn warn" id="btnLog">Data check <span class="cnt">{0}</span></button>' -f $issueTotal)
    & $add '<button class="btn" id="btnMarkers" title="Show/hide the markers on Excel errors and empty formula results">Markers: on</button>'
    & $add '<button class="btn" id="btnPresent" title="Presentation mode (P)">&#x26F6; Present</button>'
    & $add '<button class="btn" onclick="window.print()">Print / PDF</button>'
    if ($Classification) { & $add ('<span class="conf">{0}</span>' -f (& $enc $Classification)) }
    & $add "</header>`n<div class=`"layout`">`n"

    # ---- sidebar
    & $add '<aside class="sidebar"><h4>REPORT PAGES</h4><nav><ol id="nav">'
    foreach ($p in $Pages) {
        $flag = if ($p.IssueCount -gt 0) { '<span class="flag" title="Values needing attention on this page">&#9888; {0}</span>' -f $p.IssueCount } else { '' }
        & $add ('<li><a href="#p{0}"><span class="n">{0}</span><span class="t">{1}</span>{2}</a></li>' -f $p.Index, (& $enc $p.Title), $flag)
    }
    & $add '</ol></nav><div class="src">'
    & $add ('<b>Workbook:</b> {0}<br><b>Worksheet:</b> Wkly_Report<br><b>Workbook saved:</b> {1}<br><b>Snapshot taken:</b> {2}<br><b>Pages detected:</b> {3}<br><b>Generator build:</b> {4}' -f (& $enc $Meta.FileName), (& $enc $Meta.Saved), (& $enc $Meta.Snapshot), $Pages.Count, $ScriptBuild)
    & $add "</div></aside>`n<main id=`"main`">`n"

    # ---- pages
    $n = $Pages.Count
    foreach ($p in $Pages) {
        $hPt = $p.EndY - $p.StartY
        $pz = [math]::Min(1.0, [math]::Min($printW / $ContentWidth, $printH / $hPt))
        & $add ('<section class="page" id="p{0}" data-title="{1}">' -f $p.Index, (ConvertTo-Attr $p.Title))
        & $add ('<div class="pbar"><div class="pt"><span class="pn">{0:D2}</span><h2>{1}</h2><span class="rows">rows {2}&ndash;{3}</span></div><div class="tools">' -f $p.Index, (& $enc $p.Title), $p.StartRow, $p.EndRow)
        if ($p.IssueCount -gt 0) { & $add ('<span class="badge" title="See Data check">&#9888; {0} to check</span>' -f $p.IssueCount) }
        if ($p.ExcelView) { & $add ('<button class="btn xbtn" data-xl="{0}">Excel view</button>' -f $p.ExcelView) }
        else { & $add '<button class="btn xbtn" disabled title="Excel view was not exported for this run">Excel view</button>' }
        & $add ('<span class="pnum">{0} / {1}</span></div></div>' -f $p.Index, $n)
        & $add ('<div class="pbody"><div class="cvwrap" style="--w:{0}pt;--h:{1}pt;--pz:{2}">' -f (Format-Num $ContentWidth), (Format-Num $hPt), (Format-Num $pz))
        & $add ('<div class="cv" data-w="{0}" data-h="{1}" style="width:{2}pt;height:{3}pt">' -f (Format-Num ($ContentWidth * $pt2px)), (Format-Num ($hPt * $pt2px)), (Format-Num $ContentWidth), (Format-Num $hPt))
        foreach ($it in $p.Items) { & $add (ConvertTo-ItemHtml $it $p.StartY); & $add "`n" }
        & $add '</div></div><div class="xlview"><iframe title="Excel rendering of this page"></iframe></div></div>'
        & $add ('<div class="pf"><span>{0}</span><span>Wkly_Report &middot; page {1} of {2} &middot; rows {3}&ndash;{4}</span><span class="c">{5}</span></div></section>' -f (& $enc $DashboardTitle), $p.Index, $n, $p.StartRow, $p.EndRow, (& $enc $Classification))
        & $add "`n"
    }
    & $add "</main>`n</div>`n"

    # ---- data-check drawer
    & $add '<div class="present-hud"><span id="hudPos"></span><button class="btn" id="hudPrev">&#9664;</button><button class="btn" id="hudNext">&#9654;</button><button class="btn" id="hudExit">Exit (Esc)</button></div>'
    & $add '<aside class="drawer" id="drawer"><header><b>Data check &amp; extraction log</b><button class="btn" id="btnClose">&#10005;</button></header><div class="body">'
    & $add '<p class="lead">Everything the dashboard could not reproduce exactly is listed here, so nothing is silently replaced. Values themselves are never changed.</p>'
    & $add '<h5>SOURCE</h5><ul>'
    & $add ('<li>{0}</li>' -f (& $enc $Meta.SourcePath))
    & $add '<li>A temporary copy was opened read-only in a separate Excel instance: links not updated, calculation Manual (values as last saved), macros disabled, closed without saving, copy deleted.</li>'
    $hashState = if ($Meta.HashBefore -eq $Meta.HashAfter) { '<span class="ok">unchanged</span>' } else { '<span class="er">CHANGED during the run - the file was modified by someone else (this script never writes to it)</span>' }
    & $add ('<li>Original file SHA-256 before / after: {0}<br><code>{1}</code></li>' -f $hashState, $Meta.HashBefore)
    & $add '</ul><h5>CAPTURED</h5><ul>'
    & $add ('<li><span class="ok">&#10004;</span> {0}</li>' -f (& $enc $Meta.PageMethod))
    & $add ('<li><span class="ok">&#10004;</span> {0} non-empty cells, {1} shapes/text boxes, {2} charts, {3} pictures</li>' -f $Meta.CellCount, $Meta.ShapeCount, $Meta.ChartCount, $Meta.PictureCount)
    & $add ('<li><span class="ok">&#10004;</span> Excel view: {0}</li>' -f (& $enc $Meta.ExcelViewNote))
    & $add '</ul><h5>NEEDS ATTENTION</h5>'
    $shown = @($Issues | Sort-Object @{ Expression = { if ($_.Page) { $_.Page } else { 0 } } }, @{ Expression = { switch ($_.Severity) { 'Error' { 0 } 'Warning' { 1 } default { 2 } } } })
    if ($shown.Count -eq 0) { & $add '<p class="ok">Nothing to report.</p>' }
    else {
        & $add '<table class="iss"><thead><tr><th>Page</th><th>Where</th><th>What</th></tr></thead><tbody>'
        foreach ($iss in $shown) {
            $cls = switch ($iss.Severity) { 'Error' { 'er' } 'Warning' { 'wn' } default { 'in' } }
            $pg = if ($iss.Page) { '<a href="#p{0}">{0}</a>' -f $iss.Page } else { '&ndash;' }
            & $add ('<tr class="{0}"><td>{1}</td><td>{2}</td><td>{3}</td></tr>' -f $cls, $pg, (& $enc $iss.Location), (& $enc $iss.Message))
        }
        & $add '</tbody></table>'
    }
    & $add ('<h5>LOG</h5><p>Full log: <code>extraction-log.txt</code> in the output folder. Build {0}.</p>' -f $ScriptBuild)
    & $add "</div></aside>`n<script src=`"assets/dashboard.js`"></script>`n</body>`n</html>`n"

    $index = Join-Path $OutDir 'index.html'
    [System.IO.File]::WriteAllText($index, $sb.ToString(), (New-Object System.Text.UTF8Encoding($false)))
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
.brand{width:270px;flex:none;display:flex;align-items:center;gap:10px;padding-left:20px;font-weight:700;letter-spacing:.6px;font-size:14px}
.brand .mark{width:34px;height:34px;border-radius:7px;background:linear-gradient(135deg,var(--gold-2),#8c6d10);display:grid;place-items:center;color:var(--navy);font-size:12px;font-weight:800}
.tb-title{flex:1;min-width:0}
.tb-title b{display:block;font-size:15px;white-space:nowrap;overflow:hidden;text-overflow:ellipsis}
.tb-title span{font-size:12px;color:#c3cbe4}
.chip{border:1px solid rgba(201,162,39,.65);color:var(--gold-2);border-radius:999px;padding:4px 12px;font-size:12px;white-space:nowrap;max-width:340px;overflow:hidden;text-overflow:ellipsis}
.conf{background:var(--red);color:#fff;border-radius:4px;padding:4px 10px;font-size:11px;font-weight:700;letter-spacing:1.5px}
.btn{background:rgba(255,255,255,.08);border:1px solid rgba(255,255,255,.22);color:#fff;border-radius:6px;padding:7px 12px;font-size:13px;cursor:pointer;font-family:inherit;white-space:nowrap}
.btn:hover{background:rgba(255,255,255,.18)}
.btn[disabled]{opacity:.45;cursor:not-allowed}
.btn.warn{border-color:#f0b4ad}
.btn .cnt{background:var(--red);border-radius:999px;padding:0 6px;margin-left:6px;font-size:11px}
.layout{display:flex}
.sidebar{position:sticky;top:65px;height:calc(100vh - 65px);width:270px;flex:none;background:#fff;border-right:1px solid var(--line);display:flex;flex-direction:column}
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
main{flex:1;min-width:0;padding:22px 26px}
.page{background:#fff;border-radius:12px;box-shadow:0 1px 3px rgba(16,24,40,.08),0 10px 28px rgba(16,24,40,.07);margin:0 auto 28px;overflow:hidden;scroll-margin-top:84px;display:flex;flex-direction:column}
.pbar{display:flex;justify-content:space-between;align-items:center;gap:12px;padding:10px 18px;background:linear-gradient(90deg,var(--navy),var(--navy-2));border-left:6px solid var(--gold);color:#fff}
.pbar .pt{display:flex;align-items:center;gap:12px;min-width:0}
.pbar .pn{background:var(--gold);color:var(--navy);font-weight:800;border-radius:6px;padding:3px 8px;font-size:13px}
.pbar h2{margin:0;font-size:17px;white-space:nowrap;overflow:hidden;text-overflow:ellipsis}
.pbar .rows{font-size:11px;color:#aab3cf;white-space:nowrap}
.pbar .tools{display:flex;align-items:center;gap:10px}
.pbar .badge{background:#fdecea;color:#b42318;border-radius:999px;padding:3px 10px;font-size:12px;font-weight:700;white-space:nowrap}
.pnum{font-size:12px;color:#aab3cf;white-space:nowrap}
.xbtn{font-size:12px;padding:5px 10px}
.pbody{padding:14px;background:#f6f7fa;flex:1}
.cvwrap{position:relative;overflow:hidden;margin:0 auto;border-radius:6px;box-shadow:0 0 0 1px var(--line)}
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
.present-hud{display:none}
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

  function fit() {
    var present = body.classList.contains('present');
    pages.forEach(function (p) {
      var wrap = p.querySelector('.cvwrap'), cv = p.querySelector('.cv');
      if (!wrap || !cv) return;
      var W = parseFloat(cv.dataset.w), H = parseFloat(cv.dataset.h);
      var avail = p.querySelector('.pbody').clientWidth - 28;
      if (present) avail = window.innerWidth - 12;
      var s = avail / W;
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
  window.addEventListener('resize', fit);
  fit();
  if (document.fonts && document.fonts.ready) document.fonts.ready.then(fit);

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

  // Excel view (PDF rendered by Excel itself), loaded on first use
  [].slice.call(document.querySelectorAll('.xbtn[data-xl]')).forEach(function (b) {
    b.addEventListener('click', function () {
      var p = b.closest('.page'), f = p.querySelector('.xlview iframe');
      if (!f.getAttribute('src')) f.setAttribute('src', b.getAttribute('data-xl') + '#view=FitH');
      p.classList.toggle('xl');
      b.textContent = p.classList.contains('xl') ? 'Dashboard view' : 'Excel view';
    });
  });

  // data check drawer
  var drawer = document.getElementById('drawer');
  document.getElementById('btnLog').addEventListener('click', function () { drawer.classList.add('open'); });
  document.getElementById('btnClose').addEventListener('click', function () { drawer.classList.remove('open'); });
  [].slice.call(drawer.querySelectorAll('a[href^="#p"]')).forEach(function (a) { a.addEventListener('click', function () { drawer.classList.remove('open'); }); });

  // markers on/off
  var mk = document.getElementById('btnMarkers');
  mk.addEventListener('click', function () {
    body.classList.toggle('no-markers');
    mk.textContent = body.classList.contains('no-markers') ? 'Markers: off' : 'Markers: on';
  });

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
    if (e.key === 'Escape') drawer.classList.remove('open');
    if (!body.classList.contains('present')) {
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

function Export-DashboardPdf {
    param([string]$IndexPath, [string]$PdfPath)
    $edge = @(
        (Join-Path ${env:ProgramFiles(x86)} 'Microsoft\Edge\Application\msedge.exe'),
        (Join-Path $env:ProgramFiles 'Microsoft\Edge\Application\msedge.exe')
    ) | Where-Object { $_ -and (Test-Path -LiteralPath $_) } | Select-Object -First 1
    if (-not $edge) { Add-Issue -Severity Warning -Location 'Dashboard PDF' -Message 'Microsoft Edge was not found; dashboard PDF not created. Use Print / PDF in the browser instead.'; return }
    $profileDir = Join-Path $env:TEMP ('ITSMDash_Edge_' + [guid]::NewGuid().ToString('N'))
    $uri = ([System.Uri](Resolve-Path -LiteralPath $IndexPath).ProviderPath).AbsoluteUri
    $argList = @('--headless=new', '--disable-gpu', '--no-first-run', '--no-default-browser-check',
        "--user-data-dir=`"$profileDir`"", '--no-pdf-header-footer', '--virtual-time-budget=10000',
        "--print-to-pdf=`"$PdfPath`"", $uri)
    try {
        if (Test-Path -LiteralPath $PdfPath) { Remove-Item -LiteralPath $PdfPath -Force }
        $proc = Start-Process -FilePath $edge -ArgumentList $argList -PassThru -WindowStyle Hidden
        if (-not $proc.WaitForExit(120000)) { try { $proc.Kill() } catch { } }
        if (Test-Path -LiteralPath $PdfPath) { Write-Log "Dashboard PDF written: $PdfPath" 'OK' }
        else { Add-Issue -Severity Warning -Location 'Dashboard PDF' -Message 'Edge did not produce the PDF. Use Print / PDF in the browser instead.' }
    }
    catch { Add-Issue -Severity Warning -Location 'Dashboard PDF' -Message ("Edge PDF export failed: {0}" -f $_.Exception.Message) }
    finally { Remove-Item -LiteralPath $profileDir -Recurse -Force -ErrorAction SilentlyContinue }
}

# ======================================================================================
#  Main
# ======================================================================================

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
    $assets = Join-Path $outDir 'assets'
    if (Test-Path -LiteralPath $assets) {
        if (Test-Path -LiteralPath (Join-Path $assets $GeneratedMarker)) { Remove-Item -LiteralPath $assets -Recurse -Force }
        else { throw "'$assets' exists but was not created by this script. Choose another -OutputDirectory." }
    }
    New-Item -ItemType Directory -Path $assets -Force | Out-Null
    New-Item -ItemType Directory -Path (Join-Path $assets 'excel-view') -Force | Out-Null
    Set-Content -LiteralPath (Join-Path $assets $GeneratedMarker) -Value "Generated by Export-WklyReportDashboard.ps1 $ScriptBuild - safe to delete." -Encoding ASCII

    # ---- hash + private temp copy (the original is never opened)
    $hashBefore = (Get-FileHash -LiteralPath $source -Algorithm SHA256).Hash
    Write-Log "Source: $source (SHA-256 $hashBefore)"
    $tempDir = Join-Path $env:TEMP ('ITSMDash_' + [guid]::NewGuid().ToString('N'))
    New-Item -ItemType Directory -Path $tempDir -Force | Out-Null
    $tempCopy = Join-Path $tempDir ('snapshot' + $ext)
    try { Copy-Item -LiteralPath $source -Destination $tempCopy -Force }
    catch { throw "Could not copy the workbook (is it locked by another program?): $($_.Exception.Message)" }

    $mediaMap = @{}
    if ($ext -in '.xlsx', '.xlsm') {
        try { $mediaMap = Get-PictureMediaMap -PackagePath $tempCopy -Sheet $SheetName -AssetsDir $assets }
        catch { Add-Issue -Severity Warning -Location 'Pictures' -Message ("Pictures could not be read from the workbook package: {0}" -f $_.Exception.Message) }
    }

    $excel = $null; $wb = $null; $blank = $null; $excelPid = 0
    $items = New-Object System.Collections.Generic.List[object]
    $result = $null; $geo = $null; $contentWidth = 0; $cellCount = 0; $shapeCount = 0
    $excelViewNote = 'skipped (-SkipExcelView)'
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
        try { $excel.Calculation = -4135 } catch { Write-Log 'Could not switch calculation to Manual.' 'WARN' }   # xlCalculationManual

        $m = [Type]::Missing
        Write-Log 'Opening the temporary copy read-only (links not updated)...'
        $wb = $excel.Workbooks.Open($tempCopy, 0, $true, $m, $m, $m, $true, $m, $m, $false, $false, $m, $false)

        $ws = $null
        try { $ws = $wb.Worksheets.Item($SheetName) }
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
        foreach ($it in (Read-FillItems -Ws $ws -Geo $geo -R1 ([int]$used.Row) -C1 ([int]$used.Column) -R2 ([int]$used.Row + [int]$used.Rows.Count - 1) -C2 ([int]$used.Column + [int]$used.Columns.Count - 1))) { $items.Add($it) }
        Write-Log 'Reading cells (displayed text and formatting)...'
        $cellItems = Read-CellItems -Ws $ws -Geo $geo
        foreach ($it in $cellItems) { $items.Add($it) }
        $cellCount = @($cellItems | Where-Object { $_.Kind -eq 'cell' }).Count
        foreach ($b in @($cellItems | Where-Object { $_.Kind -eq 'blank' })) {
            Add-Issue -Severity Info -Location "Wkly_Report!$($b.Address)" -Message 'Formula returns an empty value - shown as blank (dashed marker), not as zero.' -Y $b.T
        }
        Write-Log 'Reading shapes, text boxes, charts and pictures...'
        $shapeItems = Read-ShapeItems -Ws $ws -AssetsDir $assets -MediaMap $mediaMap
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

        $result = Find-ReportPages -Items $items -Geo $geo -ContentWidth $contentWidth -ContentBottom $contentBottom `
            -Mode $PageMode -StartRows $PageStartRows -BreakRows $breakRows -MinFont $TitleMinFontSize -MinHeight $MinPageHeightPt

        $lastColLetter = ConvertTo-ColumnLetter ([math]::Max(1, $lastCol))
        if (-not $SkipExcelView) {
            $ok = 0
            foreach ($p in $result.Pages) {
                $pdf = Join-Path $assets ('excel-view\page-{0:D2}.pdf' -f $p.Index)
                try {
                    $rng = $ws.Range(('A{0}:{1}{2}' -f $p.StartRow, $lastColLetter, $p.EndRow))
                    $rng.ExportAsFixedFormat(0, $pdf, 0, $true, $true)
                    if (Test-Path -LiteralPath $pdf) { $p.ExcelView = 'assets/excel-view/page-{0:D2}.pdf' -f $p.Index; $ok++ }
                }
                catch { Add-Issue -Severity Warning -Location ("Page {0}" -f $p.Index) -Message ("Excel view PDF could not be exported: {0}" -f $_.Exception.Message) -Y $p.StartY }
            }
            $excelViewNote = '{0} of {1} pages exported as PDF by Excel' -f $ok, $result.Pages.Count
        }
        if (-not $SkipSheetPdf) {
            $sheetPdf = Join-Path $outDir 'Wkly_Report (Excel export).pdf'
            try { $ws.ExportAsFixedFormat(0, $sheetPdf, 0, $true, $false); Write-Log "Excel PDF of Wkly_Report written: $sheetPdf" 'OK' }
            catch { Add-Issue -Severity Warning -Location 'Wkly_Report PDF' -Message ("Excel could not export the worksheet PDF: {0}" -f $_.Exception.Message) }
        }
    }
    finally {
        Write-Log 'Closing the temporary copy without saving and closing the private Excel instance...'
        if ($wb) { try { $wb.Close($false) } catch { } }
        if ($blank) { try { $blank.Close($false) } catch { } }
        if ($excel) { try { $excel.Quit() } catch { } }
        foreach ($o in @($ws, $used, $wb, $blank, $excel)) {
            if ($null -ne $o -and [System.Runtime.InteropServices.Marshal]::IsComObject($o)) { try { [void][System.Runtime.InteropServices.Marshal]::ReleaseComObject($o) } catch { } }
        }
        $ws = $null; $used = $null; $wb = $null; $blank = $null; $excel = $null
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

    $hashAfter = (Get-FileHash -LiteralPath $source -Algorithm SHA256).Hash
    if ($hashAfter -eq $hashBefore) { Write-Log 'Original workbook unchanged (SHA-256 identical before and after).' 'OK' }
    else { Add-Issue -Severity Error -Location 'Source workbook' -Message 'The original file changed during the run (saved by someone else?). This script never writes to it. Re-run to get a consistent snapshot.' }

    $meta = [pscustomobject]@{
        FileName        = $sourceInfo.Name
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
        ReportingPeriod = (Find-LabelValue -Items $items -Label 'Reporting Period')
        ReportDate      = (Find-LabelValue -Items $items -Label 'Report Date')
    }
    $index = New-DashboardSite -Pages $result.Pages -OutDir $outDir -Meta $meta -Issues $script:Issues -ContentWidth $contentWidth
    Write-Log "Dashboard written: $index" 'OK'

    if ($ExportDashboardPdf) { Export-DashboardPdf -IndexPath $index -PdfPath (Join-Path $outDir 'ITSM_Dashboard.pdf') }

    $errs = @($script:Issues | Where-Object { $_.Severity -eq 'Error' }).Count
    $warns = @($script:Issues | Where-Object { $_.Severity -eq 'Warning' }).Count
    Write-Log ("Done: {0} page(s), {1} chart(s), {2} Excel error value(s)/error(s), {3} warning(s). See 'Data check' in the dashboard." -f $result.Pages.Count, $script:ChartIndex, $errs, $warns) 'OK'
    if ($OpenWhenDone) { Start-Process -FilePath $index }
    return $outDir
}

# Run only when executed (dot-sourcing just loads the functions, e.g. for testing).
if ($MyInvocation.InvocationName -ne '.') {
    $outForLog = $null
    try { $outForLog = Invoke-Main }
    catch {
        Write-Log $_.Exception.Message 'ERROR'
        Write-Log 'No dashboard was produced. The original workbook was not modified.' 'ERROR'
        $exitCode = 1
    }
    finally {
        $logDir = if ($outForLog) { $outForLog } elseif (Test-Path -LiteralPath $OutputDirectory) { $OutputDirectory } else { $env:TEMP }
        try { [System.IO.File]::WriteAllLines((Join-Path $logDir 'extraction-log.txt'), $script:LogLines, (New-Object System.Text.UTF8Encoding($false))) } catch { }
    }
    if ($exitCode) { exit $exitCode }
}
