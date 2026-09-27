# Copy-MasterStyle.ps1
# Gives Patching.xlsx the same look as Master.xlsx: tab colours, title rows, header row,
# blue/white row banding, fonts, borders, row heights, frozen header, filters, column widths.
# Tabs are matched by position (1st -> 1st, 2nd -> 2nd ...; extra tabs reuse the Master tabs in order).
# Cell contents, header names and tab names are NOT changed.
# Safe: both files are only read; the result is saved as a NEW file. Needs Microsoft Excel.
# Run:  .\Copy-MasterStyle.ps1
param(
    [string]$Master   = 'C:\temp\Master.xlsx',            # <-- style source
    [string]$Patching = 'C:\temp\Patching.xlsx',          # <-- file to restyle
    [string]$Output   = 'C:\temp\Patching_styled.xlsx'    # <-- result (new file)
)
$ErrorActionPreference = 'Stop'

function Find-HeaderRow($ws) {
    # The table header row if there is a table, otherwise the row (top 10) with the most filled cells
    if ($ws.ListObjects.Count -gt 0) { return $ws.ListObjects.Item(1).HeaderRowRange.Row }
    $used = $ws.UsedRange
    $best = 0; $bestRow = 0
    for ($r = 1; $r -le [Math]::Min(10, $used.Row + $used.Rows.Count - 1); $r++) {
        $cnt = $excel.WorksheetFunction.CountA($ws.Rows.Item($r))
        if ($cnt -gt $best) { $best = $cnt; $bestRow = $r }
    }
    return $bestRow
}

function Get-CellStyle($cell) {
    # What the cell really looks like (including table-style colours)
    $d = $cell.DisplayFormat
    @{
        Fill       = $d.Interior.Color
        NoFill     = ($d.Interior.ColorIndex -eq -4142)
        FontName   = $d.Font.Name
        FontSize   = $d.Font.Size
        FontColor  = $d.Font.Color
        Bold       = $d.Font.Bold
        Italic     = $d.Font.Italic
        HAlign     = $cell.HorizontalAlignment
        VAlign     = $cell.VerticalAlignment
        RowHeight  = $cell.RowHeight
        Border     = $d.Borders.Item(9).Color       # bottom edge
        BorderLine = $d.Borders.Item(9).LineStyle
    }
}

function Set-RangeStyle($rng, $st, [switch]$Fill, [switch]$Align) {
    if ($Fill) { if ($st.NoFill) { $rng.Interior.Pattern = -4142 } else { $rng.Interior.Color = $st.Fill } }
    $rng.Font.Name = $st.FontName
    $rng.Font.Size = $st.FontSize
    $rng.Font.Color = $st.FontColor
    $rng.Font.Bold = $st.Bold
    $rng.Font.Italic = $st.Italic
    if ($Align) { $rng.HorizontalAlignment = $st.HAlign }
    $rng.VerticalAlignment = $st.VAlign
    $rng.RowHeight = $st.RowHeight
}

foreach ($p in $Master, $Patching) { if (-not (Test-Path $p)) { Write-Host "Not found: $p" -ForegroundColor Red; return } }
$pIn = (Resolve-Path $Patching).Path; $out = [IO.Path]::GetFullPath($Output)
if ($pIn -eq $out) { Write-Host 'Output must be a different file than Patching.' -ForegroundColor Red; return }

$excel = $null; $wbM = $null; $wbP = $null
try {
    $excel = New-Object -ComObject Excel.Application
    $excel.Visible = $false
    $excel.DisplayAlerts = $false
    $excel.ScreenUpdating = $false

    # ---------- 1. Read the Master look (per tab) ----------
    $wbM = $excel.Workbooks.Open((Resolve-Path $Master).Path, 0, $true)
    $styles = @()
    foreach ($ws in $wbM.Worksheets) {
        $h = Find-HeaderRow $ws
        if ($h -lt 1) { continue }
        $c = 2; if (-not $ws.Cells.Item($h, $c).Value2) { $c = 1 }          # a normal column (not the "#" one)
        $styles += @{
            Tab     = $ws.Name
            TabColor = $ws.Tab.Color
            HdrRow  = $h
            Rows    = $(if ($ws.ListObjects.Count -gt 0 -and $ws.ListObjects.Item(1).DataBodyRange) { $ws.ListObjects.Item(1).DataBodyRange.Rows.Count } else { $ws.UsedRange.Rows.Count - $h })
            Title   = $(if ($h -ge 2) { Get-CellStyle $ws.Cells.Item(1, 1) } else { $null })
            Info    = $(if ($h -ge 3) { Get-CellStyle $ws.Cells.Item(2, 1) } else { $null })
            Header  = Get-CellStyle $ws.Cells.Item($h, $c)
            Odd     = Get-CellStyle $ws.Cells.Item($h + 1, $c)
            Even    = Get-CellStyle $ws.Cells.Item($h + 2, $c)
        }
    }
    $wbM.Close($false); $wbM = $null
    if ($styles.Count -eq 0) { throw 'No header row found in Master.' }
    # Header / row colours come from a Master tab with at least 2 data rows (so both band colours exist);
    # tab colour and title colours come from the Master tab in the same position.
    $base = @($styles | Where-Object { $_.Rows -ge 2 })[0]
    if (-not $base) { $base = $styles[0] }

    # ---------- 2. Apply it to Patching ----------
    $wbP = $excel.Workbooks.Open($pIn)
    $i = 0
    foreach ($ws in $wbP.Worksheets) {
        $st = $styles[$i % $styles.Count]; $i++
        $st = @{ Tab = $st.Tab; TabColor = $st.TabColor; Title = $st.Title; Info = $st.Info; Header = $base.Header; Odd = $base.Odd; Even = $base.Even }
        $h = Find-HeaderRow $ws
        $used = $ws.UsedRange
        $lastRow = $used.Row + $used.Rows.Count - 1
        $lastCol = $used.Column + $used.Columns.Count - 1
        if ($h -lt 1 -or $lastCol -lt 1) { Write-Host ("{0,-24} empty - skipped" -f $ws.Name) -ForegroundColor Yellow; continue }

        if ($st.TabColor -is [int] -or $st.TabColor -is [double]) { $ws.Tab.Color = $st.TabColor } else { $ws.Tab.ColorIndex = -4142 }
        $ws.Cells.Font.Name = $st.Header.FontName

        # Rows above the header: first one like the Master title, second like the Master info line
        if ($h -ge 2 -and $st.Title) {
            $t = $ws.Range($ws.Cells.Item(1, 1), $ws.Cells.Item(1, $lastCol))
            Set-RangeStyle $t $st.Title -Fill
            if ($excel.WorksheetFunction.CountA($t) -eq 1 -and $ws.Cells.Item(1, 1).Value2) { $t.HorizontalAlignment = 7 }   # centre across, like Master
        }
        if ($h -ge 3 -and $st.Info) { Set-RangeStyle ($ws.Range($ws.Cells.Item(2, 1), $ws.Cells.Item(2, $lastCol))) $st.Info -Fill }

        # Header row
        $hdr = $ws.Range($ws.Cells.Item($h, 1), $ws.Cells.Item($h, $lastCol))
        Set-RangeStyle $hdr $st.Header -Fill -Align
        $hdr.WrapText = $false

        # Data rows: remove old colours / highlight rules, then Master banding
        if ($lastRow -gt $h) {
            $data = $ws.Range($ws.Cells.Item($h + 1, 1), $ws.Cells.Item($lastRow, $lastCol))
            $data.FormatConditions.Delete()
            if ($ws.ListObjects.Count -gt 0) { $ws.ListObjects.Item(1).TableStyle = ''; $ws.ListObjects.Item(1).ShowTableStyleRowStripes = $false }
            Set-RangeStyle $data $st.Even -Fill
            $band = $data.FormatConditions.Add(2, 0, "=MOD(ROW()-$h,2)=1")      # odd data rows
            if ($st.Odd.NoFill) { $band.Interior.Pattern = -4142 } else { $band.Interior.Color = $st.Odd.Fill }
            $band.StopIfTrue = $false
        }

        # Borders, filter, widths, frozen header, no gridlines
        $all = $ws.Range($ws.Cells.Item($h, 1), $ws.Cells.Item([Math]::Max($lastRow, $h), $lastCol))
        if ($st.Odd.BorderLine -ne -4142) { $all.Borders.LineStyle = 1; $all.Borders.Weight = 2; $all.Borders.Color = $st.Odd.Border }
        if ($ws.ListObjects.Count -eq 0 -and -not $ws.AutoFilterMode) { [void]$all.AutoFilter() }
        [void]$all.Columns.AutoFit()
        for ($c = 1; $c -le $lastCol; $c++) {
            $col = $ws.Columns.Item($c)
            if ($col.ColumnWidth -lt 8) { $col.ColumnWidth = 8 }
            if ($col.ColumnWidth -gt 60) { $col.ColumnWidth = 60 }
            $col.ColumnWidth = $col.ColumnWidth + 2                             # room for the filter button
        }
        $ws.Activate()
        $excel.ActiveWindow.FreezePanes = $false
        $excel.ActiveWindow.SplitColumn = 0
        $excel.ActiveWindow.SplitRow = $h
        $excel.ActiveWindow.FreezePanes = $true
        $excel.ActiveWindow.DisplayGridlines = $false
        Write-Host ("{0,-24} styled like Master tab '{1}' (header row {2}, {3} data rows)" -f $ws.Name, $st.Tab, $h, [Math]::Max(0, $lastRow - $h)) -ForegroundColor Green
    }
    $wbP.Worksheets.Item(1).Activate()
    $excel.ScreenUpdating = $true
    if (Test-Path $out) { Remove-Item $out -Force }
    $wbP.SaveAs($out, 51)
    Write-Host "`nSaved: $out" -ForegroundColor Green
}
catch { Write-Host "FAILED: $($_.Exception.Message)  (input files not changed)" -ForegroundColor Red }
finally {
    foreach ($w in @($wbM, $wbP)) { if ($w) { try { $w.Close($false) } catch { } } }
    if ($excel) { $excel.Quit() }
    foreach ($o in @($wbM, $wbP, $excel)) { if ($o) { [void][Runtime.InteropServices.Marshal]::ReleaseComObject($o) } }
    [GC]::Collect(); [GC]::WaitForPendingFinalizers()
}
