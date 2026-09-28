# Add-SiteSummary.ps1
# For the consolidated master sheet (main VM list + "OT - VM Inventory", "Test - VM Inventory",
# "Template - VM Inventory" banners, each followed by its own header row and VMs):
#   * The count line in row 2 keeps its look but becomes LIVE (Excel formula, updates by itself):
#       474 VMs  |  SixFlags 85  |  ...  |  OT 4  |  Test 5  |  Template 3  |  No site 0
#     - site counts   = all rows with that value in the Site column (every section)
#     - OT/Test/Template = VMs listed under that banner
#   * Site names shown in their site colour; OT / Test / Template sections get their own two-tone rows.
#   * All sections become ONE table, so a filter on any column works on the whole sheet.
#     Merged banners are unmerged (tables cannot hold merged cells) and keep their look:
#     same colours, text centred across the row.
# Only row 2 (the old count line) is replaced; all other contents are not changed.
# Safe: the input file is only read; the result is saved as a NEW file. Needs Microsoft Excel.
# Run:  .\Add-SiteSummary.ps1
param(
    [string]$InputFile = 'C:\temp\Master_Last.xlsx',             # <-- master sheet
    [string]$Output    = 'C:\temp\Master_Last_counts.xlsx'       # <-- result (new file)
)

#region ======================= SETTINGS (edit here) =======================
# Site value (as typed in the Site column) = colour
$Sites = [ordered]@{
    'SixFlags'      = '1F4E79'
    'AquaArabia'    = '2E7D32'
    'SEVEN Tabuk'   = 'C55A11'
    'SEVEN ABHA'    = '7030A0'
    'SEVEN Alhamra' = 'A50021'
    'AMC'           = 'BF8F00'
}
# Banner starts with = box colour, row shade 1, row shade 2
$Sections = [ordered]@{
    'OT'       = @('00838F', 'A8E0E6', 'DDF3F5')
    'Test'     = @('C2185B', 'F8C8DA', 'FCE8F0')
    'Template' = @('595959', 'C9C9C9', 'EDEDED')
}
#endregion ==================================================================

$ErrorActionPreference = 'Stop'
function Ole([string]$hex) { [Convert]::ToInt32($hex.Substring(0, 2), 16) + 256 * [Convert]::ToInt32($hex.Substring(2, 2), 16) + 65536 * [Convert]::ToInt32($hex.Substring(4, 2), 16) }
function Norm($s) { (([string]$s) -replace '[\s\u00A0]+', ' ').Trim().ToLower() }
function ColLetter([int]$n) { $s = ''; while ($n -gt 0) { $m = ($n - 1) % 26; $s = [char](65 + $m) + $s; $n = [int](($n - $m - 1) / 26) }; $s }
function Q([string]$s) { '"' + $s.Replace('"', '""') + '"' }

if (-not (Test-Path $InputFile)) { Write-Host "Not found: $InputFile" -ForegroundColor Red; return }
$in = (Resolve-Path $InputFile).Path; $out = [IO.Path]::GetFullPath($Output)
if ($in -eq $out) { Write-Host 'Output must be a different file than the input.' -ForegroundColor Red; return }

$excel = $null; $wb = $null
try {
    $excel = New-Object -ComObject Excel.Application
    $excel.Visible = $false
    $excel.DisplayAlerts = $false
    $wb = $excel.Workbooks.Open($in, 0, $true)            # read-only; saved under the new name

    # ---------- Sheet, header row, columns ----------
    $ws = $null; $hdr = 0; $nameCol = 0
    foreach ($s in $wb.Worksheets) {
        $hit = $s.Range('A1:ZZ20').Find('VM NAME', [Type]::Missing, -4163, 1)
        if ($hit) { $ws = $s; $hdr = $hit.Row; $nameCol = $hit.Column; break }
    }
    if (-not $ws) { throw "No 'VM NAME' header found" }
    if ($hdr -lt 3) { throw "The header is in row $hdr - there is no row 2 above it for the counts" }
    $hv = $ws.Range($ws.Cells.Item($hdr, 1), $ws.Cells.Item($hdr, 60)).Value2
    $siteCol = 0; $lastCol = 0
    for ($c = 1; $c -le 60; $c++) {
        $v = Norm $hv[1, $c]
        if ($v -eq 'site' -and -not $siteCol) { $siteCol = $c }
        if ($v -and $v -notmatch '^(column\s*)?\d*$') { $lastCol = $c }
    }
    if (-not $siteCol) { throw "No 'Site' column found in the header row" }
    $lo = $null
    foreach ($t in $ws.ListObjects) { if ($t.Range.Row -eq $hdr) { $lo = $t } }
    $lastRow = $hdr + 1
    for ($c = 1; $c -le $lastCol; $c++) { $lastRow = [Math]::Max($lastRow, $ws.Cells.Item($ws.Rows.Count, $c).End(-4162).Row) }
    $VL = ColLetter $nameCol
    Write-Host "Sheet '$($ws.Name)': header row $hdr, rows to $lastRow, Site = column $(ColLetter $siteCol), VM NAME = column $VL"

    # ---------- Find the banners (first filled cell of the row starts with OT / Test / Template) ----------
    $grid = $ws.Range($ws.Cells.Item(1, 1), $ws.Cells.Item($lastRow, $lastCol)).Value2
    $banners = New-Object System.Collections.Generic.List[object]
    for ($r = $hdr + 1; $r -le $lastRow; $r++) {
        $first = $null; $fc = 0; $cnt = 0
        for ($c = 1; $c -le $lastCol; $c++) { $v = ([string]$grid[$r, $c]).Trim(); if ($v) { $cnt++; if (-not $first) { $first = $v; $fc = $c } } }
        if ($cnt -ne 1) { continue }                                  # a banner has one text only
        foreach ($k in $Sections.Keys) {
            if ($first -match "^$k(\b|[^A-Za-z])") { $banners.Add(@{ Key = $k; Row = $r; Text = $first; Col = ColLetter $fc }); break }
        }
    }
    foreach ($k in $Sections.Keys) { if (-not @($banners | Where-Object { $_.Key -eq $k }).Count) { Write-Host "No '$k ...' banner found - $k count will be 0" -ForegroundColor Yellow } }
    foreach ($b in $banners) { Write-Host ("Banner row {0}: '{1}'" -f $b.Row, $b.Text) }

    # ---------- One table over all sections (filter works on the whole sheet) ----------
    # repeated header rows under the banners: remember their look (fill / text colour of the VM NAME cell)
    $hdrLooks = New-Object System.Collections.Generic.List[object]
    for ($r = $hdr + 1; $r -le $lastRow; $r++) {
        if ((Norm $grid[$r, $nameCol]) -eq 'vm name') {
            $hc = $ws.Cells.Item($r, $nameCol)
            $hdrLooks.Add(@{ Row = $r; NoFill = ($hc.Interior.ColorIndex -eq -4142); Fill = $hc.Interior.Color; Font = $hc.Font.Color; Bold = $hc.Font.Bold })
        }
    }
    foreach ($t in @($ws.ListObjects)) { if (-not $lo -or $t.Name -ne $lo.Name) { Write-Host "Joining table '$($t.Name)' into the main table"; $t.Unlist() } }
    $looks = New-Object System.Collections.Generic.List[object]
    for ($r = $hdr + 1; $r -le $lastRow; $r++) {
        $rowRange = $ws.Range($ws.Cells.Item($r, 1), $ws.Cells.Item($r, $lastCol))
        if ($rowRange.MergeCells -eq $false) { continue }
        for ($c = 1; $c -le $lastCol; $c++) {
            $cell = $ws.Cells.Item($r, $c)
            if ($cell.MergeCells -eq $true) {
                $ma = $cell.MergeArea
                $looks.Add(@{ Addr = $ma.Address(); Centred = ($ma.HorizontalAlignment -eq -4108 -or $ma.HorizontalAlignment -eq 7)
                              NoFill = ($cell.Interior.ColorIndex -eq -4142); Fill = $cell.Interior.Color
                              Font = $cell.Font.Color; Bold = $cell.Font.Bold; Size = $cell.Font.Size })
                [void]$ma.UnMerge()
            }
        }
    }
    if ($lo) { $lo.Resize($ws.Range($ws.Cells.Item($hdr, $lo.Range.Column), $ws.Cells.Item($lastRow, [Math]::Max($lastCol, $lo.Range.Column + $lo.Range.Columns.Count - 1)))) }
    else { $lo = $ws.ListObjects.Add(1, $ws.Range($ws.Cells.Item($hdr, 1), $ws.Cells.Item($lastRow, $lastCol)), $null, 1); $lo.TableStyle = 'TableStyleMedium2' }
    $ok = { param($v) ($null -ne $v) -and ($v -isnot [DBNull]) }
    foreach ($lk in $looks) {                                         # banners keep their look
        $rg = $ws.Range($lk.Addr)
        try { if ($lk.NoFill) { $rg.Interior.Pattern = -4142 } elseif (& $ok $lk.Fill) { $rg.Interior.Color = [double]$lk.Fill } } catch { }
        try { if (& $ok $lk.Font) { $rg.Font.Color = [double]$lk.Font } } catch { }
        try { if (& $ok $lk.Bold) { $rg.Font.Bold = [bool]$lk.Bold } } catch { }
        try { if (& $ok $lk.Size) { $rg.Font.Size = [double]$lk.Size } } catch { }
        try { if ($lk.Centred -and $rg.Columns.Count -gt 1) { $rg.HorizontalAlignment = [int]7 } } catch { }
    }
    foreach ($hl in $hdrLooks) {                                      # repeated header rows keep their look
        $rg = $ws.Range($ws.Cells.Item($hl.Row, 1), $ws.Cells.Item($hl.Row, $lastCol))
        try { if ($hl.NoFill) { $rg.Interior.Pattern = -4142 } elseif (& $ok $hl.Fill) { $rg.Interior.Color = [double]$hl.Fill } } catch { }
        try { if (& $ok $hl.Font) { $rg.Font.Color = [double]$hl.Font } } catch { }
        try { if (& $ok $hl.Bold) { $rg.Font.Bold = [bool]$hl.Bold } } catch { }
    }
    Write-Host "One table now covers rows $hdr-$lastRow ($($looks.Count) merged banner cell(s) unmerged, $($hdrLooks.Count) header row(s) - look kept)"

    # ---------- Formulas ----------
    $S = '$' + (ColLetter $siteCol) + '$' + ($hdr + 1) + ':$' + (ColLetter $siteCol) + '$1048576'
    $V = '$' + $VL + '$' + ($hdr + 1) + ':$' + $VL + '$1048576'
    $sorted = @($banners | Sort-Object { $_.Row })
    $secStart = @{}; $secEnd = @{}          # formula pieces that locate each section by its banner text
    for ($i = 0; $i -lt $sorted.Count; $i++) {
        $b = $sorted[$i]
        # search only below the header (a search of the whole column would include row 2 itself = circular)
        $secStart[$b.Key] = "(MATCH($(Q $b.Text),`$$($b.Col)`$$($hdr + 1):`$$($b.Col)`$1048576,0)+$hdr)"
        $secEnd[$b.Key] = $(if ($i + 1 -lt $sorted.Count) { $n = $sorted[$i + 1]; "(MATCH($(Q $n.Text),`$$($n.Col)`$$($hdr + 1):`$$($n.Col)`$1048576,0)+$hdr)" } else { '1048577' })
    }

    # ---------- Row 2: same line as before (your formatting kept), but the numbers are live formulas ----------
    $parts = @("COUNTA($V)-COUNTIF($V,""VM NAME"")&"" VMs""")
    foreach ($k in $Sites.Keys) { $parts += """$k ""&COUNTIF($S,""$k"")" }
    foreach ($k in $Sections.Keys) {
        if ($secStart.ContainsKey($k)) {
            $rng = "INDEX(`$$($VL):`$$($VL),$($secStart[$k])+1):INDEX(`$$($VL):`$$($VL),$($secEnd[$k])-1)"
            $parts += """$k ""&(COUNTA($rng)-COUNTIF($rng,""VM NAME""))"
        }
    }
    $parts += """No site ""&SUMPRODUCT(($V<>"""")*($V<>""VM NAME"")*($S=""""))"
    $target = $ws.Cells.Item(2, 1).MergeArea.Cells.Item(1, 1)        # top-left cell of row 2's line
    $target.Formula = '=' + ($parts -join '&"  |  "&')
    $excel.Calculate()
    Write-Host "Row 2: $($target.Text)"
    # ---------- Colours ----------
    $body = $lo.DataBodyRange
    $siteCells = $ws.Range($ws.Cells.Item($body.Row, $siteCol), $ws.Cells.Item($body.Row + $body.Rows.Count - 1, $siteCol))
    foreach ($k in $Sites.Keys) {                                   # site name in its colour
        $fc = $siteCells.FormatConditions.Add(1, 3, "=""$k""")
        $fc.Font.Color = Ole $Sites[$k]; $fc.Font.Bold = $true
        $fc.StopIfTrue = $false; [void]$fc.SetLastPriority()
    }
    $vmFirst = '$' + $VL + $body.Row                                # e.g. $E4 (row-relative)
    foreach ($k in $Sections.Keys) {                                # section rows: own two-tone colours
        if (-not $secStart.ContainsKey($k)) { continue }
        foreach ($shade in 0, 1) {
            $cond = "=AND(ROW()>$($secStart[$k]),ROW()<$($secEnd[$k]),$vmFirst<>""VM NAME"",MOD(ROW(),2)=$shade)"
            $fr = $body.FormatConditions.Add(2, 0, $cond)
            $fr.Interior.Color = Ole $Sections[$k][1 + $shade]
            $fr.StopIfTrue = $false; [void]$fr.SetLastPriority()      # Power Status colours stay on top
        }
    }

    if (Test-Path $out) { Remove-Item $out -Force }
    $wb.SaveAs($out, 51)
    if (-not (Test-Path $out)) { throw "Excel did not write $out" }
    $excel.Calculate()
    Write-Host ''
    Write-Host "Row 2 now: $($ws.Cells.Item(2, 1).MergeArea.Cells.Item(1, 1).Text)"
    Write-Host "`nSaved: $out" -ForegroundColor Green
}
catch { Write-Host "FAILED at line $($_.InvocationInfo.ScriptLineNumber): $($_.Exception.Message)  (input file not changed)" -ForegroundColor Red }
finally {
    if ($wb) { try { $wb.Close($false) } catch { } }
    if ($excel) { $excel.Quit() }
    foreach ($o in @($wb, $excel)) { if ($o) { [void][Runtime.InteropServices.Marshal]::ReleaseComObject($o) } }
    [GC]::Collect(); [GC]::WaitForPendingFinalizers()
}
