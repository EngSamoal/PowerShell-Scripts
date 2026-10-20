<#
.SYNOPSIS
  Creates a copy of a workbook with FAKE data but the SAME design and formulas.
  Tab 1 recalculates from the fake inputs, so charts/tables keep their layout.

.HOW
  - Original is opened ReadOnly and is NEVER saved. Output is a new .xlsx.
  - Every typed-in (constant) number and text on ALL sheets is replaced.
  - Replacement is consistent: same original value -> same fake value everywhere,
    so lookups (VLOOKUP/XLOOKUP/SUMIF/COUNTIF...) keep working.
  - Text literals inside formulas (e.g. ="Riyadh") are replaced with the same mapping.
  - Formulas are kept. Dates, 0, 1, -1, TRUE/FALSE are kept.
  - Pivot tables: their cells are not overwritten and their source header row is kept,
    so the pivot keeps its fields; caches are refreshed and old items purged.
  - Protected sheets are unlocked with -SheetPassword; if one can't be unlocked the
    script stops (it never leaves a sheet with real data in the copy).
  - Removes: external links, data connections/queries, comments/notes,
    hyperlinks, document properties, VBA (saved as .xlsx).
  - NOT faked (check manually): sheet names, chart/shape/text-box text,
    headers/footers, data-validation lists, defined-name constants, Data Model.

.EXAMPLE
  .\New-FakeDataCopy.ps1 -Path "C:\path\file.xlsx"
  .\New-FakeDataCopy.ps1 -Path "C:\path\file.xlsx" -SheetPassword "pw"
#>
param(
    [Parameter(Mandatory = $true)][string]$Path,
    [string]$OutFile = "",
    [string]$SheetPassword = "",
    [int]$Seed = 12345
)

$ErrorActionPreference = "Stop"
$Path = (Resolve-Path -LiteralPath $Path).Path
if (-not $OutFile) {
    $OutFile = [IO.Path]::Combine((Split-Path $Path), [IO.Path]::GetFileNameWithoutExtension($Path) + "_FAKE.xlsx")
} else {
    $OutFile = $ExecutionContext.SessionState.Path.GetUnresolvedProviderPathFromPSPath($OutFile)   # Excel needs a full path
}
if ($OutFile -eq $Path) { throw "OutFile must be different from the original." }

$rnd      = New-Object System.Random($Seed)
$textMap  = New-Object 'System.Collections.Generic.Dictionary[string,string]' ([StringComparer]::OrdinalIgnoreCase)
$numMap   = New-Object 'System.Collections.Generic.Dictionary[double,double]'
$keepText = New-Object 'System.Collections.Generic.HashSet[string]' ([StringComparer]::OrdinalIgnoreCase)
$skip     = New-Object 'System.Collections.Generic.HashSet[string]'    # "Sheet|row|col" cells never written
$stats    = @{ Numbers = 0; Texts = 0; Formulas = 0; ArraySkipped = 0 }
$warn     = New-Object System.Collections.Generic.List[string]

function Test-DateFormat([string]$fmt) {
    if (-not $fmt) { return $false }
    $f = $fmt -replace '"[^"]*"', '' -replace '\[[^\]]*\]', '' -replace '\\.', ''
    return ($f -match '(?i)[dmyhs]')
}

function Get-FakeNumber([double]$v) {
    if ($v -eq 0 -or $v -eq 1 -or $v -eq -1) { return $v }
    if ($numMap.ContainsKey($v)) { return $numMap[$v] }
    $f = 0.6 + ($rnd.NextDouble() * 0.8)          # 0.6 .. 1.4
    $n = $v * $f
    if ([Math]::Floor($v) -eq $v) { $n = [Math]::Round($n); if ($n -eq 0) { $n = [Math]::Sign($v) } }
    else { $n = [Math]::Round($n, 4) }
    $numMap[$v] = $n
    return $n
}

function Get-FakeText([string]$s) {
    if ([string]::IsNullOrWhiteSpace($s)) { return $s }
    if ($keepText.Contains($s)) { return $s }     # pivot field names stay real everywhere
    if ($textMap.ContainsKey($s)) { return $textMap[$s] }
    $t = "Txt" + ($textMap.Count + 1).ToString("0000")
    $textMap[$s] = $t
    return $t
}

# PowerShell's COM binder can cache Value2's setter as [string] and then fail with
# "Unable to cast ... to String" on numbers/arrays; InvokeMember always works.
function Set-Value2($rng, $v) {
    [void][System.__ComObject].InvokeMember("Value2", [Reflection.BindingFlags]::SetProperty, $null, $rng, @(,$v))
}

function Add-Skip($rng, [switch]$Header) {
    foreach ($c in $rng.Cells) {
        [void]$skip.Add("$($c.Worksheet.Name)|$($c.Row)|$($c.Column)")
        if ($Header -and $c.Value2 -is [string]) { [void]$keepText.Add($c.Value2) }
    }
}

$xl = $null; $wb = $null
try {
    $xl = New-Object -ComObject Excel.Application
    $xl.Visible = $false
    $xl.DisplayAlerts = $false
    $xl.AskToUpdateLinks = $false
    $xl.EnableEvents = $false
    $xl.ScreenUpdating = $false

    Write-Host "Opening (read-only)..."
    $wb = $xl.Workbooks.Open($Path, 0, $true)     # UpdateLinks=0, ReadOnly=True
    $xl.Calculation = -4135                         # manual while editing

    try {
        $mt = $wb.Model.ModelTables.Count
        if ($mt -gt 0) { $warn.Add("Data Model (Power Pivot) has $mt table(s) - its data is NOT faked") }
    } catch {}

    # ---- Break external links (converts them to values, then they get faked) ----
    try {
        $links = $wb.LinkSources(1)
        if ($links) { foreach ($l in $links) { $wb.BreakLink($l, 1) } ; $warn.Add("External links broken: $($links.Count)") }
    } catch {}

    # ---- Remove data connections / queries ----
    try { for ($i = $wb.Queries.Count; $i -ge 1; $i--) { $wb.Queries.Item($i).Delete() } } catch {}
    try { for ($i = $wb.Connections.Count; $i -ge 1; $i--) { $wb.Connections.Item($i).Delete() } } catch {}

    # ---- Unprotect; stop rather than leave real data behind ----
    $sheets = @(); $locked = @()
    foreach ($ws in $wb.Worksheets) {
        if ($ws.ProtectContents) {
            try { $ws.Unprotect($SheetPassword) } catch { $locked += $ws.Name; continue }   # explicit password: no hidden prompt
        }
        $sheets += $ws
    }
    if ($locked) { throw "Protected sheet(s) could not be unlocked: $($locked -join ', '). Re-run with -SheetPassword." }

    # ---- Pivots: don't write into pivot cells, keep source headers (= field names) ----
    foreach ($ws in $sheets) {
        $pts = $ws.PivotTables()
        for ($i = 1; $i -le $pts.Count; $i++) {
            $pt = $pts.Item($i)
            Add-Skip $pt.TableRange2
            if ($pt.PivotCache().SourceType -ne 1) { continue }
            try {
                $src = $xl.Range($xl.ConvertFormula([string]$pt.SourceData, -4150, 1))   # R1C1 -> A1
                if ($src.ListObject) { Add-Skip $src.ListObject.HeaderRowRange -Header }
                else { Add-Skip $src.Rows.Item(1) -Header }
            } catch { $warn.Add("Pivot source not resolved, headers faked (pivot may lose fields): $($ws.Name)/$($pt.Name) - $($_.Exception.Message)") }
        }
    }

    # ---- Pass 1: fake constants ----
    foreach ($ws in $sheets) {
        Write-Host "Constants: $($ws.Name)"
        try { $ws.Cells.ClearComments() } catch {}
        try { $ws.Hyperlinks.Delete() } catch {}
        $const = $null
        try { $const = $ws.UsedRange.SpecialCells(2) } catch { continue }   # xlCellTypeConstants
        $sn = $ws.Name

        foreach ($area in $const.Areas) {
            $rows = $area.Rows.Count; $cols = $area.Columns.Count
            $r0 = $area.Row; $c0 = $area.Column

            # Date detection per column (per cell only if column is mixed)
            $isDate = New-Object 'bool[,]' ($rows + 1), ($cols + 1)
            for ($c = 1; $c -le $cols; $c++) {
                $nf = $area.Columns.Item($c).NumberFormat
                if ($nf -is [string]) {
                    $d = Test-DateFormat $nf
                    for ($r = 1; $r -le $rows; $r++) { $isDate[$r, $c] = $d }
                } else {
                    for ($r = 1; $r -le $rows; $r++) { $isDate[$r, $c] = Test-DateFormat ([string]$area.Cells.Item($r, $c).NumberFormat) }
                }
            }

            $val = $area.Value2
            if ($rows -eq 1 -and $cols -eq 1) {
                if ($skip.Contains("$sn|$r0|$c0")) { continue }
                if ($val -is [double] -and -not $isDate[1, 1]) { Set-Value2 $area (Get-FakeNumber $val); $stats.Numbers++ }
                elseif ($val -is [string]) { Set-Value2 $area (Get-FakeText $val); $stats.Texts++ }
                continue
            }
            $touched = $false; $hit = $false
            for ($r = 1; $r -le $rows; $r++) {
                for ($c = 1; $c -le $cols; $c++) {
                    if ($skip.Count -and $skip.Contains("$sn|$($r0 + $r - 1)|$($c0 + $c - 1)")) { $hit = $true; continue }
                    $v = $val[$r, $c]
                    if ($v -is [double] -and -not $isDate[$r, $c]) { $val[$r, $c] = Get-FakeNumber $v; $stats.Numbers++; $touched = $true }
                    elseif ($v -is [string]) { $val[$r, $c] = Get-FakeText $v; $stats.Texts++; $touched = $true }
                }
            }
            if (-not $touched) { continue }
            # An area overlapping a pivot can't be written as a block -> write cell by cell
            if ($hit) {
                for ($r = 1; $r -le $rows; $r++) {
                    for ($c = 1; $c -le $cols; $c++) {
                        if ($skip.Contains("$sn|$($r0 + $r - 1)|$($c0 + $c - 1)")) { continue }
                        $cell = $area.Cells.Item($r, $c)
                        if ($cell.Value2 -ne $val[$r, $c]) { Set-Value2 $cell $val[$r, $c] }
                    }
                }
            } else {
                Set-Value2 $area $val
            }
        }
    }

    # ---- Pass 2: replace text literals inside formulas ----
    $useF2 = $true
    try { $null = $sheets[0].Range("A1").Formula2 } catch { $useF2 = $false }
    $rx = [regex]'"((?:[^"]|"")*)"'

    foreach ($ws in $sheets) {
        Write-Host "Formulas:  $($ws.Name)"
        $fr = $null
        try { $fr = $ws.UsedRange.SpecialCells(-4123) } catch { continue }   # xlCellTypeFormulas
        foreach ($area in $fr.Areas) {
            $rows = $area.Rows.Count; $cols = $area.Columns.Count
            if ($useF2) { $fa = $area.Formula2 } else { $fa = $area.Formula }
            for ($r = 1; $r -le $rows; $r++) {
                for ($c = 1; $c -le $cols; $c++) {
                    if ($rows -eq 1 -and $cols -eq 1) { $f = [string]$fa } else { $f = [string]$fa[$r, $c] }
                    if ($f.IndexOf('"') -lt 0) { continue }
                    $new = $rx.Replace($f, {
                        param($m)
                        $inner = $m.Groups[1].Value.Replace('""', '"')
                        if ($inner -and $textMap.ContainsKey($inner)) { return '"' + $textMap[$inner].Replace('"', '""') + '"' }
                        return $m.Value
                    })
                    if ($new -ne $f) {
                        $cell = $area.Cells.Item($r, $c)
                        if ($cell.HasArray) { $stats.ArraySkipped++; continue }
                        if ($useF2) { $cell.Formula2 = $new } else { $cell.Formula = $new }
                        $stats.Formulas++
                    }
                }
            }
        }
    }

    # ---- Recalculate, refresh internal pivots (purge old real items from caches) ----
    Write-Host "Recalculating..."
    $xl.Calculation = -4105                          # automatic
    $xl.CalculateFull()
    foreach ($ws in $sheets) {
        $pts = $ws.PivotTables()
        for ($i = 1; $i -le $pts.Count; $i++) {
            $pt = $pts.Item($i); $pc = $pt.PivotCache()
            if ($pc.SourceType -ne 1) { $warn.Add("Pivot not refreshed (external source, cache keeps REAL data): $($ws.Name)/$($pt.Name)"); continue }
            try { $pc.MissingItemsLimit = 0; $pc.Refresh() }   # xlMissingItemsNone
            catch { $warn.Add("Pivot refresh failed (cache keeps REAL data): $($ws.Name)/$($pt.Name) - $($_.Exception.Message)") }
        }
    }

    try { $wb.RemoveDocumentInformation(99) } catch {}

    # ---- Save as NEW .xlsx (drops VBA). Original untouched. ----
    $wb.SaveAs($OutFile, 51)
    Write-Host ""
    Write-Host "Done: $OutFile"
    Write-Host ("Numbers faked: {0} | Texts faked: {1} | Formulas updated: {2} | Array formulas skipped: {3}" -f $stats.Numbers, $stats.Texts, $stats.Formulas, $stats.ArraySkipped)
    foreach ($w in $warn) { Write-Host "WARNING: $w" -ForegroundColor Yellow }
}
finally {
    if ($wb) { try { $wb.Close($false) } catch {} }
    if ($xl) { try { $xl.Quit() } catch {}; [void][Runtime.InteropServices.Marshal]::ReleaseComObject($xl) }
    # drop every COM reference so EXCEL.EXE exits even in a long-lived shell (ISE/VS Code)
    $wb = $xl = $ws = $sheets = $area = $const = $fr = $cell = $pts = $pt = $pc = $src = $null
    [GC]::Collect(); [GC]::WaitForPendingFinalizers(); [GC]::Collect()
}
