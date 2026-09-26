# Set-UniformColors.ps1
# Makes the data rows look the same in every tab: only the table's blue/white banding,
# black normal text - except the "Power Status" and "Implementor" columns, which are kept as they are.
# Safe: the input file is only read; the result is saved as a NEW file. Needs Microsoft Excel.
# Run:  .\Set-UniformColors.ps1
param(
    [string]$InputFile = 'C:\temp\Master1_withOwners.xlsx',          # <-- input
    [string]$Output    = 'C:\temp\Master1_withOwners_uniform.xlsx'   # <-- result (new file)
)

$KeepColumns = @('power status', 'implementor', 'implementer')       # these columns are not changed
$ErrorActionPreference = 'Stop'

if (-not (Test-Path $InputFile)) { Write-Host "Not found: $InputFile" -ForegroundColor Red; return }
$in = (Resolve-Path $InputFile).Path
$out = [IO.Path]::GetFullPath($Output)
if ($in -eq $out) { Write-Host 'Output must be a different file than the input.' -ForegroundColor Red; return }

$excel = $null; $wb = $null
try {
    $excel = New-Object -ComObject Excel.Application
    $excel.Visible = $false
    $excel.DisplayAlerts = $false
    $wb = $excel.Workbooks.Open($in)

    foreach ($ws in $wb.Worksheets) {
        if ($ws.ListObjects.Count -eq 0) { Write-Host ("{0,-16} no table - skipped" -f $ws.Name) -ForegroundColor Yellow; continue }
        $lo = $ws.ListObjects.Item(1)
        if (-not $lo.DataBodyRange) { continue }
        $done = @()
        foreach ($col in $lo.ListColumns) {
            if ($KeepColumns -contains ([string]$col.Name).Trim().ToLower()) { continue }
            $r = $col.DataBodyRange
            $r.FormatConditions.Delete()          # remove highlight rules (orange duplicates, grey N/A ...)
            $r.Interior.Pattern = -4142           # no fill -> table banding shows
            $r.Font.ColorIndex = -4105            # automatic (black)
            $r.Font.Italic = $false
            $done += $col.Name
        }
        Write-Host ("{0,-16} uniform: {1}" -f $ws.Name, ($done -join ', ')) -ForegroundColor Green
    }

    if (Test-Path $out) { Remove-Item $out -Force }
    $wb.SaveAs($out, 51)
    Write-Host "`nSaved: $out" -ForegroundColor Green
}
catch { Write-Host "FAILED: $($_.Exception.Message)  (input file not changed)" -ForegroundColor Red }
finally {
    if ($wb) { try { $wb.Close($false) } catch { } }
    if ($excel) { $excel.Quit() }
    foreach ($o in @($wb, $excel)) { if ($o) { [void][Runtime.InteropServices.Marshal]::ReleaseComObject($o) } }
    [GC]::Collect(); [GC]::WaitForPendingFinalizers()
}
