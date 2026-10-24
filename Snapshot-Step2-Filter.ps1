# Snapshot-Step2-Filter.ps1
# STEP 2 of 3 - Picks the snapshots to delete from the Step 1 list.
# Read-only: does not talk to vCenter at all, it only reads the Step 1 CSV.
#
# Run    : .\Snapshot-Step2-Filter.ps1 -NameKeyword 'patch' -DescriptionKeyword 'CHG12345'
#          (a snapshot is picked only if its NAME contains the first word AND its
#           DESCRIPTION contains the second word; upper/lower case does not matter)
# Input  : C:\temp\SnapshotCleanup\1-AllSnapshots.csv
# Output : C:\temp\SnapshotCleanup\2-SnapshotsToDelete.csv
#          Step 3 deletes every snapshot in this file - nothing else. No editing needed.
# Next   : .\Snapshot-Step3-Delete.ps1
param(
    [string]$NameKeyword,                                   # word that must be in the snapshot name
    [string]$DescriptionKeyword,                            # word that must be in the snapshot description
    [ValidateSet('And', 'Or')][string]$MatchMode = 'And',   # Or = either word is enough
    [switch]$WholeWord,                                     # 'patch' will not match 'patching'
    [string]$OutputFolder = 'C:\temp\SnapshotCleanup'
)

$ScriptBuild = 'Snapshot-Step2-Filter 2026.10.24-r5'
Write-Host $ScriptBuild -ForegroundColor Cyan
$ErrorActionPreference = 'Stop'

function Write-SnapshotTable($Rows) {
    # One line per snapshot, full text, columns aligned. Line breaks inside a
    # description are shown as ' / ' so a row never splits over several lines.
    $cols = 'VCenter', 'VM', 'Created', 'SizeGB', 'SnapshotName', 'Description'
    $lines = @(foreach ($r in @($Rows)) {
        , @(foreach ($c in $cols) { (([string]$r.$c) -replace '\s*[\r\n]+\s*', ' / ' -replace '\t', ' ').Trim() })
    })
    $widths = @(for ($i = 0; $i -lt $cols.Count; $i++) {
        $w = $cols[$i].Length
        foreach ($l in $lines) { if ($l[$i].Length -gt $w) { $w = $l[$i].Length } }
        $w
    })
    $fmt = { param($vals) (@(for ($i = 0; $i -lt $vals.Count; $i++) { if ($i -eq $vals.Count - 1) { $vals[$i] } else { $vals[$i].PadRight($widths[$i]) } }) -join '  ') }
    Write-Host ''
    Write-Host (& $fmt $cols)
    Write-Host (& $fmt @($widths | ForEach-Object { '-' * $_ }))
    foreach ($l in $lines) { Write-Host (& $fmt $l) }
    Write-Host ''
}

function Test-KeywordMatch([string]$Text, [string]$Keyword, [bool]$Whole) {
    # Case-insensitive, keyword is plain text (no wildcards / regex)
    if ([string]::IsNullOrEmpty($Text)) { return $false }
    if ($Whole) { return [regex]::IsMatch($Text, '(?<![\w])' + [regex]::Escape($Keyword) + '(?![\w])', 'IgnoreCase') }
    return ($Text.IndexOf($Keyword, [StringComparison]::OrdinalIgnoreCase) -ge 0)
}

if (-not $NameKeyword -and -not $DescriptionKeyword) {
    Write-Host 'Give -NameKeyword and/or -DescriptionKeyword. Example:' -ForegroundColor Red
    Write-Host "  .\Snapshot-Step2-Filter.ps1 -NameKeyword 'patch' -DescriptionKeyword 'CHG12345'"; return
}
if ($NameKeyword) { $NameKeyword = $NameKeyword.Trim() }
if ($DescriptionKeyword) { $DescriptionKeyword = $DescriptionKeyword.Trim() }
foreach ($kw in @($NameKeyword, $DescriptionKeyword)) {
    if ($null -ne $kw -and $kw -ne '' -and $kw.Length -lt 3) { Write-Host "Keyword '$kw' is too short (minimum 3 characters)." -ForegroundColor Red; return }
}

$inCsv  = Join-Path $OutputFolder '1-AllSnapshots.csv'
$outCsv = Join-Path $OutputFolder '2-SnapshotsToDelete.csv'
if (-not (Test-Path -LiteralPath $inCsv)) { Write-Host "Not found: $inCsv  - run .\Snapshot-Step1-ListAll.ps1 first." -ForegroundColor Red; return }
$age = (Get-Date) - (Get-Item -LiteralPath $inCsv).LastWriteTime
if ($age.TotalHours -gt 24) { Write-Host ("WARNING: the Step 1 list is {0:N0} hours old - consider running Step 1 again." -f $age.TotalHours) -ForegroundColor Yellow }

$all = @(Import-Csv -LiteralPath $inCsv)
$selected = @($all | Where-Object {
    $checks = @()
    if ($NameKeyword) { $checks += (Test-KeywordMatch $_.SnapshotName $NameKeyword $WholeWord.IsPresent) }
    if ($DescriptionKeyword) { $checks += (Test-KeywordMatch $_.Description $DescriptionKeyword $WholeWord.IsPresent) }
    if ($MatchMode -eq 'Or') { $checks -contains $true } else { -not ($checks -contains $false) }
})

$desc = @()
if ($NameKeyword) { $desc += "name contains '$NameKeyword'" }
if ($DescriptionKeyword) { $desc += "description contains '$DescriptionKeyword'" }
$joiner = ' AND '; if ($MatchMode -eq 'Or') { $joiner = ' OR ' }
$wholeTxt = ''; if ($WholeWord) { $wholeTxt = ' (whole word)' }
Write-Host ("Filter: {0}{1}" -f ($desc -join $joiner), $wholeTxt)

$selected | Export-Csv -Path $outCsv -NoTypeInformation -Encoding UTF8
if ($selected.Count -eq 0) {
    Write-Host "No snapshot matched (out of $($all.Count)). Nothing to delete." -ForegroundColor Yellow; return
}
Write-SnapshotTable $selected
$totalGB = [math]::Round(($selected | Measure-Object SizeGB -Sum).Sum + 0, 2)
$vmCount = @($selected | ForEach-Object { "$($_.VCenter)|$($_.VMId)" } | Select-Object -Unique).Count
Write-Host ("Matched: {0} of {1} snapshots, on {2} VM(s), {3} GB" -f $selected.Count, $all.Count, $vmCount, $totalGB) -ForegroundColor Green
Write-Host "Saved to: $outCsv" -ForegroundColor Green
Write-Host ''
Write-Host 'NEXT: run  .\Snapshot-Step3-Delete.ps1   to delete the snapshots listed above' -ForegroundColor Cyan
