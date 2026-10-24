# Snapshot-Step2-Filter.ps1
# STEP 2 of 3 - Lists the snapshots that match your words (these are the ones Step 3 deletes).
# Read-only: nothing is changed in vCenter.
#
# Run    : .\Snapshot-Step2-Filter.ps1 -NameKeyword 'patch' -DescriptionKeyword 'CHG12345'
#          (a snapshot is picked only if its NAME contains the first word AND its
#           DESCRIPTION contains the second word; upper/lower case does not matter)
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

$ScriptBuild = 'Snapshot-Step2-Filter 2026.10.24-r6'
Write-Host $ScriptBuild -ForegroundColor Cyan
$ErrorActionPreference = 'Stop'

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

$outCsv = Join-Path $OutputFolder '2-SnapshotsToDelete.csv'
if (-not (Test-Path -LiteralPath $OutputFolder)) { New-Item -ItemType Directory -Path $OutputFolder -Force | Out-Null }

$servers = @($global:DefaultVIServers | Where-Object { $_.IsConnected })
if ($servers.Count -eq 0) { Write-Host 'Not connected to any vCenter. Run Connect-VIServer first.' -ForegroundColor Red; return }

# Read all snapshots per vCenter (vCenter is remembered so Step 3 deletes on the right one)
# A VM that cannot be read is reported and skipped; the rest is still read.
$all = @()
$readErrors = @()
foreach ($srv in $servers) {
    Write-Host "Reading snapshots from $($srv.Name) ..." -ForegroundColor Yellow
    $all += @(Get-VM -Server $srv | Get-Snapshot -ErrorAction Continue -ErrorVariable +readErrors |
        ForEach-Object { Add-Member -InputObject $_ -NotePropertyName VCenter -NotePropertyValue $srv.Name -Force -PassThru })
}
if ($readErrors.Count -gt 0) { Write-Host "WARNING: $($readErrors.Count) VM(s) could not be read - see red messages above. Their snapshots are not in this list." -ForegroundColor Red }
$selected = @($all | Where-Object {
    $checks = @()
    if ($NameKeyword) { $checks += (Test-KeywordMatch $_.Name $NameKeyword $WholeWord.IsPresent) }
    if ($DescriptionKeyword) { $checks += (Test-KeywordMatch $_.Description $DescriptionKeyword $WholeWord.IsPresent) }
    if ($MatchMode -eq 'Or') { $checks -contains $true } else { -not ($checks -contains $false) }
} | Sort-Object Created)
$desc = @()
if ($NameKeyword) { $desc += "name contains '$NameKeyword'" }
if ($DescriptionKeyword) { $desc += "description contains '$DescriptionKeyword'" }
$joiner = ' AND '; if ($MatchMode -eq 'Or') { $joiner = ' OR ' }
$wholeTxt = ''; if ($WholeWord) { $wholeTxt = ' (whole word)' }
Write-Host ("Filter: {0}{1}" -f ($desc -join $joiner), $wholeTxt)

$selected | Select-Object VCenter, VM, @{N='SnapshotName';E={$_.Name}}, Description,
    @{N='Created';E={$_.Created.ToString('yyyy-MM-dd HH:mm')}}, @{N='SizeGB';E={[math]::Round($_.SizeGB,2)}},
    @{N='Parent';E={$_.ParentSnapshot.Name}}, @{N='VMId';E={$_.VM.Id}}, @{N='SnapshotId';E={$_.Id}} |
    Export-Csv -Path $outCsv -NoTypeInformation -Encoding UTF8
if ($selected.Count -eq 0) {
    Write-Host "No snapshot matched (out of $($all.Count)). Nothing to delete." -ForegroundColor Yellow; return
}
$selected | Select-Object VM,
    @{N='SnapshotName';E={$_.Name}},
    Description,
    Created,
    @{N='SizeGB';E={[math]::Round($_.SizeGB,2)}},
    @{N='Parent';E={$_.ParentSnapshot.Name}} | Format-Table -AutoSize | Out-Host
$totalGB = [math]::Round(($selected | Measure-Object SizeGB -Sum).Sum + 0, 2)
$vmCount = @($selected | ForEach-Object { "$($_.VCenter)|$($_.VM.Id)" } | Select-Object -Unique).Count
Write-Host ("Matched: {0} of {1} snapshots, on {2} VM(s), {3} GB" -f $selected.Count, $all.Count, $vmCount, $totalGB) -ForegroundColor Green
Write-Host "Saved to: $outCsv" -ForegroundColor Green
Write-Host ''
Write-Host 'NEXT: run  .\Snapshot-Step3-Delete.ps1   to delete the snapshots listed above' -ForegroundColor Cyan
