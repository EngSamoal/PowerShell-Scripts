#Requires -Version 5.1
<#
.SYNOPSIS
    Injects evidence photos into an already-generated Network Weekly Health Check dashboard.
    Optional and separate on purpose: a month with no photos supplied needs no change to
    Network_Weekly_HealthCheck.ps1 at all - just don't run this script.

.DESCRIPTION
    Network_Weekly_HealthCheck.ps1 tags every IDF Room Issues row in the dashboard HTML
    with a data-idf-row="<No.>" attribute and an empty evidence cell. This script:
      1. Reads that dashboard HTML.
      2. Walks -PicturesPath, one subfolder per site (subfolder name must match the site name
         exactly as it appears in the dashboard, e.g. "SixFlags", "AquaArabia").
      3. In each site subfolder, picks up every image file named after the IDF Summary row
         number it is evidence for - "10.jpg" for a row with one photo, or "10-1.jpg",
         "10-2.jpg", "10-3.jpg", ... for a row with several (any image extension: jpg/jpeg/
         png/gif/bmp).
      4. Embeds each matching photo (as a base64 data URI, so the dashboard stays a single
         portable .html file with no separate image files to keep alongside it) into that row's
         evidence cell, scoped to the correct site so a "row 10" picture for SixFlags can never
         land on AquaArabia's row 10 (or vice versa).
      5. Writes the result to a new file (never overwrites the original dashboard) and reports
         which rows got photos, which picture files didn't match any row, and which rows with
         issues have no photos supplied.

    NEVER modifies the source dashboard HTML file or the picture files themselves.

.PARAMETER DashboardPath
    Path to the dashboard .html file produced by Network_Weekly_HealthCheck.ps1.
.PARAMETER PicturesPath
    Folder containing one subfolder per site (e.g. Pictures\SixFlags\10-1.jpg). Defaults to a
    "Pictures" folder next to this script.
.PARAMETER OutputPath
    Where to write the updated dashboard. Defaults to the dashboard's own name with
    "-WithPhotos" inserted before the .html extension, in the same folder.

.EXAMPLE
    .\Network_Evidence_Photos.ps1 -DashboardPath "C:\...\Network_HealthCheck_Dashboard_2026-10-01.html" -PicturesPath "C:\...\Pictures"
#>

[CmdletBinding()]
param(
    [Parameter(Mandatory)][string]$DashboardPath,
    [string]$PicturesPath = (Join-Path $PSScriptRoot "Pictures"),
    [string]$OutputPath = ''
)

$ErrorActionPreference = 'Stop'

if (-not (Test-Path $DashboardPath)) { throw "Dashboard file not found: $DashboardPath" }
if (-not (Test-Path $PicturesPath)) { throw "Pictures folder not found: $PicturesPath" }

if (-not $OutputPath) {
    $dir = Split-Path -Parent $DashboardPath
    $baseName = [System.IO.Path]::GetFileNameWithoutExtension($DashboardPath)
    $OutputPath = Join-Path $dir "$baseName-WithPhotos.html"
}

$MimeByExt = @{ '.jpg' = 'image/jpeg'; '.jpeg' = 'image/jpeg'; '.png' = 'image/png'; '.gif' = 'image/gif'; '.bmp' = 'image/bmp' }

function Get-RowPhotoMap {
    param([string]$SiteFolder)
    # Groups every image file in a site's picture folder by the row number in its filename:
    # "10.jpg" -> row 10, one photo. "10-1.jpg"/"10-2.jpg" -> row 10, two photos (sorted by the
    # trailing -N). Any file that doesn't match this pattern is reported back as unmatched.
    $map = @{}
    $unmatched = @()
    $files = Get-ChildItem -Path $SiteFolder -File -ErrorAction SilentlyContinue
    foreach ($f in $files) {
        if ($f.Name -match '^(?<row>\d+)(-(?<seq>\d+))?(?<ext>\.(jpe?g|png|gif|bmp))$') {
            $row = $Matches.row
            $seq = if ($Matches.seq) { [int]$Matches.seq } else { 0 }
            if (-not $map.ContainsKey($row)) { $map[$row] = [System.Collections.Generic.List[object]]::new() }
            $map[$row].Add([pscustomobject]@{ Seq = $seq; Path = $f.FullName; Ext = $Matches.ext.ToLower() })
        } else {
            $unmatched += $f.FullName
        }
    }
    foreach ($k in @($map.Keys)) { $map[$k] = $map[$k] | Sort-Object Seq }
    return [pscustomobject]@{ Map = $map; Unmatched = $unmatched }
}

function ConvertTo-ImgTagsHtml {
    param([object[]]$Photos)
    $tags = foreach ($p in $Photos) {
        $mime = $MimeByExt[$p.Ext]
        if (-not $mime) { continue }
        $bytes = [System.IO.File]::ReadAllBytes($p.Path)
        $b64 = [System.Convert]::ToBase64String($bytes)
        $dataUri = "data:$mime;base64,$b64"
        "<a href=`"$dataUri`" target=`"_blank`"><img src=`"$dataUri`" alt=`"Evidence photo`"></a>"
    }
    return ($tags -join ' ')
}

Write-Host "Network_Evidence_Photos.ps1" -ForegroundColor Magenta
$html = Get-Content -Path $DashboardPath -Raw

$siteFolders = Get-ChildItem -Path $PicturesPath -Directory -ErrorAction SilentlyContinue
if (-not $siteFolders) {
    Write-Warning "No site subfolders found under '$PicturesPath' - nothing to inject. Expected e.g. Pictures\SixFlags\10-1.jpg"
}

$totalInjected = 0
$totalUnmatchedFiles = 0
$totalRowsNoPhoto = 0

foreach ($siteFolder in $siteFolders) {
    $siteName = $siteFolder.Name
    Write-Host "`n=== $siteName ===" -ForegroundColor Green

    # Scope to this site's own <section ... data-site="SiteName"> ... </section> block, so a
    # row number can never match the wrong site's dashboard page.
    $sectionPattern = "(?s)(<section class=`"page`"[^>]*data-site=`"" + [regex]::Escape($siteName) + "`"[^>]*>)(.*?)(</section>)"
    $sectionMatch = [regex]::Match($html, $sectionPattern)
    if (-not $sectionMatch.Success) {
        Write-Warning "No dashboard page found for site '$siteName' (folder name must exactly match the site name shown in the dashboard) - skipping."
        continue
    }
    $sectionBody = $sectionMatch.Groups[2].Value

    $result = Get-RowPhotoMap -SiteFolder $siteFolder.FullName
    $rowPhotoMap = $result.Map
    if ($result.Unmatched.Count -gt 0) {
        $totalUnmatchedFiles += $result.Unmatched.Count
        foreach ($u in $result.Unmatched) { Write-Warning "File name doesn't match the <row>.ext or <row>-<n>.ext pattern, skipped: $u" }
    }

    # Every data-idf-row="N" present in this site's dashboard page tells us which rows actually
    # have an issue reported (and so a real evidence cell to fill); any picture folder row
    # number that doesn't match one of these is reported as unmatched too.
    $dashboardRows = [regex]::Matches($sectionBody, 'data-idf-row="([^"]*)"') | ForEach-Object { $_.Groups[1].Value } | Select-Object -Unique

    foreach ($rowNum in $dashboardRows) {
        if ($rowPhotoMap.ContainsKey($rowNum)) {
            $imgHtml = ConvertTo-ImgTagsHtml -Photos $rowPhotoMap[$rowNum]
            $rowPattern = '(?s)(<div class="idf-issue-card" data-idf-row="' + [regex]::Escape($rowNum) + '"[^>]*>.*?<div class="idf-evidence-cell">)(</div>)'
            $newSectionBody = [regex]::Replace($sectionBody, $rowPattern, { param($m) $m.Groups[1].Value + $imgHtml + $m.Groups[2].Value }, 1)
            if ($newSectionBody -eq $sectionBody) {
                Write-Warning "Row $rowNum - found photo(s) but could not locate its table row in the HTML (unexpected)."
            } else {
                $sectionBody = $newSectionBody
                $totalInjected += $rowPhotoMap[$rowNum].Count
                Write-Host "  Row $rowNum - $($rowPhotoMap[$rowNum].Count) photo(s) added"
            }
        } else {
            $totalRowsNoPhoto++
            Write-Host "  Row $rowNum - no photo supplied" -ForegroundColor DarkGray
        }
    }

    $matchedRowNums = $dashboardRows | Where-Object { $rowPhotoMap.ContainsKey($_) }
    $orphanRowNums = $rowPhotoMap.Keys | Where-Object { $_ -notin $dashboardRows }
    foreach ($o in $orphanRowNums) {
        $totalUnmatchedFiles += $rowPhotoMap[$o].Count
        Write-Warning "Pictures found for row $o but the dashboard has no IDF issue for that row at $siteName - skipped."
    }

    $html = $html.Substring(0, $sectionMatch.Groups[2].Index) + $sectionBody + $html.Substring($sectionMatch.Groups[2].Index + $sectionMatch.Groups[2].Length)
}

$html | Out-File -FilePath $OutputPath -Encoding UTF8

Write-Host "`nDone. Photos injected: $totalInjected | Rows with no photo: $totalRowsNoPhoto | Unmatched/orphan files: $totalUnmatchedFiles" -ForegroundColor Green
Write-Host "Dashboard with photos written: $OutputPath" -ForegroundColor Cyan
