<#
.SYNOPSIS
    Creates one folder per server name listed in a text file.

.DESCRIPTION
    Reads server names (one per line) from the list file and creates a folder for each
    under the target path. Blank lines and lines starting with '#' are ignored, names are
    trimmed, duplicates are skipped, and folders that already exist are left untouched.

.EXAMPLE
    .\New-ServerFolders.ps1
    .\New-ServerFolders.ps1 -ListFile 'D:\lists\servers.txt' -TargetPath 'D:\upgrade'
#>
param(
    [string]$ListFile   = 'C:\temp\vmlist.txt',
    [string]$TargetPath = 'C:\temp\upgrade'
)

if (-not (Test-Path -LiteralPath $ListFile -PathType Leaf)) {
    Write-Error "Server list not found: $ListFile"
    exit 1
}

if (-not (Test-Path -LiteralPath $TargetPath)) {
    New-Item -ItemType Directory -Path $TargetPath -Force | Out-Null
    Write-Host "Created base folder: $TargetPath"
}

$invalidChars = [System.IO.Path]::GetInvalidFileNameChars()
$created = 0; $existing = 0; $skipped = 0

$servers = Get-Content -LiteralPath $ListFile |
    ForEach-Object { $_.Trim() } |
    Where-Object { $_ -and -not $_.StartsWith('#') } |
    Select-Object -Unique

foreach ($server in $servers) {
    if ($server.IndexOfAny($invalidChars) -ge 0) {
        Write-Warning "Skipped (invalid folder name): $server"
        $skipped++
        continue
    }

    $folder = Join-Path $TargetPath $server
    if (Test-Path -LiteralPath $folder) {
        Write-Host "Exists : $folder" -ForegroundColor Yellow
        $existing++
    }
    else {
        try {
            New-Item -ItemType Directory -Path $folder -ErrorAction Stop | Out-Null
            Write-Host "Created: $folder" -ForegroundColor Green
            $created++
        }
        catch {
            Write-Warning "Failed to create ${folder}: $($_.Exception.Message)"
            $skipped++
        }
    }
}

Write-Host ""
Write-Host "Done. Created: $created | Already existed: $existing | Skipped/failed: $skipped"
