<#
.SYNOPSIS
    Compares a vCenter VM inventory (List A) against an Azure Arc inventory (List B)
    and produces an enriched copy of List A with Arc onboarding status.

.DESCRIPTION
    READ-ONLY / REPORTING SCRIPT.
    - Does not connect to vCenter, Azure, or Azure Arc.
    - Does not modify, tag, move, or delete any VM or any source file.
    - Only reads two CSV files and writes a new CSV file.

    Matching logic:
      1. Exact match  : VM Name trimmed of leading/trailing whitespace, compared
                         CASE-INSENSITIVELY (e.g. 'tb-crbl-prx01' = 'TB-CRBL-PRX01' -
                         this is a display-convention difference, not a naming
                         ambiguity: Linux hosts commonly report a lowercase
                         hostname to the Arc agent while the vCenter display name
                         was set in uppercase). Only an exact match sets
                         ArcOnboardingStatus = 'Onboarded'.
      2. Possible match: Only evaluated for names that did NOT exact-match.
                         Names are normalized (uppercased, hyphens/underscores/
                         spaces stripped) into a "loose key". If the loose keys
                         match, the VM is flagged as a high-confidence candidate
                         for manual review only - it never sets Onboarded. This
                         bucket now only catches genuine structural differences:
                         hyphen vs underscore, missing/extra hyphen, accidental
                         spaces - since pure case differences are handled by the
                         exact match above.

.PARAMETER VCenterCsvPath
    Path to List A - the vCenter inventory CSV (master list).

.PARAMETER ArcCsvPath
    Path to List B - the Azure Arc inventory CSV.

.PARAMETER OutputPath
    Path to write the enriched CSV. Defaults to
    ".\vCenter-Inventory-With-Arc-Status.csv" next to the vCenter file.

.PARAMETER VCenterNameColumn
    Name of the VM Name column in List A. Default: 'VM Name'.

.PARAMETER ArcNameColumn
    Name of the VM Name column in List B. Default: 'VM Name'.

.PARAMETER ArcStatusColumn
    Name of the Arc agent/machine status column in List B.
    If not supplied, the script auto-detects common column names
    (Status, AgentStatus, Machine Status, Connectivity Status, etc.).
    If none is found, ArcAgentStatus will be blank for onboarded VMs
    rather than inventing a value.

.EXAMPLE
    .\Compare-VCenterArcInventory.ps1 -VCenterCsvPath .\vcenter.csv -ArcCsvPath .\arc.csv

.EXAMPLE
    .\Compare-VCenterArcInventory.ps1 -VCenterCsvPath .\vcenter.csv -ArcCsvPath .\arc.csv `
        -ArcStatusColumn 'Azure Arc Machine - Status' -OutputPath .\Reports\Result.csv
#>

[CmdletBinding()]
param(
    [Parameter(Mandatory = $true)]
    [ValidateScript({ Test-Path $_ })]
    [string]$VCenterCsvPath,

    [Parameter(Mandatory = $true)]
    [ValidateScript({ Test-Path $_ })]
    [string]$ArcCsvPath,

    [string]$OutputPath,

    [string]$VCenterNameColumn = 'VM Name',

    [string]$ArcNameColumn = 'VM Name',

    [string]$ArcStatusColumn
)

$ErrorActionPreference = 'Stop'

# ---------------------------------------------------------------------------
# Helpers
# ---------------------------------------------------------------------------

function Get-ExactKey {
    # Trim only. Case-insensitivity is applied via the dictionary comparer
    # (OrdinalIgnoreCase) at lookup time, so the original casing is preserved
    # here for display purposes.
    param([string]$Name)
    if ([string]::IsNullOrWhiteSpace($Name)) { return $null }
    return $Name.Trim()
}

function Get-LooseKey {
    # Uppercase + strip hyphens/underscores/spaces. Used ONLY for the
    # "possible match" review bucket, never for Onboarded status.
    param([string]$Name)
    if ([string]::IsNullOrWhiteSpace($Name)) { return $null }
    $upper = $Name.Trim().ToUpperInvariant()
    return ($upper -replace '[-_\s]', '')
}

function Resolve-ArcStatusColumn {
    param(
        [string[]]$AvailableColumns,
        [string]$Preferred
    )
    if ($Preferred) {
        if ($AvailableColumns -contains $Preferred) { return $Preferred }
        Write-Warning "Specified -ArcStatusColumn '$Preferred' was not found in List B. Falling back to auto-detection."
    }
    $candidates = @(
        'ArcAgentStatus', 'Arc Agent Status', 'ARC AGENT STATUS', 'Agent Status', 'AgentStatus',
        'Status', 'Machine Status', 'MachineStatus',
        'Connectivity Status', 'ConnectivityStatus',
        'Azure Arc Machine - Status', 'Arc Status', 'ArcStatus'
    )
    foreach ($c in $candidates) {
        if ($AvailableColumns -contains $c) { return $c }
    }
    return $null
}

# ---------------------------------------------------------------------------
# Load source data (read-only)
# ---------------------------------------------------------------------------

Write-Host "Loading vCenter inventory (List A): $VCenterCsvPath"
$vCenterVMs = Import-Csv -Path $VCenterCsvPath

Write-Host "Loading Azure Arc inventory (List B): $ArcCsvPath"
$arcVMs = Import-Csv -Path $ArcCsvPath

if (-not $vCenterVMs -or $vCenterVMs.Count -eq 0) {
    throw "List A (vCenter) contains no rows: $VCenterCsvPath"
}
if (-not $arcVMs -or $arcVMs.Count -eq 0) {
    throw "List B (Azure Arc) contains no rows: $ArcCsvPath"
}

$vCenterColumns = $vCenterVMs[0].PSObject.Properties.Name
$arcColumns     = $arcVMs[0].PSObject.Properties.Name

if ($vCenterColumns -notcontains $VCenterNameColumn) {
    throw "Column '$VCenterNameColumn' not found in List A. Available columns: $($vCenterColumns -join ', ')"
}
if ($arcColumns -notcontains $ArcNameColumn) {
    throw "Column '$ArcNameColumn' not found in List B. Available columns: $($arcColumns -join ', ')"
}

$resolvedArcStatusColumn = Resolve-ArcStatusColumn -AvailableColumns $arcColumns -Preferred $ArcStatusColumn
if ($resolvedArcStatusColumn) {
    Write-Host "Using '$resolvedArcStatusColumn' from List B as the Arc agent status column."
} else {
    Write-Warning "No Arc agent status column found in List B. ArcAgentStatus will be left blank for onboarded VMs instead of being invented."
}

if (-not $OutputPath) {
    $OutputPath = Join-Path (Split-Path -Parent $VCenterCsvPath) 'vCenter-Inventory-With-Arc-Status.csv'
}

# ---------------------------------------------------------------------------
# Build lookups from List B (Azure Arc)
# ---------------------------------------------------------------------------

# Exact key -> list of Arc rows (trimmed, case-INsensitive comparer)
$arcExactIndex = [System.Collections.Generic.Dictionary[string, System.Collections.Generic.List[object]]]::new([StringComparer]::OrdinalIgnoreCase)
# Loose key -> list of Arc rows (case/format-insensitive; already uppercased by Get-LooseKey)
$arcLooseIndex = [System.Collections.Generic.Dictionary[string, System.Collections.Generic.List[object]]]::new([StringComparer]::Ordinal)

$arcBlankNameCount = 0

foreach ($row in $arcVMs) {
    $rawName = $row.$ArcNameColumn
    $exactKey = Get-ExactKey $rawName
    $looseKey = Get-LooseKey $rawName

    if (-not $exactKey) {
        $arcBlankNameCount++
        continue
    }

    if (-not $arcExactIndex.ContainsKey($exactKey)) {
        $arcExactIndex[$exactKey] = [System.Collections.Generic.List[object]]::new()
    }
    $arcExactIndex[$exactKey].Add($row)

    if (-not $arcLooseIndex.ContainsKey($looseKey)) {
        $arcLooseIndex[$looseKey] = [System.Collections.Generic.List[object]]::new()
    }
    $arcLooseIndex[$looseKey].Add($row)
}

# Report duplicate names inside List B itself
$arcDuplicateGroups = $arcExactIndex.GetEnumerator() | Where-Object { $_.Value.Count -gt 1 }
if ($arcDuplicateGroups) {
    Write-Warning "List B (Azure Arc) contains duplicate VM names: $((($arcDuplicateGroups | ForEach-Object { $_.Key }) -join ', '))"
}

# ---------------------------------------------------------------------------
# Compare List A against List B and build enriched output
# ---------------------------------------------------------------------------

$results = [System.Collections.Generic.List[object]]::new()

$blankNameCount        = 0
$exactMatchCount       = 0
$notOnboardedCount     = 0
$possibleMatchCount    = 0
$vCenterNameSeen        = [System.Collections.Generic.Dictionary[string, int]]::new([StringComparer]::OrdinalIgnoreCase)

foreach ($vmRow in $vCenterVMs) {
    $rawName  = $vmRow.$VCenterNameColumn
    $exactKey = Get-ExactKey $rawName
    $looseKey = Get-LooseKey $rawName

    # Track duplicate VM names within List A
    if ($exactKey) {
        if ($vCenterNameSeen.ContainsKey($exactKey)) {
            $vCenterNameSeen[$exactKey]++
        } else {
            $vCenterNameSeen[$exactKey] = 1
        }
    }

    $onboardingStatus = 'Not-Onboarded'
    $agentStatus      = 'N/A'
    $possibleMatch    = ''
    $matchType        = 'No Match'

    if (-not $exactKey) {
        # Blank/null VM name - cannot be matched, handled safely (no crash, clearly labeled)
        $blankNameCount++
        $onboardingStatus = 'Not-Onboarded'
        $agentStatus      = 'N/A'
        $matchType        = 'No Match'
    }
    elseif ($arcExactIndex.ContainsKey($exactKey)) {
        # Exact match found
        $exactMatchCount++
        $onboardingStatus = 'Onboarded'
        $matchType        = 'Exact'

        $matchedArcRows = $arcExactIndex[$exactKey]
        if ($resolvedArcStatusColumn) {
            $statuses = $matchedArcRows | ForEach-Object { $_.$resolvedArcStatusColumn } | Where-Object { $_ } | Select-Object -Unique
            $agentStatus = if ($statuses) { $statuses -join ' / ' } else { '' }
        } else {
            $agentStatus = ''
        }
    }
    else {
        # No exact match - check loose/possible match
        $notOnboardedCount++
        $onboardingStatus = 'Not-Onboarded'
        $agentStatus      = 'N/A'

        if ($looseKey -and $arcLooseIndex.ContainsKey($looseKey)) {
            $possibleMatchCount++
            $matchType = 'Possible'
            $candidateNames = $arcLooseIndex[$looseKey] | ForEach-Object { $_.$ArcNameColumn } | Select-Object -Unique
            $possibleMatch = $candidateNames -join ', '
        }
    }

    # Clone the original row and append new columns, preserving all original data
    $enriched = [ordered]@{}
    foreach ($p in $vmRow.PSObject.Properties) {
        $enriched[$p.Name] = $p.Value
    }
    $enriched['ArcOnboardingStatus'] = $onboardingStatus
    $enriched['ArcAgentStatus']      = $agentStatus
    $enriched['PossibleArcMatch']    = $possibleMatch
    $enriched['MatchType']           = $matchType

    $results.Add([pscustomobject]$enriched)
}

$duplicateVCenterNames = $vCenterNameSeen.GetEnumerator() | Where-Object { $_.Value -gt 1 }

# ---------------------------------------------------------------------------
# Export (original files are never touched)
# ---------------------------------------------------------------------------

$results | Export-Csv -Path $OutputPath -NoTypeInformation -Encoding UTF8

# ---------------------------------------------------------------------------
# Summary
# ---------------------------------------------------------------------------

Write-Host ""
Write-Host "===================== SUMMARY =====================" -ForegroundColor Cyan
Write-Host ("Total vCenter VMs (List A rows)      : {0}" -f $vCenterVMs.Count)
Write-Host ("Exact Arc matches (Onboarded)        : {0}" -f $exactMatchCount)
Write-Host ("Not onboarded                        : {0}" -f $notOnboardedCount)
Write-Host ("Possible matches (manual review)     : {0}" -f $possibleMatchCount)
Write-Host ("Blank/null VM names in List A         : {0}" -f $blankNameCount)
Write-Host ("Duplicate VM names in List A          : {0}" -f ($duplicateVCenterNames | Measure-Object).Count)
if ($duplicateVCenterNames) {
    foreach ($d in $duplicateVCenterNames) {
        Write-Host ("    - '{0}' appears {1} times" -f $d.Key, $d.Value)
    }
}
Write-Host ("Duplicate VM names in List B (Arc)    : {0}" -f ($arcDuplicateGroups | Measure-Object).Count)
if ($arcDuplicateGroups) {
    foreach ($d in $arcDuplicateGroups) {
        Write-Host ("    - '{0}' appears {1} times" -f $d.Key, $d.Value.Count)
    }
}
Write-Host ("Blank/null VM names in List B (Arc)   : {0}" -f $arcBlankNameCount)
Write-Host "====================================================" -ForegroundColor Cyan
Write-Host ""
Write-Host "Output written to: $OutputPath" -ForegroundColor Green
Write-Host "Source files were not modified." -ForegroundColor Green
