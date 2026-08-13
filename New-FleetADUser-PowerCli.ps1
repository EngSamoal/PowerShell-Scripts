#Requires -Modules ActiveDirectory
<#
.SYNOPSIS
    Creates multiple Active Directory user accounts from a spreadsheet, using a saved
    AD admin credential - no interactive prompts, safe to run unattended.

.DESCRIPTION
    - Reads new-user details from an Excel file (C:\temp\new-AD-users.xlsx by default) via
      the ImportExcel module.
    - Authenticates to AD using a credential previously saved with Export-Clixml
      (C:\temp\ADcred.xml by default).
    - Per row: validates required fields, confirms the target OU exists, checks whether the
      SamAccountName already exists, and ONLY THEN creates the account.
    - NEVER modifies an existing account. If SamAccountName is already present in AD, that
      row is skipped and logged - no attribute is touched, no password is reset.
    - Passwords come directly from the "Password" column in the spreadsheet (fixed, not
      generated). Because of that, both the input spreadsheet and the Secrets CSV this script
      writes contain plaintext passwords - secure or delete both after the run.
    - Group membership (the "Groups" column) is applied as a separate step after the account
      is created. If the account is created but one or more groups fail (typo, group doesn't
      exist, etc.), the account is NOT rolled back - it is reported as CreatedWithWarnings so
      the partial state is visible rather than silently swallowed.
    - Every row ends up in exactly one of four result CSVs:
        1. CreatedSuccess
        2. CreatedWithWarnings          (account created, but group membership partially failed)
        3. SkippedAlreadyExists
        4. SkippedPrerequisiteFailed    (missing/invalid field, OU not found, DC unreachable,
                                          password policy rejection, etc.)
      A bucket produces no CSV file if it has zero rows in it, rather than a header-only file.
    - Full exception detail (message + inner exception chain) is captured to the log for
      every failure.

.NOTES
    Requires: RSAT "ActiveDirectory" PowerShell module, and the "ImportExcel" PowerShell
    module (Install-Module ImportExcel -Scope CurrentUser - no Microsoft Excel installation
    required, it reads the .xlsx file format directly).
    Run against a handful of test rows first before a full batch.
#>

# ============================== CONFIGURATION ==============================
# 1) AD admin credential used to create the accounts (must have "Create User objects" rights
#    on every target OU, and rights to modify membership of every group referenced).
$ADCredPath = "C:\temp\ADcred.xml"

# To create the saved credential file ONE TIME (run manually, not part of this script):
#   Get-Credential | Export-Clixml -Path "C:\temp\ADcred.xml"

if (-not (Test-Path $ADCredPath)) {
    throw "AD credential file not found at $ADCredPath. Create it first with: Get-Credential | Export-Clixml -Path '$ADCredPath'"
}
$ADCred = Import-Clixml -Path $ADCredPath

# 2) Optional: pin a specific domain controller. Leave blank ("") to let the ActiveDirectory
#    module locate one automatically.
$DomainController = ""

# 3) Spreadsheet of new users. See the column layout in the accompanying description -
#    required columns: SamAccountName, FirstName, LastName, Password, DisplayName, Groups.
#    Optional columns: OUPath, UserPrincipalName, Title, Department, ChangePasswordAtLogon,
#    Enabled.
#    OUPath: leave blank to use the domain's default "Users" container (resolved below from
#    AD itself, since that container's DN is domain-specific and it is NOT an OU object - it's
#    a special "Container" object, so Get-ADOrganizationalUnit can't be used to validate it).
$UsersXlsxPath = "C:\temp\new-AD-users.xlsx"

# 4) Used to build UserPrincipalName for rows that don't supply one explicitly.
$DefaultUPNSuffix = "yourdomain.com"

# 5) Output/logging paths
$OutputFolder = "C:\temp\fleet"
if (-not (Test-Path -Path $OutputFolder)) {
    New-Item -Path $OutputFolder -ItemType Directory -Force | Out-Null
}

$LogPath              = Join-Path $OutputFolder "FleetADUserCreation_Log.log"
$SecretsPath          = Join-Path $OutputFolder "FleetADUserCreation_Secrets.csv"
$CreatedSuccessPath   = Join-Path $OutputFolder "FleetAD_01_CreatedSuccess.csv"
$CreatedWarningsPath  = Join-Path $OutputFolder "FleetAD_02_CreatedWithWarnings.csv"
$SkippedExistsPath    = Join-Path $OutputFolder "FleetAD_03_SkippedAlreadyExists.csv"
$SkippedPrereqPath    = Join-Path $OutputFolder "FleetAD_04_SkippedPrerequisiteFailed.csv"
# NOTE: file names are static (no timestamp) - each run OVERWRITES the CSVs from the previous
# run. The log file is appended to (not overwritten), so each run's entries stack on top of
# earlier ones - every log line has its own timestamp, so runs stay distinguishable.
# =============================================================================

function Write-Log {
    param([string]$Message, [string]$Level = "INFO")
    $line = "[{0}] [{1}] {2}" -f (Get-Date -Format "yyyy-MM-dd HH:mm:ss"), $Level, $Message
    Add-Content -Path $LogPath -Value $line
    Write-Host $line
}

function Get-FullExceptionDetail {
    param($ErrorRecord)
    $sb = New-Object System.Text.StringBuilder
    $ex = $ErrorRecord.Exception
    $depth = 0
    while ($ex) {
        [void]$sb.AppendLine("  [$depth] Type: $($ex.GetType().FullName)")
        [void]$sb.AppendLine("  [$depth] Message: $($ex.Message)")
        $ex = $ex.InnerException
        $depth++
    }
    if ($ErrorRecord.ErrorDetails) {
        [void]$sb.AppendLine("  ErrorDetails: $($ErrorRecord.ErrorDetails.Message)")
    }
    [void]$sb.AppendLine("  CategoryInfo: $($ErrorRecord.CategoryInfo.ToString())")
    return $sb.ToString().TrimEnd()
}

# Excel boolean cells can come back as actual booleans, "TRUE"/"FALSE" strings, or blank -
# normalize all three instead of assuming one representation.
function Get-BoolValue {
    param($Value, [bool]$Default)
    if ($null -eq $Value -or "$Value".Trim() -eq "") { return $Default }
    $text = "$Value".Trim()
    try { return [bool]::Parse($text) } catch {}
    if ($text -eq "1") { return $true }
    if ($text -eq "0") { return $false }
    return $Default
}

function Get-TrimmedValue {
    param($Value)
    if ($null -eq $Value) { return "" }
    return "$Value".Trim()
}

# --- Pre-flight checks (script-level, not per-row) ---
if (-not (Get-Module -Name ActiveDirectory -ListAvailable)) {
    throw "ActiveDirectory module not found. Install RSAT: Add-WindowsCapability -Online -Name 'Rsat.ActiveDirectory.DS-LDS.Tools~~~~0.0.1.0'"
}
if (-not (Get-Module -Name ImportExcel -ListAvailable)) {
    throw "ImportExcel module not found. Install with: Install-Module ImportExcel -Scope CurrentUser"
}
Import-Module ActiveDirectory -ErrorAction Stop
Import-Module ImportExcel -ErrorAction Stop

if (-not (Test-Path $UsersXlsxPath)) {
    throw "Users spreadsheet not found at $UsersXlsxPath."
}

$rows = Import-Excel -Path $UsersXlsxPath
if (-not $rows -or $rows.Count -eq 0) {
    throw "No rows found in $UsersXlsxPath."
}

$adSplat = @{ Credential = $ADCred }
if ($DomainController) { $adSplat["Server"] = $DomainController }

# Resolve the domain's actual default "Users" container DN from AD itself (varies per domain,
# e.g. "CN=Users,DC=corp,DC=local") - used as the fallback for any row that leaves OUPath blank.
try {
    $DefaultUsersContainer = (Get-ADDomain @adSplat -ErrorAction Stop).UsersContainer
} catch {
    throw "Could not resolve the domain's default Users container (is the DC reachable / credential valid?). $(Get-FullExceptionDetail $_)"
}
Write-Log "Default OU for rows with a blank OUPath: $DefaultUsersContainer"

Write-Log "=== Fleet AD user creation started. Row count: $($rows.Count) | Spreadsheet: $UsersXlsxPath ==="
Write-Log "Passwords are being read from the spreadsheet in plaintext - secure or delete '$UsersXlsxPath' and the Secrets CSV this run produces once you're done." "WARN"

$CreatedSuccess  = @()
$CreatedWarnings = @()
$SkippedExists   = @()
$SkippedPrereq   = @()
$Secrets         = @()

foreach ($r in $rows) {

    $sam = Get-TrimmedValue $r.SamAccountName
    $row = [PSCustomObject]@{
        SamAccountName = $sam
        Reason         = ""
        Timestamp      = Get-Date -Format "yyyy-MM-dd HH:mm:ss"
    }

    try {
        # --- Required-field validation ---
        $firstName = Get-TrimmedValue $r.FirstName
        $lastName  = Get-TrimmedValue $r.LastName
        $ouPath    = Get-TrimmedValue $r.OUPath
        $password  = Get-TrimmedValue $r.Password

        $displayName = Get-TrimmedValue $r.DisplayName
        $groupsRaw   = Get-TrimmedValue $r.Groups

        if ($sam -eq "")         { throw "PREREQ_FAIL: SamAccountName is blank." }
        if ($sam.Length -gt 20)  { throw "PREREQ_FAIL: SamAccountName '$sam' is longer than 20 characters (AD limit)." }
        if ($firstName -eq "")   { throw "PREREQ_FAIL: FirstName is blank." }
        if ($lastName -eq "")    { throw "PREREQ_FAIL: LastName is blank." }
        if ($password -eq "")    { throw "PREREQ_FAIL: Password is blank." }
        if ($displayName -eq "") { throw "PREREQ_FAIL: DisplayName is blank." }
        if ($groupsRaw -eq "")   { throw "PREREQ_FAIL: Groups is blank - at least one group is required." }

        if ($ouPath -eq "") { $ouPath = $DefaultUsersContainer }

        $upn = Get-TrimmedValue $r.UserPrincipalName
        if ($upn -eq "") { $upn = "$sam@$DefaultUPNSuffix" }

        $title       = Get-TrimmedValue $r.Title
        $department  = Get-TrimmedValue $r.Department
        $changePwd   = Get-BoolValue -Value $r.ChangePasswordAtLogon -Default $true
        $enabled     = Get-BoolValue -Value $r.Enabled -Default $true

        $groupNames = $groupsRaw -split ';' | ForEach-Object { $_.Trim() } | Where-Object { $_ -ne "" }
        if ($groupNames.Count -eq 0) { throw "PREREQ_FAIL: Groups value '$groupsRaw' did not contain any usable group names." }

        # --- Prerequisite checks - run BEFORE any creation attempt ---
        # Get-ADObject (not Get-ADOrganizationalUnit) - the default "Users" container is a
        # "Container" object, not an "organizationalUnit" object, so the OU-specific cmdlet
        # would reject it even though New-ADUser -Path accepts it fine.
        try {
            Get-ADObject -Identity $ouPath @adSplat -ErrorAction Stop | Out-Null
        } catch {
            throw "PREREQ_FAIL: Target path '$ouPath' could not be found or is not reachable. $(Get-FullExceptionDetail $_)"
        }

        $existing = $null
        try {
            $existing = Get-ADUser -Filter "SamAccountName -eq '$sam'" @adSplat -ErrorAction Stop
        } catch {
            throw "PREREQ_FAIL: Could not query AD for existing account '$sam' (is the DC reachable / credential valid?). $(Get-FullExceptionDetail $_)"
        }

        if ($existing) {
            $row.Reason = "Account already exists in AD at '$($existing.DistinguishedName)' - left untouched."
            $SkippedExists += $row
            Write-Log "$sam -> SKIPPED (already exists)"
            continue
        }

        # --- Create the account ---
        $securePwd = ConvertTo-SecureString -String $password -AsPlainText -Force
        $newUserParams = @{
            SamAccountName        = $sam
            Name                  = $displayName
            GivenName             = $firstName
            Surname               = $lastName
            DisplayName           = $displayName
            UserPrincipalName     = $upn
            Path                  = $ouPath
            AccountPassword       = $securePwd
            Enabled               = $enabled
            ChangePasswordAtLogon = $changePwd
            PassThru              = $true
        }
        if ($title -ne "")      { $newUserParams["Title"] = $title }
        if ($department -ne "") { $newUserParams["Department"] = $department }

        try {
            New-ADUser @newUserParams @adSplat -ErrorAction Stop | Out-Null
        } catch {
            throw "PREREQ_FAIL: New-ADUser failed (often a password-policy rejection or insufficient rights on this OU). $(Get-FullExceptionDetail $_)"
        }

        # Account exists now - from here on we report Created (with or without warnings),
        # never SkippedPrerequisiteFailed, since the account is real and must not be silently
        # unreported.
        $Secrets += [PSCustomObject]@{ SamAccountName = $sam; UserPrincipalName = $upn; Password = $password }

        $groupFailures = @()
        foreach ($g in $groupNames) {
            try {
                Add-ADGroupMember -Identity $g -Members $sam @adSplat -ErrorAction Stop
            } catch {
                $groupFailures += "$g ($($_.Exception.Message))"
                Write-Log "$sam -> failed to add to group '$g'. $(Get-FullExceptionDetail $_)" "WARN"
            }
        }

        if ($groupFailures.Count -eq 0) {
            $row.Reason = "Account created" + $(if ($groupNames.Count -gt 0) { " and added to group(s): $($groupNames -join ', ')" } else { "" })
            $CreatedSuccess += $row
            Write-Log "$sam -> CREATED"
        } else {
            $row.Reason = "Account created, but failed to add to group(s): $($groupFailures -join '; '). Account itself is fine - group membership needs manual follow-up."
            $CreatedWarnings += $row
            Write-Log "$sam -> CREATED WITH WARNINGS (group membership incomplete)" "WARN"
        }
    }
    catch {
        $row.Reason = $_.Exception.Message
        $SkippedPrereq += $row
        Write-Log "$sam -> SKIPPED (prerequisite/execution failure): $(Get-FullExceptionDetail $_)" "ERROR"
    }
}

# --- Reporting ---
function Export-ReportCsv {
    param($Data, [string]$Path)
    if ($Data.Count -gt 0) {
        $Data | Export-Csv -Path $Path -NoTypeInformation
    } else {
        Write-Log "No records for $Path - file not created (empty result set)."
    }
}

Export-ReportCsv -Data $CreatedSuccess  -Path $CreatedSuccessPath
Export-ReportCsv -Data $CreatedWarnings -Path $CreatedWarningsPath
Export-ReportCsv -Data $SkippedExists   -Path $SkippedExistsPath
Export-ReportCsv -Data $SkippedPrereq   -Path $SkippedPrereqPath

if ($Secrets.Count -gt 0) {
    $Secrets | Export-Csv -Path $SecretsPath -NoTypeInformation
    Write-Log "Per-account passwords (as supplied in the spreadsheet) written to $SecretsPath - SECURE OR DELETE THIS FILE AFTER HANDOFF." "WARN"
}

Write-Log "=== SUMMARY: $($CreatedSuccess.Count) created, $($CreatedWarnings.Count) created with group-membership warnings, $($SkippedExists.Count) skipped (already existed), $($SkippedPrereq.Count) skipped (prerequisite failure), out of $($rows.Count) total ==="
Write-Log "Files: $CreatedSuccessPath | $CreatedWarningsPath | $SkippedExistsPath | $SkippedPrereqPath | Log: $LogPath"

Write-Host "`n--- CREATED SUCCESS ($($CreatedSuccess.Count)) ---"
$CreatedSuccess | Format-Table -AutoSize
Write-Host "`n--- CREATED WITH WARNINGS ($($CreatedWarnings.Count)) ---"
$CreatedWarnings | Format-Table -AutoSize
Write-Host "`n--- SKIPPED: ALREADY EXISTS ($($SkippedExists.Count)) ---"
$SkippedExists | Format-Table -AutoSize
Write-Host "`n--- SKIPPED: PREREQUISITE FAILED ($($SkippedPrereq.Count)) ---"
$SkippedPrereq | Format-Table -AutoSize
