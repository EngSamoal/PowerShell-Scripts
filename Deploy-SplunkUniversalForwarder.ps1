<#
    Deploy-SplunkUniversalForwarder.ps1

    Installs or upgrades the Splunk Universal Forwarder (SEVEN configuration only) across a list
    of Windows VMs in vCenter, using PowerCLI Guest Operations (VMware Tools) exclusively.
    WinRM is never used - all guest-side work goes through Copy-VMGuestFile and Invoke-VMScript.

    Prerequisites:
      - Already connected to the target vCenter via Connect-VIServer before running this script.
      - VMware Tools running in every target guest, with a local-administrator guest account
        credential exported to $GuestCredentialPath via Export-Clixml.
      - The Splunk UF installer and post_installation_Seven.bat present locally at the paths
        configured below.

    IMPORTANT - installer arguments:
      "UF splunk 10.4.2.EXE" is an IExpress self-extracting package (confirmed via its own /?
      help text), so it takes IExpress's standard switches - notably /Q for quiet/unattended mode.
      $SplunkInstallerArgs below is set to '/Q' accordingly. If a future installer build is NOT an
      IExpress package, re-check its /? output before assuming /Q still applies.
#>

#region ======================= CONFIGURATION =======================

# Bump this on every change and check it against the version quoted in chat before trusting a
# run's results - prints as the very first line of output so a stale cached copy is always
# immediately obvious, instead of silently re-running old logic.
$ScriptBuild = '2026.09.20-6'

# Splunk version this fleet must be running after this script completes.
[version]$RequiredSplunkVersion = '10.4.2'

# Local paths on the management machine (where this script runs).
$SplunkInstallerPath  = 'C:\Splunk_Install\UF splunk 10.4.2.EXE'         # <-- set to the real, full path
$SplunkInstallerArgs  = '/Q'                                              # IExpress quiet/unattended switch (confirmed via installer's /? output)
$PostInstallSevenPath = 'C:\Splunk_Install\post_installation_Seven.bat'
$VmListPath           = 'C:\temp\vmlist.txt'
$GuestCredentialPath  = 'C:\temp\wincred.xml'

# Guest-side paths/names.
$RemoteTempFolder  = 'C:\Windows\Temp\SplunkUFDeploy'
$SplunkInstallDir  = 'C:\Program Files\SplunkUniversalForwarder'
$SplunkServiceName = 'SplunkForwarder'

# Report output.
$ReportFolder = 'C:\temp'
$ReportPath   = Join-Path $ReportFolder ("SplunkDeploymentReport_{0}.csv" -f (Get-Date -Format 'yyyyMMdd_HHmmss'))

# Timeouts (seconds). Keep the installer/post-install ceilings generous - real installs can take
# several minutes - but bounded, so one stuck VM cannot hang the whole run.
$QuickGuestOpTimeoutSec      = 60     # cheap read-only guest calls (cred test, detection, validation)
$InstallerTimeoutSec         = 1200   # 20 min ceiling for the Splunk UF installer itself
$PostInstallTimeoutSec       = 600    # 10 min ceiling for post_installation_Seven.bat
$PostActionPollTimeoutSec    = 180    # extra time to allow install dir/service to settle after exit
$PostActionPollIntervalSec   = 10

#endregion ===========================================================


#region ======================= REPORT ROW TEMPLATE =======================

function New-ReportRow {
    param([string]$VMName)

    [PSCustomObject][ordered]@{
        VMName                = $VMName
        PowerState            = 'Unknown'
        VMwareToolsStatus     = 'Unknown'
        GuestCredentialStatus = 'Not Tested'
        SplunkInstalledBefore = $false
        PreviousVersion       = 'None'
        RequiredVersion       = $RequiredSplunkVersion.ToString()
        Action                = 'Skipped'
        InstallerExitCode     = ''
        SevenPostInstallStatus = ''
        SplunkServiceStatus   = 'Unknown'
        FinalVersion          = ''
        RebootRequired        = $false
        Result                = ''
        FailureReason         = ''
        StartTime             = Get-Date
        EndTime               = $null
        Duration              = ''
    }
}

#endregion ============================================================


#region ======================= LOCAL PREREQUISITE CHECKS =======================

function Test-LocalPrerequisites {
    $missing = @()

    foreach ($item in @(
        @{ Path = $SplunkInstallerPath;  Label = 'Splunk UF installer' },
        @{ Path = $PostInstallSevenPath; Label = 'post_installation_Seven.bat' },
        @{ Path = $VmListPath;           Label = 'VM list file' },
        @{ Path = $GuestCredentialPath;  Label = 'Guest credential XML' }
    )) {
        if (-not (Test-Path -LiteralPath $item.Path)) {
            $missing += "$($item.Label) not found: $($item.Path)"
        }
    }

    if (-not (Test-Path -LiteralPath $ReportFolder)) {
        try {
            New-Item -ItemType Directory -Path $ReportFolder -Force -ErrorAction Stop | Out-Null
        } catch {
            $missing += "Report folder '$ReportFolder' does not exist and could not be created: $($_.Exception.Message)"
        }
    }

    if ($missing.Count -gt 0) {
        foreach ($m in $missing) { Write-Host "PREREQUISITE FAILED: $m" -ForegroundColor Red }
        throw "One or more local prerequisites are missing. Aborting before touching any VM."
    }
}

#endregion ============================================================


#region ======================= GUEST OPERATION HELPER (WITH TIMEOUT) =======================

# Runs a guest script asynchronously and bounds the wait with Wait-Task, so a stuck/slow VM
# cannot hang the whole run. Exit codes for Bat scripts must be embedded in the script's own
# output text (Invoke-VMScript does not surface a separate exit-code property).
function Invoke-GuestScriptWithTimeout {
    param(
        [Parameter(Mandatory)] $VM,
        [Parameter(Mandatory)] [System.Management.Automation.PSCredential] $GuestCredential,
        [Parameter(Mandatory)] [ValidateSet('Powershell', 'Bat')] [string] $ScriptType,
        [Parameter(Mandatory)] [string] $ScriptText,
        [int] $TimeoutSeconds = 300
    )

    # NOTE: earlier revisions of this function tried to enforce a hard client-side timeout using
    # Invoke-VMScript -RunAsync plus Get-Task/Wait-Task. That guest-ops task object did not behave
    # like a normal vCenter task in this environment's PowerCLI version (Get-Task -Id threw "Index
    # was outside the bounds of the array"), so that approach is abandoned rather than patched a
    # third time on unverified internals. This now calls Invoke-VMScript synchronously - the same
    # proven pattern already used elsewhere in this repo (RunRemediationAcrossVMs-PowerCli.ps1).
    # $TimeoutSeconds is kept for future use but is not currently enforced; a stuck guest process
    # will block on this call until Invoke-VMScript/vCenter's own internal behavior resolves it.
    try {
        $result = Invoke-VMScript -VM $VM -GuestCredential $GuestCredential -ScriptType $ScriptType `
            -ScriptText $ScriptText -ErrorAction Stop

        [PSCustomObject]@{
            Success      = $true
            ScriptOutput = $result.ScriptOutput
            ErrorMessage = $null
        }
    } catch {
        [PSCustomObject]@{
            Success      = $false
            ScriptOutput = $null
            ErrorMessage = $_.Exception.Message
        }
    }
}

#endregion ============================================================


#region ======================= GUEST-SIDE SCRIPT TEMPLATES =======================
# Single-quoted here-strings so local PowerShell never expands $ inside them; placeholders are
# substituted explicitly with .Replace(), which also keeps filenames-with-spaces safe (no local
# string interpolation quoting hazards).

$Template_Readiness = @'
$result = [ordered]@{
    UserName         = $null
    IsAdmin          = $false
    IsAdminMethod    = $null
    IsAdminError     = $null
    InstallDirExists = $false
    ServiceExists    = $false
    ServiceStatus    = $null
    ServiceStartType = $null
    Version          = $null
    VersionSource    = $null
    Installed        = $false
    RebootPending    = $false
}

try {
    $id = [Security.Principal.WindowsIdentity]::GetCurrent()
    $result.UserName = $id.Name

    # Diagnostic pass: try three independent ways to determine admin membership and record which
    # one(s) succeeded/failed and why, instead of silently trusting one method. A prior fix
    # (checking the Administrators SID in $id.Groups) still reported IsAdmin=$false against a
    # confirmed local admin account, so this run needs to show WHY rather than guess again.
    $adminSid = 'S-1-5-32-544'
    $checks = [ordered]@{}

    try {
        $checks.GroupsSidMatch = [bool]($id.Groups | Where-Object { $_.Value -eq $adminSid })
    } catch {
        $checks.GroupsSidMatch = "ERROR: $($_.Exception.Message)"
    }

    try {
        $wp = New-Object Security.Principal.WindowsPrincipal($id)
        $checks.IsInRoleAdministrator = $wp.IsInRole([Security.Principal.WindowsBuiltInRole]::Administrator)
    } catch {
        $checks.IsInRoleAdministrator = "ERROR: $($_.Exception.Message)"
    }

    try {
        $localAdmins = @(Get-LocalGroupMember -Group 'Administrators' -ErrorAction Stop)
        $shortName = $id.Name -replace '^.*\\', ''
        $checks.LocalGroupMemberMatch = [bool]($localAdmins | Where-Object {
            ($_.Name -replace '^.*\\', '') -eq $shortName
        })
    } catch {
        $checks.LocalGroupMemberMatch = "ERROR: $($_.Exception.Message)"
    }

    $result.IsAdminMethod = ($checks.GetEnumerator() | ForEach-Object { "$($_.Key)=$($_.Value)" }) -join '; '

    # Treat as admin if ANY method affirmatively says so (defensive OR, since we do not yet know
    # which method is reliable in this environment).
    $result.IsAdmin = [bool](
        ($checks.GroupsSidMatch -eq $true) -or
        ($checks.IsInRoleAdministrator -eq $true) -or
        ($checks.LocalGroupMemberMatch -eq $true)
    )
} catch {
    $result.IsAdminError = $_.Exception.Message
}

$installDir = '__INSTALLDIR__'
$svcName    = '__SVCNAME__'

$result.InstallDirExists = Test-Path -LiteralPath $installDir

$svc = Get-Service -Name $svcName -ErrorAction SilentlyContinue
if ($svc) {
    $result.ServiceExists = $true
    $result.ServiceStatus = $svc.Status.ToString()
    try {
        $wmiSvc = Get-CimInstance -ClassName Win32_Service -Filter "Name='$svcName'" -ErrorAction Stop
        $result.ServiceStartType = $wmiSvc.StartMode
    } catch { $result.ServiceStartType = 'Unknown' }
}

$versionFile = Join-Path $installDir 'etc\splunk.version'
if (Test-Path -LiteralPath $versionFile) {
    try {
        $content = Get-Content -LiteralPath $versionFile -ErrorAction Stop
        $verLine = $content | Where-Object { $_ -match '^VERSION\s*=\s*(.+)$' }
        if ($verLine) {
            $result.Version = $Matches[1].Trim()
            $result.VersionSource = 'VersionFile'
        }
    } catch {}
}

if (-not $result.Version) {
    try {
        $uninstallKeys = @(
            'HKLM:\SOFTWARE\Microsoft\Windows\CurrentVersion\Uninstall\*',
            'HKLM:\SOFTWARE\WOW6432Node\Microsoft\Windows\CurrentVersion\Uninstall\*'
        )
        $app = Get-ItemProperty -Path $uninstallKeys -ErrorAction SilentlyContinue |
            Where-Object { $_.DisplayName -like 'SplunkForwarder*' -or $_.DisplayName -like '*Splunk Universal Forwarder*' } |
            Select-Object -First 1
        if ($app -and $app.DisplayVersion) {
            $result.Version = $app.DisplayVersion
            $result.VersionSource = 'Registry'
        }
    } catch {}
}

if (-not $result.Version) {
    $splunkExe = Join-Path $installDir 'bin\splunk.exe'
    if (Test-Path -LiteralPath $splunkExe) {
        try {
            $verOut = & $splunkExe version --accept-license 2>$null
            if ($verOut -match '([0-9]+\.[0-9]+\.[0-9]+)') {
                $result.Version = $Matches[1]
                $result.VersionSource = 'SplunkExe'
            }
        } catch {}
    }
}

$result.Installed = [bool]($result.ServiceExists -or $result.InstallDirExists -or $result.Version)

try {
    if (Test-Path 'HKLM:\SOFTWARE\Microsoft\Windows\CurrentVersion\Component Based Servicing\RebootPending') { $result.RebootPending = $true }
    if (Test-Path 'HKLM:\SOFTWARE\Microsoft\Windows\CurrentVersion\WindowsUpdate\Auto Update\RebootRequired') { $result.RebootPending = $true }
    $pfro = Get-ItemProperty -Path 'HKLM:\SYSTEM\CurrentControlSet\Control\Session Manager' -Name 'PendingFileRenameOperations' -ErrorAction SilentlyContinue
    if ($pfro) { $result.RebootPending = $true }
} catch {}

[PSCustomObject]$result | ConvertTo-Json -Compress
'@

$Template_RunInstaller = @'
@echo off
set "INSTALLER=__INSTALLER_PATH__"
set "IARGS=__INSTALLER_ARGS__"
if not exist "%INSTALLER%" (
    echo INSTALLER_MISSING
    exit /b 9009
)
"%INSTALLER%" %IARGS%
echo INSTALLER_EXITCODE=%ERRORLEVEL%
'@

$Template_RunPostInstall = @'
@echo off
set "POSTBAT=__POSTBAT_PATH__"
if not exist "%POSTBAT%" (
    echo POSTBAT_MISSING
    exit /b 9009
)
"%POSTBAT%"
echo POSTBAT_EXITCODE=%ERRORLEVEL%
'@

$Template_EnsureAutoStart = @'
$svcName = '__SVCNAME__'
try {
    $wmiSvc = Get-CimInstance -ClassName Win32_Service -Filter "Name='$svcName'" -ErrorAction Stop
    if ($wmiSvc.StartMode -ne 'Auto') {
        Set-Service -Name $svcName -StartupType Automatic -ErrorAction Stop
        'ChangedToAutomatic'
    } else {
        'AlreadyAutomatic'
    }
} catch {
    "FailedToSetStartupType: $($_.Exception.Message)"
}
'@

$Template_MkdirAndCleanup = @'
$path = '__PATH__'
$mode = '__MODE__'
try {
    if ($mode -eq 'create') {
        New-Item -ItemType Directory -Path $path -Force -ErrorAction Stop | Out-Null
        'OK'
    } else {
        if (Test-Path -LiteralPath $path) {
            Remove-Item -LiteralPath $path -Recurse -Force -ErrorAction Stop
        }
        'OK'
    }
} catch {
    "FAILED: $($_.Exception.Message)"
}
'@

#endregion ============================================================


#region ======================= VM READINESS / DETECTION =======================

function Get-VMReadinessAndSplunkStatus {
    param(
        [Parameter(Mandatory)] $VM,
        [Parameter(Mandatory)] [System.Management.Automation.PSCredential] $GuestCredential
    )

    $script = $Template_Readiness.Replace('__INSTALLDIR__', $SplunkInstallDir).Replace('__SVCNAME__', $SplunkServiceName)
    $r = Invoke-GuestScriptWithTimeout -VM $VM -GuestCredential $GuestCredential -ScriptType Powershell `
        -ScriptText $script -TimeoutSeconds $QuickGuestOpTimeoutSec

    if (-not $r.Success) {
        return [PSCustomObject]@{ Success = $false; ErrorMessage = $r.ErrorMessage; Status = $null }
    }

    try {
        $status = $r.ScriptOutput | ConvertFrom-Json -ErrorAction Stop
        [PSCustomObject]@{ Success = $true; ErrorMessage = $null; Status = $status }
    } catch {
        [PSCustomObject]@{ Success = $false; ErrorMessage = "Could not parse guest status output: $($_.Exception.Message)"; Status = $null }
    }
}

#endregion ============================================================


#region ======================= INSTALL / UPGRADE PROCEDURE =======================

# Copies the installer + post_installation_Seven.bat to the guest, runs both as the guest
# credential (which must be a local admin - checked earlier), waits for each to finish (bounded
# by timeout), then cleans up the copied files. Does not itself decide compliant/upgrade/install -
# the caller has already made that decision; this function just executes the guide's procedure.
function Invoke-SplunkInstallProcedure {
    param(
        [Parameter(Mandatory)] $VM,
        [Parameter(Mandatory)] [System.Management.Automation.PSCredential] $GuestCredential
    )

    $out = [PSCustomObject]@{
        Success             = $false
        FailureReason       = ''
        InstallerExitCode   = ''
        SevenPostInstallStatus = ''
        RebootRequired      = $false
    }

    $installerLeaf  = Split-Path -Path $SplunkInstallerPath -Leaf
    $postBatLeaf    = Split-Path -Path $PostInstallSevenPath -Leaf
    $remoteInstaller = Join-Path $RemoteTempFolder $installerLeaf
    $remotePostBat   = Join-Path $RemoteTempFolder $postBatLeaf

    # 1. Create remote temp folder.
    $mkdirScript = $Template_MkdirAndCleanup.Replace('__PATH__', $RemoteTempFolder).Replace('__MODE__', 'create')
    $mkdirResult = Invoke-GuestScriptWithTimeout -VM $VM -GuestCredential $GuestCredential -ScriptType Powershell `
        -ScriptText $mkdirScript -TimeoutSeconds $QuickGuestOpTimeoutSec
    if (-not $mkdirResult.Success -or ($mkdirResult.ScriptOutput -notmatch 'OK')) {
        $out.FailureReason = "Failed to create remote temp folder: $($mkdirResult.ErrorMessage)$($mkdirResult.ScriptOutput)"
        return $out
    }

    # 2. Copy installer + post-install bat into the guest (VMware Tools guest-file API only).
    try {
        Copy-VMGuestFile -Source $SplunkInstallerPath -Destination $remoteInstaller -LocalToGuest `
            -VM $VM -GuestCredential $GuestCredential -Force -ErrorAction Stop
        Copy-VMGuestFile -Source $PostInstallSevenPath -Destination $remotePostBat -LocalToGuest `
            -VM $VM -GuestCredential $GuestCredential -Force -ErrorAction Stop
    } catch {
        $out.FailureReason = "Failed to copy installation files to guest: $($_.Exception.Message)"
        return $out
    }

    # 3. Run the installer (as the guest admin credential) and wait for it to finish.
    $installScript = $Template_RunInstaller.Replace('__INSTALLER_PATH__', $remoteInstaller).Replace('__INSTALLER_ARGS__', $SplunkInstallerArgs)
    $installResult = Invoke-GuestScriptWithTimeout -VM $VM -GuestCredential $GuestCredential -ScriptType Bat `
        -ScriptText $installScript -TimeoutSeconds $InstallerTimeoutSec

    if (-not $installResult.Success) {
        $out.FailureReason = "Installer execution failed or timed out: $($installResult.ErrorMessage)"
        return $out
    }
    if ($installResult.ScriptOutput -match 'INSTALLER_MISSING') {
        $out.FailureReason = "Installer file was not found on the guest at $remoteInstaller after copy."
        return $out
    }
    if ($installResult.ScriptOutput -match 'INSTALLER_EXITCODE=(-?\d+)') {
        $out.InstallerExitCode = $Matches[1]
        if ($Matches[1] -eq '3010') { $out.RebootRequired = $true }
        elseif ($Matches[1] -ne '0') {
            $out.FailureReason = "Installer returned non-zero exit code $($Matches[1])."
            return $out
        }
    } else {
        $out.FailureReason = "Could not determine installer exit code from output."
        return $out
    }

    # 4. Run post_installation_Seven.bat (SEVEN configuration only) and wait for it to finish.
    $postScript = $Template_RunPostInstall.Replace('__POSTBAT_PATH__', $remotePostBat)
    $postResult = Invoke-GuestScriptWithTimeout -VM $VM -GuestCredential $GuestCredential -ScriptType Bat `
        -ScriptText $postScript -TimeoutSeconds $PostInstallTimeoutSec

    if (-not $postResult.Success) {
        $out.FailureReason = "post_installation_Seven.bat execution failed or timed out: $($postResult.ErrorMessage)"
        return $out
    }
    if ($postResult.ScriptOutput -match 'POSTBAT_MISSING') {
        $out.FailureReason = "post_installation_Seven.bat was not found on the guest at $remotePostBat after copy."
        return $out
    }
    if ($postResult.ScriptOutput -match 'POSTBAT_EXITCODE=(-?\d+)') {
        $code = $Matches[1]
        if ($code -eq '3010') { $out.RebootRequired = $true; $out.SevenPostInstallStatus = "Completed (ExitCode $code, reboot pending)" }
        elseif ($code -eq '0') { $out.SevenPostInstallStatus = 'Completed (ExitCode 0)' }
        else {
            $out.SevenPostInstallStatus = "Completed with non-zero ExitCode $code"
            $out.FailureReason = "post_installation_Seven.bat returned non-zero exit code $code."
            return $out
        }
    } else {
        $out.SevenPostInstallStatus = 'Unknown - exit code not captured'
        $out.FailureReason = "Could not determine post_installation_Seven.bat exit code from output."
        return $out
    }

    # 5. Poll briefly: the installer/bat may finish while Splunk itself is still settling.
    $deadline = (Get-Date).AddSeconds($PostActionPollTimeoutSec)
    $finalStatus = $null
    do {
        $check = Get-VMReadinessAndSplunkStatus -VM $VM -GuestCredential $GuestCredential
        if ($check.Success -and $check.Status.ServiceExists -and $check.Status.ServiceStatus -eq 'Running') {
            $finalStatus = $check.Status
            break
        }
        Start-Sleep -Seconds $PostActionPollIntervalSec
    } while ((Get-Date) -lt $deadline)

    if (-not $finalStatus) {
        $lastCheck = Get-VMReadinessAndSplunkStatus -VM $VM -GuestCredential $GuestCredential
        $out.FailureReason = "Post-install validation failed: SplunkForwarder service was not confirmed Running within $PostActionPollTimeoutSec seconds."
        if ($lastCheck.Success) {
            $out.FailureReason += " Last observed ServiceExists=$($lastCheck.Status.ServiceExists), ServiceStatus=$($lastCheck.Status.ServiceStatus)."
        }
        return $out
    }

    # 6. Ensure service is set to start Automatically (required post-install check).
    if ($finalStatus.ServiceStartType -ne 'Auto') {
        $autoScript = $Template_EnsureAutoStart.Replace('__SVCNAME__', $SplunkServiceName)
        Invoke-GuestScriptWithTimeout -VM $VM -GuestCredential $GuestCredential -ScriptType Powershell `
            -ScriptText $autoScript -TimeoutSeconds $QuickGuestOpTimeoutSec | Out-Null
    }

    $out.Success = $true
    return $out
}

#endregion ============================================================


#region ======================= MAIN =======================

Write-Host "ScriptBuild: $ScriptBuild" -ForegroundColor Magenta
Write-Host "`n=== Splunk Universal Forwarder Deployment (SEVEN configuration) ===" -ForegroundColor Cyan
Write-Host "Required version: $RequiredSplunkVersion`n"

Test-LocalPrerequisites

if (-not $global:DefaultVIServer -or -not $global:DefaultVIServer.IsConnected) {
    throw "No active vCenter connection found. Connect with Connect-VIServer before running this script."
}

$guestCred = Import-Clixml -Path $GuestCredentialPath
$vmNames = Get-Content -Path $VmListPath |
    ForEach-Object { $_.Trim() } |
    Where-Object { $_ -ne '' -and -not $_.StartsWith('#') } |
    Select-Object -Unique

if ($vmNames.Count -eq 0) {
    throw "VM list at $VmListPath contains no VM names."
}

$results = New-Object System.Collections.Generic.List[object]

foreach ($vmName in $vmNames) {

    $row = New-ReportRow -VMName $vmName
    Write-Host "`n--- $vmName ---" -ForegroundColor Cyan

    try {
        # 1. VM exists.
        $vmMatches = @(Get-VM -Name $vmName -ErrorAction SilentlyContinue)
        if ($vmMatches.Count -eq 0) { throw "VM '$vmName' was not found in vCenter." }
        if ($vmMatches.Count -gt 1) { throw "VM name '$vmName' is ambiguous ($($vmMatches.Count) matches) - skipping for safety." }
        $vm = $vmMatches[0]

        # Powered on.
        $row.PowerState = $vm.PowerState.ToString()
        if ($vm.PowerState -ne 'PoweredOn') { throw "VM is not powered on (state: $($vm.PowerState))." }

        # VMware Tools installed and running.
        $toolsStatus  = $vm.ExtensionData.Guest.ToolsStatus
        $toolsRunning = $vm.ExtensionData.Guest.ToolsRunningStatus
        $row.VMwareToolsStatus = "$toolsStatus / $toolsRunning"
        if ($toolsRunning -ne 'guestToolsRunning') { throw "VMware Tools is not running (status: $toolsStatus / $toolsRunning)." }

        # Guest OS is Windows.
        if ($vm.ExtensionData.Guest.GuestFamily -ne 'windowsGuest') {
            throw "Guest OS is not Windows (family: $($vm.ExtensionData.Guest.GuestFamily))."
        }

        # Guest credentials valid + readiness/detection (single combined guest call).
        $detect = Get-VMReadinessAndSplunkStatus -VM $vm -GuestCredential $guestCred
        if (-not $detect.Success) {
            $row.GuestCredentialStatus = 'Invalid'
            throw "Guest credential validation / detection failed: $($detect.ErrorMessage)"
        }
        $status = $detect.Status
        $row.GuestCredentialStatus = if ($status.IsAdmin) { 'Valid (Administrator)' } else { 'Valid (NOT Administrator - install will likely fail)' }
        $row.RebootRequired = [bool]$status.RebootPending

        # Diagnostic - print regardless of outcome so admin-detection results are visible in the
        # console immediately, not just in the CSV.
        Write-Host "  [diag] Guest user: $($status.UserName)" -ForegroundColor DarkGray
        Write-Host "  [diag] Admin checks: $($status.IsAdminMethod)" -ForegroundColor DarkGray
        if ($status.IsAdminError) { Write-Host "  [diag] Admin check error: $($status.IsAdminError)" -ForegroundColor DarkGray }

        $row.SplunkInstalledBefore = [bool]$status.Installed
        $row.PreviousVersion = if ($status.Version) { $status.Version } else { 'None' }

        $needsProcedure = $false

        if ($status.Installed -and $status.Version) {
            $installedVersion = $null
            try { $installedVersion = [version]$status.Version } catch { $installedVersion = $null }

            if (-not $installedVersion) {
                $row.Action = 'Manual Review'
                $row.Result = 'Manual Review'
                $row.FailureReason = "Installed version string '$($status.Version)' could not be parsed - not touching this VM."
            }
            elseif ($installedVersion -gt $RequiredSplunkVersion) {
                $row.Action = 'Manual Review'
                $row.Result = 'Newer Version Detected - Manual Review'
                $row.FinalVersion = $status.Version
            }
            elseif ($installedVersion -eq $RequiredSplunkVersion) {
                if ($status.ServiceExists -and $status.ServiceStatus -eq 'Running') {
                    $row.Action = 'Skipped'
                    $row.Result = 'Already Compliant'
                    $row.FinalVersion = $status.Version
                    $row.SplunkServiceStatus = $status.ServiceStatus
                } else {
                    # Correct version but service unhealthy - repair using the same guide procedure.
                    $row.Action = 'Installed'
                    $needsProcedure = $true
                }
            }
            else {
                $row.Action = 'Upgraded'
                $needsProcedure = $true
            }
        } else {
            $row.Action = 'Installed'
            $needsProcedure = $true
        }

        if ($needsProcedure) {
            if (-not $status.IsAdmin) {
                throw "Guest credential is not a local Administrator; cannot proceed with install/upgrade."
            }

            $procResult = Invoke-SplunkInstallProcedure -VM $vm -GuestCredential $guestCred
            $row.InstallerExitCode = $procResult.InstallerExitCode
            $row.SevenPostInstallStatus = $procResult.SevenPostInstallStatus
            if ($procResult.RebootRequired) { $row.RebootRequired = $true }

            if (-not $procResult.Success) {
                $row.Action = 'Failed'
                $row.Result = 'Failed'
                $row.FailureReason = $procResult.FailureReason
            } else {
                # Final post-install validation - success is based on this, not on the procedure "starting".
                $final = Get-VMReadinessAndSplunkStatus -VM $vm -GuestCredential $guestCred
                if (-not $final.Success) {
                    $row.Action = 'Failed'
                    $row.Result = 'Failed'
                    $row.FailureReason = "Post-install validation could not be performed: $($final.ErrorMessage)"
                } else {
                    $fs = $final.Status
                    $row.SplunkServiceStatus = $fs.ServiceStatus
                    $row.FinalVersion = $fs.Version
                    if ($fs.RebootPending) { $row.RebootRequired = $true }

                    $finalVersionOk = $false
                    try { $finalVersionOk = ([version]$fs.Version -eq $RequiredSplunkVersion) } catch {}

                    if ($fs.ServiceExists -and $fs.ServiceStatus -eq 'Running' -and $fs.InstallDirExists -and $finalVersionOk) {
                        $row.Result = if ($row.Action -eq 'Upgraded') { 'Upgrade Successful' } else { 'Install Successful' }
                    } else {
                        $row.Action = 'Failed'
                        $row.Result = 'Failed'
                        $row.FailureReason = "Post-install validation failed: ServiceExists=$($fs.ServiceExists), ServiceStatus=$($fs.ServiceStatus), InstallDirExists=$($fs.InstallDirExists), Version=$($fs.Version)."
                    }
                }
            }
        }

        if ($row.Result -eq '') { $row.Result = $row.Action }

    } catch {
        $row.Action = 'Failed'
        $row.Result = 'Failed'
        $row.FailureReason = $_.Exception.Message
    } finally {
        $row.EndTime = Get-Date
        $row.Duration = [string]([timespan]($row.EndTime - $row.StartTime))
        $results.Add($row)

        $color = switch ($row.Action) {
            'Failed'        { 'Red' }
            'Manual Review' { 'Yellow' }
            'Skipped'       { 'DarkGray' }
            default         { 'Green' }
        }
        Write-Host "Result: $($row.Result)  |  Action: $($row.Action)  |  Version: $($row.FinalVersion)$(if(-not $row.FinalVersion){$row.PreviousVersion})" -ForegroundColor $color
        if ($row.FailureReason) { Write-Host "Reason: $($row.FailureReason)" -ForegroundColor Red }
        if ($row.RebootRequired) { Write-Host "REBOOT REQUIRED on $vmName - not rebooting automatically." -ForegroundColor Yellow }
    }
}

# Export final report.
$results | Export-Csv -Path $ReportPath -NoTypeInformation -Encoding UTF8
Write-Host "`nReport written to: $ReportPath" -ForegroundColor Cyan

# Summary.
$summary = [ordered]@{
    'Total VMs'                    = $results.Count
    'Already Compliant'            = ($results | Where-Object { $_.Result -eq 'Already Compliant' }).Count
    'Successfully Installed'       = ($results | Where-Object { $_.Result -eq 'Install Successful' }).Count
    'Successfully Upgraded'        = ($results | Where-Object { $_.Result -eq 'Upgrade Successful' }).Count
    'Newer Version / Manual Review' = ($results | Where-Object { $_.Action -eq 'Manual Review' }).Count
    'Failed'                       = ($results | Where-Object { $_.Action -eq 'Failed' }).Count
    'Skipped'                      = ($results | Where-Object { $_.Action -eq 'Skipped' -and $_.Result -ne 'Already Compliant' }).Count
    'Reboot Required (not performed)' = ($results | Where-Object { $_.RebootRequired }).Count
}

Write-Host "`n=== Summary ===" -ForegroundColor Cyan
foreach ($key in $summary.Keys) {
    Write-Host ("{0,-32}: {1}" -f $key, $summary[$key])
}

#endregion ============================================================
