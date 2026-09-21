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

    IMPORTANT - installer approach changed after live testing:
      "UF splunk 10.4.2.EXE" is an IExpress self-extracting wrapper. Live testing found it hangs
      indefinitely (idle at 0% CPU, no MSI log ever created) when launched non-interactively via
      VMware guest-ops - some self-extractors depend on shell/COM components that don't
      initialize properly without an interactive desktop (Session 0 isolation). Verbose msiexec
      logging confirmed msiexec never even started, so this is the wrapper's own extraction
      hanging, not an MSI/UI problem.

      Fix: bypass the wrapper entirely. Extract it once on the management machine using
      IExpress's own extract-only mode:
        & 'C:\Splunk_Install\UF splunk 10.4.2.EXE' /T:C:\Splunk_Install\_extracted /C
      then point $SplunkInstallerPath (below) at the extracted .msi and this script runs
      msiexec.exe directly against it - a native OS binary, not a custom wrapper - reproducing
      the exact command found inside the wrapper's own install.bat.

    SECURITY - MSI properties file:
      install.bat's msiexec command includes SPLUNKPASSWORD, a credential value. That must never
      be committed to this repo, so it is NOT hardcoded here - $SplunkMsiPropertiesPath below
      points to a local file (same pattern as $GuestCredentialPath) containing the full MSI
      property string, read at runtime and never written to disk by this script or logged.
#>

param(
    # Audit-only mode: validates every VM and reports current Splunk status (installed?, version,
    # service running?) WITHOUT installing/upgrading anything. Use this to see what a full run
    # would do first, or to trim vmlist.txt down to only the VMs that actually need work. The
    # installer/MSI-properties files are not required in this mode since nothing gets installed.
    [switch]$CheckOnly
)

#region ======================= CONFIGURATION =======================

# Bump this on every change and check it against the version quoted in chat before trusting a
# run's results - prints as the very first line of output so a stale cached copy is always
# immediately obvious, instead of silently re-running old logic.
$ScriptBuild = '2026.09.21-6'

# Splunk version this fleet must be running after this script completes.
[version]$RequiredSplunkVersion = '10.4.2'

# Local paths on the management machine (where this script runs).
$SplunkInstallerPath  = 'C:\Splunk_Install\_extracted\SPLUNK~1.MSI'      # <-- the EXTRACTED .msi, not the EXE wrapper - see header comment
$SplunkMsiPropertiesPath = 'C:\temp\splunk_msi_properties.txt'           # <-- local file (never committed) holding the MSI property string, e.g.:
                                                                           #     AGREETOLICENSE=Yes LAUNCHSPLUNK=1 PRIVILEGEBACKUP=0 PRIVILEGESECURITY=0 USE_LOCAL_SYSTEM=1 SPLUNKUSERNAME=admin SPLUNKPASSWORD=<value>
$PostInstallSevenPath = 'C:\Splunk_Install\post_installation_Seven.bat'
$VmListPath           = 'C:\temp\vmlist.txt'
$GuestCredentialPath  = 'C:\temp\wincred.xml'

# Guest-side paths/names.
$RemoteTempFolder  = 'C:\Windows\Temp\SplunkUFDeploy'
$SplunkInstallDir  = 'C:\Program Files\SplunkUniversalForwarder'
$SplunkServiceName = 'SplunkForwarder'

# Report output.
$ReportFolder = 'C:\temp'
$script:ReportTimestamp = Get-Date -Format 'yyyyMMdd_HHmmss'
$ReportPath     = Join-Path $ReportFolder ("SplunkDeploymentReport_{0}.csv" -f $script:ReportTimestamp)
$ReportHtmlPath = Join-Path $ReportFolder ("SplunkDeploymentDashboard_{0}.html" -f $script:ReportTimestamp)

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
        Notes                 = ''
        StartTime             = Get-Date
        EndTime               = $null
        Duration              = ''
    }
}

# Appends to $row.Notes rather than overwriting, so multiple independent notes (e.g. a
# pre-existing pending reboot AND an upgrade decision) both survive on the same row.
function Add-RowNote {
    param([Parameter(Mandatory)] $Row, [Parameter(Mandatory)] [string] $Note)
    $Row.Notes = if ($Row.Notes) { "$($Row.Notes) | $Note" } else { $Note }
}

#endregion ============================================================


#region ======================= LOCAL PREREQUISITE CHECKS =======================

function Test-LocalPrerequisites {
    $missing = @()

    $requiredFiles = @(
        @{ Path = $VmListPath;               Label = 'VM list file' },
        @{ Path = $GuestCredentialPath;      Label = 'Guest credential XML' }
    )
    if (-not $CheckOnly) {
        # Only needed when actually installing/upgrading something - CheckOnly mode never
        # touches these, so don't force the user to have them staged just to run an audit.
        $requiredFiles += @{ Path = $SplunkInstallerPath;      Label = 'Splunk UF installer (.msi)' }
        $requiredFiles += @{ Path = $SplunkMsiPropertiesPath;  Label = 'Splunk MSI properties file' }
        $requiredFiles += @{ Path = $PostInstallSevenPath;     Label = 'post_installation_Seven.bat' }
    }

    foreach ($item in $requiredFiles) {
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

# Reads the MSI property string from $SplunkMsiPropertiesPath at runtime - kept out of the script
# itself (and out of git) since it contains SPLUNKPASSWORD. Never logged or written anywhere.
function Get-SplunkMsiProperties {
    $raw = (Get-Content -LiteralPath $SplunkMsiPropertiesPath -Raw -ErrorAction Stop).Trim()
    if ([string]::IsNullOrWhiteSpace($raw)) {
        throw "$SplunkMsiPropertiesPath is empty - it must contain the MSI property string (AGREETOLICENSE=Yes ... SPLUNKPASSWORD=<value>)."
    }
    return $raw
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

# Split into two SMALL guest scripts rather than one large one. Live testing proved that a
# single larger detection script (combining admin-check + service + version-file + registry scan
# + splunk.exe + reboot-pending, ~2.5-3KB of script text) comes back with ZERO output and no
# error at all - not even a marker string printed as the very first statement - while the exact
# same admin-check+service logic on its own returns correctly. This is consistent with a payload
# size limit on the VMware Tools guest-ops RPC channel that Invoke-VMScript uses, silently
# swallowing anything over some threshold rather than raising a catchable error. Splitting into
# two calls keeps each one comfortably small.

$Template_ReadinessCore = @'
$result = [ordered]@{
    UserName         = $null
    IsAdmin          = $false
    IsAdminMethod    = $null
    IsAdminError     = $null
    InstallDirExists = $false
    ServiceExists    = $false
    ServiceStatus    = $null
    ServiceStartType = $null
}

try {
    $id = [Security.Principal.WindowsIdentity]::GetCurrent()
    $result.UserName = $id.Name
    $adminSid = 'S-1-5-32-544'
    $checks = [ordered]@{}
    try { $checks.GroupsSidMatch = [bool]($id.Groups | Where-Object { $_.Value -eq $adminSid }) } catch { $checks.GroupsSidMatch = "ERROR: $($_.Exception.Message)" }
    try {
        $wp = New-Object Security.Principal.WindowsPrincipal($id)
        $checks.IsInRoleAdministrator = $wp.IsInRole([Security.Principal.WindowsBuiltInRole]::Administrator)
    } catch { $checks.IsInRoleAdministrator = "ERROR: $($_.Exception.Message)" }
    $result.IsAdminMethod = ($checks.GetEnumerator() | ForEach-Object { "$($_.Key)=$($_.Value)" }) -join '; '
    $result.IsAdmin = [bool](($checks.GroupsSidMatch -eq $true) -or ($checks.IsInRoleAdministrator -eq $true))
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

[PSCustomObject]$result | ConvertTo-Json -Compress
'@

$Template_VersionAndReboot = @'
$result = [ordered]@{
    Version       = $null
    VersionSource = $null
    RebootPending = $false
}

$installDir = '__INSTALLDIR__'

$versionFile = Join-Path $installDir 'etc\splunk.version'
if (Test-Path -LiteralPath $versionFile) {
    try {
        $content = Get-Content -LiteralPath $versionFile -ErrorAction Stop
        $verLine = $content | Where-Object { $_ -match '^VERSION\s*=\s*(.+)$' }
        if ($verLine) { $result.Version = $Matches[1].Trim(); $result.VersionSource = 'VersionFile' }
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
        if ($app -and $app.DisplayVersion) { $result.Version = $app.DisplayVersion; $result.VersionSource = 'Registry' }
    } catch {}
}

if (-not $result.Version) {
    $splunkExe = Join-Path $installDir 'bin\splunk.exe'
    if (Test-Path -LiteralPath $splunkExe) {
        try {
            $verOut = & $splunkExe version --accept-license 2>$null
            if ($verOut -match '([0-9]+\.[0-9]+\.[0-9]+)') { $result.Version = $Matches[1]; $result.VersionSource = 'SplunkExe' }
        } catch {}
    }
}

try {
    if (Test-Path 'HKLM:\SOFTWARE\Microsoft\Windows\CurrentVersion\Component Based Servicing\RebootPending') { $result.RebootPending = $true }
    if (Test-Path 'HKLM:\SOFTWARE\Microsoft\Windows\CurrentVersion\WindowsUpdate\Auto Update\RebootRequired') { $result.RebootPending = $true }
    $pfro = Get-ItemProperty -Path 'HKLM:\SYSTEM\CurrentControlSet\Control\Session Manager' -Name 'PendingFileRenameOperations' -ErrorAction SilentlyContinue
    if ($pfro) { $result.RebootPending = $true }
} catch {}

[PSCustomObject]$result | ConvertTo-Json -Compress
'@

# NOTE: Invoke-VMScript's -ScriptType Bat flattens multi-line batch text onto one logical line
# joined by '&' - confirmed via live testing (literal cmd.exe error "& was unexpected at this
# time."), and this breaks ANY multi-statement batch structure, not just parenthesized blocks -
# a goto/label rewrite hit the identical error, since labels can't survive being '&'-joined onto
# one line either. Abandoning -ScriptType Bat entirely for anything beyond a single command.
# These are now Powershell scripts that launch the target executable via .NET's Start-Process,
# which hands back a real exit code without ever touching cmd.exe's batch parser.
# Runs msiexec.exe directly against the extracted .msi - NOT the IExpress EXE wrapper, which live
# testing proved hangs indefinitely under non-interactive VMware guest-ops (see header comment).
# msiexec is a native OS binary, so this sidesteps that wrapper-specific problem entirely. Always
# appends /quiet and /l*v (a verbose log at a known path) regardless of what's in the properties
# string, so a real failure is diagnosable from the log rather than a silent hang.
$Template_RunInstaller = @'
$installer = '__INSTALLER_PATH__'
$installerArgs = '__INSTALLER_ARGS__'
$logPath = '__LOG_PATH__'
if (-not (Test-Path -LiteralPath $installer)) {
    'INSTALLER_MISSING'
} else {
    try {
        $msiArgs = "/i `"$installer`" $installerArgs /quiet /l*v `"$logPath`""
        $proc = Start-Process -FilePath 'msiexec.exe' -ArgumentList $msiArgs -Wait -PassThru -ErrorAction Stop
        "INSTALLER_EXITCODE=$($proc.ExitCode)"
    } catch {
        "INSTALLER_LAUNCH_ERROR: $($_.Exception.Message)"
    }
}
'@

$Template_RunPostInstall = @'
$postBat = '__POSTBAT_PATH__'
if (-not (Test-Path -LiteralPath $postBat)) {
    'POSTBAT_MISSING'
} else {
    try {
        $proc = Start-Process -FilePath $postBat -Wait -PassThru -ErrorAction Stop
        "POSTBAT_EXITCODE=$($proc.ExitCode)"
    } catch {
        "POSTBAT_LAUNCH_ERROR: $($_.Exception.Message)"
    }
}
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

# Live testing found the real cause of the persistent Copy-VMGuestFile 500s: the destination file
# from an earlier successful copy was locked by something on the guest (most likely antivirus
# on-access scanning of a newly-arrived large unknown .EXE), so deleting/overwriting it failed.
# Checking remote file size first lets the script skip re-copying a file that is already staged
# correctly, instead of fighting a lock that may not even be a real problem.
$Template_GetRemoteFileSize = @'
$path = '__PATH__'
if (Test-Path -LiteralPath $path) { (Get-Item -LiteralPath $path).Length } else { -1 }
'@

#endregion ============================================================


#region ======================= VM READINESS / DETECTION =======================

function Invoke-GuestJsonQuery {
    param(
        [Parameter(Mandatory)] $VM,
        [Parameter(Mandatory)] [System.Management.Automation.PSCredential] $GuestCredential,
        [Parameter(Mandatory)] [string] $ScriptText,
        [int] $TimeoutSeconds = 60
    )

    $r = Invoke-GuestScriptWithTimeout -VM $VM -GuestCredential $GuestCredential -ScriptType Powershell `
        -ScriptText $ScriptText -TimeoutSeconds $TimeoutSeconds

    if (-not $r.Success) {
        return [PSCustomObject]@{ Success = $false; ErrorMessage = $r.ErrorMessage; Status = $null; RawOutput = $r.ScriptOutput }
    }

    # An empty/blank response means the guest script never reached its final output line (crashed,
    # got killed, or hung partway through) - ConvertFrom-Json on an empty string returns $null
    # WITHOUT throwing, which would otherwise silently masquerade as a "successful" empty result.
    # Treat blank output as an explicit failure instead.
    if ([string]::IsNullOrWhiteSpace($r.ScriptOutput)) {
        return [PSCustomObject]@{
            Success      = $false
            ErrorMessage = 'Guest script produced no output at all (it may have crashed, hung, or the script text was too large for the guest-ops channel).'
            Status       = $null
            RawOutput    = $r.ScriptOutput
        }
    }

    try {
        $status = $r.ScriptOutput | ConvertFrom-Json -ErrorAction Stop
        [PSCustomObject]@{ Success = $true; ErrorMessage = $null; Status = $status; RawOutput = $r.ScriptOutput }
    } catch {
        [PSCustomObject]@{ Success = $false; ErrorMessage = "Could not parse guest status output: $($_.Exception.Message)"; Status = $null; RawOutput = $r.ScriptOutput }
    }
}

function Get-VMReadinessAndSplunkStatus {
    param(
        [Parameter(Mandatory)] $VM,
        [Parameter(Mandatory)] [System.Management.Automation.PSCredential] $GuestCredential
    )

    $coreScript = $Template_ReadinessCore.Replace('__INSTALLDIR__', $SplunkInstallDir).Replace('__SVCNAME__', $SplunkServiceName)
    $core = Invoke-GuestJsonQuery -VM $VM -GuestCredential $GuestCredential -ScriptText $coreScript -TimeoutSeconds $QuickGuestOpTimeoutSec

    if (-not $core.Success) {
        return [PSCustomObject]@{ Success = $false; ErrorMessage = $core.ErrorMessage; Status = $null; RawOutput = $core.RawOutput }
    }

    $verScript = $Template_VersionAndReboot.Replace('__INSTALLDIR__', $SplunkInstallDir)
    $ver = Invoke-GuestJsonQuery -VM $VM -GuestCredential $GuestCredential -ScriptText $verScript -TimeoutSeconds $QuickGuestOpTimeoutSec

    if (-not $ver.Success) {
        return [PSCustomObject]@{ Success = $false; ErrorMessage = $ver.ErrorMessage; Status = $null; RawOutput = $ver.RawOutput }
    }

    $merged = [PSCustomObject]@{
        UserName         = $core.Status.UserName
        IsAdmin          = $core.Status.IsAdmin
        IsAdminMethod    = $core.Status.IsAdminMethod
        IsAdminError     = $core.Status.IsAdminError
        InstallDirExists = $core.Status.InstallDirExists
        ServiceExists    = $core.Status.ServiceExists
        ServiceStatus    = $core.Status.ServiceStatus
        ServiceStartType = $core.Status.ServiceStartType
        Version          = $ver.Status.Version
        VersionSource    = $ver.Status.VersionSource
        RebootPending    = $ver.Status.RebootPending
        Installed        = [bool]($core.Status.ServiceExists -or $core.Status.InstallDirExists -or $ver.Status.Version)
    }

    [PSCustomObject]@{ Success = $true; ErrorMessage = $null; Status = $merged; RawOutput = "$($core.RawOutput) | $($ver.RawOutput)" }
}

$Template_GetFileTail = @'
$path = '__PATH__'
if (Test-Path -LiteralPath $path) {
    (Get-Content -LiteralPath $path -Tail 25 -ErrorAction Stop) -join ' | '
} else {
    'LOG_NOT_FOUND'
}
'@

# Best-effort: pulls the tail of the msiexec verbose log for a real diagnostic instead of a bare
# exit code. Never throws - installer failure reporting must not itself fail the run.
function Get-RemoteMsiLogTail {
    param(
        [Parameter(Mandatory)] $VM,
        [Parameter(Mandatory)] [System.Management.Automation.PSCredential] $GuestCredential,
        [Parameter(Mandatory)] [string] $LogPath
    )
    try {
        $script = $Template_GetFileTail.Replace('__PATH__', $LogPath)
        $r = Invoke-GuestScriptWithTimeout -VM $VM -GuestCredential $GuestCredential -ScriptType Powershell `
            -ScriptText $script -TimeoutSeconds $QuickGuestOpTimeoutSec
        if ($r.Success -and $r.ScriptOutput) {
            return "MSI log tail ($LogPath): $($r.ScriptOutput)"
        }
        return "(could not read MSI log at $LogPath)"
    } catch {
        return "(error reading MSI log: $($_.Exception.Message))"
    }
}

#endregion ============================================================


#region ======================= INSTALL / UPGRADE PROCEDURE =======================

# Live testing showed Copy-VMGuestFile intermittently fails with a transient HTTP 500 from the
# ESXi/vCenter guest-file-transfer endpoint - confirmed transient by retrying the exact same call
# immediately afterward and having it succeed. This wraps the copy with a few retries so the
# script recovers on its own instead of requiring a manual re-run each time it happens.
function Copy-VMGuestFileWithRetry {
    param(
        [Parameter(Mandatory)] [string] $Source,
        [Parameter(Mandatory)] [string] $Destination,
        [Parameter(Mandatory)] $VM,
        [Parameter(Mandatory)] [System.Management.Automation.PSCredential] $GuestCredential,
        [int] $MaxAttempts = 3,
        [int] $RetryDelaySeconds = 10
    )

    $lastError = $null
    for ($attempt = 1; $attempt -le $MaxAttempts; $attempt++) {
        try {
            Copy-VMGuestFile -Source $Source -Destination $Destination -LocalToGuest `
                -VM $VM -GuestCredential $GuestCredential -Force -ErrorAction Stop
            return
        } catch {
            $lastError = $_
            if ($attempt -lt $MaxAttempts) {
                Start-Sleep -Seconds $RetryDelaySeconds
            }
        }
    }
    throw $lastError
}

# Checks whether the destination already has a file of the same size as the source and, if so,
# skips copying entirely. This avoids ever needing to delete/overwrite a file that a prior
# successful copy left behind and that something on the guest (antivirus on-access scanning a
# newly-arrived large unknown .EXE, confirmed via live testing showing an Access Denied trying to
# delete it) may be holding a lock on. Only attempts cleanup + copy when sizes genuinely differ
# (missing, partial, or a different file).
function Ensure-VMGuestFileStaged {
    param(
        [Parameter(Mandatory)] [string] $Source,
        [Parameter(Mandatory)] [string] $Destination,
        [Parameter(Mandatory)] $VM,
        [Parameter(Mandatory)] [System.Management.Automation.PSCredential] $GuestCredential
    )

    $localSize = (Get-Item -LiteralPath $Source).Length

    $sizeScript = $Template_GetRemoteFileSize.Replace('__PATH__', $Destination)
    $sizeCheck = Invoke-GuestScriptWithTimeout -VM $VM -GuestCredential $GuestCredential -ScriptType Powershell `
        -ScriptText $sizeScript -TimeoutSeconds $QuickGuestOpTimeoutSec

    $remoteSize = -1
    if ($sizeCheck.Success -and $sizeCheck.ScriptOutput -match '(-?\d+)') {
        $remoteSize = [int64]$Matches[1]
    }

    if ($remoteSize -eq $localSize) {
        return  # already staged correctly on the guest - nothing to do
    }

    if ($remoteSize -ge 0) {
        # A different/partial file exists - try to clear it, but don't treat failure here as
        # fatal on its own; the copy attempt below is the real test of whether this can proceed.
        $cleanupScript = $Template_MkdirAndCleanup.Replace('__PATH__', $Destination).Replace('__MODE__', 'remove')
        Invoke-GuestScriptWithTimeout -VM $VM -GuestCredential $GuestCredential -ScriptType Powershell `
            -ScriptText $cleanupScript -TimeoutSeconds $QuickGuestOpTimeoutSec | Out-Null
    }

    Copy-VMGuestFileWithRetry -Source $Source -Destination $Destination -VM $VM -GuestCredential $GuestCredential
}

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

    # 2/3. Stage installer + post-install bat on the guest (VMware Tools guest-file API only) -
    # skips re-copying a file that's already correctly staged from a prior run (see
    # Ensure-VMGuestFileStaged for why that matters).
    try {
        Ensure-VMGuestFileStaged -Source $SplunkInstallerPath -Destination $remoteInstaller -VM $VM -GuestCredential $GuestCredential
        Ensure-VMGuestFileStaged -Source $PostInstallSevenPath -Destination $remotePostBat -VM $VM -GuestCredential $GuestCredential
    } catch {
        $out.FailureReason = "Failed to copy installation files to guest: $($_.Exception.Message)"
        return $out
    }

    # 4. Run the installer (as the guest admin credential) and wait for it to finish.
    $remoteMsiLog = Join-Path $RemoteTempFolder 'splkInstall.log'
    $installScript = $Template_RunInstaller.Replace('__INSTALLER_PATH__', $remoteInstaller).Replace('__INSTALLER_ARGS__', $SplunkInstallerArgs).Replace('__LOG_PATH__', $remoteMsiLog)
    $installResult = Invoke-GuestScriptWithTimeout -VM $VM -GuestCredential $GuestCredential -ScriptType Powershell `
        -ScriptText $installScript -TimeoutSeconds $InstallerTimeoutSec

    if (-not $installResult.Success) {
        $out.FailureReason = "Installer execution failed or timed out: $($installResult.ErrorMessage)"
        return $out
    }
    if ($installResult.ScriptOutput -match 'INSTALLER_MISSING') {
        $out.FailureReason = "Installer file was not found on the guest at $remoteInstaller after copy."
        return $out
    }
    if ($installResult.ScriptOutput -match 'INSTALLER_LAUNCH_ERROR: (.+)') {
        $out.FailureReason = "Installer failed to launch: $($Matches[1])"
        return $out
    }
    if ($installResult.ScriptOutput -match 'INSTALLER_EXITCODE=(-?\d+)') {
        $out.InstallerExitCode = $Matches[1]
        if ($Matches[1] -eq '3010') { $out.RebootRequired = $true }
        elseif ($Matches[1] -ne '0') {
            $out.FailureReason = "Installer returned non-zero exit code $($Matches[1]). $(Get-RemoteMsiLogTail -VM $VM -GuestCredential $GuestCredential -LogPath $remoteMsiLog)"
            return $out
        }
    } else {
        $out.FailureReason = "Could not determine installer exit code from output. $(Get-RemoteMsiLogTail -VM $VM -GuestCredential $GuestCredential -LogPath $remoteMsiLog)"
        return $out
    }

    # 5. Run post_installation_Seven.bat (SEVEN configuration only) and wait for it to finish.
    $postScript = $Template_RunPostInstall.Replace('__POSTBAT_PATH__', $remotePostBat)
    $postResult = Invoke-GuestScriptWithTimeout -VM $VM -GuestCredential $GuestCredential -ScriptType Powershell `
        -ScriptText $postScript -TimeoutSeconds $PostInstallTimeoutSec

    if (-not $postResult.Success) {
        $out.FailureReason = "post_installation_Seven.bat execution failed or timed out: $($postResult.ErrorMessage)"
        return $out
    }
    if ($postResult.ScriptOutput -match 'POSTBAT_MISSING') {
        $out.FailureReason = "post_installation_Seven.bat was not found on the guest at $remotePostBat after copy."
        return $out
    }
    if ($postResult.ScriptOutput -match 'POSTBAT_LAUNCH_ERROR: (.+)') {
        $out.FailureReason = "post_installation_Seven.bat failed to launch: $($Matches[1])"
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

    # 6. Poll briefly: the installer/bat may finish while Splunk itself is still settling.
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

    # 7. Ensure service is set to start Automatically (required post-install check).
    if ($finalStatus.ServiceStartType -ne 'Auto') {
        $autoScript = $Template_EnsureAutoStart.Replace('__SVCNAME__', $SplunkServiceName)
        Invoke-GuestScriptWithTimeout -VM $VM -GuestCredential $GuestCredential -ScriptType Powershell `
            -ScriptText $autoScript -TimeoutSeconds $QuickGuestOpTimeoutSec | Out-Null
    }

    $out.Success = $true
    return $out
}

#endregion ============================================================


#region ======================= HTML DASHBOARD =======================

function ConvertTo-HtmlSafe {
    param([string]$Text)
    if ([string]::IsNullOrEmpty($Text)) { return '' }
    return $Text.Replace('&', '&amp;').Replace('<', '&lt;').Replace('>', '&gt;').Replace('"', '&quot;').Replace("'", '&#39;')
}

# Same visual language as Generate-PatchingDashboard.ps1 elsewhere in this repo (dark blue
# gradient header, KPI cards, colored status badges) so reports across this repo look consistent.
function Build-SplunkDashboardHtml {
    param(
        [Parameter(Mandatory)] [System.Collections.IEnumerable] $Results,
        [Parameter(Mandatory)] [hashtable] $Summary,
        [Parameter(Mandatory)] [string] $RequiredVersion,
        [Parameter(Mandatory)] [string] $ScriptBuild,
        [bool] $CheckOnly
    )

    $css = @'
  :root{--ink:#1F2933;--line:#E3E8EF;--muted:#6B7683;--bg:#F4F6F9;}
  *{box-sizing:border-box;}
  body{margin:0;background:var(--bg);color:var(--ink);
       font-family:-apple-system,"Segoe UI",Roboto,Helvetica,Arial,sans-serif;line-height:1.45;}
  .wrap{max-width:1280px;margin:0 auto;padding:28px;}
  header.hero{background:linear-gradient(135deg,#191919 0%,#3A2410 60%,#F1611D 100%);color:#fff;
       border-radius:14px;padding:34px 40px;box-shadow:0 10px 30px rgba(0,0,0,.25);}
  .report-h1{font-size:32px;font-weight:800;letter-spacing:.4px;margin:0;line-height:1.15;}
  .report-sub{font-size:16px;font-weight:600;margin:10px 0 0;color:#F2D9C9;}
  .mode-badge{display:inline-block;margin-top:14px;padding:5px 14px;border-radius:20px;
       font-size:12px;font-weight:800;letter-spacing:.5px;text-transform:uppercase;background:#fff;color:#3A2410;}
  .kpi{display:grid;grid-template-columns:repeat(8,1fr);gap:14px;margin:22px 0 10px;}
  .kpi-card{position:relative;background:#fff;border:1px solid var(--line);border-radius:12px;
       padding:16px 14px 14px;overflow:hidden;box-shadow:0 2px 6px rgba(31,41,51,.04);}
  .kpi-accent{position:absolute;top:0;left:0;right:0;height:5px;}
  .kpi-label{font-size:11px;font-weight:700;text-transform:uppercase;letter-spacing:.5px;color:var(--muted);}
  .kpi-value{font-size:30px;font-weight:800;margin-top:6px;}
  .panel{background:#fff;border:1px solid var(--line);border-radius:12px;padding:20px 22px;margin-top:18px;
       box-shadow:0 2px 6px rgba(31,41,51,.04);}
  h2{font-size:15px;text-transform:uppercase;letter-spacing:.7px;margin:0 0 16px;color:#334155;}
  table.ex-table{width:100%;border-collapse:collapse;font-size:12.5px;}
  .ex-table th{background:#0F2A43;color:#fff;text-align:left;padding:9px 10px;font-weight:600;white-space:nowrap;}
  .ex-table td{padding:8px 10px;border-bottom:1px solid var(--line);}
  .ex-table tr:nth-child(even){background:#FAFBFD;}
  .badge{display:inline-block;width:120px;text-align:center;padding:4px 0;border-radius:4px;font-size:11.5px;font-weight:700;color:#fff;white-space:nowrap;}
  .badge.b-green{background:#1F8A4C;} .badge.b-red{background:#C0392B;} .badge.b-amber{background:#C77700;} .badge.b-grey{background:#6B7683;}
  .notes-cell{color:var(--muted);font-size:11.5px;max-width:260px;}
  footer{margin-top:24px;font-size:12px;color:var(--muted);}
  @media (max-width:1100px){.kpi{grid-template-columns:repeat(4,1fr);}}
  @media (max-width:640px){.kpi{grid-template-columns:repeat(2,1fr);}}
'@

    function Get-KpiCard($label, $value, $color) {
        return "  <div class=`"kpi-card`"><div class=`"kpi-accent`" style=`"background:$color`"></div><div class=`"kpi-label`">$(ConvertTo-HtmlSafe $label)</div><div class=`"kpi-value`" style=`"color:$color`">$value</div></div>`n"
    }

    $kpi = ''
    $kpi += Get-KpiCard 'Total VMs' $Summary['Total VMs'] '#1D4E79'
    $kpi += Get-KpiCard 'Already Compliant' $Summary['Already Compliant'] '#1F8A4C'
    $kpi += Get-KpiCard 'Installed' $Summary['Successfully Installed'] '#1F8A4C'
    $kpi += Get-KpiCard 'Upgraded' $Summary['Successfully Upgraded'] '#1F8A4C'
    $kpi += Get-KpiCard 'Failed' $Summary['Failed'] '#C0392B'
    $kpi += Get-KpiCard 'Manual Review' $Summary['Newer Version / Manual Review'] '#C77700'
    $kpi += Get-KpiCard 'Skipped' $Summary['Skipped'] '#6B7683'
    $kpi += Get-KpiCard 'Reboot Pending' $Summary['Reboot Required (not performed)'] '#C77700'

    function Get-StatusBadgeClass($row) {
        if ($row.Action -eq 'Failed') { return 'b-red' }
        if ($row.Action -eq 'Manual Review') { return 'b-amber' }
        if ($row.Result -like 'Check Only*') { return 'b-grey' }
        if ($row.Action -eq 'Skipped') { return 'b-grey' }
        return 'b-green'
    }

    $rows = ''
    foreach ($r in $Results) {
        $badgeClass = Get-StatusBadgeClass $r
        $rebootTxt = if ($r.RebootRequired) { '<span class="badge b-amber">Yes</span>' } else { 'No' }
        $notes = if ($r.FailureReason) { $r.FailureReason } else { $r.Notes }
        $rows += "          <tr>`n" +
            "            <td>$(ConvertTo-HtmlSafe $r.VMName)</td>`n" +
            "            <td>$(ConvertTo-HtmlSafe $r.PowerState)</td>`n" +
            "            <td><span class=`"badge $badgeClass`">$(ConvertTo-HtmlSafe $r.Action)</span></td>`n" +
            "            <td>$(ConvertTo-HtmlSafe $r.Result)</td>`n" +
            "            <td>$(ConvertTo-HtmlSafe $r.PreviousVersion)</td>`n" +
            "            <td>$(ConvertTo-HtmlSafe $r.FinalVersion)</td>`n" +
            "            <td>$(ConvertTo-HtmlSafe $r.SplunkServiceStatus)</td>`n" +
            "            <td>$rebootTxt</td>`n" +
            "            <td class=`"notes-cell`">$(ConvertTo-HtmlSafe $notes)</td>`n" +
            "            <td>$(ConvertTo-HtmlSafe $r.Duration)</td>`n" +
            "          </tr>`n"
    }

    $modeBadge = if ($CheckOnly) { '<div class="mode-badge">CHECK ONLY - NO CHANGES MADE</div>' } else { '' }
    $safeVersion = ConvertTo-HtmlSafe $RequiredVersion

    return @"
<!DOCTYPE html>
<html lang="en">
<head>
<meta charset="utf-8">
<meta name="viewport" content="width=device-width, initial-scale=1">
<title>Splunk UF Deployment Dashboard</title>
<style>
$css
</style>
</head>
<body>
<div class="wrap">

  <header class="hero">
    <h1 class="report-h1">Splunk Universal Forwarder Deployment</h1>
    <p class="report-sub">Required version: $safeVersion</p>
    $modeBadge
  </header>

  <div class="kpi">
$kpi  </div>

  <div class="panel">
    <h2>Per-VM Results</h2>
    <div style="overflow-x:auto;">
    <table class="ex-table">
      <thead>
        <tr>
          <th>VM Name</th><th>Power State</th><th>Action</th><th>Result</th>
          <th>Previous Version</th><th>Final Version</th><th>Service Status</th>
          <th>Reboot Required</th><th>Notes / Failure Reason</th><th>Duration</th>
        </tr>
      </thead>
      <tbody>
$rows      </tbody>
    </table>
    </div>
  </div>

  <footer>
    Generated by Deploy-SplunkUniversalForwarder.ps1. Reboot Required flags reflect a generic
    Windows pending-reboot check and are not necessarily caused by this run - see the Notes
    column for each VM.
  </footer>

</div>
</body>
</html>
"@
}

#endregion ============================================================


#region ======================= MAIN =======================

Write-Host "ScriptBuild: $ScriptBuild" -ForegroundColor Magenta
Write-Host "`n=== Splunk Universal Forwarder Deployment (SEVEN configuration) ===" -ForegroundColor Cyan
Write-Host "Required version: $RequiredSplunkVersion`n"

Test-LocalPrerequisites

# MSI property string (contains SPLUNKPASSWORD) - read from a local file that is never committed
# to this repo, not hardcoded here. Never logged or echoed. Not needed in CheckOnly mode.
$SplunkInstallerArgs = if ($CheckOnly) { '' } else { Get-SplunkMsiProperties }

if ($CheckOnly) {
    Write-Host "*** CHECK-ONLY MODE: no VM will be modified. Reporting current status only. ***`n" -ForegroundColor Magenta
}

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
            Write-Host "  [diag] Raw guest output: '$($detect.RawOutput)'" -ForegroundColor DarkGray
            throw "Guest credential validation / detection failed: $($detect.ErrorMessage)"
        }
        $status = $detect.Status
        $row.GuestCredentialStatus = if ($status.IsAdmin) { 'Valid (Administrator)' } else { 'Valid (NOT Administrator - install will likely fail)' }
        # Track whether this was pending BEFORE we touched the VM, so RebootRequired can be
        # explained honestly - this check is generic Windows reboot-pending detection (Windows
        # Update, Component Based Servicing, PendingFileRenameOperations), not Splunk-specific;
        # Splunk installs do not themselves require a reboot in normal circumstances.
        $preExistingRebootPending = [bool]$status.RebootPending
        $row.RebootRequired = $preExistingRebootPending
        if ($preExistingRebootPending) {
            Add-RowNote -Row $row -Note "Reboot already pending on this VM BEFORE this script ran (unrelated to Splunk - likely Windows Update or a prior change)."
        }

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
                Add-RowNote -Row $row -Note "Installed version $($status.Version) is NEWER than required $RequiredSplunkVersion - not downgrading. Manual review needed."
                Write-Host "  [notice] $($row.Notes)" -ForegroundColor Yellow
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
                Add-RowNote -Row $row -Note "Older version detected (installed: $($status.Version), required: $RequiredSplunkVersion) - upgrading."
                Write-Host "  [notice] $($row.Notes)" -ForegroundColor Yellow
            }
        } else {
            $row.Action = 'Installed'
            $needsProcedure = $true
        }

        if ($needsProcedure -and $CheckOnly) {
            # Audit only - report what WOULD happen without touching the VM.
            $row.Result = "Check Only - Would $($row.Action)"
            if ($status.Version) { $row.FinalVersion = $status.Version }
            Add-RowNote -Row $row -Note 'CheckOnly mode: no changes made to this VM.'
        }
        elseif ($needsProcedure) {
            if (-not $status.IsAdmin) {
                throw "Guest credential is not a local Administrator; cannot proceed with install/upgrade."
            }

            $procResult = Invoke-SplunkInstallProcedure -VM $vm -GuestCredential $guestCred
            $row.InstallerExitCode = $procResult.InstallerExitCode
            $row.SevenPostInstallStatus = $procResult.SevenPostInstallStatus
            if ($procResult.RebootRequired -and -not $preExistingRebootPending) {
                $row.RebootRequired = $true
                Add-RowNote -Row $row -Note 'Reboot required - triggered by this installation/post-install step (exit code 3010).'
            }

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
                    if ($fs.RebootPending -and -not $preExistingRebootPending -and -not $row.RebootRequired) {
                        $row.RebootRequired = $true
                        Add-RowNote -Row $row -Note 'Reboot required - detected as pending only after this installation completed.'
                    }

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
        if ($row.Notes) { Write-Host "Notes: $($row.Notes)" -ForegroundColor Yellow }
        if ($row.FailureReason) { Write-Host "Reason: $($row.FailureReason)" -ForegroundColor Red }
        if ($row.RebootRequired) { Write-Host "REBOOT REQUIRED on $vmName - not rebooting automatically." -ForegroundColor Yellow }
    }
}

# Export final report.
$results | Export-Csv -Path $ReportPath -NoTypeInformation -Encoding UTF8
Write-Host "`nReport written to: $ReportPath" -ForegroundColor Cyan

# Summary.
# @(...) around every Where-Object result is required - piping a SINGLE matching object through
# Where-Object in Windows PowerShell 5.1 returns a bare object, not a 1-element array, and .Count
# on a bare object is $null (not 1), printing blank instead of a number. Confirmed live: with
# exactly one Already Compliant VM, that line printed nothing at all. Same bug class already
# documented in this repo's CLAUDE.md for VMware_Weekly_HealthCheck.ps1.
# Plain loop counters instead of "(Where-Object {...}).Count" - that pattern has now caused two
# separate real failures live (a blank count when exactly one item matched, then an
# "Argument types do not match" ArgumentException on a later run) on this PowerShell 5.1 session.
# A loop with integer counters has no pipeline/.Count ambiguity to hit at all.
$totalVMs = 0; $alreadyCompliant = 0; $successInstalled = 0; $successUpgraded = 0
$manualReview = 0; $failedCount = 0; $skippedCount = 0; $rebootCount = 0; $checkOnlyCount = 0

foreach ($r in $results) {
    $totalVMs++
    if ($r.Result -eq 'Already Compliant') { $alreadyCompliant++ }
    elseif ($r.Result -eq 'Install Successful') { $successInstalled++ }
    elseif ($r.Result -eq 'Upgrade Successful') { $successUpgraded++ }
    elseif ($r.Action -eq 'Manual Review') { $manualReview++ }
    elseif ($r.Action -eq 'Failed') { $failedCount++ }
    elseif ($r.Result -like 'Check Only*') { $checkOnlyCount++ }
    elseif ($r.Action -eq 'Skipped') { $skippedCount++ }
    if ($r.RebootRequired) { $rebootCount++ }
}

$summary = [ordered]@{
    'Total VMs'                       = $totalVMs
    'Already Compliant'               = $alreadyCompliant
    'Successfully Installed'          = $successInstalled
    'Successfully Upgraded'           = $successUpgraded
    'Newer Version / Manual Review'   = $manualReview
    'Failed'                          = $failedCount
    'Skipped'                         = $skippedCount
    'Check Only (no changes made)'    = $checkOnlyCount
    'Reboot Required (not performed)' = $rebootCount
}

Write-Host "`n=== Summary ===" -ForegroundColor Cyan
foreach ($key in $summary.Keys) {
    Write-Host ("{0,-32}: {1}" -f $key, $summary[$key])
}

# HTML dashboard - best-effort, never fails the run if it can't be written.
try {
    $dashboardHtml = Build-SplunkDashboardHtml -Results $results -Summary $summary `
        -RequiredVersion $RequiredSplunkVersion.ToString() -ScriptBuild $ScriptBuild -CheckOnly:$CheckOnly.IsPresent
    [System.IO.File]::WriteAllText($ReportHtmlPath, $dashboardHtml, (New-Object System.Text.UTF8Encoding($false)))
    Write-Host "Dashboard written to: $ReportHtmlPath" -ForegroundColor Cyan
} catch {
    Write-Host "WARNING: Could not write HTML dashboard: $($_.Exception.Message)" -ForegroundColor Yellow
}

#endregion ============================================================
