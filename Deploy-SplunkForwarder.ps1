#Requires -Version 5.1
<#
.SYNOPSIS
    Deploys the Splunk Universal Forwarder to a list of Windows VMs via VMware Tools
    (Invoke-VMScript), without WinRM.

.DESCRIPTION
    Run from a laptop that already has an active PowerCLI session
    (Connect-VIServer done beforehand). For each VM in -InputFile:

      1. Validates: VM exists in vCenter, name/IP match the input list, VM is
         powered on, VMware Tools is running, guest credentials work, and Splunk UF
         is not already installed.
      2. Copies the local Splunk UF installer into the guest with Copy-VMGuestFile.
      3. Runs the installer silently (msiexec /quiet) via Invoke-VMScript.
      4. Verifies install: SplunkForwarder service exists/running, installed version.
      5. Does NOT reboot the VM - Splunk UF's MSI does not require a reboot.

    Any validation or step failure skips that VM only; the run always continues to
    the next VM. Nothing here configures a deployment server, indexer, outputs.conf,
    or forwarding - that is intentionally left to you (see $InstallArguments below).

.NOTES
    Author:  (fill in)
    Requires: VMware PowerCLI module, an already-connected vCenter session,
              VMware Tools running in every target guest.

.EXAMPLE
    Connect-VIServer vcenter.corp.local
    $GuestCred = Get-Credential -Message 'Local admin creds valid on all target guests'
    .\Deploy-SplunkForwarder.ps1 -InputFile .\targets.csv -InstallerPath 'C:\Installers\splunkforwarder-9.2.1-x64-release.msi' -GuestCredential $GuestCred
#>

[CmdletBinding()]
param(
    # ===== REQUIRED - CHANGE THESE FOR YOUR RUN =====================================

    # CSV or TXT file with the target VM list. Must have headers: VMName,IPAddress
    # (a plain TXT file works fine as long as it is comma-delimited with that header row).
    [Parameter(Mandatory)]
    [string]$InputFile,

    # Full local path to the Splunk Universal Forwarder installer (.msi) you already downloaded.
    [Parameter(Mandatory)]
    [string]$InstallerPath,

    # Credential valid inside every target guest (local admin or domain account with
    # local admin rights on the guest). Prompted interactively - never hard-code passwords.
    [Parameter(Mandatory)]
    [System.Management.Automation.PSCredential]$GuestCredential,

    # ===== OPTIONAL - defaults are reasonable, review before a large run ============

    # Silent-install arguments passed to msiexec. AGENTPASSWORD/RECEIVING_INDEXER/
    # DEPLOYMENT_SERVER etc. are deliberately NOT set here - add them yourself as
    # "PROPERTY=value" entries if your environment requires them.
    [string]$InstallArguments = 'AGREETOLICENSE=Yes /quiet',

    # Where the installer is staged inside the guest before running it.
    [string]$RemoteStagingPath = 'C:\Windows\Temp\SplunkUF\splunkforwarder.msi',

    # Splunk UF's own install folder, used for the post-install version/service check.
    [string]$SplunkInstallDir = 'C:\Program Files\SplunkUniversalForwarder',

    # Output folder for the CSV report and log file.
    [string]$OutputPath = (Join-Path $PSScriptRoot "SplunkUF_Deployment_Reports"),

    # Seconds to wait for the msiexec install to finish inside the guest before giving up.
    [int]$InstallTimeoutSeconds = 600
)

$ErrorActionPreference = 'Stop'
$RunStart   = Get-Date
$RunStamp   = $RunStart.ToString('yyyyMMdd_HHmmss')
if (-not (Test-Path $OutputPath)) { New-Item -ItemType Directory -Path $OutputPath -Force | Out-Null }

$CsvReportPath = Join-Path $OutputPath "SplunkUF_Deployment_$RunStamp.csv"
$LogPath       = Join-Path $OutputPath "SplunkUF_Deployment_$RunStamp.log"

# ============================================================================
# 0. LOGGING / RESULT TRACKING
# ============================================================================
$Global:Results = [System.Collections.Generic.List[object]]::new()

function Write-Log {
    param(
        [Parameter(Mandatory)][string]$Message,
        [ValidateSet('INFO','WARN','ERROR')][string]$Level = 'INFO'
    )
    $line = "[{0}] [{1}] {2}" -f (Get-Date).ToString('s'), $Level, $Message
    Add-Content -Path $LogPath -Value $line
    switch ($Level) {
        'WARN'  { Write-Warning $Message }
        'ERROR' { Write-Warning $Message }
        default { Write-Host $Message }
    }
}

# Records one VM's final outcome. Status is one of:
# AlreadyInstalled | Success | Failed | Skipped
function Add-Result {
    param(
        [string]$VMName,
        [string]$IPAddress,
        [string]$Status,
        [string]$SplunkVersion = '',
        [string]$ServiceStatus = '',
        [string]$Reason = ''
    )
    $Global:Results.Add([pscustomobject]@{
        Timestamp     = (Get-Date).ToString('s')
        VMName        = $VMName
        IPAddress     = $IPAddress
        Status        = $Status
        SplunkVersion = $SplunkVersion
        ServiceStatus = $ServiceStatus
        Reason        = $Reason
    })
}

# ============================================================================
# 1. PRE-FLIGHT
# ============================================================================
if (-not (Get-Module -ListAvailable -Name VMware.VimAutomation.Core)) {
    throw "VMware PowerCLI is not installed. Install-Module VMware.PowerCLI first."
}
Import-Module VMware.VimAutomation.Core -ErrorAction Stop

if (-not $global:DefaultVIServers -or $global:DefaultVIServers.Count -eq 0) {
    throw "No active vCenter connection found. Run Connect-VIServer before this script."
}

if (-not (Test-Path $InputFile)) {
    throw "Input file not found: $InputFile"
}
if (-not (Test-Path $InstallerPath)) {
    throw "Splunk installer not found: $InstallerPath"
}

$Targets = Import-Csv -Path $InputFile
if (-not $Targets -or -not ($Targets | Get-Member -Name VMName) -or -not ($Targets | Get-Member -Name IPAddress)) {
    throw "Input file must be a CSV/TXT with headers 'VMName,IPAddress'."
}

Write-Log "Starting Splunk UF deployment run. Targets: $($Targets.Count). Installer: $InstallerPath"

# ============================================================================
# 2. PER-VM DEPLOYMENT
# ============================================================================
foreach ($target in $Targets) {

    $vmName = $target.VMName.Trim()
    $expectedIP = $target.IPAddress.Trim()

    Write-Log "----- Processing '$vmName' ($expectedIP) -----"

    try {
        # --- 2.1 VM exists in vCenter -------------------------------------------------
        $vm = Get-VM -Name $vmName -ErrorAction SilentlyContinue
        if (-not $vm) {
            Write-Log "VM '$vmName' not found in vCenter." 'WARN'
            Add-Result -VMName $vmName -IPAddress $expectedIP -Status 'Skipped' -Reason 'VM not found in vCenter'
            continue
        }
        if (@($vm).Count -gt 1) {
            Write-Log "Multiple VMs named '$vmName' found; ambiguous, skipping." 'WARN'
            Add-Result -VMName $vmName -IPAddress $expectedIP -Status 'Skipped' -Reason 'Multiple VMs with this name in vCenter'
            continue
        }

        # --- 2.2 Powered on -------------------------------------------------------------
        if ($vm.PowerState -ne 'PoweredOn') {
            Write-Log "VM '$vmName' is not powered on (state: $($vm.PowerState))." 'WARN'
            Add-Result -VMName $vmName -IPAddress $expectedIP -Status 'Skipped' -Reason "VM not powered on ($($vm.PowerState))"
            continue
        }

        # --- 2.3 VMware Tools installed and running -------------------------------------
        $vmView = Get-View -VIObject $vm -Property Guest
        $toolsStatus = $vmView.Guest.ToolsStatus
        $toolsRunning = $vmView.Guest.ToolsRunningStatus
        if ($toolsRunning -ne 'guestToolsRunning') {
            Write-Log "VMware Tools not running on '$vmName' (status: $toolsStatus / $toolsRunning)." 'WARN'
            Add-Result -VMName $vmName -IPAddress $expectedIP -Status 'Skipped' -Reason "VMware Tools not running ($toolsStatus)"
            continue
        }

        # --- 2.4 Name/IP match the input list --------------------------------------------
        $guestIPs = $vmView.Guest.Net | ForEach-Object { $_.IpAddress } | Where-Object { $_ }
        if ($expectedIP -notin $guestIPs) {
            Write-Log "IP mismatch for '$vmName'. Expected $expectedIP, VM reports: $($guestIPs -join ', ')" 'WARN'
            Add-Result -VMName $vmName -IPAddress $expectedIP -Status 'Skipped' -Reason "IP mismatch (VM reports: $($guestIPs -join ', '))"
            continue
        }

        # --- 2.5 Guest credentials valid (also confirms guest OS is responsive) -----------
        try {
            $osCheck = Invoke-VMScript -VM $vm -GuestCredential $GuestCredential `
                -ScriptType Powershell -ScriptText '$env:COMPUTERNAME' -ErrorAction Stop
        } catch {
            Write-Log "Guest credential/connectivity check failed for '$vmName': $($_.Exception.Message)" 'WARN'
            Add-Result -VMName $vmName -IPAddress $expectedIP -Status 'Skipped' -Reason "Guest credential/connectivity check failed: $($_.Exception.Message)"
            continue
        }

        # --- 2.6 Already installed? --------------------------------------------------------
        $checkInstalledScript = @"
if (Test-Path '$SplunkInstallDir\bin\splunk.exe') {
    `$svc = Get-Service -Name 'SplunkForwarder' -ErrorAction SilentlyContinue
    `$ver = (Get-Item '$SplunkInstallDir\bin\splunk.exe').VersionInfo.ProductVersion
    "INSTALLED|`$(`$svc.Status)|`$ver"
} else {
    "NOTINSTALLED"
}
"@
        $installedCheck = Invoke-VMScript -VM $vm -GuestCredential $GuestCredential `
            -ScriptType Powershell -ScriptText $checkInstalledScript -ErrorAction Stop
        $installedOut = $installedCheck.ScriptOutput.Trim()

        if ($installedOut -like 'INSTALLED|*') {
            $parts = $installedOut.Split('|')
            Write-Log "Splunk UF already installed on '$vmName' (version $($parts[2]), service: $($parts[1]))."
            Add-Result -VMName $vmName -IPAddress $expectedIP -Status 'AlreadyInstalled' `
                -SplunkVersion $parts[2] -ServiceStatus $parts[1] -Reason 'Already installed'
            continue
        }

        # --- 2.7 Copy installer to guest ---------------------------------------------------
        $remoteDir = Split-Path $RemoteStagingPath -Parent
        $mkdirScript = "New-Item -ItemType Directory -Path '$remoteDir' -Force | Out-Null"
        Invoke-VMScript -VM $vm -GuestCredential $GuestCredential -ScriptType Powershell -ScriptText $mkdirScript -ErrorAction Stop | Out-Null

        Write-Log "Copying installer to '$vmName':$RemoteStagingPath ..."
        Copy-VMGuestFile -Source $InstallerPath -Destination $RemoteStagingPath -VM $vm `
            -LocalToGuest -GuestCredential $GuestCredential -Force -ErrorAction Stop

        # --- 2.8 Silent install --------------------------------------------------------------
        $installScript = @"
`$p = Start-Process -FilePath 'msiexec.exe' -ArgumentList '/i `"$RemoteStagingPath`" $InstallArguments /norestart /l*v `"$remoteDir\install.log`"' -Wait -PassThru
"ExitCode=`$(`$p.ExitCode)"
"@
        Write-Log "Installing Splunk UF on '$vmName' (timeout ${InstallTimeoutSeconds}s)..."
        $installResult = Invoke-VMScript -VM $vm -GuestCredential $GuestCredential `
            -ScriptType Powershell -ScriptText $installScript -ToolsWaitSecs $InstallTimeoutSeconds -ErrorAction Stop
        $installOut = $installResult.ScriptOutput.Trim()

        if ($installOut -notmatch 'ExitCode=0') {
            Write-Log "Install on '$vmName' returned non-zero: $installOut" 'WARN'
            Add-Result -VMName $vmName -IPAddress $expectedIP -Status 'Failed' -Reason "msiexec $installOut"
            continue
        }

        # --- 2.9 Post-install verification -----------------------------------------------
        $verifyScript = @"
`$svc = Get-Service -Name 'SplunkForwarder' -ErrorAction SilentlyContinue
if (-not `$svc) { "NOSERVICE"; exit }
if (`$svc.Status -ne 'Running') { Start-Service -Name 'SplunkForwarder' -ErrorAction SilentlyContinue }
Start-Sleep -Seconds 5
`$svc.Refresh()
`$ver = (Get-Item '$SplunkInstallDir\bin\splunk.exe' -ErrorAction SilentlyContinue).VersionInfo.ProductVersion
"OK|`$(`$svc.Status)|`$ver"
"@
        $verify = Invoke-VMScript -VM $vm -GuestCredential $GuestCredential `
            -ScriptType Powershell -ScriptText $verifyScript -ErrorAction Stop
        $verifyOut = $verify.ScriptOutput.Trim()

        if ($verifyOut -eq 'NOSERVICE' -or -not $verifyOut) {
            Write-Log "Install ran but SplunkForwarder service not found on '$vmName'." 'WARN'
            Add-Result -VMName $vmName -IPAddress $expectedIP -Status 'Failed' -Reason 'Installer completed but SplunkForwarder service missing'
            continue
        }

        $vparts = $verifyOut.Split('|')
        $svcStatus = $vparts[1]
        $version   = $vparts[2]

        if ($svcStatus -ne 'Running') {
            Write-Log "SplunkForwarder installed on '$vmName' but service is '$svcStatus', not Running." 'WARN'
            Add-Result -VMName $vmName -IPAddress $expectedIP -Status 'Failed' -SplunkVersion $version -ServiceStatus $svcStatus `
                -Reason "Service present but not running ($svcStatus)"
            continue
        }

        Write-Log "Splunk UF $version successfully installed and running on '$vmName'."
        Add-Result -VMName $vmName -IPAddress $expectedIP -Status 'Success' -SplunkVersion $version -ServiceStatus $svcStatus -Reason ''
    }
    catch {
        Write-Log "Unhandled error processing '$vmName': $($_.Exception.Message)" 'ERROR'
        Add-Result -VMName $vmName -IPAddress $expectedIP -Status 'Failed' -Reason $_.Exception.Message
    }
}

# ============================================================================
# 3. REPORT / SUMMARY
# ============================================================================
$Global:Results | Export-Csv -Path $CsvReportPath -NoTypeInformation -Encoding UTF8

$total    = $Global:Results.Count
$success  = ($Global:Results | Where-Object Status -eq 'Success').Count
$already  = ($Global:Results | Where-Object Status -eq 'AlreadyInstalled').Count
$skipped  = ($Global:Results | Where-Object Status -eq 'Skipped').Count
$failed   = ($Global:Results | Where-Object Status -eq 'Failed').Count

Write-Host ""
Write-Host "================ Splunk UF Deployment Summary ================"
Write-Host "Total VMs processed : $total"
Write-Host "Successfully Installed : $success"
Write-Host "Already Installed      : $already"
Write-Host "Skipped                : $skipped"
Write-Host "Failed                 : $failed"
Write-Host "================================================================"
Write-Host "CSV report: $CsvReportPath"
Write-Host "Log file  : $LogPath"

Write-Log "Run complete. Total=$total Success=$success AlreadyInstalled=$already Skipped=$skipped Failed=$failed"
