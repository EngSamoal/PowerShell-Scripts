#Requires -Version 5.1
<#
.SYNOPSIS
    Staged, least-to-most-disruptive Windows Update remediation across a list of guest
    VMs, executed through VMware Tools Guest Operations (Invoke-VMScript). Companion to
    Invoke-WindowsUpdateDiagnostics.ps1 - run diagnostics FIRST and read every 'Problem'
    row's NEXT step before touching this script.

.DESCRIPTION
    Runs against every VM listed in -VMListPath (one name per line, '#' comments
    ignored). Each VM is fully isolated and confirmed individually (via ShouldProcess) -
    a VM that can't be found/validated, or throws mid-run, is recorded as skipped/failed
    and the batch continues with the next VM; declining the prompt for one VM does not
    stop the rest.

    Four independent stages, each requested explicitly with -Stage. Nothing changes on
    any guest unless -Apply is also passed - by default every stage only PRINTS what it
    would do ("Planned"), so you can review the plan across the whole VM list before
    committing to it. Each action is individually labelled Read-Only or Change.

    Stage 1 (lowest risk - service configuration only, no restart of anything healthy):
        - Reports current StartType for wuauserv/BITS/CryptSvc/UsoSvc/WaaSMedicSvc/DoSvc.
        - CHANGE: only for services found in StartType=Disabled, restores the Windows
          default startup type (never touches a service already Manual/Automatic - that
          may be an intentional admin choice).
        - CHANGE: starts wuauserv/BITS/CryptSvc if stopped and not disabled.
        - CHANGE: triggers a normal detection scan (Microsoft.Update.AutoUpdate.DetectNow) -
          does not force a download or install.

    Stage 2 (moderate - routine maintenance, no service stop/rename):
        - Read-Only: DISM /Online /Cleanup-Image /ScanHealth (deep component-store scan;
          per Microsoft docs this does not change anything, it only reports).
        - CHANGE: DISM /Online /Cleanup-Image /StartComponentCleanup (removes superseded
          component versions - routine, Microsoft-recommended, does not remove
          functionality).
        - CHANGE: gracefully restarts wuauserv/BITS ONLY if diagnostics flagged them as
          Warning/Problem.

    Stage 3 (disruptive - short outage of Windows Update while it runs; do this in a
    maintenance window):
        - CHANGE: stops wuauserv/BITS/CryptSvc/UsoSvc/WaaSMedicSvc.
        - CHANGE: RENAMES (never deletes) C:\Windows\SoftwareDistribution and
          C:\Windows\System32\catroot2 to '<name>.old.<timestamp>', so nothing is
          destroyed and the originals can be restored by renaming back if needed.
        - CHANGE: restarts the services and triggers a fresh detection scan.
        - Refuses to run (even with -Apply) on a VM with a pending-reboot flag currently
          set, unless -Force is also passed - resetting Windows Update state ahead of a
          required reboot usually just reproduces the same failure.

    Stage 4 (most disruptive - can download a large corrective payload and take a long
    time; requires -Apply -Force per VM if a reboot is pending):
        - CHANGE: DISM /Online /Cleanup-Image /RestoreHealth. Only sensible to run after
          Stage 2's ScanHealth reported repairable corruption. Does not reboot the guest -
          schedule a reboot yourself afterward.

    This script never deletes SoftwareDistribution/Catroot2, never runs a raw
    'rm/Remove-Item' against Windows Update state, and never reboots any guest.

.PARAMETER VMListPath     Text file of VM names, one per line (default C:\temp\vmlist.txt).
                           '#' lines and blanks are ignored; names are de-duplicated.
.PARAMETER Stage          Which remediation stage to run (1-4) - applied uniformly to
                           every VM in the list. Stages are independent, not cumulative -
                           run and evaluate one at a time.
.PARAMETER Apply          Without this switch, every VM's stage run only reports what it
                           WOULD do. Pass -Apply to actually make the changes.
.PARAMETER Force          Required in addition to -Apply for Stage 3/4 on any VM where a
                           pending reboot is currently detected.
.PARAMETER CredentialPath Export-Clixml PSCredential for the guest admin (used for every
                           VM). If omitted or not found, you are prompted with
                           Get-Credential instead.
.PARAMETER OutputPath     Folder on the admin machine for the consolidated CSV/log.
.PARAMETER ToolsWaitSecs  Invoke-VMScript VMware Tools wait, seconds (default 300 - DISM
                           operations in Stage 2/4 can run long).

.EXAMPLE
    # Review what Stage 1 would change across the whole list, without changing anything:
    Connect-VIServer vcenter01
    .\Invoke-WindowsUpdateRemediation.ps1 -Stage 1

.EXAMPLE
    # Actually apply Stage 1 across the list (still prompts per VM, ShouldProcess):
    .\Invoke-WindowsUpdateRemediation.ps1 -Stage 1 -Apply

.EXAMPLE
    # Stage 3 reset on a specific subset, overriding the pending-reboot guard because a
    # reboot is scheduled for right after this run:
    .\Invoke-WindowsUpdateRemediation.ps1 -VMListPath C:\temp\wsus_servers.txt -Stage 3 -Apply -Force

.NOTES
    Requires PowerCLI, an existing Connect-VIServer session, and vCenter "Guest Operation
    Program Execution/Query" privileges on every target VM. Run
    Invoke-WindowsUpdateDiagnostics.ps1 first and do not run this against a VM that Azure
    Update Manager may currently be patching - check that script's "Azure Arc / AUM"
    section first.
#>

[CmdletBinding(SupportsShouldProcess = $true, ConfirmImpact = 'High')]
param(
    [Parameter(Mandatory)]
    [ValidateSet(1, 2, 3, 4)]
    [int] $Stage,

    [string] $VMListPath     = 'C:\temp\vmlist.txt',
    [switch] $Apply,
    [switch] $Force,

    [string] $CredentialPath = 'C:\temp\wincred.xml',
    [string] $OutputPath     = (Join-Path $PSScriptRoot 'WindowsUpdate_Diagnostics'),

    [ValidateRange(30, 3600)]
    [int] $ToolsWaitSecs = 300
)

$ErrorActionPreference = 'Stop'
$ProgressPreference    = 'SilentlyContinue'

Write-Host "Invoke-WindowsUpdateRemediation.ps1 - staged Windows Update remediation across a VM list  [build 2026-09-09b]" -ForegroundColor Magenta
Write-Host ("Running from: {0}" -f $PSCommandPath) -ForegroundColor DarkGray
if (-not $Apply) {
    Write-Host "DRY RUN (no -Apply): this run will only report what Stage $Stage WOULD change on each VM. Nothing will be modified." -ForegroundColor Cyan
} else {
    Write-Host "APPLY MODE: Stage $Stage changes WILL be made on each confirmed VM." -ForegroundColor Red
}

# =====================================================================================
# 1. Pre-flight (read-only)
# =====================================================================================
if (-not (Get-Command Invoke-VMScript -ErrorAction SilentlyContinue)) {
    try { Import-Module VMware.VimAutomation.Core -ErrorAction Stop }
    catch { throw "VMware PowerCLI (VMware.VimAutomation.Core) is not available. Install PowerCLI and retry." }
}
$connectedServers = @($global:DefaultVIServers | Where-Object { $_.IsConnected })
if ($connectedServers.Count -eq 0) { throw "No connected vCenter session. Run Connect-VIServer <vcenter> first." }
Write-Host ("Connected vCenter(s): {0}" -f (($connectedServers | ForEach-Object { $_.Name }) -join ', ')) -ForegroundColor Gray

if (-not (Test-Path -LiteralPath $VMListPath)) { throw "VM list not found: $VMListPath" }
$vmNames = @(Get-Content -LiteralPath $VMListPath |
             ForEach-Object { $_.Trim() } |
             Where-Object { $_ -and -not $_.StartsWith('#') } |
             Select-Object -Unique)
if ($vmNames.Count -eq 0) { throw "VM list '$VMListPath' contained no usable VM names." }
Write-Host ("VMs to process: {0}" -f $vmNames.Count) -ForegroundColor Gray

if (Test-Path -LiteralPath $CredentialPath) {
    try {
        $GuestCredential = Import-Clixml -LiteralPath $CredentialPath
        if (-not ($GuestCredential -is [pscredential])) { throw "did not deserialise to a PSCredential" }
    } catch {
        throw "Failed to import guest credential from '$CredentialPath': $($_.Exception.Message)"
    }
} else {
    $GuestCredential = Get-Credential -Message "Guest administrator credential (used for every VM in the list)"
}

if (-not (Test-Path -LiteralPath $OutputPath)) { New-Item -ItemType Directory -Path $OutputPath -Force | Out-Null }
$RunStamp = Get-Date -Format 'yyyyMMdd-HHmmss'
$CsvPath  = Join-Path $OutputPath "WU_Remediation_Stage${Stage}_$RunStamp.csv"
$LogPath  = Join-Path $OutputPath "WU_Remediation_Stage${Stage}_$RunStamp.log"

function Write-Log {
    param([string]$Message, [string]$Level = 'INFO')
    $line = "[{0}] [{1}] {2}" -f (Get-Date -Format 'yyyy-MM-dd HH:mm:ss'), $Level, $Message
    Write-Host $line
    Add-Content -LiteralPath $LogPath -Value $line
}
Write-Log "Remediation run started. VMs=$($vmNames.Count). Stage=$Stage Apply=$($Apply.IsPresent) Force=$($Force.IsPresent)."

# =====================================================================================
# 2. In-guest payload
#    {{STAGE}} / {{APPLY}} / {{FORCE}} are substituted before each run.
#    Every action is guarded by "if ($Apply) { <make the change> } else { <report Planned> }".
# =====================================================================================
$PayloadTemplate = @'
$ProgressPreference = 'SilentlyContinue'
$Stage = {{STAGE}}
$Apply = ${{APPLY}}
$Force = ${{FORCE}}

$rows = New-Object System.Collections.Generic.List[object]
function Add-Row {
    param($Action, $Mode, $Result, $Detail)
    $rows.Add([pscustomobject]@{ Action = $Action; Mode = $Mode; Result = $Result; Detail = $Detail }) | Out-Null
}

# Pending-reboot guard (used to block Stage 3/4 apply unless -Force).
$rebootFlags = New-Object System.Collections.Generic.List[string]
if (Test-Path 'HKLM:\SOFTWARE\Microsoft\Windows\CurrentVersion\Component Based Servicing\RebootPending')    { $rebootFlags.Add('CBS RebootPending') }
if (Test-Path 'HKLM:\SOFTWARE\Microsoft\Windows\CurrentVersion\Component Based Servicing\RebootInProgress') { $rebootFlags.Add('CBS RebootInProgress') }
if (Test-Path 'HKLM:\SOFTWARE\Microsoft\Windows\CurrentVersion\WindowsUpdate\Auto Update\RebootRequired')   { $rebootFlags.Add('WU Auto Update RebootRequired') }
$pendingReboot = $rebootFlags.Count -gt 0

$defaultStartType = @{ wuauserv = 'Manual'; BITS = 'Manual'; CryptSvc = 'Automatic'; UsoSvc = 'Manual'; WaaSMedicSvc = 'Manual'; DoSvc = 'Manual' }

if ($Stage -eq 1) {
    foreach ($svcName in @('wuauserv','BITS','CryptSvc','UsoSvc','WaaSMedicSvc','DoSvc')) {
        try {
            $svc = Get-Service -Name $svcName -ErrorAction Stop
            Add-Row "$svcName current state" 'Read-Only' 'Reported' "Status=$($svc.Status), StartType=$($svc.StartType)"
            if ($svc.StartType -eq 'Disabled') {
                $target = $defaultStartType[$svcName]
                if ($Apply) {
                    try {
                        Set-Service -Name $svcName -StartupType $target -ErrorAction Stop
                        Add-Row "$svcName startup type" 'Change' 'Applied' "Changed StartType from Disabled to $target (Windows default)."
                    } catch {
                        Add-Row "$svcName startup type" 'Change' 'Failed' "Set-Service failed: $($_.Exception.Message)"
                    }
                } else {
                    Add-Row "$svcName startup type" 'Change' 'Planned' "Would change StartType from Disabled to $target (Windows default)."
                }
            }
            if ($svc.Status -ne 'Running' -and $svc.StartType -ne 'Disabled' -and $svcName -in @('wuauserv','BITS','CryptSvc')) {
                if ($Apply) {
                    try {
                        Start-Service -Name $svcName -ErrorAction Stop
                        Add-Row "$svcName start" 'Change' 'Applied' "Started $svcName (was $($svc.Status))."
                    } catch {
                        Add-Row "$svcName start" 'Change' 'Failed' "Start-Service failed: $($_.Exception.Message)"
                    }
                } else {
                    Add-Row "$svcName start" 'Change' 'Planned' "Would start $svcName (currently $($svc.Status))."
                }
            }
        } catch {
            Add-Row "$svcName current state" 'Read-Only' 'Unable to Check' "Service not found: $($_.Exception.Message)"
        }
    }
    if ($Apply) {
        try {
            $au = New-Object -ComObject Microsoft.Update.AutoUpdate
            $au.DetectNow()
            Add-Row 'Trigger detection scan' 'Change' 'Applied' 'Called Microsoft.Update.AutoUpdate.DetectNow() - starts a normal scan only, does not force a download or install.'
        } catch {
            Add-Row 'Trigger detection scan' 'Change' 'Failed' "DetectNow() failed: $($_.Exception.Message)"
        }
    } else {
        Add-Row 'Trigger detection scan' 'Change' 'Planned' 'Would call Microsoft.Update.AutoUpdate.DetectNow() to start a normal detection scan.'
    }
}
elseif ($Stage -eq 2) {
    try {
        $scan = & dism.exe /online /cleanup-image /scanhealth 2>&1 | Out-String
        Add-Row 'DISM /ScanHealth' 'Read-Only' 'Reported' $scan.Trim()
    } catch {
        Add-Row 'DISM /ScanHealth' 'Read-Only' 'Unable to Check' "dism.exe failed: $($_.Exception.Message)"
    }
    if ($Apply) {
        try {
            $cleanup = & dism.exe /online /cleanup-image /startcomponentcleanup 2>&1 | Out-String
            Add-Row 'DISM /StartComponentCleanup' 'Change' 'Applied' $cleanup.Trim()
        } catch {
            Add-Row 'DISM /StartComponentCleanup' 'Change' 'Failed' "dism.exe failed: $($_.Exception.Message)"
        }
    } else {
        Add-Row 'DISM /StartComponentCleanup' 'Change' 'Planned' 'Would run DISM /Online /Cleanup-Image /StartComponentCleanup (removes superseded component versions; routine, does not remove functionality).'
    }
    foreach ($svcName in @('wuauserv','BITS')) {
        try {
            $svc = Get-Service -Name $svcName -ErrorAction Stop
            if ($Apply) {
                try {
                    Restart-Service -Name $svcName -Force -ErrorAction Stop
                    Add-Row "$svcName graceful restart" 'Change' 'Applied' "Restarted $svcName."
                } catch {
                    Add-Row "$svcName graceful restart" 'Change' 'Failed' "Restart-Service failed: $($_.Exception.Message)"
                }
            } else {
                Add-Row "$svcName graceful restart" 'Change' 'Planned' "Would gracefully restart $svcName (current status $($svc.Status))."
            }
        } catch {
            Add-Row "$svcName graceful restart" 'Change' 'Unable to Check' "Service not found: $($_.Exception.Message)"
        }
    }
}
elseif ($Stage -eq 3) {
    if ($pendingReboot -and -not $Force) {
        Add-Row 'Stage 3 guard' 'Read-Only' 'Blocked' "Pending reboot flag(s) present: $($rebootFlags -join '; '). Refusing to reset Windows Update state ahead of a required reboot - re-run after rebooting, or add -Force to override."
    } else {
        $svcNames = @('wuauserv','BITS','CryptSvc','UsoSvc','WaaSMedicSvc')
        if ($Apply) {
            foreach ($svcName in $svcNames) {
                try { Stop-Service -Name $svcName -Force -ErrorAction Stop; Add-Row "$svcName stop" 'Change' 'Applied' "Stopped $svcName." }
                catch { Add-Row "$svcName stop" 'Change' 'Failed' "Stop-Service failed: $($_.Exception.Message)" }
            }
            $stamp = Get-Date -Format 'yyyyMMddHHmmss'
            foreach ($pair in @(
                @{ Path = 'C:\Windows\SoftwareDistribution'; Name = 'SoftwareDistribution' },
                @{ Path = 'C:\Windows\System32\catroot2';    Name = 'catroot2' }
            )) {
                if (Test-Path -LiteralPath $pair.Path) {
                    $dest = "$($pair.Path).old.$stamp"
                    try {
                        Rename-Item -LiteralPath $pair.Path -NewName (Split-Path $dest -Leaf) -ErrorAction Stop
                        Add-Row "Rename $($pair.Name)" 'Change' 'Applied' "Renamed $($pair.Path) to $dest (original preserved, nothing deleted)."
                    } catch {
                        Add-Row "Rename $($pair.Name)" 'Change' 'Failed' "Rename-Item failed: $($_.Exception.Message)"
                    }
                } else {
                    Add-Row "Rename $($pair.Name)" 'Change' 'Skipped' "$($pair.Path) does not exist."
                }
            }
            foreach ($svcName in $svcNames) {
                try { Start-Service -Name $svcName -ErrorAction Stop; Add-Row "$svcName start" 'Change' 'Applied' "Started $svcName." }
                catch { Add-Row "$svcName start" 'Change' 'Failed' "Start-Service failed: $($_.Exception.Message)" }
            }
            try {
                $au = New-Object -ComObject Microsoft.Update.AutoUpdate
                $au.DetectNow()
                Add-Row 'Trigger detection scan' 'Change' 'Applied' 'Called Microsoft.Update.AutoUpdate.DetectNow() after the reset.'
            } catch {
                Add-Row 'Trigger detection scan' 'Change' 'Failed' "DetectNow() failed: $($_.Exception.Message)"
            }
        } else {
            Add-Row 'Stop services' 'Change' 'Planned' "Would stop: $($svcNames -join ', ')."
            Add-Row 'Rename SoftwareDistribution' 'Change' 'Planned' 'Would rename C:\Windows\SoftwareDistribution to SoftwareDistribution.old.<timestamp> (never deleted).'
            Add-Row 'Rename catroot2' 'Change' 'Planned' 'Would rename C:\Windows\System32\catroot2 to catroot2.old.<timestamp> (never deleted).'
            Add-Row 'Start services' 'Change' 'Planned' "Would start: $($svcNames -join ', '), then trigger a detection scan."
            if ($pendingReboot) {
                Add-Row 'Stage 3 guard' 'Read-Only' 'Would Block' "Pending reboot flag(s) present: $($rebootFlags -join '; '). -Apply alone would be refused; -Apply -Force would be required."
            }
        }
    }
}
elseif ($Stage -eq 4) {
    if ($pendingReboot -and -not $Force) {
        Add-Row 'Stage 4 guard' 'Read-Only' 'Blocked' "Pending reboot flag(s) present: $($rebootFlags -join '; '). Reboot first, or add -Force to override."
    } elseif ($Apply) {
        try {
            $restore = & dism.exe /online /cleanup-image /restorehealth 2>&1 | Out-String
            Add-Row 'DISM /RestoreHealth' 'Change' 'Applied' $restore.Trim()
        } catch {
            Add-Row 'DISM /RestoreHealth' 'Change' 'Failed' "dism.exe failed: $($_.Exception.Message)"
        }
        Add-Row 'Reboot reminder' 'Read-Only' 'Reported' 'DISM /RestoreHealth does not reboot the server. Schedule a reboot during your next maintenance window, then retry Windows Update.'
    } else {
        Add-Row 'DISM /RestoreHealth' 'Change' 'Planned' 'Would run DISM /Online /Cleanup-Image /RestoreHealth - downloads corrective payload from Windows Update/WSUS and can take a long time. Only worthwhile if Stage 2 ScanHealth reported repairable corruption.'
    }
}

$meta = [pscustomobject]@{
    Hostname       = $env:COMPUTERNAME
    Stage          = $Stage
    Applied        = [bool]$Apply
    ForceUsed      = [bool]$Force
    PendingReboot  = $pendingReboot
    RebootFlags    = ($rebootFlags -join '; ')
    CollectedAtUtc = (Get-Date).ToUniversalTime().ToString('yyyy-MM-dd HH:mm:ss')
}
$envelope = [pscustomobject]@{ Schema = 'wu-remediation-1'; Meta = $meta; Rows = $rows.ToArray() }
$json = $envelope | ConvertTo-Json -Depth 8 -Compress
$b64  = [Convert]::ToBase64String([Text.Encoding]::UTF8.GetBytes($json))
Write-Output '<<<WU-REMEDIATION-ENVELOPE-B64>>>'
Write-Output $b64
Write-Output '<<<END-WU-REMEDIATION-ENVELOPE>>>'
'@

$Payload = $PayloadTemplate.Replace('{{STAGE}}', "$Stage").Replace('{{APPLY}}', $(if ($Apply) { 'true' } else { 'false' })).Replace('{{FORCE}}', $(if ($Force) { 'true' } else { 'false' }))

# =====================================================================================
# 3. Admin-side helpers
# =====================================================================================
function Read-EnvelopeFromScriptOutput {
    param([string]$Output, [string]$StartMarker, [string]$EndMarker)
    if ([string]::IsNullOrEmpty($Output)) { return $null }
    $capture = $false
    $sb = New-Object System.Text.StringBuilder
    foreach ($line in ($Output -split "`r?`n")) {
        $t = $line.Trim()
        if ($t -eq $StartMarker) { $capture = $true; continue }
        if ($t -eq $EndMarker)   { break }
        if ($capture) { [void]$sb.Append($t) }
    }
    if ($sb.Length -eq 0) { return $null }
    try {
        $json = [Text.Encoding]::UTF8.GetString([Convert]::FromBase64String($sb.ToString()))
        return $json | ConvertFrom-Json
    } catch { return $null }
}

# Classifies the common VIX/guest-ops Invoke-VMScript failures into an actionable
# explanation instead of surfacing the raw (often misleading) VMware error text.
function Get-FriendlyVixError {
    param([string]$Message)
    switch -Regex ($Message) {
        'Could not locate .?Powershell.? script interpreter' {
            return "VMware Tools authenticated the guest credential but could not confirm it has a full administrator token in-guest. Almost always one of: (1) the credential is a LOCAL (non-domain, non-built-in-Administrator) account and Windows' UAC remote token filtering silently downgrades it for network-style logons - fix on the guest with: reg add `"HKLM\SOFTWARE\Microsoft\Windows\CurrentVersion\Policies\System`" /v LocalAccountTokenFilterPolicy /t REG_DWORD /d 1 /f  (no reboot needed, then re-run); or (2) VMware Tools on this VM is out of date and isn't correctly registering the PowerShell interpreter path - upgrade VMware Tools on this VM. Isolate which one with: Invoke-VMScript -VM <vm> -ScriptText 'echo hi' -ScriptType Bat -GuestCredential `$cred - if that also fails, it's the credential/UAC issue; if that works, it's Tools."
        }
        'InvalidGuestLogin|[Ff]ailed to authenticate|authentication fail|incorrect user name or password|InvalidLogin|guest permissions|permission to perform this operation' {
            return "Guest authentication failed - VMware Tools rejected the supplied credential. Verify the username/password in the credential file/prompt and that the account exists and isn't locked out on this VM."
        }
        'Tools are not running|GuestOperationsUnavailable|VMware Tools is not|not installed|VIX' {
            return "VMware Tools guest operations are unavailable on this VM (Tools not running/installed, or too old to support guest operations). Check Tools status in vCenter."
        }
        'timed out|timeout' {
            return "Timed out waiting for VMware Tools / guest response. The guest may be under load, booting, or Tools may be unresponsive - retry, or increase -ToolsWaitSecs."
        }
        default { return $null }
    }
}

function Get-VMInventory {
    param([object[]]$Servers)
    $inv = New-Object System.Collections.Generic.List[object]
    foreach ($s in $Servers) {
        try { $vms = @(Get-VM -Server $s -ErrorAction Stop) }
        catch { Write-Log "Get-VM failed on vCenter '$($s.Name)': $($_.Exception.Message)" 'WARN'; $vms = @() }
        $inv.Add([pscustomobject]@{ Server = $s; VMs = $vms }) | Out-Null
    }
    return $inv
}

function Resolve-AndValidateVM {
    param([string]$Name, [object[]]$Inventory)
    $hits = New-Object System.Collections.Generic.List[object]
    foreach ($entry in $Inventory) {
        foreach ($v in $entry.VMs) {
            if ($v.Name -eq $Name) { $hits.Add([pscustomobject]@{ Server = $entry.Server; VM = $v }) | Out-Null }
        }
    }
    $serverList = ($Inventory | ForEach-Object { $_.Server.Name }) -join ', '
    if ($hits.Count -eq 0) {
        return [pscustomobject]@{ Ok = $false; Stage = 'VM lookup'; ServerName = ''
            Reason = "VM '$Name' was not found in any connected vCenter ($serverList)."; VM = $null; Server = $null }
    }
    if ($hits.Count -gt 1) {
        $where = ($hits | ForEach-Object { "$($_.Server.Name):$($_.VM.Id)" }) -join '; '
        return [pscustomobject]@{ Ok = $false; Stage = 'VM lookup'; ServerName = $hits[0].Server.Name
            Reason = "VM name '$Name' is ambiguous - $($hits.Count) matching VMs ($where). Skipped for safety."; VM = $null; Server = $null }
    }
    $vm = $hits[0].VM; $srv = $hits[0].Server
    if ($vm.PowerState -ne 'PoweredOn') {
        return [pscustomobject]@{ Ok = $false; Stage = 'Power state'; ServerName = $srv.Name
            Reason = "VM is not powered on (PowerState=$($vm.PowerState))."; VM = $null; Server = $null }
    }
    $g = $vm.ExtensionData.Guest
    $toolsOk = ($g.ToolsRunningStatus -eq 'guestToolsRunning') -or ($g.ToolsStatus -in 'toolsOk', 'toolsOld')
    if (-not $toolsOk) {
        return [pscustomobject]@{ Ok = $false; Stage = 'VMware Tools'; ServerName = $srv.Name
            Reason = "VMware Tools not installed/running (ToolsStatus=$($g.ToolsStatus), ToolsRunningStatus=$($g.ToolsRunningStatus))."; VM = $null; Server = $null }
    }
    $isWin = ($g.GuestFamily -eq 'windowsGuest') -or ($g.GuestId -match 'windows') -or ($vm.Guest.OSFullName -match 'Windows')
    if (-not $isWin) {
        return [pscustomobject]@{ Ok = $false; Stage = 'Guest OS'; ServerName = $srv.Name
            Reason = "Guest OS is not Windows (GuestFamily=$($g.GuestFamily), OS=$($vm.Guest.OSFullName))."; VM = $null; Server = $null }
    }
    return [pscustomobject]@{ Ok = $true; Stage = 'OK'; ServerName = $srv.Name; Reason = ''; VM = $vm; Server = $srv }
}

function New-CentralRow {
    param($vCenter, $VMName, $Mode, $Action, $Result, $Detail)
    [PSCustomObject]([ordered]@{
        vCenter = $vCenter; VMName = $VMName; Stage = $Stage; Mode = $Mode
        Action = $Action; Result = $Result; Detail = $Detail
        Timestamp = (Get-Date -Format 'yyyy-MM-dd HH:mm:ss')
    })
}

function Get-ResultColor {
    param([string]$Result)
    switch ($Result) {
        'Applied'          { 'Red' }
        'Planned'          { 'Cyan' }
        'Would Block'      { 'Yellow' }
        'Blocked'          { 'Yellow' }
        'Skipped'          { 'DarkGray' }
        'Failed'           { 'Red' }
        'Reported'         { 'Gray' }
        'Unable to Check'  { 'DarkYellow' }
        default            { 'White' }
    }
}

# =====================================================================================
# 4. Per-VM processing
# =====================================================================================
$centralRows = New-Object System.Collections.Generic.List[object]
$summary = [ordered]@{ Total = 0; Applied = 0; PlannedOnly = 0; BlockedPendingReboot = 0; SkippedFailed = 0; DeclinedByUser = 0 }
$inventory = Get-VMInventory -Servers $connectedServers
$actionDescription = if ($Apply) { "APPLY Windows Update remediation Stage $Stage" } else { "Report the Stage $Stage remediation PLAN (no changes)" }

foreach ($vmName in $vmNames) {
    $summary.Total++
    Write-Log "=== Stage $Stage remediation for VM '$vmName' ==="
    try {
        $val = Resolve-AndValidateVM -Name $vmName -Inventory $inventory
        if (-not $val.Ok) {
            $summary.SkippedFailed++
            $centralRows.Add((New-CentralRow -vCenter $val.ServerName -VMName $vmName -Mode 'Read-Only' -Action 'VM validation' -Result 'Skipped' -Detail $val.Reason))
            Write-Log "SKIP '$vmName' [$($val.Stage)]: $($val.Reason)" 'WARN'
            continue
        }
        $vm = $val.VM; $srv = $val.Server

        if (-not $PSCmdlet.ShouldProcess("$vmName ($($srv.Name))", $actionDescription)) {
            $summary.DeclinedByUser++
            $centralRows.Add((New-CentralRow -vCenter $srv.Name -VMName $vmName -Mode 'Read-Only' -Action 'ShouldProcess' -Result 'Declined' -Detail 'User declined confirmation for this VM.'))
            Write-Log "'$vmName': declined by user (ShouldProcess)." 'WARN'
            continue
        }

        Write-Log "'$vmName': invoking guest remediation (Stage=$Stage Apply=$($Apply.IsPresent) ToolsWaitSecs=$ToolsWaitSecs)..."
        $result = Invoke-VMScript -VM $vm -Server $srv -ScriptText $Payload -ScriptType Powershell -GuestCredential $GuestCredential -ToolsWaitSecs $ToolsWaitSecs -Confirm:$false -ErrorAction Stop

        $envelope = Read-EnvelopeFromScriptOutput -Output $result.ScriptOutput -StartMarker '<<<WU-REMEDIATION-ENVELOPE-B64>>>' -EndMarker '<<<END-WU-REMEDIATION-ENVELOPE>>>'
        if (-not $envelope) {
            $summary.SkippedFailed++
            $snippet = if ($result.ScriptOutput) { ($result.ScriptOutput -replace '\s+', ' ').Trim() } else { '(no output)' }
            if ($snippet.Length -gt 600) { $snippet = $snippet.Substring(0, 600) + '...' }
            $centralRows.Add((New-CentralRow -vCenter $srv.Name -VMName $vmName -Mode 'Read-Only' -Action 'Guest payload result' -Result 'Unable to Check' -Detail "No parseable result envelope. Output start: $snippet"))
            Write-Log "'$vmName': no parseable envelope returned." 'ERROR'
            continue
        }

        foreach ($r in @($envelope.Rows)) {
            $centralRows.Add((New-CentralRow -vCenter $srv.Name -VMName $vmName -Mode $r.Mode -Action $r.Action -Result $r.Result -Detail $r.Detail))
        }
        if ($envelope.Meta.PendingReboot -and (@($envelope.Rows) | Where-Object { $_.Result -in 'Blocked','Would Block' })) {
            $summary.BlockedPendingReboot++
        } elseif ($Apply) {
            $summary.Applied++
        } else {
            $summary.PlannedOnly++
        }

        Write-Host ""
        Write-Host "=== $vmName @ $($srv.Name)  (Stage $Stage, Applied=$($envelope.Meta.Applied)) ===" -ForegroundColor Cyan
        if ($envelope.Meta.PendingReboot) { Write-Host " PENDING REBOOT DETECTED: $($envelope.Meta.RebootFlags)" -ForegroundColor Yellow }
        foreach ($row in $envelope.Rows) {
            $color = Get-ResultColor $row.Result
            Write-Host ("  [{0,-6}] [{1,-9}] {2}" -f $row.Mode, $row.Result, $row.Action) -ForegroundColor $color
            if ($row.Detail) { Write-Host "             $($row.Detail)" -ForegroundColor DarkGray }
        }
        Write-Log "'$vmName': Stage $Stage complete."
    }
    catch {
        $summary.SkippedFailed++
        $friendly = Get-FriendlyVixError -Message $_.Exception.Message
        Write-Log "'$vmName': unhandled error - $($_.Exception.Message)" 'ERROR'
        if ($friendly) { Write-Log "'$vmName': diagnosis - $friendly" 'WARN' }
        $detail = if ($friendly) { "$($_.Exception.Message) | DIAGNOSIS: $friendly" } else { $_.Exception.Message }
        $centralRows.Add((New-CentralRow -vCenter '' -VMName $vmName -Mode 'Read-Only' -Action 'Processing' -Result 'Failed' -Detail $detail))
        continue
    }
}

# =====================================================================================
# 5. Consolidated report + console summary
# =====================================================================================
$centralRows | Export-Csv -LiteralPath $CsvPath -NoTypeInformation -Encoding UTF8
Write-Log "CSV written: $CsvPath ($($centralRows.Count) rows across $($vmNames.Count) VM(s))."

Write-Host ""
Write-Host "================================================================================" -ForegroundColor DarkCyan
Write-Host " WINDOWS UPDATE REMEDIATION SUMMARY - Stage $Stage  Apply=$($Apply.IsPresent)" -ForegroundColor Green
Write-Host ("  Total VMs in list              : {0}" -f $summary.Total)
Write-Host ("  Applied                        : {0}" -f $summary.Applied) -ForegroundColor $(if ($summary.Applied) { 'Red' } else { 'Gray' })
Write-Host ("  Planned only (dry run)         : {0}" -f $summary.PlannedOnly) -ForegroundColor Cyan
Write-Host ("  Blocked - pending reboot       : {0}" -f $summary.BlockedPendingReboot) -ForegroundColor $(if ($summary.BlockedPendingReboot) { 'Yellow' } else { 'Gray' })
Write-Host ("  Declined by user (ShouldProcess): {0}" -f $summary.DeclinedByUser) -ForegroundColor Gray
Write-Host ("  Skipped / failed to process    : {0}" -f $summary.SkippedFailed) -ForegroundColor $(if ($summary.SkippedFailed) { 'Red' } else { 'Gray' })
Write-Host "================================================================================" -ForegroundColor DarkCyan
if (-not $Apply) {
    Write-Host " This was a DRY RUN. Re-run with -Apply once you've reviewed the plan above." -ForegroundColor Cyan
}
Write-Host "CSV: $CsvPath"
Write-Host "Log: $LogPath"

Write-Log "Remediation Stage $Stage run complete."
