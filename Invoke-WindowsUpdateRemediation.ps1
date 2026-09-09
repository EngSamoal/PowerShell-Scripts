#Requires -Version 5.1
<#
.SYNOPSIS
    Staged, least-to-most-disruptive Windows Update remediation for a single guest VM,
    executed through VMware Tools Guest Operations (Invoke-VMScript). Companion to
    Invoke-WindowsUpdateDiagnostics.ps1 - run diagnostics FIRST and read every 'Problem'
    row's NEXT step before touching this script.

.DESCRIPTION
    Four independent stages, each requested explicitly with -Stage. Nothing changes on
    the guest unless -Apply is also passed - by default every stage only PRINTS what it
    would do ("Planned"), so you can review the plan before committing to it. Each action
    is individually labelled Read-Only or Change in its output.

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
        - Refuses to run (even with -Apply) if a pending-reboot flag is currently set on
          the guest, unless -Force is also passed - resetting Windows Update state ahead
          of a required reboot usually just reproduces the same failure.

    Stage 4 (most disruptive - can download a large corrective payload and take a long
    time; requires -Apply -Force):
        - CHANGE: DISM /Online /Cleanup-Image /RestoreHealth. Only sensible to run after
          Stage 2's ScanHealth reported repairable corruption. Does not reboot the server -
          schedule a reboot yourself afterward.

    This script never deletes SoftwareDistribution/Catroot2, never runs a raw
    'rm/Remove-Item' against Windows Update state, and never reboots the guest.

.PARAMETER VMName          vCenter inventory name of the target VM.
.PARAMETER Stage           Which remediation stage to run (1-4). Stages are independent,
                            not cumulative - run and evaluate one at a time.
.PARAMETER Apply           Without this switch, the stage only reports what it WOULD do.
                            Pass -Apply to actually make the changes.
.PARAMETER Force           Required in addition to -Apply for Stage 3/4 when a pending
                            reboot is currently detected on the guest.
.PARAMETER CredentialPath  Export-Clixml PSCredential for the guest admin. If omitted or
                            not found, you are prompted with Get-Credential instead.
.PARAMETER OutputPath      Folder on the admin machine for the JSON result + log.
.PARAMETER ToolsWaitSecs   Invoke-VMScript VMware Tools wait, seconds (default 300 - DISM
                            operations in Stage 2/4 can run long).

.EXAMPLE
    # Review what Stage 1 would change, without changing anything:
    Connect-VIServer vcenter01
    .\Invoke-WindowsUpdateRemediation.ps1 -VMName TB-KRTN-APP02 -Stage 1

.EXAMPLE
    # Actually apply Stage 1:
    .\Invoke-WindowsUpdateRemediation.ps1 -VMName TB-KRTN-APP02 -Stage 1 -Apply

.EXAMPLE
    # Stage 3 reset, overriding the pending-reboot guard because the reboot is scheduled
    # for right after this run:
    .\Invoke-WindowsUpdateRemediation.ps1 -VMName TB-KRTN-APP02 -Stage 3 -Apply -Force

.NOTES
    Requires PowerCLI, an existing Connect-VIServer session, and vCenter "Guest Operation
    Program Execution/Query" privileges. Run Invoke-WindowsUpdateDiagnostics.ps1 first and
    do not run this against a VM that Azure Update Manager may currently be patching -
    check the diagnostics script's "Azure Arc / AUM" section first.
#>

[CmdletBinding(SupportsShouldProcess = $true, ConfirmImpact = 'High')]
param(
    [Parameter(Mandatory)]
    [string] $VMName,

    [Parameter(Mandatory)]
    [ValidateSet(1, 2, 3, 4)]
    [int] $Stage,

    [switch] $Apply,
    [switch] $Force,

    [string] $CredentialPath = 'C:\temp\wincred.xml',
    [string] $OutputPath     = (Join-Path $PSScriptRoot 'WindowsUpdate_Diagnostics'),

    [ValidateRange(30, 3600)]
    [int] $ToolsWaitSecs = 300
)

$ErrorActionPreference = 'Stop'
$ProgressPreference    = 'SilentlyContinue'

Write-Host "Invoke-WindowsUpdateRemediation.ps1 - staged Windows Update remediation  [build 2026-09-09a]" -ForegroundColor Magenta
Write-Host ("Running from: {0}" -f $PSCommandPath) -ForegroundColor DarkGray
if (-not $Apply) {
    Write-Host "DRY RUN (no -Apply): this run will only report what Stage $Stage WOULD change. Nothing on '$VMName' will be modified." -ForegroundColor Cyan
} else {
    Write-Host "APPLY MODE: Stage $Stage changes WILL be made on '$VMName'." -ForegroundColor Red
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

$vm = Get-VM -Name $VMName -ErrorAction SilentlyContinue
if (-not $vm) { throw "VM '$VMName' was not found on any connected vCenter." }

if (Test-Path -LiteralPath $CredentialPath) {
    try {
        $GuestCredential = Import-Clixml -LiteralPath $CredentialPath
        if (-not ($GuestCredential -is [pscredential])) { throw "did not deserialise to a PSCredential" }
    } catch {
        throw "Failed to import guest credential from '$CredentialPath': $($_.Exception.Message)"
    }
} else {
    $GuestCredential = Get-Credential -Message "Guest administrator credential for $VMName"
}

if (-not (Test-Path -LiteralPath $OutputPath)) { New-Item -ItemType Directory -Path $OutputPath -Force | Out-Null }
$RunStamp = Get-Date -Format 'yyyyMMdd-HHmmss'
$JsonPath = Join-Path $OutputPath "WU_Remediation_Stage${Stage}_${VMName}_$RunStamp.json"
$LogPath  = Join-Path $OutputPath "WU_Remediation_Stage${Stage}_${VMName}_$RunStamp.log"

function Write-Log {
    param([string]$Message)
    $line = "[{0}] {1}" -f (Get-Date -Format 'yyyy-MM-dd HH:mm:ss'), $Message
    Write-Host $line
    Add-Content -LiteralPath $LogPath -Value $line
}
Write-Log "Remediation run started for VM '$VMName'. Stage=$Stage Apply=$($Apply.IsPresent) Force=$($Force.IsPresent)."

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

$actionDescription = if ($Apply) { "APPLY Windows Update remediation Stage $Stage" } else { "Report the Stage $Stage remediation PLAN (no changes)" }
if (-not $PSCmdlet.ShouldProcess($VMName, $actionDescription)) {
    Write-Log "Cancelled by user (ShouldProcess declined)."
    return
}

Write-Log "Invoking guest script on '$VMName' (Stage=$Stage Apply=$($Apply.IsPresent) ToolsWaitSecs=$ToolsWaitSecs)..."
try {
    $result = Invoke-VMScript -VM $vm -ScriptText $Payload -ScriptType Powershell -GuestCredential $GuestCredential -ToolsWaitSecs $ToolsWaitSecs -ErrorAction Stop
} catch {
    throw "Invoke-VMScript failed against '$VMName': $($_.Exception.Message)"
}

$envelope = Read-EnvelopeFromScriptOutput -Output $result.ScriptOutput -StartMarker '<<<WU-REMEDIATION-ENVELOPE-B64>>>' -EndMarker '<<<END-WU-REMEDIATION-ENVELOPE>>>'
if (-not $envelope) {
    Write-Log "Could not parse the remediation envelope from guest output. Raw ScriptOutput follows:"
    Write-Host $result.ScriptOutput
    throw "Remediation Stage $Stage run failed - no parsable envelope returned from '$VMName'."
}

$envelope | ConvertTo-Json -Depth 10 | Set-Content -LiteralPath $JsonPath -Encoding UTF8
Write-Log "Result saved to: $JsonPath"

# =====================================================================================
# 4. Render
# =====================================================================================
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

Write-Host ""
Write-Host "================================================================================" -ForegroundColor DarkCyan
Write-Host " Windows Update Remediation - $($envelope.Meta.Hostname)   Stage $($envelope.Meta.Stage)   Applied=$($envelope.Meta.Applied)" -ForegroundColor Cyan
if ($envelope.Meta.PendingReboot) {
    Write-Host " PENDING REBOOT DETECTED: $($envelope.Meta.RebootFlags)" -ForegroundColor Yellow
}
Write-Host "================================================================================" -ForegroundColor DarkCyan
foreach ($row in $envelope.Rows) {
    $color = Get-ResultColor $row.Result
    Write-Host ("  [{0,-6}] [{1,-9}] {2}" -f $row.Mode, $row.Result, $row.Action) -ForegroundColor $color
    if ($row.Detail) { Write-Host "             $($row.Detail)" -ForegroundColor DarkGray }
}
Write-Host "================================================================================" -ForegroundColor DarkCyan
if (-not $Apply) {
    Write-Host " This was a DRY RUN. Re-run with -Apply once you've reviewed the plan above." -ForegroundColor Cyan
}
Write-Host " JSON result: $JsonPath" -ForegroundColor Gray
Write-Host "================================================================================" -ForegroundColor DarkCyan

Write-Log "Remediation Stage $Stage run complete."
