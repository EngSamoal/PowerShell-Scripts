#Requires -Version 5.1
#Requires -RunAsAdministrator
<#
.SYNOPSIS
    Check / start / auto-start IIS + Dynamics NAV services and IIS app pools after patching/reboot.
    Built for SF-NAV-APP01 (works on any IIS app server).

.DESCRIPTION
    After patching reboots, the IIS / NAV services or app pools sometimes stay stopped and have to
    be started by hand. This script does that for you and can make it permanent.

    -Mode Check      (default) Read-only. Shows services, app pools, sites, and every Automatic
                     service that is NOT running. Changes nothing.
    -Mode Start      One-time fix: starts the services (IIS first, then NAV), then any stopped
                     app pools and sites. Same as doing it by hand.
    -Mode Install    Permanent fix:
                       1. Services   -> Startup type Automatic (Delayed Start) + recovery
                                        "Restart the service" on 1st/2nd/later failures.
                       2. App pools  -> autoStart = true.
                       3. Sites      -> serverAutoStart = true.
                       4. Scheduled task "AppServiceAutoStart" (runs as SYSTEM at every boot,
                          waits -StartupDelaySeconds, then runs -Mode Start as a safety net).
                     Then runs -Mode Start once so everything is up now.
    -Mode Uninstall  Removes the scheduled task (service/app pool settings are left as they are).

    Log: C:\ProgramData\AppServiceAutoStart\AutoStart.log  (check this after the next reboot).

.EXAMPLE
    .\Set-NavAppAutoStart.ps1                    # check only
.EXAMPLE
    .\Set-NavAppAutoStart.ps1 -Mode Start        # start everything now
.EXAMPLE
    .\Set-NavAppAutoStart.ps1 -Mode Install      # make it start automatically after every restart
.EXAMPLE
    .\Set-NavAppAutoStart.ps1 -Mode Install -ServiceName 'W3SVC','WAS','MicrosoftDynamicsNav*','MyOtherSvc'
#>
[CmdletBinding()]
param(
    [ValidateSet('Check','Start','Install','Uninstall')]
    [string]  $Mode = 'Check',

    # Service names or wildcards. Order matters for starting: IIS core first, then NAV.
    [string[]]$ServiceName = @('WAS','W3SVC','MicrosoftDynamicsNav*','MicrosoftDynamics365BusinessCentral*'),

    [int]     $StartupDelaySeconds = 120,
    [int]     $ServiceStartTimeoutSeconds = 120
)

$ScriptBuild = 'Set-NavAppAutoStart 2026-11-23.1'
$TaskName    = 'AppServiceAutoStart'
$InstallDir  = Join-Path $env:ProgramData 'AppServiceAutoStart'
$LogFile     = Join-Path $InstallDir 'AutoStart.log'
$AppCmd      = Join-Path $env:windir 'System32\inetsrv\appcmd.exe'

if (-not (Test-Path $InstallDir)) { New-Item -ItemType Directory -Path $InstallDir -Force | Out-Null }

function Write-Log {
    param([string]$Message, [string]$Level = 'INFO')
    $line = '{0} [{1}] {2}' -f (Get-Date -Format 'yyyy-MM-dd HH:mm:ss'), $Level, $Message
    $color = switch ($Level) { 'ERROR' {'Red'} 'WARN' {'Yellow'} 'OK' {'Green'} default {'Gray'} }
    Write-Host $line -ForegroundColor $color
    try { Add-Content -Path $LogFile -Value $line -ErrorAction Stop } catch { }
}

# ---------------------------------------------------------------- helpers
function Get-TargetServices {
    $list = @()
    foreach ($pattern in $ServiceName) {
        $found = @(Get-Service -Name $pattern -ErrorAction SilentlyContinue)
        if ($found.Count -eq 0) { Write-Log "No service matches '$pattern' on this server (skipped)." 'WARN' }
        foreach ($s in $found) { if ($list.Name -notcontains $s.Name) { $list += $s } }
    }
    return ,$list
}

function Get-IisItems {
    param([ValidateSet('apppool','site')][string]$Kind)
    if (-not (Test-Path $AppCmd)) { return ,@() }
    try {
        [xml]$xml = & $AppCmd list $Kind /xml 2>$null
    } catch { Write-Log "appcmd list $Kind failed: $($_.Exception.Message)" 'ERROR'; return ,@() }
    $nodes = if ($Kind -eq 'apppool') { @($xml.appcmd.APPPOOL) } else { @($xml.appcmd.SITE) }
    $items = @()
    foreach ($n in $nodes) {
        if ($null -eq $n) { continue }
        $name = if ($Kind -eq 'apppool') { $n.'APPPOOL.NAME' } else { $n.'SITE.NAME' }
        $auto = if ($Kind -eq 'apppool') { & $AppCmd list apppool "$name" /text:autoStart 2>$null }
                else                     { & $AppCmd list site "$name" /text:serverAutoStart 2>$null }
        $items += [pscustomobject]@{ Name = $name; State = $n.state; AutoStart = "$auto".Trim() }
    }
    return ,$items
}

function Start-TargetService {
    param($Svc)
    $Svc.Refresh()
    if ($Svc.Status -eq 'Running') { Write-Log "Service $($Svc.Name) already running." 'OK'; return }
    if ($Svc.StartType -eq 'Disabled') { Write-Log "Service $($Svc.Name) is DISABLED - not starting it (fix startup type first or use -Mode Install)." 'WARN'; return }
    try {
        Write-Log "Starting service $($Svc.Name) ($($Svc.DisplayName))..."
        Start-Service -Name $Svc.Name -ErrorAction Stop
        $Svc.WaitForStatus('Running', [TimeSpan]::FromSeconds($ServiceStartTimeoutSeconds))
        Write-Log "Service $($Svc.Name) is Running." 'OK'
    } catch {
        Write-Log "Service $($Svc.Name) FAILED to start: $($_.Exception.Message). Check Event Viewer > Windows Logs > System/Application." 'ERROR'
    }
}

# ---------------------------------------------------------------- modes
function Invoke-Check {
    Write-Log '--- Target services ---'
    foreach ($s in (Get-TargetServices)) {
        $lvl = if ($s.Status -eq 'Running') { 'OK' } else { 'WARN' }
        Write-Log ("{0,-45} Status={1,-8} StartType={2}" -f $s.Name, $s.Status, $s.StartType) $lvl
    }
    if (Test-Path $AppCmd) {
        Write-Log '--- IIS application pools ---'
        foreach ($p in (Get-IisItems apppool)) {
            $lvl = if ($p.State -eq 'Started' -and $p.AutoStart -eq 'true') { 'OK' } else { 'WARN' }
            Write-Log ("{0,-45} State={1,-8} autoStart={2}" -f $p.Name, $p.State, $p.AutoStart) $lvl
        }
        Write-Log '--- IIS sites ---'
        foreach ($w in (Get-IisItems site)) {
            $lvl = if ($w.State -eq 'Started' -and $w.AutoStart -eq 'true') { 'OK' } else { 'WARN' }
            Write-Log ("{0,-45} State={1,-8} serverAutoStart={2}" -f $w.Name, $w.State, $w.AutoStart) $lvl
        }
    } else {
        Write-Log "IIS (appcmd.exe) not found on this server - skipping app pools/sites." 'WARN'
    }
    Write-Log '--- ALL Automatic services that are NOT running (may include normal trigger-start services) ---'
    $notRunning = @(Get-CimInstance Win32_Service -Filter "StartMode='Auto' AND State<>'Running'" -ErrorAction SilentlyContinue)
    if ($notRunning.Count -eq 0) { Write-Log 'None.' 'OK' }
    foreach ($s in $notRunning) { Write-Log ("{0,-45} {1}  ({2})" -f $s.Name, $s.State, $s.DisplayName) 'WARN' }
}

function Invoke-Start {
    Write-Log '--- Starting services ---'
    foreach ($s in (Get-TargetServices)) { Start-TargetService $s }

    if (-not (Test-Path $AppCmd)) { Write-Log 'IIS not installed - no app pools/sites to start.' 'WARN'; return }

    Write-Log '--- Starting app pools ---'
    foreach ($p in (Get-IisItems apppool)) {
        if ($p.State -eq 'Started') { Write-Log "App pool '$($p.Name)' already started." 'OK'; continue }
        $out = & $AppCmd start apppool /apppool.name:"$($p.Name)" 2>&1
        if ($LASTEXITCODE -eq 0) { Write-Log "App pool '$($p.Name)' started." 'OK' }
        else { Write-Log "App pool '$($p.Name)' FAILED to start: $out" 'ERROR' }
    }
    Write-Log '--- Starting sites ---'
    foreach ($w in (Get-IisItems site)) {
        if ($w.State -eq 'Started') { Write-Log "Site '$($w.Name)' already started." 'OK'; continue }
        $out = & $AppCmd start site /site.name:"$($w.Name)" 2>&1
        if ($LASTEXITCODE -eq 0) { Write-Log "Site '$($w.Name)' started." 'OK' }
        else { Write-Log "Site '$($w.Name)' FAILED to start: $out" 'ERROR' }
    }
}

function Invoke-Install {
    Write-Log '--- Configuring services: Automatic (Delayed Start) + restart on failure ---'
    foreach ($s in (Get-TargetServices)) {
        # WAS/W3SVC stay plain Automatic (other things depend on them); NAV services get Delayed Start
        # so they come up after SQL/network/IIS are ready - the usual cause of "stopped after patching".
        $startArg = if ($s.Name -in @('WAS','W3SVC')) { 'auto' } else { 'delayed-auto' }
        $o1 = & sc.exe config "$($s.Name)" start= $startArg 2>&1
        if ($LASTEXITCODE -eq 0) { Write-Log "$($s.Name): startup type set to $startArg." 'OK' }
        else { Write-Log "$($s.Name): could not set startup type: $o1" 'ERROR' }

        # Restart after 60s, 60s, 120s; reset failure counter after 1 day.
        $o2 = & sc.exe failure "$($s.Name)" reset= 86400 actions= restart/60000/restart/60000/restart/120000 2>&1
        if ($LASTEXITCODE -eq 0) { Write-Log "$($s.Name): recovery set to restart on failure." 'OK' }
        else { Write-Log "$($s.Name): could not set recovery actions: $o2" 'ERROR' }
        & sc.exe failureflag "$($s.Name)" 1 2>&1 | Out-Null   # also restart if it stops with an error code
    }

    if (Test-Path $AppCmd) {
        Write-Log '--- Configuring app pools (autoStart=true) ---'
        foreach ($p in (Get-IisItems apppool)) {
            $out = & $AppCmd set apppool "$($p.Name)" /autoStart:true 2>&1
            if ($LASTEXITCODE -eq 0) { Write-Log "App pool '$($p.Name)': autoStart=true." 'OK' }
            else { Write-Log "App pool '$($p.Name)': failed to set autoStart: $out" 'ERROR' }
        }
        Write-Log '--- Configuring sites (serverAutoStart=true) ---'
        foreach ($w in (Get-IisItems site)) {
            $out = & $AppCmd set site "$($w.Name)" /serverAutoStart:true 2>&1
            if ($LASTEXITCODE -eq 0) { Write-Log "Site '$($w.Name)': serverAutoStart=true." 'OK' }
            else { Write-Log "Site '$($w.Name)': failed to set serverAutoStart: $out" 'ERROR' }
        }
    }

    Write-Log "--- Registering startup scheduled task '$TaskName' ---"
    $target = Join-Path $InstallDir 'Set-NavAppAutoStart.ps1'
    if ($PSCommandPath -and ($PSCommandPath -ne $target)) { Copy-Item -Path $PSCommandPath -Destination $target -Force }
    $svcArg = ($ServiceName | ForEach-Object { "'$_'" }) -join ','
    $psArgs = "-NoProfile -ExecutionPolicy Bypass -Command `"Start-Sleep -Seconds $StartupDelaySeconds; & '$target' -Mode Start -ServiceName $svcArg`""
    try {
        $action    = New-ScheduledTaskAction -Execute 'powershell.exe' -Argument $psArgs
        $trigger   = New-ScheduledTaskTrigger -AtStartup
        $principal = New-ScheduledTaskPrincipal -UserId 'SYSTEM' -LogonType ServiceAccount -RunLevel Highest
        $settings  = New-ScheduledTaskSettingsSet -AllowStartIfOnBatteries -DontStopIfGoingOnBatteries -ExecutionTimeLimit (New-TimeSpan -Minutes 30)
        Register-ScheduledTask -TaskName $TaskName -Action $action -Trigger $trigger -Principal $principal -Settings $settings `
            -Description 'Starts IIS/NAV services and IIS app pools after reboot (safety net after patching).' -Force -ErrorAction Stop | Out-Null
        Write-Log "Scheduled task '$TaskName' registered (runs as SYSTEM at startup, $StartupDelaySeconds s delay)." 'OK'
    } catch {
        Write-Log "Failed to register scheduled task: $($_.Exception.Message)" 'ERROR'
    }

    Write-Log '--- Starting everything now ---'
    Invoke-Start
}

function Invoke-Uninstall {
    if (Get-ScheduledTask -TaskName $TaskName -ErrorAction SilentlyContinue) {
        Unregister-ScheduledTask -TaskName $TaskName -Confirm:$false
        Write-Log "Scheduled task '$TaskName' removed." 'OK'
    } else { Write-Log "Scheduled task '$TaskName' not found." 'WARN' }
}

# ---------------------------------------------------------------- main
Write-Log "$ScriptBuild  |  Host=$env:COMPUTERNAME  |  Mode=$Mode  |  User=$env:USERNAME"
switch ($Mode) {
    'Check'     { Invoke-Check }
    'Start'     { Invoke-Start; Invoke-Check }
    'Install'   { Invoke-Install; Invoke-Check }
    'Uninstall' { Invoke-Uninstall }
}
Write-Log "Done. Log file: $LogFile"
