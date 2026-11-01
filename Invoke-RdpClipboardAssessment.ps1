<#
.SYNOPSIS
    Read-only RDP clipboard / drive-redirection assessment of Windows Server VMs via VMware Tools
    (PowerCLI Invoke-VMScript). Produces a formatted Excel workbook under C:\temp.

.DESCRIPTION
    Runs from the management laptop using the EXISTING PowerCLI vCenter connections
    ($global:DefaultVIServers). No WinRM / PS Remoting / SSH / RDP is used.

    SAFETY - the script is strictly read-only:
      * Guest side: only reads registry, services, processes, event logs, files, ADMX, AppLocker
        effective policy, network listeners. No Set-*, no service/process control, no gpupdate.
      * vSphere side: only Get-VM / property reads (ExtraConfig read from the already-retrieved view).
      * Credentials are imported from an Export-Clixml file and never written to output or logs.
      * Invoke-VMScript itself creates and deletes a temporary script/output file in the guest
        %TEMP% folder - this is inherent to VMware Tools guest operations and is the only
        guest-side write that occurs.

.PARAMETER VMName
    One or more VM names (wildcards allowed) to assess.
.PARAMETER VMListFile
    Text file with one VM name per line (# comments allowed).
.PARAMETER AllWindowsVMs
    Assess every eligible Windows Server VM across all connected vCenters.
.PARAMETER IncludeWindowsClient
    Also include Windows client OS VMs (excluded by default - scope is Windows Server).
.PARAMETER EventLookbackDays
    How far back to search guest event logs (default 7).

.EXAMPLE
    .\Invoke-RdpClipboardAssessment.ps1 -VMName 'APP01'
.EXAMPLE
    .\Invoke-RdpClipboardAssessment.ps1 -VMName 'APP01','APP02','SQL*'
.EXAMPLE
    .\Invoke-RdpClipboardAssessment.ps1 -VMListFile C:\temp\vms.txt
.EXAMPLE
    .\Invoke-RdpClipboardAssessment.ps1 -AllWindowsVMs
#>
[CmdletBinding(DefaultParameterSetName = 'Named')]
param(
    [Parameter(ParameterSetName = 'Named', Position = 0)][string[]]$VMName,
    [Parameter(ParameterSetName = 'File', Mandatory)][string]$VMListFile,
    [Parameter(ParameterSetName = 'All', Mandatory)][switch]$AllWindowsVMs,
    [string]$CredentialPath = 'C:\temp\wincred.xml',
    [string]$OutputFolder = 'C:\temp',
    [ValidateRange(1, 90)][int]$EventLookbackDays = 7,
    [switch]$IncludeWindowsClient
)

$ErrorActionPreference = 'Stop'
$ts = Get-Date -Format 'yyyyMMdd_HHmmss'
if (-not (Test-Path $OutputFolder)) { throw "Output folder '$OutputFolder' does not exist." }
$LogFile = Join-Path $OutputFolder "RDP_Clipboard_Assessment_$ts.log"
$XlFile  = Join-Path $OutputFolder "RDP_Clipboard_Assessment_$ts.xlsx"

function Write-Log {
    param([string]$Msg, [string]$Level = 'INFO')
    $line = '{0} [{1}] {2}' -f (Get-Date -Format 'yyyy-MM-dd HH:mm:ss'), $Level, $Msg
    Add-Content -Path $LogFile -Value $line
    $color = @{ INFO = 'Gray'; OK = 'Green'; WARN = 'Yellow'; ERR = 'Red' }[$Level]
    Write-Host $line -ForegroundColor $color
}

#region ------------------------------------------------------------------ Pre-flight
if (-not $VMName -and -not $VMListFile -and -not $AllWindowsVMs) {
    throw 'Specify -VMName <names>, -VMListFile <file> or -AllWindowsVMs.'
}
if (-not (Get-Command Invoke-VMScript -ErrorAction SilentlyContinue)) {
    Import-Module VMware.VimAutomation.Core -ErrorAction Stop
}
$viServers = @($global:DefaultVIServers | Where-Object { $_.IsConnected })
if ($viServers.Count -eq 0) { throw 'No connected vCenter found in $global:DefaultVIServers. Connect-VIServer first.' }
Write-Log ("Using existing vCenter connections: {0}" -f (($viServers | ForEach-Object Name) -join ', '))

if (-not (Test-Path $CredentialPath)) { throw "Credential file '$CredentialPath' not found." }
$GuestCred = Import-Clixml -Path $CredentialPath
if ($GuestCred -isnot [System.Management.Automation.PSCredential]) { throw "'$CredentialPath' does not contain a PSCredential." }
Write-Log 'Guest credential loaded (value not logged).'

$HaveImportExcel = [bool](Get-Module -ListAvailable -Name ImportExcel)
if (-not $HaveImportExcel) { Write-Log 'ImportExcel module not found - report will fall back to CSV files. Install with: Install-Module ImportExcel -Scope CurrentUser' 'WARN' }
#endregion

#region ------------------------------------------------------------------ Guest payload (read-only)
$GuestPayload = @'
$ErrorActionPreference = 'Stop'
$ProgressPreference = 'SilentlyContinue'
$Lookback = [int]'__LOOKBACK__'
$F = New-Object System.Collections.ArrayList
$S = [ordered]@{ ClipPolicy='UNKNOWN'; DrivePolicy='UNKNOWN'; RdpService='UNKNOWN'; ClipProc='UNKNOWN'; Session='UNKNOWN'; Security='No restriction detected' }
$SecIssues = New-Object System.Collections.ArrayList

function Add-F { param($c,$s,$p,$a,$e,$st,$x,$r='',$rc='')
    [void]$F.Add([ordered]@{ c=$c; s=$s; p=$p; a=[string]$a; e=[string]$e; st=$st; x=$x; r=$r; rc=$rc }) }
function Invoke-Section { param($Name,[scriptblock]$Body)
    try { . $Body } catch { Add-F $Name 'Check execution' 'n/a' ('ERROR: ' + $_.Exception.Message) 'Check completes' 'UNKNOWN' 'The check failed to run (access denied or command error). It is NOT treated as a pass.' 'Re-run with a local Administrator credential / investigate the error.' } }
function Get-Reg { param($Path,$Name)
    try { $k = Get-Item -LiteralPath $Path -ErrorAction Stop }
    catch [System.Management.Automation.ItemNotFoundException] { return @{ State='KeyMissing'; Value=$null } }
    catch { return @{ State='Error'; Value=$null; Err=$_.Exception.Message } }
    try { $v = $k.GetValue($Name, $null) } catch { return @{ State='Error'; Value=$null; Err=$_.Exception.Message } }
    if ($null -eq $v) { @{ State='NotSet'; Value=$null } } else { @{ State='Set'; Value=$v } } }
function Show-Reg($r) { switch ($r.State) { 'Set' { [string]$r.Value } 'NotSet' { '(value not present)' } 'KeyMissing' { '(key not present)' } default { "(read error: $($r.Err))" } } }
function Get-Ev { param($Log,$Ids,$Match,$Level)
    $h = @{ LogName=$Log; StartTime=(Get-Date).AddDays(-$Lookback) }
    if ($Ids) { $h.Id = $Ids }; if ($Level) { $h.Level = $Level }
    try { $e = @(Get-WinEvent -FilterHashtable $h -MaxEvents 3000 -ErrorAction Stop) }
    catch {
        $id = [string]$_.FullyQualifiedErrorId; $m = $_.Exception.Message
        if ($id -like 'NoMatchingEventsFound*' -or $m -match 'No events were found') { return @{ State='OK'; Events=@() } }
        if ($id -like 'NoMatchingLogsFound*' -or $m -match 'There is not an event log') { return @{ State='NoLog'; Events=@() } }
        return @{ State='Error'; Events=@(); Err=$m } }
    if ($Match) { $e = @($e | Where-Object { $_.Message -match $Match }) }
    @{ State='OK'; Events=$e } }
function Short($t) { $t = ([string]$t) -replace '\s+',' '; if ($t.Length -gt 300) { $t.Substring(0,300) + '...' } else { $t } }

# ---------------------------------------------------------------- Context
$M = [ordered]@{}
Invoke-Section 'Execution Context' {
    $os = Get-CimInstance Win32_OperatingSystem
    $cs = Get-CimInstance Win32_ComputerSystem
    $idn = [Security.Principal.WindowsIdentity]::GetCurrent()
    $elev = (New-Object Security.Principal.WindowsPrincipal $idn).IsInRole([Security.Principal.WindowsBuiltInRole]::Administrator)
    $M.Host = $env:COMPUTERNAME; $M.OS = $os.Caption; $M.Build = $os.Version; $M.LastBoot = [string]$os.LastBootUpTime
    $M.Elevated = $elev; $M.PartOfDomain = $cs.PartOfDomain; $M.PS = $PSVersionTable.PSVersion.ToString()
    if ($elev) { Add-F 'Execution Context' 'Guest script runs with elevated Administrator token' 'WindowsPrincipal.IsInRole(Administrators)' 'True' 'True' 'PASS' 'All checks can read protected logs and policy.' }
    else { Add-F 'Execution Context' 'Guest script runs with elevated Administrator token' 'WindowsPrincipal.IsInRole(Administrators)' 'False' 'True' 'WARNING' 'Guest operations token is not elevated (UAC-filtered non-builtin admin, or non-admin account). Some checks may return UNKNOWN (access denied) - these are never counted as PASS.' 'Use the built-in Administrator or an account whose token is elevated for guest operations.' }
    Add-F 'Execution Context' 'Workgroup membership' 'Win32_ComputerSystem.PartOfDomain' $cs.PartOfDomain 'False (workgroup)' 'NOT APPLICABLE' 'Workgroup host: only Local Group Policy (and MDM, if enrolled) can set RDS policy; no domain GPO precedence applies.'
}

# ---------------------------------------------------------------- ADMX map (proves policy names/values on THIS OS)
$Admx = @{}
$TSPk = 'SOFTWARE\Policies\Microsoft\Windows NT\Terminal Services'
$TSP  = 'HKLM:\' + $TSPk
Invoke-Section 'A. Policy definitions (ADMX)' {
    $ax = "$env:windir\PolicyDefinitions\TerminalServer.admx"; $al = "$env:windir\PolicyDefinitions\en-US\TerminalServer.adml"
    if (-not (Test-Path $ax)) { Add-F 'A. Policy definitions (ADMX)' 'TerminalServer.admx present' $ax 'Missing' 'Present' 'WARNING' 'ADMX not found; policy display names cannot be resolved (registry values are still evaluated).'; return }
    [xml]$x = Get-Content -LiteralPath $ax -Raw
    $str = @{}
    if (Test-Path $al) { [xml]$l = Get-Content -LiteralPath $al -Raw; foreach ($strE in $l.policyDefinitionResources.resources.stringTable.string) { $str[$strE.id] = $strE.'#text' } }
    $res = { param($r) if ($r -match '^\$\(string\.(.+)\)$' -and $str.ContainsKey($Matches[1])) { $str[$Matches[1]] } else { $r } }
    foreach ($p in $x.policyDefinitions.policies.policy) {
        $disp = & $res $p.displayName
        if ($p.valueName) { $Admx["$($p.key)|$($p.valueName)"] = @{ Policy=$disp; Class=$p.class; Enum=$null } }
        if ($p.elements) { foreach ($el in $p.elements.ChildNodes) { if ($el.valueName) {
            $map = @{}; foreach ($it in @($el.item)) { if ($it) { $map[[string]$it.value.decimal.value] = & $res $it.displayName } }
            $Admx["$($p.key)|$($el.valueName)"] = @{ Policy=$disp; Class=$p.class; Enum=$map } } } }
    }
    Add-F 'A. Policy definitions (ADMX)' 'TerminalServer.admx parsed' $ax ("{0} value mappings" -f $Admx.Count) 'Parsed' 'PASS' 'Policy names and enumerations below are resolved from this server''s own ADMX/ADML, not assumed.'
}
function PolName($v) { $k = "$TSPk|$v"; if ($Admx.ContainsKey($k)) { "'" + $Admx[$k].Policy + "'" } else { '(not defined in this OS TerminalServer.admx)' } }

# ---------------------------------------------------------------- Local GPO Registry.pol (configured vs. effective)
$PolText = $null; $PolState = 'Missing'
Invoke-Section 'A. Local Group Policy source' {
    $pp = "$env:windir\System32\GroupPolicy\Machine\Registry.pol"
    if (Test-Path -LiteralPath $pp) { $script:PolText = [Text.Encoding]::Unicode.GetString([IO.File]::ReadAllBytes($pp)); $script:PolState = 'Read' }
    $st = if ($script:PolState -eq 'Read') { 'PASS' } else { 'NOT APPLICABLE' }
    Add-F 'A. Local Group Policy source' 'Local GPO machine Registry.pol' $pp $script:PolState 'Readable (if LGPO used)' $st 'Used to show whether an effective registry value originates from Local Group Policy or was written directly to the registry (e.g. by a hardening script).'
}
function In-Pol($n) { if ($PolState -ne 'Read') { "LGPO file: $PolState" } elseif ($PolText -match ('\*\*del\.' + [regex]::Escape($n))) { 'LGPO: delete-marker for value' } elseif ($PolText -match [regex]::Escape($n)) { 'LGPO: configured in Registry.pol' } else { 'LGPO: not in Registry.pol' } }

# Generic evaluator for "Do not allow X redirection" style policies (1 = BLOCK)
function Test-DisablePolicy { param($Name,$Cat,$Label,$Impact,$Fix,$Rc)
    $r = Get-Reg $TSP $Name
    $src = "HKLM\$TSPk\$Name  [$(In-Pol $Name)]"
    $pn = PolName $Name
    switch -Regex ($r.State) {
        '^Set$' {
            if ([int]$r.Value -eq 1) { Add-F $Cat "$Label policy $pn" $src 1 '0 (Disabled) or not configured' 'FAIL' "Policy is set to ENABLED (1). For this 'Do not allow...' policy, Enabled means the redirection is BLOCKED. $Impact" $Fix $Rc; return 'Blocked' }
            elseif ([int]$r.Value -eq 0) { Add-F $Cat "$Label policy $pn" $src 0 '0 (Disabled) or not configured' 'PASS' 'Policy is explicitly DISABLED (0) = redirection allowed by policy. Policy takes precedence over the listener (RDP-Tcp) setting.'; return 'Allowed' }
            else { Add-F $Cat "$Label policy $pn" $src $r.Value '0 or 1' 'WARNING' 'Unexpected value; behaviour undefined.' 'Correct the value via Local Group Policy.'; return 'Unknown' } }
        '^(NotSet|KeyMissing)$' { Add-F $Cat "$Label policy $pn" $src (Show-Reg $r) '0 (Disabled) or not configured' 'NOT CONFIGURED' 'No policy restriction. The listener-level (RDP-Tcp) value below is the effective server-side setting.'; return 'NotConfigured' }
        default { Add-F $Cat "$Label policy $pn" $src (Show-Reg $r) '0 or not configured' 'UNKNOWN' 'Registry could not be read; effective state unknown.' 'Re-run with elevated credential.'; return 'Unknown' }
    } }
function Test-ListenerFlag { param($Name,$Cat,$Label,$PolicyState,$Fix,$Rc)
    $out = 'Unknown'
    $root = 'HKLM:\SYSTEM\CurrentControlSet\Control\Terminal Server\WinStations'
    foreach ($l in @(Get-ChildItem -LiteralPath $root | Where-Object { $null -ne $_.GetValue('PortNumber', $null) })) {
        $r = Get-Reg $l.PSPath $Name
        $src = "HKLM\SYSTEM\CurrentControlSet\Control\Terminal Server\WinStations\$($l.PSChildName)\$Name"
        if ($PolicyState -in 'Blocked','Allowed') { Add-F $Cat "$Label - listener $($l.PSChildName)" $src (Show-Reg $r) '0 / not present' 'NOT APPLICABLE' 'Overridden: the computer policy above is configured and takes precedence over listener settings.'; $out = $PolicyState; continue }
        if ($r.State -eq 'Error') { Add-F $Cat "$Label - listener $($l.PSChildName)" $src (Show-Reg $r) '0 / not present' 'UNKNOWN' 'Could not read listener setting.'; continue }
        if ($r.State -eq 'Set' -and [int]$r.Value -eq 1) { Add-F $Cat "$Label - listener $($l.PSChildName)" $src 1 '0 / not present' 'FAIL' "Listener-level setting disables this redirection and no policy overrides it - effective BLOCK for sessions on this listener." $Fix $Rc; $out = 'Blocked' }
        else { Add-F $Cat "$Label - listener $($l.PSChildName)" $src (Show-Reg $r) '0 / not present' 'PASS' 'Listener does not disable this redirection.'; if ($out -ne 'Blocked') { $out = 'Allowed' } }
    }
    $out }

$FixClip  = "gpedit.msc > Computer Configuration > Administrative Templates > Windows Components > Remote Desktop Services > Remote Desktop Session Host > Device and Resource Redirection > 'Do not allow Clipboard redirection' = Disabled or Not Configured; then gpupdate /force and have the user sign out/reconnect (applies to new sessions). Change-controlled - not done by this script."
$FixDrive = "Same GPO path > 'Do not allow drive redirection' = Disabled or Not Configured; gpupdate /force; reconnect session. Required only if file transfer via redirected drives (\\tsclient) is in scope and approved."

# ---------------------------------------------------------------- A. Clipboard / drive policies
$ClipState = 'Unknown'; $DriveState = 'Unknown'
Invoke-Section 'A. RDP Clipboard Policy' {
    $p = Test-DisablePolicy 'fDisableClip' 'A. RDP Clipboard Policy' 'Do not allow Clipboard redirection' 'Blocks BOTH text and file clipboard redirection for all RDP sessions, regardless of BeyondTrust settings.' $FixClip 'Confirmed'
    $l = Test-ListenerFlag 'fDisableClip' 'A. RDP Clipboard Policy' 'Clipboard mapping disabled (Tsconfig client setting)' $p $FixClip 'Confirmed'
    $ClipState = if ($p -in 'Blocked','Allowed') { $p } else { $l }
    $S.ClipPolicy = switch ($ClipState) { 'Blocked' { 'BLOCKED' } 'Allowed' { 'Allowed' } default { 'UNKNOWN' } }
    if ($p -eq 'NotConfigured' -and $ClipState -eq 'Allowed') { $S.ClipPolicy = 'Allowed (policy not configured, listener allows)' }
    if ($p -eq 'Blocked') { $S.ClipPolicy = 'BLOCKED by policy fDisableClip=1' }

    # Directional clipboard format restriction policies (only on OS builds whose ADMX defines them)
    foreach ($d in @(@{ n='CSClipLevel'; t='client -> server (laptop to VM)' }, @{ n='SCClipLevel'; t='server -> client (VM to laptop)' })) {
        $def = $Admx["$TSPk|$($d.n)"]
        $r = Get-Reg $TSP $d.n
        $src = "HKLM\$TSPk\$($d.n)  [$(In-Pol $d.n)]"
        $nm = if ($def) { "Policy '$($def.Policy)'" } else { "$($d.n) (policy not defined in this OS ADMX)" }
        if ($r.State -eq 'Set') {
            $txt = if ($def -and $def.Enum -and $def.Enum.ContainsKey([string][int]$r.Value)) { $def.Enum[[string][int]$r.Value] } else { $null }
            if ($txt -and $txt -match 'Disable') { Add-F 'A. RDP Clipboard Policy' $nm $src "$($r.Value) = $txt" 'Not configured' 'FAIL' "Clipboard transfers $($d.t) are DISABLED by policy." "Set this policy to Not Configured (or a level that permits the required formats) via gpedit, same Device and Resource Redirection path; reconnect session." 'Confirmed'; $S.ClipPolicy += " | $($d.n)=blocked" }
            elseif ($txt) { Add-F 'A. RDP Clipboard Policy' $nm $src "$($r.Value) = $txt" 'Not configured' 'WARNING' "Clipboard $($d.t) is restricted to the listed formats only. Plain text may work, but other formats - including FILE copy/paste - are blocked in this direction." 'If file copy is required, set to Not Configured (change-controlled).' 'Potential-Files'; $S.ClipPolicy += " | $($d.n) restricted" }
            else { Add-F 'A. RDP Clipboard Policy' $nm $src $r.Value 'Not configured' 'WARNING' 'Value present but could not be mapped to an ADMX option on this OS - effect cannot be confirmed.' 'Review in gpedit / remove if not intended.' 'Potential' }
        }
        elseif ($r.State -eq 'Error') { Add-F 'A. RDP Clipboard Policy' $nm $src (Show-Reg $r) 'Not configured' 'UNKNOWN' 'Could not read.' }
        elseif ($def) { Add-F 'A. RDP Clipboard Policy' $nm $src (Show-Reg $r) 'Not configured' 'NOT CONFIGURED' "No format restriction for $($d.t)." }
        else { Add-F 'A. RDP Clipboard Policy' $nm $src (Show-Reg $r) 'Not configured' 'NOT APPLICABLE' 'Policy does not exist on this OS build and no value is present.' }
    }
}
Invoke-Section 'A. RDP Drive Redirection Policy' {
    $p = Test-DisablePolicy 'fDisableCdm' 'A. RDP Drive Redirection Policy' 'Do not allow drive redirection' 'Blocks redirected client drives (\\tsclient\X). Text clipboard is NOT affected; file transfer that relies on redirected drives is blocked, and some client/OS combinations also depend on it for file copy/paste.' $FixDrive 'Confirmed-Files'
    $l = Test-ListenerFlag 'fDisableCdm' 'A. RDP Drive Redirection Policy' 'Drive mapping disabled (Tsconfig client setting)' $p $FixDrive 'Confirmed-Files'
    $DriveState = if ($p -in 'Blocked','Allowed') { $p } else { $l }
    $S.DrivePolicy = switch ($DriveState) { 'Blocked' { 'BLOCKED' } 'Allowed' { 'Allowed' } default { 'UNKNOWN' } }
}

# ---------------------------------------------------------------- A. Other RDS policy values, MDM
Invoke-Section 'A. Other RDS policy values' {
    $handled = 'fDisableClip','fDisableCdm','CSClipLevel','SCClipLevel','fDenyTSConnections'
    try { $k = Get-Item -LiteralPath $TSP -ErrorAction Stop } catch [System.Management.Automation.ItemNotFoundException] { $k = $null }
    if ($k) { foreach ($vn in $k.GetValueNames()) { if ($handled -notcontains $vn) {
        Add-F 'A. Other RDS policy values' "$vn $(PolName $vn)" "HKLM\$TSPk\$vn  [$(In-Pol $vn)]" $k.GetValue($vn) 'n/a' 'NOT APPLICABLE' 'Configured RDS policy with no documented effect on clipboard redirection; listed for completeness and review.' } } }
    $mdm = 'HKLM:\SOFTWARE\Microsoft\PolicyManager\current\device\RemoteDesktopServices'
    if (Test-Path -LiteralPath $mdm) {
        $mk = Get-Item -LiteralPath $mdm
        $vals = ($mk.GetValueNames() | Where-Object { $_ -notmatch '_(ProviderSet|WinningProvider|LastWrite)' } | ForEach-Object { "$_=$($mk.GetValue($_))" }) -join '; '
        Add-F 'A. Other RDS policy values' 'MDM (Policy CSP) RemoteDesktopServices settings' $mdm $vals 'None, or not restricting redirection' 'WARNING' 'An MDM/Intune policy area for RDS exists. MDM-delivered ADMX policies write to the same Terminal Services policy key and can re-apply after local changes.' 'Confirm the owning MDM policy with the endpoint management team.' 'Potential'
    } else { Add-F 'A. Other RDS policy values' 'MDM (Policy CSP) RemoteDesktopServices settings' $mdm '(not present)' 'Not present' 'PASS' 'No MDM-delivered RDS policy area found.' }
}

# ---------------------------------------------------------------- C. RDP enabled / listener
$RdpPort = 3389
Invoke-Section 'C. RDP Session Configuration' {
    $pol = Get-Reg $TSP 'fDenyTSConnections'
    $sys = Get-Reg 'HKLM:\SYSTEM\CurrentControlSet\Control\Terminal Server' 'fDenyTSConnections'
    $eff = if ($pol.State -eq 'Set') { [int]$pol.Value } elseif ($sys.State -eq 'Set') { [int]$sys.Value } else { $null }
    $src = "Policy: HKLM\$TSPk\fDenyTSConnections = $(Show-Reg $pol); System: HKLM\SYSTEM\CurrentControlSet\Control\Terminal Server\fDenyTSConnections = $(Show-Reg $sys)"
    if ($eff -eq 0) { Add-F 'C. RDP Session Configuration' 'Remote Desktop connections allowed' $src 0 0 'PASS' 'RDP is enabled (policy value wins over system value when present).' }
    elseif ($eff -eq 1) { Add-F 'C. RDP Session Configuration' 'Remote Desktop connections allowed' $src 1 0 'FAIL' 'RDP is disabled. BeyondTrust RDP Jump Items cannot connect at all - this is a connectivity problem, not a clipboard-specific one.' 'Enable RDP only if approved.' }
    else { Add-F 'C. RDP Session Configuration' 'Remote Desktop connections allowed' $src 'unknown' 0 'UNKNOWN' 'Could not determine effective value.' }

    $l = 'HKLM:\SYSTEM\CurrentControlSet\Control\Terminal Server\WinStations\RDP-Tcp'
    $pn = Get-Reg $l 'PortNumber'; $en = Get-Reg $l 'fEnableWinStation'
    if ($pn.State -eq 'Set') { $RdpPort = [int]$pn.Value }
    $st = if ($en.State -eq 'Set' -and [int]$en.Value -eq 0) { 'FAIL' } elseif ($en.State -eq 'Error') { 'UNKNOWN' } else { 'PASS' }
    Add-F 'C. RDP Session Configuration' 'RDP-Tcp listener enabled / port' "$l\fEnableWinStation, PortNumber" "fEnableWinStation=$(Show-Reg $en); PortNumber=$(Show-Reg $pn)" 'fEnableWinStation=1' $st 'Listener state (connectivity prerequisite, not a clipboard control).'
    try {
        $lis = @(Get-NetTCPConnection -State Listen -LocalPort $RdpPort -ErrorAction Stop)
        Add-F 'C. RDP Session Configuration' "TCP $RdpPort listening" "Get-NetTCPConnection -State Listen -LocalPort $RdpPort" "Listening on $(($lis | ForEach-Object LocalAddress | Select-Object -Unique) -join ', ')" 'Listening' 'PASS' 'RDP listener is bound.'
    } catch { Add-F 'C. RDP Session Configuration' "TCP $RdpPort listening" "Get-NetTCPConnection -State Listen -LocalPort $RdpPort" 'Not listening / not queryable' 'Listening' 'WARNING' 'Listener not found on the configured port (connectivity, not clipboard).' }
}

# ---------------------------------------------------------------- B. Services
Invoke-Section 'B. RDP Services' {
    $svcTxt = @()
    foreach ($n in 'TermService','UmRdpService','SessionEnv') {
        $svc = Get-CimInstance Win32_Service -Filter "Name='$n'"
        if (-not $svc) { Add-F 'B. RDP Services' $n "Win32_Service Name='$n'" 'Not installed' 'Installed' $(if ($n -eq 'SessionEnv') { 'WARNING' } else { 'FAIL' }) 'Required RDS service is missing.' 'Investigate OS integrity (sfc /scannow, change-controlled).' $(if ($n -eq 'TermService') { 'Confirmed' } else { 'Potential' }); continue }
        $val = "State=$($svc.State); StartMode=$($svc.StartMode)"; $svcTxt += "$n=$($svc.State)/$($svc.StartMode)"
        switch ($n) {
            'TermService' {
                if ($svc.State -eq 'Running') { Add-F 'B. RDP Services' 'Remote Desktop Services (TermService)' 'Win32_Service TermService' $val 'Running (StartMode Manual is the Windows default)' 'PASS' 'Core RDP service running. Manual start mode is normal - it is started on demand/at boot by the system.' }
                else { Add-F 'B. RDP Services' 'Remote Desktop Services (TermService)' 'Win32_Service TermService' $val 'Running' 'FAIL' 'TermService is not running - no RDP sessions possible.' 'Start-Service TermService (change-controlled; not executed).' 'Confirmed' } }
            'UmRdpService' {
                if ($svc.StartMode -eq 'Disabled') { Add-F 'B. RDP Services' 'RDS UserMode Port Redirector (UmRdpService)' 'Win32_Service UmRdpService' $val 'Manual, Running when sessions exist' 'FAIL' 'Disabled. This service handles redirection of drives, printers and ports; drive redirection / \\tsclient file transfer will fail. Text clipboard (rdpclip) does not depend on it.' "Set-Service UmRdpService -StartupType Manual (change-controlled; not executed)." 'Confirmed-Files' }
                elseif ($svc.State -eq 'Running') { Add-F 'B. RDP Services' 'RDS UserMode Port Redirector (UmRdpService)' 'Win32_Service UmRdpService' $val 'Running' 'PASS' 'Device redirection helper is running.' }
                else { Add-F 'B. RDP Services' 'RDS UserMode Port Redirector (UmRdpService)' 'Win32_Service UmRdpService' $val 'Manual; started by TermService' 'WARNING' 'Not running. It normally runs alongside TermService; if RDP sessions are active this impacts drive/printer redirection (not text clipboard).' 'Check System log 7000/7023/7031 for this service.' 'Potential-Files' } }
            'SessionEnv' {
                $st = if ($svc.StartMode -eq 'Disabled') { 'WARNING' } else { 'PASS' }
                Add-F 'B. RDP Services' 'Remote Desktop Configuration (SessionEnv)' 'Win32_Service SessionEnv' $val 'Manual (trigger/on demand)' $st 'Configuration helper; Manual/Stopped is normal. Not a clipboard dependency.' }
        }
    }
    $S.RdpService = $svcTxt -join '; '
}

# ---------------------------------------------------------------- C. Sessions (+ source address from LSM log)
$Sessions = @(); $RdpActive = @()
Invoke-Section 'C. RDP Sessions' {
    $q = @(& "$env:windir\System32\qwinsta.exe" 2>&1 | ForEach-Object { [string]$_ })
    $hdr = $q | Where-Object { $_ -match 'SESSIONNAME' } | Select-Object -First 1
    if (-not $hdr) { Add-F 'C. RDP Sessions' 'Session enumeration' 'qwinsta.exe' (Short ($q -join ' ')) 'Session table' 'UNKNOWN' 'qwinsta output could not be parsed (non-English OS or error). rdpclip-per-session check cannot be evaluated.'; $S.Session = 'UNKNOWN'; return }
    $pu = $hdr.IndexOf('USERNAME')
    foreach ($ln in $q) {
        if ($ln -eq $hdr -or $ln.Length -le $pu) { continue }
        $nm = $ln.Substring(1, $pu - 1).Trim(); $rest = $ln.Substring($pu)
        if ($rest -match '^\s*(\d+)\s+(\S+)') { $u = ''; $id = [int]$Matches[1]; $stt = $Matches[2] }
        elseif ($rest -match '^\s*(\S+)\s+(\d+)\s+(\S+)') { $u = $Matches[1]; $id = [int]$Matches[2]; $stt = $Matches[3] } else { continue }
        $script:Sessions += [pscustomobject]@{ Name=$nm; User=$u; Id=$id; State=$stt; Addr='' }
    }
    $lsm = Get-Ev 'Microsoft-Windows-TerminalServices-LocalSessionManager/Operational' @(21,25)
    if ($lsm.State -eq 'OK') {
        foreach ($e in ($lsm.Events | Sort-Object TimeCreated)) { try { $x = ([xml]$e.ToXml()).Event.UserData.EventXML
            foreach ($sx in $script:Sessions) { if ([string]$sx.Id -eq [string]$x.SessionID) { $sx.Addr = [string]$x.Address } } } catch {} }
    }
    $script:RdpActive = @($script:Sessions | Where-Object { $_.Name -like 'rdp-tcp#*' -and $_.State -eq 'Active' })
    $disc = @($script:Sessions | Where-Object { $_.State -like 'Disc*' -and $_.Id -ne 0 })
    $list = ($script:Sessions | Where-Object { $_.Id -ne 0 -and $_.State -ne 'Listen' } | ForEach-Object { "#$($_.Id) $($_.Name) $($_.User) $($_.State)$(if ($_.Addr) { ' from ' + $_.Addr })" }) -join '; '
    $S.Session = "Active RDP: $($RdpActive.Count); Disconnected: $($disc.Count)"
    Add-F 'C. RDP Sessions' 'Session enumeration' 'qwinsta.exe + LocalSessionManager/Operational events 21/25 (source address)' $(if ($list) { $list } else { '(no user sessions)' }) 'n/a' 'PASS' 'Sessions enumerated. Source address identifies whether a session came from the BeyondTrust Jumpoint (Jumpoint IP) or a direct RDP client.'
}

# ---------------------------------------------------------------- B. rdpclip per session
Invoke-Section 'B. Clipboard Process' {
    $clip = @(Get-CimInstance Win32_Process -Filter "Name='rdpclip.exe'")
    if ($S.Session -eq 'UNKNOWN') { Add-F 'B. Clipboard Process' 'rdpclip.exe in active RDP session(s)' 'Win32_Process rdpclip.exe vs qwinsta' "rdpclip PIDs/sessions: $(($clip | ForEach-Object { "$($_.ProcessId)/s$($_.SessionId)" }) -join ', ')" 'Running in each active RDP session' 'UNKNOWN' 'Sessions could not be enumerated, so per-session presence cannot be judged.'; $S.ClipProc = 'UNKNOWN'; return }
    if ($RdpActive.Count -eq 0) { Add-F 'B. Clipboard Process' 'rdpclip.exe in active RDP session(s)' 'Win32_Process rdpclip.exe vs qwinsta' "No active RDP session (rdpclip instances: $($clip.Count))" 'n/a' 'NOT APPLICABLE' 'rdpclip.exe is started per RDP session; without an active RDP session its absence is expected. Re-run the script WHILE the BeyondTrust session is connected for this check to be meaningful.' 'Re-run during a live BeyondTrust session.'; $S.ClipProc = 'N/A - no active RDP session'; return }
    $res = @()
    foreach ($rs in $RdpActive) {
        $m = @($clip | Where-Object { $_.SessionId -eq $rs.Id })
        $who = "session $($rs.Id) ($($rs.Name), user $($rs.User)$(if ($rs.Addr) { ', from ' + $rs.Addr }))"
        if ($m.Count) { Add-F 'B. Clipboard Process' "rdpclip.exe in $who" 'Win32_Process rdpclip.exe SessionId' "Running, PID $(($m | ForEach-Object ProcessId) -join ',')" 'Running in session' 'PASS' 'Server-side clipboard component is running in this RDP session. NOTE: this does not prove the client (BeyondTrust) negotiated/enabled the clipboard channel.'; $res += "s$($rs.Id):running" }
        else { Add-F 'B. Clipboard Process' "rdpclip.exe in $who" 'Win32_Process rdpclip.exe SessionId' 'NOT running in this session' 'Running in session' 'WARNING' 'Server-side clipboard component is absent in an active RDP session. Causes: clipboard blocked by policy (section A), execution blocked (section D), crash (event checks), or the RDP client/BeyondTrust did not request clipboard redirection. Not a confirmed root cause by itself.' 'In that session, the user can launch rdpclip.exe (Win+R) to test; compare with a direct mstsc session with Clipboard ticked. Not executed by this script.' 'Potential'; $res += "s$($rs.Id):MISSING" }
    }
    $S.ClipProc = $res -join '; '
}

# ---------------------------------------------------------------- D. Execution restrictions affecting rdpclip.exe
Invoke-Section 'D. Security / Restrictions' {
    $exe = "$env:windir\System32\rdpclip.exe"
    if (Test-Path -LiteralPath $exe) {
        $fv = (Get-Item -LiteralPath $exe).VersionInfo.FileVersion
        $sig = (Get-AuthenticodeSignature -FilePath $exe).Status
        $st = if ($sig -eq 'Valid') { 'PASS' } else { 'WARNING' }
        Add-F 'D. Security / Restrictions' 'rdpclip.exe binary present and signed' $exe "Version $fv; Signature $sig" 'Present; Valid' $st 'Binary integrity check.' $(if ($st -ne 'PASS') { 'Run sfc /scannow (change-controlled).' } else { '' }) $(if ($st -ne 'PASS') { 'Potential' } else { '' })
    } else { Add-F 'D. Security / Restrictions' 'rdpclip.exe binary present and signed' $exe 'MISSING' 'Present' 'FAIL' 'rdpclip.exe is missing - clipboard redirection cannot work.' 'Repair OS files (sfc /scannow / DISM), change-controlled.' 'Confirmed'; [void]$SecIssues.Add('rdpclip.exe missing') }

    $ifeo = 'HKLM:\SOFTWARE\Microsoft\Windows NT\CurrentVersion\Image File Execution Options\rdpclip.exe'
    $r = Get-Reg $ifeo 'Debugger'
    if ($r.State -eq 'Set') { Add-F 'D. Security / Restrictions' 'IFEO Debugger hijack/block on rdpclip.exe' "$ifeo\Debugger" $r.Value 'Not present' 'FAIL' 'An Image File Execution Options Debugger value redirects/blocks rdpclip.exe launch (a known hardening technique for blocking executables).' 'Remove the IFEO Debugger value after confirming its owner (change-controlled).' 'Confirmed'; [void]$SecIssues.Add('IFEO on rdpclip') }
    elseif ($r.State -eq 'Error') { Add-F 'D. Security / Restrictions' 'IFEO Debugger on rdpclip.exe' "$ifeo\Debugger" (Show-Reg $r) 'Not present' 'UNKNOWN' 'Could not read.' }
    else { Add-F 'D. Security / Restrictions' 'IFEO Debugger on rdpclip.exe' "$ifeo\Debugger" (Show-Reg $r) 'Not present' 'PASS' 'No IFEO redirection of rdpclip.exe.' }

    $srp = 'HKLM:\SOFTWARE\Policies\Microsoft\Windows\Safer\CodeIdentifiers'
    $r = Get-Reg $srp 'DefaultLevel'
    if ($r.State -eq 'KeyMissing') { Add-F 'D. Security / Restrictions' 'Software Restriction Policies' $srp '(not configured)' 'Not configured, or rdpclip allowed' 'PASS' 'SRP not configured.' }
    elseif ($r.State -eq 'Error') { Add-F 'D. Security / Restrictions' 'Software Restriction Policies' $srp (Show-Reg $r) 'n/a' 'UNKNOWN' 'Could not read.' }
    else {
        # DefaultLevel 262144 (0x40000) = Unrestricted; disallow rules live under level subkey 0
        $dis = @('Paths','Hashes','UrlZones') | Where-Object { Test-Path -LiteralPath "$srp\0\$_" }
        if (($r.State -eq 'Set' -and [int]$r.Value -ne 262144) -or $dis) {
            Add-F 'D. Security / Restrictions' 'Software Restriction Policies' "$srp (DefaultLevel, 0\*)" "DefaultLevel=$(Show-Reg $r); Disallowed rule types: $(if ($dis) { $dis -join ',' } else { 'none' })" 'Unrestricted / rdpclip.exe permitted' 'WARNING' 'SRP is restrictive (non-Unrestricted default level or Disallowed rules exist). Verify %SystemRoot%\System32\rdpclip.exe is not blocked.' 'Review secpol.msc > Software Restriction Policies.' 'Potential'; [void]$SecIssues.Add('SRP restrictive') }
        else { Add-F 'D. Security / Restrictions' 'Software Restriction Policies' "$srp (DefaultLevel, 0\*)" "DefaultLevel=$(Show-Reg $r); no Disallowed rules" 'Unrestricted / rdpclip.exe permitted' 'PASS' 'SRP key exists but enforces no restriction.' } }

    try {
        [xml]$ap = Get-AppLockerPolicy -Effective -Xml -ErrorAction Stop
        $exeRc = @($ap.AppLockerPolicy.RuleCollection | Where-Object { $_.Type -eq 'Exe' })
        $mode = if ($exeRc) { [string]$exeRc[0].EnforcementMode } else { 'NotConfigured' }
        $appid = (Get-CimInstance Win32_Service -Filter "Name='AppIDSvc'").State
        if ($mode -eq 'Enabled') { Add-F 'D. Security / Restrictions' 'AppLocker Exe rules' 'Get-AppLockerPolicy -Effective' "EnforcementMode=Enabled; AppIDSvc=$appid" 'rdpclip.exe allowed' 'WARNING' 'AppLocker executable rules are enforced. Confirm a rule allows %SYSTEM32%\rdpclip.exe (default rules do). Block events are checked below.' 'Review AppLocker rules (secpol.msc).' 'Potential'; [void]$SecIssues.Add('AppLocker enforced') }
        else { Add-F 'D. Security / Restrictions' 'AppLocker Exe rules' 'Get-AppLockerPolicy -Effective' "EnforcementMode=$mode; AppIDSvc=$appid" 'Not enforced, or rdpclip allowed' 'PASS' 'AppLocker executable rules are not enforced.' }
    } catch { Add-F 'D. Security / Restrictions' 'AppLocker Exe rules' 'Get-AppLockerPolicy -Effective' ('ERROR: ' + $_.Exception.Message) 'Readable' 'UNKNOWN' 'Could not read effective AppLocker policy.' }

    foreach ($chk in @(@{ L='Microsoft-Windows-AppLocker/EXE and DLL'; I=@(8004); N='AppLocker block events for rdpclip.exe' }, @{ L='Microsoft-Windows-CodeIntegrity/Operational'; I=@(3077,3033); N='WDAC / Code Integrity block events for rdpclip.exe' })) {
        $ev = Get-Ev $chk.L $chk.I 'rdpclip'
        if ($ev.State -eq 'NoLog') { Add-F 'D. Security / Restrictions' $chk.N "$($chk.L) IDs $($chk.I -join ',')" 'Log not present' 'No events' 'NOT APPLICABLE' 'Log not present on this OS.' }
        elseif ($ev.State -ne 'OK') { Add-F 'D. Security / Restrictions' $chk.N "$($chk.L) IDs $($chk.I -join ',')" "ERROR: $($ev.Err)" 'No events' 'UNKNOWN' 'Log could not be read (access denied?).' }
        elseif ($ev.Events.Count) { $last = $ev.Events | Sort-Object TimeCreated -Descending | Select-Object -First 1
            Add-F 'D. Security / Restrictions' $chk.N "$($chk.L) IDs $($chk.I -join ',')" "$($ev.Events.Count) event(s); latest $($last.TimeCreated): $(Short $last.Message)" 'No events' 'FAIL' 'rdpclip.exe was blocked from running by application control - clipboard redirection cannot start.' 'Add an allow rule for %SystemRoot%\System32\rdpclip.exe (security team, change-controlled).' 'Confirmed'; [void]$SecIssues.Add($chk.N) }
        else { Add-F 'D. Security / Restrictions' $chk.N "$($chk.L) IDs $($chk.I -join ',')" "0 in last $Lookback days" 'No events' 'PASS' 'No block events for rdpclip.exe.' }
    }

    $sense = Get-CimInstance Win32_Service -Filter "Name='Sense'"
    if ($sense -and $sense.State -eq 'Running') { Add-F 'D. Security / Restrictions' 'Endpoint DLP indicator (Defender for Endpoint "Sense")' 'Win32_Service Sense' 'Running' 'n/a' 'UNKNOWN' 'Device is onboarded to Microsoft Defender for Endpoint. If Microsoft Purview Endpoint DLP is licensed and scoped to this device, it can restrict copy-to-clipboard / copy-to-RDP for sensitive content. This cannot be verified from the guest.' 'Ask the security team whether Purview Endpoint DLP policies target this server or the technician laptop.'; [void]$SecIssues.Add('MDE present (DLP unverified)') }
    else { Add-F 'D. Security / Restrictions' 'Endpoint DLP indicator (Defender for Endpoint "Sense")' 'Win32_Service Sense' $(if ($sense) { $sense.State } else { 'Not installed' }) 'n/a' 'NOT APPLICABLE' 'MDE sensor not running; Purview Endpoint DLP is not active on this server. (Third-party DLP cannot be detected generically.)' }

    try {
        $fr = @(Get-NetFirewallRule -Name 'RemoteDesktop-UserMode-In-TCP','RemoteDesktop-UserMode-In-UDP' -ErrorAction Stop | ForEach-Object { "$($_.Name)=$($_.Enabled)" }) -join '; '
        Add-F 'D. Security / Restrictions' 'Windows Firewall RDP inbound rules' 'Get-NetFirewallRule RemoteDesktop-UserMode-In-*' $fr 'n/a' 'NOT APPLICABLE' 'Informational only. Clipboard and drive redirection travel as virtual channels INSIDE the already-established RDP connection; Windows Firewall cannot block them separately. If the session connects, the firewall is not the clipboard cause.'
    } catch { Add-F 'D. Security / Restrictions' 'Windows Firewall RDP inbound rules' 'Get-NetFirewallRule' 'Not queryable' 'n/a' 'NOT APPLICABLE' 'Informational only; firewall does not filter RDP virtual channels.' }

    $S.Security = if ($SecIssues.Count) { $SecIssues -join '; ' } else { 'No restriction detected' }
}

# ---------------------------------------------------------------- B. Event logs
Invoke-Section 'B. Event Logs' {
    $chk = @(
        @{ N='rdpclip.exe crashes (Application Error/WER)'; L='Application'; I=@(1000,1001); M='rdpclip'; Lv=$null; Rc='Potential'; X='rdpclip.exe crashed - clipboard stops working in that session until rdpclip restarts.' },
        @{ N='RDS service failures (Service Control Manager)'; L='System'; I=@(7000,7001,7023,7024,7031,7034); M='Remote Desktop|TermService|UmRdpService'; Lv=$null; Rc='Potential'; X='RDS service start failures / unexpected terminations.' },
        @{ N='RdpCoreTS virtual-channel errors'; L='Microsoft-Windows-RemoteDesktopServices-RdpCoreTS/Operational'; I=$null; M='channel|clip|rdpdr'; Lv=2; Rc='Potential'; X='Errors in the RDP core stack mentioning virtual channels (clipboard = CLIPRDR, drives = RDPDR).' }
    )
    foreach ($c in $chk) {
        $ev = Get-Ev $c.L $c.I $c.M $c.Lv
        $src = "$($c.L)$(if ($c.I) { ' IDs ' + ($c.I -join ',') })$(if ($c.Lv) { ' Level=Error' }) match /$($c.M)/"
        if ($ev.State -eq 'NoLog') { Add-F 'B. Event Logs' $c.N $src 'Log not present' 'No events' 'NOT APPLICABLE' 'Log not present on this OS.' }
        elseif ($ev.State -ne 'OK') { Add-F 'B. Event Logs' $c.N $src "ERROR: $($ev.Err)" 'No events' 'UNKNOWN' 'Log could not be read.' }
        elseif ($ev.Events.Count) { $last = $ev.Events | Sort-Object TimeCreated -Descending | Select-Object -First 1
            Add-F 'B. Event Logs' $c.N $src "$($ev.Events.Count) event(s) in $Lookback d; latest $($last.TimeCreated) ID $($last.Id): $(Short $last.Message)" 'No events' 'WARNING' $c.X 'Correlate timestamps with the failed BeyondTrust session.' $c.Rc }
        else { Add-F 'B. Event Logs' $c.N $src "0 in last $Lookback days" 'No events' 'PASS' 'No relevant errors found.' }
    }
}

# ---------------------------------------------------------------- Items not verifiable from the guest
Add-F 'C. Live-session only' 'Clipboard virtual channel (CLIPRDR) requested by BeyondTrust / Jumpoint RDP client' 'Not observable via VMware Tools guest operations' 'Not verifiable' 'Requested/enabled' 'UNKNOWN' 'Channel negotiation happens between the RDP client (BeyondTrust Jumpoint/console) and the server at connect time. A guest-side script in session 0 cannot observe what the client requested.' 'Live test: compare direct mstsc vs BeyondTrust session (see guide).'
Add-F 'C. Live-session only' 'Drive redirection (RDPDR) offered by client (\\tsclient visible in session)' 'Not observable from session 0' 'Not verifiable' 'Per approved design' 'UNKNOWN' '\\tsclient shares exist only inside the user''s session namespace.' 'In the session, open \\tsclient in Explorer during a live test.'

$out = [ordered]@{ m=$M; s=$S; f=$F }
'###RDPJSON###' + ($out | ConvertTo-Json -Depth 6 -Compress) + '###END###'
'@
$GuestPayload = $GuestPayload.Replace('__LOOKBACK__', [string]$EventLookbackDays)
#endregion

#region ------------------------------------------------------------------ VM selection
$Assess  = New-Object System.Collections.ArrayList   # per-VM summary rows
$Detail  = New-Object System.Collections.ArrayList   # per-check rows
$Skipped = New-Object System.Collections.ArrayList   # errors / skipped
$Seen    = @{}

function Add-Skip { param($vc, $name, $reason, $next, $kind)
    if ($kind -ne 'None') { Write-Log "SKIPPED $name ($kind): $reason" 'WARN' }
    [void]$Skipped.Add([pscustomobject]@{ vCenter = $vc; 'VM Name' = $name; Category = $kind; 'Error / Reason' = $reason; 'Recommended Next Step' = $next }) }
function Get-VCName($vm) { if ($vm.Uid -match '@([^:/]+)') { $Matches[1] } else { '?' } }

$targets = @()
if ($AllWindowsVMs) {
    Write-Log 'Enumerating all VMs on connected vCenters...'
    $targets = @(Get-VM -Server $viServers)
} else {
    $names = if ($VMListFile) { @(Get-Content $VMListFile | ForEach-Object { $_.Trim() } | Where-Object { $_ -and $_ -notmatch '^#' }) } else { $VMName }
    foreach ($n in $names) {
        $found = @(Get-VM -Name $n -Server $viServers -ErrorAction SilentlyContinue)
        if ($found.Count -eq 0) { Add-Skip '-' $n 'Not found in any connected vCenter' 'Check the VM name / vCenter connection.' 'Not found' }
        $targets += $found
    }
}

$eligible = @()
foreach ($vm in $targets) {
    $v = $vm.ExtensionData
    $key = if ($v.Config.InstanceUuid) { $v.Config.InstanceUuid } else { $vm.Uid }
    if ($Seen.ContainsKey($key)) { continue }; $Seen[$key] = $true
    $vc = Get-VCName $vm
    $gfull = if ($v.Guest.GuestFullName) { $v.Guest.GuestFullName } else { $v.Config.GuestFullName }
    if ($v.Config.Template) { Add-Skip $vc $vm.Name 'Template' 'Out of scope.' 'Excluded'; continue }
    if ($v.Config.ManagedBy -or $vm.Name -match '^vCLS') { Add-Skip $vc $vm.Name "Appliance / solution-managed VM ($($v.Config.ManagedBy.ExtensionKey))" 'Out of scope.' 'Excluded'; continue }
    if ($v.Config.GuestId -notlike 'windows*' -and $v.Guest.GuestFamily -ne 'windowsGuest') { Add-Skip $vc $vm.Name "Non-Windows guest ($gfull)" 'Out of scope.' 'Excluded'; continue }
    if ($v.Config.VAppConfig -and $v.Config.VAppConfig.Product) { Add-Skip $vc $vm.Name 'vApp/OVF appliance' 'Out of scope.' 'Excluded'; continue }
    if (-not $IncludeWindowsClient -and $gfull -notmatch 'Server') { Add-Skip $vc $vm.Name "Windows client OS ($gfull)" 'Use -IncludeWindowsClient to include.' 'Excluded'; continue }
    if ($vm.PowerState -ne 'PoweredOn') { Add-Skip $vc $vm.Name "Power state: $($vm.PowerState)" 'Power on (if appropriate) and re-run.' 'Not assessed'; continue }
    if ($v.Guest.ToolsRunningStatus -ne 'guestToolsRunning') { Add-Skip $vc $vm.Name "VMware Tools not running ($($v.Guest.ToolsRunningStatus), $($v.Guest.ToolsVersionStatus2))" 'Start/repair VMware Tools in the guest (change-controlled).' 'Not assessed'; continue }
    if (-not $v.Guest.GuestOperationsReady) { Add-Skip $vc $vm.Name 'VMware Tools guest operations not ready' 'Check VMware Tools health / VGAuth service in guest.' 'Not assessed'; continue }
    $eligible += $vm
}
Write-Log ("Eligible Windows Server VMs: {0}; skipped/excluded: {1}" -f $eligible.Count, $Skipped.Count)
#endregion

#region ------------------------------------------------------------------ Assessment loop
$i = 0
foreach ($vm in $eligible) {
    $i++
    $v = $vm.ExtensionData; $vc = Get-VCName $vm
    Write-Progress -Activity 'RDP clipboard assessment' -Status "$($vm.Name) ($i/$($eligible.Count))" -PercentComplete ($i / [math]::Max(1, $eligible.Count) * 100)
    Write-Log "[$i/$($eligible.Count)] $($vm.Name) @ $vc"

    $ip = (@($v.Guest.Net | ForEach-Object { $_.IpAddress }) | Where-Object { $_ -match '^\d{1,3}(\.\d{1,3}){3}$' -and $_ -notlike '169.254.*' }) -join ', '
    if (-not $ip) { $ip = $v.Guest.IpAddress }
    $toolsTxt = "$($v.Guest.ToolsRunningStatus) / $($v.Guest.ToolsVersionStatus2) (build $($v.Guest.ToolsVersion))"
    $guestOS = if ($v.Guest.GuestFullName) { $v.Guest.GuestFullName } else { $v.Config.GuestFullName }

    # --- E. VMware rows (laptop side, read-only)
    $vmRows = @()
    $vmRows += [pscustomobject]@{ c = 'E. VMware'; s = 'VM power / guest OS'; p = 'Get-VM'; a = "$($vm.PowerState); $guestOS"; e = 'PoweredOn; Windows Server'; st = 'PASS'; x = 'In scope.'; r = ''; rc = '' }
    $tst = if ($v.Guest.ToolsVersionStatus2 -in 'guestToolsCurrent', 'guestToolsSupportedNew', 'guestToolsUnmanaged') { 'PASS' } else { 'WARNING' }
    $vmRows += [pscustomobject]@{ c = 'E. VMware'; s = 'VMware Tools status / version'; p = 'ExtensionData.Guest.ToolsRunningStatus, ToolsVersionStatus2, ToolsVersion'; a = $toolsTxt; e = 'guestToolsRunning / current'; st = $tst; x = 'VMware Tools is only the transport for THIS assessment. RDP clipboard redirection does not use VMware Tools.'; r = $(if ($tst -eq 'WARNING') { 'Plan a VMware Tools upgrade (not related to the RDP clipboard issue).' } else { '' }); rc = '' }
    $xc = @{}; foreach ($o in @($v.Config.ExtraConfig)) { if ($o.Key -match '^isolation\.tools\.(copy|paste)\.disable$|^isolation\.tools\.setGUIOptions\.enable$') { $xc[$o.Key] = $o.Value } }
    $xcTxt = if ($xc.Count) { ($xc.GetEnumerator() | ForEach-Object { "$($_.Key)=$($_.Value)" }) -join '; ' } else { 'not set (vSphere default: console copy/paste disabled)' }
    $vmRows += [pscustomobject]@{ c = 'E. VMware'; s = 'VM console (VMRC/web console) copy/paste isolation'; p = 'ExtraConfig isolation.tools.copy.disable / paste.disable / setGUIOptions.enable'; a = $xcTxt; e = 'n/a for RDP'; st = 'NOT APPLICABLE'; x = 'These settings only govern the VMware remote console clipboard (VMRC/HTML5 console via VMware Tools). BeyondTrust RDP Jump Items use the RDP protocol to the guest OS, so these values do NOT affect RDP clipboard redirection. Do not change them to fix this issue.'; r = ''; rc = '' }

    $row = [ordered]@{ vCenter = $vc; 'VM Name' = $vm.Name; 'IP Address' = $ip; OS = $guestOS; 'VMware Tools Status' = $toolsTxt
        'RDP Service Status' = ''; 'RDP Clipboard Policy' = ''; 'RDP Drive Redirection Policy' = ''; 'Clipboard Process Status' = ''
        'RDP Session Status' = ''; 'Security/Policy Restrictions' = ''; 'Overall Assessment' = ''; 'Potential Root Cause' = ''; 'Recommended Action' = ''; 'PAM Investigation Required' = '' }

    try {
        $res = Invoke-VMScript -VM $vm -ScriptText $GuestPayload -ScriptType Powershell -GuestCredential $GuestCred -ErrorAction Stop -WarningAction SilentlyContinue
        $txt = [string]$res.ScriptOutput
        if ($txt -notmatch '###RDPJSON###(.*)###END###') { throw "Guest script returned no parsable result (exit code $($res.ExitCode)). Output: $($o = $txt -replace '\s+', ' '; $o.Substring(0, [math]::Min(300, $o.Length)))" }
        $data = $Matches[1] | ConvertFrom-Json
    } catch {
        $msg = $_.Exception.Message
        $next = if ($msg -match 'InvalidGuestLogin|incorrect user name or password|Failed to authenticate') { 'Guest credential rejected - verify the account in wincred.xml exists on this VM.' }
                elseif ($msg -match 'GuestPermissionDenied|NoPermission') { 'vCenter account lacks Guest Operations privileges (VirtualMachine.GuestOperations.*).' }
                elseif ($msg -match 'not digitally signed|execution polic') { 'Guest PowerShell execution policy blocked the script; review policy (do not change without approval).' }
                else { 'Check VMware Tools / guest operations health and retry.' }
        Write-Log "$($vm.Name): $msg" 'ERR'
        Add-Skip $vc $vm.Name $msg $next 'Inaccessible'
        $row.'Overall Assessment' = 'INACCESSIBLE'; $row.'Potential Root Cause' = 'Not assessed'; $row.'Recommended Action' = $next; $row.'PAM Investigation Required' = 'n/a'
        [void]$Assess.Add([pscustomobject]$row)
        foreach ($f in $vmRows) { [void]$Detail.Add([pscustomobject]@{ vCenter = $vc; 'VM Name' = $vm.Name; 'Check Category' = $f.c; 'Setting / Dependency' = $f.s; 'Registry Path / Service / Command' = $f.p; 'Actual Value' = $f.a; 'Expected Value' = $f.e; Status = $f.st; 'Technical Explanation' = $f.x; 'Recommended Corrective Action' = $f.r; 'Root-Cause Class' = $f.rc }) }
        continue
    }

    $all = @($vmRows) + @($data.f)
    foreach ($f in $all) {
        [void]$Detail.Add([pscustomobject]@{ vCenter = $vc; 'VM Name' = $vm.Name; 'Check Category' = $f.c; 'Setting / Dependency' = $f.s; 'Registry Path / Service / Command' = $f.p
            'Actual Value' = $f.a; 'Expected Value' = $f.e; Status = $f.st; 'Technical Explanation' = $f.x; 'Recommended Corrective Action' = $f.r; 'Root-Cause Class' = $f.rc })
    }

    $fails = @($all | Where-Object { $_.st -eq 'FAIL' })
    $warns = @($all | Where-Object { $_.st -eq 'WARNING' })
    $unk   = @($all | Where-Object { $_.st -eq 'UNKNOWN' -and $_.c -ne 'C. Live-session only' -and $_.s -notlike 'Endpoint DLP*' })
    $confClip  = @($fails | Where-Object { $_.rc -eq 'Confirmed' })
    $confFiles = @($fails | Where-Object { $_.rc -eq 'Confirmed-Files' })

    $row.'RDP Service Status' = $data.s.RdpService
    $row.'RDP Clipboard Policy' = $data.s.ClipPolicy
    $row.'RDP Drive Redirection Policy' = $data.s.DrivePolicy
    $row.'Clipboard Process Status' = $data.s.ClipProc
    $row.'RDP Session Status' = $data.s.Session
    $row.'Security/Policy Restrictions' = $data.s.Security
    $row.'Overall Assessment' = if ($fails.Count) { 'FAIL' } elseif ($warns.Count) { 'WARNING' } elseif ($unk.Count) { 'PASS (Windows-side) - with UNKNOWN checks' } else { 'PASS (Windows-side) - PAM/live test still required' }

    $rcTxt = @()
    if ($confClip.Count)  { $rcTxt += 'CONFIRMED (text+file clipboard): ' + (($confClip | ForEach-Object s) -join '; ') }
    if ($confFiles.Count) { $rcTxt += 'CONFIRMED (file transfer/drive redirection): ' + (($confFiles | ForEach-Object s) -join '; ') }
    $other = @(($fails + $warns) | Where-Object { $_.rc -like 'Potential*' -or ($_.st -eq 'FAIL' -and -not $_.rc) })
    if ($other.Count) { $rcTxt += 'POTENTIAL: ' + (($other | ForEach-Object s) -join '; ') }
    if (-not $rcTxt) { $rcTxt += 'No Windows-side blocker detected. Investigate BeyondTrust session policy / Jump Item / client (see PAM Verification sheet).' }
    $row.'Potential Root Cause' = $rcTxt -join ' || '
    $acts = @(($fails + $warns) | Where-Object { $_.r } | ForEach-Object r | Select-Object -Unique)
    $row.'Recommended Action' = if ($acts) { $acts -join ' || ' } else { 'Windows: none. PAM team: complete PAM Verification sheet; run live session test while re-running this script.' }
    $row.'PAM Investigation Required' = if ($confClip.Count) { 'After Windows fix' } else { 'YES' }
    [void]$Assess.Add([pscustomobject]$row)
    Write-Log ("{0}: {1} (FAIL {2}, WARN {3})" -f $vm.Name, $row.'Overall Assessment', $fails.Count, $warns.Count) $(if ($fails.Count) { 'WARN' } else { 'OK' })
}
Write-Progress -Activity 'RDP clipboard assessment' -Completed
#endregion

#region ------------------------------------------------------------------ PAM verification (static, cannot be checked from the VM)
$Pam = @(
  [pscustomobject]@{ 'Verification Owner'='PAM console (/login)'; Setting='Effective session permission: clipboard (text) for RDP sessions'; 'Why It Matters'='BeyondTrust brokers the RDP connection; if its effective permission does not allow clipboard, the client never offers the clipboard channel even when Windows allows it.'; 'What PAM Must Verify'='The EFFECTIVE permission for the affected user on the affected Jump Item - user account permissions, group policy, Jump Policy and any Session Policy assigned to the Jump Item / Jump Group (a more specific policy can override "allowed" at another level). Note both directions if the product version exposes clipboard direction.'; 'Expected Configuration'='Clipboard allowed in both required directions for this user + Jump Item.'; 'Evidence Required'='Screenshots of user/group policy and Session Policy pages showing the clipboard setting, plus which policy wins for this Jump Item; PRA appliance and Access Console version numbers.' }
  [pscustomobject]@{ 'Verification Owner'='PAM console (/login)'; Setting='File transfer permission (separate from clipboard)'; 'Why It Matters'='Copying files is a different capability from copying text. Allowing clipboard does not imply file transfer is allowed.'; 'What PAM Must Verify'='Whether file transfer is permitted for RDP sessions for this user/Jump Item, in which direction(s), and any path restrictions; whether the RDP Jump Item type supports file copy/paste in your PRA version or only BeyondTrust''s own file-transfer feature.'; 'Expected Configuration'='Per approved design: file transfer allowed (upload/download) if required; otherwise document that files are intentionally blocked.'; 'Evidence Required'='Screenshot of file-transfer permission in effective policy; vendor release notes for RDP file-transfer support in the deployed version.' }
  [pscustomobject]@{ 'Verification Owner'='PAM console (/login)'; Setting='RDP Jump Item configuration'; 'Why It Matters'='The Jump Item defines how the Jumpoint launches the RDP session (session policy assignment, Jumpoint used, console/admin session, quality options, external vs built-in RDP client where supported).'; 'What PAM Must Verify'='Jump Item for each affected VM: assigned Jumpoint, assigned Session Policy, Jump Group, and any RDP-specific options related to clipboard/drive redirection in your version.'; 'Expected Configuration'='Jump Item uses the intended Session Policy and Jumpoint; no option that disables redirection.'; 'Evidence Required'='Export/screenshot of the Jump Item properties.' }
  [pscustomobject]@{ 'Verification Owner'='PAM console (/login)'; Setting='Jumpoint health and version'; 'Why It Matters'='RDP sessions are proxied by the Jumpoint; an outdated or mismatched Jumpoint can lack features or exhibit channel bugs.'; 'What PAM Must Verify'='Jumpoint online status, version matches appliance, host OS of the Jumpoint, any recent errors.'; 'Expected Configuration'='Jumpoint online, same version as appliance.'; 'Evidence Required'='Jumpoint status page screenshot / version.' }
  [pscustomobject]@{ 'Verification Owner'='PAM console (/login)'; Setting='Session recording / session reports for the failed attempt'; 'Why It Matters'='Session reports show which session policy was applied and events (clipboard / file-transfer attempts, denials).'; 'What PAM Must Verify'='Session report for a failing session: applied policy, user, Jump Item, Jumpoint, any clipboard/file-transfer events.'; 'Expected Configuration'='Applied policy matches expectation; no denial events.'; 'Evidence Required'='Session report export (with session ID and timestamp) to correlate with the Windows source address in this workbook.' }
  [pscustomobject]@{ 'Verification Owner'='PAM console (/login)'; Setting='Access request / approval and time-based policy conditions'; 'Why It Matters'='Some policies grant different permissions depending on approval, schedule or location.'; 'What PAM Must Verify'='Whether the user''s session was granted under a restricted policy variant (e.g. no-approval, outside schedule).'; 'Expected Configuration'='Same permissions for the session as tested.'; 'Evidence Required'='Policy condition screenshots.' }
  [pscustomobject]@{ 'Verification Owner'='Technician laptop'; Setting='Access Console version / clipboard behaviour'; 'Why It Matters'='The BeyondTrust Access Console on the laptop is the clipboard endpoint on the technician side.'; 'What PAM Must Verify'='Access Console version (must match appliance), whether the console exposes a clipboard/send-clipboard control that must be used, and whether it is running elevated vs the app copying data (UIPI can block clipboard between integrity levels).'; 'Expected Configuration'='Current console; console and source app at same integrity level.'; 'Evidence Required'='Console version screenshot; test copying from Notepad (non-elevated).' }
  [pscustomobject]@{ 'Verification Owner'='Technician laptop'; Setting='Laptop clipboard / DLP controls'; 'Why It Matters'='Endpoint DLP (e.g. Microsoft Purview Endpoint DLP or third-party agents) or clipboard-manager tools on the laptop can block or intercept copy to remote sessions.'; 'What PAM Must Verify'='With the endpoint/security team: DLP agent and policies targeting the laptop; clipboard managers.'; 'Expected Configuration'='No policy blocking copy to BeyondTrust console.'; 'Evidence Required'='DLP policy confirmation; DLP agent event logs at test time.' }
  [pscustomobject]@{ 'Verification Owner'='Live session test'; Setting='Direct RDP vs BeyondTrust comparison'; 'Why It Matters'='Isolates Windows-side from BeyondTrust-side causes.'; 'What PAM Must Verify'='From a permitted admin host, mstsc to the VM with Local Resources > Clipboard (and Drives, if needed) ticked, test text and file copy both ways; repeat via BeyondTrust.'; 'Expected Configuration'='Both work, or only BeyondTrust fails (=> PAM-side).'; 'Evidence Required'='Test matrix result with timestamps; re-run this script during each live session to capture rdpclip state and source address.' }
  [pscustomobject]@{ 'Verification Owner'='Live session test'; Setting='rdpclip.exe in the BeyondTrust session'; 'Why It Matters'='Shows whether the server-side clipboard component runs in the PAM-brokered session.'; 'What PAM Must Verify'='Keep a BeyondTrust session open and re-run this script for that VM; check "Clipboard Process Status".'; 'Expected Configuration'='rdpclip running in the session whose source address is the Jumpoint.'; 'Evidence Required'='Workbook from the live run.' }
  [pscustomobject]@{ 'Verification Owner'='Live session test'; Setting='Text vs file copy behaviour'; 'Why It Matters'='Text clipboard, file clipboard and drive redirection are separate capabilities with separate controls.'; 'What PAM Must Verify'='Test separately: text laptop->VM, text VM->laptop, file laptop->VM, file VM->laptop, \\tsclient visibility.'; 'Expected Configuration'='Per approved design.'; 'Evidence Required'='Result per test case.' }
)
#endregion

#region ------------------------------------------------------------------ Executive summary
$assessed = @($Assess | Where-Object { $_.'Overall Assessment' -ne 'INACCESSIBLE' })
$cnt = { param($p) @($Assess | Where-Object { $_.'Overall Assessment' -like $p }).Count }
$Summary = @(
  [pscustomobject]@{ Metric = 'Report generated'; Value = (Get-Date -Format 'yyyy-MM-dd HH:mm'); Notes = "vCenters: $(($viServers | ForEach-Object Name) -join ', ')" }
  [pscustomobject]@{ Metric = 'Total VMs assessed (guest script completed)'; Value = $assessed.Count; Notes = '' }
  [pscustomobject]@{ Metric = 'PASS (Windows-side)'; Value = (& $cnt 'PASS*'); Notes = 'Windows configuration shows no blocker. NOT proof that clipboard works - BeyondTrust/client must still be verified.' }
  [pscustomobject]@{ Metric = 'FAIL'; Value = (& $cnt 'FAIL'); Notes = 'At least one Windows-side setting fails.' }
  [pscustomobject]@{ Metric = 'WARNING'; Value = (& $cnt 'WARNING'); Notes = 'Potential contributors; no confirmed blocker.' }
  [pscustomobject]@{ Metric = 'Inaccessible'; Value = (& $cnt 'INACCESSIBLE'); Notes = 'Guest script could not run - see Errors and Skipped.' }
  [pscustomobject]@{ Metric = 'Not assessed (powered off / Tools not usable)'; Value = @($Skipped | Where-Object Category -eq 'Not assessed').Count; Notes = '' }
  [pscustomobject]@{ Metric = 'Excluded (non-Windows, client OS, template, appliance)'; Value = @($Skipped | Where-Object Category -eq 'Excluded').Count; Notes = '' }
  [pscustomobject]@{ Metric = 'Not found'; Value = @($Skipped | Where-Object Category -eq 'Not found').Count; Notes = '' }
  [pscustomobject]@{ Metric = 'VMs with CONFIRMED Windows configuration issue'; Value = @($assessed | Where-Object { $_.'Potential Root Cause' -like '*CONFIRMED*' }).Count; Notes = 'Owner: Windows/server team.' }
  [pscustomobject]@{ Metric = 'VMs requiring PAM-side investigation'; Value = @($assessed | Where-Object { $_.'PAM Investigation Required' -eq 'YES' }).Count; Notes = 'No confirmed Windows blocker for text clipboard - owner: BeyondTrust PAM team (see PAM Verification sheet).' }
  [pscustomobject]@{ Metric = 'Method / limitations'; Value = ''; Notes = "Read-only checks via VMware Tools Invoke-VMScript (runs in session 0 as the guest credential). Cannot observe RDP channel negotiation by BeyondTrust; rdpclip check is only meaningful while an RDP session is active. UNKNOWN = could not verify (never counted as PASS). Event lookback: $EventLookbackDays days." }
)
if ($Skipped.Count -eq 0) { Add-Skip '-' '-' 'No errors or skipped VMs' '-' 'None' }
if ($Assess.Count -eq 0) { [void]$Assess.Add([pscustomobject]@{ vCenter = '-'; 'VM Name' = 'No VM assessed'; 'Overall Assessment' = '-' }) }
if ($Detail.Count -eq 0) { [void]$Detail.Add([pscustomobject]@{ vCenter = '-'; 'VM Name' = 'No VM assessed'; Status = '-' }) }
#endregion

#region ------------------------------------------------------------------ Export
if ($HaveImportExcel) {
    Import-Module ImportExcel
    $ct = @(
        New-ConditionalText -Text 'FAIL'           -BackgroundColor LightPink           -ConditionalTextColor DarkRed
        New-ConditionalText -Text 'INACCESSIBLE'   -BackgroundColor LightPink           -ConditionalTextColor DarkRed
        New-ConditionalText -Text 'CONFIRMED'      -BackgroundColor LightPink           -ConditionalTextColor DarkRed
        New-ConditionalText -Text 'WARNING'        -BackgroundColor LightGoldenrodYellow -ConditionalTextColor DarkGoldenrod
        New-ConditionalText -Text 'UNKNOWN'        -BackgroundColor Gainsboro           -ConditionalTextColor DimGray
        New-ConditionalText -Text 'NOT CONFIGURED' -BackgroundColor AliceBlue           -ConditionalTextColor SteelBlue
        New-ConditionalText -Text 'NOT APPLICABLE' -BackgroundColor WhiteSmoke          -ConditionalTextColor Gray
        New-ConditionalText -Text 'PASS'           -BackgroundColor Honeydew            -ConditionalTextColor DarkGreen
    )
    $common = @{ Path = $XlFile; AutoSize = $true; FreezeTopRow = $true; AutoFilter = $true; BoldTopRow = $true }
    $Summary | Export-Excel @common -WorksheetName 'Executive Summary' -TableStyle Medium2 -TableName ExecSummary
    $Assess  | Export-Excel @common -WorksheetName 'VM Assessment' -TableStyle Medium2 -TableName VMAssessment -ConditionalText $ct
    $Detail  | Export-Excel @common -WorksheetName 'Technical Findings' -TableStyle Medium2 -TableName Findings -ConditionalText $ct
    $Pam     | Export-Excel @common -WorksheetName 'PAM Verification' -TableStyle Medium7 -TableName PamVerification
    $Skipped | Export-Excel @common -WorksheetName 'Errors and Skipped' -TableStyle Medium3 -TableName Skipped -ConditionalText $ct
    # Cap column widths and wrap long text for readability
    $pkg = Open-ExcelPackage -Path $XlFile
    foreach ($ws in $pkg.Workbook.Worksheets) {
        if (-not $ws.Dimension) { continue }
        for ($c = 1; $c -le $ws.Dimension.End.Column; $c++) { $col = $ws.Column($c); if ($col.Width -gt 55) { $col.Width = 55 }; $col.Style.WrapText = $true }
        $ws.Cells[$ws.Dimension.Address].Style.VerticalAlignment = [OfficeOpenXml.Style.ExcelVerticalAlignment]::Top
    }
    Close-ExcelPackage $pkg
    Write-Log "Excel report: $XlFile" 'OK'
} else {
    $base = Join-Path $OutputFolder "RDP_Clipboard_Assessment_$ts"
    $Summary | Export-Csv "${base}_1_ExecutiveSummary.csv" -NoTypeInformation
    $Assess  | Export-Csv "${base}_2_VMAssessment.csv" -NoTypeInformation
    $Detail  | Export-Csv "${base}_3_TechnicalFindings.csv" -NoTypeInformation
    $Pam     | Export-Csv "${base}_4_PAMVerification.csv" -NoTypeInformation
    $Skipped | Export-Csv "${base}_5_ErrorsSkipped.csv" -NoTypeInformation
    Write-Log "ImportExcel not installed - CSV reports written to ${base}_*.csv" 'WARN'
}
Write-Log "Log file: $LogFile"
#endregion
