<#
.SYNOPSIS
    Windows-side remediation for RDP clipboard / drive redirection blockers found by
    Invoke-RdpClipboardAssessment.ps1. DRY RUN by default; -Apply makes changes; -Rollback reverts.

.DESCRIPTION
    Runs from the management laptop over existing PowerCLI vCenter connections using
    Invoke-VMScript (VMware Tools) and C:\temp\wincred.xml. No WinRM / RDP / SSH.

    Default fix (always planned):
      * 'Do not allow Clipboard redirection' Enabled (fDisableClip=1) -> Disabled (0), in the
        Local GPO (Registry.pol) AND the live registry so it is not reverted at next policy refresh.
      * RDP listener fDisableClip=1 -> 0 (only when no computer policy overrides the listener).
    Opt-in fixes:
      -FixClipboardFormatPolicies  Remove 'Restrict clipboard transfer client->server / server->client'
                                   (CSClipLevel / SCClipLevel) -> Not Configured.
      -FixDriveRedirection         'Do not allow drive redirection' -> Disabled (0); listener fDisableCdm -> 0;
                                   UmRdpService Disabled -> Manual. Only if file transfer is approved.
      -StartRdpServices            Start TermService / UmRdpService if stopped (never changes a Disabled TermService).
      -RemoveRdpclipIfeo           Remove an IFEO 'Debugger' value that blocks rdpclip.exe.
    Never touched (report-only): AppLocker / WDAC / SRP, DLP, firewall, VMware settings, BeyondTrust.

    Safety:
      * Nothing changes without -Apply, and each VM asks for confirmation (use -Confirm:$false to skip).
      * Before any change: Registry.pol + gpt.ini copied and every registry/service change recorded to
        C:\Windows\Temp\RdpClipRemediation\<timestamp>\changes.json in the guest (used by -Rollback).
      * Registry.pol is only rewritten if the parser can round-trip the original file byte-for-byte.
      * No reboot, no service stop/restart, no session disconnect. New settings apply to NEW RDP
        sessions - users must sign out / reconnect.
      * If an MDM (Intune) RDS policy exists it is flagged: MDM may re-apply its own value.

.EXAMPLE
    .\Invoke-RdpClipboardRemediation.ps1 -VMName APP01                      # dry run (plan only)
.EXAMPLE
    .\Invoke-RdpClipboardRemediation.ps1 -VMName APP01,APP02 -Apply
.EXAMPLE
    .\Invoke-RdpClipboardRemediation.ps1 -VMListFile C:\temp\fail.txt -Apply -FixClipboardFormatPolicies -StartRdpServices
.EXAMPLE
    .\Invoke-RdpClipboardRemediation.ps1 -VMName APP01 -Rollback            # revert last applied run on APP01
#>
[CmdletBinding(SupportsShouldProcess, ConfirmImpact = 'High', DefaultParameterSetName = 'Named')]
param(
    [Parameter(ParameterSetName = 'Named', Mandatory, Position = 0)][string[]]$VMName,
    [Parameter(ParameterSetName = 'File', Mandatory)][string]$VMListFile,
    [switch]$Apply,
    [switch]$Rollback,
    [switch]$FixClipboardFormatPolicies,
    [switch]$FixDriveRedirection,
    [switch]$StartRdpServices,
    [switch]$RemoveRdpclipIfeo,
    [switch]$SkipGpUpdate,
    [string]$CredentialPath = 'C:\temp\wincred.xml',
    [string]$OutputFolder = 'C:\temp'
)

$ErrorActionPreference = 'Stop'
if ($Apply -and $Rollback) { throw 'Use either -Apply or -Rollback, not both.' }
$Mode = if ($Rollback) { 'Rollback' } elseif ($Apply) { 'Apply' } else { 'Plan' }
$ts = Get-Date -Format 'yyyyMMdd_HHmmss'
if (-not (Test-Path $OutputFolder)) { throw "Output folder '$OutputFolder' does not exist." }
$LogFile = Join-Path $OutputFolder "RDP_Clipboard_Remediation_${Mode}_$ts.log"
$XlFile  = Join-Path $OutputFolder "RDP_Clipboard_Remediation_${Mode}_$ts.xlsx"

function Write-Log {
    param([string]$Msg, [string]$Level = 'INFO')
    $line = '{0} [{1}] {2}' -f (Get-Date -Format 'yyyy-MM-dd HH:mm:ss'), $Level, $Msg
    Add-Content -Path $LogFile -Value $line
    Write-Host $line -ForegroundColor (@{ INFO = 'Gray'; OK = 'Green'; WARN = 'Yellow'; ERR = 'Red' }[$Level])
}

#region ------------------------------------------------------------------ Pre-flight
if (-not (Get-Command Invoke-VMScript -ErrorAction SilentlyContinue)) { Import-Module VMware.VimAutomation.Core -ErrorAction Stop }
$viServers = @($global:DefaultVIServers | Where-Object { $_.IsConnected })
if ($viServers.Count -eq 0) { throw 'No connected vCenter found in $global:DefaultVIServers. Connect-VIServer first.' }
if (-not (Test-Path $CredentialPath)) { throw "Credential file '$CredentialPath' not found." }
$GuestCred = Import-Clixml -Path $CredentialPath
if ($GuestCred -isnot [System.Management.Automation.PSCredential]) { throw "'$CredentialPath' does not contain a PSCredential." }
$opts = @()
if ($FixClipboardFormatPolicies) { $opts += 'Format' }
if ($FixDriveRedirection) { $opts += 'Drive' }
if ($StartRdpServices) { $opts += 'Services' }
if ($RemoveRdpclipIfeo) { $opts += 'Ifeo' }
if ($SkipGpUpdate) { $opts += 'NoGpUpdate' }
Write-Log ("Mode: {0}; options: {1}; vCenters: {2}" -f $Mode, $(if ($opts) { $opts -join ',' } else { 'default' }), (($viServers | ForEach-Object Name) -join ', '))
if ($Mode -eq 'Plan') { Write-Log 'DRY RUN - no changes will be made. Add -Apply to remediate.' 'WARN' }
#endregion

#region ------------------------------------------------------------------ Guest payload
$GuestPayload = @'
$ErrorActionPreference = 'Stop'
$ProgressPreference = 'SilentlyContinue'
$Mode = '__MODE__'
$Opt = @('__OPTS__' -split ',' | Where-Object { $_ })
$RunId = '__RUNID__'
$A = New-Object System.Collections.ArrayList
$Changes = New-Object System.Collections.ArrayList
function Add-A { param($i,$t,$c,$p,$r,$n='') [void]$A.Add([ordered]@{ i=$i; t=$t; c=[string]$c; p=[string]$p; r=$r; n=$n }) }
function Emit { '###REMJSON###' + ([ordered]@{ a=$A; bk=$script:Bk } | ConvertTo-Json -Depth 5 -Compress) + '###END###' }

$TSPk = 'SOFTWARE\Policies\Microsoft\Windows NT\Terminal Services'; $TSP = "HKLM:\$TSPk"
$WSRoot = 'HKLM:\SYSTEM\CurrentControlSet\Control\Terminal Server\WinStations'
$IfeoKey = 'HKLM:\SOFTWARE\Microsoft\Windows NT\CurrentVersion\Image File Execution Options\rdpclip.exe'
$GPDir = "$env:windir\System32\GroupPolicy"; $PolFile = "$GPDir\Machine\Registry.pol"; $GptIni = "$GPDir\gpt.ini"
$BkRoot = "$env:windir\Temp\RdpClipRemediation"; $script:Bk = ''

$idn = [Security.Principal.WindowsIdentity]::GetCurrent()
$Elev = (New-Object Security.Principal.WindowsPrincipal $idn).IsInRole([Security.Principal.WindowsBuiltInRole]::Administrator)
if (-not $Elev -and $Mode -ne 'Plan') { Add-A 'Pre-check' 'Elevation' 'Not elevated' 'Elevated admin' 'FAILED' 'Guest credential token is not elevated - no changes attempted. Use the built-in Administrator.'; Emit; return }

# ---------------- registry helpers (record old state for rollback)
function Get-RegState($Path, $Name) {
    if (-not (Test-Path -LiteralPath $Path)) { return @{ Exists=$false; Value=$null; Kind=$null } }
    $k = Get-Item -LiteralPath $Path; $v = $k.GetValue($Name, $null)
    if ($null -eq $v) { @{ Exists=$false; Value=$null; Kind=$null } } else { @{ Exists=$true; Value=$v; Kind=[string]$k.GetValueKind($Name) } } }
function Save-Reg($Path, $Name) { $o = Get-RegState $Path $Name; [void]$Changes.Add([ordered]@{ Type='Reg'; Path=$Path; Name=$Name; Existed=$o.Exists; Value=$o.Value; Kind=$o.Kind }) }
function Set-Dword($Path, $Name, [int]$Val) { Save-Reg $Path $Name; if (-not (Test-Path -LiteralPath $Path)) { New-Item -Path $Path -Force | Out-Null }; New-ItemProperty -LiteralPath $Path -Name $Name -Value $Val -PropertyType DWord -Force | Out-Null }
function Remove-Val($Path, $Name) { Save-Reg $Path $Name; Remove-ItemProperty -LiteralPath $Path -Name $Name -ErrorAction SilentlyContinue }
function Save-Changes { if ($script:Bk) { ConvertTo-Json -InputObject @($Changes) -Depth 4 | Set-Content -LiteralPath "$script:Bk\changes.json" -Encoding UTF8 } }

# ---------------- Registry.pol (MS-GPREG format) reader / writer
function Read-Pol([string]$Path) {
    $b = [IO.File]::ReadAllBytes($Path)
    if ($b.Length -lt 8 -or [Text.Encoding]::ASCII.GetString($b, 0, 4) -ne 'PReg') { throw 'Registry.pol header invalid' }
    $u = [Text.Encoding]::Unicode; $list = New-Object System.Collections.ArrayList; $i = 8
    while ($i -lt $b.Length) {
        if ([BitConverter]::ToUInt16($b, $i) -ne 0x5B) { throw "Registry.pol parse error ('[') at offset $i" }; $i += 2
        $st = $i; while ([BitConverter]::ToUInt16($b, $i) -ne 0) { $i += 2 }; $key = $u.GetString($b, $st, $i - $st); $i += 4
        $st = $i; while ([BitConverter]::ToUInt16($b, $i) -ne 0) { $i += 2 }; $val = $u.GetString($b, $st, $i - $st); $i += 4
        $type = [BitConverter]::ToUInt32($b, $i); $i += 6
        $size = [BitConverter]::ToUInt32($b, $i); $i += 6
        $data = New-Object byte[] $size; [Array]::Copy($b, $i, $data, 0, $size); $i += $size
        if ([BitConverter]::ToUInt16($b, $i) -ne 0x5D) { throw "Registry.pol parse error (']') at offset $i" }; $i += 2
        [void]$list.Add([pscustomobject]@{ Key=$key; Value=$val; Type=$type; Data=$data })
    }
    , $list }
function ConvertTo-PolBytes($List) {
    $ms = New-Object IO.MemoryStream; $w = New-Object IO.BinaryWriter($ms); $u = [Text.Encoding]::Unicode
    $w.Write([byte[]](0x50, 0x52, 0x65, 0x67)); $w.Write([uint32]1)
    foreach ($e in $List) {
        $w.Write($u.GetBytes('[' + $e.Key + [char]0 + ';' + $e.Value + [char]0 + ';'))
        $w.Write([uint32]$e.Type); $w.Write($u.GetBytes(';')); $w.Write([uint32]$e.Data.Length); $w.Write($u.GetBytes(';'))
        $w.Write([byte[]]$e.Data); $w.Write($u.GetBytes(']'))
    }
    $w.Flush(); , $ms.ToArray() }

$Pol = $null; $PolOk = $false; $PolDirty = $false; $PolNote = 'No Local GPO Registry.pol'
if (Test-Path -LiteralPath $PolFile) {
    try {
        $Pol = Read-Pol $PolFile
        $orig = [IO.File]::ReadAllBytes($PolFile); $rt = ConvertTo-PolBytes $Pol
        $PolOk = [System.Linq.Enumerable]::SequenceEqual([byte[]]$orig, [byte[]]$rt)
        $PolNote = if ($PolOk) { 'Registry.pol parsed and round-trip verified' } else { 'Registry.pol round-trip mismatch - LGPO will NOT be edited' }
    } catch { $PolNote = 'Registry.pol unreadable: ' + $_.Exception.Message }
}
function Get-PolEntry($Name) { if ($Pol) { @($Pol | Where-Object { $_.Key -eq $TSPk -and $_.Value -eq $Name }) } else { @() } }
function PolDword($e) { if ($e.Type -eq 4 -and $e.Data.Length -ge 4) { [BitConverter]::ToUInt32($e.Data, 0) } else { 'non-DWORD' } }
function Show-State($s) { if ($s.Exists) { [string]$s.Value } else { '(not set)' } }

$MdmNote = if (Test-Path 'HKLM:\SOFTWARE\Microsoft\PolicyManager\current\device\RemoteDesktopServices') { ' WARNING: MDM RemoteDesktopServices policy present - MDM may re-apply its value.' } else { '' }

# ---------------- Rollback mode
if ($Mode -eq 'Rollback') {
    $last = Get-ChildItem -LiteralPath $BkRoot -Directory -ErrorAction SilentlyContinue | Where-Object { $_.Name -notlike '*_rolledback' -and (Test-Path "$($_.FullName)\changes.json") } | Sort-Object Name -Descending | Select-Object -First 1
    if (-not $last) { Add-A 'Rollback' $BkRoot 'No backup found' 'n/a' 'SKIPPED' 'Nothing to roll back.'; Emit; return }
    $script:Bk = $last.FullName
    # PS 5.1 ConvertFrom-Json emits a JSON array as ONE object - unroll it
    $ch = @((Get-Content -LiteralPath "$script:Bk\changes.json" -Raw | ConvertFrom-Json) | ForEach-Object { $_ })
    [array]::Reverse($ch)
    foreach ($c in $ch) {
        try {
            if ($c.Type -eq 'Reg') {
                if ($c.Existed) { if (-not (Test-Path -LiteralPath $c.Path)) { New-Item -Path $c.Path -Force | Out-Null }
                    New-ItemProperty -LiteralPath $c.Path -Name $c.Name -Value $c.Value -PropertyType $c.Kind -Force | Out-Null
                    Add-A 'Rollback registry' "$($c.Path)\$($c.Name)" '' $c.Value 'RESTORED' }
                else { Remove-ItemProperty -LiteralPath $c.Path -Name $c.Name -ErrorAction SilentlyContinue; Add-A 'Rollback registry' "$($c.Path)\$($c.Name)" '' '(removed - did not exist before)' 'RESTORED' }
            } elseif ($c.Type -eq 'Service') {
                Set-Service -Name $c.Name -StartupType $c.StartMode
                Add-A 'Rollback service start type' $c.Name '' $c.StartMode 'RESTORED' 'Running state not changed (services are never stopped by this tool).'
            } elseif ($c.Type -eq 'Pol') {
                Copy-Item -LiteralPath "$script:Bk\Registry.pol" -Destination $PolFile -Force
                if (Test-Path "$script:Bk\gpt.ini") { Copy-Item -LiteralPath "$script:Bk\gpt.ini" -Destination $GptIni -Force }
                Add-A 'Rollback Local GPO' $PolFile '' 'Original Registry.pol restored' 'RESTORED'; $PolDirty = $true
            }
        } catch { Add-A 'Rollback' "$($c.Type) $($c.Path) $($c.Name)" '' '' 'FAILED' $_.Exception.Message }
    }
    if ($PolDirty -and $Opt -notcontains 'NoGpUpdate') { $g = & gpupdate.exe /target:computer /force 2>&1 | Out-String; Add-A 'gpupdate /target:computer /force' 'Local policy refresh' '' "exit $LASTEXITCODE" $(if ($LASTEXITCODE -eq 0) { 'APPLIED' } else { 'FAILED' }) (($g -replace '\s+', ' ').Trim()) }
    if (-not @($A | Where-Object { $_.r -eq 'FAILED' }).Count) { Rename-Item -LiteralPath $script:Bk -NewName ((Split-Path $script:Bk -Leaf) + '_rolledback') }
    Emit; return
}

# ---------------- Build the plan (Act = action name for Invoke-Act; $null = report only)
$Plan = New-Object System.Collections.ArrayList
function Add-P { param($Item,$Target,$Current,$New,$Act,$Arg,$Enabled=$true,$Need='',$Note='') [void]$Plan.Add([pscustomobject]@{ Item=$Item; Target=$Target; Current=$Current; New=$New; Act=$Act; Arg=$Arg; Enabled=$Enabled; Need=$Need; Note=$Note }) }
function Invoke-Act($p) {
    $g = $p.Arg
    switch ($p.Act) {
        'PolicyDword' {   # live registry + matching LGPO entries
            Set-Dword $TSP $g.Name $g.Val
            if ($script:PolOk) { foreach ($e in (Get-PolEntry $g.Name)) { $e.Type = 4; $e.Data = [BitConverter]::GetBytes([uint32]$g.Val); $script:PolDirty = $true } } }
        'PolicyRemove' {
            Remove-Val $TSP $g.Name
            if ($script:PolOk -and (Get-PolEntry $g.Name)) { $script:Pol = [System.Collections.ArrayList]@($script:Pol | Where-Object { -not ($_.Key -eq $TSPk -and $_.Value -eq $g.Name) }); $script:PolDirty = $true } }
        'Dword'        { Set-Dword $g.Path $g.Name $g.Val }
        'RemoveValue'  { Remove-Val $g.Path $g.Name }
        'ServiceManual' { [void]$Changes.Add([ordered]@{ Type='Service'; Name=$g.Name; StartMode='Disabled' }); Set-Service -Name $g.Name -StartupType Manual }
        'StartService' { Start-Service -Name $g.Name }
        default { throw "Unknown action $($p.Act)" }
    } }

# Policy-type fix: set to explicit Disabled (0) in live registry + LGPO entry (if LGPO holds it)
function Plan-PolicyDword { param($Name,$Label,$Bad,$Good,$Enabled,$Need)
    $live = Get-RegState $TSP $Name; $pe = Get-PolEntry $Name
    $polBad = @($pe | Where-Object { (PolDword $_) -eq $Bad }).Count -gt 0
    if (($live.Exists -and [int]$live.Value -eq $Bad) -or $polBad) {
        $cur = "live=$(Show-State $live); LGPO=$(if ($pe) { ($pe | ForEach-Object { PolDword $_ }) -join ',' } else { 'not in Registry.pol' })"
        $note = if ($pe -and -not $PolOk) { "LGPO holds this value but Registry.pol cannot be safely edited ($PolNote) - live value is fixed but LGPO will re-apply it; fix in gpedit.msc." } elseif (-not $pe -and $live.Exists) { 'Value is not from Local GPO (written directly, e.g. hardening script/baseline) - that tool may re-apply it.' } else { 'LGPO and live registry updated together.' }
        Add-P $Label "HKLM\$TSPk\$Name" $cur "$Good (Disabled = redirection allowed)" 'PolicyDword' @{ Name=$Name; Val=$Good } $Enabled $Need ($note + $MdmNote)
        return $true }
    return $false }
function Plan-Listener { param($Name,$Label,$Enabled,$Need,$PolicyWillBeSet)
    $polLive = Get-RegState $TSP $Name
    foreach ($l in @(Get-ChildItem -LiteralPath $WSRoot | Where-Object { $null -ne $_.GetValue('PortNumber', $null) })) {
        $st = Get-RegState $l.PSPath $Name
        if ($st.Exists -and [int]$st.Value -eq 1) {
            if ($PolicyWillBeSet -or $polLive.Exists) { Add-P $Label "$($l.Name)\$Name" 1 0 $null $null $false '' 'Not needed: computer policy is configured and overrides the listener value.'; continue }
            Add-P $Label "$($l.Name)\$Name" 1 0 'Dword' @{ Path=$l.PSPath; Name=$Name; Val=0 } $Enabled $Need
        } } }

# 1. Clipboard (default)
$clipPol = Plan-PolicyDword 'fDisableClip' "Policy 'Do not allow Clipboard redirection' -> Disabled" 1 0 $true ''
Plan-Listener 'fDisableClip' 'Listener clipboard mapping (fDisableClip) -> 0' $true '' $clipPol

# 2. Directional clipboard format restrictions (opt-in)
foreach ($n in 'CSClipLevel', 'SCClipLevel') {
    $live = Get-RegState $TSP $n; $pe = Get-PolEntry $n
    if ($live.Exists -or $pe) {
        $note = if ($pe -and -not $PolOk) { "Registry.pol cannot be safely edited ($PolNote) - use gpedit.msc." } else { '' }
        Add-P "Policy $n (clipboard format restriction) -> Not Configured" "HKLM\$TSPk\$n" "live=$(Show-State $live); LGPO=$(if ($pe) { 'configured' } else { 'not in Registry.pol' })" '(removed)' 'PolicyRemove' @{ Name=$n } ($Opt -contains 'Format') '-FixClipboardFormatPolicies' ($note + $MdmNote)
    } }

# 3. Drive redirection (opt-in)
$drv = $Opt -contains 'Drive'
$cdmPol = Plan-PolicyDword 'fDisableCdm' "Policy 'Do not allow drive redirection' -> Disabled" 1 0 $drv '-FixDriveRedirection'
Plan-Listener 'fDisableCdm' 'Listener drive mapping (fDisableCdm) -> 0' $drv '-FixDriveRedirection' $cdmPol
$um = Get-CimInstance Win32_Service -Filter "Name='UmRdpService'"
if ($um -and $um.StartMode -eq 'Disabled') { Add-P 'UmRdpService start type Disabled -> Manual' 'UmRdpService' 'Disabled' 'Manual' 'ServiceManual' @{ Name='UmRdpService' } $drv '-FixDriveRedirection' }

# 4. Services (opt-in) - start only, never stop/restart
foreach ($n in 'TermService', 'UmRdpService') {
    $sv = Get-CimInstance Win32_Service -Filter "Name='$n'"
    if ($sv -and $sv.State -ne 'Running') {
        if ($sv.StartMode -eq 'Disabled' -and $n -eq 'TermService') { Add-P "Start $n" $n "State=$($sv.State); StartMode=Disabled" 'Running' $null $null $false '' 'TermService is Disabled - not changed automatically (RDP may be intentionally disabled). Decide manually.'; continue }
        Add-P "Start $n" $n "State=$($sv.State); StartMode=$($sv.StartMode)" 'Running' 'StartService' @{ Name=$n } ($Opt -contains 'Services') '-StartRdpServices' 'Start only; existing sessions unaffected.'
    } }

# 5. IFEO (opt-in)
$ifeo = Get-RegState $IfeoKey 'Debugger'
if ($ifeo.Exists) { Add-P 'Remove IFEO Debugger blocking rdpclip.exe' "$IfeoKey\Debugger" $ifeo.Value '(removed)' 'RemoveValue' @{ Path=$IfeoKey; Name='Debugger' } ($Opt -contains 'Ifeo') '-RemoveRdpclipIfeo' 'Confirm with the security team that this block is not intentional.' }

# 6. Report-only items
$ap = $null; try { [xml]$ap = Get-AppLockerPolicy -Effective -Xml } catch {}
if ($ap -and @($ap.AppLockerPolicy.RuleCollection | Where-Object { $_.Type -eq 'Exe' -and $_.EnforcementMode -eq 'Enabled' }).Count) { Add-P 'AppLocker Exe rules enforced' 'Get-AppLockerPolicy -Effective' 'Enabled' 'rdpclip.exe allowed' $null $null $false '' 'Not remediated automatically - security team to verify an allow rule for %SystemRoot%\System32\rdpclip.exe.' }

if ($Plan.Count -eq 0) { Add-A 'Plan' $env:COMPUTERNAME 'No Windows-side blocker found' 'n/a' 'NO ACTION' 'Nothing to remediate on Windows; continue with BeyondTrust PAM verification.'; Emit; return }

# ---------------- Plan output or apply
if ($Mode -eq 'Plan') {
    foreach ($p in $Plan) { $r = if ($p.Note -like 'Not needed*') { 'NOT NEEDED' } elseif (-not $p.Act) { 'REPORT ONLY' } elseif ($p.Enabled) { 'PLANNED' } else { "SKIPPED (needs $($p.Need))" }; Add-A $p.Item $p.Target $p.Current $p.New $r $p.Note }
    Add-A 'Local GPO' $PolFile '' '' 'INFO' $PolNote
    Emit; return
}

# Apply
$todo = @($Plan | Where-Object { $_.Act -and $_.Enabled })
foreach ($p in @($Plan | Where-Object { -not ($_.Act -and $_.Enabled) })) { $r = if ($p.Note -like 'Not needed*') { 'NOT NEEDED' } elseif (-not $p.Act) { 'REPORT ONLY' } else { "SKIPPED (needs $($p.Need))" }; Add-A $p.Item $p.Target $p.Current $p.New $r $p.Note }
if ($todo.Count -eq 0) { Emit; return }

$script:Bk = Join-Path $BkRoot $RunId
New-Item -ItemType Directory -Path $script:Bk -Force | Out-Null
if ($PolOk) { Copy-Item -LiteralPath $PolFile -Destination "$script:Bk\Registry.pol" -Force; if (Test-Path $GptIni) { Copy-Item -LiteralPath $GptIni -Destination "$script:Bk\gpt.ini" -Force } }
Save-Changes

foreach ($p in $todo) {
    try { Invoke-Act $p; Add-A $p.Item $p.Target $p.Current $p.New 'APPLIED' $p.Note }
    catch { Add-A $p.Item $p.Target $p.Current $p.New 'FAILED' $_.Exception.Message }
    Save-Changes
}

if ($PolDirty) {
    try {
        [IO.File]::WriteAllBytes($PolFile, (ConvertTo-PolBytes $Pol))
        [void]$Changes.Add([ordered]@{ Type='Pol'; Path=$PolFile }); Save-Changes
        if (Test-Path $GptIni) {
            $g = Get-Content -LiteralPath $GptIni
            $g = $g | ForEach-Object { if ($_ -match '^\s*Version\s*=\s*(\d+)') { $v = [uint32]$Matches[1]; $mv = ($v -band 0xFFFF) + 1; if ($mv -gt 0xFFFF) { $mv = 1 }; "Version=$(($v -band 0xFFFF0000) + $mv)" } else { $_ } }
            Set-Content -LiteralPath $GptIni -Value $g -Encoding ASCII
        }
        Add-A 'Local GPO Registry.pol updated' $PolFile '' 'Matching LGPO entries updated; gpt.ini machine version incremented' 'APPLIED' 'gpedit.msc now shows the new setting.'
    } catch { Add-A 'Local GPO Registry.pol update' $PolFile '' '' 'FAILED' $_.Exception.Message }
}
if ($Opt -notcontains 'NoGpUpdate') {
    $g = & gpupdate.exe /target:computer /force 2>&1 | Out-String
    Add-A 'gpupdate /target:computer /force' 'Local policy refresh' '' "exit $LASTEXITCODE" $(if ($LASTEXITCODE -eq 0) { 'APPLIED' } else { 'FAILED' }) (($g -replace '\s+', ' ').Trim())
}

# Post-verification
foreach ($n in 'fDisableClip', 'fDisableCdm', 'CSClipLevel', 'SCClipLevel') { $s = Get-RegState $TSP $n; Add-A "Verify $n (after)" "HKLM\$TSPk\$n" '' (Show-State $s) 'VERIFY' '' }
Add-A 'Next step' 'RDP users' '' '' 'ACTION REQUIRED' 'Users must sign out and reconnect (new session) for redirection settings to take effect. Then re-run Invoke-RdpClipboardAssessment.ps1 during a live BeyondTrust session.'
Emit
'@
#endregion

function Invoke-GuestPayload {
    # Runs the payload with -ScriptType Powershell; if PowerCLI cannot locate the guest PowerShell
    # interpreter (known detection issue), copies the payload to the guest and runs it via Bat instead.
    param($VM, [string]$Payload, $Cred)
    try { return Invoke-VMScript -VM $VM -ScriptText $Payload -ScriptType Powershell -GuestCredential $Cred -ErrorAction Stop -WarningAction SilentlyContinue }
    catch { if ($_.Exception.Message -notmatch 'Could not locate "Powershell"') { throw } }
    Write-Log "$($VM.Name): PowerShell interpreter not detected by PowerCLI - using copy + Bat fallback." 'WARN'
    $name = 'rdpclip_' + [guid]::NewGuid().ToString('N') + '.ps1'
    $local = Join-Path $env:TEMP $name; $remote = "C:\Windows\Temp\$name"
    Set-Content -Path $local -Value $Payload -Encoding ASCII
    try {
        Copy-VMGuestFile -LocalToGuest -Source $local -Destination $remote -VM $VM -GuestCredential $Cred -Force -ErrorAction Stop | Out-Null
        $bat = "%SystemRoot%\System32\WindowsPowerShell\v1.0\powershell.exe -NoProfile -NonInteractive -ExecutionPolicy Bypass -File `"$remote`" & del /f /q `"$remote`""
        return Invoke-VMScript -VM $VM -ScriptText $bat -ScriptType Bat -GuestCredential $Cred -ErrorAction Stop -WarningAction SilentlyContinue
    } finally { Remove-Item -LiteralPath $local -Force -ErrorAction SilentlyContinue }
}

#region ------------------------------------------------------------------ VM resolution
$Rows = New-Object System.Collections.ArrayList
$Errs = New-Object System.Collections.ArrayList
function Get-VCName($vm) { if ($vm.Uid -match '@([^:/]+)') { $Matches[1] } else { '?' } }
$names = if ($VMListFile) { @(Get-Content $VMListFile | ForEach-Object { $_.Trim() } | Where-Object { $_ -and $_ -notmatch '^#' }) } else { $VMName }
$seen = @{}; $vms = @()
foreach ($n in $names) {
    $found = @(Get-VM -Name $n -Server $viServers -ErrorAction SilentlyContinue)
    if (-not $found) { [void]$Errs.Add([pscustomobject]@{ vCenter = '-'; 'VM Name' = $n; Error = 'Not found in any connected vCenter'; 'Next Step' = 'Check name / connection.' }); continue }
    foreach ($vm in $found) {
        $v = $vm.ExtensionData; $key = if ($v.Config.InstanceUuid) { $v.Config.InstanceUuid } else { $vm.Uid }
        if ($seen.ContainsKey($key)) { continue }; $seen[$key] = $true
        $why = if ($v.Config.Template) { 'Template' }
               elseif ($v.Config.GuestId -notlike 'windows*' -and $v.Guest.GuestFamily -ne 'windowsGuest') { 'Not a Windows guest' }
               elseif ($vm.PowerState -ne 'PoweredOn') { "Power state $($vm.PowerState)" }
               elseif ($v.Guest.ToolsRunningStatus -ne 'guestToolsRunning' -or -not $v.Guest.GuestOperationsReady) { 'VMware Tools / guest operations not ready' }
        if ($why) { [void]$Errs.Add([pscustomobject]@{ vCenter = (Get-VCName $vm); 'VM Name' = $vm.Name; Error = $why; 'Next Step' = 'Skipped.' }); continue }
        $vms += $vm
    }
}
#endregion

#region ------------------------------------------------------------------ Execute
foreach ($vm in $vms) {
    $vc = Get-VCName $vm
    if ($Mode -ne 'Plan' -and -not $PSCmdlet.ShouldProcess("$($vm.Name) @ $vc", "$Mode RDP clipboard remediation")) {
        Write-Log "$($vm.Name): not confirmed - skipped." 'WARN'
        [void]$Errs.Add([pscustomobject]@{ vCenter = $vc; 'VM Name' = $vm.Name; Error = 'Not confirmed by operator'; 'Next Step' = 'Re-run if intended.' }); continue
    }
    Write-Log "$($vm.Name) @ ${vc}: $Mode"
    $payload = $GuestPayload.Replace('__MODE__', $Mode).Replace('__OPTS__', ($opts -join ',')).Replace('__RUNID__', $ts)
    try {
        $res = Invoke-GuestPayload -VM $vm -Payload $payload -Cred $GuestCred
        $txt = [string]$res.ScriptOutput
        if ($txt -notmatch '###REMJSON###(.*)###END###') { $o = $txt -replace '\s+', ' '; throw "No parsable result (exit $($res.ExitCode)): $($o.Substring(0, [math]::Min(300, $o.Length)))" }
        $data = $Matches[1] | ConvertFrom-Json
        foreach ($a in @($data.a)) {
            [void]$Rows.Add([pscustomobject]@{ vCenter = $vc; 'VM Name' = $vm.Name; Mode = $Mode; Item = $a.i; Target = $a.t; 'Current Value' = $a.c; 'New Value' = $a.p; Result = $a.r; Note = $a.n; 'Guest Backup Folder' = $data.bk })
        }
        $failed = @($data.a | Where-Object { $_.r -eq 'FAILED' }).Count
        Write-Log ("{0}: {1} action rows, {2} failed{3}" -f $vm.Name, @($data.a).Count, $failed, $(if ($data.bk) { "; backup $($data.bk)" })) $(if ($failed) { 'ERR' } else { 'OK' })
    } catch {
        Write-Log "$($vm.Name): $($_.Exception.Message)" 'ERR'
        [void]$Errs.Add([pscustomobject]@{ vCenter = $vc; 'VM Name' = $vm.Name; Error = $_.Exception.Message; 'Next Step' = 'Verify credential / VMware Tools; check guest state before retrying.' })
    }
}
#endregion

#region ------------------------------------------------------------------ Report
if ($Rows.Count -eq 0) { [void]$Rows.Add([pscustomobject]@{ vCenter = '-'; 'VM Name' = '-'; Mode = $Mode; Item = 'No results'; Result = '-' }) }
if ($Errs.Count -eq 0) { [void]$Errs.Add([pscustomobject]@{ vCenter = '-'; 'VM Name' = '-'; Error = 'None'; 'Next Step' = '-' }) }
if (Get-Module -ListAvailable -Name ImportExcel) {
    Import-Module ImportExcel
    $ct = @(
        New-ConditionalText -Text 'FAILED'          -BackgroundColor LightPink            -ConditionalTextColor DarkRed
        New-ConditionalText -Text 'SKIPPED'         -BackgroundColor LightGoldenrodYellow -ConditionalTextColor DarkGoldenrod
        New-ConditionalText -Text 'ACTION REQUIRED' -BackgroundColor LightGoldenrodYellow -ConditionalTextColor DarkGoldenrod
        New-ConditionalText -Text 'REPORT ONLY'     -BackgroundColor Gainsboro            -ConditionalTextColor DimGray
        New-ConditionalText -Text 'PLANNED'         -BackgroundColor AliceBlue            -ConditionalTextColor SteelBlue
        New-ConditionalText -Text 'APPLIED'         -BackgroundColor Honeydew             -ConditionalTextColor DarkGreen
        New-ConditionalText -Text 'RESTORED'        -BackgroundColor Honeydew             -ConditionalTextColor DarkGreen
    )
    $common = @{ Path = $XlFile; AutoSize = $true; FreezeTopRow = $true; AutoFilter = $true; BoldTopRow = $true; ConditionalText = $ct }
    $Rows | Export-Excel @common -WorksheetName 'Remediation Actions' -TableStyle Medium2 -TableName Actions
    $Errs | Export-Excel @common -WorksheetName 'Errors and Skipped' -TableStyle Medium3 -TableName Errors
    $pkg = Open-ExcelPackage -Path $XlFile
    foreach ($ws in $pkg.Workbook.Worksheets) { if ($ws.Dimension) { for ($c = 1; $c -le $ws.Dimension.End.Column; $c++) { $col = $ws.Column($c); if ($col.Width -gt 55) { $col.Width = 55 }; $col.Style.WrapText = $true } } }
    Close-ExcelPackage $pkg
    Write-Log "Report: $XlFile" 'OK'
} else {
    $base = $XlFile -replace '\.xlsx$', ''
    $Rows | Export-Csv "${base}_Actions.csv" -NoTypeInformation
    $Errs | Export-Csv "${base}_Errors.csv" -NoTypeInformation
    Write-Log "ImportExcel not installed - CSV written: ${base}_*.csv" 'WARN'
}
Write-Log "Log: $LogFile"
#endregion
