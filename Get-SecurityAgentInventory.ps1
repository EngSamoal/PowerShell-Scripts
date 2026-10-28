<#
.SYNOPSIS
    READ-ONLY inventory of Qualys / Binalyze AIR / Splunk UF / Azure Arc agents on the
    Windows + Linux VMs of ABHA, ALHAMRA and GOLF. Output: C:\temp\Security_Agent_Inventory.xlsx

.NOTES
    - Reuses the vCenter sessions already open in this PowerShell window (no Connect-VIServer).
    - Runs one small read-only query in each guest with Invoke-VMScript (VMware Tools, no WinRM).
      Nothing is installed, started, stopped or changed. Unavoidable side effects of
      Invoke-VMScript: a guest logon (Windows event 4624, and a user profile on the very first
      logon of that account), a temporary script/output file in the guest temp folder that
      VMware Tools deletes afterwards, and guest-operation entries in vCenter/ESXi logs.
      'azcmagent show' also writes a line to the Arc agent's own log.
    - Anything that can't be proven is reported as 'Unknown' / 'Unable to Verify', never guessed.
    - Linux needs its own credential (Windows credentials don't work on Linux). Create once:
          Get-Credential | Export-Clixml C:\temp\linuxcred.xml
      (Export-Clixml encrypts the password with DPAPI - only your user on this laptop can read it.)
#>

$ScriptBuild = 'SecAgentInv 2026-10-28.1'
Write-Host "Build: $ScriptBuild" -ForegroundColor Cyan

#region ---------- CONFIGURATION (edit here) ----------
# Site -> vCenter(s) exactly as they appear in $global:DefaultVIServers (IP or FQDN).
# Datacenter = @() means every datacenter on that vCenter; list names if sites share a vCenter.
$Sites = [ordered]@{
    'ABHA'    = @{ VCenter = @('10.61.6.140'); Datacenter = @(); TableStyle = 'Medium2'; Tab = '4472C4' }  # blue
    'ALHAMRA' = @{ VCenter = @('10.11.7.120'); Datacenter = @(); TableStyle = 'Medium3'; Tab = 'ED7D31' }  # orange
    'GOLF'    = @{ VCenter = @('GOLF-VCENTER-CHANGE-ME'); Datacenter = @(); TableStyle = 'Medium7'; Tab = '70AD47' }  # green
}
# Credential files are tried in order; the next one is only tried if the guest rejects the login.
$WinCredFiles   = @('C:\temp\wincred.xml')
$LinuxCredFiles = @('C:\temp\linuxcred.xml')
# VMs whose NAME matches this are treated as appliances and excluded (logged, never silently).
$ApplianceNameRegex = 'vcsa|vcenter|nsx|vrops|vrli|aria|vcls|fortigate|forti|paloalto|panorama|bigip|f5-|cohesity|rubrik|esxi'
$OutFile = 'C:\temp\Security_Agent_Inventory.xlsx'
$LogFile = 'C:\temp\Security_Agent_Inventory_Log.csv'
#endregion

#region ---------- Guest queries (read-only) ----------
# Windows: uninstall registry + Win32_Service + the service's .exe on disk. Pattern = regex used
# to find the product; preferred service name is tried first.
$WinScript = @'
$ErrorActionPreference='SilentlyContinue'
$u=@(Get-ItemProperty HKLM:\SOFTWARE\Microsoft\Windows\CurrentVersion\Uninstall\*,HKLM:\SOFTWARE\WOW6432Node\Microsoft\Windows\CurrentVersion\Uninstall\*)
$s=@(Get-WmiObject Win32_Service)
foreach($a in 'Q;QualysAgent;Qualys','B;-;Binalyze','S;SplunkForwarder;SplunkForwarder|UniversalForwarder','A;himds;himds|Azure ?Connected ?Machine'){
$t,$n,$p=$a.Split(';')
$r=$u|Where-Object{$_.DisplayName -match $p}|Select-Object -First 1
$v=$s|Where-Object{$_.Name -match $p -or $_.PathName -match $p}|Sort-Object {$_.Name -ne $n}|Select-Object -First 1
$e=if($v){$v.PathName -replace '^"?([^"]+?\.exe).*','$1'}
$x=[bool]($e -and (Test-Path -LiteralPath $e))
$f=if($x){(Get-Item -LiteralPath $e).VersionInfo.ProductVersion}
"$t|$($r.DisplayVersion)|$($r.DisplayName)|$($v.Name)|$($v.State)|$x|$f"}
$z="$env:ProgramFiles\AzureConnectedMachineAgent\azcmagent.exe"
if(Test-Path $z){$o=& $z show 2>&1|Out-String;if($o -match 'Agent Status\s*:\s*(\S+)'){"ARC|$($matches[1])"}}
'END'
'@
# Linux: installed package (rpm or dpkg, works on RHEL/CentOS/Rocky/SUSE/Ubuntu/Debian) + systemd
# unit whose file mentions the product + systemctl is-active. No root needed except for azcmagent.
$LinuxScript = @'
for a in 'Q;qualys;qualys' 'B;binalyze;binalyze' 'S;splunkforwarder;splunkforwarder' 'A;azcmagent;himds'; do
IFS=';' read -r t p s <<< "$a"
k=$({ rpm -qa --qf '%{NAME} %{VERSION}-%{RELEASE}\n'; dpkg-query -W -f='${Status} ${Package} ${Version}\n' | grep '^install ok installed' | cut -d' ' -f4-; } 2>/dev/null | grep -iE "^[^ ]*($p)" | head -n1)
f=$(grep -lsiE "$s" /etc/systemd/system/*.service /usr/lib/systemd/system/*.service /lib/systemd/system/*.service 2>/dev/null | head -n1)
n=${f##*/}; st=; [ -n "$n" ] && st=$(systemctl is-active "$n" 2>/dev/null)
[ -z "$k" ] && [ -n "$n" ] && [ "$t" = S ] && [ -r /opt/splunkforwarder/etc/splunk.version ] && k="splunk.version $(sed -n 's/^VERSION=//p' /opt/splunkforwarder/etc/splunk.version)"
echo "$t|$k|$n|$st"
done
command -v azcmagent >/dev/null 2>&1 && azcmagent show 2>/dev/null | sed -n 's/^[[:space:]]*Agent Status[[:space:]]*:[[:space:]]*\([^[:space:]]*\).*/ARC|\1/p'
echo END
'@
#endregion

#region ---------- Helpers ----------
$Log = [System.Collections.Generic.List[object]]::new()
function Write-Log($Site, $VM, $Check, $Reason) {
    $Log.Add([pscustomobject]@{ Timestamp = (Get-Date -Format 'yyyy-MM-dd HH:mm:ss'); Site = $Site; VM = $VM; Check = $Check; Reason = $Reason })
    Write-Host "  [$Check] $VM - $Reason" -ForegroundColor Yellow
}
function Get-Creds($Files) { foreach ($f in $Files) { if (Test-Path $f) { Import-Clixml $f } } }

# Turns one guest output line into Version / Status / Evidence for one agent.
function Resolve-Agent($Os, $Line) {
    if (-not $Line) { return @('Unable to Verify', 'Unknown', 'no output for this agent') }
    $f = $Line.Split('|')
    if ($Os -eq 'Windows') {   # t|regVer|regName|svc|state|exeExists|fileVer
        if ($f[3] -and $f[5] -eq 'True') {
            $ver = if ($f[1]) { $f[1]; $src = 'uninstall registry' } elseif ($f[6]) { $f[6]; $src = 'exe file version' } else { 'Unable to Verify'; $src = 'none' }
            return @($ver, $f[4], "service '$($f[3])'=$($f[4]); version from $src")
        }
        if ($f[3]) { return @('Unable to Verify', 'Unknown', "service '$($f[3])' exists but its exe is missing") }
        if ($f[1] -or $f[2]) { return @('Unable to Verify', 'Unknown', "uninstall entry '$($f[2])' but no service (possible leftover)") }
        return @('Not Installed', 'Not Installed', 'no service, no uninstall entry')
    }
    # Linux: t|"pkg version"|unit|is-active
    $ver = if ($f[1]) { ($f[1] -split ' ', 2)[1] } else { 'Unable to Verify' }
    if ($f[2]) {
        $st = switch ($f[3]) { 'active' { 'Running' } { $_ -in 'inactive', 'failed' } { 'Stopped' } default { 'Unknown' } }
        if (-not $f[1] -and $st -ne 'Running') { return @('Unable to Verify', 'Unknown', "unit '$($f[2])'=$($f[3]) but no package (possible leftover)") }
        return @($ver, $st, "unit '$($f[2])'=$($f[3]); version from $(if ($f[1]) { ($f[1] -split ' ')[0] } else { 'none' })")
    }
    if ($f[1]) { return @($ver, 'Unknown', "package '$($f[1])' but no systemd unit found") }
    return @('Not Installed', 'Not Installed', 'no package, no systemd unit')
}
#endregion

#region ---------- Pre-checks ----------
if (-not (Get-Module -ListAvailable ImportExcel)) { throw 'ImportExcel module missing. Run: Install-Module ImportExcel -Scope CurrentUser' }
$servers = @($global:DefaultVIServers | Where-Object IsConnected)
if (-not $servers) { throw 'No connected vCenter. Run Connect-VIServer first.' }
$WinCreds = @(Get-Creds $WinCredFiles); $LinuxCreds = @(Get-Creds $LinuxCredFiles)
if (-not $WinCreds) { Write-Warning "No Windows credential file found - Windows VMs will be 'Unable to Verify'." }
if (-not $LinuxCreds) { Write-Warning "No Linux credential file found - Linux VMs will be 'Unable to Verify'." }
#endregion

$Rows = [System.Collections.Generic.List[object]]::new()
$seen = @{}
foreach ($site in $Sites.Keys) {
    $cfg = $Sites[$site]
    $siteServers = @($servers | Where-Object { $_.Name -in $cfg.VCenter -or $_.ServiceUri.Host -in $cfg.VCenter })
    if (-not $siteServers) { Write-Log $site '-' 'vCenter' "Not connected to $($cfg.VCenter -join ', ') - site NOT inventoried"; continue }
    foreach ($srv in $siteServers) {
        $vms = if ($cfg.Datacenter) { Get-Datacenter -Server $srv -Name $cfg.Datacenter | Get-VM -Server $srv } else { Get-VM -Server $srv }
        foreach ($vm in $vms) {
            $v = $vm.ExtensionData
            if ($seen[$v.Config.InstanceUuid]) { continue }   # same VM via a second connection
            $seen[$v.Config.InstanceUuid] = $true
            if ($vm.PowerState -ne 'PoweredOn' -or $v.Config.Template) { continue }
            $full = "$($v.Config.GuestFullName) $($v.Guest.GuestFullName)"
            $os = if ($v.Guest.GuestFamily -eq 'windowsGuest' -or $v.Config.GuestId -like 'win*') { 'Windows' }
                  elseif ($full -notmatch 'Photon' -and ($v.Guest.GuestFamily -eq 'linuxGuest' -or $full -match 'Linux|Red Hat|CentOS|Ubuntu|Debian|SUSE|Rocky|Alma|Oracle')) { 'Linux' }
            $why = if (-not $os) { "guest OS '$($v.Config.GuestFullName)' is not Windows/Linux" }
                   elseif ($vm.Name -match $ApplianceNameRegex) { 'name matches appliance pattern' }
                   elseif ($v.Config.VAppConfig.Product) { 'deployed as OVF appliance (vApp product info)' }
                   elseif ($v.Config.ManagedBy) { "managed by $($v.Config.ManagedBy.ExtensionKey)" }
            if ($why) { Write-Log $site $vm.Name 'Excluded' $why; continue }

            Write-Host "[$site] $($vm.Name)"
            $ip = @($v.Guest.IpAddress) + @($vm.Guest.IPAddress) | Where-Object { $_ -match '^\d+\.\d+\.\d+\.\d+$' -and $_ -notlike '169.254.*' } | Select-Object -First 1
            $row = [ordered]@{ Site = $site; 'VM Name' = $vm.Name; 'IP' = $(if ($ip) { $ip } else { 'Unable to Verify' }); 'OS' = $os }
            $lines = @(); $result = 'OK'
            if (-not $v.Guest.GuestOperationsReady) { $result = "VMware Tools not ready for guest operations (Tools: $($v.Guest.ToolsRunningStatus))" }
            else {
                $creds = if ($os -eq 'Windows') { $WinCreds } else { $LinuxCreds }
                if (-not $creds) { $result = "No $os credential file" }
                foreach ($c in $creds) {
                    try {
                        $out = Invoke-VMScript -VM $vm -Server $srv -GuestCredential $c -ScriptText $(if ($os -eq 'Windows') { $WinScript } else { $LinuxScript }) `
                            -ScriptType $(if ($os -eq 'Windows') { 'Powershell' } else { 'Bash' }) -ToolsWaitSecs 60 -ErrorAction Stop
                        $lines = @("$($out.ScriptOutput)" -split "`r?`n" | ForEach-Object Trim)
                        $result = if ($lines -contains 'END') { 'OK' } else { 'Guest query did not complete (incomplete output)' }
                        break
                    } catch {
                        $m = $_.Exception.Message
                        $result = if ($m -match 'InvalidGuestLogin|password|logon|authenticat') { 'Guest authentication failed' }
                                  elseif ($m -match 'NoPermission|permission|not authorized') { 'Insufficient vCenter permission (needs VirtualMachine.GuestOperations.*)' }
                                  elseif ($m -match 'time[d ]?out') { 'Guest operation timed out' }
                                  else { "Guest operation failed: $($m -replace '\s+', ' ')" }
                        if ($result -ne 'Guest authentication failed') { break }   # only try next credential on bad login
                    }
                }
            }
            $evidence = @()
            foreach ($a in @(@('Q', 'Qualys'), @('B', 'Binalyze'), @('S', 'Splunk'), @('A', 'Arc'))) {
                $r = if ($result -eq 'OK') { Resolve-Agent $os ($lines | Where-Object { $_ -like "$($a[0])|*" } | Select-Object -First 1) } else { @('Unable to Verify', 'Unknown', '') }
                if ($a[1] -eq 'Arc') {
                    $arcLine = $lines | Where-Object { $_ -like 'ARC|*' } | Select-Object -First 1
                    $row['Azure Arc Status (local agent)'] = if ($result -ne 'OK') { 'Unknown' } elseif ($r[1] -eq 'Not Installed') { 'Not Installed' } elseif ($arcLine) { $arcLine.Split('|')[1] } else { 'Unknown' }
                } else {
                    $row["$($a[1]) Agent Version"] = $r[0]
                    $row["$($a[1]) Service Status"] = $r[1]
                }
                if ($r[2]) { $evidence += "$($a[1]): $($r[2])" }
            }
            $row['Verification'] = if ($result -eq 'OK') { 'OK' } else { "Unable to Verify: $result" }
            $row['Evidence / Source'] = $evidence -join ' | '
            $row['vCenter'] = $srv.Name
            if ($result -ne 'OK') { Write-Log $site $vm.Name 'Guest query' $result }
            $Rows.Add([pscustomobject]$row)
        }
    }
}

#region ---------- Excel ----------
if (Test-Path $OutFile) { Remove-Item $OutFile -ErrorAction Stop }   # fails clearly if the file is open in Excel
$cols = 'VM Name', 'IP', 'OS', 'Qualys Agent Version', 'Qualys Service Status', 'Binalyze Agent Version', 'Binalyze Service Status',
        'Splunk Agent Version', 'Splunk Service Status', 'Azure Arc Status (local agent)', 'Verification', 'Evidence / Source', 'vCenter'
function Get-Color($Hex) { [System.Drawing.ColorTranslator]::FromHtml("#$Hex") }
# green = running/connected, red = missing, orange = stopped/disconnected, yellow = could not verify
$cond = foreach ($c in @(('Running', 'C6EFCE', '006100'), ('Not Installed', 'FFC7CE', '9C0006'), ('Stopped', 'F8CBAD', '833C0B'),
                        ('Disconnected', 'F8CBAD', '833C0B'), ('Expired', 'F8CBAD', '833C0B'), ('Unknown', 'FFEB9C', '9C5700'), ('Unable', 'FFEB9C', '9C5700'))) {
    New-ConditionalText -Range 'D:K' -Text $c[0] -BackgroundColor (Get-Color $c[1]) -ConditionalTextColor (Get-Color $c[2])
}
$cond += New-ConditionalText -Range 'D:K' -ConditionalType BeginsWith -Text 'Connected' -BackgroundColor (Get-Color 'C6EFCE') -ConditionalTextColor (Get-Color '006100')

$Summary = foreach ($site in $Sites.Keys) {
    $r = @($Rows | Where-Object Site -eq $site)
    $s = [ordered]@{ Site = $site; 'Eligible VMs' = $r.Count; 'Verified' = @($r | Where-Object Verification -eq 'OK').Count; 'Unable to Verify' = @($r | Where-Object Verification -ne 'OK').Count }
    foreach ($a in 'Qualys', 'Binalyze', 'Splunk') {
        $st = @($r | ForEach-Object { $_."$a Service Status" })
        # confirmed installed = service state proven, or version proven by package/registry
        $s["$a Installed"]     = @($r | Where-Object { $_."$a Service Status" -notin 'Not Installed', 'Unknown' -or $_."$a Agent Version" -notin 'Not Installed', 'Unable to Verify' }).Count
        $s["$a Running"]       = @($st | Where-Object { $_ -eq 'Running' }).Count
        $s["$a Not Installed"] = @($st | Where-Object { $_ -eq 'Not Installed' }).Count
        $s["$a Unknown"]       = @($st | Where-Object { $_ -eq 'Unknown' }).Count
    }
    $arc = @($r | ForEach-Object { $_.'Azure Arc Status (local agent)' })
    $s['Arc Connected']     = @($arc | Where-Object { $_ -eq 'Connected' }).Count
    $s['Arc Disconnected/Expired'] = @($arc | Where-Object { $_ -in 'Disconnected', 'Expired' }).Count
    $s['Arc Not Installed'] = @($arc | Where-Object { $_ -eq 'Not Installed' }).Count
    $s['Arc Unknown']       = @($arc | Where-Object { $_ -notin 'Connected', 'Disconnected', 'Expired', 'Not Installed' }).Count
    [pscustomobject]$s
}
$xl = $Summary | Export-Excel -Path $OutFile -WorksheetName 'Summary' -TableName 'Summary' -TableStyle Medium4 -AutoSize -FreezeTopRow -PassThru
$xl.Workbook.Worksheets['Summary'].TabColor = Get-Color 'A5A5A5'
Close-ExcelPackage $xl
foreach ($site in $Sites.Keys) {
    $data = @($Rows | Where-Object Site -eq $site | Select-Object $cols)
    if (-not $data) { $data = @([pscustomobject]@{ 'VM Name' = 'No eligible VMs collected - see Verification Log' }) }
    $xl = $data | Export-Excel -Path $OutFile -WorksheetName $site -TableName "T_$site" -TableStyle $Sites[$site].TableStyle -AutoSize -FreezeTopRow -ConditionalText $cond -PassThru
    $xl.Workbook.Worksheets[$site].TabColor = Get-Color $Sites[$site].Tab
    Close-ExcelPackage $xl
}
if ($Log.Count) { $Log | Export-Excel -Path $OutFile -WorksheetName 'Verification Log' -TableName 'VerificationLog' -TableStyle Light1 -AutoSize -FreezeTopRow }
$Log | Export-Csv $LogFile -NoTypeInformation
#endregion

Write-Host "`nDone. Report: $OutFile   Log: $LogFile" -ForegroundColor Green
$Summary | Format-Table Site, 'Eligible VMs', Verified, 'Unable to Verify' -AutoSize
