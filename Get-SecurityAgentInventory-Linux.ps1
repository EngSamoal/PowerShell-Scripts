<#
.SYNOPSIS
    READ-ONLY inventory of Qualys / Binalyze AIR / Splunk UF agents on the LINUX VMs of
    ABHA, ALHAMRA and GOLF. Output: C:\temp\Security_Agent_Inventory_Linux.xlsx
    (Windows VMs are covered by Get-SecurityAgentInventory.ps1.)

.NOTES
    - Reuses the vCenter sessions already open in this PowerShell window (no Connect-VIServer).
    - Runs one small read-only bash query in each guest with Invoke-VMScript (VMware Tools, no SSH).
      Nothing is installed, started, stopped or changed. Unavoidable side effects of
      Invoke-VMScript: a guest login (recorded in /var/log/secure or auth.log), a temporary
      script/output file in /tmp that VMware Tools deletes afterwards, and guest-operation entries
      in vCenter/ESXi logs.
    - Root is NOT required: rpm/dpkg queries, reading systemd unit files and 'systemctl is-active'
      all work as a normal user.
    - Anything that can't be proven is reported as 'Unknown' / 'Unable to Verify', never guessed.
    - Linux credential file, created once with:  Get-Credential | Export-Clixml C:\temp\linuxcred.xml
      (Export-Clixml encrypts the password with DPAPI - only your user on this laptop can read it.)
#>

$ScriptBuild = 'SecAgentInv-Linux 2026-10-29.1'
Write-Host "Build: $ScriptBuild" -ForegroundColor Cyan

#region ---------- CONFIGURATION (edit here) ----------
# Site -> vCenter(s) exactly as they appear in $global:DefaultVIServers (IP or FQDN).
# Datacenter = @() means every datacenter on that vCenter; list names if sites share a vCenter.
$Sites = [ordered]@{
    'ABHA'    = @{ VCenter = @('10.61.6.140'); Datacenter = @(); TableStyle = 'Medium2'; Tab = '4472C4' }  # blue
    'ALHAMRA' = @{ VCenter = @('10.11.7.120'); Datacenter = @(); TableStyle = 'Medium3'; Tab = 'ED7D31' }  # orange
    'GOLF'    = @{ VCenter = @('GOLF-VCENTER-CHANGE-ME'); Datacenter = @(); TableStyle = 'Medium7'; Tab = '70AD47' }  # green - set the real GOLF vCenter IP/FQDN
}
# Credential files are tried in order; the next one is only tried if the guest rejects the login.
$LinuxCredFiles = @('C:\temp\linuxcred.xml')
# VMs whose NAME matches this are treated as appliances and excluded (logged, never silently).
$ApplianceNameRegex = 'vcsa|vcenter|nsx|vrops|vrli|aria|vcls|fortigate|forti|paloalto|panorama|bigip|f5-|cohesity|rubrik|esxi'
$OutFile = 'C:\temp\Security_Agent_Inventory_Linux.xlsx'
$LogFile = 'C:\temp\Security_Agent_Inventory_Linux_Log.csv'
#endregion

# Read-only guest query: installed package (rpm or dpkg - RHEL/CentOS/Rocky/Alma/Oracle/SUSE/Ubuntu/
# Debian) + the systemd unit whose file mentions the product + 'systemctl is-active'.
$LinuxScript = @'
for a in 'Q;qualys;qualys' 'B;binalyze;binalyze' 'S;splunkforwarder;splunkforwarder'; do
IFS=';' read -r t p s <<< "$a"
k=$({ rpm -qa --qf '%{NAME} %{VERSION}-%{RELEASE}\n'; dpkg-query -W -f='${Status} ${Package} ${Version}\n' | grep '^install ok installed' | cut -d' ' -f4-; } 2>/dev/null | grep -iE "^[^ ]*($p)" | head -n1)
f=$(grep -lsiE "$s" /etc/systemd/system/*.service /usr/lib/systemd/system/*.service /lib/systemd/system/*.service 2>/dev/null | head -n1)
n=${f##*/}; st=; [ -n "$n" ] && st=$(systemctl is-active "$n" 2>/dev/null)
[ -z "$k" ] && [ -n "$n" ] && [ "$t" = S ] && [ -r /opt/splunkforwarder/etc/splunk.version ] && k="splunk.version $(sed -n 's/^VERSION=//p' /opt/splunkforwarder/etc/splunk.version)"
echo "$t|$k|$n|$st"
done
echo END
'@

#region ---------- Helpers ----------
$Log = [System.Collections.Generic.List[object]]::new()
function Write-Log($Site, $VM, $Check, $Reason) {
    $Log.Add([pscustomobject]@{ Timestamp = (Get-Date -Format 'yyyy-MM-dd HH:mm:ss'); Site = $Site; VM = $VM; Check = $Check; Reason = $Reason })
    Write-Host "  [$Check] $VM - $Reason" -ForegroundColor Yellow
}

# Turns one guest output line (t|"pkg version"|unit|is-active) into Version / Status / Evidence.
function Resolve-Agent($Line) {
    if (-not $Line) { return @('Unable to Verify', 'Unknown', 'no output for this agent') }
    $f = $Line.Split('|')
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
$LinuxCreds = @(foreach ($f in $LinuxCredFiles) { if (Test-Path $f) { Import-Clixml $f } })
if (-not $LinuxCreds) { Write-Warning "No Linux credential file found - all VMs will be 'Unable to Verify'." }
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
            if ($v.Guest.GuestFamily -eq 'windowsGuest' -or $v.Config.GuestId -like 'win*') { continue }   # Linux only
            $full = "$($v.Config.GuestFullName) $($v.Guest.GuestFullName)"
            if (-not ($v.Guest.GuestFamily -eq 'linuxGuest' -or $full -match 'Linux|Red Hat|CentOS|Ubuntu|Debian|SUSE|Rocky|Alma|Oracle')) { continue }
            $why = if ($full -match 'Photon') { 'VMware Photon OS appliance' }
                   elseif ($vm.Name -match $ApplianceNameRegex) { 'name matches appliance pattern' }
                   elseif ($v.Config.VAppConfig.Product) { 'deployed as OVF appliance (vApp product info)' }
                   elseif ($v.Config.ManagedBy) { "managed by $($v.Config.ManagedBy.ExtensionKey)" }
            if ($why) { Write-Log $site $vm.Name 'Excluded' $why; continue }

            Write-Host "[$site] $($vm.Name)"
            $ip = @($v.Guest.IpAddress) + @($vm.Guest.IPAddress) | Where-Object { $_ -match '^\d+\.\d+\.\d+\.\d+$' -and $_ -notlike '169.254.*' } | Select-Object -First 1
            $row = [ordered]@{ Site = $site; 'VM Name' = $vm.Name; 'IP' = $(if ($ip) { $ip } else { 'Unable to Verify' }); 'OS' = 'Linux' }
            $lines = @(); $result = 'OK'
            if (-not $v.Guest.GuestOperationsReady) { $result = "VMware Tools not ready for guest operations (Tools: $($v.Guest.ToolsRunningStatus))" }
            else {
                if (-not $LinuxCreds) { $result = 'No Linux credential file' }
                foreach ($c in $LinuxCreds) {
                    try {
                        $out = Invoke-VMScript -VM $vm -Server $srv -GuestCredential $c -ScriptText $LinuxScript -ScriptType Bash -ToolsWaitSecs 60 -ErrorAction Stop
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
            foreach ($a in @(@('Q', 'Qualys'), @('B', 'Binalyze'), @('S', 'Splunk'))) {
                $r = if ($result -eq 'OK') { Resolve-Agent ($lines | Where-Object { $_ -like "$($a[0])|*" } | Select-Object -First 1) } else { @('Unable to Verify', 'Unknown', '') }
                $row["$($a[1]) Agent Version"] = $r[0]
                $row["$($a[1]) Service Status"] = $r[1]
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
        'Splunk Agent Version', 'Splunk Service Status', 'Verification', 'Evidence / Source', 'vCenter'
function Get-Color($Hex) { [System.Drawing.ColorTranslator]::FromHtml("#$Hex") }
# green = running, red = missing, orange = stopped, yellow = could not verify
$cond = foreach ($c in @(('Running', 'C6EFCE', '006100'), ('Not Installed', 'FFC7CE', '9C0006'), ('Stopped', 'F8CBAD', '833C0B'),
                        ('Unknown', 'FFEB9C', '9C5700'), ('Unable', 'FFEB9C', '9C5700'))) {
    New-ConditionalText -Range 'D:J' -Text $c[0] -BackgroundColor (Get-Color $c[1]) -ConditionalTextColor (Get-Color $c[2])
}

$Summary = foreach ($site in $Sites.Keys) {
    $r = @($Rows | Where-Object Site -eq $site)
    $s = [ordered]@{ Site = $site; 'Eligible VMs' = $r.Count; 'Verified' = @($r | Where-Object Verification -eq 'OK').Count; 'Unable to Verify' = @($r | Where-Object Verification -ne 'OK').Count }
    foreach ($a in 'Qualys', 'Binalyze', 'Splunk') {
        $st = @($r | ForEach-Object { $_."$a Service Status" })
        # confirmed installed = service state proven, or version proven by the package database
        $s["$a Installed"]     = @($r | Where-Object { $_."$a Service Status" -notin 'Not Installed', 'Unknown' -or $_."$a Agent Version" -notin 'Not Installed', 'Unable to Verify' }).Count
        $s["$a Running"]       = @($st | Where-Object { $_ -eq 'Running' }).Count
        $s["$a Not Installed"] = @($st | Where-Object { $_ -eq 'Not Installed' }).Count
        $s["$a Unknown"]       = @($st | Where-Object { $_ -eq 'Unknown' }).Count
    }
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
