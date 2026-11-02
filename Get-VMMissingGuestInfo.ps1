# Get-VMMissingGuestInfo.ps1  (Script 2 of 2)
# READ-ONLY retrieval of IP / OS for the VMs listed by Get-VMToolsHealth.ps1.
# Uses existing vCenter connections. Changes nothing on VMs, Tools, guest OS, network or vCenter.
# Input : C:\temp\VMTools_Health.csv   (output of Script 1)
# Output: C:\temp\VM_IP_OS_Results.csv (one row per virtual NIC)
# Run   : .\Get-VMMissingGuestInfo.ps1
#   Optional in-guest query (ONLY used when VMware Tools is running and IP/OS is still missing):
#   .\Get-VMMissingGuestInfo.ps1 -WindowsGuestCredential (Get-Credential) -LinuxGuestCredential (Get-Credential)
param(
    [string]$InputCsv = 'C:\temp\VMTools_Health.csv',
    [string]$OutFile  = 'C:\temp\VM_IP_OS_Results.csv',
    [pscredential]$WindowsGuestCredential,
    [pscredential]$LinuxGuestCredential
)

$servers = @{}
$global:DefaultVIServers | Where-Object IsConnected | ForEach-Object { $servers[$_.Name] = $_ }
if (-not $servers.Count) { Write-Error 'No connected vCenters. Run Connect-VIServer first.'; return }

# Read-only in-guest queries. Line 1 = OS, line 2 = IPv4 list.
$winScript = '$o = Get-WmiObject Win32_OperatingSystem; "$($o.Caption) $($o.Version)"; ' +
             '(Get-WmiObject Win32_NetworkAdapterConfiguration -Filter "IPEnabled=True" | ForEach-Object { $_.IPAddress } | Where-Object { $_ -match "^\d+\.\d+\.\d+\.\d+$" }) -join ","'
$linScript = @(
    'PATH=$PATH:/sbin:/usr/sbin'
    'if [ -r /etc/os-release ]; then . /etc/os-release; echo "$PRETTY_NAME"; else head -1 /etc/redhat-release 2>/dev/null || uname -sr; fi'
    "ip -4 -o addr show scope global 2>/dev/null | awk '{print `$4}' | cut -d/ -f1 | paste -sd, -"
) -join "`n"   # LF only - bash fails on CRLF

$results = foreach ($r in Import-Csv $InputCsv) {
    $row = [ordered]@{ 'vCenter' = $r.vCenter; 'VM Name' = $r.'VM Name'; 'MAC Address' = $null; 'Port Group' = $null
                       'IP Address' = $null; 'OS' = $null; 'Information Source' = $null
                       'Retrieval Status' = 'Failed'; 'Failure Reason' = $null; 'VM Id' = $r.'VM Id' }

    # --- Locate the exact VM: vCenter + MoRef (unique). Name match only as a fallback, and only if unique.
    $vc = $servers[$r.vCenter]
    if (-not $vc) { $row.'Failure Reason' = 'vCenter in CSV is not currently connected'; [pscustomobject]$row; continue }
    try {
        if ($r.'VM Id') { $vm = @(Get-VM -Server $vc -Id "VirtualMachine-$($r.'VM Id')" -ErrorAction Stop) }
        else {
            $vm = @(Get-View -Server $vc -ViewType VirtualMachine -Property Name -Filter @{ Name = "^$([regex]::Escape($r.'VM Name'))$" } |
                    Get-VIObjectByVIView)
        }
    } catch { $vm = @() }
    if ($vm.Count -ne 1) {
        $row.'Failure Reason' = "VM lookup returned $($vm.Count) matches (deleted, renamed, no permission, or duplicate name) - skipped"
        [pscustomobject]$row; continue
    }
    $vm = $vm[0]; $row.'VM Name' = $vm.Name
    $g = $vm.ExtensionData.Guest; $cfg = $vm.ExtensionData.Config

    $running   = $g.ToolsRunningStatus -in 'guestToolsRunning','guestToolsExecutingScripts'
    $isWin     = $g.GuestFamily -eq 'windowsGuest' -or $cfg.GuestId -like 'win*'
    $toolsIPs  = @($g.Net | ForEach-Object { $_.IpAddress } | Where-Object { $_ -and $_ -notlike 'fe80*' })
    $stale     = $(if ($running) { '' } else { ' [last known - Tools not running, may be stale]' })
    $reasons   = @()

    # --- Optional in-guest query: only when Tools is RUNNING, data is missing, and a credential was supplied
    $inGuestOS = $null; $inGuestIP = $null
    $cred = $(if ($isWin) { $WindowsGuestCredential } else { $LinuxGuestCredential })
    if ($running -and (-not $toolsIPs -or -not $g.GuestFullName)) {
        if ($cred) {
            try {
                $out = Invoke-VMScript -VM $vm -GuestCredential $cred -ErrorAction Stop `
                         -ScriptText $(if ($isWin) { $winScript } else { $linScript }) -ScriptType $(if ($isWin) { 'Powershell' } else { 'Bash' })
                $lines = @("$($out.ScriptOutput)" -split "`r?`n" | Where-Object { $_.Trim() })
                $inGuestOS = $lines[0]; $inGuestIP = $lines[1]
            } catch { $reasons += "Invoke-VMScript failed: $($_.Exception.Message)" }
        } else {
            $reasons += 'Tools running but IP/OS not reported; rerun with guest credential to query inside the guest'
        }
    }
    elseif (-not $running) {
        $reasons += "VMware Tools not running or status unknown (running=$(if ($g.ToolsRunningStatus) { $g.ToolsRunningStatus } else { 'unknown' }), version=$(if ($g.ToolsVersionStatus2) { $g.ToolsVersionStatus2 } else { 'unknown' })): guest IP and actual OS cannot be read via vSphere or Invoke-VMScript. Use MAC + port group with DHCP/IPAM/ARP/CMDB records (outside PowerCLI)"
    }

    # --- OS: actual (in-guest) > Tools-reported > configured (unverified)
    if     ($inGuestOS)        { $row.OS = $inGuestOS;          $osSrc = 'OS: actual, in-guest query (Invoke-VMScript)' }
    elseif ($g.GuestFullName)  { $row.OS = $g.GuestFullName;    $osSrc = "OS: reported by VMware Tools via vCenter$stale" }
    elseif ($cfg.GuestFullName){ $row.OS = $cfg.GuestFullName;  $osSrc = 'OS: CONFIGURED in VM settings only - not verified, may differ from actual OS' }
    else                       { $osSrc = 'OS: not available' }
    $osVerified = [bool]($inGuestOS -or $g.GuestFullName)

    # --- One row per vNIC: MAC + port group always come from VM hardware config (no Tools needed)
    $nics = @(Get-NetworkAdapter -VM $vm -ErrorAction SilentlyContinue)
    if (-not $nics) { $nics = @($null); $reasons += 'No virtual NIC found' }
    foreach ($n in $nics) {
        $nicRow  = [pscustomobject]$row   # fresh copy per NIC
        $nicWhy  = @($reasons)
        $nicIPs  = @($g.Net | Where-Object { $n -and $_.MacAddress -eq $n.MacAddress } |
                     ForEach-Object { $_.IpAddress } | Where-Object { $_ -and $_ -notlike 'fe80*' })
        if ($n) {
            $nicRow.'MAC Address' = $n.MacAddress
            $nicRow.'Port Group'  = $n.NetworkName
            if (-not $n.ConnectionState.Connected) { $nicWhy += 'vNIC is not connected' }
        }
        if     ($nicIPs)    { $nicRow.'IP Address' = $nicIPs -join ', '; $ipSrc = "IP: VMware Tools (Guest.Net, matched by MAC)$stale" }
        elseif ($inGuestIP) { $nicRow.'IP Address' = $inGuestIP;         $ipSrc = 'IP: in-guest query (Invoke-VMScript, all guest IPv4 - not mapped per NIC)' }
        else                { $ipSrc = 'IP: not available' }

        $nicRow.'Information Source' = "$ipSrc; $osSrc"
        $nicRow.'Retrieval Status'   = $(if ($nicRow.'IP Address' -and $osVerified) { 'Success' }
                                         elseif ($nicRow.'IP Address' -or $row.OS -or $n) { 'Partial' } else { 'Failed' })
        $nicRow.'Failure Reason'     = $nicWhy -join '; '
        $nicRow
    }
}

$results = @($results)
New-Item -ItemType Directory -Path (Split-Path $OutFile) -Force -ErrorAction SilentlyContinue | Out-Null
$results | Export-Csv -Path $OutFile -NoTypeInformation -Encoding UTF8
Write-Host "Rows: $($results.Count)  ->  $OutFile" -ForegroundColor Cyan
$results | Group-Object 'Retrieval Status' | Format-Table Name, Count -AutoSize
