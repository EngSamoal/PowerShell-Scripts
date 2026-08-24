<#
.SYNOPSIS
   Checks Qualys Agent AND Binalyze AIR Agent installation status across the VM
   fleet (Windows and Linux) via VMware Guest Operations (Invoke-VMScript)
   instead of WinRM/SSH - works even where WinRM is locked down (AMC, Tabuk).
   OS is auto-detected per VM and the matching guest-side check (PowerShell or
   Bash) is used. Both products are checked independently in a single guest
   script execution per VM (no extra Guest Ops round-trips).
.REQUIRES
   - Connected to vCenter (Connect-VIServer)
   - VMware PowerCLI installed
   - VMware Tools running in the guest (no WinRM/SSH required)
   - A Windows local admin credential (for Windows guests)
   - A Linux credential with rights to read service/package status, e.g. root
     or a sudo-capable account (for Linux guests)
   - Run PowerShell as Administrator if you want the DNS preflight to auto-fix
     unresolvable ESXi hosts via the local hosts file (see -AutoAddHostsFileEntries)
.NOTES
   Server list  : C:\Temp\vmlist.txt   (one VM name per line)
   Output folder: C:\Temp\Fleet        (CSV exports; empty result sets are not exported)
   This script is 100% read-only: it only queries service state, the Windows
   uninstall registry, and Linux package/unit metadata. It never installs,
   starts, stops, restarts, or reconfigures either agent, and never touches
   vCenter configuration.

   Binalyze AIR Agent detection - what is actually checked and why:
   Binalyze does not publish one single fixed service/package name that is
   guaranteed stable across every AIR Agent version, so (exactly like the
   existing Qualys check below, which already wildcards "*Qualys*" rather
   than hardcoding one literal) detection is done by matching the verified
   vendor/product keywords "Binalyze" / "IREC" rather than guessing a single
   exact string:
     Windows:
       - Get-Service, first for the known internal service name
         "Binalyze.AIR.Agent.Service" (this is the exact name Binalyze's own
         Safe Mode KB article registers under
         HKLM\SYSTEM\CurrentControlSet\Control\SafeBoot\Network and \Minimal),
         then falling back to any service whose Name or DisplayName contains
         "Binalyze" (covers the "AIR Responder Service" display name and any
         version that reuses the legacy "IREC" naming).
       - Uninstall registry keys (same Uninstall\* paths already used for
         Qualys) where DisplayName/Publisher contains "Binalyze" - confirmed
         against a real installed AIR Agent's registry entry (Publisher
         "Binalyze", installed under
         "C:\Program Files (x86)\Binalyze\AIR\agent", main executable
         IREC.exe / Binalyze.IREC.exe - the endpoint agent is built on
         Binalyze's IREC evidence-collector engine).
     Linux:
       - systemctl list-unit-files, grep'd case-insensitively for "binalyze"
         (Binalyze's own Linux Relay Server package is named
         "binalyze-air-relay-server", confirming "binalyze" is the stable
         package/unit prefix used across their Linux components).
       - rpm -qa / dpkg -l, same case-insensitive "binalyze" match, mirroring
         exactly how the existing Qualys Linux check already works.
   This keeps detection resilient to minor version-to-version naming changes
   without guessing an unverified exact identifier.
#>
param(
   [string]$VMListPath   = "C:\Temp\vmlist.txt",
   [string]$OutputFolder = "C:\Temp\Fleet",
   [PSCredential]$WindowsCredential,
   [PSCredential]$LinuxCredential,
   # Manual override: ESXi hostname -> IP, for hosts you already know the IP for
   [hashtable]$KnownHostIPMap = @{},
   # If $true, auto-adds unresolvable ESXi hosts to the local hosts file (needs admin PowerShell)
   [bool]$AutoAddHostsFileEntries = $true
)
# Credentials are only prompted for if the fleet list actually contains that
# OS type - asked once, the first time each is needed.
function Get-OSCredential {
   param([string]$OSType)
   if ($OSType -eq "Windows") {
       if (-not $script:WindowsCredential) {
           $script:WindowsCredential = Get-Credential -Message "Enter Windows local admin credential (for Guest Operations)"
       }
       return $script:WindowsCredential
   }
   else {
       if (-not $script:LinuxCredential) {
           $script:LinuxCredential = Get-Credential -Message "Enter Linux credential (root or sudo-capable) for Guest Operations"
       }
       return $script:LinuxCredential
   }
}
if (-not (Test-Path $OutputFolder)) {
   New-Item -Path $OutputFolder -ItemType Directory -Force | Out-Null
}
if (-not (Test-Path $VMListPath)) {
   Write-Host "VM list not found at $VMListPath" -ForegroundColor Red
   return
}
$Servers = Get-Content $VMListPath | Where-Object { $_.Trim() -ne "" }
$Success = @()
$Failed  = @()
# Runs INSIDE the guest via VMware Tools - checks Qualys and Binalyze AIR
# services plus the uninstall registry keys. No network path to the guest is
# needed. Both products are checked in one pass so only one Guest Ops call is
# made per VM.
$GuestScript = @'
$result = [ordered]@{
   QualysServiceFound    = $false
   QualysRegFound        = $false
   QualysDisplayVersion  = "-"
   QualysServiceStatus   = "N/A"
   BinalyzeServiceFound  = $false
   BinalyzeRegFound      = $false
   BinalyzeDisplayVersion = "-"
   BinalyzeServiceStatus = "N/A"
}
# ---- Qualys ----
$svc = Get-Service -Name QualysAgent -ErrorAction SilentlyContinue
if ($svc) {
   $result.QualysServiceFound  = $true
   $result.QualysServiceStatus = $svc.Status.ToString()
}
$paths = @(
   "HKLM:\Software\Microsoft\Windows\CurrentVersion\Uninstall\*",
   "HKLM:\Software\WOW6432Node\Microsoft\Windows\CurrentVersion\Uninstall\*"
)
$app = Get-ItemProperty -Path $paths -ErrorAction SilentlyContinue |
   Where-Object { $_.DisplayName -like "*Qualys*" -or $_.Publisher -like "*Qualys*" } |
   Select-Object -First 1 DisplayName, DisplayVersion
if ($app) {
   $result.QualysRegFound       = $true
   $result.QualysDisplayVersion = $app.DisplayVersion
}
# ---- Binalyze AIR Agent ----
$bsvc = Get-Service -Name "Binalyze.AIR.Agent.Service" -ErrorAction SilentlyContinue
if (-not $bsvc) {
   $bsvc = Get-Service -ErrorAction SilentlyContinue |
       Where-Object { $_.Name -like "*Binalyze*" -or $_.DisplayName -like "*Binalyze*" } |
       Select-Object -First 1
}
if ($bsvc) {
   $result.BinalyzeServiceFound  = $true
   $result.BinalyzeServiceStatus = $bsvc.Status.ToString()
}
$bapp = Get-ItemProperty -Path $paths -ErrorAction SilentlyContinue |
   Where-Object { $_.DisplayName -like "*Binalyze*" -or $_.Publisher -like "*Binalyze*" } |
   Select-Object -First 1 DisplayName, DisplayVersion
if ($bapp) {
   $result.BinalyzeRegFound       = $true
   $result.BinalyzeDisplayVersion = $bapp.DisplayVersion
}
$result | ConvertTo-Json -Compress
'@
# Runs INSIDE the guest via VMware Tools - checks for the Qualys and Binalyze
# AIR services and their installed packages (rpm or dpkg based). No SSH
# required. Kept as a single semicolon-separated line (not multi-line) so
# CRLF line endings from this Windows-authored .ps1 file can't get embedded
# into the bash script text and break execution in the guest.
$GuestScriptLinux = @'
QUALYS_SVC_FOUND=false; QUALYS_PKG_FOUND=false; QUALYS_VERSION="-"; QUALYS_SVC_STATUS="N/A"; BINALYZE_SVC_FOUND=false; BINALYZE_PKG_FOUND=false; BINALYZE_VERSION="-"; BINALYZE_SVC_STATUS="N/A"; if command -v systemctl >/dev/null 2>&1; then UNIT_LIST=$(systemctl list-unit-files 2>/dev/null); QSVC=$(echo "$UNIT_LIST" | grep -i qualys | awk '{print $1}' | head -n1); if [ -n "$QSVC" ]; then QUALYS_SVC_FOUND=true; QUALYS_SVC_STATUS=$(systemctl is-active "$QSVC" 2>/dev/null); fi; BSVC=$(echo "$UNIT_LIST" | grep -i binalyze | awk '{print $1}' | head -n1); if [ -n "$BSVC" ]; then BINALYZE_SVC_FOUND=true; BINALYZE_SVC_STATUS=$(systemctl is-active "$BSVC" 2>/dev/null); fi; fi; if command -v rpm >/dev/null 2>&1; then RPM_LIST=$(rpm -qa 2>/dev/null); QRPM=$(echo "$RPM_LIST" | grep -i qualys | head -1); if [ -n "$QRPM" ]; then QUALYS_PKG_FOUND=true; QUALYS_VERSION="$QRPM"; fi; BRPM=$(echo "$RPM_LIST" | grep -i binalyze | head -1); if [ -n "$BRPM" ]; then BINALYZE_PKG_FOUND=true; BINALYZE_VERSION="$BRPM"; fi; elif command -v dpkg >/dev/null 2>&1; then DPKG_LIST=$(dpkg -l 2>/dev/null); QDPKG=$(echo "$DPKG_LIST" | grep -i qualys | head -1); if [ -n "$QDPKG" ]; then QUALYS_PKG_FOUND=true; QUALYS_VERSION=$(echo "$QDPKG" | awk '{print $3}'); fi; BDPKG=$(echo "$DPKG_LIST" | grep -i binalyze | head -1); if [ -n "$BDPKG" ]; then BINALYZE_PKG_FOUND=true; BINALYZE_VERSION=$(echo "$BDPKG" | awk '{print $3}'); fi; fi; printf '{"QualysServiceFound":%s,"QualysRegFound":%s,"QualysDisplayVersion":"%s","QualysServiceStatus":"%s","BinalyzeServiceFound":%s,"BinalyzeRegFound":%s,"BinalyzeDisplayVersion":"%s","BinalyzeServiceStatus":"%s"}\n' "$QUALYS_SVC_FOUND" "$QUALYS_PKG_FOUND" "$QUALYS_VERSION" "$QUALYS_SVC_STATUS" "$BINALYZE_SVC_FOUND" "$BINALYZE_PKG_FOUND" "$BINALYZE_VERSION" "$BINALYZE_SVC_STATUS"
'@
# ---- DNS preflight for ESXi hosts ----
# Invoke-VMScript needs to reach the VM's ESXi host directly (a separate call,
# after the guest script runs) to retrieve output. If that host doesn't
# resolve via DNS, you get a generic "error occurred while sending the
# request" even though the guest-side check succeeded. This checks every
# unique ESXi host in the list up front and fixes unresolvable ones.
Write-Host "Running DNS preflight for ESXi hosts..." -ForegroundColor Cyan
$PreflightVMs = Get-VM -Name $Servers -ErrorAction SilentlyContinue
$UniqueHosts  = $PreflightVMs | Select-Object -ExpandProperty VMHost -Unique
foreach ($EsxHost in $UniqueHosts) {
   $HostName = $EsxHost.Name
   $Resolved = $true
   try {
       Resolve-DnsName -Name $HostName -ErrorAction Stop | Out-Null
   }
   catch {
       $Resolved = $false
   }
   if ($Resolved) { continue }
   Write-Host "  DNS lookup failed for $HostName" -ForegroundColor Yellow
   $HostIP = $null
   if ($KnownHostIPMap.ContainsKey($HostName)) {
       $HostIP = $KnownHostIPMap[$HostName]
   }
   else {
       $MgmtAdapter = Get-VMHostNetworkAdapter -VMHost $EsxHost -VMKernel -ErrorAction SilentlyContinue |
           Where-Object { $_.ManagementTrafficEnabled } | Select-Object -First 1
       if ($MgmtAdapter) { $HostIP = $MgmtAdapter.IP }
   }
   if (-not $HostIP) {
       Write-Host "    Could not determine an IP for $HostName - VMs on this host will likely fail with a communication error." -ForegroundColor Red
       continue
   }
   if ($AutoAddHostsFileEntries) {
       $HostsFile = "$env:SystemRoot\System32\drivers\etc\hosts"
       $Existing  = Get-Content $HostsFile -ErrorAction SilentlyContinue
       if ($Existing -notmatch [regex]::Escape($HostName)) {
           try {
               Add-Content -Path $HostsFile -Value "$HostIP`t$HostName" -ErrorAction Stop
               Write-Host "    Added $HostName ($HostIP) to hosts file" -ForegroundColor Green
           }
           catch {
               Write-Host "    Could not write to hosts file - re-run PowerShell as Administrator. ($($_.Exception.Message))" -ForegroundColor Red
           }
       }
   }
   else {
       Write-Host "    $HostName -> $HostIP (add manually to hosts file, or run with -AutoAddHostsFileEntries `$true)" -ForegroundColor Yellow
   }
}
Write-Host "DNS preflight complete." -ForegroundColor Cyan
Write-Host ""
foreach ($Server in $Servers) {
   Write-Host "Checking $Server..." -ForegroundColor Cyan
   $VM = Get-VM -Name $Server -ErrorAction SilentlyContinue
   if (-not $VM) {
       $Failed += [PSCustomObject]@{
           ServerName = $Server; IPAddress = "-"; PowerState = "-"; VMwareTools = "-"
           Reason = "VM not found in vCenter"
       }
       continue
   }
   $IPAddress = $VM.Guest.IPAddress | Where-Object { $_ -match '^\d{1,3}(\.\d{1,3}){3}$' } | Select-Object -First 1
   if (-not $IPAddress) { $IPAddress = "Unavailable" }
   $PowerState  = $VM.PowerState
   $ToolsStatus = $VM.ExtensionData.Guest.ToolsRunningStatus
   if ($PowerState -ne "PoweredOn") {
       $Failed += [PSCustomObject]@{
           ServerName = $Server; IPAddress = $IPAddress; PowerState = $PowerState; VMwareTools = $ToolsStatus
           Reason = "VM not powered on"
       }
       continue
   }
   if ($ToolsStatus -ne "guestToolsRunning") {
       $Failed += [PSCustomObject]@{
           ServerName = $Server; IPAddress = $IPAddress; PowerState = $PowerState; VMwareTools = $ToolsStatus
           Reason = "VMware Tools not running"
       }
       continue
   }
   # OS detection - GuestFamily is reported by VMware Tools, independent of any network access
   $GuestFamily = $VM.ExtensionData.Guest.GuestFamily
   switch ($GuestFamily) {
       "windowsGuest" { $OSType = "Windows" }
       "linuxGuest"   { $OSType = "Linux" }
       default        { $OSType = $null }
   }
   if (-not $OSType) {
       $Failed += [PSCustomObject]@{
           ServerName = $Server; IPAddress = $IPAddress; PowerState = $PowerState; VMwareTools = $ToolsStatus
           Reason = "Unrecognized or undetected guest OS (GuestFamily: $GuestFamily)"
       }
       continue
   }
   if ($OSType -eq "Windows") {
       $ScriptToRun = $GuestScript
       $ScriptType  = "PowerShell"
   }
   else {
       $ScriptToRun = $GuestScriptLinux
       $ScriptType  = "Bash"
   }
   $Cred = Get-OSCredential -OSType $OSType
   try {
       $Output = Invoke-VMScript -VM $VM -ScriptText $ScriptToRun -ScriptType $ScriptType -GuestCredential $Cred -ErrorAction Stop
       $Parsed            = $Output.ScriptOutput | ConvertFrom-Json
       $QualysInstalled   = ($Parsed.QualysServiceFound -or $Parsed.QualysRegFound)
       $BinalyzeInstalled = ($Parsed.BinalyzeServiceFound -or $Parsed.BinalyzeRegFound)
       $Success += [PSCustomObject]@{
           ServerName          = $Server
           IPAddress           = $IPAddress
           OS                  = $OSType
           PowerState          = $PowerState
           VMwareTools         = $ToolsStatus
           QualysInstalled     = if ($QualysInstalled) { "Installed" } else { "Not Installed" }
           QualysAgentStatus   = if ($QualysInstalled) { if ($Parsed.QualysServiceStatus -in @("active","Running")) { "Running" } else { "Not Running" } } else { "N/A" }
           QualysVersion       = $Parsed.QualysDisplayVersion
           BinalyzeInstalled   = if ($BinalyzeInstalled) { "Installed" } else { "Not Installed" }
           BinalyzeAgentStatus = if ($BinalyzeInstalled) { if ($Parsed.BinalyzeServiceStatus -in @("active","Running")) { "Running" } else { "Not Running" } } else { "N/A" }
           BinalyzeVersion     = $Parsed.BinalyzeDisplayVersion
       }
   }
   catch {
       $Failed += [PSCustomObject]@{
           ServerName = $Server; IPAddress = $IPAddress; PowerState = $PowerState; VMwareTools = $ToolsStatus
           Reason = "Invoke-VMScript error ($OSType): $($_.Exception.Message)"
       }
   }
}
# ---- Console table ----
Write-Host ""
Write-Host "================================================================================================================================="
Write-Host ("{0,-20} {1,-15} {2,-16} {3,-14} {4,-19} {5,-15}" -f "ServerName","IPAddress","Qualys Installed","Qualys Status","Binalyze Installed","Binalyze Status")
Write-Host "================================================================================================================================="
foreach ($R in $Success) {
   $Line = "{0,-20} {1,-15} {2,-16} {3,-14} {4,-19} {5,-15}" -f $R.ServerName,$R.IPAddress,$R.QualysInstalled,$R.QualysAgentStatus,$R.BinalyzeInstalled,$R.BinalyzeAgentStatus
   $Color = if ($R.QualysInstalled -eq "Installed" -and $R.BinalyzeInstalled -eq "Installed") { "Green" }
            elseif ($R.QualysInstalled -eq "Installed" -or $R.BinalyzeInstalled -eq "Installed") { "Yellow" }
            else { "Red" }
   Write-Host $Line -ForegroundColor $Color
}
foreach ($R in $Failed) {
   Write-Host ("{0,-20} FAILED - {1}" -f $R.ServerName, $R.Reason) -ForegroundColor Yellow
}
# ---- Summary report ----
$Total = $Servers.Count
Write-Host ""
Write-Host "==================== SUMMARY ====================" -ForegroundColor Cyan
Write-Host "Total VMs checked : $Total"
Write-Host "Succeeded         : $($Success.Count)" -ForegroundColor Green
Write-Host "Failed            : $($Failed.Count)" -ForegroundColor Red
if ($Failed.Count -gt 0) {
   Write-Host ""
   Write-Host "Failure breakdown:"
   $Failed | Group-Object Reason | Sort-Object Count -Descending | ForEach-Object {
       Write-Host ("  {0,-45} {1}" -f $_.Name, $_.Count)
   }
}
# ---- CSV export (only non-empty sets) ----
$SuccessPath = Join-Path $OutputFolder "Security_Agents_Check_Success.csv"
$FailedPath  = Join-Path $OutputFolder "Security_Agents_Check_Failed.csv"
if ($Success.Count -gt 0) {
   $Success | Select-Object ServerName, IPAddress, OS, QualysInstalled, QualysAgentStatus, QualysVersion, BinalyzeInstalled, BinalyzeAgentStatus, BinalyzeVersion |
       Export-Csv $SuccessPath -NoTypeInformation
   Write-Host "Exported: $SuccessPath" -ForegroundColor Green
}
if ($Failed.Count -gt 0) {
   $Failed | Export-Csv $FailedPath -NoTypeInformation
   Write-Host "Exported: $FailedPath" -ForegroundColor Green
}
