#Requires -Version 5.1
<#
.SYNOPSIS
    Centralised (PowerCLI) front-end for "Check (Kerberos SHA256SHA384SHA512).ps1".

    Collects event-based evidence of the Kerberos ticket encryption types in use, by
    running the Get-WinEvent (Event ID 4768) query inside each Windows VM through
    VMware Tools Guest Operations. Read-only. No WinRM / PsExec / SMB / RDP / guest
    network path.

.DESCRIPTION
    WHAT THE ORIGINAL DID
        Get-WinEvent -FilterHashtable @{LogName='Security'; Id=4768} -MaxEvents 50
        then read TicketEncryptionType / TargetUserName from each event's XML.

    PROBLEM WITH RUNNING THAT BLINDLY ON EVERY SERVER
        Event ID 4768 (Kerberos Authentication Service - TGT / AS-REQ) is written by
        the KDC, i.e. only on a Domain Controller. On an ordinary member server the
        Security log contains no 4768 events and Get-WinEvent throws
        "No events were found". A naive wrapper would mark those member servers
        "failed / non-compliant", which is wrong.

    WHAT THIS SCRIPT DOES INSTEAD (evidence preserved, not replaced by a registry read)
        * Detects the machine's domain role (Win32_ComputerSystem.DomainRole).
        * On a Domain Controller: runs the 4768 query, decodes every
          TicketEncryptionType, groups by encryption type with counts, newest
          timestamp and sample account names, and rates each:
              etype 0x11/0x12 (17/18)  AES128/AES256-SHA1   -> Compliant
              etype 0x13/0x14 (19/20)  AES128/AES256-SHA2   -> Compliant  (these
                                       are the RFC 8009 "SHA256 / SHA384" suites the
                                       requirement is about)
              etype 0x17/0x18 (23/24)  RC4-HMAC             -> Non-Compliant
              etype 0x1 /0x3  (1/3)    DES                  -> Non-Compliant
              etype 0x0 / 0xFFFFFFFF   unknown / failure    -> Unable to Check
          If no 4768 events exist on a DC, that is reported "Unable to Check" with
          the note that the "Audit Kerberos Authentication Service" subcategory is
          probably disabled - NOT as a pass or a fail.
        * On a non-DC (standalone / member server): reports
          "Manual Verification Required" with an explicit explanation that 4768 is a
          KDC-only event and its absence here is expected. It still reads any 4768
          events that happen to be present (e.g. if the host is in fact a promoted
          DC) and reports them.

    Same per-VM validation battery as the other two wrappers. A validation failure
    skips only that VM; one bad VM never stops the batch.

.PARAMETER VMListPath      Text file of VM names (default C:\temp\vmlist.txt). '#' lines ignored.
.PARAMETER CredentialPath  Export-Clixml PSCredential for the guest admin (default C:\temp\wincred.xml). In-memory only.
.PARAMETER OutputPath      Folder on the admin machine for the CSV + log.
.PARAMETER MaxEvents       Max 4768 events to pull per guest (default 200).
.PARAMETER ToolsWaitSecs   Invoke-VMScript VMware Tools wait, seconds (default 180).

.EXAMPLE
    .\Invoke-RemoteKerberosCheck.ps1
    # Run it with a DC-only list for meaningful ticket-encryption evidence:
    .\Invoke-RemoteKerberosCheck.ps1 -VMListPath C:\temp\domaincontrollers.txt

.NOTES
    Reading the Security log requires the guest session to be elevated; VMware Tools
    guest operations run with a full administrator token, so this normally succeeds.
    Prerequisites: PowerCLI imported; already Connect-VIServer'd; vCenter account has
    the "Guest Operation ..." privileges. No vSphere write is performed.
#>

[CmdletBinding()]
param(
    [string] $VMListPath     = 'C:\temp\vmlist.txt',
    [string] $CredentialPath = 'C:\temp\wincred.xml',
    [string] $OutputPath     = (Join-Path $PSScriptRoot 'SecurityGap_Reports'),

    [ValidateRange(1, 5000)]
    [int]    $MaxEvents = 200,

    [ValidateRange(30, 3600)]
    [int]    $ToolsWaitSecs = 180
)

$ErrorActionPreference = 'Stop'
$ProgressPreference     = 'SilentlyContinue'

Write-Host "Invoke-RemoteKerberosCheck.ps1 - event-based Kerberos encryption evidence (read-only)" -ForegroundColor Magenta

# =====================================================================================
# 1. Pre-flight
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

if (-not (Test-Path -LiteralPath $CredentialPath)) { throw "Guest credential file not found: $CredentialPath" }
try {
    $GuestCredential = Import-Clixml -LiteralPath $CredentialPath
} catch {
    throw ("Failed to import guest credential from '$CredentialPath'. Export-Clixml credentials are DPAPI-protected " +
           "and only readable by the same Windows account on the same machine that created them. Recreate with: " +
           "Get-Credential | Export-Clixml '$CredentialPath'. Underlying: $($_.Exception.Message)")
}
if (-not ($GuestCredential -is [pscredential])) { throw "'$CredentialPath' did not deserialise to a PSCredential." }

if (-not (Test-Path -LiteralPath $OutputPath)) { New-Item -ItemType Directory -Path $OutputPath -Force | Out-Null }
$RunStamp = Get-Date -Format 'yyyyMMdd-HHmmss'
$CsvPath  = Join-Path $OutputPath "Kerberos_Check_$RunStamp.csv"
$LogPath  = Join-Path $OutputPath "SecurityGap_$RunStamp.log"

function Write-Log {
    param([string]$Message, [string]$Level = 'INFO')
    $line = "[{0}] [{1}] {2}" -f (Get-Date -Format 'yyyy-MM-dd HH:mm:ss'), $Level, $Message
    Write-Host $line
    Add-Content -LiteralPath $LogPath -Value $line
}
Write-Log "Kerberos check run started. VMs=$($vmNames.Count). MaxEvents=$MaxEvents."

# =====================================================================================
# 2. In-guest READ-ONLY payload
#    {{MAXEVENTS}} is replaced with an integer literal before each run.
# =====================================================================================
$PayloadTemplate = @'
$ProgressPreference = 'SilentlyContinue'
$MaxEvents = {{MAXEVENTS}}

$ComputerName = $env:COMPUTERNAME
try {
    $IPAddresses = (Get-NetIPAddress -AddressFamily IPv4 -ErrorAction Stop |
        Where-Object { $_.InterfaceAlias -notmatch 'Loopback' -and $_.IPAddress -notmatch '^169\.254\.' }).IPAddress -join ', '
} catch { $IPAddresses = $null }
if ([string]::IsNullOrWhiteSpace($IPAddresses)) {
    try { $IPAddresses = ([System.Net.Dns]::GetHostAddresses($ComputerName) | Where-Object { $_.AddressFamily -eq 'InterNetwork' } | Select-Object -First 1).IPAddressToString } catch { $IPAddresses = 'Unknown' }
}
try { $OSCaption = (Get-CimInstance -ClassName Win32_OperatingSystem -ErrorAction Stop).Caption } catch { $OSCaption = 'Unknown' }

try {
    $role = [int](Get-CimInstance -ClassName Win32_ComputerSystem -ErrorAction Stop).DomainRole
} catch { $role = -99 }
$roleName = switch ($role) {
    0 { 'Standalone Workstation' } 1 { 'Member Workstation' } 2 { 'Standalone Server' }
    3 { 'Member Server' } 4 { 'Backup Domain Controller' } 5 { 'Primary Domain Controller' }
    default { "Unknown ($role)" }
}
$isDC = $role -in 4, 5

function ConvertFrom-EType {
    param($Raw)
    if ($null -eq $Raw -or "$Raw" -eq '') { return [pscustomobject]@{ Code = $null; Name = '(none)' } }
    $s = "$Raw".Trim()
    $code = $null
    try {
        if ($s -match '^0x[0-9A-Fa-f]+$') {
            $u = [Convert]::ToUInt32($s, 16)
            $code = if ($u -eq [uint32]4294967295) { -1 } else { [int]$u }
        } elseif ($s -match '^-?\d+$') {
            $code = [int]$s
        }
    } catch { $code = $null }
    if ($null -eq $code) { return [pscustomobject]@{ Code = $null; Name = "(unparsed: $s)" } }
    $name = switch ($code) {
        1  { 'DES-CBC-CRC' }
        3  { 'DES-CBC-MD5' }
        17 { 'AES128-CTS-HMAC-SHA1-96' }
        18 { 'AES256-CTS-HMAC-SHA1-96' }
        19 { 'AES128-CTS-HMAC-SHA256-128 (RFC 8009 / SHA-2)' }
        20 { 'AES256-CTS-HMAC-SHA384-192 (RFC 8009 / SHA-2)' }
        23 { 'RC4-HMAC' }
        24 { 'RC4-HMAC-EXP' }
        0  { 'Unknown / additional pre-auth required (0)' }
        -1 { 'Unknown (0xFFFFFFFF)' }
        default { "Other ($code)" }
    }
    [pscustomobject]@{ Code = $code; Name = $name }
}
function Get-ETypeVerdict {
    param($Code)
    if ($Code -in 17, 18, 19, 20) { 'Compliant' }
    elseif ($Code -in 23, 24)     { 'Non-Compliant' }
    elseif ($Code -in 1, 3)       { 'Non-Compliant' }
    else                          { 'Unable to Check' }
}

$rows = New-Object System.Collections.Generic.List[object]
$queryNote = $null
$eventsReturned = 0
$evts = @()
try {
    $evts = @(Get-WinEvent -FilterHashtable @{ LogName = 'Security'; Id = 4768 } -MaxEvents $MaxEvents -ErrorAction Stop)
} catch [System.UnauthorizedAccessException] {
    $queryNote = "Access denied reading the Security log ($($_.Exception.Message))."
} catch {
    if ($_.Exception.Message -match 'No events were found') {
        $queryNote = 'no-events'
    } else {
        $queryNote = "Get-WinEvent failed: $($_.Exception.Message)"
    }
}

$byType = @{}
foreach ($e in $evts) {
    $eventsReturned++
    $et = $null; $acct = $null
    try {
        $xml  = [xml]$e.ToXml()
        $et   = ($xml.Event.EventData.Data | Where-Object { $_.Name -eq 'TicketEncryptionType' }).'#text'
        $acct = ($xml.Event.EventData.Data | Where-Object { $_.Name -eq 'TargetUserName' }).'#text'
    } catch { }
    $d = ConvertFrom-EType $et
    $key = if ($null -ne $d.Code) { [string]$d.Code } else { 'null' }
    if (-not $byType.ContainsKey($key)) {
        $byType[$key] = [pscustomobject]@{
            Code = $d.Code; Name = $d.Name; Count = 0
            Newest = $e.TimeCreated
            Accounts = New-Object System.Collections.Generic.List[string]
        }
    }
    $slot = $byType[$key]
    $slot.Count++
    if ($e.TimeCreated -gt $slot.Newest) { $slot.Newest = $e.TimeCreated }
    if ($acct -and $slot.Accounts.Count -lt 8 -and -not $slot.Accounts.Contains([string]$acct)) { $slot.Accounts.Add([string]$acct) }
}

$anyRc4OrDes = $false
$anyAes      = $false
$anySha2     = $false
foreach ($k in $byType.Keys) {
    $slot = $byType[$k]
    $verdict = Get-ETypeVerdict $slot.Code
    if ($slot.Code -in 23, 24, 1, 3) { $anyRc4OrDes = $true }
    if ($slot.Code -in 17, 18, 19, 20) { $anyAes = $true }
    if ($slot.Code -in 19, 20) { $anySha2 = $true }
    $rows.Add([pscustomobject]@{
        SecurityCheck = "Kerberos ticket encryption observed: $($slot.Name)"
        EncTypeCode   = $(if ($null -ne $slot.Code) { ('0x{0:X} ({0})' -f $slot.Code) } else { 'unparsed' })
        EventCount    = $slot.Count
        NewestSample  = $(if ($slot.Newest) { $slot.Newest.ToString('yyyy-MM-dd HH:mm:ss') } else { '' })
        SampleAccounts = ($slot.Accounts -join '; ')
        Status        = $verdict
        Detail        = "Event ID 4768, $($slot.Count) occurrence(s), newest $(if ($slot.Newest) { $slot.Newest } else { 'n/a' })."
    }) | Out-Null
}

# Overall / scope row.
if (-not $isDC) {
    $rows.Add([pscustomobject]@{
        SecurityCheck  = 'Kerberos AS-REQ (Event 4768) evidence scope'
        EncTypeCode    = ''
        EventCount     = $eventsReturned
        NewestSample   = ''
        SampleAccounts = ''
        Status         = 'Manual Verification Required'
        Detail         = "DomainRole=$role ($roleName). Event ID 4768 (Kerberos TGT / AS-REQ) is generated by the KDC and is therefore expected only on Domain Controllers. This host is not a DC, so an absence of 4768 events here is expected and is NOT a compliance failure. Run this check against your Domain Controllers to assess Kerberos ticket encryption; on member servers, encryption negotiated for them is visible via Event ID 4769 on the DCs. 4768 events actually present on this host: $eventsReturned."
    }) | Out-Null
    $overall = 'Manual Verification Required'
    $overallDetail = "Not a Domain Controller (role $role / $roleName) - 4768 ticket-encryption evidence is only meaningful on a KDC."
} elseif ($queryNote -eq 'no-events' -or ($eventsReturned -eq 0 -and $null -eq $queryNote)) {
    $overall = 'Unable to Check'
    $overallDetail = "Domain Controller ($roleName) but no Event ID 4768 records are present in the Security log. The 'Audit Kerberos Authentication Service' subcategory is probably disabled, or the log has rolled. Enable it (auditpol /set /subcategory:'Kerberos Authentication Service' /success:enable /failure:enable, or the equivalent Advanced Audit Policy GPO) and re-check. This is neither a pass nor a fail."
    $rows.Add([pscustomobject]@{ SecurityCheck = 'Kerberos AS-REQ (Event 4768) evidence'; EncTypeCode = ''; EventCount = 0; NewestSample = ''; SampleAccounts = ''; Status = $overall; Detail = $overallDetail }) | Out-Null
} elseif ($null -ne $queryNote -and $queryNote -ne 'no-events') {
    $overall = 'Unable to Check'
    $overallDetail = "Domain Controller ($roleName). $queryNote"
    $rows.Add([pscustomobject]@{ SecurityCheck = 'Kerberos AS-REQ (Event 4768) evidence'; EncTypeCode = ''; EventCount = 0; NewestSample = ''; SampleAccounts = ''; Status = $overall; Detail = $overallDetail }) | Out-Null
} else {
    if ($anyRc4OrDes) {
        $overall = 'Non-Compliant'
        $overallDetail = "Domain Controller ($roleName): $eventsReturned event(s) sampled. RC4 and/or DES ticket encryption is still being issued - see the per-type rows. AES-SHA2 (SHA256/SHA384) observed: $anySha2."
    } elseif ($anyAes) {
        $overall = 'Compliant'
        $overallDetail = "Domain Controller ($roleName): $eventsReturned event(s) sampled. Only AES ticket encryption observed. AES-SHA2 (SHA256/SHA384, RFC 8009) observed: $anySha2."
    } else {
        $overall = 'Unable to Check'
        $overallDetail = "Domain Controller ($roleName): $eventsReturned event(s) sampled but none had a decodable AES/RC4/DES TicketEncryptionType (all unknown / pre-auth). Re-sample after normal logon activity."
    }
    $rows.Add([pscustomobject]@{ SecurityCheck = 'Kerberos AS-REQ (Event 4768) overall'; EncTypeCode = ''; EventCount = $eventsReturned; NewestSample = ''; SampleAccounts = ''; Status = $overall; Detail = $overallDetail }) | Out-Null
}

$meta = [PSCustomObject]@{
    Hostname       = $ComputerName
    IPAddress      = $IPAddresses
    OS             = $OSCaption
    DomainRole     = $role
    DomainRoleName = $roleName
    IsDomainController = $isDC
    EventsReturned = $eventsReturned
    QueryNote      = $queryNote
    Overall        = $overall
    OverallDetail  = $overallDetail
}
# .ToArray() rather than @(...) - @() around a generic List instance is unreliable
# on some PowerShell builds; .ToArray() is well-defined on 5.1 and 7.x alike.
$envelope = [PSCustomObject]@{ Schema = 'kerberos-check-1'; Meta = $meta; Rows = $rows.ToArray() }
$json = $envelope | ConvertTo-Json -Depth 8 -Compress
$b64  = [Convert]::ToBase64String([Text.Encoding]::UTF8.GetBytes($json))
Write-Output '<<<KRB-ENVELOPE-B64>>>'
Write-Output $b64
Write-Output '<<<END-KRB-ENVELOPE>>>'
'@

# =====================================================================================
# 3. Admin-side helpers (identical validation to the other wrappers)
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
    $b64 = $sb.ToString()
    if ([string]::IsNullOrWhiteSpace($b64)) { return $null }
    try { return ([Text.Encoding]::UTF8.GetString([Convert]::FromBase64String($b64)) | ConvertFrom-Json) }
    catch { return $null }
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
            Reason = "VM '$Name' was not found in any connected vCenter ($serverList)."; Detected = '0 matches'; VM = $null; Server = $null }
    }
    if ($hits.Count -gt 1) {
        $where = ($hits | ForEach-Object { "$($_.Server.Name):$($_.VM.Id)" }) -join '; '
        return [pscustomobject]@{ Ok = $false; Stage = 'VM lookup'; ServerName = $hits[0].Server.Name
            Reason = "VM name '$Name' is ambiguous - $($hits.Count) matching VMs ($where). Skipped for safety."; Detected = "$($hits.Count) matches"; VM = $null; Server = $null }
    }
    $vm = $hits[0].VM; $srv = $hits[0].Server
    if ($vm.PowerState -ne 'PoweredOn') {
        return [pscustomobject]@{ Ok = $false; Stage = 'Power state'; ServerName = $srv.Name
            Reason = "VM is not powered on (PowerState=$($vm.PowerState))."; Detected = "PowerState=$($vm.PowerState)"; VM = $null; Server = $null }
    }
    $g = $vm.ExtensionData.Guest
    $toolsOk = ($g.ToolsRunningStatus -eq 'guestToolsRunning') -or ($g.ToolsStatus -in 'toolsOk', 'toolsOld')
    if (-not $toolsOk) {
        return [pscustomobject]@{ Ok = $false; Stage = 'VMware Tools'; ServerName = $srv.Name
            Reason = "VMware Tools not installed/running (ToolsStatus=$($g.ToolsStatus), ToolsRunningStatus=$($g.ToolsRunningStatus))."; Detected = "ToolsStatus=$($g.ToolsStatus)"; VM = $null; Server = $null }
    }
    $isWin = ($g.GuestFamily -eq 'windowsGuest') -or ($g.GuestId -match 'windows') -or ($vm.Guest.OSFullName -match 'Windows')
    if (-not $isWin) {
        return [pscustomobject]@{ Ok = $false; Stage = 'Guest OS'; ServerName = $srv.Name
            Reason = "Guest OS is not Windows (GuestFamily=$($g.GuestFamily), OS=$($vm.Guest.OSFullName))."; Detected = "$($vm.Guest.OSFullName)"; VM = $null; Server = $null }
    }
    return [pscustomobject]@{ Ok = $true; Stage = 'OK'; ServerName = $srv.Name; Reason = ''; Detected = ''; VM = $vm; Server = $srv }
}

function Test-GuestPowerShell {
    param($VM, $Server, [pscredential]$Credential, [int]$ToolsWaitSecs)
    $probe = @'
try {
    $id = [Security.Principal.WindowsIdentity]::GetCurrent()
    $pr = New-Object Security.Principal.WindowsPrincipal($id)
    $o = [pscustomobject]@{ PS = $PSVersionTable.PSVersion.ToString(); User = $id.Name; Elevated = [bool]$pr.IsInRole([Security.Principal.WindowsBuiltInRole]::Administrator); Host = $env:COMPUTERNAME }
    Write-Output ('<<<PROBE>>>' + ($o | ConvertTo-Json -Compress) + '<<<ENDPROBE>>>')
} catch {
    Write-Output ('<<<PROBE-ERR>>>' + $_.Exception.Message + '<<<ENDPROBE-ERR>>>')
}
'@
    try {
        $r = Invoke-VMScript -VM $VM -Server $Server -GuestCredential $Credential -ScriptText $probe `
                             -ScriptType Powershell -ToolsWaitSecs $ToolsWaitSecs -Confirm:$false -ErrorAction Stop
    } catch {
        $m = $_.Exception.Message
        $reason = switch -Regex ($m) {
            'InvalidGuestLogin|authentication fail|incorrect user name or password|InvalidLogin|guest permissions|permission to perform this operation' {
                "Guest authentication failed - VMware Tools rejected the supplied credential ('$($Credential.UserName)'). Underlying: $m"; break }
            'Tools are not running|GuestOperationsUnavailable|VMware Tools is not|not installed|VIX' { "VMware Tools guest operations unavailable: $m"; break }
            'timed out|timeout' { "Timed out waiting for VMware Tools / guest response: $m"; break }
            default { "Guest PowerShell probe failed: $m" }
        }
        return [pscustomobject]@{ Ok = $false; Reason = $reason; PSVersion = ''; Elevated = $false; Hostname = '' }
    }
    $out = [string]$r.ScriptOutput
    if ($out -match '<<<PROBE>>>(.*?)<<<ENDPROBE>>>') {
        try { $p = $Matches[1] | ConvertFrom-Json }
        catch { return [pscustomobject]@{ Ok = $false; Reason = "Guest probe returned unreadable JSON: $($Matches[1])"; PSVersion = ''; Elevated = $false; Hostname = '' } }
        try { $ver = [version]$p.PS } catch { $ver = [version]'0.0' }
        if ($ver -lt [version]'5.1') {
            return [pscustomobject]@{ Ok = $false; Reason = "Guest PowerShell is $($p.PS); 5.1+ required."; PSVersion = $p.PS; Elevated = [bool]$p.Elevated; Hostname = $p.Host }
        }
        return [pscustomobject]@{ Ok = $true; Reason = ''; PSVersion = $p.PS; Elevated = [bool]$p.Elevated; Hostname = $p.Host }
    }
    if ($out -match '<<<PROBE-ERR>>>(.*?)<<<ENDPROBE-ERR>>>') {
        return [pscustomobject]@{ Ok = $false; Reason = "Guest probe raised: $($Matches[1])"; PSVersion = ''; Elevated = $false; Hostname = '' }
    }
    return [pscustomobject]@{ Ok = $false; Reason = "Guest probe produced no recognisable output. Raw: $((($out -replace '\s+', ' ').Trim()))"; PSVersion = ''; Elevated = $false; Hostname = '' }
}

# See Invoke-RemoteSecurityGapRemediation.ps1 for the full rationale. A single
# Invoke-VMScript carrying the whole payload fails in some environments with the
# misleading "Could not locate Powershell script interpreter" error (the script is
# placed on a command line that overruns a guest-ops length limit). Primary method
# here: Copy-VMGuestFile (one transfer, no limit); fallback: write Base64 into a
# guest file in small pieces, then decode in-guest. Guest operations only - no
# WinRM/PsExec/SMB/ESXi-or-guest network path. Staged file holds no credential and
# is always deleted.
function Invoke-LargeGuestPayload {
    param(
        $VM, $Server, [pscredential]$Credential,
        [string]$PayloadText, [int]$ToolsWaitSecs,
        [int]$ChunkSize = 1800
    )
    $tag    = [guid]::NewGuid().ToString('N').Substring(0, 12)
    $gB64   = "C:\Windows\Temp\sg$tag.b64"
    $gPs1   = "C:\Windows\Temp\sg$tag.ps1"
    $b64    = [Convert]::ToBase64String([Text.Encoding]::UTF8.GetBytes($PayloadText))
    $staged = $false

    try {
        if (Get-Command Copy-VMGuestFile -ErrorAction SilentlyContinue) {
            $local = Join-Path $env:TEMP "sg$tag.ps1"
            try {
                [System.IO.File]::WriteAllText($local, $PayloadText, (New-Object System.Text.UTF8Encoding($false)))
                Copy-VMGuestFile -Source $local -Destination $gPs1 -VM $VM -Server $Server `
                                 -GuestCredential $Credential -LocalToGuest -Force -ErrorAction Stop | Out-Null
                $staged = $true
                Write-Log "  staged payload into guest via Copy-VMGuestFile." 'INFO'
            } catch {
                Write-Log "  Copy-VMGuestFile not usable ($(($_.Exception.Message -split "`n")[0].Trim())); using chunked staging." 'INFO'
            } finally {
                Remove-Item -LiteralPath $local -Force -ErrorAction SilentlyContinue
            }
        }

        if (-not $staged) {
            $total = [Math]::Ceiling($b64.Length / $ChunkSize)
            Write-Log "  staging payload into guest in $total chunk(s) of <= $ChunkSize chars." 'INFO'
            for ($i = 0; $i -lt $total; $i++) {
                $part = $b64.Substring($i * $ChunkSize, [Math]::Min($ChunkSize, $b64.Length - ($i * $ChunkSize)))
                $verb = if ($i -eq 0) { 'Set-Content' } else { 'Add-Content' }
                $st = "$verb -LiteralPath '$gB64' -Value '$part'"
                Invoke-VMScript -VM $VM -Server $Server -GuestCredential $Credential -ScriptText $st `
                                -ScriptType Powershell -ToolsWaitSecs $ToolsWaitSecs -Confirm:$false -ErrorAction Stop | Out-Null
            }
            $decoder = @"
`$raw = [System.IO.File]::ReadAllText('$gB64')
`$bytes = [Convert]::FromBase64String(( `$raw -replace '\s','' ))
[System.IO.File]::WriteAllText('$gPs1', [System.Text.Encoding]::UTF8.GetString(`$bytes))
"@
            Invoke-VMScript -VM $VM -Server $Server -GuestCredential $Credential -ScriptText $decoder `
                            -ScriptType Powershell -ToolsWaitSecs $ToolsWaitSecs -Confirm:$false -ErrorAction Stop | Out-Null
        }

        $runner = "& powershell.exe -NonInteractive -NoProfile -ExecutionPolicy Bypass -File '$gPs1' 2>&1"
        $res = Invoke-VMScript -VM $VM -Server $Server -GuestCredential $Credential -ScriptText $runner `
                               -ScriptType Powershell -ToolsWaitSecs $ToolsWaitSecs -Confirm:$false -ErrorAction Stop
        return [string]$res.ScriptOutput
    }
    finally {
        $cleanup = "Remove-Item -LiteralPath '$gB64','$gPs1' -Force -ErrorAction SilentlyContinue"
        try {
            Invoke-VMScript -VM $VM -Server $Server -GuestCredential $Credential -ScriptText $cleanup `
                            -ScriptType Powershell -ToolsWaitSecs $ToolsWaitSecs -Confirm:$false -ErrorAction SilentlyContinue | Out-Null
        } catch { }
    }
}

function New-CentralRow {
    param($vCenter, $VMName, $GuestHostname, $IPAddress, $OS, $Category, $SecurityCheck, $DetectedValue, $ExpectedValue, $Status, $Action, $RequiresReboot, $RiskNote, $Reason)
    [PSCustomObject]([ordered]@{
        vCenter = $vCenter; VMName = $VMName; GuestHostname = $GuestHostname; IPAddress = $IPAddress; OS = $OS
        Category = $Category; SecurityCheck = $SecurityCheck; DetectedValue = $DetectedValue; ExpectedValue = $ExpectedValue
        Status = $Status; Action = $Action; RequiresReboot = $RequiresReboot; RiskNote = $RiskNote
        'Error/Reason' = $Reason; Timestamp = (Get-Date -Format 'yyyy-MM-dd HH:mm:ss')
    })
}

$ExpectedEnc = 'AES128/AES256 (etype 0x11/0x12) or AES-SHA2 (etype 0x13/0x14); no RC4 (0x17) or DES (0x1/0x3)'

# =====================================================================================
# 4. Per-VM processing
# =====================================================================================
$centralRows = New-Object System.Collections.Generic.List[object]
$summary = [ordered]@{ Total = 0; Checked = 0; Compliant = 0; NonCompliant = 0; Manual = 0; Unable = 0; SkippedFailed = 0 }
$inventory = Get-VMInventory -Servers $connectedServers
$guestPayload = $PayloadTemplate.Replace('{{MAXEVENTS}}', [string][int]$MaxEvents)

foreach ($vmName in $vmNames) {
    $summary.Total++
    Write-Log "=== Kerberos check for VM '$vmName' ===" 'INFO'
    try {
        $val = Resolve-AndValidateVM -Name $vmName -Inventory $inventory
        if (-not $val.Ok) {
            $summary.SkippedFailed++
            $centralRows.Add((New-CentralRow -vCenter $val.ServerName -VMName $vmName -GuestHostname '' -IPAddress '' -OS '' `
                -Category 'Kerberos' -SecurityCheck $val.Stage -DetectedValue $val.Detected -ExpectedValue '' `
                -Status 'Unable to Check' -Action 'Skipped - VM not processed' -RequiresReboot $false -RiskNote '' -Reason $val.Reason))
            Write-Log "SKIP '$vmName' [$($val.Stage)]: $($val.Reason)" 'WARN'
            continue
        }
        $vm = $val.VM; $srv = $val.Server

        $probe = Test-GuestPowerShell -VM $vm -Server $srv -Credential $GuestCredential -ToolsWaitSecs $ToolsWaitSecs
        if (-not $probe.Ok) {
            $summary.SkippedFailed++
            $centralRows.Add((New-CentralRow -vCenter $srv.Name -VMName $vmName -GuestHostname $vm.Guest.HostName -IPAddress ($vm.Guest.IPAddress -join ',') -OS $vm.Guest.OSFullName `
                -Category 'Kerberos' -SecurityCheck 'Guest authentication / PowerShell' -DetectedValue '' -ExpectedValue 'PowerShell 5.1+ reachable via VMware Tools' `
                -Status 'Unable to Check' -Action 'Skipped - VM not processed' -RequiresReboot $false -RiskNote '' -Reason $probe.Reason))
            Write-Log "SKIP '$vmName' [probe]: $($probe.Reason)" 'WARN'
            continue
        }
        if (-not $probe.Elevated) {
            Write-Log "'$vmName': guest session not elevated - reading the Security log may fail; proceeding and reporting per-item." 'WARN'
        }

        Write-Log "'$vmName': staging + running Kerberos 4768 query payload." 'INFO'
        $scriptOutput = Invoke-LargeGuestPayload -VM $vm -Server $srv -Credential $GuestCredential -PayloadText $guestPayload -ToolsWaitSecs $ToolsWaitSecs

        $envelope = Read-EnvelopeFromScriptOutput -Output $scriptOutput -StartMarker '<<<KRB-ENVELOPE-B64>>>' -EndMarker '<<<END-KRB-ENVELOPE>>>'
        if ($null -eq $envelope) {
            $summary.SkippedFailed++
            $snippet = if ($scriptOutput) { ($scriptOutput -replace '\s+', ' ').Trim() } else { '(no output)' }
            if ($snippet.Length -gt 600) { $snippet = $snippet.Substring(0, 600) + '...' }
            $centralRows.Add((New-CentralRow -vCenter $srv.Name -VMName $vmName -GuestHostname $probe.Hostname -IPAddress ($vm.Guest.IPAddress -join ',') -OS $vm.Guest.OSFullName `
                -Category 'Kerberos' -SecurityCheck 'Guest payload result' -DetectedValue '' -ExpectedValue 'Base64 result envelope' `
                -Status 'Unable to Check' -Action 'Skipped - unparseable result' -RequiresReboot $false -RiskNote '' -Reason "No parseable result envelope. Output start: $snippet"))
            Write-Log "'$vmName': no parseable envelope returned." 'ERROR'
            continue
        }

        $meta  = $envelope.Meta
        $ghost = if ($meta.Hostname)  { $meta.Hostname }  else { $vm.Guest.HostName }
        $gip   = if ($meta.IPAddress) { $meta.IPAddress } else { ($vm.Guest.IPAddress -join ',') }
        $gos   = if ($meta.OS)        { $meta.OS }        else { $vm.Guest.OSFullName }
        $roleNote = "DomainRole=$($meta.DomainRole) ($($meta.DomainRoleName)); IsDC=$($meta.IsDomainController); 4768 events sampled=$($meta.EventsReturned)"

        $summary.Checked++
        foreach ($r in @($envelope.Rows)) {
            $detected = @()
            if ($r.EncTypeCode)         { $detected += "etype=$($r.EncTypeCode)" }
            if ($null -ne $r.EventCount) { $detected += "count=$($r.EventCount)" }
            if ($r.NewestSample)        { $detected += "newest=$($r.NewestSample)" }
            if ($r.SampleAccounts) { $detected += "accounts=$($r.SampleAccounts)" }
            $detectedStr = ($detected -join '; ')
            if (-not $detectedStr) { $detectedStr = $roleNote }

            $centralRows.Add((New-CentralRow -vCenter $srv.Name -VMName $vmName -GuestHostname $ghost -IPAddress $gip -OS $gos `
                -Category 'Kerberos' -SecurityCheck $r.SecurityCheck -DetectedValue $detectedStr -ExpectedValue $ExpectedEnc `
                -Status $r.Status -Action 'Event-based verification (read-only)' -RequiresReboot $false -RiskNote $roleNote -Reason $r.Detail))
        }

        switch ($meta.Overall) {
            'Compliant'                    { $summary.Compliant++ }
            'Non-Compliant'                { $summary.NonCompliant++ }
            'Manual Verification Required' { $summary.Manual++ }
            default                        { $summary.Unable++ }
        }
        Write-Log "'$vmName': $roleNote -> overall $($meta.Overall). $($meta.OverallDetail)" 'INFO'
    }
    catch {
        $summary.SkippedFailed++
        Write-Log "'$vmName': unhandled error - $($_.Exception.Message)" 'ERROR'
        try {
            $centralRows.Add((New-CentralRow -vCenter '' -VMName $vmName -GuestHostname '' -IPAddress '' -OS '' `
                -Category 'Kerberos' -SecurityCheck 'Processing' -DetectedValue '' -ExpectedValue '' `
                -Status 'Unable to Check' -Action 'Skipped - exception' -RequiresReboot $false -RiskNote '' -Reason $_.Exception.Message))
        } catch { }
        continue
    }
}

# =====================================================================================
# 5. Central report + console summary
# =====================================================================================
$centralRows | Export-Csv -LiteralPath $CsvPath -NoTypeInformation -Encoding UTF8
Write-Log "CSV written: $CsvPath ($($centralRows.Count) rows)."

Write-Host ""
Write-Host "============= KERBEROS CHECK SUMMARY =============" -ForegroundColor Green
Write-Host ("  Total VMs                          : {0}" -f $summary.Total)
Write-Host ("  Successfully Checked               : {0}" -f $summary.Checked)
Write-Host ("  Compliant (AES only, DC evidence)  : {0}" -f $summary.Compliant)
Write-Host ("  Non-Compliant (RC4/DES observed)   : {0}" -f $summary.NonCompliant) -ForegroundColor $(if ($summary.NonCompliant) { 'Red' } else { 'Gray' })
Write-Host ("  Manual Verification Required       : {0}" -f $summary.Manual) -ForegroundColor DarkYellow
Write-Host ("  Unable to Check                    : {0}" -f $summary.Unable) -ForegroundColor $(if ($summary.Unable) { 'Yellow' } else { 'Gray' })
Write-Host ("  Skipped / Failed to process        : {0}" -f $summary.SkippedFailed) -ForegroundColor $(if ($summary.SkippedFailed) { 'Red' } else { 'Gray' })
Write-Host "================================================" -ForegroundColor Green
Write-Host ""
Write-Host "Report : $CsvPath"
Write-Host "Log    : $LogPath"
Write-Host "Note   : 'Manual Verification Required' for a non-DC is expected - Event 4768 is a KDC-only event." -ForegroundColor DarkYellow
Write-Log "Kerberos check run complete."
