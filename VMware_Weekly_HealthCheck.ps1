#Requires -Version 5.1
<#
.SYNOPSIS
    VMware Weekly Health Check - Read-only automated collection and DOCX report generation,
    built around the "VMware Weekly Health Check - Full Checklist & Structure".

.DESCRIPTION
    Replaces the manual weekly vCenter/ESXi health-check reports with a single read-only
    PowerCLI collection run from an already-authenticated laptop session.

    - Uses whatever vCenter sessions are already connected (Connect-VIServer done beforehand).
    - Auto-discovers clusters, hosts, datastores, vSAN, vDS/port groups and VMs per vCenter -
      nothing about the infrastructure is hard-coded.
    - A vCenter with multiple clusters/sites is discovered and reported on per-cluster
      automatically - no hard-coded cluster list.
    - A failure collecting one item is logged and skipped; the script always continues.
    - NEVER writes/modifies anything in vCenter. Every cmdlet used below is read-only
      (Get-*, no Set-*/New-*/Remove-*, no config changes, no SSH enablement).

    Report follows the checklist's 10-part structure exactly:
      1. Executive Summary   6. VM Fleet
      2. vCenter & Hosts      7. Security
      3. Cluster (DRS/HA)     8. Backup & Licensing
      4. Storage              9. Alarms
      5. Networking          10. Action Plan

    Some checklist items (storage-array controller/disk/power hardware, and per-VM backup
    success/failure from a backup product) have NO vendor-neutral vCenter/PowerCLI API.
    Those items are always emitted as Status = 'Manual/External Required' with a Notes
    explanation, rather than fabricated - see NOTES below.

.NOTES
    Every threshold below (capacity %, latency, cert/license expiry windows, snapshot age,
    powered-off aging, DRS imbalance, action-plan SLA days) is a CONFIGURABLE PARAMETER, not
    an assumed company standard - raw values/status are always shown alongside any computed
    flag.

    Data NOT available from vCenter/PowerCLI (by design, not an oversight):
      - Physical storage-array hardware health (controllers/disks/power) - vendor-specific
        (NetApp/Dell/HPE/etc.) API, not exposed through vCenter.
      - Per-VM backup job success/failure and backup-policy coverage - vendor-specific
        (Veeam/Cohesity/Commvault/etc.) API, not exposed through vCenter, UNLESS the backup
        product writes its status into a vCenter Custom Attribute, in which case this script
        will pick that up automatically and report on it.
      - "Action Plan" ownership/assignment - vCenter has no concept of who owns a fix; the
        Owner column is intentionally left blank for manual assignment.

.EXAMPLE
    # Already connected: Connect-VIServer sf-vc.sixflags.local, tb-vc.aq.local, amc-vc.amc.local
    .\VMware_Weekly_HealthCheck.ps1 -SiteMap @{'sf-vc.sixflags.local'='SF'; 'tb-vc.aq.local'='TB'; 'amc-vc.amc.local'='AMC'}
#>

[CmdletBinding()]
param(
    # Maps a connected vCenter server (Name as shown in $global:DefaultVIServers) to a friendly
    # site label used in the report. If a connected vCenter isn't in this map, its own server
    # name is used as the label - nothing is hard-coded or required.
    [hashtable]$SiteMap = @{},

    [string]$OutputPath = (Join-Path $PSScriptRoot "VMware_HealthCheck_Reports"),

    [string]$PreparedBy = 'VMware Weekly Health Check Automation',
    [string]$PreparedByTitle = '',
    [string]$CompanyName = '',

    # Configurable thresholds - NOT vendor/company-defined standards. Raw values are always
    # shown regardless of these; these only drive the Warning/Critical flag shown alongside them.
    [double]$CapacityWarningPct       = 80,
    [double]$CapacityCriticalPct      = 90,
    [double]$DatastoreLatencyWarningMs  = 15,
    [double]$DatastoreLatencyCriticalMs = 30,
    [double]$DrsImbalanceWarningPct   = 20,
    [int]$CertExpiryWarningDays       = 60,
    [int]$LicenseExpiryWarningDays    = 30,
    [int]$SnapshotAgeWarningDays      = 3,
    [int]$PoweredOffAgingDays         = 30,

    # Action Plan suggested SLA windows (calendar days from report date) - suggested only,
    # not a company-mandated SLA; edit per your change-management process.
    [int]$CriticalRemediationDays = 3,
    [int]$WarningRemediationDays  = 14,

    # How many Top Risks to surface in the Executive Summary.
    [int]$TopRisksCount = 5,

    # Historical performance window for capacity stats (hours), matches the 24-72h window used
    # in the manual reports.
    [int]$PerfHistoryHours = 24
)

$ErrorActionPreference = 'Stop'
$ScriptStart = Get-Date
$RunDate     = $ScriptStart.ToString('yyyy-MM-dd')
if (-not (Test-Path $OutputPath)) { New-Item -ItemType Directory -Path $OutputPath -Force | Out-Null }

# ============================================================================
# 0. LOGGING / SAFE-EXECUTION HELPERS
# ============================================================================
$Global:FailureLog = [System.Collections.Generic.List[object]]::new()
$Global:AllResults = [System.Collections.Generic.List[object]]::new()
$Global:ObjectClusterMap = @{}   # maps a host/datastore/vDS/port-group/VM name -> the cluster it
                                  # belongs to, used only when building the per-site report.
$Global:SiteClusterMap = @{}     # maps a Site label -> ordered list of cluster names discovered.

function Write-CheckLog {
    param([string]$VCenter, [string]$Site, [string]$Object, [string]$CheckName, [string]$ErrorMessage)
    $Global:FailureLog.Add([pscustomobject]@{
        Timestamp = (Get-Date).ToString('s')
        VCenter   = $VCenter
        Site      = $Site
        Object    = $Object
        Check     = $CheckName
        Error     = $ErrorMessage
    })
    Write-Warning "[$Site/$VCenter] $CheckName on '$Object' failed: $ErrorMessage"
}

# Runs a scriptblock; on failure logs it and returns $null instead of stopping the run.
function Invoke-SafeCheck {
    param(
        [Parameter(Mandatory)][scriptblock]$Script,
        [Parameter(Mandatory)][string]$CheckName,
        [string]$VCenter = 'n/a',
        [string]$Site = 'n/a',
        [string]$ObjectName = 'n/a'
    )
    try {
        & $Script
    } catch {
        Write-CheckLog -VCenter $VCenter -Site $Site -Object $ObjectName -CheckName $CheckName -ErrorMessage $_.Exception.Message
        return $null
    }
}

function New-Finding {
    param(
        [string]$Site, [string]$VCenter, [string]$Area, [string]$Object,
        [string]$Item, [string]$Value, [string]$Status, [string]$Notes = ''
    )
    # Area must be one of the 8 checklist categories:
    #   'vCenter & Host Health' | 'Cluster (DRS/HA)' | 'Storage' | 'Networking' |
    #   'VM Fleet' | 'Security' | 'Backup & Licensing' | 'Alarms'
    # Status must be one of: Healthy / Warning / Critical / Information / Unable to Check / Manual/External Required
    $obj = [pscustomobject]@{
        Site    = $Site
        VCenter = $VCenter
        Area    = $Area
        Object  = $Object
        Item    = $Item
        Value   = $Value
        Status  = $Status
        Notes   = $Notes
    }
    $Global:AllResults.Add($obj)
    return $obj
}

function Get-PctStatus {
    param([double]$Pct)
    if ($Pct -ge $CapacityCriticalPct) { return 'Critical' }
    elseif ($Pct -ge $CapacityWarningPct) { return 'Warning' }
    else { return 'Healthy' }
}

# Reads the leaf TLS certificate of a remote endpoint (read-only handshake, no data sent/changed)
# so certificate expiry can be checked without needing a separate VAMI/CIS session.
function Get-RemoteCertExpiry {
    param(
        [Parameter(Mandatory)][string]$ComputerName,
        [int]$Port = 443,
        [int]$TimeoutMs = 5000
    )
    $tcp = New-Object System.Net.Sockets.TcpClient
    $ssl = $null
    try {
        $iar = $tcp.BeginConnect($ComputerName, $Port, $null, $null)
        if (-not $iar.AsyncWaitHandle.WaitOne($TimeoutMs)) {
            throw "Connection to $ComputerName`:$Port timed out."
        }
        $tcp.EndConnect($iar)
        $callback = [System.Net.Security.RemoteCertificateValidationCallback]{ param($se,$cert,$chain,$errs) $true }
        $ssl = New-Object System.Net.Security.SslStream($tcp.GetStream(), $false, $callback)
        $ssl.AuthenticateAsClient($ComputerName)
        $cert = [System.Security.Cryptography.X509Certificates.X509Certificate2]$ssl.RemoteCertificate
        [pscustomobject]@{
            Subject   = $cert.Subject
            NotBefore = $cert.NotBefore
            NotAfter  = $cert.NotAfter
        }
    } finally {
        if ($ssl) { $ssl.Dispose() }
        $tcp.Close()
    }
}

# ============================================================================
# 1. PREREQUISITES / CONNECTED SESSION DISCOVERY
# ============================================================================
if (-not (Get-Module -ListAvailable -Name VMware.PowerCLI, VMware.VimAutomation.Core |
        Where-Object { $_.Name -eq 'VMware.VimAutomation.Core' })) {
    throw "VMware.PowerCLI is not installed. Run: Install-Module VMware.PowerCLI -Scope CurrentUser"
}
Import-Module VMware.VimAutomation.Core -ErrorAction SilentlyContinue

$VsanModuleAvailable = [bool](Get-Module -ListAvailable -Name VMware.VimAutomation.Vsan)
if ($VsanModuleAvailable) { Import-Module VMware.VimAutomation.Vsan -ErrorAction SilentlyContinue }

$Connections = $global:DefaultVIServers | Where-Object { $_.IsConnected }
if (-not $Connections -or $Connections.Count -eq 0) {
    throw "No connected vCenter sessions found. Connect first, e.g.:`n  Connect-VIServer sf-vc.sixflags.local, tb-vc.aq.local, amc-vc.amc.local"
}

Write-Host "Connected vCenter sessions: $($Connections.Name -join ', ')" -ForegroundColor Cyan

# ============================================================================
# 2. PER-VCENTER / PER-SITE COLLECTION
# ============================================================================
foreach ($VC in $Connections) {

    $VCName = $VC.Name
    $Site   = if ($SiteMap.ContainsKey($VCName)) { $SiteMap[$VCName] } else { $VCName }

    Write-Host "`n=== Collecting: $Site ($VCName) ===" -ForegroundColor Green

    # ------------------------------------------------------------------
    # 2.1 vCenter & Host Health - vCenter appliance level
    # ------------------------------------------------------------------
    Invoke-SafeCheck -CheckName 'vCenter version/build' -VCenter $VCName -Site $Site -ObjectName $VCName -Script {
        New-Finding -Site $Site -VCenter $VCName -Area 'vCenter & Host Health' -Object $VCName `
            -Item 'vCenter Version/Build' -Value "$($VC.Version) (Build $($VC.Build))" -Status 'Information'
    }

    # Licensing (vCenter/ESXi licenses via LicenseManager - reliably exposed via API, no CIS needed)
    Invoke-SafeCheck -CheckName 'Licensing' -VCenter $VCName -Site $Site -ObjectName $VCName -Script {
        $lm  = Get-View -Server $VC ($VC.ExtensionData.Content.LicenseManager)
        foreach ($lic in $lm.Licenses) {
            $expProp = $lic.Properties | Where-Object { $_.Key -eq 'expirationDate' }
            $status  = 'Information'
            $expStr  = 'No expiration / not set'
            if ($expProp) {
                $expDate = [datetime]$expProp.Value
                $expStr  = $expDate.ToString('yyyy-MM-dd')
                $daysLeft = ($expDate - (Get-Date)).Days
                if ($daysLeft -le 0) { $status = 'Critical' }
                elseif ($daysLeft -le $LicenseExpiryWarningDays) { $status = 'Warning' }
                else { $status = 'Healthy' }
            }
            New-Finding -Site $Site -VCenter $VCName -Area 'Backup & Licensing' -Object $VCName `
                -Item "License: $($lic.Name)" -Value "Used $($lic.Used)/$($lic.Total) - Expires $expStr" -Status $status
        }
    }

    # vCenter appliance CPU/Mem/Disk/Services/NTP require the VAMI/CIS REST API (Connect-CisServer),
    # a separate authenticated session from the vSphere API. If connected it is used; otherwise
    # this is explicitly marked, not fabricated.
    $CisSession = $global:DefaultCisServers | Where-Object { $_.Name -eq $VCName -and $_.IsConnected }
    if ($CisSession) {
        Invoke-SafeCheck -CheckName 'Appliance health (VAMI)' -VCenter $VCName -Site $Site -ObjectName $VCName -Script {
            $healthSvc = Get-CisService -Name 'com.vmware.appliance.health.system' -Server $CisSession
            $overall = $healthSvc.get()
            New-Finding -Site $Site -VCenter $VCName -Area 'vCenter & Host Health' -Object $VCName `
                -Item 'Overall Appliance Health' -Value $overall -Status $(if ($overall -eq 'green') {'Healthy'} else {'Warning'})
            foreach ($comp in 'cpu','mem','storage') {
                $svc = Get-CisService -Name "com.vmware.appliance.health.$comp" -Server $CisSession
                $val = $svc.get()
                $label = @{ cpu = 'CPU'; mem = 'Memory'; storage = 'Disk Usage' }[$comp]
                New-Finding -Site $Site -VCenter $VCName -Area 'vCenter & Host Health' -Object $VCName `
                    -Item "Appliance $label" -Value $val -Status $(if ($val -eq 'green') {'Healthy'} else {'Warning'})
            }
        }
        Invoke-SafeCheck -CheckName 'Appliance NTP (VAMI)' -VCenter $VCName -Site $Site -ObjectName $VCName -Script {
            $ntpSvc = Get-CisService -Name 'com.vmware.appliance.ntp' -Server $CisSession
            $ntpStatus = $ntpSvc.test()
            New-Finding -Site $Site -VCenter $VCName -Area 'vCenter & Host Health' -Object $VCName `
                -Item 'Appliance NTP Sync' -Value ($ntpStatus | Out-String).Trim() -Status 'Information'
        }
        Invoke-SafeCheck -CheckName 'Appliance services (VAMI)' -VCenter $VCName -Site $Site -ObjectName $VCName -Script {
            $svcListSvc = Get-CisService -Name 'com.vmware.appliance.services' -Server $CisSession
            $services = $svcListSvc.list()
            $notRunning = $services.GetEnumerator() | Where-Object { $_.Value.state -ne 'STARTED' }
            $status = if ($notRunning) { 'Warning' } else { 'Healthy' }
            $val = if ($notRunning) {
                ($notRunning | ForEach-Object { "$($_.Key): $($_.Value.state)" }) -join '; '
            } else {
                "All $($services.Count) services running"
            }
            New-Finding -Site $Site -VCenter $VCName -Area 'vCenter & Host Health' -Object $VCName `
                -Item 'vCenter Services' -Value $val -Status $status
        }
    } else {
        New-Finding -Site $Site -VCenter $VCName -Area 'vCenter & Host Health' -Object $VCName `
            -Item 'Appliance CPU/Memory/Disk/Services/NTP' -Value 'n/a' `
            -Status 'Manual/External Required' `
            -Notes 'Requires a VAMI/CIS session (Connect-CisServer <vcenter>) in addition to the vSphere API session. Not connected in this run.'
    }

    # vCenter TLS certificate expiry (read-only TLS handshake, no CIS session required)
    Invoke-SafeCheck -CheckName 'vCenter TLS certificate' -VCenter $VCName -Site $Site -ObjectName $VCName -Script {
        $cert = Get-RemoteCertExpiry -ComputerName $VCName
        $daysLeft = ($cert.NotAfter - (Get-Date)).Days
        $status = if ($daysLeft -le 0) { 'Critical' } elseif ($daysLeft -le $CertExpiryWarningDays) { 'Warning' } else { 'Healthy' }
        New-Finding -Site $Site -VCenter $VCName -Area 'Security' -Object $VCName `
            -Item 'vCenter TLS Certificate Expiry' -Value ("Expires {0:yyyy-MM-dd} ({1} days)" -f $cert.NotAfter, $daysLeft) -Status $status
    }

    # Active vCenter-level alarms
    Invoke-SafeCheck -CheckName 'vCenter-level alarms' -VCenter $VCName -Site $Site -ObjectName $VCName -Script {
        $rootFolder = Get-View -Server $VC $VC.ExtensionData.Content.RootFolder
        $alarms = $rootFolder.TriggeredAlarmState
        if ($alarms -and $alarms.Count -gt 0) {
            foreach ($a in $alarms) {
                $alarmDef = Get-View -Server $VC $a.Alarm
                New-Finding -Site $Site -VCenter $VCName -Area 'Alarms' -Object $VCName `
                    -Item $alarmDef.Info.Name -Value $a.OverallStatus.ToString() -Status ($a.OverallStatus.ToString().Substring(0,1).ToUpper() + $a.OverallStatus.ToString().Substring(1))
            }
        } else {
            New-Finding -Site $Site -VCenter $VCName -Area 'Alarms' -Object $VCName -Item 'Triggered Alarms' -Value 'None' -Status 'Healthy'
        }
    }

    # --- 2.2 Discover clusters (a multi-cluster vCenter's sites come out of this automatically) ---
    $Clusters = Invoke-SafeCheck -CheckName 'Cluster discovery' -VCenter $VCName -Site $Site -ObjectName $VCName -Script {
        Get-Cluster -Server $VC
    }
    if (-not $Clusters) { continue }

    foreach ($Cluster in $Clusters) {
        $ClusterName = $Cluster.Name
        Write-Host "  - Cluster: $ClusterName"

        if (-not $Global:SiteClusterMap.ContainsKey($Site)) { $Global:SiteClusterMap[$Site] = [System.Collections.Generic.List[string]]::new() }
        $Global:SiteClusterMap[$Site].Add($ClusterName)

        # --- Cluster (DRS/HA) ---------------------------------------------------------------
        Invoke-SafeCheck -CheckName 'HA/DRS status' -VCenter $VCName -Site $Site -ObjectName $ClusterName -Script {
            New-Finding -Site $Site -VCenter $VCName -Area 'Cluster (DRS/HA)' -Object $ClusterName `
                -Item 'vSphere HA' -Value $Cluster.HAEnabled -Status $(if ($Cluster.HAEnabled) {'Healthy'} else {'Warning'})
            New-Finding -Site $Site -VCenter $VCName -Area 'Cluster (DRS/HA)' -Object $ClusterName `
                -Item 'vSphere DRS' -Value $Cluster.DrsEnabled -Status $(if ($Cluster.DrsEnabled) {'Healthy'} else {'Warning'})
        }

        Invoke-SafeCheck -CheckName 'HA admission control' -VCenter $VCName -Site $Site -ObjectName $ClusterName -Script {
            $dasConfig = $Cluster.ExtensionData.Configuration.DasConfig
            if (-not $dasConfig -or -not $dasConfig.Enabled) {
                New-Finding -Site $Site -VCenter $VCName -Area 'Cluster (DRS/HA)' -Object $ClusterName `
                    -Item 'HA Admission Control / Failover Headroom' -Value 'HA not enabled' -Status 'Information'
                return
            }
            $policy = $dasConfig.AdmissionControlPolicy
            $desc = 'n/a'
            if ($policy) {
                switch -Regex ($policy.GetType().Name) {
                    'FailoverResourcesAdmissionControlPolicy' { $desc = "Reserved failover capacity: CPU $($policy.CpuFailoverResourcesPercent)% / Memory $($policy.MemoryFailoverResourcesPercent)%" }
                    'FailoverHostAdmissionControlPolicy'      { $desc = "Dedicated failover host(s) reserved" }
                    'FailoverLevelAdmissionControlPolicy'     { $desc = "Host failures to tolerate: $($policy.FailoverLevel)" }
                    default                                   { $desc = "Policy type: $($policy.GetType().Name)" }
                }
            }
            New-Finding -Site $Site -VCenter $VCName -Area 'Cluster (DRS/HA)' -Object $ClusterName `
                -Item 'HA Admission Control / Failover Headroom' -Value $desc -Status 'Information'
        }

        # Cluster-level alarms
        Invoke-SafeCheck -CheckName 'Cluster alarms' -VCenter $VCName -Site $Site -ObjectName $ClusterName -Script {
            $alarms = $Cluster.ExtensionData.TriggeredAlarmState
            if ($alarms -and $alarms.Count -gt 0) {
                foreach ($a in $alarms) {
                    $alarmDef = Get-View -Server $VC $a.Alarm
                    New-Finding -Site $Site -VCenter $VCName -Area 'Alarms' -Object $ClusterName `
                        -Item $alarmDef.Info.Name -Value $a.OverallStatus.ToString() -Status ($a.OverallStatus.ToString().Substring(0,1).ToUpper() + $a.OverallStatus.ToString().Substring(1))
                }
            } else {
                New-Finding -Site $Site -VCenter $VCName -Area 'Alarms' -Object $ClusterName -Item 'Triggered Alarms' -Value 'None' -Status 'Healthy'
            }
        }

        # --- Hosts in this cluster ---
        $Hosts = Invoke-SafeCheck -CheckName 'Host discovery' -VCenter $VCName -Site $Site -ObjectName $ClusterName -Script {
            Get-VMHost -Server $VC -Location $Cluster
        }
        if (-not $Hosts) { continue }

        $HostVersions = @()

        foreach ($VMHost in $Hosts) {
            $HName = $VMHost.Name
            $Global:ObjectClusterMap[$HName] = $ClusterName
            $HostVersions += "$($VMHost.Version) build $($VMHost.Build)"

            # Connection/power state
            Invoke-SafeCheck -CheckName 'Host connection state' -VCenter $VCName -Site $Site -ObjectName $HName -Script {
                $ok = $VMHost.ConnectionState -eq 'Connected'
                New-Finding -Site $Site -VCenter $VCName -Area 'vCenter & Host Health' -Object $HName `
                    -Item 'Connection State' -Value $VMHost.ConnectionState -Status $(if ($ok) {'Healthy'} else {'Critical'})
            }

            # Version/build (for consistency check, per cluster, below)
            Invoke-SafeCheck -CheckName 'Host version' -VCenter $VCName -Site $Site -ObjectName $HName -Script {
                New-Finding -Site $Site -VCenter $VCName -Area 'vCenter & Host Health' -Object $HName `
                    -Item 'ESXi Version/Build' -Value "$($VMHost.Version) (Build $($VMHost.Build))" -Status 'Information'
            }

            # Per-host CPU/Memory/Uptime - explicitly NOT a fleet average, one row per host.
            Invoke-SafeCheck -CheckName 'Per-host CPU/Mem/Uptime' -VCenter $VCName -Site $Site -ObjectName $HName -Script {
                $qs = $VMHost.ExtensionData.Summary.QuickStats
                $cpuTotalMhz = $VMHost.CpuTotalMhz
                $memTotalMB  = $VMHost.MemoryTotalMB
                $cpuPct = if ($qs -and $cpuTotalMhz -gt 0) { ($qs.OverallCpuUsage / $cpuTotalMhz) * 100 } else { $null }
                $memPct = if ($qs -and $memTotalMB -gt 0) { ($qs.OverallMemoryUsage / $memTotalMB) * 100 } else { $null }
                if ($cpuPct -ne $null) {
                    New-Finding -Site $Site -VCenter $VCName -Area 'vCenter & Host Health' -Object $HName `
                        -Item 'CPU Usage %' -Value ("{0:N1}% ({1} / {2} MHz)" -f $cpuPct, $qs.OverallCpuUsage, $cpuTotalMhz) -Status (Get-PctStatus $cpuPct)
                }
                if ($memPct -ne $null) {
                    New-Finding -Site $Site -VCenter $VCName -Area 'vCenter & Host Health' -Object $HName `
                        -Item 'Memory Usage %' -Value ("{0:N1}% ({1} / {2} MB)" -f $memPct, $qs.OverallMemoryUsage, $memTotalMB) -Status (Get-PctStatus $memPct)
                }
                $uptimeDays = if ($qs -and $qs.Uptime) { [Math]::Round($qs.Uptime / 86400, 1) } else { $null }
                New-Finding -Site $Site -VCenter $VCName -Area 'vCenter & Host Health' -Object $HName `
                    -Item 'Uptime (days)' -Value $(if ($uptimeDays -ne $null) { $uptimeDays } else { 'n/a' }) -Status 'Information'
            }

            # Host-level alarms
            Invoke-SafeCheck -CheckName 'Host alarms' -VCenter $VCName -Site $Site -ObjectName $HName -Script {
                $alarms = $VMHost.ExtensionData.TriggeredAlarmState
                if ($alarms -and $alarms.Count -gt 0) {
                    foreach ($a in $alarms) {
                        $alarmDef = Get-View -Server $VC $a.Alarm
                        New-Finding -Site $Site -VCenter $VCName -Area 'Alarms' -Object $HName `
                            -Item $alarmDef.Info.Name -Value $a.OverallStatus.ToString() -Status ($a.OverallStatus.ToString().Substring(0,1).ToUpper() + $a.OverallStatus.ToString().Substring(1))
                    }
                } else {
                    New-Finding -Site $Site -VCenter $VCName -Area 'Alarms' -Object $HName -Item 'Triggered Alarms' -Value 'None' -Status 'Healthy'
                }
            }

            # Hardware health via built-in host Health System (numeric sensors) - no SSH required.
            # Covers PSU, fans, RAID controllers and disks in one pass, grouped by sensor type.
            Invoke-SafeCheck -CheckName 'Hardware sensors' -VCenter $VCName -Site $Site -ObjectName $HName -Script {
                $sensors = $VMHost.ExtensionData.Runtime.HealthSystemRuntime.SystemHealthInfo.NumericSensorInfo
                if (-not $sensors) {
                    New-Finding -Site $Site -VCenter $VCName -Area 'vCenter & Host Health' -Object $HName `
                        -Item 'Hardware Sensors (PSU/Fans/RAID/Disks)' -Value 'n/a' -Status 'Unable to Check' `
                        -Notes 'Host does not expose CIM/IPMI sensor data to vCenter (common on some blade/BMC configs).'
                    return
                }
                $groups = $sensors | Group-Object -Property SensorType
                foreach ($g in $groups) {
                    $bad = $g.Group | Where-Object { $_.HealthState.Key -notin @('green','Green') }
                    $status = if ($bad) { 'Warning' } else { 'Healthy' }
                    $summary = if ($bad) { ($bad | ForEach-Object { "$($_.Name): $($_.HealthState.Label)" }) -join '; ' } else { 'All normal' }
                    New-Finding -Site $Site -VCenter $VCName -Area 'vCenter & Host Health' -Object $HName `
                        -Item "$($g.Name) Sensors" -Value $summary -Status $status
                }
            }

            # TPM
            Invoke-SafeCheck -CheckName 'TPM' -VCenter $VCName -Site $Site -ObjectName $HName -Script {
                $tpmInfo = $VMHost.ExtensionData.Capability.TpmSupported
                New-Finding -Site $Site -VCenter $VCName -Area 'Security' -Object $HName `
                    -Item 'TPM Present/Supported' -Value $tpmInfo -Status $(if ($tpmInfo) {'Healthy'} else {'Information'})
            }

            # Secure Boot (via EsxCli, no SSH required - this is the vSphere API path)
            Invoke-SafeCheck -CheckName 'Secure Boot' -VCenter $VCName -Site $Site -ObjectName $HName -Script {
                $esxcli = Get-EsxCli -VMHost $VMHost -Server $VC -V2
                $sb = $esxcli.system.settings.encryption.get.Invoke()
                $enabled = $sb.RequireSecureBoot
                New-Finding -Site $Site -VCenter $VCName -Area 'Security' -Object $HName `
                    -Item 'Secure Boot' -Value $enabled -Status $(if ($enabled -eq $true -or $enabled -eq 'true') {'Healthy'} else {'Warning'})
            }

            # Lockdown Mode
            Invoke-SafeCheck -CheckName 'Lockdown mode' -VCenter $VCName -Site $Site -ObjectName $HName -Script {
                $lockdown = $VMHost.ExtensionData.Config.LockdownMode
                New-Finding -Site $Site -VCenter $VCName -Area 'Security' -Object $HName `
                    -Item 'Lockdown Mode' -Value $lockdown -Status 'Information'
            }

            # Local ESXi users - via esxcli (API-based, no SSH) rather than Get-VMHostAccount,
            # whose -VMHost parameter isn't present on all PowerCLI versions.
            Invoke-SafeCheck -CheckName 'Local ESXi users' -VCenter $VCName -Site $Site -ObjectName $HName -Script {
                $esxcli = Get-EsxCli -VMHost $VMHost -Server $VC -V2
                $accts = $esxcli.system.account.list.Invoke()
                New-Finding -Site $Site -VCenter $VCName -Area 'Security' -Object $HName `
                    -Item 'Local Accounts' -Value (($accts | ForEach-Object { $_.UserID }) -join ', ') -Status 'Information' `
                    -Notes 'Review membership manually; automation only inventories accounts, it does not judge appropriateness.'
            }

            # Syslog
            Invoke-SafeCheck -CheckName 'Syslog config' -VCenter $VCName -Site $Site -ObjectName $HName -Script {
                $syslog = Get-VMHostSysLogServer -Server $VC -VMHost $VMHost
                $configured = -not [string]::IsNullOrWhiteSpace($syslog.Host)
                New-Finding -Site $Site -VCenter $VCName -Area 'Security' -Object $HName `
                    -Item 'Syslog Target' -Value $(if ($configured) { "$($syslog.Host):$($syslog.Port)" } else { 'Not configured' }) `
                    -Status $(if ($configured) {'Healthy'} else {'Warning'})
            }

            # ESXi TLS certificate expiry (read-only TLS handshake, no SSH/CIS required)
            Invoke-SafeCheck -CheckName 'ESXi TLS certificate' -VCenter $VCName -Site $Site -ObjectName $HName -Script {
                $cert = Get-RemoteCertExpiry -ComputerName $HName
                $daysLeft = ($cert.NotAfter - (Get-Date)).Days
                $status = if ($daysLeft -le 0) { 'Critical' } elseif ($daysLeft -le $CertExpiryWarningDays) { 'Warning' } else { 'Healthy' }
                New-Finding -Site $Site -VCenter $VCName -Area 'Security' -Object $HName `
                    -Item 'ESXi TLS Certificate Expiry' -Value ("Expires {0:yyyy-MM-dd} ({1} days)" -f $cert.NotAfter, $daysLeft) -Status $status
            }

            # NTP
            Invoke-SafeCheck -CheckName 'Host NTP' -VCenter $VCName -Site $Site -ObjectName $HName -Script {
                $ntpServers = Get-VMHostNtpServer -Server $VC -VMHost $VMHost
                $ntpSvc = Get-VMHostService -Server $VC -VMHost $VMHost | Where-Object { $_.Key -eq 'ntpd' }
                $running = $ntpSvc -and $ntpSvc.Running
                New-Finding -Site $Site -VCenter $VCName -Area 'vCenter & Host Health' -Object $HName `
                    -Item 'NTP' -Value "Servers: $($ntpServers -join ', ') | Running: $running" `
                    -Status $(if ($running -and $ntpServers) {'Healthy'} else {'Warning'})
            }

            # Physical NICs / redundancy
            Invoke-SafeCheck -CheckName 'Physical NICs' -VCenter $VCName -Site $Site -ObjectName $HName -Script {
                $pnics = Get-VMHostNetworkAdapter -Server $VC -VMHost $VMHost -Physical
                $up = $pnics | Where-Object { $_.BitRatePerSec -gt 0 }
                New-Finding -Site $Site -VCenter $VCName -Area 'Networking' -Object $HName `
                    -Item 'Physical NICs' -Value "$($pnics.Count) total, $($up.Count) linked up" `
                    -Status $(if ($pnics.Count -ge 2 -and $up.Count -ge 2) {'Healthy'} else {'Warning'})
            }

            # NIC error/drop counters via performance manager (safe, no SSH) - "network errors and
            # dropped packets on uplinks" from the checklist.
            Invoke-SafeCheck -CheckName 'NIC error counters' -VCenter $VCName -Site $Site -ObjectName $HName -Script {
                $stat = Get-Stat -Server $VC -Entity $VMHost -Stat 'net.errorsRx.summation','net.errorsTx.summation','net.droppedRx.summation','net.droppedTx.summation' -Realtime -MaxSamples 1 -ErrorAction Stop
                if ($stat) {
                    $total = ($stat | Measure-Object -Property Value -Sum).Sum
                    New-Finding -Site $Site -VCenter $VCName -Area 'Networking' -Object $HName `
                        -Item 'NIC Errors/Drops (latest sample)' -Value $total -Status $(if ($total -eq 0) {'Healthy'} else {'Warning'})
                } else {
                    New-Finding -Site $Site -VCenter $VCName -Area 'Networking' -Object $HName `
                        -Item 'NIC Errors/Drops' -Value 'n/a' -Status 'Unable to Check' -Notes 'Counter not returned by this host.'
                }
            }
        } # end per-host

        # ESXi version consistency across the cluster
        Invoke-SafeCheck -CheckName 'Version consistency' -VCenter $VCName -Site $Site -ObjectName $ClusterName -Script {
            $distinct = $HostVersions | Select-Object -Unique
            $status = if ($distinct.Count -le 1) { 'Healthy' } else { 'Warning' }
            New-Finding -Site $Site -VCenter $VCName -Area 'vCenter & Host Health' -Object $ClusterName `
                -Item 'ESXi Version Consistency' -Value ($distinct -join ' | ') -Status $status
        }

        # DRS load balance - spread between busiest and idlest host's CPU usage % in the cluster.
        # No VMware-published "imbalance" threshold exists, hence configurable, raw value always shown.
        Invoke-SafeCheck -CheckName 'DRS load balance' -VCenter $VCName -Site $Site -ObjectName $ClusterName -Script {
            if (-not $Cluster.DrsEnabled) {
                New-Finding -Site $Site -VCenter $VCName -Area 'Cluster (DRS/HA)' -Object $ClusterName `
                    -Item 'DRS Load Balance (CPU spread across hosts)' -Value 'DRS not enabled' -Status 'Information'
                return
            }
            $usagePcts = foreach ($h in $Hosts) {
                $qs = $h.ExtensionData.Summary.QuickStats
                if ($qs -and $h.CpuTotalMhz -gt 0) { ($qs.OverallCpuUsage / $h.CpuTotalMhz) * 100 }
            }
            if ($usagePcts) {
                $spread = ($usagePcts | Measure-Object -Maximum).Maximum - ($usagePcts | Measure-Object -Minimum).Minimum
                $status = if ($spread -ge $DrsImbalanceWarningPct) { 'Warning' } else { 'Healthy' }
                New-Finding -Site $Site -VCenter $VCName -Area 'Cluster (DRS/HA)' -Object $ClusterName `
                    -Item 'DRS Load Balance (CPU spread across hosts)' -Value ("{0:N1}% spread (max-min host CPU usage)" -f $spread) -Status $status
            } else {
                New-Finding -Site $Site -VCenter $VCName -Area 'Cluster (DRS/HA)' -Object $ClusterName `
                    -Item 'DRS Load Balance (CPU spread across hosts)' -Value 'n/a' -Status 'Unable to Check'
            }
        }

        # Cluster CPU/Memory capacity - historical stats over the configured window (not instantaneous)
        Invoke-SafeCheck -CheckName 'Cluster CPU/Mem capacity (historical)' -VCenter $VCName -Site $Site -ObjectName $ClusterName -Script {
            $start = (Get-Date).AddHours(-1 * $PerfHistoryHours)
            $cpuStat = Get-Stat -Server $VC -Entity $Cluster -Stat 'cpu.usage.average' -Start $start -Finish (Get-Date) -ErrorAction Stop
            $memStat = Get-Stat -Server $VC -Entity $Cluster -Stat 'mem.usage.average' -Start $start -Finish (Get-Date) -ErrorAction Stop

            $cpuTotalMHz = ($Hosts | Measure-Object -Property CpuTotalMhz -Sum).Sum
            $memTotalMB  = ($Hosts | Measure-Object -Property MemoryTotalMB -Sum).Sum

            if ($cpuStat) {
                $avgCpuPct = ($cpuStat | Measure-Object -Property Value -Average).Average
                New-Finding -Site $Site -VCenter $VCName -Area 'Cluster (DRS/HA)' -Object $ClusterName `
                    -Item ("Cluster CPU Usage % (avg, last {0}h)" -f $PerfHistoryHours) `
                    -Value ("{0:N2}%  (Capacity {1:N1} GHz)" -f $avgCpuPct, ($cpuTotalMHz/1000)) -Status (Get-PctStatus $avgCpuPct)
            } else {
                New-Finding -Site $Site -VCenter $VCName -Area 'Cluster (DRS/HA)' -Object $ClusterName `
                    -Item 'Cluster CPU Usage %' -Value 'n/a' -Status 'Unable to Check' -Notes 'No historical performance samples returned (check vCenter stats level/retention).'
            }
            if ($memStat) {
                $avgMemPct = ($memStat | Measure-Object -Property Value -Average).Average
                New-Finding -Site $Site -VCenter $VCName -Area 'Cluster (DRS/HA)' -Object $ClusterName `
                    -Item ("Cluster Memory Usage % (avg, last {0}h)" -f $PerfHistoryHours) `
                    -Value ("{0:N2}%  (Capacity {1:N1} GB)" -f $avgMemPct, ($memTotalMB/1024)) -Status (Get-PctStatus $avgMemPct)
            } else {
                New-Finding -Site $Site -VCenter $VCName -Area 'Cluster (DRS/HA)' -Object $ClusterName `
                    -Item 'Cluster Memory Usage %' -Value 'n/a' -Status 'Unable to Check' -Notes 'No historical performance samples returned.'
            }
        }

        # Datastores in this cluster - capacity AND latency
        Invoke-SafeCheck -CheckName 'Datastore capacity/latency' -VCenter $VCName -Site $Site -ObjectName $ClusterName -Script {
            $datastores = Get-Datastore -Server $VC -RelatedObject $Cluster
            foreach ($ds in $datastores) {
                $Global:ObjectClusterMap[$ds.Name] = $ClusterName
                $usedPct = if ($ds.CapacityGB -gt 0) { (($ds.CapacityGB - $ds.FreeSpaceGB) / $ds.CapacityGB) * 100 } else { 0 }
                $accessible = $ds.ExtensionData.Summary.Accessible
                $capStatus = if (-not $accessible) { 'Critical' } else { Get-PctStatus $usedPct }
                New-Finding -Site $Site -VCenter $VCName -Area 'Storage' -Object $ds.Name `
                    -Item 'Capacity/Used/Free/Util%' `
                    -Value ("Cap {0:N0} GB | Free {1:N0} GB | Used {2:N1}% | Accessible: {3}" -f $ds.CapacityGB, $ds.FreeSpaceGB, $usedPct, $accessible) `
                    -Status $capStatus

                $latencyValue = 'n/a'
                $latencyStatus = 'Unable to Check'
                try {
                    $lat = Get-Stat -Server $VC -Entity $ds -Stat 'datastore.totalReadLatency.average','datastore.totalWriteLatency.average' -Realtime -MaxSamples 1 -ErrorAction Stop
                    if ($lat) {
                        $read  = ($lat | Where-Object { $_.MetricId -eq 'datastore.totalReadLatency.average' }  | Measure-Object -Property Value -Average).Average
                        $write = ($lat | Where-Object { $_.MetricId -eq 'datastore.totalWriteLatency.average' } | Measure-Object -Property Value -Average).Average
                        $worst = [Math]::Max([double]$read, [double]$write)
                        $latencyStatus = if ($worst -ge $DatastoreLatencyCriticalMs) { 'Critical' } elseif ($worst -ge $DatastoreLatencyWarningMs) { 'Warning' } else { 'Healthy' }
                        $latencyValue = ("Read {0:N1} ms | Write {1:N1} ms" -f $read, $write)
                    }
                } catch { }
                New-Finding -Site $Site -VCenter $VCName -Area 'Storage' -Object $ds.Name `
                    -Item 'Latency (Read/Write, realtime)' -Value $latencyValue -Status $latencyStatus `
                    -Notes $(if ($latencyStatus -eq 'Unable to Check') { 'Realtime datastore latency counters not returned - may require Storage I/O Control or a different stats level.' } else { '' })
            }
        }

        # Storage array hardware (controllers/disks/power) - no vendor-neutral API via vCenter.
        Invoke-SafeCheck -CheckName 'Storage array hardware' -VCenter $VCName -Site $Site -ObjectName $ClusterName -Script {
            New-Finding -Site $Site -VCenter $VCName -Area 'Storage' -Object $ClusterName `
                -Item 'Storage Array Hardware (Controllers/Disks/Power)' -Value 'n/a' -Status 'Manual/External Required' `
                -Notes 'No vendor-neutral API is exposed through vCenter/PowerCLI for physical storage-array health. Connect the array vendor''s own management API/module (e.g. NetApp ONTAP, Dell, HPE) to automate this line.'
        }

        # vSAN (only if the cluster actually has vSAN enabled and the module is available)
        Invoke-SafeCheck -CheckName 'vSAN health/capacity' -VCenter $VCName -Site $Site -ObjectName $ClusterName -Script {
            if (-not $Cluster.VsanEnabled) {
                New-Finding -Site $Site -VCenter $VCName -Area 'Storage' -Object $ClusterName `
                    -Item 'vSAN' -Value 'Not enabled on this cluster' -Status 'Information'
                return
            }
            if (-not $VsanModuleAvailable) {
                New-Finding -Site $Site -VCenter $VCName -Area 'Storage' -Object $ClusterName `
                    -Item 'vSAN Health/Capacity' -Value 'n/a' -Status 'Unable to Check' `
                    -Notes 'VMware.VimAutomation.Vsan module not installed on this laptop.'
                return
            }
            $space = Get-VsanSpaceUsage -Server $VC -Cluster $Cluster -ErrorAction Stop
            $usedPct = if ($space.CapacityGB -gt 0) { ($space.UsedGB / $space.CapacityGB) * 100 } else { 0 }
            New-Finding -Site $Site -VCenter $VCName -Area 'Storage' -Object $ClusterName `
                -Item 'vSAN Capacity/Used/Free/Util%' `
                -Value ("Cap {0:N0} GB | Used {1:N0} GB | Free {2:N0} GB | Util {3:N1}%" -f $space.CapacityGB, $space.UsedGB, ($space.CapacityGB - $space.UsedGB), $usedPct) `
                -Status (Get-PctStatus $usedPct)

            $healthTest = Get-VsanClusterHealth -Server $VC -Cluster $Cluster -ErrorAction Stop
            foreach ($grp in $healthTest.HealthGroups) {
                $bad = $grp.HealthTests | Where-Object { $_.Status -ne 'green' }
                $status = if ($bad) { 'Warning' } else { 'Healthy' }
                New-Finding -Site $Site -VCenter $VCName -Area 'Storage' -Object $ClusterName `
                    -Item "vSAN Health: $($grp.GroupName)" -Value $(if ($bad) { ($bad.TestName -join '; ') } else { 'OK' }) -Status $status
            }
        }

        # --- VM Fleet -------------------------------------------------------------------------
        $VMs = Invoke-SafeCheck -CheckName 'VM discovery' -VCenter $VCName -Site $Site -ObjectName $ClusterName -Script {
            Get-VM -Server $VC -Location $Cluster
        }

        if ($VMs) {
            # Count / power state / guest OS distribution
            Invoke-SafeCheck -CheckName 'VM inventory' -VCenter $VCName -Site $Site -ObjectName $ClusterName -Script {
                $on  = ($VMs | Where-Object { $_.PowerState -eq 'PoweredOn' }).Count
                $off = ($VMs | Where-Object { $_.PowerState -eq 'PoweredOff' }).Count
                New-Finding -Site $Site -VCenter $VCName -Area 'VM Fleet' -Object $ClusterName `
                    -Item 'VM Count' -Value "Total $($VMs.Count) | On $on | Off $off" -Status 'Information'

                $osDist = $VMs | Group-Object { if ($_.Guest.OSFullName) { $_.Guest.OSFullName } else { $_.ExtensionData.Config.GuestFullName } } |
                    Sort-Object Count -Descending
                foreach ($g in $osDist) {
                    New-Finding -Site $Site -VCenter $VCName -Area 'VM Fleet' -Object $ClusterName `
                        -Item "Guest OS: $($g.Name)" -Value $g.Count -Status 'Information'
                }
            }

            # Production vs non-production split - via vSphere Tags if an environment-style
            # category exists, otherwise a folder-name heuristic, otherwise honestly "Unable to Check".
            Invoke-SafeCheck -CheckName 'Production vs non-production split' -VCenter $VCName -Site $Site -ObjectName $ClusterName -Script {
                $classified = $false
                try {
                    $assignments = Get-TagAssignment -Entity $VMs -Server $VC -ErrorAction Stop
                    $envAssignments = $assignments | Where-Object { $_.Tag.Category.Name -match 'env|production' }
                    if ($envAssignments) {
                        $prodCount = ($envAssignments | Where-Object { $_.Tag.Name -match 'prod' -and $_.Tag.Name -notmatch 'non|pre' } | Select-Object -ExpandProperty Entity -Unique).Count
                        $nonProdCount = $VMs.Count - $prodCount
                        New-Finding -Site $Site -VCenter $VCName -Area 'VM Fleet' -Object $ClusterName `
                            -Item 'Production vs Non-Production (by vSphere Tag)' -Value "Prod $prodCount | Non-Prod $nonProdCount" -Status 'Information'
                        $classified = $true
                    }
                } catch { }
                if (-not $classified) {
                    $prodCount = ($VMs | Where-Object { $_.Folder.Name -match 'prod' }).Count
                    if ($prodCount -gt 0) {
                        New-Finding -Site $Site -VCenter $VCName -Area 'VM Fleet' -Object $ClusterName `
                            -Item 'Production vs Non-Production (folder-name heuristic)' -Value "Prod $prodCount | Non-Prod $($VMs.Count - $prodCount)" -Status 'Information' `
                            -Notes 'No vSphere Tag category found for environment classification; inferred from folder naming - verify manually.'
                    } else {
                        New-Finding -Site $Site -VCenter $VCName -Area 'VM Fleet' -Object $ClusterName `
                            -Item 'Production vs Non-Production' -Value 'n/a' -Status 'Unable to Check' `
                            -Notes 'No vSphere Tags or folder naming convention detected to classify VMs as production/non-production.'
                    }
                }
            }

            # VMware Tools status per VM (running/outdated)
            Invoke-SafeCheck -CheckName 'VMware Tools status' -VCenter $VCName -Site $Site -ObjectName $ClusterName -Script {
                $toolsGroups = $VMs | Group-Object { $_.ExtensionData.Guest.ToolsStatus }
                foreach ($g in $toolsGroups) {
                    $status = switch ($g.Name) {
                        'toolsOk'           { 'Healthy' }
                        'toolsOld'          { 'Warning' }
                        'toolsNotRunning'   { 'Warning' }
                        'toolsNotInstalled' { 'Warning' }
                        default             { 'Information' }
                    }
                    New-Finding -Site $Site -VCenter $VCName -Area 'VM Fleet' -Object $ClusterName `
                        -Item "VMware Tools: $($g.Name)" -Value $g.Count -Status $status
                }
            }

            # Snapshot inventory - flag any older than the configured threshold (default 3 days)
            Invoke-SafeCheck -CheckName 'Snapshot inventory' -VCenter $VCName -Site $Site -ObjectName $ClusterName -Script {
                $snaps = Get-Snapshot -VM $VMs -Server $VC -ErrorAction SilentlyContinue
                New-Finding -Site $Site -VCenter $VCName -Area 'VM Fleet' -Object $ClusterName `
                    -Item 'Total Snapshots' -Value $(if ($snaps) { $snaps.Count } else { 0 }) -Status 'Information'
                if ($snaps) {
                    $old = $snaps | Where-Object { $_.Created -lt (Get-Date).AddDays(-$SnapshotAgeWarningDays) }
                    if ($old) {
                        foreach ($s in $old) {
                            $ageDays = [Math]::Round(((Get-Date) - $s.Created).TotalDays)
                            New-Finding -Site $Site -VCenter $VCName -Area 'VM Fleet' -Object $s.VM.Name `
                                -Item 'Snapshot Older Than Threshold' -Value ("'{0}' created {1:yyyy-MM-dd} ({2} days old)" -f $s.Name, $s.Created, $ageDays) -Status 'Warning'
                        }
                    } else {
                        New-Finding -Site $Site -VCenter $VCName -Area 'VM Fleet' -Object $ClusterName `
                            -Item 'Snapshots Older Than Threshold' -Value 'None' -Status 'Healthy'
                    }
                }
            }

            # Orphaned/powered-off VM aging - flag VMs powered off longer than the configured threshold.
            Invoke-SafeCheck -CheckName 'Powered-off VM aging' -VCenter $VCName -Site $Site -ObjectName $ClusterName -Script {
                $offVMs = $VMs | Where-Object { $_.PowerState -eq 'PoweredOff' }
                New-Finding -Site $Site -VCenter $VCName -Area 'VM Fleet' -Object $ClusterName `
                    -Item 'Powered-Off VM Count' -Value $offVMs.Count -Status 'Information'
                foreach ($vm in $offVMs) {
                    try {
                        $evt = Get-VIEvent -Entity $vm -Server $VC -MaxSamples 10 -ErrorAction Stop |
                            Where-Object { $_.GetType().Name -match 'PoweredOff' } | Select-Object -First 1
                        if ($evt) {
                            $daysOff = [Math]::Round(((Get-Date) - $evt.CreatedTime).TotalDays)
                            if ($daysOff -ge $PoweredOffAgingDays) {
                                New-Finding -Site $Site -VCenter $VCName -Area 'VM Fleet' -Object $vm.Name `
                                    -Item 'Powered Off Since' -Value ("{0:yyyy-MM-dd} ({1} days)" -f $evt.CreatedTime, $daysOff) -Status 'Warning' `
                                    -Notes 'Review for decommission/cleanup.'
                            }
                        }
                    } catch { }
                }
            }

            # Backup status - only if the backup product surfaces status via a vCenter Custom
            # Attribute; otherwise honestly marked Manual/External Required (see script NOTES).
            Invoke-SafeCheck -CheckName 'Backup status (VM custom attributes)' -VCenter $VCName -Site $Site -ObjectName $ClusterName -Script {
                $backupAttrs = Get-CustomAttribute -Server $VC -ErrorAction SilentlyContinue | Where-Object { $_.Name -match 'backup' }
                if (-not $backupAttrs) {
                    New-Finding -Site $Site -VCenter $VCName -Area 'Backup & Licensing' -Object $ClusterName `
                        -Item 'Backup Success/Failure per VM' -Value 'n/a' -Status 'Manual/External Required' `
                        -Notes 'No backup-related Custom Attribute found in vCenter, and no vendor-neutral backup API exists. Connect your backup product''s PowerShell module/API (e.g. Veeam, Cohesity, Commvault) to automate this section.'
                    return
                }
                foreach ($attr in $backupAttrs) {
                    $annotations = Get-Annotation -Entity $VMs -CustomAttribute $attr -Server $VC -ErrorAction SilentlyContinue
                    $withValue = $annotations | Where-Object { -not [string]::IsNullOrWhiteSpace($_.Value) }
                    $missing = $VMs.Count - $withValue.Count
                    New-Finding -Site $Site -VCenter $VCName -Area 'Backup & Licensing' -Object $ClusterName `
                        -Item "Backup Coverage ('$($attr.Name)' attribute)" -Value "$($withValue.Count)/$($VMs.Count) VMs have a value set" `
                        -Status $(if ($missing -eq 0) { 'Healthy' } else { 'Warning' }) `
                        -Notes "VMs without a value in the '$($attr.Name)' custom attribute may lack backup coverage - verify against your backup product's console."
                }
            }
        }

        # vDS / port groups / VLAN / teaming / MTU for this cluster's hosts
        Invoke-SafeCheck -CheckName 'vDS/networking config' -VCenter $VCName -Site $Site -ObjectName $ClusterName -Script {
            $vdSwitches = Get-VDSwitch -Server $VC -VMHost $Hosts -ErrorAction SilentlyContinue | Select-Object -Unique
            if (-not $vdSwitches) {
                New-Finding -Site $Site -VCenter $VCName -Area 'Networking' -Object $ClusterName `
                    -Item 'Distributed Switches' -Value 'None (standard vSwitches only, or none discovered)' -Status 'Information'
                return
            }
            $mtus = @()
            foreach ($vds in $vdSwitches) {
                $mtus += $vds.Mtu
                $Global:ObjectClusterMap[$vds.Name] = $ClusterName
                New-Finding -Site $Site -VCenter $VCName -Area 'Networking' -Object $vds.Name `
                    -Item 'vDS MTU' -Value $vds.Mtu -Status 'Information'

                $pgs = Get-VDPortgroup -Server $VC -VDSwitch $vds
                foreach ($pg in $pgs) {
                    $Global:ObjectClusterMap[$pg.Name] = $ClusterName
                    $vlan = $pg.ExtensionData.Config.DefaultPortConfig.Vlan.VlanId
                    New-Finding -Site $Site -VCenter $VCName -Area 'Networking' -Object $pg.Name `
                        -Item 'Port Group VLAN' -Value $vlan -Status 'Information'
                    $teaming = $pg.ExtensionData.Config.DefaultPortConfig.UplinkTeamingPolicy.Policy.Value
                    New-Finding -Site $Site -VCenter $VCName -Area 'Networking' -Object $pg.Name `
                        -Item 'Teaming/Failover Policy' -Value $teaming -Status 'Information'
                }
            }
            $distinctMtu = $mtus | Select-Object -Unique
            New-Finding -Site $Site -VCenter $VCName -Area 'Networking' -Object $ClusterName `
                -Item 'MTU Consistency (vDS)' -Value ($distinctMtu -join ', ') -Status $(if ($distinctMtu.Count -le 1) {'Healthy'} else {'Warning'})
        }

    } # end per-cluster
} # end per-vCenter

# ============================================================================
# 3. TREND DATA - save this run's risk counts per site, load last run's for comparison
# ============================================================================
$Global:SiteTrend = @{}
foreach ($SiteLabel in ($Global:AllResults | Select-Object -ExpandProperty Site -Unique)) {
    $siteResults = $Global:AllResults | Where-Object { $_.Site -eq $SiteLabel }
    $high   = ($siteResults | Where-Object { $_.Status -eq 'Critical' }).Count
    $medium = ($siteResults | Where-Object { $_.Status -eq 'Warning' }).Count
    $low    = ($siteResults | Where-Object { $_.Status -eq 'Unable to Check' }).Count

    $safeLabel = ($SiteLabel -replace '[\\/\?\*\[\]:<>\|]', '_')
    $snapshotPath = Join-Path $OutputPath "$safeLabel`_LastRun.json"
    $previous = $null
    if (Test-Path $snapshotPath) {
        try { $previous = Get-Content -Path $snapshotPath -Raw | ConvertFrom-Json } catch { $previous = $null }
    }
    $Global:SiteTrend[$SiteLabel] = [pscustomobject]@{ High = $high; Medium = $medium; Low = $low; Previous = $previous }

    [pscustomobject]@{ RunDate = $RunDate; High = $high; Medium = $medium; Low = $low } |
        ConvertTo-Json | Out-File -FilePath $snapshotPath -Encoding UTF8
}

# ============================================================================
# ============================================================================
# 4. REPORT GENERATION HELPERS
#    Each site's .docx is built DIRECTLY in Word via COM automation - typing text,
#    applying styles/colors, and inserting native Word tables. There is no HTML
#    step anywhere in this pipeline (opening an HTML file in Word and re-saving
#    it is what caused earlier failures - Word treats that as a web/compat
#    document and SaveAs on it is unreliable). This is slower to write but far
#    more reliable, and gives real Word styles/colors instead of an HTML import.
# ============================================================================
$StatusColors = @{
    'Healthy'         = @(22,163,74)
    'Warning'         = @(245,158,11)
    'Critical'        = @(220,38,38)
    'Information'     = @(37,99,235)
    'Unable to Check' = @(156,163,175)
}
function Get-WordColorLong {
    param([int[]]$Rgb)
    # Word/COLORREF packs color as R + G*256 + B*65536 (same packing Windows uses)
    return $Rgb[0] + ($Rgb[1] * 256) + ($Rgb[2] * 65536)
}
function Get-StatusColorLong {
    param([string]$Status)
    $rgb = $StatusColors[$Status]
    if (-not $rgb) { $rgb = @(156,163,175) }
    return Get-WordColorLong $rgb
}

function Write-Heading {
    param($Selection, [string]$Text, [int]$Level = 1)
    $Selection.Style = "Heading $Level"
    $Selection.TypeText($Text)
    $Selection.TypeParagraph()
    $Selection.Style = 'Normal'
}

function Write-Para {
    param($Selection, [string]$Text, [switch]$Bold, [switch]$Italic, [switch]$Gray)
    $Selection.Font.Bold = [bool]$Bold
    $Selection.Font.Italic = [bool]$Italic
    if ($Gray) { $Selection.Font.Color = Get-WordColorLong @(107,114,128) }
    $Selection.TypeText($Text)
    $Selection.TypeParagraph()
    $Selection.Font.Bold = $false
    $Selection.Font.Italic = $false
    $Selection.Font.Color = -16777216   # wdColorAutomatic
}

function Write-DotLine {
    param($Selection, [string]$Status, [string]$Text)
    $Selection.Font.Color = Get-StatusColorLong $Status
    $Selection.TypeText([char]0x25CF)
    $Selection.Font.Color = -16777216
    $Selection.TypeText(" $Text")
    $Selection.TypeParagraph()
}

function Write-BulletList {
    param($Selection, [string[]]$Lines, [switch]$Red)
    foreach ($line in $Lines) {
        if ($Red) { $Selection.Font.Color = Get-WordColorLong @(220,38,38) }
        $Selection.TypeText("-  $line")
        $Selection.TypeParagraph()
        if ($Red) { $Selection.Font.Color = -16777216 }
    }
}

function Add-Table {
    param($Doc, $Selection, [string[]]$Headers, [array]$Rows, [int]$StatusColumnIndex = -1)
    if (-not $Rows -or $Rows.Count -eq 0) { return }
    $rowCount = $Rows.Count + 1
    $colCount = $Headers.Count
    $table = $Doc.Tables.Add($Selection.Range, $rowCount, $colCount)
    $table.Borders.InsideLineStyle = 1
    $table.Borders.OutsideLineStyle = 1
    for ($c = 0; $c -lt $colCount; $c++) {
        $cell = $table.Cell(1, $c + 1)
        $cell.Range.Text = $Headers[$c]
        $cell.Range.Font.Bold = $true
        $cell.Range.Font.Color = Get-WordColorLong @(255,255,255)
        $cell.Shading.BackgroundPatternColor = Get-WordColorLong @(30,58,95)
    }
    for ($r = 0; $r -lt $Rows.Count; $r++) {
        for ($c = 0; $c -lt $colCount; $c++) {
            $cellRange = $table.Cell($r + 2, $c + 1).Range
            $cellText = [string]$Rows[$r][$c]
            $cellRange.Text = $cellText
            if ($c -eq $StatusColumnIndex) {
                $cellRange.Font.Color = Get-StatusColorLong $cellText
                $cellRange.Font.Bold = $true
            }
        }
    }
    $Selection.SetRange($table.Range.End, $table.Range.End)
    $Selection.TypeParagraph()
}

function Get-WorstStatus {
    param([string[]]$Statuses)
    $order = @{ 'Critical' = 4; 'Warning' = 3; 'Unable to Check' = 2; 'Healthy' = 1; 'Information' = 0 }
    if (-not $Statuses -or $Statuses.Count -eq 0) { return 'Healthy' }
    return ($Statuses | Sort-Object { $order[$_] } -Descending | Select-Object -First 1)
}

# Looks up a single finding's Value for a given Object+Item, or 'n/a' if not collected.
function Get-FindingValue {
    param($Findings, [string]$ObjectName, [string]$ItemName)
    $f = $Findings | Where-Object { $_.Object -eq $ObjectName -and $_.Item -eq $ItemName } | Select-Object -First 1
    if ($f) { return $f.Value } else { return 'n/a' }
}

function Write-SiteReportDocx {
    param($Word, [string]$SiteLabel, [string]$OutputPath, [string]$RunDate)

    $SiteFindings = $Global:AllResults | Where-Object { $_.Site -eq $SiteLabel }
    if (-not $SiteFindings) { return }

    $VCenterName = $SiteFindings | Select-Object -First 1 -ExpandProperty VCenter
    $ClusterNames = if ($Global:SiteClusterMap.ContainsKey($SiteLabel)) { $Global:SiteClusterMap[$SiteLabel] } else { @() }
    $HostObjects = $SiteFindings | Where-Object { $_.Item -eq 'ESXi Version/Build' } | Select-Object -ExpandProperty Object -Unique

    # Risk counting excludes Information/Healthy/Manual-External rows by construction of Status values.
    $CritCount = ($SiteFindings | Where-Object { $_.Status -eq 'Critical' }).Count
    $WarnCount = ($SiteFindings | Where-Object { $_.Status -eq 'Warning' }).Count
    $UnableCount = ($SiteFindings | Where-Object { $_.Status -eq 'Unable to Check' }).Count
    $OverallHealth = if ($CritCount -gt 0) { 'Critical' } elseif ($WarnCount -gt 0) { 'Warning' } else { 'Healthy' }
    $VcVersionItem = $SiteFindings | Where-Object { $_.Item -eq 'vCenter Version/Build' } | Select-Object -First 1

    $doc = $Word.Documents.Add()
    $sel = $Word.Selection

    $TitlePrefix = if ($CompanyName) { "$CompanyName - " } else { '' }
    Write-Heading $sel "$TitlePrefix$SiteLabel vCenter & ESXi Health Check Report" 1
    Write-Para $sel "Report Date: $RunDate" -Gray
    Write-Para $sel "vCenter: $VCenterName" -Gray

    # ---------------- 1. Executive Summary ----------------
    Write-Heading $sel "1. Executive Summary" 2
    Write-Para $sel "Environment: 1 vCenter ($VCenterName) | $($ClusterNames.Count) Cluster(s)/Site(s) | $($HostObjects.Count) ESXi Host(s)"
    if ($VcVersionItem) { Write-Para $sel "vCenter Version: $($VcVersionItem.Value)" }
    Write-DotLine $sel $OverallHealth "Overall Environment Health: $OverallHealth"

    Write-Para $sel "Key Findings Summary:" -Bold
    Write-BulletList $sel @(
        "High Risk Issues (Critical): $CritCount"
        "Medium Risk Issues (Warning): $WarnCount"
        "Low Risk Issues (Unable to Verify): $UnableCount"
    )
    Write-Para $sel "'Low Risk' = items this automated run could not verify (e.g. missing sensor data, no CIS session) - not a confirmed problem, but worth a manual look." -Italic -Gray

    $trend = $Global:SiteTrend[$SiteLabel]
    if ($trend -and $trend.Previous) {
        $deltaHigh = $trend.High - $trend.Previous.High
        $deltaMed  = $trend.Medium - $trend.Previous.Medium
        $fmt = { param($d) if ($d -gt 0) { "+$d" } else { "$d" } }
        Write-Para $sel ("Trend vs. Last Run ({0}): High {1}, Medium {2}" -f $trend.Previous.RunDate, (& $fmt $deltaHigh), (& $fmt $deltaMed))
    } else {
        Write-Para $sel "Trend vs. Last Run: no prior run data available for comparison yet." -Italic -Gray
    }

    $topRisks = $SiteFindings | Where-Object { $_.Status -eq 'Critical' } | Select-Object -First $TopRisksCount
    if ($topRisks.Count -lt $TopRisksCount) {
        $topRisks = @($topRisks) + @($SiteFindings | Where-Object { $_.Status -eq 'Warning' } | Select-Object -First ($TopRisksCount - $topRisks.Count))
    }
    Write-Para $sel "Top Risks:" -Bold
    if ($topRisks) {
        foreach ($r in $topRisks) { Write-DotLine $sel $r.Status "$($r.Object) - $($r.Item): $($r.Value)" }
    } else {
        Write-Para $sel "No Critical or Warning items identified this run."
    }

    # ---------------- 2. vCenter & Hosts ----------------
    Write-Heading $sel "2. vCenter & Hosts" 2
    $ApplianceRows = $SiteFindings | Where-Object { $_.Area -eq 'vCenter & Host Health' -and $_.Object -eq $VCenterName -and $_.Item -notin @('vCenter Version/Build') }
    if ($ApplianceRows) {
        $tableRows = @()
        foreach ($r in $ApplianceRows) { $tableRows += ,@($r.Item, $r.Value, $r.Status) }
        Add-Table $doc $sel @('Item','Value','Status') $tableRows -StatusColumnIndex 2
    } else {
        Write-Para $sel "No vCenter appliance-level data collected this run." -Italic
    }

    Write-Heading $sel "Per-Host Summary (not a fleet average)" 3
    $hostTableRows = @()
    foreach ($h in $HostObjects) {
        $cluster = if ($Global:ObjectClusterMap.ContainsKey($h)) { $Global:ObjectClusterMap[$h] } else { 'n/a' }
        $ver = Get-FindingValue $SiteFindings $h 'ESXi Version/Build'
        $cpu = Get-FindingValue $SiteFindings $h 'CPU Usage %'
        $mem = Get-FindingValue $SiteFindings $h 'Memory Usage %'
        $uptime = Get-FindingValue $SiteFindings $h 'Uptime (days)'
        $hostRows = $SiteFindings | Where-Object { $_.Object -eq $h }
        $hostStatus = Get-WorstStatus ($hostRows | Select-Object -ExpandProperty Status)
        $hostTableRows += ,@($h, $cluster, $ver, $cpu, $mem, $uptime, $hostStatus)
    }
    if ($hostTableRows.Count -gt 0) {
        Add-Table $doc $sel @('Host','Cluster','ESXi Version/Build','CPU %','Memory %','Uptime (days)','Status') $hostTableRows -StatusColumnIndex 6
    } else {
        Write-Para $sel "No hosts discovered." -Italic
    }

    # ---------------- 3. Cluster (DRS/HA) ----------------
    Write-Heading $sel "3. Cluster (DRS/HA)" 2
    foreach ($ClusterName in $ClusterNames) {
        Write-Heading $sel $ClusterName 3
        $clusterRows = $SiteFindings | Where-Object { $_.Area -eq 'Cluster (DRS/HA)' -and $_.Object -eq $ClusterName }
        if ($clusterRows) {
            $tableRows = @()
            foreach ($r in $clusterRows) { $tableRows += ,@($r.Item, $r.Value, $r.Status) }
            Add-Table $doc $sel @('Item','Value','Status') $tableRows -StatusColumnIndex 2
        } else {
            Write-Para $sel "No cluster configuration data collected." -Italic
        }
    }

    # ---------------- 4. Storage ----------------
    Write-Heading $sel "4. Storage" 2
    foreach ($ClusterName in $ClusterNames) {
        Write-Heading $sel $ClusterName 3
        $dsRows = $SiteFindings | Where-Object { $_.Area -eq 'Storage' -and $_.Item -like '*Capacity*' -and $Global:ObjectClusterMap[$_.Object] -eq $ClusterName }
        $latRows = $SiteFindings | Where-Object { $_.Area -eq 'Storage' -and $_.Item -like 'Latency*' -and $Global:ObjectClusterMap[$_.Object] -eq $ClusterName }
        if ($dsRows) {
            $tableRows = @()
            foreach ($r in $dsRows) {
                $lat = ($latRows | Where-Object { $_.Object -eq $r.Object } | Select-Object -First 1)
                $latVal = if ($lat) { $lat.Value } else { 'n/a' }
                $latStatus = if ($lat) { $lat.Status } else { 'Unable to Check' }
                $worst = Get-WorstStatus @($r.Status, $latStatus)
                $tableRows += ,@($r.Object, $r.Value, $latVal, $worst)
            }
            Add-Table $doc $sel @('Datastore','Capacity/Used/Free/Util%','Latency (Read/Write)','Status') $tableRows -StatusColumnIndex 3
        } else {
            Write-Para $sel "No datastore data collected for this cluster." -Italic
        }

        $vsanRows = $SiteFindings | Where-Object { $_.Area -eq 'Storage' -and $_.Object -eq $ClusterName -and $_.Item -like 'vSAN*' }
        foreach ($vr in $vsanRows) { Write-DotLine $sel $vr.Status "$($vr.Item): $($vr.Value)" }

        $arrayRow = $SiteFindings | Where-Object { $_.Area -eq 'Storage' -and $_.Object -eq $ClusterName -and $_.Item -eq 'Storage Array Hardware (Controllers/Disks/Power)' } | Select-Object -First 1
        if ($arrayRow) { Write-DotLine $sel $arrayRow.Status "$($arrayRow.Item): $($arrayRow.Notes)" }
    }

    # ---------------- 5. Networking ----------------
    Write-Heading $sel "5. Networking" 2
    foreach ($ClusterName in $ClusterNames) {
        Write-Heading $sel $ClusterName 3
        $HostsHere = $HostObjects | Where-Object { $Global:ObjectClusterMap[$_] -eq $ClusterName }

        $netConfigRows = $SiteFindings | Where-Object { $_.Area -eq 'Networking' -and $_.Object -eq $ClusterName }
        $vdsRows = $SiteFindings | Where-Object { $_.Area -eq 'Networking' -and $Global:ObjectClusterMap[$_.Object] -eq $ClusterName -and $_.Object -ne $ClusterName }
        foreach ($n in $netConfigRows) { Write-DotLine $sel $n.Status "$($n.Item): $($n.Value)" }
        foreach ($n in $vdsRows) { Write-DotLine $sel $n.Status "$($n.Object) - $($n.Item): $($n.Value)" }

        $nicTableRows = @()
        foreach ($h in $HostsHere) {
            $pnic = Get-FindingValue $SiteFindings $h 'Physical NICs'
            $errRow = $SiteFindings | Where-Object { $_.Object -eq $h -and $_.Item -like 'NIC Errors*' } | Select-Object -First 1
            $errVal = if ($errRow) { $errRow.Value } else { 'n/a' }
            $errStatus = if ($errRow) { $errRow.Status } else { 'Unable to Check' }
            $nicTableRows += ,@($h, $pnic, $errVal, $errStatus)
        }
        if ($nicTableRows.Count -gt 0) {
            Add-Table $doc $sel @('Host','Physical NICs','Errors/Drops (latest sample)','Status') $nicTableRows -StatusColumnIndex 3
        }
    }

    # ---------------- 6. VM Fleet ----------------
    Write-Heading $sel "6. VM Fleet" 2
    foreach ($ClusterName in $ClusterNames) {
        Write-Heading $sel $ClusterName 3
        $vmCountRow = $SiteFindings | Where-Object { $_.Area -eq 'VM Fleet' -and $_.Object -eq $ClusterName -and $_.Item -eq 'VM Count' } | Select-Object -First 1
        if ($vmCountRow) { Write-Para $sel $vmCountRow.Value }
        $splitRow = $SiteFindings | Where-Object { $_.Area -eq 'VM Fleet' -and $_.Object -eq $ClusterName -and $_.Item -like 'Production vs*' } | Select-Object -First 1
        if ($splitRow) { Write-DotLine $sel $splitRow.Status "$($splitRow.Item): $($splitRow.Value)" }

        $osRows = $SiteFindings | Where-Object { $_.Area -eq 'VM Fleet' -and $_.Object -eq $ClusterName -and $_.Item -like 'Guest OS:*' }
        if ($osRows) {
            $osTableRows = @()
            foreach ($o in $osRows) { $osTableRows += ,@(($o.Item -replace 'Guest OS: ',''), $o.Value) }
            Add-Table $doc $sel @('Guest OS','Count') $osTableRows
        }

        $toolsRows = $SiteFindings | Where-Object { $_.Area -eq 'VM Fleet' -and $_.Object -eq $ClusterName -and $_.Item -like 'VMware Tools:*' }
        if ($toolsRows) {
            $toolsTableRows = @()
            foreach ($t in $toolsRows) { $toolsTableRows += ,@(($t.Item -replace 'VMware Tools: ',''), $t.Value, $t.Status) }
            Add-Table $doc $sel @('VMware Tools Status','VM Count','Status') $toolsTableRows -StatusColumnIndex 2
        }

        $snapRows = $SiteFindings | Where-Object { $_.Area -eq 'VM Fleet' -and $_.Item -eq 'Snapshot Older Than Threshold' -and $Global:ObjectClusterMap[$_.Object] -eq $ClusterName }
        if ($snapRows) {
            Write-Heading $sel "Snapshots to Review" 4
            foreach ($s in $snapRows) { Write-DotLine $sel $s.Status "$($s.Object): $($s.Value)" }
        }

        $offRows = $SiteFindings | Where-Object { $_.Area -eq 'VM Fleet' -and $_.Item -eq 'Powered Off Since' -and $Global:ObjectClusterMap[$_.Object] -eq $ClusterName }
        if ($offRows) {
            Write-Heading $sel "Aging Powered-Off VMs" 4
            foreach ($o in $offRows) { Write-DotLine $sel $o.Status "$($o.Object): $($o.Value)" }
        }
    }

    # ---------------- 7. Security ----------------
    Write-Heading $sel "7. Security" 2
    $secRows = $SiteFindings | Where-Object { $_.Area -eq 'Security' } | Sort-Object Object, Item
    if ($secRows) {
        $secStatus = Get-WorstStatus ($secRows | Select-Object -ExpandProperty Status)
        Write-DotLine $sel $secStatus "Overall Security Status: $secStatus"
        $tableRows = @()
        foreach ($r in $secRows) { $tableRows += ,@($r.Object, $r.Item, $r.Value, $r.Status) }
        Add-Table $doc $sel @('Object','Item','Value','Status') $tableRows -StatusColumnIndex 3
    } else {
        Write-Para $sel "No security data collected this run." -Italic
    }

    # ---------------- 8. Backup & Licensing ----------------
    Write-Heading $sel "8. Backup & Licensing" 2
    $blRows = $SiteFindings | Where-Object { $_.Area -eq 'Backup & Licensing' } | Sort-Object Object, Item
    if ($blRows) {
        $tableRows = @()
        foreach ($r in $blRows) { $tableRows += ,@($r.Object, $r.Item, $r.Value, $r.Status) }
        Add-Table $doc $sel @('Object','Item','Value','Status') $tableRows -StatusColumnIndex 3
    } else {
        Write-Para $sel "No backup/licensing data collected this run." -Italic
    }

    # ---------------- 9. Alarms ----------------
    Write-Heading $sel "9. Alarms" 2
    $alarmRows = $SiteFindings | Where-Object { $_.Area -eq 'Alarms' -and $_.Item -ne 'Triggered Alarms' }
    if ($alarmRows) {
        $alarmLines = $alarmRows | ForEach-Object { "$($_.Object) - $($_.Item) ($($_.Value))" }
        Write-BulletList $sel $alarmLines -Red
    } else {
        Write-Para $sel "No active alarms triggered this run."
    }

    # ---------------- 10. Action Plan ----------------
    Write-Heading $sel "10. Action Plan" 2
    Write-Para $sel "Dated, owned fixes for every Critical/Warning item found this run. Suggested Target Date is a configurable SLA window ($CriticalRemediationDays days for High risk, $WarningRemediationDays days for Medium risk), not a mandated deadline. Owner is intentionally blank - vCenter has no concept of fix ownership." -Italic -Gray

    $actionItems = $SiteFindings | Where-Object { $_.Status -in 'Critical','Warning' } | Sort-Object { if ($_.Status -eq 'Critical') { 0 } else { 1 } }, Area, Object
    if ($actionItems) {
        $actionRows = @()
        foreach ($a in $actionItems) {
            $dueDays = if ($a.Status -eq 'Critical') { $CriticalRemediationDays } else { $WarningRemediationDays }
            $dueDate = (Get-Date $RunDate).AddDays($dueDays).ToString('yyyy-MM-dd')
            $recommended = if ($a.Status -eq 'Critical') { 'Immediate remediation required' } else { 'Schedule remediation' }
            $actionRows += ,@($a.Area, $a.Object, $a.Item, $a.Value, $a.Status, $recommended, $dueDate, '')
        }
        Add-Table $doc $sel @('Area','Object','Item','Value','Status','Recommended Action','Suggested Target Date','Owner') $actionRows -StatusColumnIndex 4
    } else {
        Write-Para $sel "No Critical or Warning items require action this run."
    }

    # ---------------- Conclusion / sign-off ----------------
    Write-Heading $sel "Conclusion" 2
    $concl = if ($CritCount -gt 0) {
        "The $SiteLabel VMware environment has $CritCount critical item(s) requiring prompt attention, alongside $WarnCount warning item(s). See the Action Plan above."
    } elseif ($WarnCount -gt 0) {
        "The $SiteLabel VMware environment is stable with $WarnCount warning item(s) to track. No critical issues were identified during this automated check."
    } else {
        "The $SiteLabel VMware environment is stable and healthy. No critical or warning issues were identified during this automated check."
    }
    Write-Para $sel $concl
    Write-Para $sel "Prepared By: $PreparedBy" -Bold
    if ($PreparedByTitle) { Write-Para $sel $PreparedByTitle }
    Write-Para $sel "Report Date: $RunDate" -Bold

    $siteFailures = $Global:FailureLog | Where-Object { $_.Site -eq $SiteLabel }
    if ($siteFailures) {
        Write-Para $sel "$($siteFailures.Count) item(s) could not be automatically verified this run - see the .log file for detail." -Italic -Gray
    }

    $safeLabel = ($SiteLabel -replace '[\\/\?\*\[\]:<>\|]', '_')
    $DocxPath = Join-Path $OutputPath "$safeLabel`_VMware_HealthCheck_$RunDate.docx"
    try {
        # InvokeMember avoids the [ref]-parameter marshalling bug some PowerShell/.NET
        # combinations hit on a direct $doc.SaveAs(...) call.
        $null = $doc.GetType().InvokeMember('SaveAs', [System.Reflection.BindingFlags]::InvokeMethod, $null, $doc, @([string]$DocxPath, 16))
        Write-Host "Report written: $DocxPath" -ForegroundColor Cyan
    } catch {
        Write-CheckLog -VCenter 'n/a' -Site $SiteLabel -Object 'DOCX export' -CheckName 'Word COM automation' -ErrorMessage $_.Exception.Message
        Write-Warning "Could not save .docx for $SiteLabel : $($_.Exception.Message)"
    }
    # Close without prompting regardless of whether SaveAs succeeded (0 = wdDoNotSaveChanges) -
    # avoids any invisible "save changes?" prompt hanging the run.
    $null = $doc.GetType().InvokeMember('Close', [System.Reflection.BindingFlags]::InvokeMethod, $null, $doc, @(0))
    [System.Runtime.Interopservices.Marshal]::ReleaseComObject($doc) | Out-Null
}

# ============================================================================
# 5. OUTPUT: ONE DOCX PER SITE - built natively, no HTML step
# ============================================================================
$SiteLabels = $Global:AllResults | Select-Object -ExpandProperty Site -Unique
$WordAvailable = $true
try {
    $Word = New-Object -ComObject Word.Application
    $Word.Visible = $false
} catch {
    $WordAvailable = $false
    Write-CheckLog -VCenter 'n/a' -Site 'n/a' -Object 'DOCX export' -CheckName 'Word COM automation' -ErrorMessage $_.Exception.Message
    Write-Warning "Microsoft Word is not available on this machine - cannot generate .docx reports."
}

if ($WordAvailable) {
    foreach ($SiteLabel in $SiteLabels) {
        Write-SiteReportDocx -Word $Word -SiteLabel $SiteLabel -OutputPath $OutputPath -RunDate $RunDate
    }
    $Word.Quit()
    [System.Runtime.Interopservices.Marshal]::ReleaseComObject($Word) | Out-Null
}

# ============================================================================
# 6. OUTPUT: LOG FILE (written last so it also captures any DOCX export failures)
# ============================================================================
$LogPath = Join-Path $OutputPath "VMware_Weekly_HealthCheck_$RunDate.log"
$LogLines = @()
$LogLines += "VMware Weekly Health Check run - $($ScriptStart.ToString('u'))"
$LogLines += "Sites processed: $($SiteLabels -join ', ')"
$LogLines += "Total findings collected: $($Global:AllResults.Count)"
$LogLines += "Total collection failures: $($Global:FailureLog.Count)"
$LogLines += "----------------------------------------------------------------"
foreach ($f in $Global:FailureLog) {
    $LogLines += "$($f.Timestamp) | $($f.Site) | $($f.VCenter) | $($f.Object) | $($f.Check) | $($f.Error)"
}
$LogLines | Out-File -FilePath $LogPath -Encoding UTF8
Write-Host "Log written: $LogPath" -ForegroundColor Cyan

Write-Host "`nDone. Findings: $($Global:AllResults.Count) | Failures: $($Global:FailureLog.Count) | Duration: $([Math]::Round(((Get-Date)-$ScriptStart).TotalMinutes,1)) min" -ForegroundColor Green
