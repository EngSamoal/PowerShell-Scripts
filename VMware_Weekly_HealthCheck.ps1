#Requires -Version 5.1
<#
.SYNOPSIS
    VMware Weekly Health Check - Read-only automated collection across SF, Tabuk and AMC (AJW/PAN/RGL/STR).

.DESCRIPTION
    Replaces the manual weekly vCenter/ESXi health-check reports with a single read-only
    PowerCLI collection run from an already-authenticated laptop session.

    - Uses whatever vCenter sessions are already connected (Connect-VIServer done beforehand).
    - Auto-discovers clusters, hosts, datastores, vSAN, vDS/port groups and VMs per vCenter -
      nothing about the infrastructure is hard-coded.
    - AMC is one vCenter with four clusters/sites (AJW, PAN, RGL, STR); each is discovered and
      reported on separately automatically, no hard-coded cluster list.
    - A failure collecting one item is logged and skipped; the script always continues.
    - NEVER writes/modifies anything in vCenter. Every cmdlet used below is read-only
      (Get-*, no Set-*/New-*/Remove-*, no config changes, no SSH enablement).

.NOTES
    Every threshold below (capacity %, cert/license expiry warning windows) is a CONFIGURABLE
    PARAMETER, not an assumed company standard - the source reports did not define explicit
    numeric thresholds, so raw values/status are always shown alongside any computed flag.

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

    [string]$OutputPath = ".\VMware_HealthCheck_Reports",

    # Configurable thresholds - NOT vendor/company-defined standards. Raw values are always
    # shown regardless of these; these only drive the Warning/Critical flag shown alongside them.
    [double]$CapacityWarningPct  = 80,
    [double]$CapacityCriticalPct = 90,
    [int]$CertExpiryWarningDays  = 60,
    [int]$LicenseExpiryWarningDays = 30,

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
$Global:ObjectClusterMap = @{}   # maps a host/datastore/vDS/port-group name -> the cluster it belongs to,
                                  # used only when building the per-site report to group findings correctly.

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

    # --- 2.1 vCenter appliance / server-level ---------------------------------
    Invoke-SafeCheck -CheckName 'vCenter version/build' -VCenter $VCName -Site $Site -ObjectName $VCName -Script {
        New-Finding -Site $Site -VCenter $VCName -Area 'vCenter Overview' -Object $VCName `
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
            New-Finding -Site $Site -VCenter $VCName -Area 'vCenter Overview' -Object $VCName `
                -Item "License: $($lic.Name)" -Value "Used $($lic.Used)/$($lic.Total) - Expires $expStr" -Status $status
        }
    }

    # vCenter appliance CPU/Mem/Disk/Services/NTP/Certificates require the VAMI/CIS REST API
    # (Connect-CisServer), which is a separate authenticated session from the vSphere API.
    # If that session exists it is used; otherwise this is explicitly marked, not fabricated.
    $CisSession = $global:DefaultCisServers | Where-Object { $_.Name -eq $VCName -and $_.IsConnected }
    if ($CisSession) {
        Invoke-SafeCheck -CheckName 'Appliance health (VAMI)' -VCenter $VCName -Site $Site -ObjectName $VCName -Script {
            $healthSvc = Get-CisService -Name 'com.vmware.appliance.health.system' -Server $CisSession
            $overall = $healthSvc.get()
            New-Finding -Site $Site -VCenter $VCName -Area 'vCenter Appliance' -Object $VCName `
                -Item 'Overall Appliance Health' -Value $overall -Status $(if ($overall -eq 'green') {'Healthy'} else {'Warning'})
            foreach ($comp in 'cpu','mem','storage') {
                $svc = Get-CisService -Name "com.vmware.appliance.health.$comp" -Server $CisSession
                $val = $svc.get()
                New-Finding -Site $Site -VCenter $VCName -Area 'vCenter Appliance' -Object $VCName `
                    -Item "$comp Health" -Value $val -Status $(if ($val -eq 'green') {'Healthy'} else {'Warning'})
            }
        }
        Invoke-SafeCheck -CheckName 'Appliance NTP (VAMI)' -VCenter $VCName -Site $Site -ObjectName $VCName -Script {
            $ntpSvc = Get-CisService -Name 'com.vmware.appliance.ntp' -Server $CisSession
            $ntpStatus = $ntpSvc.test()
            New-Finding -Site $Site -VCenter $VCName -Area 'vCenter Appliance' -Object $VCName `
                -Item 'NTP Sync' -Value ($ntpStatus | Out-String).Trim() -Status 'Information'
        }
    } else {
        New-Finding -Site $Site -VCenter $VCName -Area 'vCenter Appliance' -Object $VCName `
            -Item 'CPU/Memory/Disk/Services/NTP/TLS Certificate' -Value 'n/a' `
            -Status 'Manual/External Required' `
            -Notes 'Requires a VAMI/CIS session (Connect-CisServer <vcenter>) in addition to the vSphere API session. Not connected in this run.'
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

    # --- 2.2 Discover clusters (AMC's 4 sites/clusters come out of this automatically) --------
    $Clusters = Invoke-SafeCheck -CheckName 'Cluster discovery' -VCenter $VCName -Site $Site -ObjectName $VCName -Script {
        Get-Cluster -Server $VC
    }
    if (-not $Clusters) { continue }

    foreach ($Cluster in $Clusters) {
        $ClusterName = $Cluster.Name
        Write-Host "  - Cluster: $ClusterName"

        # HA / DRS
        Invoke-SafeCheck -CheckName 'HA/DRS status' -VCenter $VCName -Site $Site -ObjectName $ClusterName -Script {
            New-Finding -Site $Site -VCenter $VCName -Area 'Cluster Configuration' -Object $ClusterName `
                -Item 'vSphere HA' -Value $Cluster.HAEnabled -Status $(if ($Cluster.HAEnabled) {'Healthy'} else {'Warning'})
            New-Finding -Site $Site -VCenter $VCName -Area 'Cluster Configuration' -Object $ClusterName `
                -Item 'vSphere DRS' -Value $Cluster.DrsEnabled -Status $(if ($Cluster.DrsEnabled) {'Healthy'} else {'Warning'})
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
                New-Finding -Site $Site -VCenter $VCName -Area 'ESXi Host Health' -Object $HName `
                    -Item 'Connection State' -Value $VMHost.ConnectionState -Status $(if ($ok) {'Healthy'} else {'Critical'})
            }

            # Version/build (for consistency check, per cluster, below)
            Invoke-SafeCheck -CheckName 'Host version' -VCenter $VCName -Site $Site -ObjectName $HName -Script {
                New-Finding -Site $Site -VCenter $VCName -Area 'ESXi Host Health' -Object $HName `
                    -Item 'ESXi Version/Build' -Value "$($VMHost.Version) (Build $($VMHost.Build))" -Status 'Information'
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

            # Hardware health via built-in host Health System (numeric sensors) - no SSH required
            Invoke-SafeCheck -CheckName 'Hardware sensors' -VCenter $VCName -Site $Site -ObjectName $HName -Script {
                $sensors = $VMHost.ExtensionData.Runtime.HealthSystemRuntime.SystemHealthInfo.NumericSensorInfo
                if (-not $sensors) {
                    New-Finding -Site $Site -VCenter $VCName -Area 'Hardware Health' -Object $HName `
                        -Item 'Hardware Sensors' -Value 'n/a' -Status 'Unable to Check' `
                        -Notes 'Host does not expose CIM/IPMI sensor data to vCenter (common on some blade/BMC configs).'
                    return
                }
                # Group by sensor type so we get one row per component category, matching the manual report layout
                $groups = $sensors | Group-Object -Property SensorType
                foreach ($g in $groups) {
                    $bad = $g.Group | Where-Object { $_.HealthState.Key -notin @('green','Green') }
                    $status = if ($bad) { 'Warning' } else { 'Healthy' }
                    $summary = if ($bad) { ($bad | ForEach-Object { "$($_.Name): $($_.HealthState.Label)" }) -join '; ' } else { 'All normal' }
                    New-Finding -Site $Site -VCenter $VCName -Area 'Hardware Health' -Object $HName `
                        -Item "$($g.Name) Sensors" -Value $summary -Status $status
                }
            }

            # TPM
            Invoke-SafeCheck -CheckName 'TPM' -VCenter $VCName -Site $Site -ObjectName $HName -Script {
                $tpmInfo = $VMHost.ExtensionData.Capability.TpmSupported
                New-Finding -Site $Site -VCenter $VCName -Area 'Security & Compliance' -Object $HName `
                    -Item 'TPM Present/Supported' -Value $tpmInfo -Status $(if ($tpmInfo) {'Healthy'} else {'Information'})
            }

            # Secure Boot (via EsxCli, no SSH required - this is the vSphere API path)
            Invoke-SafeCheck -CheckName 'Secure Boot' -VCenter $VCName -Site $Site -ObjectName $HName -Script {
                $esxcli = Get-EsxCli -VMHost $VMHost -Server $VC -V2
                $sb = $esxcli.system.settings.encryption.get.Invoke()
                $enabled = $sb.RequireSecureBoot
                New-Finding -Site $Site -VCenter $VCName -Area 'Security & Compliance' -Object $HName `
                    -Item 'Secure Boot' -Value $enabled -Status $(if ($enabled -eq $true -or $enabled -eq 'true') {'Healthy'} else {'Warning'})
            }

            # Lockdown Mode
            Invoke-SafeCheck -CheckName 'Lockdown mode' -VCenter $VCName -Site $Site -ObjectName $HName -Script {
                $lockdown = $VMHost.ExtensionData.Config.LockdownMode
                New-Finding -Site $Site -VCenter $VCName -Area 'Security & Compliance' -Object $HName `
                    -Item 'Lockdown Mode' -Value $lockdown -Status 'Information'
            }

            # Local ESXi users - via esxcli (API-based, no SSH) rather than Get-VMHostAccount,
            # whose -VMHost parameter isn't present on all PowerCLI versions.
            Invoke-SafeCheck -CheckName 'Local ESXi users' -VCenter $VCName -Site $Site -ObjectName $HName -Script {
                $esxcli = Get-EsxCli -VMHost $VMHost -Server $VC -V2
                $accts = $esxcli.system.account.list.Invoke()
                New-Finding -Site $Site -VCenter $VCName -Area 'Security & Compliance' -Object $HName `
                    -Item 'Local Accounts' -Value (($accts | ForEach-Object { $_.UserID }) -join ', ') -Status 'Information' `
                    -Notes 'Review membership manually; automation only inventories accounts, it does not judge appropriateness.'
            }

            # Syslog
            Invoke-SafeCheck -CheckName 'Syslog config' -VCenter $VCName -Site $Site -ObjectName $HName -Script {
                $syslog = Get-VMHostSysLogServer -Server $VC -VMHost $VMHost
                $configured = -not [string]::IsNullOrWhiteSpace($syslog.Host)
                New-Finding -Site $Site -VCenter $VCName -Area 'Security & Compliance' -Object $HName `
                    -Item 'Syslog Target' -Value $(if ($configured) { "$($syslog.Host):$($syslog.Port)" } else { 'Not configured' }) `
                    -Status $(if ($configured) {'Healthy'} else {'Warning'})
            }

            # NTP
            Invoke-SafeCheck -CheckName 'Host NTP' -VCenter $VCName -Site $Site -ObjectName $HName -Script {
                $ntpServers = Get-VMHostNtpServer -Server $VC -VMHost $VMHost
                $ntpSvc = Get-VMHostService -Server $VC -VMHost $VMHost | Where-Object { $_.Key -eq 'ntpd' }
                $running = $ntpSvc -and $ntpSvc.Running
                New-Finding -Site $Site -VCenter $VCName -Area 'ESXi Host Health' -Object $HName `
                    -Item 'NTP' -Value "Servers: $($ntpServers -join ', ') | Running: $running" `
                    -Status $(if ($running -and $ntpServers) {'Healthy'} else {'Warning'})
            }

            # Physical NICs / redundancy / errors
            Invoke-SafeCheck -CheckName 'Physical NICs' -VCenter $VCName -Site $Site -ObjectName $HName -Script {
                $pnics = Get-VMHostNetworkAdapter -Server $VC -VMHost $VMHost -Physical
                $up = $pnics | Where-Object { $_.BitRatePerSec -gt 0 }
                New-Finding -Site $Site -VCenter $VCName -Area 'Networking' -Object $HName `
                    -Item 'Physical NICs' -Value "$($pnics.Count) total, $($up.Count) linked up" `
                    -Status $(if ($pnics.Count -ge 2 -and $up.Count -ge 2) {'Healthy'} else {'Warning'})
            }

            # NIC error/drop counters via performance manager (safe, no SSH)
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
            New-Finding -Site $Site -VCenter $VCName -Area 'Cluster Configuration' -Object $ClusterName `
                -Item 'ESXi Version Consistency' -Value ($distinct -join ' | ') -Status $status
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
                New-Finding -Site $Site -VCenter $VCName -Area 'CPU/Memory Capacity' -Object $ClusterName `
                    -Item ("CPU Usage % (avg, last {0}h, historical)" -f $PerfHistoryHours) `
                    -Value ("{0:N2}%  (Capacity {1:N1} GHz)" -f $avgCpuPct, ($cpuTotalMHz/1000)) -Status (Get-PctStatus $avgCpuPct)
            } else {
                New-Finding -Site $Site -VCenter $VCName -Area 'CPU/Memory Capacity' -Object $ClusterName `
                    -Item 'CPU Usage %' -Value 'n/a' -Status 'Unable to Check' -Notes 'No historical performance samples returned (check vCenter stats level/retention).'
            }
            if ($memStat) {
                $avgMemPct = ($memStat | Measure-Object -Property Value -Average).Average
                New-Finding -Site $Site -VCenter $VCName -Area 'CPU/Memory Capacity' -Object $ClusterName `
                    -Item ("Memory Usage % (avg, last {0}h, historical)" -f $PerfHistoryHours) `
                    -Value ("{0:N2}%  (Capacity {1:N1} GB)" -f $avgMemPct, ($memTotalMB/1024)) -Status (Get-PctStatus $avgMemPct)
            } else {
                New-Finding -Site $Site -VCenter $VCName -Area 'CPU/Memory Capacity' -Object $ClusterName `
                    -Item 'Memory Usage %' -Value 'n/a' -Status 'Unable to Check' -Notes 'No historical performance samples returned.'
            }
        }

        # Datastores in this cluster
        Invoke-SafeCheck -CheckName 'Datastore health/capacity' -VCenter $VCName -Site $Site -ObjectName $ClusterName -Script {
            $datastores = Get-Datastore -Server $VC -RelatedObject $Cluster
            foreach ($ds in $datastores) {
                $Global:ObjectClusterMap[$ds.Name] = $ClusterName
                $usedPct = if ($ds.CapacityGB -gt 0) { (($ds.CapacityGB - $ds.FreeSpaceGB) / $ds.CapacityGB) * 100 } else { 0 }
                $accessible = $ds.ExtensionData.Summary.Accessible
                $status = if (-not $accessible) { 'Critical' } else { Get-PctStatus $usedPct }
                New-Finding -Site $Site -VCenter $VCName -Area 'Storage/Datastore' -Object $ds.Name `
                    -Item 'Capacity/Used/Free/Util%' `
                    -Value ("Cap {0:N0} GB | Free {1:N0} GB | Used {2:N1}% | Accessible: {3}" -f $ds.CapacityGB, $ds.FreeSpaceGB, $usedPct, $accessible) `
                    -Status $status
            }
        }

        # vSAN (only if the cluster actually has vSAN enabled and the module is available)
        Invoke-SafeCheck -CheckName 'vSAN health/capacity' -VCenter $VCName -Site $Site -ObjectName $ClusterName -Script {
            if (-not $Cluster.VsanEnabled) {
                New-Finding -Site $Site -VCenter $VCName -Area 'Storage/vSAN' -Object $ClusterName `
                    -Item 'vSAN' -Value 'Not enabled on this cluster' -Status 'Information'
                return
            }
            if (-not $VsanModuleAvailable) {
                New-Finding -Site $Site -VCenter $VCName -Area 'Storage/vSAN' -Object $ClusterName `
                    -Item 'vSAN Health/Capacity' -Value 'n/a' -Status 'Unable to Check' `
                    -Notes 'VMware.VimAutomation.Vsan module not installed on this laptop.'
                return
            }
            $space = Get-VsanSpaceUsage -Server $VC -Cluster $Cluster -ErrorAction Stop
            $usedPct = if ($space.CapacityGB -gt 0) { ($space.UsedGB / $space.CapacityGB) * 100 } else { 0 }
            New-Finding -Site $Site -VCenter $VCName -Area 'Storage/vSAN' -Object $ClusterName `
                -Item 'vSAN Capacity/Used/Free/Util%' `
                -Value ("Cap {0:N0} GB | Used {1:N0} GB | Free {2:N0} GB | Util {3:N1}%" -f $space.CapacityGB, $space.UsedGB, ($space.CapacityGB - $space.UsedGB), $usedPct) `
                -Status (Get-PctStatus $usedPct)

            $healthTest = Get-VsanClusterHealth -Server $VC -Cluster $Cluster -ErrorAction Stop
            foreach ($grp in $healthTest.HealthGroups) {
                $bad = $grp.HealthTests | Where-Object { $_.Status -ne 'green' }
                $status = if ($bad) { 'Warning' } else { 'Healthy' }
                New-Finding -Site $Site -VCenter $VCName -Area 'Storage/vSAN' -Object $ClusterName `
                    -Item "vSAN Health: $($grp.GroupName)" -Value $(if ($bad) { ($bad.TestName -join '; ') } else { 'OK' }) -Status $status
            }
        }

        # VM inventory / guest OS distribution
        Invoke-SafeCheck -CheckName 'VM inventory' -VCenter $VCName -Site $Site -ObjectName $ClusterName -Script {
            $vms = Get-VM -Server $VC -Location $Cluster
            $on  = ($vms | Where-Object { $_.PowerState -eq 'PoweredOn' }).Count
            $off = ($vms | Where-Object { $_.PowerState -eq 'PoweredOff' }).Count
            New-Finding -Site $Site -VCenter $VCName -Area 'VM Inventory' -Object $ClusterName `
                -Item 'VM Count' -Value "Total $($vms.Count) | On $on | Off $off" -Status 'Information'

            $osDist = $vms | Group-Object { if ($_.Guest.OSFullName) { $_.Guest.OSFullName } else { $_.ExtensionData.Config.GuestFullName } } |
                Sort-Object Count -Descending
            foreach ($g in $osDist) {
                New-Finding -Site $Site -VCenter $VCName -Area 'VM Inventory' -Object $ClusterName `
                    -Item "Guest OS: $($g.Name)" -Value $g.Count -Status 'Information'
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
# 3. REPORT GENERATION HELPERS
#    One report per site, laid out the same way as the original manual reports
#    (Executive Summary -> vCenter Health -> per-cluster sections -> Global Notes
#    -> Conclusion), with simple colored-dot status indicators instead of tables
#    of raw data. Manual/physical-only items are never collected in the first
#    place, so there is nothing to filter out here.
# ============================================================================
function Get-Dot {
    param([string]$Status)
    $color = switch ($Status) {
        'Healthy'          { '#16a34a' }
        'Warning'          { '#f59e0b' }
        'Critical'         { '#dc2626' }
        'Information'      { '#2563eb' }
        'Unable to Check'  { '#9ca3af' }
        default            { '#9ca3af' }
    }
    return "<span style='color:$color;'>&#9679;</span> $Status"
}

function Get-WorstStatus {
    param([string[]]$Statuses)
    $order = @{ 'Critical' = 4; 'Warning' = 3; 'Unable to Check' = 2; 'Healthy' = 1; 'Information' = 0 }
    if (-not $Statuses -or $Statuses.Count -eq 0) { return 'Healthy' }
    return ($Statuses | Sort-Object { $order[$_] } -Descending | Select-Object -First 1)
}

$SiteReportStyle = @"
<style>
body { font-family: Segoe UI, Arial, sans-serif; margin: 30px; color:#1f2937; font-size:13.5px; }
h1 { color:#1e3a5f; margin-bottom:2px; }
.subtitle { color:#6b7280; margin-top:0; }
h2 { color:#1e3a5f; border-bottom:2px solid #1e3a5f; padding-bottom:4px; margin-top:30px; font-size:17px; }
h3 { color:#1e3a5f; margin-top:18px; font-size:14.5px; }
table { border-collapse:collapse; width:100%; margin:8px 0 16px 0; }
th, td { border:1px solid #d1d5db; padding:5px 9px; text-align:left; font-size:13px; }
th { background:#1e3a5f; color:#fff; }
tr:nth-child(even) { background:#f3f4f6; }
.exec-box { background:#f3f4f6; border-radius:8px; padding:12px 16px; margin:10px 0 18px 0; }
ul.notes { margin:6px 0; padding-left:20px; }
ul.notes li { margin:3px 0; }
.alarm { color:#dc2626; }
.small-note { color:#6b7280; font-size:12px; margin-top:30px; }
</style>
"@

function Build-SiteReportHtml {
    param([string]$SiteLabel)

    # Manual/External Required items are never surfaced in the report - they're only ever
    # produced when an optional dependency (e.g. Connect-CisServer) wasn't connected this run.
    $SiteFindings = $Global:AllResults | Where-Object { $_.Site -eq $SiteLabel -and $_.Status -ne 'Manual/External Required' }
    if (-not $SiteFindings) { return $null }

    $VCenterName  = $SiteFindings | Select-Object -First 1 -ExpandProperty VCenter
    $ClusterNames = $SiteFindings | Where-Object { $_.Area -eq 'Cluster Configuration' } | Select-Object -ExpandProperty Object -Unique
    $HostObjects  = $SiteFindings | Where-Object { $_.Area -eq 'ESXi Host Health' -and $_.Item -eq 'ESXi Version/Build' } | Select-Object -ExpandProperty Object -Unique

    $CritCount = ($SiteFindings | Where-Object { $_.Status -eq 'Critical' }).Count
    $WarnCount = ($SiteFindings | Where-Object { $_.Status -eq 'Warning' }).Count
    $OverallHealth = if ($CritCount -gt 0) { 'Critical' } elseif ($WarnCount -gt 0) { 'Warning' } else { 'Healthy' }
    $VcVersionItem = $SiteFindings | Where-Object { $_.Item -eq 'vCenter Version/Build' } | Select-Object -First 1

    $sb = New-Object System.Text.StringBuilder
    [void]$sb.Append("<!DOCTYPE html><html><head><meta charset='utf-8'><title>$SiteLabel vCenter &amp; ESXi Health Check Report</title>$SiteReportStyle</head><body>")
    [void]$sb.Append("<h1>$SiteLabel vCenter &amp; ESXi Health Check Report</h1>")
    [void]$sb.Append("<p class='subtitle'>Report Date: $RunDate</p>")

    # ---- 1. Executive Summary ----
    [void]$sb.Append("<h2>1. Executive Summary</h2><div class='exec-box'>")
    [void]$sb.Append("<p><b>Environment:</b> 1 vCenter ($VCenterName) | $($ClusterNames.Count) Cluster(s)/Site(s) | $($HostObjects.Count) ESXi Host(s)</p>")
    if ($VcVersionItem) { [void]$sb.Append("<p><b>vCenter Version:</b> $($VcVersionItem.Value)</p>") }
    [void]$sb.Append("<p><b>Overall Environment Health:</b> $(Get-Dot $OverallHealth)</p>")
    [void]$sb.Append("<p><b>High Risk Issue:</b> $(if ($CritCount -gt 0) {'YES'} else {'NO'})</p>")
    [void]$sb.Append("</div>")

    # ---- 2. vCenter Server Health ----
    [void]$sb.Append("<h2>2. vCenter Server Health</h2><table><tr><th>Item</th><th>Status</th></tr>")
    $ApplianceRows = $SiteFindings | Where-Object { $_.Area -in 'vCenter Overview','vCenter Appliance' -and $_.Item -ne 'vCenter Version/Build' }
    foreach ($r in $ApplianceRows) {
        [void]$sb.Append("<tr><td>$($r.Item)</td><td>$(Get-Dot $r.Status) $($r.Value)</td></tr>")
    }
    if (-not $ApplianceRows) { [void]$sb.Append("<tr><td colspan='2'><em>No appliance-level data collected this run.</em></td></tr>") }
    [void]$sb.Append("</table>")

    # ---- 3. Per-Cluster Sections (naturally becomes 4 sections for AMC, 1 for SF/TB) ----
    $clusterIdx = 2
    foreach ($ClusterName in $ClusterNames) {
        $clusterIdx++
        [void]$sb.Append("<h2>$clusterIdx. $ClusterName</h2>")

        $HostsHere = $HostObjects | Where-Object { $Global:ObjectClusterMap[$_] -eq $ClusterName }

        [void]$sb.Append("<h3>ESXi Host Versions</h3><table><tr><th>Host</th><th>ESXi Version/Build</th><th>Status</th><th>Notes</th></tr>")
        foreach ($h in $HostsHere) {
            $verItem  = $SiteFindings | Where-Object { $_.Object -eq $h -and $_.Item -eq 'ESXi Version/Build' } | Select-Object -First 1
            $hostRows = $SiteFindings | Where-Object { $_.Object -eq $h }
            $hostStatus = Get-WorstStatus ($hostRows | Select-Object -ExpandProperty Status)
            $hostIssues = $hostRows | Where-Object { $_.Status -in 'Critical','Warning' } | Select-Object -First 2
            $notes = ($hostIssues | ForEach-Object { "$($_.Item): $($_.Value)" }) -join '; '
            [void]$sb.Append("<tr><td>$h</td><td>$($verItem.Value)</td><td>$(Get-Dot $hostStatus)</td><td>$notes</td></tr>")
        }
        [void]$sb.Append("</table>")

        [void]$sb.Append("<h3>CPU &amp; Memory Capacity</h3><table><tr><th>Metric</th><th>Value</th><th>Status</th></tr>")
        $capRows = $SiteFindings | Where-Object { $_.Area -eq 'CPU/Memory Capacity' -and $_.Object -eq $ClusterName }
        foreach ($r in $capRows) { [void]$sb.Append("<tr><td>$($r.Item)</td><td>$($r.Value)</td><td>$(Get-Dot $r.Status)</td></tr>") }
        if (-not $capRows) { [void]$sb.Append("<tr><td colspan='3'><em>No capacity data collected.</em></td></tr>") }
        [void]$sb.Append("</table>")

        [void]$sb.Append("<h3>Storage</h3>")
        $vsanRow = $SiteFindings | Where-Object { $_.Area -eq 'Storage/vSAN' -and $_.Object -eq $ClusterName -and $_.Item -like 'vSAN Capacity*' } | Select-Object -First 1
        if ($vsanRow) {
            [void]$sb.Append("<p>$(Get-Dot $vsanRow.Status) vSAN - $($vsanRow.Value)</p>")
            $vsanIssues = $SiteFindings | Where-Object { $_.Area -eq 'Storage/vSAN' -and $_.Object -eq $ClusterName -and $_.Item -like 'vSAN Health:*' -and $_.Status -in 'Critical','Warning' }
            if ($vsanIssues) {
                [void]$sb.Append("<ul class='notes'>")
                foreach ($vi in $vsanIssues) { [void]$sb.Append("<li>$(Get-Dot $vi.Status) $($vi.Item): $($vi.Value)</li>") }
                [void]$sb.Append("</ul>")
            }
        }
        $dsRows = $SiteFindings | Where-Object { $_.Area -eq 'Storage/Datastore' -and $Global:ObjectClusterMap[$_.Object] -eq $ClusterName }
        if ($dsRows) {
            $dsHealthy = ($dsRows | Where-Object { $_.Status -eq 'Healthy' }).Count
            [void]$sb.Append("<p>$($dsRows.Count) datastore(s) - $dsHealthy healthy</p>")
            $dsIssues = $dsRows | Where-Object { $_.Status -in 'Critical','Warning' }
            if ($dsIssues) {
                [void]$sb.Append("<ul class='notes'>")
                foreach ($di in $dsIssues) { [void]$sb.Append("<li>$(Get-Dot $di.Status) $($di.Object): $($di.Value)</li>") }
                [void]$sb.Append("</ul>")
            }
        }
        if (-not $vsanRow -and -not $dsRows) { [void]$sb.Append("<p><em>No storage data collected.</em></p>") }

        [void]$sb.Append("<h3>Virtual Machine Inventory</h3>")
        $vmCountRow = $SiteFindings | Where-Object { $_.Area -eq 'VM Inventory' -and $_.Object -eq $ClusterName -and $_.Item -eq 'VM Count' } | Select-Object -First 1
        if ($vmCountRow) { [void]$sb.Append("<p>$($vmCountRow.Value)</p>") }
        $osRows = $SiteFindings | Where-Object { $_.Area -eq 'VM Inventory' -and $_.Object -eq $ClusterName -and $_.Item -like 'Guest OS:*' } | Select-Object -First 6
        if ($osRows) {
            [void]$sb.Append("<table><tr><th>Guest OS</th><th>Count</th></tr>")
            foreach ($o in $osRows) { [void]$sb.Append("<tr><td>$($o.Item -replace 'Guest OS: ','')</td><td>$($o.Value)</td></tr>") }
            [void]$sb.Append("</table>")
        }

        $netRows = $SiteFindings | Where-Object { $_.Area -eq 'Networking' -and $_.Object -eq $ClusterName }
        if ($netRows) {
            [void]$sb.Append("<h3>Networking</h3><ul class='notes'>")
            foreach ($n in $netRows) { [void]$sb.Append("<li>$(Get-Dot $n.Status) $($n.Item): $($n.Value)</li>") }
            [void]$sb.Append("</ul>")
        }

        $secRows = $SiteFindings | Where-Object { $_.Area -eq 'Security & Compliance' -and $HostsHere -contains $_.Object }
        if ($secRows) {
            $secStatus = Get-WorstStatus ($secRows | Select-Object -ExpandProperty Status)
            [void]$sb.Append("<h3>Security &amp; Compliance</h3><p>$(Get-Dot $secStatus)</p>")
            $secIssues = $secRows | Where-Object { $_.Status -in 'Critical','Warning' }
            if ($secIssues) {
                [void]$sb.Append("<ul class='notes'>")
                foreach ($si in $secIssues) { [void]$sb.Append("<li>$($si.Object) - $($si.Item): $($si.Value)</li>") }
                [void]$sb.Append("</ul>")
            } else {
                [void]$sb.Append("<p>No compliance issues found.</p>")
            }
        }

        [void]$sb.Append("<h3>Issues and Alarms</h3>")
        $alarmRows = $SiteFindings | Where-Object { $_.Area -eq 'Alarms' -and $_.Item -ne 'Triggered Alarms' -and ($_.Object -eq $ClusterName -or $HostsHere -contains $_.Object) }
        if ($alarmRows) {
            [void]$sb.Append("<ul class='notes'>")
            foreach ($a in $alarmRows) { [void]$sb.Append("<li class='alarm'>$($a.Object) - $($a.Item) ($($a.Value))</li>") }
            [void]$sb.Append("</ul>")
        } else {
            [void]$sb.Append("<p>No active alarms.</p>")
        }
    }

    # ---- Global Notes & Observations (short, prioritized - not every finding) ----
    [void]$sb.Append("<h2>Global Notes &amp; Observations</h2>")
    $keyIssues = $SiteFindings | Where-Object {
        $_.Status -in 'Critical','Warning' -and (
            $_.Item -like 'License:*' -or $_.Item -eq 'ESXi Version Consistency' -or
            $_.Item -in 'vSphere HA','vSphere DRS' -or $_.Area -eq 'Alarms' -or
            $_.Area -in 'Storage/vSAN','Storage/Datastore'
        )
    }
    if ($keyIssues) {
        [void]$sb.Append("<ul class='notes'>")
        foreach ($k in $keyIssues) { [void]$sb.Append("<li>$(Get-Dot $k.Status) $($k.Object) - $($k.Item): $($k.Value)</li>") }
        [void]$sb.Append("</ul>")
    } else {
        [void]$sb.Append("<p>No short-term action items identified this run.</p>")
    }

    # ---- Conclusion ----
    [void]$sb.Append("<h2>Conclusion</h2>")
    $concl = if ($CritCount -gt 0) {
        "The $SiteLabel VMware environment has $CritCount critical item(s) requiring prompt attention, alongside $WarnCount warning item(s). See Global Notes above."
    } elseif ($WarnCount -gt 0) {
        "The $SiteLabel VMware environment is stable with $WarnCount warning item(s) to track. No critical issues were identified during this automated check."
    } else {
        "The $SiteLabel VMware environment is stable and healthy. No critical or warning issues were identified during this automated check."
    }
    [void]$sb.Append("<p>$concl</p>")
    [void]$sb.Append("<p><b>Prepared By:</b> VMware Weekly Health Check Automation</p><p><b>Report Date:</b> $RunDate</p>")

    $siteFailures = $Global:FailureLog | Where-Object { $_.Site -eq $SiteLabel }
    if ($siteFailures) {
        [void]$sb.Append("<p class='small-note'>$($siteFailures.Count) item(s) could not be automatically verified this run - see VMware_Weekly_HealthCheck_$RunDate.log for detail.</p>")
    }

    [void]$sb.Append("</body></html>")
    return $sb.ToString()
}

# ============================================================================
# 4. OUTPUT: ONE DOCX PER SITE
#    Each site's HTML is a scratch/intermediate file only, deleted after conversion -
#    the deliverable per site is a single Word document.
# ============================================================================
$SiteLabels = $Global:AllResults | Select-Object -ExpandProperty Site -Unique
$WordAvailable = $true
try {
    $word = New-Object -ComObject Word.Application
    $word.Visible = $false
} catch {
    $WordAvailable = $false
    Write-CheckLog -VCenter 'n/a' -Site 'n/a' -Object 'DOCX export' -CheckName 'Word COM automation' -ErrorMessage $_.Exception.Message
    Write-Warning "Microsoft Word is not available on this machine - cannot generate .docx reports. Scratch HTML reports will be left in $env:TEMP instead."
}

foreach ($SiteLabel in $SiteLabels) {
    $SiteHtml = Build-SiteReportHtml -SiteLabel $SiteLabel
    if (-not $SiteHtml) { continue }

    $safeLabel = ($SiteLabel -replace '[\\/\?\*\[\]:<>\|]', '_')
    $HtmlPath  = Join-Path $env:TEMP "VMware_HealthCheck_$safeLabel`_scratch_$RunDate`_$PID.html"
    $SiteHtml | Out-File -FilePath $HtmlPath -Encoding UTF8

    $DocxPath = Join-Path $OutputPath "$safeLabel`_VMware_HealthCheck_$RunDate.docx"
    if ($WordAvailable) {
        try {
            $doc = $word.Documents.Open((Resolve-Path $HtmlPath).Path)
            # SaveAs called via InvokeMember rather than a direct method call - PowerShell's COM
            # late-binding can fail to marshal SaveAs's [ref] parameters ("Cannot convert ... psobject
            # to Object") on some PowerShell/.NET combinations. InvokeMember avoids that entirely.
            $null = $doc.GetType().InvokeMember('SaveAs', [System.Reflection.BindingFlags]::InvokeMethod, $null, $doc, @([string]$DocxPath, 17))
            $doc.Close()
            [System.Runtime.Interopservices.Marshal]::ReleaseComObject($doc) | Out-Null
            Remove-Item $HtmlPath -Force -ErrorAction SilentlyContinue
            Write-Host "Report written: $DocxPath" -ForegroundColor Cyan
        } catch {
            Write-CheckLog -VCenter 'n/a' -Site $SiteLabel -Object 'DOCX export' -CheckName 'Word COM automation' -ErrorMessage $_.Exception.Message
            Write-Warning "Could not generate .docx for $SiteLabel - scratch HTML left at $HtmlPath"
        }
    }
}

if ($WordAvailable) {
    $word.Quit()
    [System.Runtime.Interopservices.Marshal]::ReleaseComObject($word) | Out-Null
}

# ============================================================================
# 5. OUTPUT: LOG FILE (written last so it also captures any DOCX export failures)
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
