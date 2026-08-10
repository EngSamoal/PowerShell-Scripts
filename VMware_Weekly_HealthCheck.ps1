https://rhpds.github.io/ocp-virt-roadshow-2026-showroom/modules/index.html
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

    [string]$OutputPath = (Join-Path $PSScriptRoot "VMware_HealthCheck_Reports"),

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
# ============================================================================
# 3. REPORT GENERATION HELPERS
#    Each site's .docx is built DIRECTLY in Word via COM automation - typing text,
#    applying styles/colors, and inserting native Word tables. There is no HTML
#    step anywhere in this pipeline (opening an HTML file in Word and re-saving
#    it is what caused the earlier failures - Word treats that as a web/compat
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

function Write-SiteReportDocx {
    param($Word, [string]$SiteLabel, [string]$OutputPath, [string]$RunDate)

    $SiteFindings = $Global:AllResults | Where-Object { $_.Site -eq $SiteLabel -and $_.Status -ne 'Manual/External Required' }
    if (-not $SiteFindings) { return }

    $VCenterName  = $SiteFindings | Select-Object -First 1 -ExpandProperty VCenter
    $ClusterNames = $SiteFindings | Where-Object { $_.Area -eq 'Cluster Configuration' } | Select-Object -ExpandProperty Object -Unique
    $HostObjects  = $SiteFindings | Where-Object { $_.Area -eq 'ESXi Host Health' -and $_.Item -eq 'ESXi Version/Build' } | Select-Object -ExpandProperty Object -Unique
    $CritCount = ($SiteFindings | Where-Object { $_.Status -eq 'Critical' }).Count
    $WarnCount = ($SiteFindings | Where-Object { $_.Status -eq 'Warning' }).Count
    $OverallHealth = if ($CritCount -gt 0) { 'Critical' } elseif ($WarnCount -gt 0) { 'Warning' } else { 'Healthy' }
    $VcVersionItem = $SiteFindings | Where-Object { $_.Item -eq 'vCenter Version/Build' } | Select-Object -First 1

    $doc = $Word.Documents.Add()
    $sel = $Word.Selection

    Write-Heading $sel "$SiteLabel vCenter & ESXi Health Check Report" 1
    Write-Para $sel "Report Date: $RunDate" -Gray

    Write-Heading $sel "1. Executive Summary" 2
    Write-Para $sel "Environment: 1 vCenter ($VCenterName) | $($ClusterNames.Count) Cluster(s)/Site(s) | $($HostObjects.Count) ESXi Host(s)"
    if ($VcVersionItem) { Write-Para $sel "vCenter Version: $($VcVersionItem.Value)" }
    Write-DotLine $sel $OverallHealth "Overall Environment Health: $OverallHealth"
    Write-Para $sel "High Risk Issue: $(if ($CritCount -gt 0) {'YES'} else {'NO'})"

    Write-Heading $sel "2. vCenter Server Health" 2
    $ApplianceRows = $SiteFindings | Where-Object { $_.Area -in 'vCenter Overview','vCenter Appliance' -and $_.Item -ne 'vCenter Version/Build' }
    if ($ApplianceRows) {
        $tableRows = @()
        foreach ($r in $ApplianceRows) { $tableRows += ,@($r.Item, $r.Value, $r.Status) }
        Add-Table $doc $sel @('Item','Value','Status') $tableRows -StatusColumnIndex 2
    } else {
        Write-Para $sel "No appliance-level data collected this run." -Italic
    }

    $clusterIdx = 2
    foreach ($ClusterName in $ClusterNames) {
        $clusterIdx++
        Write-Heading $sel "$clusterIdx. $ClusterName" 2
        $HostsHere = $HostObjects | Where-Object { $Global:ObjectClusterMap[$_] -eq $ClusterName }

        Write-Heading $sel "ESXi Host Versions" 3
        $hostTableRows = @()
        foreach ($h in $HostsHere) {
            $verItem  = $SiteFindings | Where-Object { $_.Object -eq $h -and $_.Item -eq 'ESXi Version/Build' } | Select-Object -First 1
            $hostRows = $SiteFindings | Where-Object { $_.Object -eq $h }
            $hostStatus = Get-WorstStatus ($hostRows | Select-Object -ExpandProperty Status)
            $hostIssues = $hostRows | Where-Object { $_.Status -in 'Critical','Warning' } | Select-Object -First 2
            $notes = ($hostIssues | ForEach-Object { "$($_.Item): $($_.Value)" }) -join '; '
            $hostTableRows += ,@($h, $verItem.Value, $hostStatus, $notes)
        }
        if ($hostTableRows.Count -gt 0) {
            Add-Table $doc $sel @('Host','ESXi Version/Build','Status','Notes') $hostTableRows -StatusColumnIndex 2
        } else {
            Write-Para $sel "No hosts discovered for this cluster." -Italic
        }

        Write-Heading $sel "CPU & Memory Capacity" 3
        $capRows = $SiteFindings | Where-Object { $_.Area -eq 'CPU/Memory Capacity' -and $_.Object -eq $ClusterName }
        if ($capRows) {
            $capTableRows = @()
            foreach ($r in $capRows) { $capTableRows += ,@($r.Item, $r.Value, $r.Status) }
            Add-Table $doc $sel @('Metric','Value','Status') $capTableRows -StatusColumnIndex 2
        } else {
            Write-Para $sel "No capacity data collected." -Italic
        }

        Write-Heading $sel "Storage" 3
        $vsanRow = $SiteFindings | Where-Object { $_.Area -eq 'Storage/vSAN' -and $_.Object -eq $ClusterName -and $_.Item -like 'vSAN Capacity*' } | Select-Object -First 1
        $dsRows  = $SiteFindings | Where-Object { $_.Area -eq 'Storage/Datastore' -and $Global:ObjectClusterMap[$_.Object] -eq $ClusterName }
        if ($vsanRow) {
            Write-DotLine $sel $vsanRow.Status "vSAN - $($vsanRow.Value)"
            $vsanIssues = $SiteFindings | Where-Object { $_.Area -eq 'Storage/vSAN' -and $_.Object -eq $ClusterName -and $_.Item -like 'vSAN Health:*' -and $_.Status -in 'Critical','Warning' }
            foreach ($vi in $vsanIssues) { Write-DotLine $sel $vi.Status "$($vi.Item): $($vi.Value)" }
        }
        if ($dsRows) {
            $dsHealthy = ($dsRows | Where-Object { $_.Status -eq 'Healthy' }).Count
            Write-Para $sel "$($dsRows.Count) datastore(s) - $dsHealthy healthy"
            $dsIssues = $dsRows | Where-Object { $_.Status -in 'Critical','Warning' }
            foreach ($di in $dsIssues) { Write-DotLine $sel $di.Status "$($di.Object): $($di.Value)" }
        }
        if (-not $vsanRow -and -not $dsRows) { Write-Para $sel "No storage data collected." -Italic }

        Write-Heading $sel "Virtual Machine Inventory" 3
        $vmCountRow = $SiteFindings | Where-Object { $_.Area -eq 'VM Inventory' -and $_.Object -eq $ClusterName -and $_.Item -eq 'VM Count' } | Select-Object -First 1
        if ($vmCountRow) { Write-Para $sel $vmCountRow.Value }
        $osRows = $SiteFindings | Where-Object { $_.Area -eq 'VM Inventory' -and $_.Object -eq $ClusterName -and $_.Item -like 'Guest OS:*' } | Select-Object -First 6
        if ($osRows) {
            $osTableRows = @()
            foreach ($o in $osRows) { $osTableRows += ,@(($o.Item -replace 'Guest OS: ',''), $o.Value) }
            Add-Table $doc $sel @('Guest OS','Count') $osTableRows
        }

        $netRows = $SiteFindings | Where-Object { $_.Area -eq 'Networking' -and $_.Object -eq $ClusterName }
        if ($netRows) {
            Write-Heading $sel "Networking" 3
            foreach ($n in $netRows) { Write-DotLine $sel $n.Status "$($n.Item): $($n.Value)" }
        }

        $secRows = $SiteFindings | Where-Object { $_.Area -eq 'Security & Compliance' -and $HostsHere -contains $_.Object }
        if ($secRows) {
            Write-Heading $sel "Security & Compliance" 3
            $secStatus = Get-WorstStatus ($secRows | Select-Object -ExpandProperty Status)
            Write-DotLine $sel $secStatus "Overall: $secStatus"
            $secIssues = $secRows | Where-Object { $_.Status -in 'Critical','Warning' }
            if ($secIssues) {
                foreach ($si in $secIssues) { Write-Para $sel "$($si.Object) - $($si.Item): $($si.Value)" }
            } else {
                Write-Para $sel "No compliance issues found."
            }
        }

        Write-Heading $sel "Issues and Alarms" 3
        $alarmRows = $SiteFindings | Where-Object { $_.Area -eq 'Alarms' -and $_.Item -ne 'Triggered Alarms' -and ($_.Object -eq $ClusterName -or $HostsHere -contains $_.Object) }
        if ($alarmRows) {
            $alarmLines = $alarmRows | ForEach-Object { "$($_.Object) - $($_.Item) ($($_.Value))" }
            Write-BulletList $sel $alarmLines -Red
        } else {
            Write-Para $sel "No active alarms."
        }
    }

    Write-Heading $sel "Global Notes & Observations" 2
    $keyIssues = $SiteFindings | Where-Object {
        $_.Status -in 'Critical','Warning' -and (
            $_.Item -like 'License:*' -or $_.Item -eq 'ESXi Version Consistency' -or
            $_.Item -in 'vSphere HA','vSphere DRS' -or $_.Area -eq 'Alarms' -or
            $_.Area -in 'Storage/vSAN','Storage/Datastore'
        )
    }
    if ($keyIssues) {
        foreach ($k in $keyIssues) { Write-DotLine $sel $k.Status "$($k.Object) - $($k.Item): $($k.Value)" }
    } else {
        Write-Para $sel "No short-term action items identified this run."
    }

    Write-Heading $sel "Conclusion" 2
    $concl = if ($CritCount -gt 0) {
        "The $SiteLabel VMware environment has $CritCount critical item(s) requiring prompt attention, alongside $WarnCount warning item(s). See Global Notes above."
    } elseif ($WarnCount -gt 0) {
        "The $SiteLabel VMware environment is stable with $WarnCount warning item(s) to track. No critical issues were identified during this automated check."
    } else {
        "The $SiteLabel VMware environment is stable and healthy. No critical or warning issues were identified during this automated check."
    }
    Write-Para $sel $concl
    Write-Para $sel "Prepared By: VMware Weekly Health Check Automation" -Bold
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

function Get-WorstStatus {
    param([string[]]$Statuses)
    $order = @{ 'Critical' = 4; 'Warning' = 3; 'Unable to Check' = 2; 'Healthy' = 1; 'Information' = 0 }
    if (-not $Statuses -or $Statuses.Count -eq 0) { return 'Healthy' }
    return ($Statuses | Sort-Object { $order[$_] } -Descending | Select-Object -First 1)
}

# ============================================================================
# 4. OUTPUT: ONE DOCX PER SITE - built natively, no HTML step
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
