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
    [int]$PerfHistoryHours = 24,

    [switch]$SkipExcel   # skip .xlsx generation if ImportExcel module isn't available
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

$ExcelModuleAvailable = (-not $SkipExcel) -and [bool](Get-Module -ListAvailable -Name ImportExcel)
if ($ExcelModuleAvailable) { Import-Module ImportExcel -ErrorAction SilentlyContinue }

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
                    -Item $alarmDef.Info.Name -Value $a.OverallStatus -Status ($a.OverallStatus.Substring(0,1).ToUpper() + $a.OverallStatus.Substring(1))
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
                        -Item $alarmDef.Info.Name -Value $a.OverallStatus -Status ($a.OverallStatus.Substring(0,1).ToUpper() + $a.OverallStatus.Substring(1))
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
                            -Item $alarmDef.Info.Name -Value $a.OverallStatus -Status ($a.OverallStatus.Substring(0,1).ToUpper() + $a.OverallStatus.Substring(1))
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

            # Local ESXi users
            Invoke-SafeCheck -CheckName 'Local ESXi users' -VCenter $VCName -Site $Site -ObjectName $HName -Script {
                $accts = Get-VMHostAccount -Server $VC -VMHost $VMHost -ErrorAction Stop
                New-Finding -Site $Site -VCenter $VCName -Area 'Security & Compliance' -Object $HName `
                    -Item 'Local Accounts' -Value ($accts.Id -join ', ') -Status 'Information' `
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
                New-Finding -Site $Site -VCenter $VCName -Area 'Networking' -Object $vds.Name `
                    -Item 'vDS MTU' -Value $vds.Mtu -Status 'Information'

                $pgs = Get-VDPortgroup -Server $VC -VDSwitch $vds
                foreach ($pg in $pgs) {
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
# 3. STATIC "MANUAL / EXTERNAL REQUIRED" ITEMS
#    These are never retrievable from vCenter/PowerCLI - included so the report is complete
#    and nothing from the original manual checklist is silently dropped.
# ============================================================================
$ManualItems = @(
    'Rack/location information, U-position, device labeling (physical)',
    'Cable labeling and cable management (physical inspection)',
    'Cooling/temperature/humidity readings not exposed via CIM sensors (site environmental monitoring/BMS)',
    'PDU A/B status (requires PDU/iLO-iDRAC/Redfish or environmental monitoring, not vCenter)',
    'iLO/iDRAC connectivity and out-of-band management health (Redfish/iDRAC/iLO API, not vCenter)',
    'Physical condition / cleanliness of racks and equipment (manual visual inspection)',
    'Storage array (HPE Alletra/Cohesity) controller, disk, link and power-supply detail beyond datastore capacity (vendor array management API/UI)',
    'Backup job success/failure detail (Cohesity/backup platform API or console, not vCenter)'
)
foreach ($item in $ManualItems) {
    New-Finding -Site 'All Sites' -VCenter 'n/a' -Area 'External/Manual Checks' -Object 'n/a' `
        -Item $item -Value 'n/a' -Status 'Manual/External Required'
}

# ============================================================================
# 4. OUTPUT: LOG FILE
# ============================================================================
$LogPath = Join-Path $OutputPath "VMware_Weekly_HealthCheck_$RunDate.log"
$LogLines = @()
$LogLines += "VMware Weekly Health Check run - $($ScriptStart.ToString('u'))"
$LogLines += "Sites processed: $(($Connections | ForEach-Object { if ($SiteMap.ContainsKey($_.Name)) {$SiteMap[$_.Name]} else {$_.Name} }) -join ', ')"
$LogLines += "Total findings collected: $($Global:AllResults.Count)"
$LogLines += "Total collection failures: $($Global:FailureLog.Count)"
$LogLines += "----------------------------------------------------------------"
foreach ($f in $Global:FailureLog) {
    $LogLines += "$($f.Timestamp) | $($f.Site) | $($f.VCenter) | $($f.Object) | $($f.Check) | $($f.Error)"
}
$LogLines | Out-File -FilePath $LogPath -Encoding UTF8
Write-Host "`nLog written: $LogPath" -ForegroundColor Cyan

# ============================================================================
# 5. OUTPUT: EXCEL (.xlsx) - one worksheet per functional area
# ============================================================================
$ExcelPath = Join-Path $OutputPath "VMware_Weekly_HealthCheck_$RunDate.xlsx"
if ($ExcelModuleAvailable) {
    if (Test-Path $ExcelPath) { Remove-Item $ExcelPath -Force }
    $areas = $Global:AllResults | Select-Object -ExpandProperty Area -Unique
    foreach ($area in $areas) {
        $sheetName = ($area -replace '[\\/\?\*\[\]:]', ' ').Substring(0, [Math]::Min(31, $area.Length))
        $Global:AllResults | Where-Object { $_.Area -eq $area } |
            Export-Excel -Path $ExcelPath -WorksheetName $sheetName -AutoSize -FreezeTopRow -BoldTopRow -Append
    }
    $Global:FailureLog | Export-Excel -Path $ExcelPath -WorksheetName 'Exceptions' -AutoSize -FreezeTopRow -BoldTopRow -Append
    Write-Host "Excel report written: $ExcelPath" -ForegroundColor Cyan
} else {
    Write-Warning "ImportExcel module not available - skipping .xlsx. Install with: Install-Module ImportExcel -Scope CurrentUser"
    $CsvFallback = Join-Path $OutputPath "VMware_Weekly_HealthCheck_$RunDate.csv"
    $Global:AllResults | Export-Csv -Path $CsvFallback -NoTypeInformation -Encoding UTF8
    Write-Host "CSV fallback written: $CsvFallback" -ForegroundColor Yellow
}

# ============================================================================
# 6. OUTPUT: HTML (management report)
# ============================================================================
function Get-StatusBadge {
    param([string]$Status)
    $color = switch ($Status) {
        'Healthy'  { '#10b981' }
        'Warning'  { '#f59e0b' }
        'Critical' { '#ef4444' }
        'Information' { '#3b82f6' }
        'Unable to Check' { '#9ca3af' }
        'Manual/External Required' { '#a78bfa' }
        default { '#9ca3af' }
    }
    return "<span style='background:$color;color:#fff;padding:2px 8px;border-radius:10px;font-size:12px;'>$Status</span>"
}

function ConvertTo-AreaTable {
    param([string]$Area, [string]$Title)
    $rows = $Global:AllResults | Where-Object { $_.Area -eq $Area }
    if (-not $rows) { return "<p><em>No data collected for $Title.</em></p>" }
    $sb = New-Object System.Text.StringBuilder
    [void]$sb.Append("<table><tr><th>Site</th><th>vCenter</th><th>Object</th><th>Item</th><th>Value</th><th>Status</th><th>Notes</th></tr>")
    foreach ($r in $rows) {
        [void]$sb.Append("<tr><td>$($r.Site)</td><td>$($r.VCenter)</td><td>$($r.Object)</td><td>$($r.Item)</td><td>$($r.Value)</td><td>$(Get-StatusBadge $r.Status)</td><td>$($r.Notes)</td></tr>")
    }
    [void]$sb.Append("</table>")
    return $sb.ToString()
}

$critCount = ($Global:AllResults | Where-Object { $_.Status -eq 'Critical' }).Count
$warnCount = ($Global:AllResults | Where-Object { $_.Status -eq 'Warning' }).Count
$unableCount = ($Global:AllResults | Where-Object { $_.Status -eq 'Unable to Check' }).Count
$manualCount = ($Global:AllResults | Where-Object { $_.Status -eq 'Manual/External Required' }).Count

$SitesSummary = ''
foreach ($s in ($Global:AllResults | Where-Object {$_.Site -ne 'All Sites'} | Select-Object -ExpandProperty Site -Unique)) {
    $siteCrit = ($Global:AllResults | Where-Object { $_.Site -eq $s -and $_.Status -eq 'Critical' }).Count
    $siteWarn = ($Global:AllResults | Where-Object { $_.Site -eq $s -and $_.Status -eq 'Warning' }).Count
    $overall = if ($siteCrit -gt 0) { 'Critical' } elseif ($siteWarn -gt 0) { 'Warning' } else { 'Healthy' }
    $SitesSummary += "<tr><td>$s</td><td>$(Get-StatusBadge $overall)</td><td>$siteCrit</td><td>$siteWarn</td></tr>"
}

$Html = @"
<!DOCTYPE html>
<html>
<head>
<meta charset='utf-8'>
<title>VMware Weekly Health Check - $RunDate</title>
<style>
body { font-family: Segoe UI, Arial, sans-serif; margin: 30px; color: #1f2937; }
h1 { color: #1e3a5f; } h2 { color: #1e3a5f; border-bottom: 2px solid #1e3a5f; padding-bottom: 4px; margin-top: 40px; }
table { border-collapse: collapse; width: 100%; margin: 10px 0 20px 0; font-size: 13px; }
th, td { border: 1px solid #d1d5db; padding: 6px 10px; text-align: left; }
th { background: #1e3a5f; color: #fff; }
tr:nth-child(even) { background: #f3f4f6; }
.summary-box { display:inline-block; padding:14px 22px; margin-right:14px; border-radius:8px; background:#f3f4f6; }
.summary-box .num { font-size:26px; font-weight:bold; }
</style>
</head>
<body>
<h1>VMware Weekly Health Check</h1>
<p>Report Date: $RunDate | Generated: $($ScriptStart.ToString('yyyy-MM-dd HH:mm')) | Sites: $((($Global:AllResults | Where-Object {$_.Site -ne 'All Sites'} | Select-Object -ExpandProperty Site -Unique)) -join ', ')</p>

<h2>1. Executive Summary</h2>
<div class='summary-box'>Critical<div class='num' style='color:#ef4444'>$critCount</div></div>
<div class='summary-box'>Warning<div class='num' style='color:#f59e0b'>$warnCount</div></div>
<div class='summary-box'>Unable to Check<div class='num' style='color:#9ca3af'>$unableCount</div></div>
<div class='summary-box'>Manual/External Required<div class='num' style='color:#a78bfa'>$manualCount</div></div>
<div class='summary-box'>Collection Failures<div class='num'>$($Global:FailureLog.Count)</div></div>

<h2>2-4. Overall Status by Site</h2>
<table><tr><th>Site</th><th>Overall</th><th>Critical Items</th><th>Warning Items</th></tr>$SitesSummary</table>

<h2>5. vCenter Health</h2>
$(ConvertTo-AreaTable -Area 'vCenter Overview' -Title 'vCenter Overview')
$(ConvertTo-AreaTable -Area 'vCenter Appliance' -Title 'vCenter Appliance')

<h2>6. Cluster Health</h2>
$(ConvertTo-AreaTable -Area 'Cluster Configuration' -Title 'Cluster Configuration')

<h2>7. ESXi Host Health</h2>
$(ConvertTo-AreaTable -Area 'ESXi Host Health' -Title 'ESXi Host Health')

<h2>8. Hardware Health</h2>
$(ConvertTo-AreaTable -Area 'Hardware Health' -Title 'Hardware Health')

<h2>9. Networking</h2>
$(ConvertTo-AreaTable -Area 'Networking' -Title 'Networking')

<h2>10. CPU/Memory Capacity</h2>
$(ConvertTo-AreaTable -Area 'CPU/Memory Capacity' -Title 'CPU/Memory Capacity')

<h2>11. Storage / Datastore / vSAN</h2>
$(ConvertTo-AreaTable -Area 'Storage/Datastore' -Title 'Storage/Datastore')
$(ConvertTo-AreaTable -Area 'Storage/vSAN' -Title 'Storage/vSAN')

<h2>12. VM Inventory and OS Distribution</h2>
$(ConvertTo-AreaTable -Area 'VM Inventory' -Title 'VM Inventory')

<h2>13. Security &amp; Compliance</h2>
$(ConvertTo-AreaTable -Area 'Security & Compliance' -Title 'Security & Compliance')

<h2>14. Active Warnings / Alarms</h2>
$(ConvertTo-AreaTable -Area 'Alarms' -Title 'Alarms')

<h2>15. External / Manual Checks Still Required</h2>
$(ConvertTo-AreaTable -Area 'External/Manual Checks' -Title 'External/Manual Checks')

<h2>16. Exceptions / Failed Data Collection</h2>
<table><tr><th>Timestamp</th><th>Site</th><th>vCenter</th><th>Object</th><th>Check</th><th>Error</th></tr>
$(($Global:FailureLog | ForEach-Object { "<tr><td>$($_.Timestamp)</td><td>$($_.Site)</td><td>$($_.VCenter)</td><td>$($_.Object)</td><td>$($_.Check)</td><td>$($_.Error)</td></tr>" }) -join "`n")
</table>

<h2>17. Action Items</h2>
<table><tr><th>Site</th><th>Object</th><th>Item</th><th>Value</th><th>Status</th></tr>
$(($Global:AllResults | Where-Object { $_.Status -in 'Critical','Warning' } | ForEach-Object { "<tr><td>$($_.Site)</td><td>$($_.Object)</td><td>$($_.Item)</td><td>$($_.Value)</td><td>$(Get-StatusBadge $_.Status)</td></tr>" }) -join "`n")
</table>

</body>
</html>
"@

$HtmlPath = Join-Path $OutputPath "VMware_Weekly_HealthCheck_$RunDate.html"
$Html | Out-File -FilePath $HtmlPath -Encoding UTF8
Write-Host "HTML report written: $HtmlPath" -ForegroundColor Cyan

Write-Host "`nDone. Findings: $($Global:AllResults.Count) | Failures: $($Global:FailureLog.Count) | Duration: $([Math]::Round(((Get-Date)-$ScriptStart).TotalMinutes,1)) min" -ForegroundColor Green
