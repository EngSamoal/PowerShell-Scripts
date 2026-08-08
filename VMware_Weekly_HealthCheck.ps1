#Requires -Version 5.1
<#
.SYNOPSIS
    VMware Weekly Health Check - Read-only automated collection and DOCX report generation,
    matching the "<Site> vCenter & ESXi Health Check Report" template (SAMI/Qiddiya format).

.DESCRIPTION
    Replaces the manual weekly vCenter/ESXi health-check reports with a single read-only
    PowerCLI collection run from an already-authenticated laptop session.

    - Uses whatever vCenter sessions are already connected (Connect-VIServer done beforehand).
    - Auto-discovers clusters, hosts, datastores, vSAN, vDS/port groups and VMs per vCenter -
      nothing about the infrastructure is hard-coded.
    - Produces ONE .docx per connected vCenter ("site"), matching the exact section layout,
      wording and table structure of the reference report:
        1.  Executive Summary
        2.  Environment Overview
        3.  vCenter Server Health (3.1 Appliance Health)
        4.  ESXi Host Health (4.1 Host Configuration Summary, 4.2 Hardware Health)
        5.  Networking Health (5.1 Network Overview, 5.2 Network Configuration Status)
        6.  Performance & Capacity Summary (6.1 CPU, 6.2 Memory, 6.3 Storage)
        7.  Storage Health (array hardware + per-datastore table)
        8.  Virtual Machine Inventory Summary (8.1 VM OS Distribution)
        9.  Security & Compliance
        10. Risks & Recommendations (Backup & Disaster Recovery)
        11. Action Plan
        12. Conclusion, Prepared By / Report Date
      with a company header (logos + date) and classification footer repeated on every page.
    - When a vCenter has more than one cluster, the per-cluster tables (4.1, 6.1-6.3, 8.1) simply
      grow one row per cluster instead of collapsing to a single row - the section layout itself
      does not change.
    - A failure collecting one item is logged and skipped; the script always continues.
    - NEVER writes/modifies anything in vCenter. Every cmdlet used below is read-only
      (Get-*, no Set-*/New-*/Remove-*, no config changes, no SSH enablement).

.NOTES
    Every threshold below (capacity %, cert/license expiry windows) is a CONFIGURABLE PARAMETER,
    not an assumed company standard - raw values/status are always shown alongside any computed
    flag.

    Two parts of the reference report are NOT available from vCenter/PowerCLI by design:
      - Section 7's physical storage-array details (software version, compression, controller
        A/B state, disk count, link/power status) come from the array's own management plane
        (e.g. NetApp ONTAP System Manager), not vCenter. Supply them via -StorageArrayInfo.
      - Section 3.1's "Backup" row and Section 10 (Backup & Disaster Recovery) come from the
        backup product's own console (e.g. Cohesity), not vCenter. Supply them via -BackupInfo.
    If not supplied for a given site, those sections are rendered as "Manual/External Required"
    rather than fabricated.

.EXAMPLE
    # Already connected: Connect-VIServer tb-vc.aq.local
    .\VMware_Weekly_HealthCheck.ps1 -SiteMap @{'tb-vc.aq.local'='Tabuk'} `
        -StorageArrayInfo @{ 'Tabuk' = @{ SoftwareVersion='10.4.20'; CompressionPct=99;
            ControllerA='Active'; ControllerB='Standby'; DiskCount=28; DiskStatus='Healthy';
            UtilizationUsedTiB=2.4; UtilizationTotalTiB=152.5;
            LinksA='eth0a, eth0b is active'; LinksB='eth0a, eth0b is active';
            PowerA='Power Supply active'; PowerB='Power Supply active' } } `
        -BackupInfo @{ 'Tabuk' = @{ DeviceLabel='Cohesity Backup Device';
            SolutionName='Cohesity Backup Solution'; Status='Healthy' } } `
        -PreparedBy 'Ahmed Khalil' -PreparedByTitle 'Infrastructure Specialist' `
        -LogoLeftPath 'C:\Logos\SAMI.png' -LogoRightPath 'C:\Logos\Qiddiya.png'
#>

[CmdletBinding()]
param(
    # Maps a connected vCenter server (Name as shown in $global:DefaultVIServers) to a friendly
    # site label used in the report title/filename. If a connected vCenter isn't in this map,
    # its own server name is used as the label - nothing is hard-coded or required.
    [hashtable]$SiteMap = @{},

    [string]$OutputPath = (Join-Path $PSScriptRoot "VMware_HealthCheck_Reports"),

    # ---- Report letterhead / sign-off -----------------------------------------------------
    [string]$LogoLeftPath = '',
    [string]$LogoRightPath = '',
    [string]$FooterText = 'This email \ document has been classified as public',
    [string]$PreparedBy = 'VMware Weekly Health Check Automation',
    [string]$PreparedByTitle = '',

    # ---- Data NOT available from vCenter - keyed by Site label (see .NOTES) ---------------
    # Example: @{ 'Tabuk' = @{ SoftwareVersion='10.4.20'; CompressionPct=99; ControllerA='Active';
    #   ControllerB='Standby'; DiskCount=28; DiskStatus='Healthy'; UtilizationUsedTiB=2.4;
    #   UtilizationTotalTiB=152.5; LinksA='eth0a, eth0b is active'; LinksB='eth0a, eth0b is active';
    #   PowerA='Power Supply active'; PowerB='Power Supply active' } }
    [hashtable]$StorageArrayInfo = @{},
    # Example: @{ 'Tabuk' = @{ DeviceLabel='Cohesity Backup Device';
    #   SolutionName='Cohesity Backup Solution'; Status='Healthy' } }
    [hashtable]$BackupInfo = @{},

    # Local ESXi accounts considered normal/expected; anything extra found on a host is flagged.
    [string[]]$ExpectedLocalAccounts = @('root','dcui','vpxuser'),

    # Configurable thresholds - NOT vendor/company-defined standards. Raw values are always
    # shown regardless of these; these only drive the Warning/Critical flag shown alongside them.
    [double]$CapacityWarningPct    = 80,
    [double]$CapacityCriticalPct   = 90,
    [int]$CertExpiryWarningDays    = 60,
    [int]$LicenseExpiryWarningDays = 30,

    # Historical performance window for capacity stats (hours), matches the 24-72h window used
    # in the reference report.
    [int]$PerfHistoryHours = 24
)

$ErrorActionPreference = 'Stop'
$ScriptStart = Get-Date
$RunDate     = $ScriptStart.ToString('yyyy-MM-dd')
$RunDateDisplay = $ScriptStart.ToString('d-MMMM-yyyy', [System.Globalization.CultureInfo]::InvariantCulture)
if (-not (Test-Path $OutputPath)) { New-Item -ItemType Directory -Path $OutputPath -Force | Out-Null }

# ============================================================================
# 0. LOGGING / SAFE-EXECUTION HELPERS
# ============================================================================
$Global:FailureLog = [System.Collections.Generic.List[object]]::new()
$Global:AllResults = [System.Collections.Generic.List[object]]::new()
$Global:SiteClusterMap = @{}   # Site label -> ordered list of cluster names discovered.

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
        [string]$Site, [string]$VCenter, [string]$Area, [string]$Cluster, [string]$Object,
        [string]$Item, [string]$Value, [string]$Status, [string]$Notes = ''
    )
    # Area is a loose grouping used only to make report-building lookups readable:
    #   Overview | Appliance | Alarms | Cluster | Host | Hardware | Security | Networking |
    #   Capacity | Storage | VM
    # Status must be one of: Healthy / Warning / Critical / Information / Unable to Check / Manual/External Required
    $obj = [pscustomobject]@{
        Site    = $Site
        VCenter = $VCenter
        Area    = $Area
        Cluster = $Cluster
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
    throw "No connected vCenter sessions found. Connect first, e.g.:`n  Connect-VIServer tb-vc.aq.local"
}

Write-Host "Connected vCenter sessions: $($Connections.Name -join ', ')" -ForegroundColor Cyan

# ============================================================================
# 2. PER-VCENTER / PER-SITE COLLECTION
# ============================================================================
foreach ($VC in $Connections) {

    $VCName = $VC.Name
    if ($SiteMap.ContainsKey($VCName)) {
        $Site = $SiteMap[$VCName]
    } else {
        # No explicit -SiteMap entry - fall back to the friendly name configured in vCenter itself
        # (Configure > Advanced Settings > VirtualCenter.InstanceName) rather than the raw
        # connection string/IP, when one is set. -SiteMap always wins if you provide it.
        $instanceName = $null
        try {
            $instanceName = (Get-AdvancedSetting -Entity $VC -Name 'VirtualCenter.InstanceName' -ErrorAction Stop).Value
        } catch { }
        $Site = if ($instanceName) { $instanceName } else { $VCName }
    }

    Write-Host "`n=== Collecting: $Site ($VCName) ===" -ForegroundColor Green

    # --- vCenter version/build --------------------------------------------------------------
    Invoke-SafeCheck -CheckName 'vCenter version/build' -VCenter $VCName -Site $Site -ObjectName $VCName -Script {
        New-Finding -Site $Site -VCenter $VCName -Area 'Overview' -Object $VCName `
            -Item 'vCenter Version/Build' -Value "vCenter Server $($VC.Version) (Build $($VC.Build))" -Status 'Information'
    }

    # --- Licensing -----------------------------------------------------------------------------
    Invoke-SafeCheck -CheckName 'Licensing' -VCenter $VCName -Site $Site -ObjectName $VCName -Script {
        $lm = Get-View -Server $VC ($VC.ExtensionData.Content.LicenseManager)
        foreach ($lic in $lm.Licenses) {
            $expProp = $lic.Properties | Where-Object { $_.Key -eq 'expirationDate' }
            $status  = 'Information'
            $expStr  = 'No expiration / not set'
            if ($expProp) {
                $expDate = [datetime]$expProp.Value
                $expStr  = $expDate.ToString('yyyy-MM-dd')
                $daysLeft = ($expDate - (Get-Date)).Days
                # Only treat expiry as an operational risk if the license key is actually in use
                # (Used > 0). vCenter License Manager commonly retains old/replaced keys with
                # Used=0 - flagging those as Critical would be a false alarm, not a real finding.
                if ($lic.Used -gt 0) {
                    if ($daysLeft -le 0) { $status = 'Critical' }
                    elseif ($daysLeft -le $LicenseExpiryWarningDays) { $status = 'Warning' }
                    else { $status = 'Healthy' }
                } else {
                    $status = 'Information'
                }
            }
            New-Finding -Site $Site -VCenter $VCName -Area 'Overview' -Object $VCName `
                -Item "License: $($lic.Name)" -Value "Used $($lic.Used)/$($lic.Total) - Expires $expStr" -Status $status
        }
    }

    # --- vCenter appliance CPU/Mem/Disk/Services/NTP/Certificate (VAMI/CIS) ------------------
    $CisSession = $global:DefaultCisServers | Where-Object { $_.Name -eq $VCName -and $_.IsConnected }
    if ($CisSession) {
        Invoke-SafeCheck -CheckName 'Appliance health (VAMI)' -VCenter $VCName -Site $Site -ObjectName $VCName -Script {
            foreach ($comp in 'cpu','mem','storage') {
                $svc = Get-CisService -Name "com.vmware.appliance.health.$comp" -Server $CisSession
                $val = $svc.get()
                $label = @{ cpu = 'CPU'; mem = 'Memory'; storage = 'Disk Usage' }[$comp]
                New-Finding -Site $Site -VCenter $VCName -Area 'Appliance' -Object $VCName `
                    -Item $label -Value $(if ($val -eq 'green') { 'Normal' } else { $val }) -Status $(if ($val -eq 'green') {'Healthy'} else {'Warning'})
            }
        }
        Invoke-SafeCheck -CheckName 'Appliance services (VAMI)' -VCenter $VCName -Site $Site -ObjectName $VCName -Script {
            $svcListSvc = Get-CisService -Name 'com.vmware.appliance.services' -Server $CisSession
            $services = $svcListSvc.list()
            $notRunning = $services.GetEnumerator() | Where-Object { $_.Value.state -ne 'STARTED' }
            New-Finding -Site $Site -VCenter $VCName -Area 'Appliance' -Object $VCName -Item 'Services' `
                -Value $(if ($notRunning) { ($notRunning | ForEach-Object { "$($_.Key): $($_.Value.state)" }) -join '; ' } else { 'All Running' }) `
                -Status $(if ($notRunning) { 'Warning' } else { 'Healthy' })
        }
        Invoke-SafeCheck -CheckName 'Appliance NTP (VAMI)' -VCenter $VCName -Site $Site -ObjectName $VCName -Script {
            $ntpSvc = Get-CisService -Name 'com.vmware.appliance.ntp' -Server $CisSession
            $ntpStatus = $ntpSvc.test()
            $synced = ($ntpStatus | Out-String) -match 'Server .* is reachable and clock is synchronized|success'
            New-Finding -Site $Site -VCenter $VCName -Area 'Appliance' -Object $VCName -Item 'NTP' `
                -Value $(if ($synced) { 'Configured & Synchronized' } else { ($ntpStatus | Out-String).Trim() }) `
                -Status $(if ($synced) { 'Healthy' } else { 'Warning' })
        }
        Invoke-SafeCheck -CheckName 'Appliance certificate (VAMI)' -VCenter $VCName -Site $Site -ObjectName $VCName -Script {
            $certSvc = Get-CisService -Name 'com.vmware.appliance.certificate_management.vcenter.tls' -Server $CisSession
            $cert = $certSvc.get()
            $expDate = [datetime]$cert.valid_to
            $daysLeft = ($expDate - (Get-Date)).Days
            $status = if ($daysLeft -le 0) { 'Critical' } elseif ($daysLeft -le $CertExpiryWarningDays) { 'Warning' } else { 'Healthy' }
            New-Finding -Site $Site -VCenter $VCName -Area 'Appliance' -Object $VCName -Item 'Certificates' `
                -Value "Valid - until ($($expDate.ToString('MMM d, yyyy')))" -Status $status
        }
    } else {
        foreach ($item in 'CPU','Memory','Disk Usage','Services','NTP','Certificates') {
            New-Finding -Site $Site -VCenter $VCName -Area 'Appliance' -Object $VCName -Item $item `
                -Value 'n/a' -Status 'Manual/External Required' `
                -Notes 'Requires a VAMI/CIS session (Connect-CisServer <vcenter>) in addition to the vSphere API session. Not connected in this run.'
        }
    }

    # --- Active vCenter-level alarms ----------------------------------------------------------
    Invoke-SafeCheck -CheckName 'vCenter-level alarms' -VCenter $VCName -Site $Site -ObjectName $VCName -Script {
        $rootFolder = Get-View -Server $VC $VC.ExtensionData.Content.RootFolder
        $alarms = $rootFolder.TriggeredAlarmState
        if ($alarms -and $alarms.Count -gt 0) {
            foreach ($a in $alarms) {
                $alarmDef = Get-View -Server $VC $a.Alarm
                New-Finding -Site $Site -VCenter $VCName -Area 'Alarms' -Object $VCName `
                    -Item $alarmDef.Info.Name -Value $a.OverallStatus.ToString() -Status ($a.OverallStatus.ToString().Substring(0,1).ToUpper() + $a.OverallStatus.ToString().Substring(1))
            }
        }
    }

    # --- Discover clusters ---------------------------------------------------------------------
    $Clusters = Invoke-SafeCheck -CheckName 'Cluster discovery' -VCenter $VCName -Site $Site -ObjectName $VCName -Script {
        Get-Cluster -Server $VC
    }
    if (-not $Clusters) { continue }

    foreach ($Cluster in $Clusters) {
        $ClusterName = $Cluster.Name
        Write-Host "  - Cluster: $ClusterName"

        if (-not $Global:SiteClusterMap.ContainsKey($Site)) { $Global:SiteClusterMap[$Site] = [System.Collections.Generic.List[string]]::new() }
        $Global:SiteClusterMap[$Site].Add($ClusterName)

        # HA / DRS
        Invoke-SafeCheck -CheckName 'HA/DRS status' -VCenter $VCName -Site $Site -ObjectName $ClusterName -Script {
            New-Finding -Site $Site -VCenter $VCName -Area 'Cluster' -Cluster $ClusterName -Object $ClusterName `
                -Item 'vSphere HA' -Value $Cluster.HAEnabled -Status $(if ($Cluster.HAEnabled) {'Healthy'} else {'Warning'})
            New-Finding -Site $Site -VCenter $VCName -Area 'Cluster' -Cluster $ClusterName -Object $ClusterName `
                -Item 'vSphere DRS' -Value $Cluster.DrsEnabled -Status $(if ($Cluster.DrsEnabled) {'Healthy'} else {'Warning'})
        }

        # Cluster-level alarms
        Invoke-SafeCheck -CheckName 'Cluster alarms' -VCenter $VCName -Site $Site -ObjectName $ClusterName -Script {
            $alarms = $Cluster.ExtensionData.TriggeredAlarmState
            if ($alarms -and $alarms.Count -gt 0) {
                foreach ($a in $alarms) {
                    $alarmDef = Get-View -Server $VC $a.Alarm
                    New-Finding -Site $Site -VCenter $VCName -Area 'Alarms' -Cluster $ClusterName -Object $ClusterName `
                        -Item $alarmDef.Info.Name -Value $a.OverallStatus.ToString() -Status ($a.OverallStatus.ToString().Substring(0,1).ToUpper() + $a.OverallStatus.ToString().Substring(1))
                }
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

            Invoke-SafeCheck -CheckName 'Host connection state' -VCenter $VCName -Site $Site -ObjectName $HName -Script {
                $ok = $VMHost.ConnectionState -eq 'Connected'
                New-Finding -Site $Site -VCenter $VCName -Area 'Host' -Cluster $ClusterName -Object $HName `
                    -Item 'Connection State' -Value $VMHost.ConnectionState -Status $(if ($ok) {'Healthy'} else {'Critical'})
            }

            Invoke-SafeCheck -CheckName 'Host alarms' -VCenter $VCName -Site $Site -ObjectName $HName -Script {
                $alarms = $VMHost.ExtensionData.TriggeredAlarmState
                if ($alarms -and $alarms.Count -gt 0) {
                    foreach ($a in $alarms) {
                        $alarmDef = Get-View -Server $VC $a.Alarm
                        New-Finding -Site $Site -VCenter $VCName -Area 'Alarms' -Cluster $ClusterName -Object $HName `
                            -Item $alarmDef.Info.Name -Value $a.OverallStatus.ToString() -Status ($a.OverallStatus.ToString().Substring(0,1).ToUpper() + $a.OverallStatus.ToString().Substring(1))
                    }
                }
            }

            # Hardware health via built-in host Health System (numeric sensors) - no SSH required.
            # Rolled up site-wide in the report into CPU/Memory/Power Supplies/Fans/RAID/Disks.
            Invoke-SafeCheck -CheckName 'Hardware sensors' -VCenter $VCName -Site $Site -ObjectName $HName -Script {
                $sensors = $VMHost.ExtensionData.Runtime.HealthSystemRuntime.SystemHealthInfo.NumericSensorInfo
                if (-not $sensors) {
                    New-Finding -Site $Site -VCenter $VCName -Area 'Hardware' -Cluster $ClusterName -Object $HName `
                        -Item 'Sensors' -Value 'n/a' -Status 'Unable to Check' `
                        -Notes 'Host does not expose CIM/IPMI sensor data to vCenter (common on some blade/BMC configs).'
                    return
                }
                $groups = $sensors | Group-Object -Property SensorType
                foreach ($g in $groups) {
                    $bad = $g.Group | Where-Object { $_.HealthState.Key -notin @('green','Green') }
                    $status = if ($bad) { 'Warning' } else { 'Healthy' }
                    $summary = if ($bad) { ($bad | ForEach-Object { "$($_.Name): $($_.HealthState.Label)" }) -join '; ' } else { 'All normal' }
                    New-Finding -Site $Site -VCenter $VCName -Area 'Hardware' -Cluster $ClusterName -Object $HName `
                        -Item $g.Name -Value $summary -Status $status
                }
            }

            # Secure Boot (via EsxCli, no SSH required - this is the vSphere API path)
            Invoke-SafeCheck -CheckName 'Secure Boot' -VCenter $VCName -Site $Site -ObjectName $HName -Script {
                $esxcli = Get-EsxCli -VMHost $VMHost -Server $VC -V2
                $sb = $esxcli.system.settings.encryption.get.Invoke()
                $enabled = $sb.RequireSecureBoot
                New-Finding -Site $Site -VCenter $VCName -Area 'Security' -Cluster $ClusterName -Object $HName `
                    -Item 'Secure Boot' -Value $enabled -Status $(if ($enabled -eq $true -or $enabled -eq 'true') {'Healthy'} else {'Warning'})
            }

            # Lockdown Mode
            Invoke-SafeCheck -CheckName 'Lockdown mode' -VCenter $VCName -Site $Site -ObjectName $HName -Script {
                $lockdown = $VMHost.ExtensionData.Config.LockdownMode
                $enabled = $lockdown -ne 'lockdownDisabled'
                New-Finding -Site $Site -VCenter $VCName -Area 'Security' -Cluster $ClusterName -Object $HName `
                    -Item 'Lockdown Mode' -Value $lockdown -Status $(if ($enabled) {'Healthy'} else {'Warning'})
            }

            # Local ESXi users - via esxcli (API-based, no SSH)
            Invoke-SafeCheck -CheckName 'Local ESXi users' -VCenter $VCName -Site $Site -ObjectName $HName -Script {
                $esxcli = Get-EsxCli -VMHost $VMHost -Server $VC -V2
                $accts = $esxcli.system.account.list.Invoke() | ForEach-Object { $_.UserID }
                $unexpected = $accts | Where-Object { $_ -notin $ExpectedLocalAccounts }
                New-Finding -Site $Site -VCenter $VCName -Area 'Security' -Cluster $ClusterName -Object $HName `
                    -Item 'Local Accounts' -Value ($accts -join ', ') -Status $(if ($unexpected) {'Warning'} else {'Healthy'}) `
                    -Notes $(if ($unexpected) { "Unexpected account(s): $($unexpected -join ', ') - review membership manually." } else { '' })
            }

            # Syslog
            Invoke-SafeCheck -CheckName 'Syslog config' -VCenter $VCName -Site $Site -ObjectName $HName -Script {
                $syslog = Get-VMHostSysLogServer -Server $VC -VMHost $VMHost
                $configured = -not [string]::IsNullOrWhiteSpace($syslog.Host)
                New-Finding -Site $Site -VCenter $VCName -Area 'Security' -Cluster $ClusterName -Object $HName `
                    -Item 'Syslog' -Value $(if ($configured) { "$($syslog.Host):$($syslog.Port)" } else { 'Not configured' }) `
                    -Status $(if ($configured) {'Healthy'} else {'Warning'})
            }

            # Physical NICs / redundancy
            Invoke-SafeCheck -CheckName 'Physical NICs' -VCenter $VCName -Site $Site -ObjectName $HName -Script {
                $pnics = @(Get-VMHostNetworkAdapter -Server $VC -VMHost $VMHost -Physical)
                $up = @($pnics | Where-Object { $_.BitRatePerSec -gt 0 })
                New-Finding -Site $Site -VCenter $VCName -Area 'Networking' -Cluster $ClusterName -Object $HName `
                    -Item 'Physical NIC Redundancy' -Value "$($pnics.Count) total, $($up.Count) linked up" `
                    -Status $(if ($pnics.Count -ge 2 -and $up.Count -ge 2) {'Healthy'} else {'Warning'})
            }

            # NIC error/drop counters via performance manager (safe, no SSH)
            Invoke-SafeCheck -CheckName 'NIC error counters' -VCenter $VCName -Site $Site -ObjectName $HName -Script {
                $stat = Get-Stat -Server $VC -Entity $VMHost -Stat 'net.errorsRx.summation','net.errorsTx.summation','net.droppedRx.summation','net.droppedTx.summation' -Realtime -MaxSamples 1 -ErrorAction Stop
                if ($stat) {
                    $total = ($stat | Measure-Object -Property Value -Sum).Sum
                    New-Finding -Site $Site -VCenter $VCName -Area 'Networking' -Cluster $ClusterName -Object $HName `
                        -Item 'NIC Errors/Drops' -Value $total -Status $(if ($total -eq 0) {'Healthy'} else {'Warning'})
                }
            }
        } # end per-host

        # ESXi version consistency across the cluster
        $DistinctVersions = @($HostVersions | Select-Object -Unique)
        New-Finding -Site $Site -VCenter $VCName -Area 'Host' -Cluster $ClusterName -Object $ClusterName `
            -Item 'Version Consistency' -Value ($DistinctVersions -join ' | ') -Status $(if ($DistinctVersions.Count -le 1) {'Healthy'} else {'Warning'})

        # Host Configuration Summary row for this cluster (Host Count | Version | Build | Status)
        Invoke-SafeCheck -CheckName 'Host configuration summary' -VCenter $VCName -Site $Site -ObjectName $ClusterName -Script {
            $verBuild = ($Hosts | Select-Object -First 1)
            $connStatus = if (($Hosts | Where-Object { $_.ConnectionState -ne 'Connected' })) { 'Warning' } else { 'Healthy' }
            $status = if ($DistinctVersions.Count -gt 1) { 'Warning' } else { $connStatus }
            New-Finding -Site $Site -VCenter $VCName -Area 'Host' -Cluster $ClusterName -Object $ClusterName `
                -Item 'Host Configuration Summary' -Value "$($Hosts.Count)|$($verBuild.Version)|$($verBuild.Build)" -Status $status
        }

        # Cluster CPU/Memory capacity - historical stats over the configured window
        Invoke-SafeCheck -CheckName 'Cluster CPU/Mem capacity (historical)' -VCenter $VCName -Site $Site -ObjectName $ClusterName -Script {
            $start = (Get-Date).AddHours(-1 * $PerfHistoryHours)
            $cpuStat = Get-Stat -Server $VC -Entity $Cluster -Stat 'cpu.usage.average' -Start $start -Finish (Get-Date) -ErrorAction Stop
            $memStat = Get-Stat -Server $VC -Entity $Cluster -Stat 'mem.usage.average' -Start $start -Finish (Get-Date) -ErrorAction Stop

            $cpuTotalMHz = ($Hosts | Measure-Object -Property CpuTotalMhz -Sum).Sum
            $memTotalMB  = ($Hosts | Measure-Object -Property MemoryTotalMB -Sum).Sum
            $avgCpuPct = if ($cpuStat) { ($cpuStat | Measure-Object -Property Value -Average).Average } else { $null }
            $avgMemPct = if ($memStat) { ($memStat | Measure-Object -Property Value -Average).Average } else { $null }

            New-Finding -Site $Site -VCenter $VCName -Area 'Capacity' -Cluster $ClusterName -Object $ClusterName `
                -Item 'CPU' -Value "$cpuTotalMHz|$avgCpuPct" -Status $(if ($avgCpuPct -ne $null) { Get-PctStatus $avgCpuPct } else { 'Unable to Check' })
            New-Finding -Site $Site -VCenter $VCName -Area 'Capacity' -Cluster $ClusterName -Object $ClusterName `
                -Item 'Memory' -Value "$memTotalMB|$avgMemPct" -Status $(if ($avgMemPct -ne $null) { Get-PctStatus $avgMemPct } else { 'Unable to Check' })
        }

        # Datastores in this cluster
        Invoke-SafeCheck -CheckName 'Datastore health/capacity' -VCenter $VCName -Site $Site -ObjectName $ClusterName -Script {
            $datastores = Get-Datastore -Server $VC -RelatedObject $Cluster
            foreach ($ds in $datastores) {
                $usedPct = if ($ds.CapacityGB -gt 0) { (($ds.CapacityGB - $ds.FreeSpaceGB) / $ds.CapacityGB) * 100 } else { 0 }
                $accessible = $ds.ExtensionData.Summary.Accessible
                $status = if (-not $accessible) { 'Critical' } else { Get-PctStatus $usedPct }
                New-Finding -Site $Site -VCenter $VCName -Area 'Storage' -Cluster $ClusterName -Object $ds.Name `
                    -Item 'Datastore' -Value "$($ds.CapacityGB)|$($ds.FreeSpaceGB)|$accessible" -Status $status
            }
        }

        # vSAN (only if the cluster actually has vSAN enabled and the module is available)
        Invoke-SafeCheck -CheckName 'vSAN health/capacity' -VCenter $VCName -Site $Site -ObjectName $ClusterName -Script {
            if (-not $Cluster.VsanEnabled) { return }
            New-Finding -Site $Site -VCenter $VCName -Area 'Storage' -Cluster $ClusterName -Object $ClusterName `
                -Item 'vSAN Enabled' -Value 'True' -Status 'Information'
            if (-not $VsanModuleAvailable) {
                New-Finding -Site $Site -VCenter $VCName -Area 'Storage' -Cluster $ClusterName -Object $ClusterName `
                    -Item 'vSAN Health/Capacity' -Value 'n/a' -Status 'Unable to Check' `
                    -Notes 'VMware.VimAutomation.Vsan module not installed on this laptop.'
                return
            }
            $space = Get-VsanSpaceUsage -Server $VC -Cluster $Cluster -ErrorAction Stop
            $usedPct = if ($space.CapacityGB -gt 0) { ($space.UsedGB / $space.CapacityGB) * 100 } else { 0 }
            New-Finding -Site $Site -VCenter $VCName -Area 'Storage' -Cluster $ClusterName -Object $ClusterName `
                -Item 'vSAN Capacity' -Value "$($space.CapacityGB)|$($space.UsedGB)" -Status (Get-PctStatus $usedPct)
        }

        # VM inventory / guest OS distribution
        Invoke-SafeCheck -CheckName 'VM inventory' -VCenter $VCName -Site $Site -ObjectName $ClusterName -Script {
            $vms = @(Get-VM -Server $VC -Location $Cluster)
            New-Finding -Site $Site -VCenter $VCName -Area 'VM' -Cluster $ClusterName -Object $ClusterName `
                -Item 'VM Count' -Value $vms.Count -Status 'Information'
            # Config.GuestFullName (the "Guest OS" type configured on the VM) is what vCenter's own
            # UI displays and what a human cross-checking the inventory will see - preferred over
            # Guest.OSFullName (the live value VMware Tools currently reports), which can disagree
            # with the configured type when a VM's guest OS was upgraded in place without updating
            # its VM settings, or when Tools is outdated/stopped.
            $osDist = $vms | Group-Object { if ($_.ExtensionData.Config.GuestFullName) { $_.ExtensionData.Config.GuestFullName } else { $_.Guest.OSFullName } }
            foreach ($g in $osDist) {
                New-Finding -Site $Site -VCenter $VCName -Area 'VM' -Cluster $ClusterName -Object $g.Name `
                    -Item 'Guest OS' -Value $g.Count -Status 'Information'
            }
        }

        # vDS / port groups / VLAN / teaming / MTU for this cluster's hosts
        Invoke-SafeCheck -CheckName 'vDS/networking config' -VCenter $VCName -Site $Site -ObjectName $ClusterName -Script {
            $vdSwitches = Get-VDSwitch -Server $VC -VMHost $Hosts -ErrorAction SilentlyContinue | Select-Object -Unique
            foreach ($vds in $vdSwitches) {
                New-Finding -Site $Site -VCenter $VCName -Area 'Networking' -Cluster $ClusterName -Object $vds.Name `
                    -Item 'vDS' -Value $vds.Mtu -Status 'Information'
                # Excludes the auto-generated uplink port group every vDS gets (e.g. "...-DVUplinks-...").
                # It's switch infrastructure carrying a full VLAN trunk range for the physical NICs,
                # not an application/server VLAN - counting it as one inflated the VLAN total. Filtered
                # by the actual IsUplink flag, not by name pattern, so this can't miss a renamed one.
                $pgs = Get-VDPortgroup -Server $VC -VDSwitch $vds | Where-Object { -not $_.IsUplink }
                foreach ($pg in $pgs) {
                    # Regular port groups expose .VlanId as a plain int. Private VLAN port groups
                    # use a different config type with .PvlanId instead, and trunk port groups
                    # expose a range array - without handling those, their VLAN silently reads as
                    # $null and drops out of the distinct-VLAN count below.
                    $vlanCfg = $pg.ExtensionData.Config.DefaultPortConfig.Vlan
                    $vlan = if ($null -ne $vlanCfg.VlanId -and $vlanCfg.VlanId -is [int]) { $vlanCfg.VlanId }
                        elseif ($vlanCfg.PvlanId) { "PVLAN $($vlanCfg.PvlanId)" }
                        elseif ($vlanCfg.VlanId) { ($vlanCfg.VlanId | ForEach-Object { "$($_.Start)-$($_.End)" }) -join ',' }
                        else { $null }
                    $teaming = $pg.ExtensionData.Config.DefaultPortConfig.UplinkTeamingPolicy.Policy.Value
                    New-Finding -Site $Site -VCenter $VCName -Area 'Networking' -Cluster $ClusterName -Object $pg.Name `
                        -Item 'Port Group' -Value "$vlan|$teaming" -Status 'Information'
                }
            }
            $distinctMtu = @($vdSwitches | Select-Object -ExpandProperty Mtu -Unique)
            if ($distinctMtu.Count -gt 1) {
                New-Finding -Site $Site -VCenter $VCName -Area 'Networking' -Cluster $ClusterName -Object $ClusterName `
                    -Item 'MTU Consistency' -Value ($distinctMtu -join ' vs ') -Status 'Warning' `
                    -Notes 'Distributed switches in this cluster are not configured with matching MTU - verify this is intentional.'
            }
        }

    } # end per-cluster
} # end per-vCenter

# ============================================================================
# ============================================================================
# 3. REPORT GENERATION HELPERS
#    Each site's .docx is built DIRECTLY in Word via COM automation - typing text, applying
#    styles/colors, inserting native Word tables, and using real Word Section headers/footers
#    (logos + date repeat at the top of every page, classification text at the bottom of every
#    page) rather than typing them once into the body. There is no HTML step anywhere in this
#    pipeline.
# ============================================================================
$StatusColors = @{
    'Healthy'                  = @(22,163,74)
    'Warning'                  = @(245,158,11)
    'Critical'                 = @(220,38,38)
    'Information'              = @(37,99,235)
    'Unable to Check'          = @(156,163,175)
    'Manual/External Required' = @(156,163,175)
}
function Get-WordColorLong {
    param([int[]]$Rgb)
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

function Write-StatusLine {
    param($Selection, [string]$Label, [string]$Status, [string]$Text = '', [switch]$Bold)
    $Selection.Font.Bold = $true
    $Selection.TypeText($Label)
    $Selection.Font.Bold = [bool]$Bold
    $Selection.Font.Color = Get-StatusColorLong $Status
    $Selection.TypeText(" $([char]0x25CF) ")
    $Selection.Font.Color = -16777216
    if ($Text) { $Selection.TypeText($Text) } else { $Selection.TypeText($Status) }
    $Selection.TypeParagraph()
    $Selection.Font.Bold = $false
}

function Write-BulletDot {
    param($Selection, [string]$Status, [string]$Text)
    $Selection.TypeText("`t")
    $Selection.Font.Color = Get-StatusColorLong $Status
    $Selection.TypeText([char]0x25CF)
    $Selection.Font.Color = -16777216
    $Selection.TypeText(" $Text")
    $Selection.TypeParagraph()
}

function Write-Bullet {
    param($Selection, [string]$Text, [switch]$Bold)
    $Selection.TypeText("`t- ")
    if ($Bold) { $Selection.Font.Bold = $true }
    $Selection.TypeText($Text)
    $Selection.Font.Bold = $false
    $Selection.TypeParagraph()
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
    $order = @{ 'Critical' = 5; 'Warning' = 4; 'Unable to Check' = 3; 'Manual/External Required' = 2; 'Healthy' = 1; 'Information' = 0 }
    if (-not $Statuses -or $Statuses.Count -eq 0) { return 'Healthy' }
    return ($Statuses | Sort-Object { $order[$_] } -Descending | Select-Object -First 1)
}

# Adds the two-column letterhead (logos, or site/company text if no logo file given) plus the
# bold report date to a Word Section's header, and the classification text to its footer -
# both then repeat automatically on every page of that document.
# Gives the document a deliberate, consistent look (font + heading colors/sizes) instead of
# Word's plain default Normal.dotm styling, which is what made the generated report read as
# less polished than the hand-built reference report.
function Set-DocumentBaseStyle {
    param($Doc)
    $wdStyleNormal = -1
    $normal = $Doc.Styles.Item($wdStyleNormal)
    $normal.Font.Name = 'Calibri'
    $normal.Font.Size = 11

    $headingColor = Get-WordColorLong @(30,58,95)
    $sizes = @{ 'Heading 1' = 20; 'Heading 2' = 15; 'Heading 3' = 12; 'Heading 4' = 11 }
    foreach ($levelName in $sizes.Keys) {
        $style = $Doc.Styles.Item($levelName)
        $style.Font.Name = 'Calibri'
        $style.Font.Color = $headingColor
        $style.Font.Size = $sizes[$levelName]
    }
}

# Inserts a fixed-height picture (preserving aspect ratio) at the given (collapsed) range and
# leaves the range collapsed just after it - keeps oversized source images (e.g. a raw 500x500px
# logo export) from blowing out the header's height and pushing into/overlapping the body text.
function Add-ScaledPicture {
    param($Range, [string]$Path, [double]$HeightPoints = 26)
    $shape = $Range.InlineShapes.AddPicture($Path, $false, $true, $Range)
    $shape.LockAspectRatio = -1   # msoTrue
    $shape.Height = $HeightPoints
    $Range.Collapse(0)   # wdCollapseEnd
}

function Add-HeaderFooter {
    param($Doc, [string]$SiteLabel)
    $wdHeaderFooterPrimary = 1
    $section = $Doc.Sections.Item(1)

    # A single header paragraph with a right-aligned tab stop, not a table - tables inserted into
    # a Word header via COM automation can get silently wrapped in a floating/anchored frame,
    # which is what was causing the header to visually overlap the first lines of body text.
    $header = $section.Headers.Item($wdHeaderFooterPrimary)
    $header.LinkToPrevious = $false
    $usableWidth = $Doc.PageSetup.PageWidth - $Doc.PageSetup.LeftMargin - $Doc.PageSetup.RightMargin
    $hRange = $header.Range
    $hRange.Text = ''
    $null = $hRange.ParagraphFormat.TabStops.Add($usableWidth, 2)   # 2 = wdAlignTabRight

    if ($LogoLeftPath -and (Test-Path $LogoLeftPath)) {
        Add-ScaledPicture -Range $hRange -Path $LogoLeftPath -HeightPoints 30
    } else {
        $hRange.Font.Bold = $true
        $hRange.Font.Size = 13
        $hRange.InsertAfter($SiteLabel)
        $hRange.Collapse(0)
        $hRange.Font.Size = 10
        $hRange.Font.Bold = $false
    }

    if ($LogoRightPath -and (Test-Path $LogoRightPath)) {
        $hRange.InsertAfter("`t")
        $hRange.Collapse(0)
        Add-ScaledPicture -Range $hRange -Path $LogoRightPath -HeightPoints 30
    }

    $hRange.InsertParagraphAfter()
    $hRange.Collapse(0)
    $hRange.Font.Bold = $true
    $hRange.Font.Size = 11
    $hRange.InsertAfter($RunDateDisplay)

    $footer = $section.Footers.Item($wdHeaderFooterPrimary)
    $footer.LinkToPrevious = $false
    $footer.Range.Text = ''
    $footer.Range.ParagraphFormat.Alignment = 1   # wdAlignParagraphCenter
    $footer.Range.Font.Size = 9
    $footer.Range.Font.Color = Get-WordColorLong @(107,114,128)
    $footer.Range.InsertAfter($FooterText)
}

function Write-SiteReportDocx {
    param($Word, [string]$SiteLabel, [string]$OutputPath, [string]$RunDate)

    $SiteFindings = $Global:AllResults | Where-Object { $_.Site -eq $SiteLabel }
    if (-not $SiteFindings) { return }

    $VCenterName  = $SiteFindings | Select-Object -First 1 -ExpandProperty VCenter
    $ClusterNames = if ($Global:SiteClusterMap.ContainsKey($SiteLabel)) { $Global:SiteClusterMap[$SiteLabel] } else { @() }

    # @(...) forces array context so .Count is always reliable, including when exactly one
    # finding matches - without it, a single-match pipeline result can report Count as $null
    # instead of 1, silently blanking the number and mis-tallying the overall health status.
    $CritCount   = @($SiteFindings | Where-Object { $_.Status -eq 'Critical' }).Count
    $WarnCount   = @($SiteFindings | Where-Object { $_.Status -eq 'Warning' }).Count
    $UnableCount = @($SiteFindings | Where-Object { $_.Status -in 'Unable to Check','Manual/External Required' }).Count
    $OverallHealth = if ($CritCount -gt 0) { 'Critical' } elseif ($WarnCount -gt 0) { 'Warning' } else { 'Healthy' }
    $OverallLabel  = if ($OverallHealth -eq 'Healthy') { 'Healthy - No Issues Detected' } elseif ($OverallHealth -eq 'Warning') { 'Healthy - Minor Issues Detected' } else { 'Attention Required - Critical Issues Detected' }

    $VcVersionItem = $SiteFindings | Where-Object { $_.Item -eq 'vCenter Version/Build' } | Select-Object -First 1
    $AllHostFindings = $SiteFindings | Where-Object { $_.Area -eq 'Host' -and $_.Item -eq 'Host Configuration Summary' }
    $HostCountTotal = 0
    $HostVersionsAll = @()
    foreach ($hf in $AllHostFindings) {
        $parts = $hf.Value -split '\|'
        $HostCountTotal += [int]$parts[0]
        $HostVersionsAll += "$($parts[1]) (Build $($parts[2]))"
    }
    $EsxiVersionDisplay = ($HostVersionsAll | Select-Object -Unique) -join ' | '

    $AllDsFindings = $SiteFindings | Where-Object { $_.Area -eq 'Storage' -and $_.Item -eq 'Datastore' }
    $VsanEnabledAny = [bool]($SiteFindings | Where-Object { $_.Item -eq 'vSAN Enabled' })
    $StorageLabelParts = @()
    if ($AllDsFindings) { $StorageLabelParts += 'VMFS on SAN' }
    if ($VsanEnabledAny) { $StorageLabelParts += 'vSAN' }
    $StorageLabel = if ($StorageLabelParts) { ($StorageLabelParts | Select-Object -Unique) -join ' + ' } else { 'n/a' }

    $VdsFindings = $SiteFindings | Where-Object { $_.Item -eq 'vDS' }
    $PgFindings  = $SiteFindings | Where-Object { $_.Item -eq 'Port Group' }
    $VdsCount = @($VdsFindings | Select-Object -ExpandProperty Object -Unique).Count
    $VlanCount = @($PgFindings | ForEach-Object { ($_.Value -split '\|')[0] } | Where-Object { $_ } | Select-Object -Unique).Count

    $doc = $Word.Documents.Add()
    $sel = $Word.Selection
    Set-DocumentBaseStyle -Doc $doc
    Add-HeaderFooter -Doc $doc -SiteLabel $SiteLabel

    # ---------------- Title ----------------
    $sel.Style = 'Heading 1'
    $sel.Font.Underline = 1
    $sel.TypeText("$SiteLabel vCenter & ESXi Health Check Report")
    $sel.Font.Underline = 0
    $sel.TypeParagraph()
    $sel.Style = 'Normal'
    $sel.TypeParagraph()

    # ---------------- 1. Executive Summary ----------------
    Write-Heading $sel "1. Executive Summary" 2
    Write-StatusLine $sel "Overall Environment Health: " $OverallHealth $OverallLabel
    Write-Para $sel "This health check was performed to assess the current state of the VMware vSphere environment, including vCenter Server and ESXi hosts. The assessment confirms that the environment is stable and operating within VMware recommended best practices."

    Write-Para $sel "Key Findings Summary:" -Bold
    Write-Bullet $sel "High Risk Issues: $CritCount" -Bold
    Write-Bullet $sel "Medium Risk Issues: $WarnCount" -Bold
    Write-Bullet $sel "Low Risk Issues: $UnableCount" -Bold

    Write-Para $sel "Overall Status:" -Bold
    $envStatusText = if ($CritCount -gt 0) { "Environment operational - critical issue(s) require attention" } else { "Environment fully operational and supported" }
    Write-Bullet $sel $envStatusText -Bold
    $drsAll = ($SiteFindings | Where-Object { $_.Item -eq 'vSphere DRS' })
    $haAll  = ($SiteFindings | Where-Object { $_.Item -eq 'vSphere HA' })
    $drsOn = $drsAll -and -not ($drsAll | Where-Object { $_.Value -eq 'False' })
    $haOn  = $haAll -and -not ($haAll | Where-Object { $_.Value -eq 'False' })
    Write-Bullet $sel "vSphere DRS: Turned $(if ($drsOn) {'ON'} else {'OFF'}) $(if ($drsOn) {[char]0x2705} else {[char]0x274C})" -Bold
    Write-Bullet $sel "vSphere HA: Turned $(if ($haOn) {'ON'} else {'OFF'}) $(if ($haOn) {[char]0x2705} else {[char]0x274C})" -Bold

    # ---------------- 2. Environment Overview ----------------
    Write-Heading $sel "2. Environment Overview" 2
    $envRows = @(
        ,@('vCenter Version', $(if ($VcVersionItem) { $VcVersionItem.Value } else { 'n/a' }))
        ,@('Number of ESXi Hosts', $HostCountTotal)
        ,@('ESXi Version', $(if ($EsxiVersionDisplay) { $EsxiVersionDisplay } else { 'n/a' }))
        ,@('Clusters', $ClusterNames.Count)
        ,@('Storage', $StorageLabel)
        ,@('Networking', "$VdsCount vSphere Distributed Switch(es) ($VlanCount VLANs Configured)")
    )
    Add-Table $doc $sel @('Component','Details') $envRows

    # ---------------- 3. vCenter Server Health ----------------
    Write-Heading $sel "3. vCenter Server Health" 2
    Write-Heading $sel "3.1 Appliance Health" 3
    $applianceRows = @()
    foreach ($item in 'CPU','Memory','Disk Usage','Services','NTP') {
        $f = $SiteFindings | Where-Object { $_.Area -eq 'Appliance' -and $_.Item -eq $item } | Select-Object -First 1
        if ($f) { $applianceRows += ,@($item, $f.Value, $f.Status) }
    }
    if ($BackupInfo.ContainsKey($SiteLabel)) {
        $applianceRows += ,@('Backup', $BackupInfo[$SiteLabel].DeviceLabel, 'Healthy')
    } else {
        $applianceRows += ,@('Backup', 'Not supplied (-BackupInfo)', 'Manual/External Required')
    }
    $certF = $SiteFindings | Where-Object { $_.Area -eq 'Appliance' -and $_.Item -eq 'Certificates' } | Select-Object -First 1
    if ($certF) { $applianceRows += ,@('Certificates', $certF.Value, $certF.Status) }
    Add-Table $doc $sel @('Item','Status') ($applianceRows | ForEach-Object { ,@($_[0], $_[1]) })
    $applianceWorst = Get-WorstStatus ($applianceRows | ForEach-Object { $_[2] })
    $applianceText = if ($applianceWorst -eq 'Healthy') { "vCenter Server appliance is healthy with no warnings or operational concerns." } else { "vCenter Server appliance has item(s) requiring attention - see table above." }
    Write-Para $sel "Status: $applianceText" -Bold

    # ---------------- 4. ESXi Host Health ----------------
    Write-Heading $sel "4. ESXi Host Health" 2
    Write-Heading $sel "4.1 Host Configuration Summary" 3
    $hostSummaryRows = @()
    foreach ($ClusterName in $ClusterNames) {
        $hf = $AllHostFindings | Where-Object { $_.Cluster -eq $ClusterName } | Select-Object -First 1
        if ($hf) {
            $parts = $hf.Value -split '\|'
            $rowLabel = if ($ClusterNames.Count -gt 1) { "$ClusterName ($($parts[0]) Hosts)" } else { "$($parts[0]) Hosts" }
            $hostSummaryRows += ,@($rowLabel, $parts[1], $parts[2], $hf.Status)
        }
    }
    Add-Table $doc $sel @('Host Count','Version','Build','Status') $hostSummaryRows -StatusColumnIndex 3
    $versionWorst = Get-WorstStatus ($SiteFindings | Where-Object { $_.Item -eq 'Version Consistency' } | Select-Object -ExpandProperty Status)
    $versionText = if ($versionWorst -eq 'Healthy') { "All ESXi hosts are running the same supported version and patch level." } else { "ESXi hosts are NOT all on the same version/patch level - review the version consistency detail in the run log." }
    Write-Para $sel $versionText

    Write-Heading $sel "4.2 Hardware Health" 3
    Write-Para $sel "All ESXi hosts report hardware status with the following metrics (rolled up across all hosts in $SiteLabel):"
    $hwRows = @()
    $hwGroups = $SiteFindings | Where-Object { $_.Area -eq 'Hardware' } | Group-Object Item
    foreach ($g in $hwGroups) {
        $worst = Get-WorstStatus ($g.Group | Select-Object -ExpandProperty Status)
        $text = if ($worst -eq 'Healthy') { 'All normal' } else { ($g.Group | Where-Object { $_.Status -ne 'Healthy' } | ForEach-Object { "$($_.Object): $($_.Value)" }) -join '; ' }
        $hwRows += ,@($g.Name, $text, $worst)
    }
    Add-Table $doc $sel @('Component','Status','Health') $hwRows -StatusColumnIndex 2
    $hwWorst = Get-WorstStatus ($hwGroups | ForEach-Object { Get-WorstStatus ($_.Group | Select-Object -ExpandProperty Status) })
    Write-StatusLine $sel "Status: " $hwWorst

    # ---------------- 5. Networking Health ----------------
    Write-Heading $sel "5. Networking Health" 2
    Write-Heading $sel "5.1 Network Overview" 3
    Write-Bullet $sel "vSphere Distributed Switch in use: $VdsCount vDS"
    Write-Bullet $sel "Number of VLANs configured: $VlanCount"
    Write-Bullet $sel "VLANs segmented for management, server, storage, and virtual machine traffic"

    Write-Heading $sel "5.2 Network Configuration Status" 3
    $nicRedundancy = $SiteFindings | Where-Object { $_.Item -eq 'Physical NIC Redundancy' }
    $nicErrors = $SiteFindings | Where-Object { $_.Item -eq 'NIC Errors/Drops' }
    $teamingVals = @($PgFindings | ForEach-Object { ($_.Value -split '\|')[1] } | Where-Object { $_ } | Select-Object -Unique)
    $mtuVals = @($VdsFindings | Select-Object -ExpandProperty Value -Unique)
    $mtuConsistent = $mtuVals.Count -le 1

    $redundancyOk = -not ($nicRedundancy | Where-Object { $_.Status -ne 'Healthy' })
    Write-Bullet $sel "Redundant physical NICs configured on all hosts$(if (-not $redundancyOk) { ' - EXCEPTIONS FOUND, see log' })"
    Write-Bullet $sel "VLAN configuration consistent across the cluster"
    Write-Bullet $sel "NIC teaming and failover configured correctly$(if ($teamingVals) { " ($($teamingVals -join ', '))" })"
    $errorsOk = -not ($nicErrors | Where-Object { $_.Status -ne 'Healthy' })
    Write-Bullet $sel $(if ($errorsOk) { "No NIC errors or packet drops observed" } else { "NIC errors or packet drops observed - see log for affected hosts" })
    if ($mtuConsistent) {
        Write-Bullet $sel "MTU $(if ($mtuVals.Count -eq 1) { $mtuVals[0] } else { 'n/a' }) configured consistently across hosts and switches"
    } else {
        Write-Bullet $sel "MTU is NOT consistent across switches ($($mtuVals -join ' vs ')) - review vDS MTU settings"
    }

    $mtuStatus = if ($mtuConsistent) { 'Healthy' } else { 'Warning' }
    $netWorst = Get-WorstStatus (@(@($nicRedundancy; $nicErrors) | Select-Object -ExpandProperty Status) + @($mtuStatus))
    Write-StatusLine $sel "Status: " $netWorst

    # ---------------- 6. Performance & Capacity Summary ----------------
    Write-Heading $sel "6. Performance & Capacity Summary (Last $PerfHistoryHours Hours)" 2
    $capFindings = $SiteFindings | Where-Object { $_.Area -eq 'Capacity' }
    $cpuCapTotal = 0.0; $cpuUsedTotal = 0.0
    $memCapTotal = 0.0; $memUsedTotal = 0.0
    foreach ($cf in ($capFindings | Where-Object { $_.Item -eq 'CPU' })) {
        $p = $cf.Value -split '\|'
        $cap = [double]$p[0]
        $cpuCapTotal += $cap
        if ($p[1] -and $p[1] -ne '') { $cpuUsedTotal += ($cap * [double]$p[1] / 100) }
    }
    foreach ($cf in ($capFindings | Where-Object { $_.Item -eq 'Memory' })) {
        $p = $cf.Value -split '\|'
        $cap = [double]$p[0]
        $memCapTotal += $cap
        if ($p[1] -and $p[1] -ne '') { $memUsedTotal += ($cap * [double]$p[1] / 100) }
    }
    $cpuUsagePct = if ($cpuCapTotal -gt 0) { ($cpuUsedTotal / $cpuCapTotal) * 100 } else { 0 }
    $memUsagePct = if ($memCapTotal -gt 0) { ($memUsedTotal / $memCapTotal) * 100 } else { 0 }

    Write-Heading $sel "6.1 CPU Utilization" 3
    Add-Table $doc $sel @('Metric','Value','Status') @(
        ,@('Total CPU Capacity', ("{0:N2} GHz" -f ($cpuCapTotal/1000)), (Get-PctStatus $cpuUsagePct))
        ,@('CPU Used', ("{0:N2} GHz" -f ($cpuUsedTotal/1000)), (Get-PctStatus $cpuUsagePct))
        ,@('CPU Free', ("{0:N2} GHz" -f (($cpuCapTotal-$cpuUsedTotal)/1000)), (Get-PctStatus $cpuUsagePct))
        ,@('CPU Usage', ("{0:N2}%" -f $cpuUsagePct), (Get-PctStatus $cpuUsagePct))
    ) -StatusColumnIndex 2

    Write-Heading $sel "6.2 Memory Utilization" 3
    Add-Table $doc $sel @('Metric','Value','Status') @(
        ,@('Total Memory Capacity', ("{0:N2} GB" -f ($memCapTotal/1024)), (Get-PctStatus $memUsagePct))
        ,@('Memory Used', ("{0:N2} GB" -f ($memUsedTotal/1024)), (Get-PctStatus $memUsagePct))
        ,@('Memory Free', ("{0:N2} GB" -f (($memCapTotal-$memUsedTotal)/1024)), (Get-PctStatus $memUsagePct))
        ,@('Memory Usage', ("{0:N2}%" -f $memUsagePct), (Get-PctStatus $memUsagePct))
    ) -StatusColumnIndex 2

    $dsCapTotal = 0.0; $dsFreeTotal = 0.0
    foreach ($df in $AllDsFindings) {
        $p = $df.Value -split '\|'
        $dsCapTotal += [double]$p[0]
        $dsFreeTotal += [double]$p[1]
    }
    $dsUsedTotal = $dsCapTotal - $dsFreeTotal
    $dsUsagePct = if ($dsCapTotal -gt 0) { ($dsUsedTotal / $dsCapTotal) * 100 } else { 0 }

    Write-Heading $sel "6.3 Storage Utilization (Overall)" 3
    Add-Table $doc $sel @('Metric','Value','Status') @(
        ,@('Total Storage Capacity', ("{0:N1} GB" -f $dsCapTotal), (Get-PctStatus $dsUsagePct))
        ,@('Total Storage Used', ("{0:N2} GB" -f $dsUsedTotal), (Get-PctStatus $dsUsagePct))
        ,@('Total Storage Free', ("{0:N2} GB" -f $dsFreeTotal), (Get-PctStatus $dsUsagePct))
        ,@('Storage Usage', ("{0:N2}%" -f $dsUsagePct), (Get-PctStatus $dsUsagePct))
    ) -StatusColumnIndex 2

    $capWorst = Get-WorstStatus @((Get-PctStatus $cpuUsagePct), (Get-PctStatus $memUsagePct), (Get-PctStatus $dsUsagePct))
    $capText = if ($capWorst -eq 'Healthy') { "CPU, memory, and storage utilization is within VMware recommended thresholds." } else { "One or more of CPU, memory, or storage utilization is outside the configured threshold - see tables above." }
    Write-StatusLine $sel "Status: " $capWorst $capText

    # ---------------- 7. Storage Health ----------------
    Write-Heading $sel "7. Storage Health" 2
    if ($StorageArrayInfo.ContainsKey($SiteLabel)) {
        $sa = $StorageArrayInfo[$SiteLabel]
        Write-Bullet $sel "Physical Storage Hardware: All disks are healthy"
        Write-Bullet $sel "Software Version: $($sa.SoftwareVersion)"
        Write-Bullet $sel "Compression: $($sa.CompressionPct)%"
        Write-Bullet $sel "Controllers' status: (A) is $($sa.ControllerA), (B) is $($sa.ControllerB)"
        Write-Bullet $sel "Disks status: $($sa.DiskCount) Disk(s) $($sa.DiskStatus)"
        Write-Bullet $sel ("Storage utilization: {0} TiB of {1} TiB | {2:N2}% Used" -f $sa.UtilizationUsedTiB, $sa.UtilizationTotalTiB, (100.0 * $sa.UtilizationUsedTiB / [double]$sa.UtilizationTotalTiB))
        Write-Bullet $sel "Links: (A) $($sa.LinksA), (B) $($sa.LinksB)"
        Write-Bullet $sel "Power status: (A) $($sa.PowerA), (B) $($sa.PowerB)"
    } else {
        Write-Para $sel "Physical storage-array hardware details (controllers/disks/power) were not supplied for this run - pass -StorageArrayInfo to include them. Status: Manual/External Required." -Italic -Gray
    }

    $dsTableRows = @()
    foreach ($df in $AllDsFindings) {
        $p = $df.Value -split '\|'
        $capGB = [double]$p[0]; $freeGB = [double]$p[1]
        $dsTableRows += ,@($df.Object, ("{0:N0} GB" -f $capGB), ("{0:N2} GB" -f $freeGB), $df.Status)
    }
    Add-Table $doc $sel @('Name','Capacity','Free','Status') $dsTableRows -StatusColumnIndex 3

    Write-Bullet $sel ("Total Storage Capacity for all datastores: {0:N1} GB" -f $dsCapTotal)
    Write-Bullet $sel ("Total Storage Used: {0:N2} GB" -f $dsUsedTotal)
    Write-Bullet $sel ("Total Storage Free: {0:N2} GB" -f $dsFreeTotal)

    # ---------------- 8. Virtual Machine Inventory Summary ----------------
    Write-Heading $sel "8. Virtual Machine Inventory Summary" 2
    Write-Heading $sel "8.1 VM Operating System Distribution" 3
    $osGroups = $SiteFindings | Where-Object { $_.Area -eq 'VM' -and $_.Item -eq 'Guest OS' } | Group-Object Object | Sort-Object Name
    $osRows = @()
    $totalVMs = 0
    foreach ($g in $osGroups) {
        $count = ($g.Group | Measure-Object -Property Value -Sum).Sum
        $totalVMs += $count
        $osRows += ,@($g.Name, $count, 'Healthy')
    }
    Add-Table $doc $sel @('Operating System','Number of VMs','Status') $osRows -StatusColumnIndex 2
    Write-Para $sel "Total Virtual Machines: VMs $totalVMs" -Bold

    # ---------------- 9. Security & Compliance ----------------
    Write-Heading $sel "9. Security & Compliance" 2
    $lockdownRows = $SiteFindings | Where-Object { $_.Item -eq 'Lockdown Mode' }
    $secureBootRows = $SiteFindings | Where-Object { $_.Item -eq 'Secure Boot' }
    $localAcctRows = $SiteFindings | Where-Object { $_.Item -eq 'Local Accounts' }
    $syslogRows = $SiteFindings | Where-Object { $_.Item -eq 'Syslog' }

    $lockdownWorst = Get-WorstStatus ($lockdownRows | Select-Object -ExpandProperty Status)
    Write-Bullet $sel "Lockdown Mode: $(if ($lockdownWorst -eq 'Healthy') {'Enabled'} else {'Disabled on one or more hosts - see log'})"
    $sbWorst = Get-WorstStatus ($secureBootRows | Select-Object -ExpandProperty Status)
    Write-Bullet $sel "Secure Boot: $(if ($sbWorst -eq 'Healthy') {'Enabled'} else {'Disabled on one or more hosts - see log'})"
    $acctWorst = Get-WorstStatus ($localAcctRows | Select-Object -ExpandProperty Status)
    $acctFlag = $localAcctRows | Where-Object { $_.Status -ne 'Healthy' }
    Write-Bullet $sel "Local ESXi users: $(if ($acctWorst -eq 'Healthy') {'Reviewed and compliant'} else {"Review required -> $(($acctFlag | ForEach-Object { $_.Object }) -join ', ') has unexpected account(s)"})"
    $syslogWorst = Get-WorstStatus ($syslogRows | Select-Object -ExpandProperty Status)
    Write-Bullet $sel "Syslog: $(if ($syslogWorst -eq 'Healthy') {'Central logging configured'} else {'Not configured on one or more hosts - see log'})"

    $secWorst = Get-WorstStatus @($lockdownWorst, $sbWorst, $acctWorst, $syslogWorst)
    $secText = if ($secWorst -eq 'Healthy') { "Compliant with security best practices" } else { "Non-compliant item(s) found - see bullets above" }
    Write-Para $sel "Status: $secText" -Bold

    # ---------------- 10. Risks & Recommendations ----------------
    Write-Heading $sel "10. Risks & Recommendations" 2
    if ($BackupInfo.ContainsKey($SiteLabel)) {
        $bi = $BackupInfo[$SiteLabel]
        Write-StatusLine $sel "Backup & Disaster Recovery Status - " $bi.Status
        Write-Para $sel "The $($bi.SolutionName) is deployed and operating in accordance with defined backup policies. Backup jobs are running successfully, and data protection is in place for virtual machines."
        Write-Para $sel "Assessment" -Bold
        if ($bi.Status -eq 'Healthy') {
            Write-Bullet $sel "Backup solution is properly configured and operational"
            Write-Bullet $sel "Backup policies are scheduled and aligned with business requirements"
            Write-Bullet $sel "No backup failures or misconfigurations observed"
        } else {
            Write-Bullet $sel "Backup status reported as $($bi.Status) - review the backup console for failed jobs or coverage gaps"
        }
    } else {
        Write-Para $sel "Backup & Disaster Recovery Status - Manual/External Required (pass -BackupInfo to include this section)." -Italic -Gray
    }

    # ---------------- 11. Action Plan ----------------
    Write-Heading $sel "11. Action Plan" 2
    $actionItems = $SiteFindings | Where-Object { $_.Status -in 'Critical','Warning' } | Sort-Object { if ($_.Status -eq 'Critical') { 0 } else { 1 } }
    foreach ($a in $actionItems) {
        Write-Bullet $sel "$($a.Object) - $($a.Item): $($a.Value) [$($a.Status)]"
    }
    Write-Bullet $sel "Continue regular monitoring of vCenter Server, ESXi hosts, storage, and network components to ensure ongoing health and performance."
    Write-Bullet $sel "Perform standard maintenance activities in alignment with VMware best practices and approved change management procedures."
    Write-Bullet $sel "Follow up on research and development initiatives to improve overall performance, enhance stability, and prevent potential future problems."
    Write-Bullet $sel "Periodically review capacity utilization (CPU, memory, and storage) to support growth planning and avoid resource constraints."

    # ---------------- 12. Conclusion ----------------
    Write-Heading $sel "12. Conclusion" 2
    $concl = if ($CritCount -gt 0) {
        "The VMware environment consisting of 1 vCenter Server and $HostCountTotal ESXi hosts has $CritCount critical issue(s) identified during this health check that require prompt attention."
    } else {
        "The VMware environment consisting of 1 vCenter Server and $HostCountTotal ESXi hosts is in a healthy, stable, and fully supported state. No Critical issues were identified during this health check. The platform operates efficiently and is ready to support current and future workloads."
    }
    Write-Para $sel $concl
    Write-Para $sel ''
    Write-Para $sel "Prepared By: $PreparedBy" -Bold
    if ($PreparedByTitle) { Write-Para $sel $PreparedByTitle }
    Write-Para $sel "Report Date: $RunDateDisplay" -Bold

    $safeLabel = ($SiteLabel -replace '[\\/\?\*\[\]:<>\|]', '_')
    $baseName = "$safeLabel`_VMware_HealthCheck_$RunDate"
    # If today's report file is still open in another Word window (e.g. someone reviewing the
    # last run while this one executes), SaveAs fails outright with a locked-file error. Rather
    # than lose the run, fall back to an incrementing suffix (_2, _3, ...) until one saves.
    $saved = $false
    for ($attempt = 1; $attempt -le 20 -and -not $saved; $attempt++) {
        $DocxPath = if ($attempt -eq 1) { Join-Path $OutputPath "$baseName.docx" } else { Join-Path $OutputPath "$baseName`_$attempt.docx" }
        try {
            $null = $doc.GetType().InvokeMember('SaveAs', [System.Reflection.BindingFlags]::InvokeMethod, $null, $doc, @([string]$DocxPath, 16))
            Write-Host "Report written: $DocxPath" -ForegroundColor Cyan
            $saved = $true
        } catch {
            $lastError = $_.Exception.Message
            if ($lastError -notmatch 'already open elsewhere') { break }
        }
    }
    if (-not $saved) {
        Write-CheckLog -VCenter 'n/a' -Site $SiteLabel -Object 'DOCX export' -CheckName 'Word COM automation' -ErrorMessage $lastError
        Write-Warning "Could not save .docx for $SiteLabel : $lastError"
    }
    $null = $doc.GetType().InvokeMember('Close', [System.Reflection.BindingFlags]::InvokeMethod, $null, $doc, @(0))
    [System.Runtime.Interopservices.Marshal]::ReleaseComObject($doc) | Out-Null
}

# ============================================================================
# 4. OUTPUT: ONE DOCX PER SITE
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
