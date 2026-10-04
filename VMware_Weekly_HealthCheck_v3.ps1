#Requires -Version 5.1
<#
.SYNOPSIS
    VMware Weekly Health Check - Read-only automated collection and HTML dashboard generation.

.DESCRIPTION
    Replaces the manual weekly vCenter/ESXi health-check reports with a single read-only
    PowerCLI collection run from an already-authenticated laptop session.

    - Uses whatever vCenter sessions are already connected (Connect-VIServer done beforehand).
    - Auto-discovers clusters, hosts, datastores, vSAN, vDS/port groups and VMs per vCenter -
      nothing about the infrastructure is hard-coded.
    - Produces ONE combined HTML dashboard covering every connected vCenter ("site"), with an
      Overview page (most important points across all sites) and one full-detail page per site
      covering:
        - Executive summary (risk counts, DRS/HA)
        - Environment overview (versions, hosts, clusters, storage type)
        - Networking health (vDS/VLANs, NIC teaming, MTU consistency, NIC redundancy)
        - Performance & capacity (CPU/Memory/Storage utilization)
        - Storage / datastores (full per-datastore table)
        - VM inventory / guest OS distribution
        - Security & compliance (Lockdown Mode, Secure Boot, local accounts, syslog)
        - Backup & disaster recovery
        - Action plan (itemized Warning/Critical findings)
        - Per-cluster breakdown
      Appliance health and host hardware-sensor checks are intentionally NOT collected or
      reported - out of scope for this dashboard.
    - When a vCenter has more than one cluster, per-cluster data (hosts, capacity, VM counts)
      is aggregated up to the site level for the dashboard, with its own per-cluster table too.
    - A failure collecting one item is logged and skipped; the script always continues.
    - NEVER writes/modifies anything in vCenter. Every cmdlet used below is read-only
      (Get-*, no Set-*/New-*/Remove-*, no config changes, no SSH enablement).

.NOTES
    Every threshold below (capacity %, license expiry windows) is a CONFIGURABLE PARAMETER, not
    an assumed company standard - raw values/status are always shown alongside any computed flag.

    Backup & Disaster Recovery status comes from the backup product's own console (e.g. Cohesity,
    Veeam), not vCenter - supply/override it via -BackupInfo (defaults are pre-filled for the
    current real sites). If a site isn't in -BackupInfo at all, that panel is rendered as
    "Manual/External Required" rather than fabricated.

    -SiteMapPath (default C:\temp\VMware_Weekly_Health_Check_SiteMap.xml) is loaded automatically if the file exists, so you
    don't have to retype -SiteMap on every run. Expected shape - one <Site> element per vCenter:
        <SiteMap>
          <Site vCenter="tb-dhci-vc01.seventb.local" Name="Tabuk" />
          <Site vCenter="sf-vc.sixflags.local" Name="SF" />
          <Site vCenter="amc-vc.amc.local" Name="SceneCinema" />
        </SiteMap>
    "vCenter" must match the server name exactly as it appears when connected (the same string
    shown in "Connected vCenter sessions: ..." when the script starts). Missing/unreadable file is
    not an error - the script just falls back to -SiteMap / the automatic name detection below.

.EXAMPLE
    # Already connected: Connect-VIServer tb-vc.aq.local - overrides the built-in -BackupInfo
    # default for this one site, e.g. once its Cohesity job history for the week is checked.
    .\VMware_Weekly_HealthCheck.ps1 -SiteMap @{'tb-vc.aq.local'='SEVEN Tabuk'} `
        -BackupInfo @{ 'SEVEN Tabuk' = @{ Appliance='Cohesity'; Schedule='2:00 AM';
            Retention='15 Day - 4 Weeks - 1 Month'; Status='Healthy' } }
#>

[CmdletBinding()]
param(
    # Maps a connected vCenter server (Name as shown in $global:DefaultVIServers) to a friendly
    # site label used in the dashboard. If a connected vCenter isn't in this map, its own server
    # name is used as the label - nothing is hard-coded or required. Entries here always take
    # precedence over -SiteMapPath below for the same vCenter.
    [hashtable]$SiteMap = @{},

    # Optional XML file of the same vCenter->site-label mappings, so you don't have to retype
    # -SiteMap on every run. Loaded automatically if present - see the XML shape in .NOTES.
    # Entries in -SiteMap above override a matching entry here.
    [string]$SiteMapPath = 'C:\temp\VMware_Weekly_Health_Check_SiteMap.xml',

    [string]$OutputPath = (Join-Path $PSScriptRoot "VMware_HealthCheck_Reports"),

    # ---- Data NOT available from vCenter - keyed by Site label (see .NOTES) ---------------
    # Appliance/Schedule/Retention/Status per site, plus an optional Notes reason shown when
    # Status is Warning/Critical (e.g. an expired license) - defaults reflect the current real
    # backup setup: Cohesity nightly at 2:00 AM for every site except SceneCinema, which runs Veeam at
    # 5:00 AM and is flagged Critical because its license has expired. SEVEN ALhamra is a newly
    # handed-over site whose backup management hasn't been handed over to this team yet - Status
    # 'Manual/External Required' (with a Notes reason) renders it as a distinct "not yet ours"
    # state rather than either a health status or the generic "no data supplied" panel.
    [hashtable]$BackupInfo = @{
        'SF-AQ'         = @{ Appliance = 'Cohesity'; Schedule = '2:00 AM'; Retention = '15 Day - 4 Weeks - 1 Month'; Status = 'Healthy' }
        'SEVEN Tabuk'   = @{ Appliance = 'Cohesity'; Schedule = '2:00 AM'; Retention = '15 Day - 4 Weeks - 1 Month'; Status = 'Healthy' }
        'SEVEN ABHA'    = @{ Appliance = 'Cohesity'; Schedule = '2:00 AM'; Retention = '15 Day - 4 Weeks - 1 Month'; Status = 'Healthy' }
        'SEVEN ALhamra' = @{ Status = 'Manual/External Required'; Notes = 'Site recently handed over - backup management not yet handed over to this team.' }
        'SceneCinema'   = @{ Appliance = 'Veeam'; Schedule = '5:00 AM'; Retention = '15 Day - 4 Weeks - 1 Month'; Status = 'Critical'; Notes = 'Veeam license expired' }
    },

    # Local ESXi accounts considered normal/expected; anything extra found on a host is flagged.
    [string[]]$ExpectedLocalAccounts = @('root','dcui','vpxuser'),

    # Cluster names to skip entirely (e.g. one that's been migrated to another platform but is
    # still visible in vCenter inventory) - excluded from every section of every site's report and
    # the dashboard, as if it didn't exist. Matched by exact cluster name, case-insensitive.
    # RGL defaults on since it's already been migrated off this vCenter - override with
    # -ExcludeClusters @() to include it again, or add more names as needed.
    [string[]]$ExcludeClusters = @('RGL'),

    # Display order for site tabs/rows in the HTML dashboard (Overview always comes first,
    # regardless of this list). Any connected site NOT in this list is appended afterward,
    # alphabetically, so a new/unlisted site still appears rather than being dropped.
    [string[]]$DashboardSiteOrder = @('SF-AQ','SEVEN Tabuk','SEVEN ABHA','SEVEN ALhamra','SceneCinema'),

    # Configurable thresholds - NOT vendor/company-defined standards. Raw values are always
    # shown regardless of these; these only drive the Warning/Critical flag shown alongside them.
    [double]$CapacityWarningPct    = 80,
    [double]$CapacityCriticalPct   = 90,
    [int]$LicenseExpiryWarningDays = 30,

    # Historical performance window for capacity stats (hours), matches the 24-72h window used
    # in the reference report.
    [int]$PerfHistoryHours = 24
)

# Bump this on every change. Printed first thing at startup and written into the log file, so
# it's always possible to confirm exactly which script version produced a given run/report
# instead of guessing whether an old cached copy is being executed somewhere.
$ScriptBuild = '2026-09-29-23-html-only-dashboard-no-appliance-hardware'
Write-Host "VMware_Weekly_HealthCheck.ps1 - build $ScriptBuild" -ForegroundColor Magenta

$ErrorActionPreference = 'Stop'
$ScriptStart = Get-Date
$RunDate     = $ScriptStart.ToString('yyyy-MM-dd')
$RunDateDisplay = $ScriptStart.ToString('d-MMMM-yyyy', [System.Globalization.CultureInfo]::InvariantCulture)
if (-not (Test-Path $OutputPath)) { New-Item -ItemType Directory -Path $OutputPath -Force | Out-Null }

# Load site-name mappings from -SiteMapPath, if present, so -SiteMap doesn't need retyping every
# run. An explicit -SiteMap entry for the same vCenter still wins - only fills in what's missing.
# Deliberately noisy either way (found/not found/loaded-what) so a filename or path mismatch shows
# up immediately here instead of only being noticeable later in the report filenames.
if (Test-Path $SiteMapPath) {
    try {
        [xml]$SiteMapXml = Get-Content -Path $SiteMapPath -Raw
        foreach ($entry in @($SiteMapXml.SiteMap.Site)) {
            if ($entry.vCenter -and $entry.Name -and -not $SiteMap.ContainsKey($entry.vCenter)) {
                $SiteMap[$entry.vCenter] = $entry.Name
            }
        }
        Write-Host "Loaded site names from $SiteMapPath" -ForegroundColor Cyan
    } catch {
        Write-Warning "Could not read -SiteMapPath '$SiteMapPath': $($_.Exception.Message) - continuing without it."
    }
} else {
    Write-Host "No site-map file found at '$SiteMapPath' - reports will use automatic/raw vCenter names unless -SiteMap is passed." -ForegroundColor Yellow
}
if ($SiteMap.Count -gt 0) {
    Write-Host "Site name mapping in effect:" -ForegroundColor Cyan
    $SiteMap.GetEnumerator() | ForEach-Object { Write-Host "  $($_.Key)  ->  $($_.Value)" -ForegroundColor Cyan }
} else {
    Write-Host "No site name mapping in effect - every report will be named after its raw vCenter connection string." -ForegroundColor Yellow
}

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
    #   Overview | Alarms | Cluster | Host | Security | Networking | Capacity | Storage | VM
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
        if ($instanceName) {
            $Site = $instanceName
        } elseif ($VCName -match '^\d{1,3}(\.\d{1,3}){3}$') {
            # Raw IP address - nothing shorter/friendlier to derive from it.
            $Site = $VCName
        } else {
            # No InstanceName set either - use just the short hostname (before the first dot)
            # instead of the full FQDN, e.g. "tb-dhci-vc01.seventb.local" -> "tb-dhci-vc01".
            $Site = ($VCName -split '\.')[0]
        }
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
        $allClusters = Get-Cluster -Server $VC
        # Substring match rather than exact equality - a trailing space, prefix, or suffix in the
        # real cluster name (e.g. "RGL-01" or "SceneCinema_RGL") would silently defeat an exact match, and
        # exclusion failing SILENTLY is worse than it matching a little too broadly.
        $kept = $allClusters | Where-Object {
            $clusterName = $_.Name
            -not ($ExcludeClusters | Where-Object { $clusterName -like "*$_*" })
        }
        $excluded = @($allClusters | Where-Object { $_.Name -notin $kept.Name })
        if ($excluded.Count -gt 0) {
            Write-Host "  Excluding cluster(s) for $Site : $($excluded.Name -join ', ')" -ForegroundColor Yellow
        }
        $kept
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
                # not an application/server VLAN - counting it inflated the VLAN total. Belt-and-braces:
                # check both the IsUplink flag AND the standard DVUplinks naming, since IsUplink has
                # been observed unreliable on at least one third-party-managed vDS (e.g. Apstra).
                $pgs = Get-VDPortgroup -Server $VC -VDSwitch $vds | Where-Object { -not $_.IsUplink -and $_.Name -notmatch '-DVUplinks-' }
                foreach ($pg in $pgs) {
                    # Regular port groups expose .VlanId as a plain int - only these count as one
                    # "configured VLAN" below. Private VLAN port groups use .PvlanId instead, and
                    # trunk port groups expose a range - both are shown per-port-group for visibility
                    # but deliberately excluded from the VLAN count itself, since a trunk range isn't
                    # one countable VLAN (this is also a second, independent guard against any
                    # remaining uplink-like port group that slips past the IsUplink/name filter above).
                    $vlanCfg = $pg.ExtensionData.Config.DefaultPortConfig.Vlan
                    $isSingleVlan = $null -ne $vlanCfg.VlanId -and $vlanCfg.VlanId -is [int]
                    $vlan = if ($isSingleVlan) { $vlanCfg.VlanId }
                        elseif ($vlanCfg.PvlanId) { "PVLAN $($vlanCfg.PvlanId)" }
                        elseif ($vlanCfg.VlanId) { "Trunk " + (($vlanCfg.VlanId | ForEach-Object { "$($_.Start)-$($_.End)" }) -join ',') }
                        else { $null }
                    $teaming = $pg.ExtensionData.Config.DefaultPortConfig.UplinkTeamingPolicy.Policy.Value
                    New-Finding -Site $Site -VCenter $VCName -Area 'Networking' -Cluster $ClusterName -Object $pg.Name `
                        -Item 'Port Group' -Value "$vlan|$teaming|$isSingleVlan" -Status 'Information'
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
# 3. DASHBOARD GENERATION HELPERS
# ============================================================================
function Get-WorstStatus {
    param([string[]]$Statuses)
    $order = @{ 'Critical' = 5; 'Warning' = 4; 'Unable to Check' = 3; 'Manual/External Required' = 2; 'Healthy' = 1; 'Information' = 0 }
    if (-not $Statuses -or $Statuses.Count -eq 0) { return 'Healthy' }
    return ($Statuses | Sort-Object { $order[$_] } -Descending | Select-Object -First 1)
}

# A handful of Item values are internally pipe-delimited (e.g. Capacity's "1234|56.7" or Storage's
# "capGB|freeGB|accessible") so New-Finding's caller can pack a few raw numbers in without adding
# dedicated columns - fine for the aggregation code that already knows the shape, but unreadable if
# shown to a person as-is. The Action Plan is the one place a raw Finding.Value reaches the page
# directly, so this decodes those specific Items back into plain text; everything else (MTU
# Consistency, Syslog, Local Accounts, etc.) already stores a human-readable Value and passes through.
function Get-ActionItemDisplayValue {
    param([string]$Item, [string]$Value)
    switch ($Item) {
        'CPU' {
            $p = $Value -split '\|'
            if ($p.Count -ge 2 -and $p[1]) { return ("{0:N1}% CPU usage" -f [double]$p[1]) }
        }
        'Memory' {
            $p = $Value -split '\|'
            if ($p.Count -ge 2 -and $p[1]) { return ("{0:N1}% Memory usage" -f [double]$p[1]) }
        }
        'Datastore' {
            $p = $Value -split '\|'
            if ($p.Count -ge 2) {
                $capGB = [double]$p[0]; $freeGB = [double]$p[1]
                $usedPct = if ($capGB -gt 0) { (($capGB - $freeGB) / $capGB) * 100 } else { 0 }
                return ("{0:N1}% used ({1:N2} GB free of {2:N0} GB)" -f $usedPct, $freeGB, $capGB)
            }
        }
        'vSphere HA'  { if ($Value -eq 'False') { return 'Turned OFF' } }
        'vSphere DRS' { if ($Value -eq 'False') { return 'Turned OFF' } }
    }
    return $Value
}

# Escapes text for safe inclusion in the HTML dashboard - avoids any dependency on
# System.Web (not guaranteed loaded) for a handful of characters that matter here.
function ConvertTo-HtmlSafe {
    param([string]$Text)
    if (-not $Text) { return '' }
    return $Text -replace '&','&amp;' -replace '<','&lt;' -replace '>','&gt;' -replace '"','&quot;'
}

# Computes every per-site aggregate the HTML dashboard needs (health status, risk counts,
# host/VM/cluster counts, DRS/HA, capacity, full datastore list, OS distribution, networking
# detail, security/compliance detail, backup detail, itemized action-plan findings) directly from
# $Global:AllResults. Appliance health and hardware-sensor data are never referenced here - those
# checks are not collected at all.
function Get-SiteDashboardSummary {
    param([string]$SiteLabel)
    $SiteFindings = $Global:AllResults | Where-Object { $_.Site -eq $SiteLabel }
    if (-not $SiteFindings) { return $null }

    $CritCount   = @($SiteFindings | Where-Object { $_.Status -eq 'Critical' }).Count
    $WarnCount   = @($SiteFindings | Where-Object { $_.Status -eq 'Warning' }).Count
    $UnableCount = @($SiteFindings | Where-Object { $_.Status -in 'Unable to Check','Manual/External Required' }).Count
    $OverallHealth = if ($CritCount -gt 0) { 'Critical' } elseif ($WarnCount -gt 0) { 'Warning' } else { 'Healthy' }

    $VcVersionItem = $SiteFindings | Where-Object { $_.Item -eq 'vCenter Version/Build' } | Select-Object -First 1
    $VCenterVersionText = if ($VcVersionItem) { $VcVersionItem.Value } else { 'n/a' }

    $AllHostFindings = $SiteFindings | Where-Object { $_.Area -eq 'Host' -and $_.Item -eq 'Host Configuration Summary' }
    $HostCountTotal = 0
    $HostVersionsAll = @()
    foreach ($hf in $AllHostFindings) {
        $parts = $hf.Value -split '\|'
        $HostCountTotal += [int]$parts[0]
        $HostVersionsAll += "$($parts[1]) ($($parts[2]))"
    }
    $EsxiVersionDisplay = ($HostVersionsAll | Select-Object -Unique) -join ', '

    $VmCountTotal = 0
    foreach ($g in ($SiteFindings | Where-Object { $_.Area -eq 'VM' -and $_.Item -eq 'Guest OS' })) {
        $VmCountTotal += [int]$g.Value
    }

    $ClusterNames = @(if ($Global:SiteClusterMap.ContainsKey($SiteLabel)) { $Global:SiteClusterMap[$SiteLabel] } else { @() })
    $ClusterCount = $ClusterNames.Count

    # Same "ON unless any cluster explicitly reports False" logic used everywhere else in this
    # script - keeps every DRS/HA read in perfect agreement across the dashboard.
    $drsAll = ($SiteFindings | Where-Object { $_.Item -eq 'vSphere DRS' })
    $haAll  = ($SiteFindings | Where-Object { $_.Item -eq 'vSphere HA' })
    $drsOn = [bool]($drsAll) -and -not ($drsAll | Where-Object { $_.Value -eq 'False' })
    $haOn  = [bool]($haAll)  -and -not ($haAll  | Where-Object { $_.Value -eq 'False' })

    # Capacity - same capacity-weighted average formula used throughout this script.
    $capFindings = $SiteFindings | Where-Object { $_.Area -eq 'Capacity' }
    $cpuCapTotal = 0.0; $cpuUsedTotal = 0.0
    $memCapTotal = 0.0; $memUsedTotal = 0.0
    foreach ($cf in ($capFindings | Where-Object { $_.Item -eq 'CPU' })) {
        $p = $cf.Value -split '\|'; $cap = [double]$p[0]; $cpuCapTotal += $cap
        if ($p[1] -and $p[1] -ne '') { $cpuUsedTotal += ($cap * [double]$p[1] / 100) }
    }
    foreach ($cf in ($capFindings | Where-Object { $_.Item -eq 'Memory' })) {
        $p = $cf.Value -split '\|'; $cap = [double]$p[0]; $memCapTotal += $cap
        if ($p[1] -and $p[1] -ne '') { $memUsedTotal += ($cap * [double]$p[1] / 100) }
    }
    $cpuUsagePct = if ($cpuCapTotal -gt 0) { ($cpuUsedTotal / $cpuCapTotal) * 100 } else { $null }
    $memUsagePct = if ($memCapTotal -gt 0) { ($memUsedTotal / $memCapTotal) * 100 } else { $null }

    # Storage / datastores - grouped by Object (datastore name) so a datastore shared across two
    # clusters in the same site is only counted/listed once, not once per cluster it's mounted to.
    $dsGroups = $SiteFindings | Where-Object { $_.Area -eq 'Storage' -and $_.Item -eq 'Datastore' } | Group-Object Object
    $DatastoresRaw = foreach ($g in $dsGroups) {
        $f = $g.Group | Select-Object -First 1
        $p = $f.Value -split '\|'
        $capGB = [double]$p[0]; $freeGB = [double]$p[1]
        $usedPct = if ($capGB -gt 0) { (($capGB - $freeGB) / $capGB) * 100 } else { 0 }
        [pscustomobject]@{ Name = $f.Object; CapacityGB = $capGB; FreeGB = $freeGB; UsedPct = $usedPct; Status = $f.Status }
    }
    $Datastores = @($DatastoresRaw | Sort-Object Name)
    $dsCapTotal  = [double](($Datastores | Measure-Object -Property CapacityGB -Sum).Sum)
    $dsFreeTotal = [double](($Datastores | Measure-Object -Property FreeGB -Sum).Sum)
    $dsUsagePct = if ($dsCapTotal -gt 0) { (($dsCapTotal - $dsFreeTotal) / $dsCapTotal) * 100 } else { $null }
    $dsInaccessible = [bool]($Datastores | Where-Object { $_.Status -eq 'Critical' })

    $VsanEnabledAny = [bool]($SiteFindings | Where-Object { $_.Item -eq 'vSAN Enabled' })
    $StorageLabelParts = @()
    if ($Datastores.Count -gt 0) { $StorageLabelParts += 'VMFS on SAN' }
    if ($VsanEnabledAny) { $StorageLabelParts += 'vSAN' }
    $StorageLabel = if ($StorageLabelParts.Count -gt 0) { ($StorageLabelParts | Select-Object -Unique) -join ' + ' } else { 'n/a' }

    # Networking - same uplink/VLAN-0 exclusion used throughout this script.
    $PgFindings  = $SiteFindings | Where-Object { $_.Item -eq 'Port Group' }
    $VdsFindings = $SiteFindings | Where-Object { $_.Item -eq 'vDS' }
    $VdsCountVal = @($VdsFindings | Select-Object -ExpandProperty Object -Unique).Count
    $VlanCountVal = @($PgFindings | Where-Object {
        $parts = $_.Value -split '\|'
        $parts[2] -eq 'True' -and $parts[0] -ne '0'
    } | ForEach-Object { ($_.Value -split '\|')[0] } | Select-Object -Unique).Count

    $TeamingValues = @($PgFindings | ForEach-Object { ($_.Value -split '\|')[1] } | Where-Object { $_ } | Select-Object -Unique)
    $NicTeamingPolicy = if ($TeamingValues.Count -gt 0) { $TeamingValues -join ', ' } else { 'n/a' }

    $MtuValues = @($VdsFindings | Select-Object -ExpandProperty Value -Unique)
    $MtuConsistent = $MtuValues.Count -le 1
    $MtuText = if ($MtuValues.Count -eq 1) { "Consistent ($($MtuValues[0]))" }
               elseif ($MtuValues.Count -gt 1) { "Mismatch: $($MtuValues -join ' vs ')" }
               else { 'n/a' }

    $nicRedundancy = @($SiteFindings | Where-Object { $_.Item -eq 'Physical NIC Redundancy' })
    $nicBad = @($nicRedundancy | Where-Object { $_.Status -ne 'Healthy' })
    $NicRedundancyOk = $nicRedundancy.Count -gt 0 -and $nicBad.Count -eq 0
    $NicRedundancyText = if ($nicRedundancy.Count -eq 0) { 'n/a' }
                         elseif ($NicRedundancyOk) { 'Yes, all hosts' }
                         else { "No - $($nicBad.Count) host(s) lacking redundancy" }

    $netStatuses = @($nicRedundancy | Select-Object -ExpandProperty Status) +
                   @($SiteFindings | Where-Object { $_.Item -eq 'NIC Errors/Drops' } | Select-Object -ExpandProperty Status) +
                   @($(if ($MtuConsistent) { 'Healthy' } else { 'Warning' }))
    $NetworkStatus = if ($netStatuses.Count -gt 0) { Get-WorstStatus $netStatuses } else { 'Unable to Check' }

    # Security & compliance - per-host findings rolled up to one site-wide status + short display
    # text per item, in the same spirit as the old report's bullets but sized for a dashboard tile.
    $lockdownRows   = @($SiteFindings | Where-Object { $_.Item -eq 'Lockdown Mode' })
    $secureBootRows = @($SiteFindings | Where-Object { $_.Item -eq 'Secure Boot' })
    $localAcctRows  = @($SiteFindings | Where-Object { $_.Item -eq 'Local Accounts' })
    $syslogRows     = @($SiteFindings | Where-Object { $_.Item -eq 'Syslog' })

    $lockdownWorst = if ($lockdownRows.Count -gt 0) { Get-WorstStatus ($lockdownRows | Select-Object -ExpandProperty Status) } else { 'Unable to Check' }
    $lockdownBad   = @($lockdownRows | Where-Object { $_.Status -ne 'Healthy' }).Count
    $lockdownText  = if ($lockdownRows.Count -eq 0) { 'n/a' } elseif ($lockdownWorst -eq 'Healthy') { 'Enabled' } else { "Disabled on $lockdownBad host(s)" }

    $sbWorst = if ($secureBootRows.Count -gt 0) { Get-WorstStatus ($secureBootRows | Select-Object -ExpandProperty Status) } else { 'Unable to Check' }
    $sbBad   = @($secureBootRows | Where-Object { $_.Status -ne 'Healthy' }).Count
    $sbText  = if ($secureBootRows.Count -eq 0) { 'n/a' } elseif ($sbWorst -eq 'Healthy') { 'Enabled' } else { "Disabled on $sbBad host(s)" }

    $acctWorst = if ($localAcctRows.Count -gt 0) { Get-WorstStatus ($localAcctRows | Select-Object -ExpandProperty Status) } else { 'Unable to Check' }
    $acctBad   = @($localAcctRows | Where-Object { $_.Status -ne 'Healthy' }).Count
    $acctText  = if ($localAcctRows.Count -eq 0) { 'n/a' } elseif ($acctWorst -eq 'Healthy') { 'Compliant' } else { "$acctBad host(s) with unexpected accounts" }

    $syslogWorst = if ($syslogRows.Count -gt 0) { Get-WorstStatus ($syslogRows | Select-Object -ExpandProperty Status) } else { 'Unable to Check' }
    $syslogBad   = @($syslogRows | Where-Object { $_.Status -ne 'Healthy' }).Count
    $syslogText  = if ($syslogRows.Count -eq 0) { 'n/a' } elseif ($syslogWorst -eq 'Healthy') { 'Configured on all hosts' } else { "Not configured on $syslogBad host(s)" }

    $SecurityStatus = Get-WorstStatus @($lockdownWorst, $sbWorst, $acctWorst, $syslogWorst)

    # Backup & DR - whether -BackupInfo was supplied for this site; the detail itself (appliance,
    # schedule, retention, status) comes straight from that parameter, never fabricated. A
    # Warning/Critical backup status counts toward this site's risk totals and Action Plan just
    # like any other finding - a failed/expired backup solution is exactly the kind of thing that
    # belongs in "risk issues", not a fact hidden away in its own panel.
    $backupSupplied = $BackupInfo.ContainsKey($SiteLabel)
    $backupDetail = if ($backupSupplied) { $BackupInfo[$SiteLabel] } else { $null }
    $backupStatus = if ($backupDetail -and $backupDetail.Status) { $backupDetail.Status } else { 'Manual/External Required' }
    if ($backupStatus -eq 'Critical') {
        $CritCount++
        $OverallHealth = 'Critical'
    } elseif ($backupStatus -eq 'Warning' -and $OverallHealth -ne 'Critical') {
        $WarnCount++
        $OverallHealth = 'Warning'
    }

    # Action plan - every Warning/Critical finding for this site, Critical first, with the backup
    # issue (if any) surfaced first since it's typically the most business-impacting item.
    # Appliance and Hardware findings can never appear here since those checks are no longer
    # collected at all.
    $ActionItems = @($SiteFindings | Where-Object { $_.Status -in 'Critical','Warning' } |
        Sort-Object @{Expression = { if ($_.Status -eq 'Critical') { 0 } else { 1 } }} |
        ForEach-Object {
            [pscustomobject]@{
                Severity = if ($_.Status -eq 'Critical') { 'High' } else { 'Medium' }
                Object   = $_.Object
                Item     = $_.Item
                Value    = Get-ActionItemDisplayValue -Item $_.Item -Value $_.Value
                Notes    = $_.Notes
            }
        })
    if ($backupStatus -in 'Critical','Warning') {
        $ActionItems = @([pscustomobject]@{
            Severity = if ($backupStatus -eq 'Critical') { 'High' } else { 'Medium' }
            Object   = $SiteLabel
            Item     = 'Backup & Disaster Recovery'
            Value    = "$($backupDetail.Appliance) - Status: $backupStatus"
            Notes    = $backupDetail.Notes
        }) + $ActionItems
    }

    # Per-cluster breakdown for this site's own dashboard page.
    $clusterRows = @(foreach ($cn in $ClusterNames) {
        $hf = $AllHostFindings | Where-Object { $_.Cluster -eq $cn } | Select-Object -First 1
        $clusterHosts = if ($hf) { [int]($hf.Value -split '\|')[0] } else { 0 }
        $clusterVMs = 0
        foreach ($g in ($SiteFindings | Where-Object { $_.Area -eq 'VM' -and $_.Item -eq 'Guest OS' -and $_.Cluster -eq $cn })) { $clusterVMs += [int]$g.Value }
        $cCpuF = $capFindings | Where-Object { $_.Item -eq 'CPU' -and $_.Cluster -eq $cn } | Select-Object -First 1
        $cMemF = $capFindings | Where-Object { $_.Item -eq 'Memory' -and $_.Cluster -eq $cn } | Select-Object -First 1
        $cCpuPct = if ($cCpuF) { $p = $cCpuF.Value -split '\|'; if ($p[1] -and $p[1] -ne '') { [double]$p[1] } else { $null } } else { $null }
        $cMemPct = if ($cMemF) { $p = $cMemF.Value -split '\|'; if ($p[1] -and $p[1] -ne '') { [double]$p[1] } else { $null } } else { $null }
        [pscustomobject]@{ Cluster = $cn; Hosts = $clusterHosts; VMs = $clusterVMs; CpuPct = $cCpuPct; MemPct = $cMemPct }
    })

    # VM OS distribution, most common guest OS first.
    $osGroups = $SiteFindings | Where-Object { $_.Area -eq 'VM' -and $_.Item -eq 'Guest OS' } | Group-Object Object
    $OsDistribution = @($osGroups | ForEach-Object {
        [pscustomobject]@{ Name = $_.Name; Count = ($_.Group | Measure-Object -Property Value -Sum).Sum }
    } | Sort-Object Count -Descending)

    $SummaryText = if ($CritCount -gt 0) {
        "The VMware environment consisting of 1 vCenter Server and $HostCountTotal ESXi hosts has $CritCount critical issue(s) identified during this health check that require prompt attention."
    } else {
        "The VMware environment consisting of 1 vCenter Server and $HostCountTotal ESXi hosts is in a healthy, stable, and fully supported state. No Critical issues were identified during this health check. The platform operates efficiently and is ready to support current and future workloads."
    }

    [pscustomobject]@{
        Site              = $SiteLabel
        VCenterVersion    = $VCenterVersionText
        OverallHealth     = $OverallHealth
        HighRisk          = $CritCount
        MediumRisk        = $WarnCount
        LowRisk           = $UnableCount
        HostCount         = $HostCountTotal
        VmCount           = $VmCountTotal
        ClusterCount      = $ClusterCount
        ClusterRows       = $clusterRows
        EsxiVersion       = $(if ($EsxiVersionDisplay) { $EsxiVersionDisplay } else { 'n/a' })
        StorageLabel      = $StorageLabel
        DrsOn             = $drsOn
        HaOn              = $haOn
        CpuPct            = $cpuUsagePct
        CpuStatus         = $(if ($cpuUsagePct -ne $null) { Get-PctStatus $cpuUsagePct } else { 'Unable to Check' })
        CpuCapGHz         = $cpuCapTotal / 1000
        CpuUsedGHz        = $cpuUsedTotal / 1000
        CpuFreeGHz        = ($cpuCapTotal - $cpuUsedTotal) / 1000
        MemPct            = $memUsagePct
        MemStatus         = $(if ($memUsagePct -ne $null) { Get-PctStatus $memUsagePct } else { 'Unable to Check' })
        MemCapGB          = $memCapTotal / 1024
        MemUsedGB         = $memUsedTotal / 1024
        MemFreeGB         = ($memCapTotal - $memUsedTotal) / 1024
        StoragePct        = $dsUsagePct
        StorageStatus     = $(if ($dsInaccessible) { 'Critical' } elseif ($dsUsagePct -ne $null) { Get-PctStatus $dsUsagePct } else { 'Unable to Check' })
        StorageCapGB      = $dsCapTotal
        StorageUsedGB     = $dsCapTotal - $dsFreeTotal
        StorageFreeGB     = $dsFreeTotal
        Datastores        = $Datastores
        VdsCount          = $VdsCountVal
        VlanCount         = $VlanCountVal
        NicTeamingPolicy  = $NicTeamingPolicy
        MtuConsistent     = $MtuConsistent
        MtuText           = $MtuText
        NicRedundancyOk   = $NicRedundancyOk
        NicRedundancyText = $NicRedundancyText
        NetworkStatus     = $NetworkStatus
        LockdownStatus    = $lockdownWorst
        LockdownText      = $lockdownText
        SecureBootStatus  = $sbWorst
        SecureBootText    = $sbText
        LocalAcctStatus   = $acctWorst
        LocalAcctText     = $acctText
        SyslogStatus      = $syslogWorst
        SyslogText        = $syslogText
        SecurityStatus    = $SecurityStatus
        BackupSupplied    = $backupSupplied
        BackupDetail      = $backupDetail
        ActionItems       = $ActionItems
        OsDistribution    = $OsDistribution
        SummaryText       = $SummaryText
    }
}

# Writes the combined multi-site HTML dashboard - one static, self-contained file with no
# external dependencies (no CDN/internet access assumed), so it opens correctly straight from disk
# on an offline/internal machine. This is the ONLY report artifact the script produces now.
function ConvertTo-Slug {
    param([string]$Text)
    $slug = ($Text -replace '[^a-zA-Z0-9]+', '-').Trim('-').ToLower()
    if (-not $slug) { return 'site' }
    return $slug
}

function Write-DashboardHtml {
    param([string[]]$SiteLabels, [string]$OutputPath, [string]$RunDateDisplay, [string]$ScriptBuild)

    # Sites listed in -DashboardSiteOrder appear in that exact order; any connected site NOT in
    # the list is appended afterward, alphabetically, so a new/unlisted site still shows up rather
    # than being silently dropped from the dashboard.
    $orderIndex = @{}
    for ($i = 0; $i -lt $DashboardSiteOrder.Count; $i++) { $orderIndex[$DashboardSiteOrder[$i]] = $i }
    $summaries = @($SiteLabels | ForEach-Object { Get-SiteDashboardSummary -SiteLabel $_ } | Where-Object { $_ } |
        Sort-Object @{Expression = { if ($orderIndex.ContainsKey($_.Site)) { $orderIndex[$_.Site] } else { 9999 } }}, Site)
    if ($summaries.Count -eq 0) { return }

    $healthColor = @{ Healthy = '#2e7d32'; Warning = '#e6a100'; Critical = '#c62828' }
    $healthLabelText = @{ Healthy = 'Healthy - No Issues Detected'; Warning = 'Healthy - Minor Issues Detected'; Critical = 'Attention Required - Critical Issues' }
    # "Healthy Sites" counts every site that isn't Critical - a site with only Warning-level
    # findings still shows its own badge as "Healthy - Minor Issues Detected", so it belongs here
    # too, not just the (in practice almost never reached) zero-findings case. "Sites with
    # Warnings" is a narrower, informational subset of that same count - how many of the healthy
    # sites still have at least one Warning worth a look.
    $healthCounts = @{
        Healthy  = @($summaries | Where-Object { $_.OverallHealth -ne 'Critical' }).Count
        Warning  = @($summaries | Where-Object { $_.OverallHealth -eq 'Warning' }).Count
        Critical = @($summaries | Where-Object { $_.OverallHealth -eq 'Critical' }).Count
    }
    $totalHigh   = ($summaries | Measure-Object -Property HighRisk -Sum).Sum
    $totalMedium = ($summaries | Measure-Object -Property MediumRisk -Sum).Sum
    $totalHosts  = ($summaries | Measure-Object -Property HostCount -Sum).Sum
    $totalVMs    = ($summaries | Measure-Object -Property VmCount -Sum).Sum

    # Distinct heavy tab color per site, cycling through this palette by position - never the
    # green/amber/red health colors, which already mean something specific elsewhere (risk badges,
    # KPI tiles). The Overview tab keeps its own fixed navy.
    $tabPalette = @('#2563EB','#7C3AED','#0D9488','#C026D3','#EA580C','#4F46E5','#DB2777','#0EA5E9')

    # Fixed decorative palette for the 5 non-status tiles on every site page (ESXi Hosts, Clusters,
    # VMs, ESXi Version, Storage Type), plus dedicated colors for the DRS/HA status tiles - all 7
    # distinct on every single page, never white, never a duplicate within the same page.
    $tileBlue = '#1565C0'; $tilePurple = '#6A1B9A'; $tileTeal = '#00897B'; $tileIndigo = '#283593'; $tileBrown = '#6D4C41'
    $tileDrs = '#37474F'; $tileHa = '#880E4F'

    function Get-PctBarColor {
        param([Nullable[double]]$Pct)
        if ($Pct -eq $null) { return '#888' }
        switch (Get-PctStatus $Pct) {
            'Critical' { '#c62828' }
            'Warning'  { '#e6a100' }
            default    { '#2e7d32' }
        }
    }

    function Get-BarHtml {
        param([Nullable[double]]$Pct, [string]$Status, [string]$Label, [string]$ValueText, [string]$DetailHtml = '')
        $pctColorLocal = @{ Healthy = '#2e7d32'; Warning = '#e6a100'; Critical = '#c62828'; 'Unable to Check' = '#888' }
        $color = $pctColorLocal[$Status]
        $width = if ($Pct -ne $null) { [Math]::Min(100, [Math]::Max(0, $Pct)) } else { 0 }
@"
        <div class="bar-row">
          <div class="bar-label"><span>$(ConvertTo-HtmlSafe $Label)</span><span style="color:$color; font-weight:bold">$(ConvertTo-HtmlSafe $ValueText)</span></div>
          <div class="bar-track"><div class="bar-fill" style="width:$width%;background:$color"></div></div>
          $DetailHtml
        </div>
"@
    }

    function Get-StatusPillHtml {
        param([string]$Status, [string]$Text)
        $variant = switch ($Status) {
            'Healthy' { 'ok' }
            'Warning' { 'warn' }
            'Critical' { 'bad' }
            default { 'info' }
        }
        "<span class=`"status-pill $variant`"><span class=`"dot`"></span>$(ConvertTo-HtmlSafe $Text)</span>"
    }

    function Get-MiniStatVariant {
        param([string]$Status)
        switch ($Status) {
            'Healthy' { 'ok' }
            'Warning' { 'warn' }
            'Critical' { 'bad' }
            default { '' }
        }
    }

    # --- Overview page: KPI stat row + one card per site --------------------------------------
    $overviewCards = ($summaries | ForEach-Object {
        $s = $_
        $slug = ConvertTo-Slug $s.Site
        $cpuText = if ($s.CpuPct -ne $null) { "{0:N1}%" -f $s.CpuPct } else { 'n/a' }
        $memText = if ($s.MemPct -ne $null) { "{0:N1}%" -f $s.MemPct } else { 'n/a' }
        $stgText = if ($s.StoragePct -ne $null) { "{0:N1}%" -f $s.StoragePct } else { 'n/a' }
        $color = $healthColor[$s.OverallHealth]
        $complianceText = if ($s.SecurityStatus -eq 'Healthy') { 'Compliant' } elseif ($s.SecurityStatus -in 'Unable to Check','Manual/External Required') { 'n/a' } else { 'Non-Compliant' }
        $complianceColor = if ($s.SecurityStatus -eq 'Healthy') { '#2e7d32' } elseif ($s.SecurityStatus -in 'Unable to Check','Manual/External Required') { '#888' } else { '#c62828' }
@"
      <div class="ov-card" onclick="showPage('$slug')" style="border-top-color:$color">
        <div class="ov-head"><h2>$(ConvertTo-HtmlSafe $s.Site)</h2><span class="badge" style="background:$color">$(ConvertTo-HtmlSafe $healthLabelText[$s.OverallHealth])</span></div>
        <div class="risk-row"><span class="risk risk-high">High: $($s.HighRisk)</span><span class="risk risk-med">Medium: $($s.MediumRisk)</span><span class="risk risk-low">Low: $($s.LowRisk)</span></div>
        <table class="metrics">
          <tr><td>ESXi Hosts / Clusters / VMs</td><td>$($s.HostCount) / $($s.ClusterCount) / $($s.VmCount)</td></tr>
          <tr><td>vCenter / ESXi Version</td><td>$(ConvertTo-HtmlSafe $s.VCenterVersion) / $(ConvertTo-HtmlSafe $s.EsxiVersion)</td></tr>
          <tr><td>Storage Type</td><td>$(ConvertTo-HtmlSafe $s.StorageLabel)</td></tr>
          <tr><td>vSphere DRS / HA</td><td>$(if ($s.DrsOn) {'ON'} else {'<span style="color:#c62828">OFF</span>'}) / $(if ($s.HaOn) {'ON'} else {'<span style="color:#c62828">OFF</span>'})</td></tr>
          <tr><td>CPU / Mem / Storage</td><td>$cpuText / $memText / $stgText</td></tr>
          <tr><td>Backup Configured</td><td>$(if ($s.BackupSupplied) {'Yes'} else {'No'})</td></tr>
          <tr><td>Compliance</td><td style="color:$complianceColor">$complianceText</td></tr>
        </table>
        <span class="ov-link">View Full Details &rarr;</span>
      </div>
"@
    }) -join "`n"

    # --- Tab navigation bar (circular, full site name inside each circle) ---
    $tabButtons = (@('<button class="tab overview-tab active" id="tab-overview" onclick="showPage(''overview'')"><span class="tab-circle">Overview</span></button>') + ($summaries | ForEach-Object {
        $i = $summaries.IndexOf($_)
        $slug = ConvertTo-Slug $_.Site
        $tabColor = $tabPalette[$i % $tabPalette.Count]
        "<button class=`"tab`" id=`"tab-$slug`" onclick=`"showPage('$slug')`"><span class=`"tab-circle`" style=`"background:$tabColor`">$(ConvertTo-HtmlSafe $_.Site)</span></button>"
    })) -join "`n    "

    # --- One full, panel-rich page per site ---
    $sitePages = ($summaries | ForEach-Object {
        $s = $_
        $slug = ConvertTo-Slug $s.Site
        $color = $healthColor[$s.OverallHealth]

        $clusterTableRows = if ($s.ClusterRows.Count -gt 0) {
            ($s.ClusterRows | ForEach-Object {
                $cr = $_
                $crCpuVal = if ($cr.CpuPct -ne $null) { $cr.CpuPct } else { 0 }
                $crMemVal = if ($cr.MemPct -ne $null) { $cr.MemPct } else { 0 }
                $crCpuText = if ($cr.CpuPct -ne $null) { "{0:N1}%" -f $cr.CpuPct } else { 'n/a' }
                $crMemText = if ($cr.MemPct -ne $null) { "{0:N1}%" -f $cr.MemPct } else { 'n/a' }
@"
          <tr><td>$(ConvertTo-HtmlSafe $cr.Cluster)</td><td>$($cr.Hosts)</td><td>$($cr.VMs)</td>
            <td><div class="mini-bar-wrap"><div class="mini-bar-track"><div class="mini-bar-fill" style="width:$crCpuVal%;background:$(Get-PctBarColor $cr.CpuPct)"></div></div><span>$crCpuText</span></div></td>
            <td><div class="mini-bar-wrap"><div class="mini-bar-track"><div class="mini-bar-fill" style="width:$crMemVal%;background:$(Get-PctBarColor $cr.MemPct)"></div></div><span>$crMemText</span></div></td>
          </tr>
"@
            }) -join "`n"
        } else { "<tr><td colspan='5' style='color:#999'>No cluster detail available</td></tr>" }

        $dsRows = if ($s.Datastores.Count -gt 0) {
            ($s.Datastores | ForEach-Object {
                $ds = $_
@"
              <tr><td>$(ConvertTo-HtmlSafe $ds.Name)</td><td>$("{0:N0}" -f $ds.CapacityGB) GB</td><td>$("{0:N2}" -f $ds.FreeGB) GB</td>
                <td><div class="mini-bar-wrap"><div class="mini-bar-track"><div class="mini-bar-fill" style="width:$($ds.UsedPct)%;background:$(Get-PctBarColor $ds.UsedPct)"></div></div><span>$("{0:N1}" -f $ds.UsedPct)%</span></div></td>
                <td>$(Get-StatusPillHtml -Status $ds.Status -Text $ds.Status)</td></tr>
"@
            }) -join "`n"
        } else { "<tr><td colspan='5' style='color:#999'>No datastores found</td></tr>" }

        $osRows = if ($s.OsDistribution.Count -gt 0) {
            $maxCount = ($s.OsDistribution | Select-Object -First 1).Count
            ($s.OsDistribution | ForEach-Object {
                $pctWidth = if ($maxCount -gt 0) { [Math]::Max(3, ($_.Count / $maxCount) * 100) } else { 3 }
                "<div class=`"os-bar-row`"><div class=`"os-name`">$(ConvertTo-HtmlSafe $_.Name)</div><div class=`"os-track`"><div class=`"os-fill`" style=`"width:$pctWidth%`"></div></div><div class=`"os-count`">$($_.Count)</div></div>"
            }) -join "`n"
        } else { "<p style='color:#999'>No VM inventory found</p>" }

        $actionRows = if ($s.ActionItems.Count -gt 0) {
            ($s.ActionItems | ForEach-Object {
                $sevClass = if ($_.Severity -eq 'High') { 'high' } else { 'med' }
                $subtitle = if ($_.Notes) { "$($_.Object) - $($_.Notes)" } else { $_.Object }
@"
        <li><span class="sev $sevClass">$($_.Severity)</span><div class="txt"><strong>$(ConvertTo-HtmlSafe $_.Item): $(ConvertTo-HtmlSafe $_.Value)</strong><span>$(ConvertTo-HtmlSafe $subtitle)</span></div></li>
"@
            }) -join "`n"
        } else { $null }

        $actionPlanBody = if ($actionRows) {
            "<ul class=`"action-list`">`n$actionRows`n</ul>"
        } else {
@"
      <div class="center-callout">
        $(Get-StatusPillHtml -Status 'Healthy' -Text 'No Warning or Critical items flagged for this site')
      </div>
"@
        }

        $backupBody = if ($s.BackupSupplied -and $s.BackupDetail.Status -in 'Healthy','Warning','Critical') {
            $bi = $s.BackupDetail
            $biStatus = $bi.Status
            $statusText = switch ($biStatus) {
                'Healthy'  { 'Successfully Completed' }
                'Critical' { if ($bi.Notes) { $bi.Notes } else { 'Failed' } }
                'Warning'  { if ($bi.Notes) { $bi.Notes } else { 'Attention Required' } }
            }
            $flagClass = switch ($biStatus) { 'Critical' { 'critical' }; 'Warning' { 'warn' }; default { 'healthy' } }
@"
      <table class="metrics">
        <tr><td>Appliance</td><td>$(ConvertTo-HtmlSafe $bi.Appliance) ($(ConvertTo-HtmlSafe $biStatus))</td></tr>
        <tr><td>Schedule</td><td>$(ConvertTo-HtmlSafe $bi.Schedule)</td></tr>
        <tr><td>Retention</td><td style="white-space:nowrap">$(ConvertTo-HtmlSafe $bi.Retention)</td></tr>
        <tr><td>Status</td><td>$(ConvertTo-HtmlSafe $statusText)</td></tr>
      </table>
      <div class="backup-flag $flagClass">$(ConvertTo-HtmlSafe $biStatus)</div>
"@
        } elseif ($s.BackupSupplied) {
            # Explicitly supplied but with a known reason it's not yet tracked as Healthy/Warning/
            # Critical (e.g. a newly handed-over site whose backup management isn't ours yet) -
            # distinct from "no data supplied at all" below, since there IS a specific reason to show.
            $bi = $s.BackupDetail
            $reasonText = if ($bi.Notes) { $bi.Notes } else { 'Manual/External Required' }
@"
      <div class="center-callout">
        $(Get-StatusPillHtml -Status 'Manual/External Required' -Text 'Manual/External Required')
        <p style="color:#999;font-size:13px;margin:12px 0 0">$(ConvertTo-HtmlSafe $reasonText)</p>
      </div>
      <div class="backup-flag warn">Pending</div>
"@
        } else {
@"
      <div class="center-callout">
        $(Get-StatusPillHtml -Status 'Manual/External Required' -Text 'Manual/External Required')
        <p style="color:#999;font-size:13px;margin:12px 0 0">Not supplied for this run - pass -BackupInfo to include backup device/solution status here.</p>
      </div>
"@
        }

        $mtuVariant = if ($s.MtuConsistent) { 'ok' } else { 'warn' }
        $mtuIcon = if ($s.MtuConsistent) { [char]0x2705 } else { [char]0x26A0 }
        $redundancyVariant = if ($s.NicRedundancyOk) { 'ok' } else { 'bad' }

        $cpuValueText = if ($s.CpuPct -ne $null) { "{0:N1}%" -f $s.CpuPct } else { 'n/a' }
        $memValueText = if ($s.MemPct -ne $null) { "{0:N1}%" -f $s.MemPct } else { 'n/a' }
        $stgValueText = if ($s.StoragePct -ne $null) { "{0:N1}%" -f $s.StoragePct } else { 'n/a' }
        $cpuDetail = "<div class=`"cap-detail`"><span><b>$("{0:N2}" -f $s.CpuCapGHz)</b> GHz total</span><span><b>$("{0:N2}" -f $s.CpuUsedGHz)</b> GHz used</span><span><b>$("{0:N2}" -f $s.CpuFreeGHz)</b> GHz free</span></div>"
        $memDetail = "<div class=`"cap-detail`"><span><b>$("{0:N2}" -f $s.MemCapGB)</b> GB total</span><span><b>$("{0:N2}" -f $s.MemUsedGB)</b> GB used</span><span><b>$("{0:N2}" -f $s.MemFreeGB)</b> GB free</span></div>"
        $stgDetail = "<div class=`"cap-detail`"><span><b>$("{0:N1}" -f $s.StorageCapGB)</b> GB total</span><span><b>$("{0:N2}" -f $s.StorageUsedGB)</b> GB used</span><span><b>$("{0:N2}" -f $s.StorageFreeGB)</b> GB free</span></div>"

@"
      <section class="page" id="page-$slug">
        <div class="site-hero" style="border-left-color:$color">
          <h1>$(ConvertTo-HtmlSafe $s.Site)</h1>
          <span class="badge big" style="background:$color">$(ConvertTo-HtmlSafe $healthLabelText[$s.OverallHealth])</span>
          <div class="meta">$(ConvertTo-HtmlSafe $s.VCenterVersion) &middot; Report generated $(ConvertTo-HtmlSafe $RunDateDisplay)</div>
        </div>

        <div class="risk-row big">
          <span class="risk risk-high">High Risk: $($s.HighRisk)</span>
          <span class="risk risk-med">Medium Risk: $($s.MediumRisk)</span>
          <span class="risk risk-low">Low Risk: $($s.LowRisk)</span>
        </div>

        <div class="tile-row">
          <div class="tile" style="background:$tileBlue"><span class="num">$($s.HostCount)</span><span class="label">ESXi Hosts</span></div>
          <div class="tile" style="background:$tilePurple"><span class="num">$($s.ClusterCount)</span><span class="label">Clusters</span></div>
          <div class="tile" style="background:$tileTeal"><span class="num">$($s.VmCount)</span><span class="label">Virtual Machines</span></div>
          <div class="tile" style="background:$tileIndigo"><span class="num" style="font-size:16px">$(ConvertTo-HtmlSafe $s.EsxiVersion)</span><span class="label">ESXi Version</span></div>
          <div class="tile" style="background:$tileDrs"><span class="num $(if ($s.DrsOn) {'on'} else {'off'})">$(if ($s.DrsOn) {'ON'} else {'OFF'})</span><span class="label">vSphere DRS</span></div>
          <div class="tile" style="background:$tileHa"><span class="num $(if ($s.HaOn) {'on'} else {'off'})">$(if ($s.HaOn) {'ON'} else {'OFF'})</span><span class="label">vSphere HA</span></div>
          <div class="tile" style="background:$tileBrown"><span class="num" style="font-size:18px">$(ConvertTo-HtmlSafe $s.StorageLabel)</span><span class="label">Storage Type</span></div>
        </div>

        <div class="panel-grid">
          <div class="panel" style="border-top-color:#1565C0">
            <h3><span class="n" style="background:#1565C0">&#9889;</span>Performance &amp; Capacity (Last $PerfHistoryHours Hours)</h3>
$(Get-BarHtml -Pct $s.CpuPct -Status $s.CpuStatus -Label 'CPU Usage' -ValueText $cpuValueText -DetailHtml $cpuDetail)
$(Get-BarHtml -Pct $s.MemPct -Status $s.MemStatus -Label 'Memory Usage' -ValueText $memValueText -DetailHtml $memDetail)
$(Get-BarHtml -Pct $s.StoragePct -Status $s.StorageStatus -Label 'Storage Usage' -ValueText $stgValueText -DetailHtml $stgDetail)
          </div>

          <div class="panel" style="border-top-color:#6A1B9A">
            <h3><span class="n" style="background:#6A1B9A">&#127760;</span>Networking Health</h3>
            <div class="mini-grid">
              <div class="mini-stat"><span class="mi-icon">&#128256;</span><span class="mi-val">$($s.VdsCount)</span><span class="mi-label">Distributed Switches</span></div>
              <div class="mini-stat"><span class="mi-icon">&#127991;&#65039;</span><span class="mi-val">$($s.VlanCount)</span><span class="mi-label">VLANs Configured</span></div>
              <div class="mini-stat"><span class="mi-icon">&#9878;&#65039;</span><span class="mi-val" style="font-size:13px">$(ConvertTo-HtmlSafe $s.NicTeamingPolicy)</span><span class="mi-label">NIC Teaming Policy</span></div>
              <div class="mini-stat $mtuVariant"><span class="mi-icon">$mtuIcon</span><span class="mi-val" style="font-size:13px">$(ConvertTo-HtmlSafe $s.MtuText)</span><span class="mi-label">MTU Consistency</span></div>
              <div class="mini-stat $redundancyVariant"><span class="mi-icon">&#128268;</span><span class="mi-val" style="font-size:14px">$(ConvertTo-HtmlSafe $s.NicRedundancyText)</span><span class="mi-label">Redundant NICs</span></div>
            </div>
          </div>

          <div class="panel panel-full" style="border-top-color:#00897B">
            <h3><span class="n" style="background:#00897B">&#128451;&#65039;</span>Storage / Datastores <span style="font-weight:normal;font-size:14px;color:#999">($($s.Datastores.Count) total)</span></h3>
            <div class="scroll-box">
              <table class="metrics wide">
                <thead><tr><th>Datastore</th><th>Capacity</th><th>Free</th><th>Used %</th><th>Status</th></tr></thead>
                <tbody>
$dsRows
                </tbody>
              </table>
            </div>
          </div>

          <div class="panel panel-full" style="border-top-color:#283593">
            <h3><span class="n" style="background:#283593">&#128421;&#65039;</span>VM Inventory - OS Distribution <span style="font-weight:normal;font-size:14px;color:#999">($($s.VmCount) VMs across $($s.OsDistribution.Count) OS types)</span></h3>
            <div class="scroll-box" style="padding:14px 18px">
$osRows
            </div>
          </div>

          <div class="panel" style="border-top-color:#2e7d32">
            <h3><span class="n" style="background:#2e7d32">&#128737;&#65039;</span>Security &amp; Compliance</h3>
            <div class="mini-grid">
              <div class="mini-stat $(Get-MiniStatVariant $s.LockdownStatus)"><span class="mi-icon">&#128274;</span><span class="mi-val" style="font-size:14px">$(ConvertTo-HtmlSafe $s.LockdownText)</span><span class="mi-label">Lockdown Mode</span></div>
              <div class="mini-stat $(Get-MiniStatVariant $s.SecureBootStatus)"><span class="mi-icon">&#128737;&#65039;</span><span class="mi-val" style="font-size:14px">$(ConvertTo-HtmlSafe $s.SecureBootText)</span><span class="mi-label">Secure Boot</span></div>
              <div class="mini-stat $(Get-MiniStatVariant $s.LocalAcctStatus)"><span class="mi-icon">&#128100;</span><span class="mi-val" style="font-size:13px">$(ConvertTo-HtmlSafe $s.LocalAcctText)</span><span class="mi-label">Local ESXi Users</span></div>
              <div class="mini-stat $(Get-MiniStatVariant $s.SyslogStatus)"><span class="mi-icon">&#128221;</span><span class="mi-val" style="font-size:13px">$(ConvertTo-HtmlSafe $s.SyslogText)</span><span class="mi-label">Syslog Configured</span></div>
            </div>
          </div>

          <div class="panel" style="border-top-color:#6D4C41">
            <h3><span class="n" style="background:#6D4C41">&#9729;&#65039;</span>Backup &amp; Disaster Recovery</h3>
$backupBody
          </div>
        </div>

        <div class="panel" style="border-top-color:#e6a100">
          <h3><span class="n" style="background:#e6a100">&#9888;&#65039;</span>Action Plan - Items Requiring Attention</h3>
$actionPlanBody
          <ul class="standard-notes">
            <li>Continue regular monitoring of vCenter Server, ESXi hosts, storage, and network components.</li>
            <li>Perform standard maintenance activities in line with VMware best practices and change management procedures.</li>
            <li>Periodically review capacity utilization (CPU, memory, storage) to support growth planning.</li>
          </ul>
        </div>

        <div class="panel" style="border-top-color:#1E3A5F">
          <h3><span class="n" style="background:#1E3A5F">&#128203;</span>Summary</h3>
          <p class="summary-text">$(ConvertTo-HtmlSafe $s.SummaryText)</p>
        </div>

        <div class="panel" style="border-top-color:#AD1457">
          <h3><span class="n" style="background:#AD1457">&#129513;</span>Clusters</h3>
          <table class="metrics wide">
            <thead><tr><th>Cluster</th><th>Hosts</th><th>VMs</th><th>CPU %</th><th>Memory %</th></tr></thead>
            <tbody>
$clusterTableRows
            </tbody>
          </table>
        </div>
      </section>
"@
    }) -join "`n"

    $html = @"
<!DOCTYPE html>
<html lang="en">
<head>
<meta charset="UTF-8">
<title>VMware Weekly Health Check - Dashboard</title>
<style>
  body { font-family: Calibri, Arial, sans-serif; background:#eef1f5; color:#1a1a1a; margin:0; padding:0 32px 32px; font-size:16px; line-height:1.4; }
  h1 { color:#1E3A5F; margin:0; font-size:36px; }
  .subtitle { color:#555; margin:6px 0 20px; font-size:16px; }
  .tabs { display:grid; grid-template-columns: repeat(auto-fit, minmax(170px, 1fr)); gap:20px; padding:24px 0 20px; position:sticky; top:0; background:#eef1f5; z-index:10; border-bottom:1px solid #dfe3e8; margin-bottom:32px; }
  .tab { display:flex; flex-direction:column; align-items:center; cursor:pointer; background:none; border:none; padding:0; font-family:inherit; }
  .tab .tab-circle { width:150px; height:150px; border-radius:50%; display:flex; align-items:center; justify-content:center; color:#fff; font-weight:bold; font-size:17px; text-align:center; line-height:1.25; padding:10px; box-sizing:border-box; box-shadow:0 1px 4px rgba(0,0,0,0.2); border:3px solid transparent; }
  .tab.overview-tab .tab-circle { background:#1E3A5F; }
  .tab.active .tab-circle { border-color:#1a1a1a; box-shadow:0 0 0 3px rgba(0,0,0,0.15), 0 1px 4px rgba(0,0,0,0.2); }
  .page { display:none; }
  .page.active { display:block; }

  .stat-row { display:grid; grid-template-columns: repeat(auto-fit, minmax(170px, 1fr)); gap:20px; margin-bottom:28px; }
  .stat { border-radius:10px; padding:18px 20px; text-align:center; box-shadow:0 1px 4px rgba(0,0,0,0.14); }
  .stat .num { font-size:32px; font-weight:bold; display:block; color:#fff; }
  .stat .label { font-size:14px; margin-top:4px; display:block; color:#fff; }

  .ov-grid { display:grid; grid-template-columns: repeat(auto-fit, minmax(420px, 1fr)); gap:24px; }
  .ov-card { background:#fff; border-radius:10px; padding:24px 26px; box-shadow:0 1px 4px rgba(0,0,0,0.14); border-top:6px solid; cursor:pointer; }
  .ov-card:hover { box-shadow:0 4px 14px rgba(0,0,0,0.18); }
  .ov-head { display:flex; flex-direction:column; align-items:flex-start; gap:10px; margin-bottom:14px; }
  .ov-head h2 { margin:0; font-size:24px; color:#1E3A5F; }
  .ov-link { display:inline-block; margin-top:14px; color:#1E3A5F; font-weight:bold; font-size:14px; }
  .badge { color:#fff; padding:6px 14px; border-radius:6px; font-size:14px; font-weight:bold; white-space:nowrap; display:inline-block; text-align:center; width:280px; }
  .badge.big { font-size:20px; padding:10px 22px; width:420px; }
  .risk-row { display:flex; gap:12px; margin-bottom:16px; flex-wrap:wrap; }
  .risk-row.big { margin:20px 0 28px; }
  .risk-row.big .risk { font-size:18px; padding:10px 20px; }
  .risk { font-size:14px; padding:5px 12px; border-radius:6px; font-weight:bold; }
  .risk-high { background:#fdecea; color:#c62828; }
  .risk-med  { background:#fff6e0; color:#8a6100; }
  .risk-low  { background:#eef2f5; color:#555; }
  .site-hero { border-left:8px solid; padding:14px 0 14px 26px; margin-bottom:8px; }
  .site-hero h1 { font-size:40px; }
  .site-hero .meta { color:#888; font-size:14px; margin-top:10px; }

  .tile-row { display:grid; grid-template-columns: repeat(auto-fit, minmax(170px, 1fr)); gap:20px; margin-bottom:28px; }
  .tile { border-radius:10px; padding:20px 18px; text-align:center; box-shadow:0 1px 4px rgba(0,0,0,0.14); }
  .tile .num { font-size:24px; font-weight:bold; display:block; color:#fff; }
  .tile .label { font-size:14px; margin-top:6px; display:block; color:#fff; }
  .tile .num.on { color:#69F0AE; }
  .tile .num.off { color:#FF8A80; }

  .panel-grid { display:grid; grid-template-columns: repeat(auto-fit, minmax(380px, 1fr)); gap:24px; margin-bottom:24px; }
  .panel { background:#fff; border-radius:10px; padding:26px 28px; box-shadow:0 1px 4px rgba(0,0,0,0.14); margin-bottom:24px; border-top:5px solid #ccc; }
  .status-pill { display:inline-flex; align-items:center; gap:6px; padding:5px 12px; border-radius:14px; font-size:13px; font-weight:bold; }
  .status-pill.ok { background:#e8f5e9; color:#1b5e20; }
  .status-pill.warn { background:#fff6e0; color:#8a6100; }
  .status-pill.bad { background:#fdecea; color:#c62828; }
  .status-pill.info { background:#eef2f5; color:#555; }
  .status-pill .dot { width:8px; height:8px; border-radius:50%; background:currentColor; }
  .panel h3 { margin:0 0 18px; color:#1E3A5F; font-size:19px; display:flex; align-items:center; gap:10px; }
  .panel h3 .n { color:#fff; width:34px; height:34px; border-radius:50%; display:inline-flex; align-items:center; justify-content:center; font-size:17px; flex-shrink:0; box-shadow:0 1px 3px rgba(0,0,0,0.25); }
  .bar-row { margin-bottom:14px; }
  .bar-label { display:flex; justify-content:space-between; font-size:15px; color:#555; margin-bottom:6px; }
  .bar-track { background:#eef1f4; border-radius:7px; height:14px; overflow:hidden; }
  .bar-fill { height:100%; border-radius:7px; }
  .cap-detail { display:flex; gap:18px; font-size:13px; color:#777; margin:6px 0 18px; flex-wrap:wrap; }
  .cap-detail b { color:#444; }

  .mini-grid { display:grid; grid-template-columns: repeat(auto-fit, minmax(120px, 1fr)); gap:12px; }
  .mini-stat { background:#f4f7fa; border-radius:8px; padding:16px 10px; text-align:center; }
  .mini-stat .mi-icon { font-size:24px; display:block; }
  .mini-stat .mi-val { font-weight:bold; font-size:17px; margin:8px 0 2px; color:#1a1a1a; display:block; }
  .mini-stat .mi-label { font-size:12px; color:#777; display:block; }
  .mini-stat.ok { background:#e8f5e9; } .mini-stat.ok .mi-val { color:#1b5e20; }
  .mini-stat.warn { background:#fff6e0; } .mini-stat.warn .mi-val { color:#8a6100; }
  .mini-stat.bad { background:#fdecea; } .mini-stat.bad .mi-val { color:#c62828; }

  table.metrics { width:100%; border-collapse:collapse; font-size:16px; }
  table.metrics td, table.metrics th { padding:10px 6px; border-bottom:1px solid #eee; vertical-align:top; }
  table.metrics td:first-child { color:#666; width:55%; }
  table.metrics td:last-child:not(:first-child) { text-align:right; font-weight:600; }
  table.metrics.wide th { text-align:center; color:#999; font-size:13px; text-transform:uppercase; padding-bottom:10px; }
  table.metrics.wide th:first-child { text-align:left; }
  table.metrics.wide td { text-align:center; }
  table.metrics.wide td:first-child { text-align:left; color:#333; font-weight:600; width:auto; }
  table.metrics.wide td:last-child:not(:first-child) { text-align:center; font-weight:normal; }
  .scroll-box { max-height:340px; overflow-y:auto; border:1px solid #f0f0f0; border-radius:6px; }
  .scroll-box table.metrics.wide { font-size:15px; }
  .scroll-box thead th { position:sticky; top:0; background:#fff; }
  .panel-full { grid-column: 1 / -1; }
  .mini-bar-wrap { display:flex; align-items:center; gap:8px; justify-content:center; }
  .mini-bar-track { width:60px; height:8px; background:#eef1f4; border-radius:4px; overflow:hidden; flex-shrink:0; }
  .mini-bar-fill { height:100%; border-radius:4px; }
  .os-bar-row { display:flex; align-items:center; gap:10px; margin-bottom:8px; }
  .os-name { width:230px; font-size:14px; color:#444; flex-shrink:0; }
  .os-track { flex:1; background:#eef1f4; border-radius:6px; height:14px; overflow:hidden; }
  .os-fill { height:100%; background:#1565C0; border-radius:6px; }
  .os-count { width:36px; text-align:right; font-weight:600; font-size:14px; }
  .action-list { list-style:none; margin:0; padding:0; }
  .action-list li { display:flex; gap:14px; padding:14px 0; border-bottom:1px solid #eee; align-items:flex-start; }
  .action-list li:last-child { border-bottom:none; }
  .action-list .sev { flex-shrink:0; padding:4px 12px; border-radius:6px; font-size:12px; font-weight:bold; white-space:nowrap; margin-top:2px; }
  .action-list .sev.high { background:#fdecea; color:#c62828; }
  .action-list .sev.med { background:#fff6e0; color:#8a6100; }
  .action-list .txt strong { display:block; font-size:15px; color:#1a1a1a; }
  .action-list .txt span { font-size:14px; color:#666; }
  .standard-notes { list-style:disc; padding-left:20px; color:#666; font-size:14px; margin:14px 0 0; }
  .standard-notes li { margin-bottom:6px; }
  .summary-text { color:#444; font-size:15px; line-height:1.6; }
  .center-callout { text-align:center; padding:8px 0 18px; }
  .backup-flag { margin:18px -28px -26px -28px; padding:10px 0; text-align:center; font-weight:bold; color:#fff; border-radius:0 0 10px 10px; font-size:14px; letter-spacing:0.5px; }
  .backup-flag.healthy { background:#2e7d32; }
  .backup-flag.critical { background:#c62828; }
  .backup-flag.warn { background:#e6a100; }
  footer { margin-top:36px; color:#888; font-size:14px; }
</style>
</head>
<body>
  <h1>VMware Weekly Health Check - Dashboard</h1>
  <p class="subtitle">Generated $(ConvertTo-HtmlSafe $RunDateDisplay) - $($summaries.Count) site(s)</p>

  <nav class="tabs">
    $tabButtons
  </nav>

  <section class="page active" id="page-overview">
    <div class="stat-row">
      <div class="stat" style="background:#2e7d32"><span class="num">$($healthCounts.Healthy)</span><span class="label">Healthy Sites</span></div>
      <div class="stat" style="background:#e6a100"><span class="num">$($healthCounts.Warning)</span><span class="label">Sites with Warnings</span></div>
      <div class="stat" style="background:#c62828"><span class="num">$($healthCounts.Critical)</span><span class="label">Sites Critical</span></div>
      <div class="stat" style="background:#BF360C"><span class="num">$totalHigh</span><span class="label">Total High Risk Issues</span></div>
      <div class="stat" style="background:#37474F"><span class="num">$totalMedium</span><span class="label">Total Medium Risk Issues</span></div>
      <div class="stat" style="background:#1565C0"><span class="num">$totalHosts</span><span class="label">Total ESXi Hosts</span></div>
      <div class="stat" style="background:#00897B"><span class="num">$totalVMs</span><span class="label">Total VMs</span></div>
    </div>

    <div class="ov-grid">
$overviewCards
    </div>
  </section>

$sitePages

  <footer>VMware_Weekly_HealthCheck.ps1 - build $(ConvertTo-HtmlSafe $ScriptBuild)</footer>

<script>
function showPage(slug) {
  document.querySelectorAll('.page').forEach(function(el){ el.classList.remove('active'); });
  document.querySelectorAll('.tab').forEach(function(el){ el.classList.remove('active'); });
  var page = document.getElementById('page-' + slug);
  var tab = document.getElementById('tab-' + slug);
  if (page) { page.classList.add('active'); }
  if (tab) { tab.classList.add('active'); }
  window.scrollTo(0, 0);
}
</script>
</body>
</html>
"@

    $dashboardPath = Join-Path $OutputPath "VMware_HealthCheck_Dashboard_$RunDate.html"
    $html | Out-File -FilePath $dashboardPath -Encoding UTF8
    Write-Host "Dashboard written: $dashboardPath" -ForegroundColor Cyan
}

# ============================================================================
# 4. OUTPUT: COMBINED HTML DASHBOARD (only output artifact besides the log)
# ============================================================================
$SiteLabels = $Global:AllResults | Select-Object -ExpandProperty Site -Unique
Invoke-SafeCheck -CheckName 'Dashboard generation' -VCenter 'n/a' -Site 'n/a' -ObjectName 'Dashboard' -Script {
    Write-DashboardHtml -SiteLabels $SiteLabels -OutputPath $OutputPath -RunDateDisplay $RunDateDisplay -ScriptBuild $ScriptBuild
}

# ============================================================================
# 5. OUTPUT: LOG FILE (written last so it also captures any dashboard-generation failure)
# ============================================================================
$LogPath = Join-Path $OutputPath "VMware_Weekly_HealthCheck_$RunDate.log"
$LogLines = @()
$LogLines += "VMware Weekly Health Check run - $($ScriptStart.ToString('u')) - build $ScriptBuild"
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
