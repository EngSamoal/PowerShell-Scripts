# Get-VMToolsHealth.ps1  (Script 1 of 2)
# READ-ONLY VMware Tools health assessment for powered-on Windows/Linux VMs
# across ALL vCenters you are already connected to (Connect-VIServer). Changes nothing.
# Output: C:\temp\VMTools_Health.csv  (only VMs with an issue)
# Run   : .\Get-VMToolsHealth.ps1
param([string]$OutFile = 'C:\temp\VMTools_Health.csv')

$servers = @($global:DefaultVIServers | Where-Object IsConnected)
if (-not $servers) { Write-Error 'No connected vCenters. Run Connect-VIServer first.'; return }

$props    = 'Name','Runtime.PowerState','Runtime.ConnectionState','Config.Template','Config.GuestId',
            'Config.GuestFullName','Config.ManagedBy','Guest.ToolsStatus','Guest.ToolsRunningStatus',
            'Guest.ToolsVersionStatus2','Guest.IpAddress','Guest.GuestFullName','Guest.GuestFamily'
$outdated = 'guestToolsNeedUpgrade','guestToolsTooOld','guestToolsBlacklisted','guestToolsSupportedOld'
$linuxIds = 'linux|rhel|centos|ubuntu|debian|sles|suse|oracle|rocky|alma|fedora|amazon|asianux'

$report = foreach ($vc in $servers) {
    try   { $vms = Get-View -Server $vc -ViewType VirtualMachine -Property $props -Filter @{'Runtime.PowerState' = '^poweredOn$'} -ErrorAction Stop }
    catch { Write-Warning "[$($vc.Name)] Cannot read VMs: $($_.Exception.Message)"; continue }

    foreach ($vm in $vms) {
        $g = $vm.Guest; $c = $vm.Config
        $rs = $g.ToolsRunningStatus; $vs = $g.ToolsVersionStatus2
        $issues = @()

        if ($vm.Runtime.ConnectionState -ne 'connected') {
            # Orphaned / inaccessible / host disconnected: config & guest data cannot be trusted
            $issues += "INACCESSIBLE: VM connection state is '$($vm.Runtime.ConnectionState)' - cannot assess"
        }
        else {
            # Exclusions: templates, vCLS/solution-managed VMs (ManagedBy), Photon appliances (e.g. VCSA), non-Windows/Linux
            if ($c.Template -or $c.ManagedBy -or $c.GuestId -like 'vmwarePhoton*') { continue }
            $win   = $g.GuestFamily -eq 'windowsGuest' -or $c.GuestId -like 'win*'
            $linux = $g.GuestFamily -eq 'linuxGuest'   -or $c.GuestId -match $linuxIds
            if (-not ($win -or $linux)) { continue }

            if (-not $rs -and -not $vs) {
                $issues += 'UNKNOWN: VMware Tools status not reported by vCenter'
            }
            elseif ($vs -eq 'guestToolsNotInstalled' -or $g.ToolsStatus -eq 'toolsNotInstalled') {
                $issues += 'NONFUNCTIONAL: VMware Tools not installed'
            }
            else {
                if ($rs -notin 'guestToolsRunning','guestToolsExecutingScripts') {
                    $issues += "NONFUNCTIONAL: VMware Tools installed but not running ($rs)"
                }
                else {
                    if (-not $g.IpAddress)     { $issues += 'DEGRADED: Tools running but no guest IP reported (check guest NIC/IP config, vNIC connected)' }
                    if (-not $g.GuestFullName) { $issues += 'DEGRADED: Tools running but no guest OS reported' }
                }
                # Version is reported separately: outdated Tools often still report IP/OS fine
                if ($vs -in $outdated) { $issues += "OUTDATED: $vs (version issue only - not by itself the cause of missing IP/OS)" }
            }
        }
        if (-not $issues) { continue }

        [pscustomobject]@{
            'vCenter'                     = $vc.Name
            'VM Name'                     = $vm.Name
            'Power State'                 = "$($vm.Runtime.PowerState)"
            'VMware Tools Status'         = "$($g.ToolsStatus)"
            'VMware Tools Running Status' = "$rs"
            'VMware Tools Version Status' = "$vs"
            'IP Address'                  = $g.IpAddress
            'Guest OS'                    = $(if ($g.GuestFullName) { $g.GuestFullName }
                                              elseif ($c.GuestFullName) { "Not reported by Tools (configured: $($c.GuestFullName))" })
            'Identified Issue'            = $issues -join '; '
            'VM Id'                       = $vm.MoRef.Value   # unique per vCenter; used by Script 2 to avoid name mismatches
        }
    }
}

$report = @($report)
New-Item -ItemType Directory -Path (Split-Path $OutFile) -Force -ErrorAction SilentlyContinue | Out-Null
$report | Export-Csv -Path $OutFile -NoTypeInformation -Encoding UTF8
Write-Host "Affected VMs: $($report.Count)  ->  $OutFile" -ForegroundColor Cyan
$report | Group-Object { ($_.'Identified Issue' -split ':')[0] } | Sort-Object Count -Descending |
    Format-Table @{n='Primary issue';e={$_.Name}}, Count -AutoSize
