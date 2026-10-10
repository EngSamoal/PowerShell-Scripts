<#
.SYNOPSIS
    Simple, READ-ONLY full inventory export from vCenter: VMs, templates and vApps.

.DESCRIPTION
    Collects everything that lives in vCenter's VM inventory and writes it to CSV:

        Type = VM               - every virtual machine (powered on/off, orphaned, inaccessible)
        Type = Template         - every VM template in the VM & Templates inventory
        Type = vApp             - every vApp (deployed OVAs often land as a vApp)

    Note on "OVA": an .ova file is only a package. Once deployed it becomes a normal VM
    (or vApp) and is listed above. Content Library items are intentionally NOT included.

    Output (in -OutputFolder):
        vCenter_Inventory_<timestamp>.csv      - everything in one file (filter on Type)
        vCenter_Inventory_<timestamp>.xlsx     - one sheet per type, only if the ImportExcel
                                                 module is installed (optional)

    Nothing in vCenter is changed. Only Get-* cmdlets are used.

.PARAMETER vCenter
    One or more vCenter FQDNs/IPs. If omitted, the existing PowerCLI session(s) are used.

.PARAMETER OutputFolder
    Folder for the output files. Default: C:\Temp

.EXAMPLE
    .\Export-vCenterFullInventory.ps1 -vCenter vcsa01.corp.local
.EXAMPLE
    .\Export-vCenterFullInventory.ps1 -vCenter vcsa01.corp.local, vcsa02.corp.local -OutputFolder D:\Reports

.NOTES
    Requires VMware PowerCLI:  Install-Module VMware.PowerCLI -Scope CurrentUser
    Optional Excel output:     Install-Module ImportExcel -Scope CurrentUser
#>
[CmdletBinding()]
param(
    [string[]]$vCenter,
    [string]$OutputFolder = 'C:\Temp'
)

$ScriptBuild = 'Export-vCenterFullInventory build 2026-10-10.2'
Write-Host $ScriptBuild -ForegroundColor Cyan

#region ---- Connect ----
if ($vCenter) {
    Set-PowerCLIConfiguration -InvalidCertificateAction Ignore -Scope Session -Confirm:$false | Out-Null
    $cred = Get-Credential -Message 'vCenter credentials'
    foreach ($vc in $vCenter) {
        Write-Host "Connecting to $vc ..." -ForegroundColor Cyan
        Connect-VIServer -Server $vc -Credential $cred -ErrorAction Stop | Out-Null
    }
}
if (-not $global:DefaultVIServers -or @($global:DefaultVIServers).Count -eq 0) {
    throw 'Not connected to any vCenter. Use -vCenter <name> or run Connect-VIServer first.'
}
if (-not (Test-Path $OutputFolder)) { New-Item -ItemType Directory -Path $OutputFolder -Force | Out-Null }
#endregion

$results = New-Object System.Collections.Generic.List[object]

# Builds one output row with the same columns for every object type, so the CSV lines up.
function New-Row {
    param([hashtable]$p)
    $cols = 'vCenter','Type','Name','PowerState','ConnectionState','GuestOS','IPAddress',
            'DNSName','NumCPU','MemoryGB','ProvisionedGB','UsedGB','Datacenter','Cluster','Host',
            'Folder','ResourcePoolOrvApp','Datastores','Networks','ToolsStatus','HWVersion',
            'CreatedDate','Notes'
    $o = [ordered]@{}
    foreach ($c in $cols) { $o[$c] = if ($p.ContainsKey($c)) { $p[$c] } else { '' } }
    [pscustomobject]$o
}

foreach ($vc in @($global:DefaultVIServers)) {
    $vcName = $vc.Name
    Write-Host "[$vcName] Reading inventory ..." -ForegroundColor Cyan

    # ---- Lookup tables (MoRef -> name) so each VM doesn't need extra calls ----
    $nameMap = @{}
    $hostInfo = @{}   # host MoRef -> @{Name; Cluster; Datacenter}
    $dcOf = @{}       # any MoRef -> datacenter name

    foreach ($dc in Get-View -Server $vc -ViewType Datacenter -Property Name, VmFolder, HostFolder) {
        $nameMap["$($dc.MoRef)"] = $dc.Name
        # Every folder/cluster/host/datastore/network/pool under this DC
        foreach ($type in 'Folder','ClusterComputeResource','ComputeResource','HostSystem',
                          'Datastore','Network','ResourcePool','VirtualApp') {
            foreach ($v in Get-View -Server $vc -ViewType $type -SearchRoot $dc.MoRef -Property Name, Parent -ErrorAction SilentlyContinue) {
                $nameMap["$($v.MoRef)"] = $v.Name
                $dcOf["$($v.MoRef)"] = $dc.Name
                if ($type -eq 'HostSystem') {
                    $hostInfo["$($v.MoRef)"] = @{ Name = $v.Name; Parent = "$($v.Parent)"; Datacenter = $dc.Name }
                }
            }
        }
    }
    # Host's parent is a ClusterComputeResource (cluster) or ComputeResource (standalone host)
    foreach ($h in $hostInfo.Values) {
        $h.Cluster = if ($h.Parent -like 'ClusterComputeResource-*') { $nameMap[$h.Parent] } else { '(standalone)' }
    }
    function Get-Name($moref) { if ($moref) { $nameMap["$moref"] } }

    # ---- VMs and Templates (both are VirtualMachine objects; Config.Template tells them apart) ----
    $vmProps = 'Name','Config.Template','Config.GuestFullName','Config.Version','Config.Hardware.NumCPU',
               'Config.Hardware.MemoryMB','Config.Annotation','Config.CreateDate','Runtime.PowerState',
               'Runtime.ConnectionState','Runtime.Host','Guest.IpAddress','Guest.HostName','Guest.Net',
               'Guest.ToolsStatus','Guest.GuestFullName','Summary.Storage.Committed',
               'Summary.Storage.Uncommitted','Parent','ParentVApp','ResourcePool','Datastore','Network'
    $vmViews = @(Get-View -Server $vc -ViewType VirtualMachine -Property $vmProps)
    foreach ($v in $vmViews) {
        $isTemplate = [bool]$v.Config.Template
        $hi = if ($v.Runtime.Host) { $hostInfo["$($v.Runtime.Host)"] } else { $null }

        $ips = @($v.Guest.Net | ForEach-Object { $_.IpAddress } | Where-Object { $_ -and $_ -notlike 'fe80*' } | Select-Object -Unique)
        if ($ips.Count -eq 0 -and $v.Guest.IpAddress) { $ips = @($v.Guest.IpAddress) }

        $committed   = [double]$v.Summary.Storage.Committed
        $uncommitted = [double]$v.Summary.Storage.Uncommitted
        $os = if ($v.Guest.GuestFullName) { $v.Guest.GuestFullName } else { $v.Config.GuestFullName }

        # Orphaned/inaccessible VMs may have no Config at all - say so instead of leaving blanks
        $notes = $v.Config.Annotation
        if ("$($v.Runtime.ConnectionState)" -ne 'connected') {
            $notes = ("VM is '$($v.Runtime.ConnectionState)' in vCenter - config details unavailable. " + $notes).Trim()
        }

        $results.Add((New-Row @{
            vCenter            = $vcName
            Type               = if ($isTemplate) { 'Template' } else { 'VM' }
            Name               = $v.Name
            PowerState         = if ($isTemplate) { 'n/a (template)' } else { "$($v.Runtime.PowerState)" }
            ConnectionState    = "$($v.Runtime.ConnectionState)"
            GuestOS            = $os
            IPAddress          = ($ips -join ', ')
            DNSName            = $v.Guest.HostName
            NumCPU             = $v.Config.Hardware.NumCPU
            MemoryGB           = if ($v.Config.Hardware.MemoryMB) { [math]::Round($v.Config.Hardware.MemoryMB / 1024, 2) } else { '' }
            ProvisionedGB      = [math]::Round(($committed + $uncommitted) / 1GB, 2)
            UsedGB             = [math]::Round($committed / 1GB, 2)
            Datacenter         = if ($hi) { $hi.Datacenter } else { $dcOf["$($v.Parent)"] }
            Cluster            = if ($hi) { $hi.Cluster } else { '' }
            Host               = if ($hi) { $hi.Name } else { '' }
            Folder             = Get-Name $v.Parent
            ResourcePoolOrvApp = if ($v.ParentVApp) { Get-Name $v.ParentVApp } else { Get-Name $v.ResourcePool }
            Datastores         = (@($v.Datastore | ForEach-Object { Get-Name $_ }) -join ', ')
            Networks           = (@($v.Network   | ForEach-Object { Get-Name $_ }) -join ', ')
            ToolsStatus        = "$($v.Guest.ToolsStatus)"
            HWVersion          = $v.Config.Version
            CreatedDate        = $v.Config.CreateDate
            Notes              = $notes
        }))
    }
    $vmCount  = @($vmViews | Where-Object { -not $_.Config.Template }).Count
    $tplCount = @($vmViews | Where-Object { $_.Config.Template }).Count
    Write-Host "[$vcName]   VMs: $vmCount   Templates: $tplCount"

    # ---- vApps ----
    $vapps = @(Get-View -Server $vc -ViewType VirtualApp -Property Name, Parent, Summary.VAppState, VAppConfig.Annotation, Vm -ErrorAction SilentlyContinue)
    foreach ($a in $vapps) {
        $results.Add((New-Row @{
            vCenter            = $vcName
            Type               = 'vApp'
            Name               = $a.Name
            PowerState         = "$($a.Summary.VAppState)"
            Datacenter         = $dcOf["$($a.MoRef)"]
            ResourcePoolOrvApp = Get-Name $a.Parent
            Notes              = ("Contains $(@($a.Vm).Count) VM(s). " + $a.VAppConfig.Annotation).Trim()
        }))
    }
    Write-Host "[$vcName]   vApps: $($vapps.Count)"
}

#region ---- Export ----
$stamp   = Get-Date -Format 'yyyyMMdd_HHmm'
$csvPath = Join-Path $OutputFolder "vCenter_Inventory_$stamp.csv"
$results | Sort-Object vCenter, Type, Name | Export-Csv -Path $csvPath -NoTypeInformation -Encoding UTF8
Write-Host "`nCSV saved : $csvPath" -ForegroundColor Green

if (Get-Module -ListAvailable -Name ImportExcel) {
    $xlsxPath = Join-Path $OutputFolder "vCenter_Inventory_$stamp.xlsx"
    $results | Export-Excel -Path $xlsxPath -WorksheetName 'All' -AutoSize -AutoFilter -FreezeTopRow -BoldTopRow
    foreach ($t in 'VM','Template','vApp') {
        $rows = @($results | Where-Object { $_.Type -eq $t })
        if ($rows.Count -gt 0) {
            $rows | Export-Excel -Path $xlsxPath -WorksheetName $t -AutoSize -AutoFilter -FreezeTopRow -BoldTopRow
        }
    }
    Write-Host "Excel saved: $xlsxPath" -ForegroundColor Green
} else {
    Write-Host '(Install-Module ImportExcel to also get an .xlsx with one sheet per type.)' -ForegroundColor DarkGray
}

# Summary
$results | Group-Object vCenter, Type | Select-Object @{n='vCenter / Type';e={$_.Name}}, Count | Format-Table -AutoSize | Out-Host
#endregion
