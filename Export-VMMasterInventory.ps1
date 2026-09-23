<#
.SYNOPSIS
    Builds a single, formatted master VM inventory workbook (.xlsx) with one worksheet per site.

.DESCRIPTION
    READ-ONLY. For every site defined in the SITE CONFIGURATION block, the script queries the
    site's vCenter(s) with Get-View (property-filtered, read-only calls only) and writes:

        #  |  VM NAME  |  IP  |  OS  |  Power Status  |  Implementor

    to a dedicated worksheet, formatted as an Excel Table with filters, a frozen header,
    conditional formatting on Power Status and site-coloured worksheet tabs.

    Existing PowerCLI sessions ($global:DefaultVIServers) are reused. A vCenter is only
    connected to if no live session exists (use -NoConnect to forbid that). Sessions opened
    by this script are closed again at the end; sessions you opened yourself are left alone.

    No VM, host or vCenter setting is changed. The only cmdlets used against vCenter are
    Get-View and Get-Datacenter.

.PARAMETER OutputPath
    Full path of the .xlsx to create. Default: Documents\VM_Master_Inventory_<timestamp>.xlsx

.PARAMETER SampleData
    Generate the workbook from fictitious VMs (no vCenter access, PowerCLI not required).

.PARAMETER NoConnect
    Only use existing PowerCLI sessions; never call Connect-VIServer.

.PARAMETER IncludeTemplates
    Include VM templates (excluded by default).

.EXAMPLE
    .\Export-VMMasterInventory.ps1
.EXAMPLE
    .\Export-VMMasterInventory.ps1 -OutputPath 'D:\Reports\VM_Inventory.xlsx' -NoConnect
.EXAMPLE
    .\Export-VMMasterInventory.ps1 -SampleData

.NOTES
    Requirements : Windows PowerShell 5.1 or PowerShell 7.x
                   Microsoft Excel (desktop) installed on this machine - used to build the .xlsx
                   VMware PowerCLI (not needed with -SampleData):
                       Install-Module VMware.PowerCLI -Scope CurrentUser
#>
[CmdletBinding()]
param(
    [string]$OutputPath = (Join-Path ([Environment]::GetFolderPath('C:\Temp')) ('VM_Master_Inventory_{0}.xlsx' -f (Get-Date -Format 'yyyyMMdd_HHmm'))),
    [switch]$SampleData,
    [switch]$NoConnect,
    [switch]$IncludeTemplates
)

#region ======================= SITE CONFIGURATION (edit here) =======================
# One entry per worksheet, in tab order. To add a site, copy a line and change the values.
#   Name       : worksheet / tab name (max 31 chars, no  [ ] : * ? / \ )
#   VCenter    : one or more vCenter IPs or FQDNs for this site, e.g. @('10.1.1.20')
#                (an existing PowerCLI session is reused whether it was opened by IP or by name)
#   Datacenter : optional - limit to these vCenter datacenters (use when sites share a vCenter)
#   TabColor   : hex RGB for the tab and the sheet title banner
#   VCenter = @() leaves that tab blank for manual entry (nothing is collected for it)
$Sites = @(
    [pscustomobject]@{ Name = 'SixFlags'; VCenter = @('10.50.10.10'); Datacenter = @(); TabColor = '1F4E79' }
    [pscustomobject]@{ Name = 'AquaArabia'; VCenter = @(); Datacenter = @(); TabColor = '2E7D32' }   # same vCenter as SixFlags - left blank, fill manually
    [pscustomobject]@{ Name = 'SEVEN Tabuk'; VCenter = @('10.52.6.140'); Datacenter = @(); TabColor = 'C55A11' }
    [pscustomobject]@{ Name = 'SEVEN ABHA'; VCenter = @('10.61.6.140'); Datacenter = @(); TabColor = '7030A0' }
    [pscustomobject]@{ Name = 'SEVEN ALHamra'; VCenter = @('10.11.7.120'); Datacenter = @(); TabColor = 'A50021' }
)

# OT worksheet - added after the site tabs. Same columns plus 'Site', so OT VMs from every site
# are listed together. (How OT VMs are identified in vCenter is still to be confirmed.)
$OTSheet = [pscustomobject]@{ Name = 'OT'; TabColor = '00838F' }
#endregion ==========================================================================

#region ---------- Style constants ----------
$Style = @{
    Font          = 'Arial'
    HeaderFill    = '1F3864'   # navy, identical on every sheet
    HeaderFont    = 'FFFFFF'
    GridLine      = 'BFBFBF'
    ManualFill    = 'FFF2CC'   # Implementor = manual entry
    TableStyle    = 'TableStyleMedium2'
    NotAvailable  = 'N/A'      # shown when no IP could be read
    UnknownOS     = 'Unknown'  # shown when no OS could be read
}
$Headers = '#', 'VM NAME', 'IP', 'OS', 'Power Status', 'Implementor'
$OTHeaders = '#', 'Site', 'VM NAME', 'IP', 'OS', 'Power Status', 'Implementor'
#endregion

$ErrorActionPreference = 'Stop'
$script:OpenedSessions = New-Object System.Collections.Generic.List[object]
$LogPath = [IO.Path]::ChangeExtension($OutputPath, '.log')

function Write-Log {
    param([string]$Message, [ValidateSet('INFO', 'WARN', 'ERROR', 'OK')][string]$Level = 'INFO')
    $line = '{0} [{1,-5}] {2}' -f (Get-Date -Format 'yyyy-MM-dd HH:mm:ss'), $Level, $Message
    $color = @{ INFO = 'Gray'; WARN = 'Yellow'; ERROR = 'Red'; OK = 'Green' }[$Level]
    Write-Host $line -ForegroundColor $color
    try { Add-Content -Path $LogPath -Value $line -Encoding UTF8 } catch { }
}

function ConvertTo-OleColor([string]$Hex) {
    $Hex = $Hex.TrimStart('#')
    $r = [Convert]::ToInt32($Hex.Substring(0, 2), 16)
    $g = [Convert]::ToInt32($Hex.Substring(2, 2), 16)
    $b = [Convert]::ToInt32($Hex.Substring(4, 2), 16)
    return $r + ($g * 256) + ($b * 65536)
}

function Get-SafeSheetName([string]$Name) {
    $n = ($Name -replace '[\[\]\:\*\?\/\\]', '-').Trim()
    if ($n.Length -gt 31) { $n = $n.Substring(0, 31) }
    if (-not $n) { $n = 'Site' }
    return $n
}

function New-SiteResult($Site, [string[]]$Columns = $Headers) {
    [pscustomobject]@{
        Site        = $Site.Name
        Columns     = $Columns
        VCenter     = (@($Site.VCenter) -join ', ')
        Manual      = (@($Site.VCenter).Count -eq 0)
        TabColor    = $Site.TabColor
        Rows        = New-Object System.Collections.Generic.List[object]
        QueryOk     = 0
        QueryFailed = 0
        VmErrors    = 0
        Errors      = New-Object System.Collections.Generic.List[string]
    }
}

#region ---------- vCenter collection (read-only) ----------
function Resolve-HostAddress([string]$HostName) {
    # Returns the IP address(es) for a name or IP; empty if it cannot be resolved
    $ip = $null
    if ([Net.IPAddress]::TryParse($HostName, [ref]$ip)) { return $ip.ToString() }
    try { return @([Net.Dns]::GetHostAddresses($HostName) | ForEach-Object { $_.ToString() }) } catch { return @() }
}

function Get-VCenterSession([string]$Server) {
    $live = @($global:DefaultVIServers | Where-Object { $_.IsConnected })
    $existing = @($live | Where-Object { $_.Name -eq $Server -or $_.ServiceUri.Host -eq $Server })

    # Config may use an IP while the session was opened by FQDN (or the reverse):
    # compare the resolved IP addresses before deciding a new login is needed.
    if ($existing.Count -eq 0 -and $live.Count -gt 0) {
        $wanted = @(Resolve-HostAddress $Server)
        if ($wanted.Count -gt 0) {
            $existing = @($live | Where-Object {
                    $have = @(Resolve-HostAddress $_.Name)
                    @($have | Where-Object { $wanted -contains $_ }).Count -gt 0
                })
        }
    }

    if ($existing.Count -gt 0) {
        Write-Log "Reusing existing PowerCLI session to $Server (user: $($existing[0].User))"
        return $existing[0]
    }
    if ($NoConnect) { throw "No live PowerCLI session to $Server and -NoConnect was specified." }

    Write-Log "No existing session to $Server - connecting (you may be prompted for credentials)" 'WARN'
    $session = Connect-VIServer -Server $Server -ErrorAction Stop
    $script:OpenedSessions.Add($session)
    return $session
}

function Get-PrimaryIP($View) {
    $candidates = New-Object System.Collections.Generic.List[string]
    if ($View.Guest.IpAddress) { $candidates.Add([string]$View.Guest.IpAddress) }     # Tools-reported primary
    foreach ($nic in @($View.Guest.Net)) {
        if ($nic -and $nic.IpAddress) { foreach ($a in @($nic.IpAddress)) { $candidates.Add([string]$a) } }
    }
    # Prefer a routable IPv4 address
    foreach ($ip in $candidates) {
        $addr = $null
        if ([Net.IPAddress]::TryParse($ip, [ref]$addr) -and $addr.AddressFamily -eq 'InterNetwork' -and
            $ip -notlike '169.254.*' -and $ip -notlike '127.*' -and $ip -ne '0.0.0.0') { return $ip }
    }
    # Fall back to a global IPv6 address
    foreach ($ip in $candidates) {
        $addr = $null
        if ([Net.IPAddress]::TryParse($ip, [ref]$addr) -and $addr.AddressFamily -eq 'InterNetworkV6' -and
            -not $addr.IsIPv6LinkLocal -and -not [Net.IPAddress]::IsLoopback($addr)) { return $ip }
    }
    return $null
}

function Get-SiteInventory($Site) {
    $result = New-SiteResult $Site
    $props = 'Name', 'Runtime.PowerState', 'Guest.IpAddress', 'Guest.Net', 'Guest.GuestFullName',
             'Config.GuestFullName', 'Config.Template'

    foreach ($vc in @($Site.VCenter)) {
        try {
            $server = Get-VCenterSession $vc

            $views = New-Object System.Collections.Generic.List[object]
            if (@($Site.Datacenter).Count -gt 0) {
                foreach ($dcName in @($Site.Datacenter)) {
                    $dc = Get-Datacenter -Server $server -Name $dcName -ErrorAction Stop
                    foreach ($v in @(Get-View -Server $server -ViewType VirtualMachine -Property $props -SearchRoot $dc.ExtensionData.MoRef)) { $views.Add($v) }
                }
            }
            else {
                foreach ($v in @(Get-View -Server $server -ViewType VirtualMachine -Property $props)) { $views.Add($v) }
            }
            $result.QueryOk++
            Write-Log "[$($Site.Name)] $vc returned $($views.Count) VM objects" 'OK'
        }
        catch {
            $result.QueryFailed++
            $msg = "[$($Site.Name)] Query to $vc FAILED: $($_.Exception.Message)"
            $result.Errors.Add($msg)
            Write-Log $msg 'ERROR'
            continue
        }

        foreach ($v in $views) {
            try {
                if (-not $IncludeTemplates -and $v.Config -and $v.Config.Template) { continue }

                $ip = Get-PrimaryIP $v
                $os = $v.Guest.GuestFullName                          # detected by VMware Tools
                if (-not $os) { $os = $v.Config.GuestFullName }       # configured guest OS (works when powered off)

                $state = [string]$v.Runtime.PowerState               # poweredOn / poweredOff / suspended
                if ($state) { $state = $state.Substring(0, 1).ToUpper() + $state.Substring(1) } else { $state = $Style.NotAvailable }

                $result.Rows.Add([pscustomobject]@{
                        Name  = $v.Name
                        IP    = $(if ($ip) { $ip } else { $Style.NotAvailable })
                        OS    = $(if ($os) { $os } else { $Style.UnknownOS })
                        Power = $state
                    })
            }
            catch {
                $result.VmErrors++
                $name = try { $v.Name } catch { '<unreadable>' }
                Write-Log "[$($Site.Name)] VM '$name' could not be read: $($_.Exception.Message)" 'WARN'
                $result.Rows.Add([pscustomobject]@{ Name = $name; IP = $Style.NotAvailable; OS = $Style.UnknownOS; Power = $Style.NotAvailable })
            }
        }
    }

    $sorted = @($result.Rows | Sort-Object Name)
    $result.Rows.Clear()
    foreach ($r in $sorted) { $result.Rows.Add($r) }

    $dupes = @($result.Rows | Group-Object Name | Where-Object Count -gt 1)
    if ($dupes.Count -gt 0) {
        Write-Log "[$($Site.Name)] Duplicate VM names (kept, highlighted orange): $(($dupes.Name) -join ', ')" 'WARN'
    }
    return $result
}
#endregion

#region ---------- Sample data (-SampleData) ----------
function Get-SampleInventory($Site, [int]$Index) {
    $result = New-SiteResult $Site
    $result.QueryOk = 1
    $p = 'S' + [char](65 + $Index) + '-'   # SA-, SB-, ...
    $w19 = 'Microsoft Windows Server 2019 (64-bit)'
    $w22 = 'Microsoft Windows Server 2022 (64-bit)'
    $w25 = 'Microsoft Windows Server 2025 (64-bit)'
    $r8 = 'Red Hat Enterprise Linux 8 (64-bit)'
    $r9 = 'Red Hat Enterprise Linux 9 (64-bit)'
    $ub = 'Ubuntu Linux (64-bit)'
    $net = '10.{0}.' -f (10 * ($Index + 1))

    $sets = @(
        @(  @('DC01', "${net}10.11", $w22, 'PoweredOn'), @('DC02', "${net}10.12", $w22, 'PoweredOn'),
            @('SQL01', "${net}20.21", $w19, 'PoweredOn'), @('APP01', "${net}30.31", $w25, 'PoweredOn'),
            @('WEB01', "${net}30.41", $r9, 'PoweredOn'), @('WEB02', "${net}30.42", $r9, 'PoweredOn'),
            @('FS01', $null, $w19, 'PoweredOff'), @('JUMP01', "${net}40.10", $w25, 'PoweredOn'),
            @('LEGACY01', $null, $null, 'PoweredOff') ),
        @(  @('DC01', "${net}10.11", $w25, 'PoweredOn'), @('EXCH01', "${net}20.15", $w19, 'PoweredOn'),
            @('ORA01', "${net}20.30", $r8, 'PoweredOn'), @('ORA02', $null, $r8, 'Suspended'),
            @('APP01', "${net}30.31", $w22, 'PoweredOn'), @('MON01', "${net}50.5", $ub, 'PoweredOn'),
            @('BKP01', "${net}60.10", $w22, 'PoweredOn'), @('TEST01', $null, $w25, 'PoweredOff'),
            @('K8S-N01', "${net}70.11", $r9, 'PoweredOn'), @('K8S-N02', $null, $r9, 'PoweredOn') ),
        @(  @('DC01', "${net}10.11", $w19, 'PoweredOn'), @('SQL01', "${net}20.21", $w22, 'PoweredOn'),
            @('APP01', "${net}30.31", $w22, 'PoweredOn'), @('APP01', $null, $w22, 'PoweredOff'),
            @('WEB01', "${net}30.41", $r8, 'PoweredOn'), @('SFTP01', "${net}40.20", $r9, 'PoweredOn'),
            @('PRINT01', $null, $null, 'PoweredOn'), @('OLD-FS02', $null, $w19, 'PoweredOff') ),
        @(  @('DC01', "${net}10.11", $w25, 'PoweredOn'), @('DC02', "${net}10.12", $w25, 'PoweredOn'),
            @('SQL01', "${net}20.21", $w22, 'PoweredOn'), @('SQL02', "${net}20.22", $w22, 'PoweredOff'),
            @('APP01', "${net}30.31", $w19, 'PoweredOn'), @('LNX-DB01', "${net}20.50", $r9, 'PoweredOn'),
            @('LNX-APP01', "${net}30.60", $r8, 'PoweredOn'), @('SIEM01', "${net}50.9", $r9, 'PoweredOn'),
            @('TMP-BUILD01', $null, $null, 'PoweredOff') )
    )
    foreach ($s in $sets[$Index % $sets.Count]) {
        $result.Rows.Add([pscustomobject]@{
                Name  = $p + $s[0]
                IP    = $(if ($s[1]) { $s[1] } else { $Style.NotAvailable })
                OS    = $(if ($s[2]) { $s[2] } else { $Style.UnknownOS })
                Power = $s[3]
            })
    }
    $sorted = @($result.Rows | Sort-Object Name)
    $result.Rows.Clear()
    foreach ($r in $sorted) { $result.Rows.Add($r) }
    return $result
}

function Get-SampleOTInventory($OT) {
    $result = New-SiteResult $OT $OTHeaders
    $result.QueryOk = 1
    $result.VCenter = 'OT'
    $w16 = 'Microsoft Windows Server 2016 (64-bit)'
    $w19 = 'Microsoft Windows Server 2019 (64-bit)'
    $w22 = 'Microsoft Windows Server 2022 (64-bit)'
    $w10 = 'Microsoft Windows 10 (64-bit)'
    $r8 = 'Red Hat Enterprise Linux 8 (64-bit)'
    $set = @(
        @($Sites[0 % $Sites.Count].Name, 'OT-SA-SCADA01', '172.16.10.11', $w19, 'PoweredOn'),
        @($Sites[0 % $Sites.Count].Name, 'OT-SA-HMI01', '172.16.10.21', $w10, 'PoweredOn'),
        @($Sites[0 % $Sites.Count].Name, 'OT-SA-HIST01', '172.16.10.31', $w22, 'PoweredOn'),
        @($Sites[1 % $Sites.Count].Name, 'OT-SB-SCADA01', '172.16.20.11', $w19, 'PoweredOn'),
        @($Sites[1 % $Sites.Count].Name, 'OT-SB-EWS01', $null, $w10, 'PoweredOff'),
        @($Sites[1 % $Sites.Count].Name, 'OT-SB-OPC01', '172.16.20.41', $w16, 'PoweredOn'),
        @($Sites[2 % $Sites.Count].Name, 'OT-SC-SCADA01', '172.16.30.11', $w22, 'PoweredOn'),
        @($Sites[2 % $Sites.Count].Name, 'OT-SC-HIST01', $null, $w19, 'Suspended'),
        @($Sites[3 % $Sites.Count].Name, 'OT-SD-SCADA01', '172.16.40.11', $w22, 'PoweredOn'),
        @($Sites[3 % $Sites.Count].Name, 'OT-SD-LOG01', '172.16.40.51', $r8, 'PoweredOn'),
        @($Sites[3 % $Sites.Count].Name, 'OT-SD-PLCGW01', $null, $null, 'PoweredOff')
    )
    foreach ($s in $set) {
        $result.Rows.Add([pscustomobject]@{
                Site  = $s[0]
                Name  = $s[1]
                IP    = $(if ($s[2]) { $s[2] } else { $Style.NotAvailable })
                OS    = $(if ($s[3]) { $s[3] } else { $Style.UnknownOS })
                Power = $s[4]
            })
    }
    return $result
}
#endregion

#region ---------- Excel output (Excel COM) ----------
function Format-SiteSheet($Excel, $Ws, $SiteResult, [bool]$IsSample) {
    $Ws.Name = Get-SafeSheetName $SiteResult.Site
    $Ws.Cells.Font.Name = $Style.Font
    $Ws.Cells.Font.Size = 10
    $siteColor = ConvertTo-OleColor $SiteResult.TabColor
    $Ws.Tab.Color = $siteColor

    $rows = $SiteResult.Rows
    $n = $rows.Count
    $headerRow = 3
    $lastRow = $headerRow + [Math]::Max($n, 1)

    # --- Build the whole sheet (title, info line, header, data) as ONE array and write it in a
    #     single call. Mixing scalar and array writes to .Value2 trips PowerShell's COM binder
    #     cache (arrays then get stringified), so every cell value goes through this one write.
    $on = @($rows | Where-Object Power -eq 'PoweredOn').Count
    $off = @($rows | Where-Object Power -eq 'PoweredOff').Count
    $sus = @($rows | Where-Object Power -eq 'Suspended').Count
    $src = if ($IsSample) { 'SAMPLE DATA (fictitious VMs)' } else { "vCenter: $($SiteResult.VCenter)" }
    $info = "{0}  |  Generated {1}  |  {2} VMs: {3} on, {4} off, {5} suspended  |  Yellow = manual entry, orange name = duplicate" -f `
        $src, (Get-Date -Format 'yyyy-MM-dd HH:mm'), $n, $on, $off, $sus
    if ($SiteResult.Manual) { $info = 'Filled in manually  |  Yellow = manual entry, orange name = duplicate' }
    if ($SiteResult.QueryFailed -gt 0) { $info = "COLLECTION INCOMPLETE - $($SiteResult.QueryFailed) vCenter query failed (see log)   |   " + $info }

    # Column layout is driven by the header list, so sheets can carry extra columns (e.g. OT + Site)
    $cols = @($SiteResult.Columns)
    $nc = $cols.Count
    $lc = [string][char](64 + $nc)                        # last column letter
    $ci = @{}; for ($c = 0; $c -lt $nc; $c++) { $ci[$cols[$c]] = $c + 1 }
    $field = @{ 'Site' = 'Site'; 'VM NAME' = 'Name'; 'IP' = 'IP'; 'OS' = 'OS'; 'Power Status' = 'Power' }

    $h = $headerRow - 1                                   # array row index of the header
    $data = New-Object 'object[,]' $lastRow, $nc
    $data[0, 0] = "$($SiteResult.Site)  -  VM Inventory"
    $data[1, 0] = $info
    for ($c = 0; $c -lt $nc; $c++) { $data[$h, $c] = $cols[$c] }
    for ($i = 0; $i -lt $n; $i++) {
        $r = $rows[$i]
        for ($c = 0; $c -lt $nc; $c++) {
            $name = $cols[$c]
            if ($name -eq '#') { $data[($h + 1 + $i), $c] = $i + 1 }
            elseif ($field.ContainsKey($name)) { $data[($h + 1 + $i), $c] = [string]$r.($field[$name]) }
            else { $data[($h + 1 + $i), $c] = '' }
        }
    }
    if ($n -eq 0 -and -not $SiteResult.Manual) { $data[($h + 1), ($ci['VM NAME'] - 1)] = $(if ($SiteResult.QueryFailed -gt 0) { 'No data - vCenter query failed' } else { 'No VMs found' }) }

    $Ws.Range($Ws.Cells.Item($headerRow + 1, $ci['IP']), $Ws.Cells.Item($lastRow, $ci['IP'])).NumberFormat = '@'   # IP stored as text
    try {
        $Ws.Range("A1:$lc$lastRow").Value2 = $data
        if ([string]$Ws.Cells.Item($headerRow, $nc).Value2 -ne $cols[$nc - 1]) { throw 'bulk write verification failed' }
    }
    catch {
        Write-Log "Bulk write not accepted ($($_.Exception.Message)) - writing cell by cell" 'WARN'
        $Ws.Range("A1:$lc$lastRow").ClearContents() | Out-Null
        for ($r = 0; $r -lt $lastRow; $r++) {
            for ($c = 0; $c -lt $nc; $c++) {
                if ($null -ne $data[$r, $c] -and $data[$r, $c] -ne '') { $Ws.Cells.Item($r + 1, $c + 1).Value2 = $data[$r, $c] }
            }
        }
    }

    # --- Title banner (row 1) and info line (row 2)
    $title = $Ws.Range("A1:${lc}1")
    $title.HorizontalAlignment = 7            # centre across selection (no merged cells)
    $title.VerticalAlignment = -4108
    $title.Interior.Color = $siteColor
    $title.Font.Color = ConvertTo-OleColor 'FFFFFF'
    $title.Font.Bold = $true
    $title.Font.Size = 14
    $Ws.Rows.Item(1).RowHeight = 30

    $infoRange = $Ws.Range("A2:${lc}2")
    $infoRange.Font.Size = 9
    $infoRange.Font.Italic = $true
    $infoRange.Font.Color = ConvertTo-OleColor $(if ($SiteResult.QueryFailed -gt 0) { 'C00000' } else { '595959' })
    $Ws.Rows.Item(2).RowHeight = 18

    # --- Excel Table (filters, banding, easy to extend)
    $lo = $Ws.ListObjects.Add(1, $Ws.Range("A${headerRow}:$lc$lastRow"), $null, 1)
    $lo.Name = 'tbl_' + (($SiteResult.Site -replace '[^A-Za-z0-9_]', '_'))
    $lo.TableStyle = $Style.TableStyle
    $lo.ShowTableStyleRowStripes = $true

    $hdr = $lo.HeaderRowRange
    $hdr.Interior.Color = ConvertTo-OleColor $Style.HeaderFill
    $hdr.Font.Color = ConvertTo-OleColor $Style.HeaderFont
    $hdr.Font.Bold = $true
    $hdr.Font.Size = 11
    $hdr.HorizontalAlignment = -4108
    $hdr.VerticalAlignment = -4108
    $Ws.Rows.Item($headerRow).RowHeight = 24

    $body = $lo.DataBodyRange
    $body.VerticalAlignment = -4108
    $body.RowHeight = 18
    $lo.Range.Borders.LineStyle = 1
    $lo.Range.Borders.Weight = 2
    $lo.Range.Borders.Color = ConvertTo-OleColor $Style.GridLine

    $colNum = $lo.ListColumns.Item($ci['#']).DataBodyRange
    $colName = $lo.ListColumns.Item($ci['VM NAME']).DataBodyRange
    $colIP = $lo.ListColumns.Item($ci['IP']).DataBodyRange
    $colOS = $lo.ListColumns.Item($ci['OS']).DataBodyRange
    $colPwr = $lo.ListColumns.Item($ci['Power Status']).DataBodyRange
    $colImpl = $lo.ListColumns.Item($ci['Implementor']).DataBodyRange

    # Optional Site column: centred, bold, text in that site's tab colour
    if ($ci.ContainsKey('Site')) {
        $colSite = $lo.ListColumns.Item($ci['Site']).DataBodyRange
        $colSite.HorizontalAlignment = -4108
        $colSite.Font.Bold = $true
        foreach ($s in $Sites) {
            $fc = $colSite.FormatConditions.Add(1, 3, "=""$($s.Name)""")
            $fc.Font.Color = ConvertTo-OleColor $s.TabColor
        }
    }

    $colNum.HorizontalAlignment = -4108
    $colName.HorizontalAlignment = -4131
    $colIP.HorizontalAlignment = -4131
    $colOS.HorizontalAlignment = -4131
    $colPwr.HorizontalAlignment = -4108
    $colPwr.Font.Bold = $true
    $colImpl.Interior.Color = ConvertTo-OleColor $Style.ManualFill

    # --- Conditional formatting: Power Status
    foreach ($cf in @(
            @('PoweredOn', 'C6EFCE', '006100'),
            @('PoweredOff', 'FFC7CE', '9C0006'),
            @('Suspended', 'FFEB9C', '9C5700'))) {
        $fc = $colPwr.FormatConditions.Add(1, 3, "=""$($cf[0])""")
        $fc.Interior.Color = ConvertTo-OleColor $cf[1]
        $fc.Font.Color = ConvertTo-OleColor $cf[2]
    }
    # Missing IP / OS -> grey italic
    foreach ($pair in @(@($colIP, $Style.NotAvailable), @($colOS, $Style.UnknownOS))) {
        $fc = $pair[0].FormatConditions.Add(1, 3, "=""$($pair[1])""")
        $fc.Font.Color = ConvertTo-OleColor '808080'
        $fc.Font.Italic = $true
    }
    # Duplicate VM names -> orange
    if ($n -gt 1) {
        $dup = $colName.FormatConditions.AddUniqueValues()
        $dup.DupeUnique = 1
        $dup.Interior.Color = ConvertTo-OleColor 'F8CBAD'
        $dup.Font.Color = ConvertTo-OleColor '843C0C'
    }

    # --- Column widths: longest value per column (not the banner rows), clamped to min / max
    $limitsByName = @{ '#' = @(6, 8); 'Site' = @(12, 24); 'VM NAME' = @(26, 45); 'IP' = @(16, 40); 'OS' = @(34, 55); 'Power Status' = @(19, 20); 'Implementor' = @(24, 40) }
    $limits = @{}; for ($c = 0; $c -lt $nc; $c++) { $limits[$c] = $(if ($limitsByName.ContainsKey($cols[$c])) { $limitsByName[$cols[$c]] } else { @(12, 40) }) }
    for ($c = 0; $c -lt $nc; $c++) {
        $longest = 0
        for ($r = $h; $r -le $data.GetUpperBound(0); $r++) {
            $len = ([string]$data[$r, $c]).Length
            if ($len -gt $longest) { $longest = $len }
        }
        $width = [Math]::Min([Math]::Max($longest * 1.1 + 4, $limits[$c][0]), $limits[$c][1])
        $Ws.Columns.Item($c + 1).ColumnWidth = $width
    }

    # --- View: freeze header, hide gridlines
    $Ws.Activate()
    $win = $Excel.ActiveWindow
    $win.FreezePanes = $false
    $win.SplitColumn = 0
    $win.SplitRow = $headerRow
    $win.FreezePanes = $true
    $win.DisplayGridlines = $false
    $Ws.Range('A' + ($headerRow + 1)).Select() | Out-Null

    # --- Print setup (best effort - needs a printer driver)
    try {
        $Excel.PrintCommunication = $false
        $ps = $Ws.PageSetup
        $ps.Orientation = 2
        $ps.Zoom = $false
        $ps.FitToPagesWide = 1
        $ps.FitToPagesTall = $false
        $ps.PrintTitleRows = "`$${headerRow}:`$${headerRow}"
        $ps.CenterFooter = '&P / &N'
        $ps.LeftFooter = "&A - VM Inventory"
        $Excel.PrintCommunication = $true
    }
    catch { try { $Excel.PrintCommunication = $true } catch { } }
}

function Export-InventoryWorkbook($SiteResults, [string]$Path, [bool]$IsSample) {
    $excel = $null; $wb = $null
    try {
        $excel = New-Object -ComObject Excel.Application
        $excel.Visible = $false
        $excel.DisplayAlerts = $false
        $excel.ScreenUpdating = $false

        $wb = $excel.Workbooks.Add()
        $missing = [Reflection.Missing]::Value
        while ($wb.Worksheets.Count -lt $SiteResults.Count) { [void]$wb.Worksheets.Add($missing, $wb.Worksheets.Item($wb.Worksheets.Count)) }
        while ($wb.Worksheets.Count -gt $SiteResults.Count) { $wb.Worksheets.Item($wb.Worksheets.Count).Delete() }

        for ($i = 0; $i -lt $SiteResults.Count; $i++) {
            Write-Log "Formatting worksheet '$($SiteResults[$i].Site)' ($($SiteResults[$i].Rows.Count) rows)"
            Format-SiteSheet $excel $wb.Worksheets.Item($i + 1) $SiteResults[$i] $IsSample
        }
        $wb.Worksheets.Item(1).Activate()
        $excel.ScreenUpdating = $true

        if (Test-Path $Path) {
            try { Remove-Item $Path -Force -ErrorAction Stop }
            catch {
                # Usually the previous copy is still open in Excel - save alongside instead of failing
                $Path = Join-Path (Split-Path $Path -Parent) ('{0}_{1}.xlsx' -f [IO.Path]::GetFileNameWithoutExtension($Path), (Get-Date -Format 'HHmmss'))
                Write-Log "Existing file is locked (open in Excel?) - saving as $Path instead" 'WARN'
            }
        }
        $wb.SaveAs($Path, 51)   # 51 = xlOpenXMLWorkbook (.xlsx)
        $wb.Close($false)
        return $Path
    }
    finally {
        if ($excel) { $excel.Quit() }
        foreach ($o in @($wb, $excel)) { if ($o) { [void][Runtime.InteropServices.Marshal]::ReleaseComObject($o) } }
        [GC]::Collect(); [GC]::WaitForPendingFinalizers()
    }
}
#endregion

#region ---------- Main ----------
$OutputPath = [IO.Path]::GetFullPath($OutputPath)
$outDir = Split-Path $OutputPath -Parent
if (-not (Test-Path $outDir)) { New-Item -ItemType Directory -Path $outDir -Force | Out-Null }

Write-Log "===== VM master inventory started ($(if ($SampleData) { 'SAMPLE mode' } else { 'LIVE, read-only' })) ====="
Write-Log "Sites configured: $(($Sites.Name) -join ', ')"

if (-not $SampleData) {
    try { Import-Module VMware.VimAutomation.Core -ErrorAction Stop -WarningAction SilentlyContinue }
    catch {
        Write-Log 'VMware PowerCLI is not installed. Run:  Install-Module VMware.PowerCLI -Scope CurrentUser' 'ERROR'
        return
    }
}

$results = New-Object System.Collections.Generic.List[object]
for ($i = 0; $i -lt $Sites.Count; $i++) {
    $site = $Sites[$i]
    if (-not $site.TabColor) { $site.TabColor = @('1F4E79', '2E7D32', 'C55A11', '7030A0', '00838F', '8B1A1A', '5D4037', '37474F')[$i % 8] }
    Write-Log "----- $($site.Name) -----"
    try {
        if (@($site.VCenter).Count -eq 0) { Write-Log "[$($site.Name)] No vCenter configured - tab left blank for manual entry"; $results.Add((New-SiteResult $site)) }
        elseif ($SampleData) { $results.Add((Get-SampleInventory $site $i)) }
        else { $results.Add((Get-SiteInventory $site)) }
    }
    catch {
        Write-Log "[$($site.Name)] Unexpected failure: $($_.Exception.Message)" 'ERROR'
        $r = New-SiteResult $site; $r.QueryFailed++; $results.Add($r)
    }
}

if ($OTSheet) {
    Write-Log "----- $($OTSheet.Name) -----"
    if ($SampleData) { $results.Add((Get-SampleOTInventory $OTSheet)) }
    else {
        # Live OT collection is wired in once the OT identification rule is confirmed.
        Write-Log "[$($OTSheet.Name)] OT source not configured yet - worksheet will be empty" 'WARN'
        $r = New-SiteResult $OTSheet $OTHeaders; $r.VCenter = 'not configured'; $results.Add($r)
    }
}

$saved = $false
try {
    $OutputPath = @(Export-InventoryWorkbook $results $OutputPath ([bool]$SampleData))[-1]
    $saved = $true
    Write-Log "Workbook saved: $OutputPath" 'OK'
}
catch {
    Write-Log "Excel export FAILED: $($_.Exception.Message) $($_.ScriptStackTrace)" 'ERROR'
    $csv = [IO.Path]::ChangeExtension($OutputPath, '.fallback.csv')
    $results | ForEach-Object { $s = $_.Site; $_.Rows | Select-Object @{n = 'Sheet'; e = { $s } }, * } | Export-Csv $csv -NoTypeInformation
    Write-Log "Raw data saved to fallback CSV: $csv" 'WARN'
}
finally {
    foreach ($s in $script:OpenedSessions) {
        try { Disconnect-VIServer -Server $s -Confirm:$false -ErrorAction Stop; Write-Log "Closed session opened by this script: $($s.Name)" }
        catch { }
    }
}

# ---------- Summary ----------
$summary = foreach ($r in $results) {
    [pscustomobject]@{
        Site           = $r.Site
        vCenter        = $r.VCenter
        'VMs Collected' = $r.Rows.Count
        'Queries OK'   = $r.QueryOk
        'Queries Failed' = $r.QueryFailed
        'VM Read Errors' = $r.VmErrors
        Status         = $(if ($r.Manual) { 'MANUAL' } elseif ($r.QueryFailed -gt 0 -and $r.QueryOk -eq 0) { 'FAILED' } elseif ($r.QueryFailed -gt 0 -or $r.VmErrors -gt 0) { 'PARTIAL' } else { 'OK' })
    }
}
Write-Host ''
Write-Host '==================== SUMMARY ====================' -ForegroundColor Cyan
$summary | Format-Table -AutoSize | Out-String -Width 200 | Write-Host
Write-Host ("Total VMs     : {0}" -f (($results | ForEach-Object { $_.Rows.Count } | Measure-Object -Sum).Sum))
Write-Host ("Excel file    : {0}" -f $(if ($saved) { $OutputPath } else { 'NOT CREATED - see log' })) -ForegroundColor $(if ($saved) { 'Green' } else { 'Red' })
Write-Host ("Log file      : {0}" -f $LogPath)
Write-Host '=================================================' -ForegroundColor Cyan
#endregion

