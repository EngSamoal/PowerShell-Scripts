<#
.SYNOPSIS
    Collects the AMC VM inventory and adds it as a new "AMC" tab to an existing master workbook,
    in the same format. The other tabs are not changed.

.EXAMPLE
    .\Export-AMCInventoryMerge.ps1 -InputFile 'C:\Temp\VM_Master_Inventory.xlsx'

.NOTES
    READ-ONLY against vCenter. A backup copy of the workbook is saved before it is changed.
    Requires Microsoft Excel and VMware PowerCLI (not needed with -SampleData).
#>
[CmdletBinding()]
param(
    [string]$InputFile = 'C:\Temp\master.xlsx',     # <-- input Excel file (the AMC tab is added to this file)
    [switch]$SampleData,
    [switch]$NoConnect,
    [switch]$IncludeTemplates
)

#region ======================= AMC CONFIGURATION (edit here) =======================
$Sites = @(
    [pscustomobject]@{ Name = 'AMC'; VCenter = @('10.60.1.20'); Datacenter = @(); TabColor = 'BF8F00' }
)
$InsertBefore = 'OT'     # AMC tab goes before this tab; last if not found
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
$LogPath = Join-Path $env:TEMP 'AMC_Inventory.log'

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
function Get-SampleAMCInventory($Site) {
    $result = New-SiteResult $Site
    $result.QueryOk = 1
    $w19 = 'Microsoft Windows Server 2019 (64-bit)'
    $w22 = 'Microsoft Windows Server 2022 (64-bit)'
    $w25 = 'Microsoft Windows Server 2025 (64-bit)'
    $r9 = 'Red Hat Enterprise Linux 9 (64-bit)'
    $set = @(
        @('AMC-DC01', '10.60.10.11', $w25, 'PoweredOn'),
        @('AMC-DC02', '10.60.10.12', $w25, 'PoweredOn'),
        @('AMC-SQL01', '10.60.20.21', $w22, 'PoweredOn'),
        @('AMC-APP01', '10.60.30.31', $w22, 'PoweredOn'),
        @('AMC-APP02', $null, $w19, 'PoweredOff'),
        @('AMC-WEB01', '10.60.30.41', $r9, 'PoweredOn'),
        @('AMC-FS01', '10.60.40.10', $w19, 'PoweredOn'),
        @('AMC-BKP01', $null, $w22, 'Suspended'),
        @('AMC-OLD01', $null, $null, 'PoweredOff')
    )
    foreach ($s in $set) {
        $result.Rows.Add([pscustomobject]@{
                Name  = $s[0]
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
#endregion

#region ---------- Excel formatting (same as the main script) ----------
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
    $field = @{ 'Site' = 'Site'; 'VM NAME' = 'Name'; 'IP' = 'IP'; 'OS' = 'OS'; 'Power Status' = 'Power'; 'Implementor' = 'Implementor' }

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
#endregion

#region ---------- Main ----------
$InputFile = [IO.Path]::GetFullPath($InputFile)
$LogPath = Join-Path (Split-Path $InputFile -Parent) ([IO.Path]::GetFileNameWithoutExtension($InputFile) + '_AMC.log')
Write-Log "===== AMC inventory started ($(if ($SampleData) { 'SAMPLE mode' } else { 'LIVE, read-only' })) ====="
if (-not (Test-Path $InputFile)) { Write-Log "Workbook not found: $InputFile" 'ERROR'; return }

# 1. Collect AMC
if (-not $SampleData) {
    try { Import-Module VMware.VimAutomation.Core -ErrorAction Stop -WarningAction SilentlyContinue }
    catch { Write-Log 'VMware PowerCLI is not installed. Run:  Install-Module VMware.PowerCLI -Scope CurrentUser' 'ERROR'; return }
}
$site = $Sites[0]
try { $result = $(if ($SampleData) { Get-SampleAMCInventory $site } else { Get-SiteInventory $site }) }
catch { Write-Log "[$($site.Name)] Unexpected failure: $($_.Exception.Message)" 'ERROR'; $result = New-SiteResult $site; $result.QueryFailed++ }

# 2. Add the AMC tab to the workbook
$backup = Join-Path (Split-Path $InputFile -Parent) ('{0}_backup_{1}.xlsx' -f [IO.Path]::GetFileNameWithoutExtension($InputFile), (Get-Date -Format 'yyyyMMdd_HHmmss'))
$saved = $false
$excel = $null; $wb = $null
try {
    Copy-Item $InputFile $backup -Force
    $excel = New-Object -ComObject Excel.Application
    $excel.Visible = $false
    $excel.DisplayAlerts = $false
    $excel.ScreenUpdating = $false
    $wb = $excel.Workbooks.Open($InputFile)
    if ($wb.ReadOnly) { throw 'The workbook is open in Excel - close it and run again.' }

    $names = @($wb.Worksheets | ForEach-Object { $_.Name })
    $sheetName = Get-SafeSheetName $site.Name
    if ($names -contains $sheetName) {
        Write-Log "Replacing existing '$sheetName' tab" 'WARN'
        $wb.Worksheets.Item($sheetName).Delete()
        $names = @($wb.Worksheets | ForEach-Object { $_.Name })
    }
    if ($names -contains $InsertBefore) { $ws = $wb.Worksheets.Add($wb.Worksheets.Item($InsertBefore)) }
    else { $ws = $wb.Worksheets.Add([Reflection.Missing]::Value, $wb.Worksheets.Item($wb.Worksheets.Count)) }

    Format-SiteSheet $excel $ws $result ([bool]$SampleData)
    $wb.Worksheets.Item(1).Activate()
    $excel.ScreenUpdating = $true
    $wb.Save()
    $wb.Close($false)
    $saved = $true
    Write-Log "AMC tab added: $InputFile" 'OK'
}
catch { Write-Log "Could not update the workbook: $($_.Exception.Message)" 'ERROR' }
finally {
    if ($excel) { $excel.Quit() }
    foreach ($o in @($wb, $excel)) { if ($o) { [void][Runtime.InteropServices.Marshal]::ReleaseComObject($o) } }
    [GC]::Collect(); [GC]::WaitForPendingFinalizers()
    foreach ($s in $script:OpenedSessions) { try { Disconnect-VIServer -Server $s -Confirm:$false -ErrorAction Stop } catch { } }
}

Write-Host ''
Write-Host ("AMC VMs       : {0}  (failed queries: {1})" -f $result.Rows.Count, $result.QueryFailed)
Write-Host ("Workbook      : {0}" -f $(if ($saved) { $InputFile } else { 'NOT CHANGED - see log' })) -ForegroundColor $(if ($saved) { 'Green' } else { 'Red' })
Write-Host ("Backup        : {0}" -f $backup)
Write-Host ("Log           : {0}" -f $LogPath)
#endregion
