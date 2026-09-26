# Export-AMCInventoryMerge.ps1
# Collects the AMC VM inventory and adds it as a new "AMC" tab to the master Excel file,
# in the same format. The other tabs are not changed. Excel does NOT need to be installed.
# Read-only against vCenter. A backup of the Excel file is saved first.
# Run:  .\Export-AMCInventoryMerge.ps1
param(
    [string]$InputFile = 'C:\Temp\master.xlsx',     # <-- input Excel file (the AMC tab is added to this file)
    [switch]$SampleData,
    [switch]$NoConnect,
    [switch]$IncludeTemplates
)

#region ======================= AMC CONFIGURATION (edit here) =======================
$Sites = @(
    [pscustomobject]@{ Name = 'AMC'; VCenter = @('10.60.1.20'); Datacenter = @(); TabColor = 'BF8F00' }   # <-- AMC vCenter IP
)
$InsertBefore = 'OT'     # AMC tab goes before this tab; last if not found
#endregion ==========================================================================
#region ---------- Style constants ----------
$Style = @{
    NotAvailable  = 'N/A'      # shown when no IP could be read
    UnknownOS     = 'Unknown'  # shown when no OS could be read
}
$Headers = '#', 'VM NAME', 'IP', 'OS', 'Power Status', 'Implementor'
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

function Get-SafeSheetName([string]$Name) {
    $n = ($Name -replace '[\[\]\:\*\?\/\\]', '-').Trim()
    if ($n.Length -gt 31) { $n = $n.Substring(0, 31) }
    if (-not $n) { $n = 'Site' }
    return $n
}

function New-SiteResult($Site) {
    [pscustomobject]@{
        Site        = $Site.Name
        VCenter     = (@($Site.VCenter) -join ', ')
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

#region ---------- Add the AMC tab to the .xlsx (no Excel needed) ----------
# The new tab is a copy of an existing site tab (same columns), so column widths, header,
# table style, borders and conditional formatting are identical. Other tabs are not changed.
Add-Type -AssemblyName System.IO.Compression
Add-Type -AssemblyName System.IO.Compression.FileSystem

$NsMain = 'http://schemas.openxmlformats.org/spreadsheetml/2006/main'
$NsRel  = 'http://schemas.openxmlformats.org/officeDocument/2006/relationships'
$NsPkg  = 'http://schemas.openxmlformats.org/package/2006/relationships'
$NsCT   = 'http://schemas.openxmlformats.org/package/2006/content-types'

function Read-ZipXml($Zip, [string]$Name) {
    $e = $Zip.GetEntry($Name)
    if (-not $e) { return $null }
    $sr = New-Object IO.StreamReader($e.Open())
    try {
        $x = New-Object Xml.XmlDocument
        $x.PreserveWhitespace = $true
        $x.LoadXml($sr.ReadToEnd())
        return $x
    }
    finally { $sr.Dispose() }
}

function Write-ZipXml($Zip, [string]$Name, [xml]$Doc) {
    $old = $Zip.GetEntry($Name)
    if ($old) { $old.Delete() }
    $st = $Zip.CreateEntry($Name).Open()
    try {
        $w = New-Object IO.StreamWriter($st, (New-Object Text.UTF8Encoding $false))
        $Doc.Save($w)
        $w.Flush()
    }
    finally { $st.Dispose() }
}

function New-Ns($Doc) {
    $ns = New-Object Xml.XmlNamespaceManager($Doc.NameTable)
    $ns.AddNamespace('m', $NsMain); $ns.AddNamespace('r', $NsRel); $ns.AddNamespace('p', $NsPkg); $ns.AddNamespace('ct', $NsCT)
    return , $ns
}

function Resolve-PartName([string]$BaseDir, [string]$Target) {
    if ($Target.StartsWith('/')) { return $Target.TrimStart('/') }
    $parts = New-Object System.Collections.Generic.List[string]
    foreach ($p in (($BaseDir + '/' + $Target) -split '/')) {
        if ($p -eq '..') { if ($parts.Count) { $parts.RemoveAt($parts.Count - 1) } }
        elseif ($p -and $p -ne '.') { $parts.Add($p) }
    }
    return ($parts -join '/')
}

function Get-RelsPartName([string]$Part) {
    $i = $Part.LastIndexOf('/')
    return $Part.Substring(0, $i) + '/_rels/' + $Part.Substring($i + 1) + '.rels'
}

function Set-CellValue($Doc, $Cell, $Value, [switch]$Number) {
    while ($Cell.HasChildNodes) { [void]$Cell.RemoveChild($Cell.FirstChild) }
    $Cell.RemoveAttribute('t')
    if ($null -eq $Value -or [string]$Value -eq '') { return }
    if ($Number) {
        $v = $Doc.CreateElement('v', $NsMain); $v.InnerText = [string]$Value
        [void]$Cell.AppendChild($v)
        return
    }
    $Cell.SetAttribute('t', 'inlineStr')
    $is = $Doc.CreateElement('is', $NsMain)
    $t = $Doc.CreateElement('t', $NsMain); $t.InnerText = [string]$Value
    [void]$is.AppendChild($t); [void]$Cell.AppendChild($is)
}

function Reset-Uids($Doc) {
    # xr:uid / xr3:uid must be unique per workbook - give the copies new ones
    foreach ($el in @($Doc.SelectNodes('//*'))) {
        foreach ($a in @($el.Attributes)) {
            if ($a.LocalName -eq 'uid' -and $a.NamespaceURI -like 'http://schemas.microsoft.com/office/spreadsheetml/*') {
                $a.Value = '{' + [guid]::NewGuid().ToString().ToUpper() + '}'
            }
        }
    }
}

function Add-InventorySheet([string]$Path, $Result, [string]$SheetName, [string]$TabColor, [string]$InsertBefore, [string]$Info) {
    $zip = [IO.Compression.ZipFile]::Open($Path, [IO.Compression.ZipArchiveMode]::Update)
    try {
        $wb = Read-ZipXml $zip 'xl/workbook.xml';          $wns = New-Ns $wb
        $wbRels = Read-ZipXml $zip 'xl/_rels/workbook.xml.rels'; $rns = New-Ns $wbRels
        $sheets = @($wb.SelectNodes('/m:workbook/m:sheets/m:sheet', $wns))
        if (@($sheets | Where-Object { $_.GetAttribute('name') -eq $SheetName }).Count -gt 0) {
            throw "The workbook already has a '$SheetName' tab. Run it on the file without the $SheetName tab."
        }

        # --- 1. Find a site tab with the same columns to copy the format from
        $tpl = $null
        foreach ($s in $sheets) {
            $rid = $s.GetAttribute('id', $NsRel)
            $rel = $wbRels.SelectSingleNode("/p:Relationships/p:Relationship[@Id='$rid']", $rns)
            if (-not $rel) { continue }
            $sheetPart = Resolve-PartName 'xl' $rel.GetAttribute('Target')
            $sr = Read-ZipXml $zip (Get-RelsPartName $sheetPart)
            if (-not $sr) { continue }
            $trel = @($sr.SelectNodes('/p:Relationships/p:Relationship', (New-Ns $sr)) | Where-Object { $_.GetAttribute('Type') -like '*/table' })
            if ($trel.Count -eq 0) { continue }
            $tablePart = Resolve-PartName ($sheetPart.Substring(0, $sheetPart.LastIndexOf('/'))) $trel[0].GetAttribute('Target')
            $tx = Read-ZipXml $zip $tablePart
            $cols = @($tx.SelectNodes('//m:tableColumn', (New-Ns $tx)) | ForEach-Object { $_.GetAttribute('name') }) -join '|'
            if ($cols -eq ($Headers -join '|')) {
                $tpl = @{ Name = $s.GetAttribute('name'); SheetPart = $sheetPart; TablePart = $tablePart; Index = [array]::IndexOf($sheets, $s) }
                break
            }
        }
        if (-not $tpl) { throw "No tab with the columns '$($Headers -join ', ')' was found to copy the format from." }
        Write-Log "Copying the format of tab '$($tpl.Name)'"

        # --- 2. New part names / ids
        $names = @($zip.Entries | ForEach-Object { $_.FullName })
        $sheetNo = 1 + (@($names | Where-Object { $_ -match '^xl/worksheets/sheet(\d+)\.xml$' } | ForEach-Object { [int]$Matches[1] }) | Measure-Object -Maximum).Maximum
        $tableNo = 1 + (@($names | Where-Object { $_ -match '^xl/tables/table(\d+)\.xml$' } | ForEach-Object { [int]$Matches[1] }) | Measure-Object -Maximum).Maximum
        $tableIds = @(); $tableNames = @()
        foreach ($n in @($names | Where-Object { $_ -match '^xl/tables/table\d+\.xml$' })) {
            $t = Read-ZipXml $zip $n
            $tableIds += [int]$t.DocumentElement.GetAttribute('id')
            $tableNames += $t.DocumentElement.GetAttribute('name')
        }
        $tableId = 1 + ($tableIds | Measure-Object -Maximum).Maximum
        $tableName = 'tbl_' + ($SheetName -replace '[^A-Za-z0-9_]', '_')
        while ($tableNames -contains $tableName) { $tableName += '_1' }
        $newSheetPart = "xl/worksheets/sheet$sheetNo.xml"
        $newTablePart = "xl/tables/table$tableNo.xml"

        # --- 3. Styles: title banner fill in the AMC colour
        $st = Read-ZipXml $zip 'xl/styles.xml'; $sns = New-Ns $st
        $tx = Read-ZipXml $zip $tpl.TablePart; $tns = New-Ns $tx
        $sx = Read-ZipXml $zip $tpl.SheetPart; $xns = New-Ns $sx
        $ref = $tx.DocumentElement.GetAttribute('ref')                     # e.g. A3:F12
        if ($ref -notmatch '^([A-Z]+)(\d+):([A-Z]+)(\d+)$') { throw "Unexpected table range '$ref' in tab '$($tpl.Name)'" }
        $c1 = $Matches[1]; $hdr = [int]$Matches[2]; $c2 = $Matches[3]; $tplLast = [int]$Matches[4]

        $rowsTpl = @{}
        foreach ($r in @($sx.SelectNodes('/m:worksheet/m:sheetData/m:row', $xns))) { $rowsTpl[[int]$r.GetAttribute('r')] = $r }
        $titleStyle = $null
        if ($rowsTpl.ContainsKey(1)) { $a1 = $rowsTpl[1].SelectSingleNode('m:c', $xns); if ($a1) { $titleStyle = $a1.GetAttribute('s') } }
        $newTitleStyle = $null
        if ($titleStyle -ne $null -and $titleStyle -ne '') {
            $fills = $st.SelectSingleNode('/m:styleSheet/m:fills', $sns)
            $fill = $st.CreateElement('fill', $NsMain)
            $pf = $st.CreateElement('patternFill', $NsMain); $pf.SetAttribute('patternType', 'solid')
            $fg = $st.CreateElement('fgColor', $NsMain); $fg.SetAttribute('rgb', 'FF' + $TabColor.ToUpper())
            $bg = $st.CreateElement('bgColor', $NsMain); $bg.SetAttribute('indexed', '64')
            [void]$pf.AppendChild($fg); [void]$pf.AppendChild($bg); [void]$fill.AppendChild($pf); [void]$fills.AppendChild($fill)
            $fillId = $fills.SelectNodes('m:fill', $sns).Count - 1
            $fills.SetAttribute('count', [string]($fillId + 1))
            $xfs = $st.SelectSingleNode('/m:styleSheet/m:cellXfs', $sns)
            $xf = $xfs.SelectNodes('m:xf', $sns)[[int]$titleStyle].CloneNode($true)
            $xf.SetAttribute('fillId', [string]$fillId); $xf.SetAttribute('applyFill', '1')
            [void]$xfs.AppendChild($xf)
            $newTitleStyle = [string]($xfs.SelectNodes('m:xf', $sns).Count - 1)
            $xfs.SetAttribute('count', [string]([int]$newTitleStyle + 1))
        }

        # --- 4. Sheet: copy the template, replace title/info/data
        $n = $Result.Rows.Count
        $last = $hdr + [Math]::Max($n, 1)
        $tc = $sx.SelectSingleNode('/m:worksheet/m:sheetPr/m:tabColor', $xns)
        if ($tc) { foreach ($a in 'theme', 'tint', 'indexed') { $tc.RemoveAttribute($a) }; $tc.SetAttribute('rgb', 'FF' + $TabColor.ToUpper()) }
        foreach ($sv in @($sx.SelectNodes('//m:sheetView', $xns))) { $sv.RemoveAttribute('tabSelected') }
        $dim = $sx.SelectSingleNode('/m:worksheet/m:dimension', $xns)
        if ($dim) { $dim.SetAttribute('ref', "A1:$c2$last") }

        $sd = $sx.SelectSingleNode('/m:worksheet/m:sheetData', $xns)
        $rowMid = $rowsTpl[$hdr + 1]; $rowEnd = $rowsTpl[$tplLast]
        if (-not $rowMid -or -not $rowEnd -or -not $rowsTpl.ContainsKey($hdr)) { throw "Tab '$($tpl.Name)' has no data rows to copy the format from." }
        $keep = @()
        foreach ($r in 1..$hdr) { if ($rowsTpl.ContainsKey($r)) { $keep += $rowsTpl[$r] } }
        while ($sd.HasChildNodes) { [void]$sd.RemoveChild($sd.FirstChild) }
        foreach ($r in $keep) {
            $rr = [int]$r.GetAttribute('r')
            if ($rr -eq 1 -or $rr -eq 2) {
                $cells = @($r.SelectNodes('m:c', $xns))
                for ($i = 0; $i -lt $cells.Count; $i++) {
                    if ($rr -eq 1 -and $newTitleStyle) { $cells[$i].SetAttribute('s', $newTitleStyle) }
                    if ($i -eq 0) { Set-CellValue $sx $cells[$i] $(if ($rr -eq 1) { "$SheetName  -  VM Inventory" } else { $Info }) }
                }
            }
            [void]$sd.AppendChild($r)
        }
        $field = @('#', 'Name', 'IP', 'OS', 'Power', 'Implementor')
        for ($i = 0; $i -lt [Math]::Max($n, 1); $i++) {
            $rowNo = $hdr + 1 + $i
            $src = $(if ($rowNo -eq $last) { $rowEnd } else { $rowMid })
            $row = $src.CloneNode($true)
            $row.SetAttribute('r', [string]$rowNo)
            $cells = @($row.SelectNodes('m:c', $xns))
            for ($c = 0; $c -lt $cells.Count; $c++) {
                $col = ($cells[$c].GetAttribute('r') -replace '\d', '')
                $cells[$c].SetAttribute('r', "$col$rowNo")
                if ($n -eq 0) { Set-CellValue $sx $cells[$c] $(if ($c -eq 1 -and $Result.QueryFailed -gt 0) { 'No data - vCenter query failed' } elseif ($c -eq 1) { 'No VMs found' } else { '' }); continue }
                $vm = $Result.Rows[$i]
                if ($c -eq 0) { Set-CellValue $sx $cells[$c] ($i + 1) -Number }
                elseif ($c -lt $field.Count -and $field[$c] -ne 'Implementor') { Set-CellValue $sx $cells[$c] $vm.($field[$c]) }
                else { Set-CellValue $sx $cells[$c] '' }
            }
            [void]$sd.AppendChild($row)
        }
        foreach ($cf in @($sx.SelectNodes('/m:worksheet/m:conditionalFormatting', $xns))) {
            $cf.SetAttribute('sqref', ($cf.GetAttribute('sqref') -replace '([A-Z]+)(\d+):([A-Z]+)\d+', ('${1}${2}:${3}' + $last)))
        }
        $ps = $sx.SelectSingleNode('/m:worksheet/m:pageSetup', $xns)
        if ($ps) { $ps.RemoveAttribute('id', $NsRel) }                      # printer settings are not copied
        foreach ($tag in 'drawing', 'legacyDrawing', 'legacyDrawingHF', 'picture', 'hyperlinks') {
            foreach ($x in @($sx.SelectNodes("/m:worksheet/m:$tag", $xns))) { [void]$x.ParentNode.RemoveChild($x) }
        }
        $tps = $sx.SelectSingleNode('/m:worksheet/m:tableParts', $xns)
        foreach ($tp in @($tps.SelectNodes('m:tablePart', $xns))) { [void]$tps.RemoveChild($tp) }
        $tp = $sx.CreateElement('tablePart', $NsMain)
        $ra = $sx.CreateAttribute('r', 'id', $NsRel); $ra.Value = 'rId1'; [void]$tp.Attributes.Append($ra)
        [void]$tps.AppendChild($tp); $tps.SetAttribute('count', '1')
        Reset-Uids $sx

        # --- 5. Table
        $tx.DocumentElement.SetAttribute('id', [string]$tableId)
        $tx.DocumentElement.SetAttribute('name', $tableName)
        $tx.DocumentElement.SetAttribute('displayName', $tableName)
        $tx.DocumentElement.SetAttribute('ref', "$c1${hdr}:$c2$last")
        $af = $tx.SelectSingleNode('/m:table/m:autoFilter', $tns)
        if ($af) { $af.SetAttribute('ref', "$c1${hdr}:$c2$last") }
        Reset-Uids $tx

        $srels = New-Object Xml.XmlDocument
        $srels.LoadXml("<?xml version=`"1.0`" encoding=`"UTF-8`" standalone=`"yes`"?><Relationships xmlns=`"$NsPkg`"><Relationship Id=`"rId1`" Type=`"http://schemas.openxmlformats.org/officeDocument/2006/relationships/table`" Target=`"../tables/table$tableNo.xml`"/></Relationships>")

        # --- 6. Workbook: register the new tab before $InsertBefore
        $ids = @($wbRels.SelectNodes('/p:Relationships/p:Relationship', $rns) | ForEach-Object { $_.GetAttribute('Id') } | Where-Object { $_ -match '^rId(\d+)$' } | ForEach-Object { [int]$Matches[1] })
        $newRid = 'rId' + (1 + ($ids | Measure-Object -Maximum).Maximum)
        $rel = $wbRels.CreateElement('Relationship', $NsPkg)
        $rel.SetAttribute('Id', $newRid)
        $rel.SetAttribute('Type', 'http://schemas.openxmlformats.org/officeDocument/2006/relationships/worksheet')
        $rel.SetAttribute('Target', "worksheets/sheet$sheetNo.xml")
        [void]$wbRels.DocumentElement.AppendChild($rel)

        $sheetsNode = $wb.SelectSingleNode('/m:workbook/m:sheets', $wns)
        $newSheet = $wb.CreateElement('sheet', $NsMain)
        $newSheet.SetAttribute('name', $SheetName)
        $newSheet.SetAttribute('sheetId', [string](1 + (@($sheets | ForEach-Object { [int]$_.GetAttribute('sheetId') }) | Measure-Object -Maximum).Maximum))
        $ra = $wb.CreateAttribute('r', 'id', $NsRel); $ra.Value = $newRid; [void]$newSheet.Attributes.Append($ra)
        $before = @($sheets | Where-Object { $_.GetAttribute('name') -eq $InsertBefore })
        if ($before.Count -gt 0) { [void]$sheetsNode.InsertBefore($newSheet, $before[0]); $pos = [array]::IndexOf($sheets, $before[0]) }
        else { [void]$sheetsNode.AppendChild($newSheet); $pos = $sheets.Count }

        # sheet positions after the new tab move up by one
        foreach ($dn in @($wb.SelectNodes('/m:workbook/m:definedNames/m:definedName[@localSheetId]', $wns))) {
            $l = [int]$dn.GetAttribute('localSheetId'); if ($l -ge $pos) { $dn.SetAttribute('localSheetId', [string]($l + 1)) }
        }
        foreach ($bv in @($wb.SelectNodes('/m:workbook/m:bookViews/m:workbookView', $wns))) {
            foreach ($a in 'activeTab', 'firstSheet') { if ($bv.HasAttribute($a) -and [int]$bv.GetAttribute($a) -ge $pos) { $bv.SetAttribute($a, [string]([int]$bv.GetAttribute($a) + 1)) } }
        }
        # print titles (repeat header row) for the new tab
        $dns = $wb.SelectSingleNode('/m:workbook/m:definedNames', $wns)
        if ($dns) {
            $dn = $wb.CreateElement('definedName', $NsMain)
            $dn.SetAttribute('name', '_xlnm.Print_Titles'); $dn.SetAttribute('localSheetId', [string]$pos)
            $dn.InnerText = "'" + $SheetName.Replace("'", "''") + "'!`$${hdr}:`$${hdr}"
            [void]$dns.AppendChild($dn)
        }

        # --- 7. Content types
        $ct = Read-ZipXml $zip '[Content_Types].xml'
        foreach ($pair in @(@("/$newSheetPart", 'application/vnd.openxmlformats-officedocument.spreadsheetml.worksheet+xml'),
                            @("/$newTablePart", 'application/vnd.openxmlformats-officedocument.spreadsheetml.table+xml'))) {
            $o = $ct.CreateElement('Override', $NsCT); $o.SetAttribute('PartName', $pair[0]); $o.SetAttribute('ContentType', $pair[1])
            [void]$ct.DocumentElement.AppendChild($o)
        }

        # --- 8. docProps/app.xml sheet list (informational; best effort)
        try {
            $app = Read-ZipXml $zip 'docProps/app.xml'
            if ($app) {
                $ans = New-Object Xml.XmlNamespaceManager($app.NameTable)
                $ans.AddNamespace('ep', 'http://schemas.openxmlformats.org/officeDocument/2006/extended-properties')
                $ans.AddNamespace('vt', 'http://schemas.openxmlformats.org/officeDocument/2006/docPropsVTypes')
                $vars = @($app.SelectNodes('//ep:HeadingPairs/vt:vector/vt:variant', $ans))
                $parts = $app.SelectSingleNode('//ep:TitlesOfParts/vt:vector', $ans)
                $lp = @($parts.SelectNodes('vt:lpstr', $ans))
                for ($i = 0; $i -lt $vars.Count - 1; $i++) {
                    if ($vars[$i].InnerText -eq 'Worksheets') {
                        $cnt = $vars[$i + 1].SelectSingleNode('vt:i4', $ans); $cnt.InnerText = [string]([int]$cnt.InnerText + 1)
                        $el = $app.CreateElement('vt', 'lpstr', 'http://schemas.openxmlformats.org/officeDocument/2006/docPropsVTypes'); $el.InnerText = $SheetName
                        if ($pos -lt $lp.Count) { [void]$parts.InsertBefore($el, $lp[$pos]) } else { [void]$parts.AppendChild($el) }
                        $parts.SetAttribute('size', [string]($parts.SelectNodes('vt:lpstr', $ans).Count))
                    }
                }
                Write-ZipXml $zip 'docProps/app.xml' $app
            }
        }
        catch { }

        Write-ZipXml $zip 'xl/styles.xml' $st
        Write-ZipXml $zip $newSheetPart $sx
        Write-ZipXml $zip (Get-RelsPartName $newSheetPart) $srels
        Write-ZipXml $zip $newTablePart $tx
        Write-ZipXml $zip 'xl/workbook.xml' $wb
        Write-ZipXml $zip 'xl/_rels/workbook.xml.rels' $wbRels
        Write-ZipXml $zip '[Content_Types].xml' $ct
    }
    finally { $zip.Dispose() }
}
#endregion

#region ---------- Main ----------
$InputFile = [IO.Path]::GetFullPath($InputFile)
$LogPath = Join-Path (Split-Path $InputFile -Parent) ([IO.Path]::GetFileNameWithoutExtension($InputFile) + '_AMC.log')
Write-Log "===== AMC inventory started ($(if ($SampleData) { 'SAMPLE mode' } else { 'LIVE, read-only' })) ====="
if (-not (Test-Path $InputFile)) { Write-Log "Excel file not found: $InputFile" 'ERROR'; return }

# 1. Collect AMC
if (-not $SampleData) {
    try { Import-Module VMware.VimAutomation.Core -ErrorAction Stop -WarningAction SilentlyContinue }
    catch { Write-Log 'VMware PowerCLI is not installed. Run:  Install-Module VMware.PowerCLI -Scope CurrentUser' 'ERROR'; return }
}
$site = $Sites[0]
try { $result = $(if ($SampleData) { Get-SampleAMCInventory $site } else { Get-SiteInventory $site }) }
catch { Write-Log "[$($site.Name)] Unexpected failure: $($_.Exception.Message)" 'ERROR'; $result = New-SiteResult $site; $result.QueryFailed++ }
foreach ($s in $script:OpenedSessions) { try { Disconnect-VIServer -Server $s -Confirm:$false -ErrorAction Stop } catch { } }

$on = @($result.Rows | Where-Object Power -eq 'PoweredOn').Count
$off = @($result.Rows | Where-Object Power -eq 'PoweredOff').Count
$sus = @($result.Rows | Where-Object Power -eq 'Suspended').Count
$src = if ($SampleData) { 'SAMPLE DATA (fictitious VMs)' } else { "vCenter: $($result.VCenter)" }
$info = "{0}  |  Generated {1}  |  {2} VMs: {3} on, {4} off, {5} suspended  |  Yellow = manual entry, orange name = duplicate" -f `
    $src, (Get-Date -Format 'yyyy-MM-dd HH:mm'), $result.Rows.Count, $on, $off, $sus
if ($result.QueryFailed -gt 0) { $info = "COLLECTION INCOMPLETE - vCenter query failed (see log)   |   " + $info }

# 2. Add the AMC tab (work on a temp copy; the original is replaced only when everything succeeded)
$backup = Join-Path (Split-Path $InputFile -Parent) ('{0}_backup_{1}.xlsx' -f [IO.Path]::GetFileNameWithoutExtension($InputFile), (Get-Date -Format 'yyyyMMdd_HHmmss'))
$work = Join-Path $env:TEMP ('AMC_work_{0}.xlsx' -f [guid]::NewGuid().ToString('N'))
$saved = $false
try {
    Copy-Item $InputFile $backup -Force
    Copy-Item $InputFile $work -Force
    Add-InventorySheet $work $result (Get-SafeSheetName $site.Name) $site.TabColor $InsertBefore $info
    try { Copy-Item $work $InputFile -Force -ErrorAction Stop }
    catch { throw "Cannot overwrite $InputFile - is it open in Excel? Close it and run again." }
    $saved = $true
    Write-Log "AMC tab added: $InputFile" 'OK'
}
catch { Write-Log "Could not update the Excel file: $($_.Exception.Message)" 'ERROR' }
finally { Remove-Item $work -Force -ErrorAction SilentlyContinue }

Write-Host ''
Write-Host ("AMC VMs       : {0}  (failed queries: {1})" -f $result.Rows.Count, $result.QueryFailed)
Write-Host ("Excel file    : {0}" -f $(if ($saved) { $InputFile } else { 'NOT CHANGED - see log' })) -ForegroundColor $(if ($saved) { 'Green' } else { 'Red' })
Write-Host ("Backup        : {0}" -f $backup)
Write-Host ("Log           : {0}" -f $LogPath)
#endregion