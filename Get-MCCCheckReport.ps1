<#
    Get-MCCCheckReport.ps1
    Read-only check for Microsoft Connected Cache (MCC), installed apps and
    running third-party services on target VMs via VMware Guest Operations.
    Output: colored console summary + HTML report in C:\temp
    Requirements: PowerCLI, Connect-VIServer to the relevant vCenter(s),
                  guest local admin credential, VMware Tools running.
    Usage:        .\Get-MCCCheckReport.ps1 -VMName 'VM01','VM02'
#>
param(
    [Parameter(Mandatory=$true)][string[]]$VMName,
    [string]$OutDir = 'C:\temp'
)

$script = @'
$k='mcc|connected cache|doinc|deliveryoptimization|iotedge|aziot'
$reg  = 'HKLM:\Software\Microsoft\Windows\CurrentVersion\Uninstall\*','HKLM:\Software\WOW6432Node\Microsoft\Windows\CurrentVersion\Uninstall\*'
$all  = Get-ItemProperty $reg -ErrorAction SilentlyContinue | Where-Object {$_.DisplayName -and $_.SystemComponent -ne 1}
$svc  = Get-CimInstance Win32_Service | Where-Object {$_.Name -match $k -or $_.DisplayName -match $k}
$apps = $all | Where-Object {$_.DisplayName -match $k}
$task = Get-ScheduledTask -ErrorAction SilentlyContinue | Where-Object {$_.TaskName -match $k}
$fold = Get-ChildItem C:\ -Directory -ErrorAction SilentlyContinue | Where-Object {$_.Name -match 'mcc|doinc'}
$wsl  = Get-Service WslService,LxssManager -ErrorAction SilentlyContinue
$iis  = Get-Service W3SVC -ErrorAction SilentlyContinue
$p80  = Get-NetTCPConnection -LocalPort 80 -State Listen -ErrorAction SilentlyContinue | Select-Object -First 1
$agt  = Get-CimInstance Win32_Service | Where-Object {$_.State -eq 'Running' -and $_.PathName -notmatch '\\Windows\\'} | Sort-Object DisplayName
$r = [ordered]@{
  Host     = $env:COMPUTERNAME
  MCCItems = (@($svc.Name) + @($apps.DisplayName) + @($task.TaskName) + @($fold.FullName) | Where-Object {$_}) -join ', '
  WSL      = if ($wsl) {'Installed'} else {'Not installed'}
  IIS      = if ($iis) {'Installed'} else {'Not installed'}
  Port80   = if ($p80) {'In use by ' + (Get-Process -Id $p80.OwningProcess).ProcessName} else {'Free'}
  Agents   = ($agt.DisplayName) -join ' | '
  Apps     = ($all | Sort-Object DisplayName -Unique | ForEach-Object { "$($_.DisplayName)~$($_.DisplayVersion)" }) -join ' | '
}
$r | ConvertTo-Json -Compress
'@

$vms    = $VMName
$outDir = $OutDir
$cred   = Get-Credential

$results = foreach ($vm in $vms) {
  try {
    $out = (Invoke-VMScript -VM $vm -ScriptText $script -ScriptType Powershell -GuestCredential $cred -ErrorAction Stop -WarningAction SilentlyContinue).ScriptOutput
    $j   = ($out -split "`n" | Where-Object {$_ -match '^\s*\{'} | Select-Object -Last 1) | ConvertFrom-Json
    [pscustomobject]@{
      'VM'=$vm; 'MCC Status'= if ($j.MCCItems) {'DETECTED: ' + $j.MCCItems} else {'Not installed'}
      'WSL'=$j.WSL; 'IIS'=$j.IIS; 'Port 80'=$j.Port80; 'Agents'=$j.Agents; 'Apps'=$j.Apps
    }
  } catch {
    [pscustomobject]@{'VM'=$vm;'MCC Status'='CHECK FAILED: ' + $_.Exception.Message;'WSL'='-';'IIS'='-';'Port 80'='-';'Agents'='-';'Apps'='-'}
  }
}

function Enc($s) { [System.Net.WebUtility]::HtmlEncode([string]$s) }
function Col($t) { if ($t -match '^(Not installed|Free)$') {'Green'} elseif ($t -match '^(DETECTED|CHECK FAILED)') {'Red'} else {'Yellow'} }
function Badge($t) { $c = if ($t -match '^(Not installed|Free)$') {'ok'} elseif ($t -match '^(DETECTED|CHECK FAILED)') {'bad'} else {'warn'}; "<span class='badge $c'>$(Enc $t)</span>" }
$bad = @($results | Where-Object {$_.'MCC Status' -ne 'Not installed'})

# ---------- Console ----------
Write-Host ""
Write-Host "==================== MCC CHECK SUMMARY ====================" -ForegroundColor Cyan
foreach ($r in $results) {
  Write-Host ""
  Write-Host ("  " + $r.VM) -ForegroundColor White
  foreach ($f in 'MCC Status','WSL','IIS','Port 80') {
    Write-Host ("    {0,-11}: " -f $f) -NoNewline -ForegroundColor Gray
    Write-Host $r.$f -ForegroundColor (Col $r.$f)
  }
}
Write-Host ""
if ($bad.Count -eq 0) { Write-Host "  RESULT: No MCC components found on any server." -ForegroundColor Green }
else { Write-Host "  RESULT: Review required on: $($bad.VM -join ', ')" -ForegroundColor Red }
foreach ($r in $results) {
  Write-Host "`n---------- $($r.VM) : Installed applications ----------" -ForegroundColor Cyan
  ($r.Apps -split ' \| ') | ForEach-Object { $p = $_ -split '~'; Write-Host ("  - {0,-60} {1}" -f $p[0],$p[1]) -ForegroundColor Gray }
  Write-Host "`n---------- $($r.VM) : Running third-party services ----------" -ForegroundColor Cyan
  ($r.Agents -split ' \| ') | ForEach-Object { Write-Host "  - $_" -ForegroundColor Gray }
}

# ---------- HTML ----------
$css = @'
<style>
body{font-family:Segoe UI,Arial,sans-serif;background:#f4f6f9;margin:0;color:#222}
.hdr{background:#1f4e79;color:#fff;padding:22px 30px}
.hdr h1{margin:0;font-size:24px}
.hdr p{margin:6px 0 0;opacity:.85;font-size:13px}
.wrap{padding:25px 30px}
.verdict{padding:16px 20px;border-radius:6px;font-size:18px;font-weight:600;margin-bottom:25px;color:#fff}
.vok{background:#2e7d32}.vbad{background:#c62828}
.card{background:#fff;border-radius:6px;box-shadow:0 1px 4px rgba(0,0,0,.12);padding:18px 22px;margin-bottom:25px}
h2{color:#1f4e79;margin:0 0 12px;font-size:19px}
h3{color:#1f4e79;font-size:15px;margin:18px 0 8px}
table{border-collapse:collapse;width:100%;font-size:13px}
th{background:#1f4e79;color:#fff;text-align:left;padding:8px 10px}
td{border-bottom:1px solid #e3e6ea;padding:7px 10px}
tr:nth-child(even) td{background:#f8f9fb}
.badge{padding:3px 10px;border-radius:12px;font-weight:600;font-size:12px;color:#fff}
.ok{background:#2e7d32}.warn{background:#ef8f00}.bad{background:#c62828}
.cnt{color:#666;font-weight:normal;font-size:13px}
.ft{color:#888;font-size:12px;padding:0 30px 25px}
</style>
'@
$now     = Get-Date -Format 'dd MMM yyyy HH:mm'
$verdict = if ($bad.Count -eq 0) {"<div class='verdict vok'>RESULT: No Microsoft Connected Cache (MCC) components found on any server</div>"} else {"<div class='verdict vbad'>RESULT: Review required on: $(Enc ($bad.VM -join ', '))</div>"}
$sumRows = ($results | ForEach-Object { "<tr><td><b>$(Enc $_.VM)</b></td><td>$(Badge $_.'MCC Status')</td><td>$(Badge $_.WSL)</td><td>$(Badge $_.IIS)</td><td>$(Badge $_.'Port 80')</td></tr>" }) -join ''
$cards = foreach ($r in $results) {
  $ap = @($r.Apps   -split ' \| ' | Where-Object {$_ -and $_ -ne '-'})
  $sv = @($r.Agents -split ' \| ' | Where-Object {$_ -and $_ -ne '-'})
  $apRows = ($ap | ForEach-Object { $p = $_ -split '~'; "<tr><td>$(Enc $p[0])</td><td>$(Enc $p[1])</td></tr>" }) -join ''
  $svRows = ($sv | ForEach-Object { "<tr><td>$(Enc $_)</td></tr>" }) -join ''
  "<div class='card'><h2>$(Enc $r.VM)</h2><h3>Installed applications <span class='cnt'>($($ap.Count))</span></h3><table><tr><th>Application</th><th>Version</th></tr>$apRows</table><h3>Running third-party services <span class='cnt'>($($sv.Count))</span></h3><table><tr><th>Service</th></tr>$svRows</table></div>"
}
$html = "<!DOCTYPE html><html><head><meta charset='utf-8'><title>MCC Check Report</title>$css</head><body><div class='hdr'><h1>Microsoft Connected Cache (MCC) Check</h1><p>Date: $now &nbsp;|&nbsp; Servers: $(Enc ($vms -join ', ')) &nbsp;|&nbsp; Method: read-only check via VMware Guest Operations</p></div><div class='wrap'>$verdict<div class='card'><h2>Summary</h2><table><tr><th>Server</th><th>MCC</th><th>WSL</th><th>IIS</th><th>Port 80</th></tr>$sumRows</table></div>$($cards -join '')</div><div class='ft'>Checks performed: MCC services, installed applications, scheduled tasks, folders, WSL (MCC Enterprise), IIS (ConfigMgr MCC), port 80 listener. No changes were made to the servers.</div></body></html>"

if (-not (Test-Path $outDir)) { New-Item -Path $outDir -ItemType Directory | Out-Null }
$file = Join-Path $outDir "MCC_Check_$(Get-Date -Format yyyyMMdd_HHmm).html"
$html | Out-File $file -Encoding UTF8
Write-Host "`n  Report saved: $file" -ForegroundColor Cyan
Start-Process $file
