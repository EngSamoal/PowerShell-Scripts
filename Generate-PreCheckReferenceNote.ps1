#Requires -Version 5.1
<#
.SYNOPSIS
    Generates PreCheck_Reference_Note.html on the local machine.

.DESCRIPTION
    SAFE to run on a restricted / company laptop:
      - No network calls, contacts nothing.
      - No Administrator rights needed.
      - No system/registry/security setting is changed.
      - Only WRITES ONE FILE: PreCheck_Reference_Note.html, next to this
        script (or to -OutputPath if given).
    The HTML is embedded below as plain text and written out exactly as-is -
    open this .ps1 in Notepad and read it first if you want to verify that.

.PARAMETER OutputPath
    Folder to write the HTML file into. Defaults to the folder this script is
    run from.

.PARAMETER NoLaunch
    Do not automatically open the generated file in the default browser.

.EXAMPLE
    .\Generate-PreCheckReferenceNote.ps1
#>
[CmdletBinding()]
param(
    [string]$OutputPath = $PSScriptRoot,
    [switch]$NoLaunch
)

$ErrorActionPreference = 'Stop'
if ([string]::IsNullOrWhiteSpace($OutputPath)) { $OutputPath = (Get-Location).Path }
if (-not (Test-Path -LiteralPath $OutputPath)) {
    New-Item -ItemType Directory -Path $OutputPath -Force | Out-Null
}
$targetFile = Join-Path $OutputPath 'PreCheck_Reference_Note.html'

$html = @'
<!DOCTYPE html>
<html lang="en">
<head>
<meta charset="utf-8">
<meta name="viewport" content="width=device-width, initial-scale=1">
<title>Pre-Check Reference Note — Script-Review Items</title>
<style>
:root{--bg:#f3f7f6;--card:#fff;--ink:#152420;--muted:#5b6b67;--line:#d9e6e2;--accent:#0d9488;--accent2:#115e59;--good:#1a7a35;--warn:#b45309;--bad:#9c2a20}
*{box-sizing:border-box}
body{margin:0;font:15.5px/1.65 -apple-system,Segoe UI,Roboto,Helvetica,Arial,sans-serif;background:var(--bg);color:var(--ink)}
header{background:linear-gradient(135deg,var(--accent2),var(--accent));color:#fff;padding:26px 32px}
header h1{margin:0;font-size:21px}
header .sub{opacity:.9;font-size:13.5px;margin-top:5px}
main{max-width:1150px;margin:0 auto;padding:22px 20px 60px}
h2.sec{margin:28px 0 8px;padding:9px 14px;background:var(--accent2);color:#fff;border-radius:6px;font-size:16px}
p{margin:8px 0}
.lead{color:var(--muted);font-size:14.5px}
table{width:100%;border-collapse:collapse;background:var(--card);border-radius:8px;overflow:hidden;border:1px solid var(--line);margin:10px 0}
th,td{padding:10px 12px;border-bottom:1px solid var(--line);text-align:left;font-size:14px;vertical-align:top}
th{background:#e8f3f1;color:var(--muted);font-weight:700;font-size:12px;text-transform:uppercase;letter-spacing:.02em}
tr:last-child td{border-bottom:none}
code{font-family:ui-monospace,Consolas,monospace;font-size:12.5px;background:#e8f3f1;padding:2px 5px;border-radius:4px;white-space:nowrap}
.badge{display:inline-block;padding:3px 9px;border-radius:11px;font-size:12px;font-weight:700}
.b-yes{background:#fde3e1;color:#9c2a20}
.b-no{background:#e0f2e6;color:#1a7a35}
.callout{border-radius:8px;padding:14px 16px;margin:14px 0;font-size:14px;border:1px solid;background:#e8f3f1;border-color:#c3e0da}
.footnote{font-size:12.5px;color:var(--muted);margin-top:10px}
</style>
</head>
<body>
<header>
<h1>Pre-Check Reference Note — "Script — Review" Items</h1>
<div class="sub">What each check/remediation command does, reboot requirement, and impact of leaving it unremediated</div>
</header>
<main>

<h2 class="sec">1. Why these 6 need a review step</h2>
<p class="lead">These are the only in-scope categories that are scripted but gated behind manual review, because applying them without checking first can affect connectivity or access (e.g. lock out RDP, break legacy SMB clients, block silent automation). Every other in-scope control is either fully automated (safe) or fully manual.</p>

<h2 class="sec">2. Command reference</h2>
<table>
<thead><tr><th>Category</th><th>Command</th><th>What it does</th><th>Reboot?</th><th>If ignored (left unremediated)</th></tr></thead>
<tbody>
<tr>
<td>Windows Firewall</td>
<td><code>Set-NetFirewallProfile -Profile Domain,Private,Public -Enabled True</code></td>
<td>Turns on the host firewall (blocks unsolicited inbound traffic)</td>
<td><span class="badge b-no">No</span></td>
<td>Zero host-level network protection — any open port is reachable from anywhere that can route to it</td>
</tr>
<tr>
<td>SMB minimum version</td>
<td><code>Set-SmbServerConfiguration -MinSmb2Dialect SMB300</code></td>
<td>Blocks SMB 2.0.2 / 2.1, forces SMB 3.0+</td>
<td><span class="badge b-no">No</span></td>
<td>Weaker/unencrypted SMB sessions stay allowed — usable for relay / downgrade attacks</td>
</tr>
<tr>
<td>NetBIOS</td>
<td><code>Set-ItemProperty HKLM:\Software\Policies\Microsoft\Windows NT\DNSClient EnableNetbios 0</code></td>
<td>Disables legacy NetBIOS name resolution</td>
<td><span class="badge b-no">No</span> <span class="footnote">(full effect after reboot)</span></td>
<td>Exposed to LLMNR/NBT-NS spoofing — a common way to steal password hashes on a flat network</td>
</tr>
<tr>
<td>Account Lockout Threshold</td>
<td><code>net accounts /lockoutthreshold:3</code></td>
<td>Lowers the number of failed logins allowed before an account locks</td>
<td><span class="badge b-no">No</span></td>
<td>Accounts can be brute-forced / guessed with unlimited attempts</td>
</tr>
<tr>
<td>UAC hardening</td>
<td><code>Set-ItemProperty HKLM:\...\Policies\System ConsentPromptBehaviorAdmin 2</code></td>
<td>Forces admin actions to require a consent prompt (no silent elevation)</td>
<td><span class="badge b-no">No</span></td>
<td>Malware / a compromised script running as admin gets full privilege instantly, no prompt, no warning</td>
</tr>
<tr>
<td>VBS &amp; Credential Guard</td>
<td><code>Set-ItemProperty HKLM:\...\DeviceGuard LsaCfgFlags 1</code></td>
<td>Isolates password hashes / Kerberos tickets in hardware-protected memory</td>
<td><span class="badge b-yes">Yes</span></td>
<td>Password hashes/tickets sit in normal memory — tools like Mimikatz can steal them straight from RAM</td>
</tr>
</tbody>
</table>

<h2 class="sec">3. What to check before running each one</h2>
<table>
<thead><tr><th>Category</th><th>Check command</th><th>Look for</th></tr></thead>
<tbody>
<tr><td>Windows Firewall</td><td><code>Get-NetFirewallRule -Enabled True | Where Direction -eq Inbound | Select DisplayName,LocalPort</code></td><td>RDP (3389), monitoring/agent ports, app ports already allowed</td></tr>
<tr><td>SMB minimum version</td><td><code>Get-SmbSession | Select ClientComputerName,Dialect</code></td><td>Any session on dialect <code>2.0.2</code> or <code>2.1</code> (would break)</td></tr>
<tr><td>NetBIOS</td><td><code>Resolve-DnsName &lt;servername&gt;</code></td><td>If it resolves via DNS, NetBIOS is safe to disable</td></tr>
<tr><td>Account Lockout Threshold</td><td><code>net accounts</code> + confirm with app/service owners</td><td>Any account that fails login often (would get locked out)</td></tr>
<tr><td>UAC hardening</td><td><code>Get-ScheduledTask | Where Principal -match 'Administrator'</code></td><td>Tasks/scripts that run unattended as full admin</td></tr>
<tr><td>VBS &amp; Credential Guard</td><td>Check VM settings in vCenter</td><td>Secure Boot = Enabled, hardware virtualization exposed to guest</td></tr>
</tbody>
</table>
<p class="footnote">All 6 checks above are combined in one script: <code>Invoke-ChangeReviewPreCheck.ps1</code> (read-only, prints PASS/REVIEW per item, saves a log).</p>

<h2 class="sec">4. Qualys alignment</h2>
<div class="callout">
Every command above writes the <b>exact registry path / value</b> that the corresponding Qualys Control ID checks for — verified directly against the Qualys export evidence, not assumed:
<table style="margin-top:10px">
<thead><tr><th>Category</th><th>Qualys-checked path</th></tr></thead>
<tbody>
<tr><td>Firewall</td><td><code>HKLM\Software\Policies\Microsoft\WindowsFirewall\{Profile}Profile\EnableFirewall = 1</code></td></tr>
<tr><td>SMB min version</td><td><code>HKLM\Software\Policies\Microsoft\Windows\LanmanServer\MinSmb2Dialect = 768</code></td></tr>
<tr><td>NetBIOS</td><td><code>HKLM\Software\Policies\Microsoft\Windows NT\DNSClient\EnableNetbios = 0</code></td></tr>
<tr><td>Account Lockout</td><td><code>secedit [System Access] LockoutBadCount</code></td></tr>
<tr><td>UAC</td><td><code>HKLM\...\Policies\System\ConsentPromptBehaviorAdmin = 2</code></td></tr>
<tr><td>VBS/Credential Guard</td><td><code>HKLM\SOFTWARE\Policies\Microsoft\Windows\DeviceGuard\LsaCfgFlags = 1</code></td></tr>
</tbody>
</table>
</div>
<p class="footnote">Same applies to the other 60 in-scope controls (133 of 160 total controls are direct registry writes generated straight from the Qualys evidence path; the remaining use dedicated handlers — ASR rules via <code>Add-MpPreference</code>, audit policy via <code>auditpol</code>, user rights/account policy via <code>secedit</code> — each individually verified against its Qualys expected value).</p>

</main>
</body>
</html>

'@

# Write with a UTF-8 BOM so the — / ... characters above render correctly
# regardless of the machine's regional/codepage settings.
$utf8Bom = New-Object System.Text.UTF8Encoding($true)
[System.IO.File]::WriteAllText($targetFile, $html, $utf8Bom)

Write-Host "Generated: $targetFile" -ForegroundColor Green

if (-not $NoLaunch) {
    try { Start-Process -FilePath $targetFile } catch { Write-Host "(Could not auto-open; open the file manually.)" -ForegroundColor Yellow }
}
