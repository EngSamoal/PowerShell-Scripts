#Requires -Version 5.1
<#
.SYNOPSIS
    Generates CR_Remediation_Action_Plan.html on the local machine.

.DESCRIPTION
    This script is SAFE to run on a restricted / company laptop:
      - It makes NO network calls and contacts nothing.
      - It does NOT need Administrator rights.
      - It does NOT change any system, registry, or security setting.
      - It only WRITES ONE FILE: CR_Remediation_Action_Plan.html, next to this
        script (or to -OutputPath if you pass one), and does nothing else.
    The entire HTML report is embedded below as plain text and is written out
    exactly as-is - open the .ps1 in Notepad first and read it if you want to
    verify that before running it.

.PARAMETER OutputPath
    Folder to write the HTML file into. Defaults to the folder this script is
    run from.

.PARAMETER NoLaunch
    Do not automatically open the generated file in the default browser.

.EXAMPLE
    .\Generate-RemediationActionPlan.ps1
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
$targetFile = Join-Path $OutputPath 'CR_Remediation_Action_Plan.html'

$html = @'
<!DOCTYPE html>
<html lang="en">
<head>
<meta charset="utf-8">
<meta name="viewport" content="width=device-width, initial-scale=1">
<title>Servers Policy Compliance — Remediation Action Plan</title>
<style>
:root{--bg:#f4f6f9;--card:#fff;--ink:#1c2530;--muted:#5b6876;--line:#dde3ea;--accent:#0b5cab;--accent2:#08386a;--good:#1a7a35;--warn:#a5540d;--bad:#9c2a20}
*{box-sizing:border-box}
body{margin:0;font:15.5px/1.65 -apple-system,Segoe UI,Roboto,Helvetica,Arial,sans-serif;background:var(--bg);color:var(--ink)}
header{background:linear-gradient(135deg,var(--accent2),var(--accent));color:#fff;padding:30px 32px}
header h1{margin:0;font-size:24px}
header .sub{opacity:.9;font-size:14.5px;margin-top:5px}
main{max-width:1150px;margin:0 auto;padding:24px 20px 70px}
h2.sec{margin:34px 0 6px;padding:10px 14px;background:var(--accent2);color:#fff;border-radius:6px;font-size:17px}
h3.sub{margin:20px 0 8px;color:var(--accent2);font-size:15.5px}
p{margin:8px 0}
.lead{color:var(--muted);font-size:15px}
.stepnum{display:inline-flex;align-items:center;justify-content:center;width:24px;height:24px;border-radius:50%;background:var(--accent2);color:#fff;font-size:13px;font-weight:800;margin-right:8px}
.kpis{display:grid;grid-template-columns:repeat(auto-fit,minmax(160px,1fr));gap:10px;margin:14px 0}
.kpi{background:var(--card);border:1px solid var(--line);border-radius:8px;padding:12px 14px}
.kpi .n{font-size:24px;font-weight:800;color:var(--accent2)}
.kpi.reduce .n{color:var(--good)}
.kpi .l{font-size:13px;color:var(--muted);margin-top:2px}
.callout{border-radius:8px;padding:14px 16px;margin:12px 0;font-size:14.5px;border:1px solid}
.callout.evidence{background:#eef7ee;border-color:#bfe3bf}
.callout.warn{background:#fff8e6;border-color:#f0d98a}
.callout.info{background:#eef4fb;border-color:#cfe0f2}
.quote{font-style:italic;border-left:3px solid var(--good);padding:6px 0 6px 12px;margin:8px 0;background:#fff;border-radius:0 6px 6px 0}
.quote a{color:var(--accent)}
table{width:100%;border-collapse:collapse;background:var(--card);border-radius:8px;overflow:hidden;border:1px solid var(--line);margin:10px 0}
th,td{padding:10px 12px;border-bottom:1px solid var(--line);text-align:left;font-size:14.5px;vertical-align:middle}
th{background:#eef2f7;color:var(--muted);font-weight:700;font-size:12px;text-transform:uppercase;letter-spacing:.02em}
tr.excluded td{color:var(--muted)}
tr.excluded{background:#f7f8fa}
tr.total{font-weight:700;background:#eef4fb}
.badge{display:inline-flex;align-items:center;justify-content:center;width:180px;height:34px;padding:0 8px;border-radius:12px;font-size:13px;font-weight:700;white-space:nowrap;text-align:center;box-sizing:border-box}
.b-script{background:#e2eefb;color:#0b4a8f}
.b-review{background:#fdf2d8;color:#8a5a06}
.b-manual{background:#fde3e1;color:#9c2a20}
.b-mixed{background:#f1e9fb;color:#5a3b9c}
.b-na{background:#eceff2;color:#5b6876}
.b-out{background:#eceff2;color:#5b6876}
.count{font-size:13px;color:var(--muted)}
.footnote{font-size:13px;color:var(--muted);margin-top:6px}
.legend{display:flex;flex-wrap:wrap;gap:8px;margin:10px 0}
ol,ul{margin:6px 0;padding-left:22px}
li{margin:4px 0}
code{font-family:ui-monospace,Consolas,monospace;font-size:12.5px;background:#f3f5f8;padding:1px 4px;border-radius:4px}
</style>
</head>
<body>
<header>
<h1>Family Entertainment's Servers Policy Compliance — Remediation Action Plan</h1>
<div class="sub">Qualys Policy Compliance assessment · Windows Server 2025 Datacenter 24H2 · WORKGROUP environment</div>
</header>
<main>

<h2 class="sec">1. Purpose &amp; Methodology</h2>
<p>This document is the remediation action plan for the Qualys Policy Compliance scan performed against the in-scope Windows Server 2025 servers. It covers what was found, which findings are relevant to this environment, and — for every relevant finding — how it should be fixed (automated script, script with change review, or manual action).</p>
<p class="lead">Approach: the raw Qualys export (472 individual findings) was consolidated by Control ID to remove duplication across hosts, grouped into 16 technical categories, and each unique control was assessed for applicability to this specific environment (Windows Server 2025, standalone/WORKGROUP, no Active Directory) before a remediation method was assigned.</p>
<div class="kpis">
<div class="kpi"><div class="n">472</div><div class="l">Raw findings (Qualys export)</div></div>
<div class="kpi"><div class="n">160</div><div class="l">Unique Control IDs after consolidation</div></div>
<div class="kpi"><div class="n">16</div><div class="l">Technical categories</div></div>
</div>

<h2 class="sec">2. Applicability Review — Internet Explorer Findings</h2>
<p><span class="stepnum">1</span>During consolidation, <b>94 of the 160 Control IDs (59% of all controls, 282 of 472 raw findings)</b> were found to test only <b>Internet Explorer</b> settings — browser zone security (ActiveX, scripting, downloads), IE feature-control keys, and per-user IE autofill options.</p>
<p><span class="stepnum">2</span>These controls exist in the Qualys policy because it is a generic Windows benchmark; they were never written specifically for Windows Server 2025.</p>
<p><span class="stepnum">3</span><b>Internet Explorer is not present on Windows Server 2025.</b> Microsoft removed the standalone Internet Explorer application from the product starting with this release — it is not simply disabled or deprecated, it has been taken out of the operating system entirely.</p>

<div class="callout evidence">
<b>Microsoft evidence</b>
<div class="quote">"Microsoft removed the standalone Internet Explorer application from Windows Server 2025."<br>— <i>Features Removed or No Longer Developed in Windows Server</i>, Microsoft Learn, Windows Server 2025 tab, table "Features removed", row "Internet Explorer" (last updated 2026-03-19).<br>
<a href="https://learn.microsoft.com/en-us/windows-server/get-started/removed-deprecated-features-windows-server-2025" target="_blank">learn.microsoft.com/.../removed-deprecated-features-windows-server-2025</a></div>
Full published text: <i>"The Internet Explorer 11 desktop application is retired and out of support as of June 15, 2022. Microsoft removed the standalone Internet Explorer application from Windows Server 2025. To access older, legacy sites, use IE mode in Microsoft Edge."</i>
<div class="footnote">Supporting reference (context only — this one covers the original 2022 client-only retirement and explicitly states Windows Server was <i>not</i> affected by it; the Windows Server-specific removal is the citation above): <a href="https://aka.ms/IEJune15Blog" target="_blank">Internet Explorer 11 desktop app retirement FAQ — Microsoft Tech Community</a>.</div>
</div>

<p><span class="stepnum">4</span><b>Conclusion:</b> because the application these 94 controls test for does not exist on Windows Server 2025, they are assessed as <b>Not Applicable</b> to this environment and are excluded from the remediation scope below. Recommendation: submit these 94 Control IDs to the Qualys/GRC exception process using the evidence above, rather than remediate them.</p>

<h2 class="sec">3. Resulting Remediation Scope</h2>
<table>
<thead><tr><th></th><th>Control IDs</th><th>Raw findings</th><th>% of total findings</th></tr></thead>
<tbody>
<tr><td>Total from Qualys scan</td><td>160</td><td>472</td><td>100%</td></tr>
<tr class="excluded"><td>− Internet Explorer (Not Applicable, §2)</td><td>94</td><td>282</td><td>59.7%</td></tr>
<tr class="total"><td>= Remaining — actionable remediation scope</td><td>66</td><td>190</td><td>40.3%</td></tr>
</tbody>
</table>
<p class="lead">In plain terms: excluding the Internet Explorer findings reduces the real work from 160 controls down to <b>66 controls</b> — a 59.7% reduction — across the 13 categories below.</p>

<h2 class="sec">4. Remediation Action Plan by Category</h2>
<p>For every in-scope category, the table shows how many of its controls fall into each remediation method. <b>Primary method</b> is the majority method for that category.</p>
<div class="legend">
<span class="badge b-script">Script — Automated</span>
<span class="badge b-review">Script — Review</span>
<span class="badge b-manual">Manual</span>
<span class="badge b-mixed">Mixed</span>
</div>
<table>
<thead><tr><th>Category</th><th>Control IDs</th><th>Findings</th><th>Script</th><th>Script+Review</th><th>Manual</th><th>N/A</th><th>Primary remediation method</th></tr></thead>
<tbody>
<tr><td>Audit Policy</td><td>1</td><td>3</td><td>1</td><td>0</td><td>0</td><td>0</td><td><span class="badge b-script">Script — Automated</span></td></tr>
<tr><td>PowerShell Logging</td><td>1</td><td>3</td><td>1</td><td>0</td><td>0</td><td>0</td><td><span class="badge b-script">Script — Automated</span></td></tr>
<tr><td>Microsoft Defender AV &amp; Exploit Guard</td><td>12</td><td>36</td><td>9</td><td>2</td><td>1</td><td>0</td><td><span class="badge b-script">Script — Automated</span></td></tr>
<tr><td>SMB / LanMan Hardening</td><td>12</td><td>36</td><td>8</td><td>4</td><td>0</td><td>0</td><td><span class="badge b-script">Script — Automated</span></td></tr>
<tr><td>Print Spooler / RPC</td><td>2</td><td>6</td><td>1</td><td>1</td><td>0</td><td>0</td><td><span class="badge b-mixed">Mixed</span></td></tr>
<tr><td>Kerberos / PKINIT Certificate Logon</td><td>10</td><td>30</td><td>5</td><td>0</td><td>0</td><td>5</td><td><span class="badge b-mixed">Mixed</span></td></tr>
<tr><td>Security Options &amp; UAC</td><td>6</td><td>17</td><td>2</td><td>4</td><td>0</td><td>0</td><td><span class="badge b-review">Script — Review</span></td></tr>
<tr><td>Virtualization-Based Security &amp; Credential Guard</td><td>8</td><td>24</td><td>0</td><td>7</td><td>0</td><td>1</td><td><span class="badge b-review">Script — Review</span></td></tr>
<tr><td>Windows Firewall</td><td>3</td><td>3</td><td>0</td><td>3</td><td>0</td><td>0</td><td><span class="badge b-review">Script — Review</span></td></tr>
<tr><td>Password &amp; Account Lockout Policy</td><td>2</td><td>5</td><td>0</td><td>2</td><td>0</td><td>0</td><td><span class="badge b-review">Script — Review</span></td></tr>
<tr><td>Network Security (NetBIOS)</td><td>1</td><td>3</td><td>0</td><td>1</td><td>0</td><td>0</td><td><span class="badge b-review">Script — Review</span></td></tr>
<tr><td>LSA Protection</td><td>2</td><td>6</td><td>0</td><td>1</td><td>1</td><td>0</td><td><span class="badge b-mixed">Mixed</span></td></tr>
<tr><td>User Rights Assignment</td><td>6</td><td>18</td><td>1</td><td>0</td><td>5</td><td>0</td><td><span class="badge b-manual">Manual</span></td></tr>
<tr class="total"><td>TOTAL — in-scope</td><td>66</td><td>190</td><td>28</td><td>25</td><td>7</td><td>6</td><td></td></tr>
</tbody>
</table>

<h3 class="sub">4.1 Script — safe, automated (28 controls)</h3>
<p>Registry/policy-value fixes with no functional or connectivity impact (SmartScreen, Defender exclusion-visibility &amp; update-channel settings, PowerShell script-block logging, audit policy, LSA/VBS registry values, Kerberos PKINIT client values, printer RPC). Can be applied directly via the remediation script, no reboot required for these specific items.</p>

<h3 class="sub">4.2 Script — needs change review before running (25 controls)</h3>
<p>The fix is already scripted, but it changes something that could affect connectivity or access if applied blindly, so it is gated behind an explicit opt-in switch and should be reviewed/scheduled per item:</p>
<ul>
<li><b>Windows Firewall</b> — enabling the host firewall profiles requires inbound allow-rules to be staged first (RDP, monitoring/agent ports, app ports).</li>
<li><b>VBS &amp; Credential Guard</b> — requires the VM to expose Secure Boot/virtualization and a reboot; can affect non-WHQL LSA plugins.</li>
<li><b>Security Options &amp; UAC</b> — can affect unattended/automation jobs that rely on silent elevation.</li>
<li><b>Password &amp; Account Lockout Policy</b> / <b>NetBIOS</b> / <b>SMB minimum dialect</b> — can affect legacy clients or lock out accounts if thresholds are too aggressive.</li>
</ul>

<h3 class="sub">4.3 Manual — human / process action (7 controls)</h3>
<p>Cannot be safely automated — needs a decision or coordination with another team:</p>
<ul>
<li><b>User Rights Assignment (5 of 6)</b> — rights currently held by Splunk, ServiceNow and IIS service accounts beyond the expected "Administrators" baseline; removing them needs sign-off from the app/agent owner (risk of breaking log forwarding / discovery / web app functionality).</li>
<li><b>LSA Protection (1 of 2)</b> — requires a documented security decision (Qualys's expected value is the less-hardened option; needs sign-off before changing).</li>
<li><b>Microsoft Defender (1 of 12)</b> — Network Protection setting needs a documented decision on the organization's Defender standard.</li>
</ul>

<h3 class="sub">4.4 Not Applicable (6 controls)</h3>
<p>Same treatment as the Internet Explorer findings — submit as Qualys exceptions, do not remediate:</p>
<ul>
<li><b>Machine Identity Isolation (1)</b> — only meaningful for a domain-joined machine; these servers are WORKGROUP.</li>
<li><b>KDC PKINIT hash settings (5)</b> — only apply to a Domain Controller role, which is not installed on these servers.</li>
</ul>

<h2 class="sec">5. Summary</h2>
<table>
<thead><tr><th>Remediation method</th><th>Applies to</th><th>Control IDs</th><th>Findings</th></tr></thead>
<tbody>
<tr><td><span class="badge b-script">Script — Automated</span></td><td>13 categories, safe fixes</td><td>28</td><td>83</td></tr>
<tr><td><span class="badge b-review">Script — Review</span></td><td>Scripted, needs change review</td><td>25</td><td>68</td></tr>
<tr><td><span class="badge b-manual">Manual</span></td><td>Human / process action</td><td>7</td><td>21</td></tr>
<tr><td><span class="badge b-na">Not Applicable</span></td><td>Workgroup / role not installed</td><td>6</td><td>18</td></tr>
<tr class="total"><td colspan="2">Subtotal — in-scope (§3)</td><td>66</td><td>190</td></tr>
<tr class="excluded"><td><span class="badge b-out">Not Applicable</span></td><td>Internet Explorer (§2)</td><td>94</td><td>282</td></tr>
<tr class="total"><td colspan="2">TOTAL — original Qualys scan</td><td>160</td><td>472</td></tr>
</tbody>
</table>
<p class="lead">Full per-control detail (meaning, risk, exact registry path, evidence-collection command, remediation steps) is in the companion document <code>ISO_Consolidated_Remediation_Report.html</code>. The 28 automated and 25 change-review controls are already implemented in the revised remediation scripts delivered separately.</p>

</main>
</body>
</html>

'@

# Write with a UTF-8 BOM so the — / § / · characters above render correctly
# regardless of the machine's regional/codepage settings.
$utf8Bom = New-Object System.Text.UTF8Encoding($true)
[System.IO.File]::WriteAllText($targetFile, $html, $utf8Bom)

Write-Host "Generated: $targetFile" -ForegroundColor Green

if (-not $NoLaunch) {
    try { Start-Process -FilePath $targetFile } catch { Write-Host "(Could not auto-open; open the file manually.)" -ForegroundColor Yellow }
}
