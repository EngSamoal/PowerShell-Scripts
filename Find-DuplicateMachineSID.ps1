<#
.SYNOPSIS
    Read-only: lists every powered-on Windows VM in vCenter that shares its machine SID with another VM.

.DESCRIPTION
    Windows Server 2025 / Windows 11 24H2+ with updates from 29 Aug 2025 onward block NTLM/Kerberos
    logons between machines that share a machine SID (LsaSrv Event 6167, Microsoft KB5070568).
    VMs deployed from a template that was never sysprepped all share one SID.

    Runs on your laptop (PowerCLI, already connected to vCenter). For each VM it reads the machine SID
    and OS build through VMware Tools (Invoke-VMScript). Nothing is changed in any guest.

    Guest credentials: you are asked for one or more local admin credentials; each VM is tried with
    them in order until one works. VMs where none works are listed with the reason.

.PARAMETER NameFilter
    Only scan VMs whose name matches this wildcard. Default: all.

.PARAMETER OutputFolder
    Where the CSV is saved. Default C:\temp.

.EXAMPLE
    Connect-VIServer <vcenter>
    .\Find-DuplicateMachineSID.ps1
    .\Find-DuplicateMachineSID.ps1 -NameFilter 'AQ-*'
#>
[CmdletBinding()]
param(
    [string] $NameFilter = '*',
    [string] $OutputFolder = 'C:\temp'
)

$ScriptBuild = 'Find-DuplicateMachineSID build 2026-11-09.1'
Write-Host $ScriptBuild -ForegroundColor Cyan

if (-not $global:DefaultVIServers) { Write-Host 'Not connected to vCenter. Run Connect-VIServer first.' -ForegroundColor Red; return }

$creds = @(Get-Credential -Message 'Local admin credential to try (1st)')
while ((Read-Host 'Add another credential to try? (y/n)') -eq 'y') {
    $creds += Get-Credential -Message ("Local admin credential to try ({0})" -f ($creds.Count + 1))
}

# Runs inside each guest: machine SID | OS name | build
$probe = '$s=(Get-LocalUser | Select-Object -First 1).SID.AccountDomainSid.Value; $o=Get-ItemProperty ''HKLM:\SOFTWARE\Microsoft\Windows NT\CurrentVersion''; "{0}|{1}|{2}.{3}" -f $s,$o.ProductName,$o.CurrentBuild,$o.UBR'

$vms = @(Get-VM -Name $NameFilter | Where-Object { $_.PowerState -eq 'PoweredOn' -and $_.Guest.OSFullName -like '*Windows*' } | Sort-Object Name)
Write-Host ("Scanning {0} powered-on Windows VM(s)...`n" -f $vms.Count)

$i = 0
$rows = foreach ($v in $vms) {
    $i++
    Write-Progress -Activity 'Reading machine SIDs' -Status $v.Name -PercentComplete (100 * $i / [math]::Max($vms.Count, 1))
    $out = $null; $err = 'No credential worked'
    if ($v.ExtensionData.Guest.ToolsRunningStatus -ne 'guestToolsRunning') {
        $err = 'VMware Tools not running'
    } else {
        foreach ($c in $creds) {
            try { $out = (Invoke-VMScript -VM $v -GuestCredential $c -ScriptType Powershell -ScriptText $probe -ErrorAction Stop).ScriptOutput.Trim(); break }
            catch { $err = $_.Exception.Message }
        }
    }
    $p = if ($out -and $out -match '^S-1-5-21-') { $out -split '\|' } else { @('', '', '') }
    [pscustomobject]@{
        VM         = $v.Name
        IP         = ($v.Guest.IPAddress | Where-Object { $_ -match '^\d+\.\d+\.\d+\.\d+$' }) -join ', '
        MachineSID = $p[0]
        OS         = $p[1]
        Build      = $p[2]
        Error      = $(if ($p[0]) { '' } elseif ($out) { "Unexpected output: $out" } else { $err })
    }
}
Write-Progress -Activity 'Reading machine SIDs' -Completed

if (-not (Test-Path $OutputFolder)) { New-Item -ItemType Directory -Path $OutputFolder -Force | Out-Null }
$csv = Join-Path $OutputFolder ("VM_MachineSID_Scan_{0}.csv" -f (Get-Date -Format 'yyyyMMdd_HHmm'))
$rows | Export-Csv -Path $csv -NoTypeInformation -Encoding UTF8

$dupes = @($rows | Where-Object MachineSID | Group-Object MachineSID | Where-Object Count -gt 1 | Sort-Object Count -Descending)
Write-Host '================ VMs SHARING A MACHINE SID ================' -ForegroundColor Cyan
if ($dupes.Count -eq 0) {
    Write-Host 'None found - every scanned VM has a unique machine SID.' -ForegroundColor Green
} else {
    foreach ($d in $dupes) {
        Write-Host ("`nSID {0}  ({1} VMs)" -f $d.Name, $d.Count) -ForegroundColor Red
        $d.Group | Format-Table VM, IP, OS, Build -AutoSize | Out-String -Width 200 | Write-Host
    }
    Write-Host 'Any two VMs in the same group that connect to each other (PAM/RDP, file shares) will be blocked.' -ForegroundColor Yellow
}

$failed = @($rows | Where-Object Error)
if ($failed.Count) {
    Write-Host ("`n{0} VM(s) could not be checked:" -f $failed.Count) -ForegroundColor Yellow
    $failed | Format-Table VM, Error -AutoSize -Wrap | Out-String -Width 200 | Write-Host
}

Write-Host ("`nUnique: {0}   In duplicate groups: {1}   Not checked: {2}" -f
    @($rows | Where-Object MachineSID | Group-Object MachineSID | Where-Object Count -eq 1).Count,
    @($dupes | ForEach-Object Group).Count, $failed.Count)
Write-Host "Saved: $csv" -ForegroundColor Cyan
Write-Host 'Nothing was changed on any VM.' -ForegroundColor Cyan
