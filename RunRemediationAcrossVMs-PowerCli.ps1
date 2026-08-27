# Runs RepairWindowsServerSecurityGapsLocal.ps1 against every VM listed in vmlist.txt,
# via vCenter/PowerCLI (Invoke-VMScript) - no network/WinRM path needed.

$vCenterServer = '10.50.10.10'
$vmListPath    = 'C:\temp\vmlist.txt'
$credPath      = 'C:\temp\guestcred.xml'
$scriptPath    = '.\RepairWindowsServerSecurityGapsLocal.ps1'
$Apply         = $false   # set to $true to actually apply changes on every VM in the list

Connect-VIServer -Server $vCenterServer

$guestCred = Import-Clixml -Path $credPath
$vmNames   = Get-Content -Path $vmListPath | Where-Object { $_.Trim() -ne '' }
$rawScript = Get-Content -Path $scriptPath -Raw

# Wrap the script body in a scriptblock invocation so -Apply actually gets bound
# to its param() block - just appending text after the raw script does NOT do this.
$applyArg   = if ($Apply) { ' -Apply' } else { '' }
$scriptText = "& {`n$rawScript`n}$applyArg"

foreach ($vmName in $vmNames) {
    Write-Host "`n=== $vmName ===" -ForegroundColor Cyan
    try {
        $result = Invoke-VMScript -VM $vmName -ScriptText $scriptText -ScriptType Powershell -GuestCredential $guestCred -ErrorAction Stop
        Write-Host $result.ScriptOutput
    } catch {
        Write-Host "FAILED: $($_.Exception.Message)" -ForegroundColor Red
    }
}

Disconnect-VIServer -Server $vCenterServer -Confirm:$false
