# Clones every VM listed in C:\temp\vmlist.txt, one after another.
# Run it in a PowerShell window that is already connected to vCenter (Connect-VIServer).
# Each clone is named <VMName>_clone and created on the same host, datastore and folder as the original.

$vmNames = Get-Content -Path 'C:\temp\vmlist.txt' | Where-Object { $_.Trim() -ne '' }

foreach ($vmName in $vmNames) {
    $vmName    = $vmName.Trim()
    $cloneName = "${vmName}_clone"
    Write-Host "Cloning $vmName -> $cloneName ..." -ForegroundColor Cyan
    try {
        $vm        = Get-VM -Name $vmName -ErrorAction Stop
        $datastore = Get-Datastore -RelatedObject $vm | Select-Object -First 1
        New-VM -Name $cloneName -VM $vm -VMHost $vm.VMHost -Datastore $datastore -Location $vm.Folder -ErrorAction Stop | Out-Null
        Write-Host "Done: $cloneName" -ForegroundColor Green
    } catch {
        Write-Host "FAILED: $vmName - $($_.Exception.Message)" -ForegroundColor Red
    }
}
