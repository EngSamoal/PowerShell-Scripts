# Quick post-remediation check — read-only, changes nothing.

$checks = @(
    @{ Name='VBS Enabled';               Path='HKLM:\SYSTEM\CurrentControlSet\Control\DeviceGuard'; Value='EnableVirtualizationBasedSecurity'; Expected=1 }
    @{ Name='HVCI Enabled';              Path='HKLM:\SYSTEM\CurrentControlSet\Control\DeviceGuard\Scenarios\HypervisorEnforcedCodeIntegrity'; Value='Enabled'; Expected=1 }
    @{ Name='Credential Guard';          Path='HKLM:\SYSTEM\CurrentControlSet\Control\LSA'; Value='LsaCfgFlags'; Expected=1 }
    @{ Name='Machine Identity Isolation';Path='HKLM:\SOFTWARE\Policies\Microsoft\Windows\DeviceGuard'; Value='MachineIdentityIsolation'; Expected=2 }
    @{ Name='WDigest Disabled';          Path='HKLM:\SYSTEM\CurrentControlSet\Control\SecurityProviders\WDigest'; Value='UseLogonCredential'; Expected=0 }
    @{ Name='WinRM Unencrypted Blocked'; Path='HKLM:\SOFTWARE\Microsoft\Windows\CurrentVersion\WSMAN\Service'; Value='AllowUnencrypted'; Expected=0 }
    @{ Name='WinRM Basic Auth Blocked';  Path='HKLM:\SOFTWARE\Microsoft\Windows\CurrentVersion\WSMAN\Service'; Value='auth_Basic'; Expected=0 }
    @{ Name='Printer RPC Privacy';       Path='HKLM:\SYSTEM\CurrentControlSet\Control\Print'; Value='RpcAuthnLevelPrivacyEnabled'; Expected=1 }
    @{ Name='Point-and-Print Restricted';Path='HKLM:\SOFTWARE\Policies\Microsoft\Windows NT\Printers\PointAndPrint'; Value='RestrictDriverInstallationToAdministrators'; Expected=1 }
    @{ Name='SmartScreen Enabled';       Path='HKLM:\SOFTWARE\Policies\Microsoft\Windows\System'; Value='EnableSmartScreen'; Expected=1 }
)

foreach ($c in $checks) {
    $actual = (Get-ItemProperty -Path $c.Path -Name $c.Value -ErrorAction SilentlyContinue).($c.Value)
    $status = if ($actual -eq $c.Expected) { "OK" } else { "NOT SET (found: $actual)" }
    "{0,-30} {1}" -f $c.Name, $status
}