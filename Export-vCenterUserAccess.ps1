<#
.SYNOPSIS
    READ-ONLY report of every user and group that has access to each connected vCenter.

.DESCRIPTION
    For every vCenter in the current PowerCLI session (or the ones passed in -vCenter) this
    reads all vCenter permissions (Get-VIPermission) and writes:

        Permissions   - one row per permission: who, role, on which object, propagate, etc.
        Principals    - one row per user/group per vCenter, with all of its roles and objects
        Matrix        - one row per user/group, one column per vCenter (role(s) or blank)
        GroupMembers  - only with -ExpandGroups: the users inside each group that has access
                        (AD groups via the ActiveDirectory module, vsphere.local/SSO groups via
                        the VMware.vSphere.SsoAdmin module if you are connected with
                        Connect-SsoAdminServer). Anything that can't be expanded is listed
                        with the reason, not skipped.

    Output (in -OutputFolder):
        vCenter_UserAccess_<timestamp>.xlsx            - all sheets, if ImportExcel is installed
        vCenter_UserAccess_<sheet>_<timestamp>.csv     - one CSV per sheet (always)

    NOTE: "Global Permissions" (Administration > Access Control > Global Permissions) are NOT
    returned by Get-VIPermission / the vSphere API, so they are not in this report. Check them
    in the vSphere Client for each vCenter (they replicate across linked-mode vCenters).

    Nothing in vCenter, SSO or AD is changed. Only Get-* cmdlets are used.

.PARAMETER vCenter
    One or more vCenter FQDNs/IPs. If omitted, the existing PowerCLI session(s) are used.

.PARAMETER OutputFolder
    Folder for the output files. Default: C:\Temp

.PARAMETER ExpandGroups
    Also list the individual users inside each group that has access.

.PARAMETER IncludeSystemAccounts
    Include built-in solution/service accounts (vpxd-*, vsphere-webclient-*, VSPHERE.LOCAL\
    vStatsGroup, etc.). Hidden by default to keep the list readable.

.EXAMPLE
    .\Export-vCenterUserAccess.ps1
    Uses the vCenters you are already connected to.
.EXAMPLE
    .\Export-vCenterUserAccess.ps1 -ExpandGroups -OutputFolder D:\Reports
.EXAMPLE
    .\Export-vCenterUserAccess.ps1 -vCenter vc01.corp.local, vc02.corp.local

.NOTES
    Requires VMware PowerCLI:  Install-Module VMware.PowerCLI -Scope CurrentUser
    Optional Excel output:     Install-Module ImportExcel -Scope CurrentUser
    Optional -ExpandGroups:    RSAT ActiveDirectory module (AD groups)
                               Install-Module VMware.vSphere.SsoAdmin + Connect-SsoAdminServer
                               (vsphere.local groups)
#>
[CmdletBinding()]
param(
    [string[]]$vCenter,
    [string]$OutputFolder = 'C:\Temp',
    [switch]$ExpandGroups,
    [switch]$IncludeSystemAccounts
)

$ScriptBuild = 'Export-vCenterUserAccess build 2026-10-26.1'
Write-Host $ScriptBuild -ForegroundColor Cyan

#region ---- Connect ----
if ($vCenter) {
    Set-PowerCLIConfiguration -InvalidCertificateAction Ignore -Scope Session -Confirm:$false | Out-Null
    $cred = Get-Credential -Message 'vCenter credentials'
    foreach ($vc in $vCenter) {
        Write-Host "Connecting to $vc ..." -ForegroundColor Cyan
        Connect-VIServer -Server $vc -Credential $cred -ErrorAction Stop | Out-Null
    }
}
if (-not $global:DefaultVIServers -or @($global:DefaultVIServers).Count -eq 0) {
    throw 'Not connected to any vCenter. Use -vCenter <name> or run Connect-VIServer first.'
}
if (-not (Test-Path $OutputFolder)) { New-Item -ItemType Directory -Path $OutputFolder -Force | Out-Null }
#endregion

# Built-in service/solution accounts that every vCenter has; hidden unless -IncludeSystemAccounts.
$systemPatterns = @(
    '\\vpxd-', '\\vpxd-extension-', '\\vsphere-webclient-', '\\machine-', '\\vsan-health-',
    '\\vStatsGroup$', '\\ComponentManager\.', '\\SyncUsers$', '\\SystemConfiguration\.',
    '\\wcp-', '\\hvc-', '\\infraprofile-', '\\eam-', '\\vlcm-', '\\TrustedAdmins$',
    '\\RegistryAdministrators$', '\\NsxAdministrators$', '\\NsxAuditors$', '\\NsxViAdministrators$'
)
function Test-SystemAccount([string]$principal) {
    foreach ($p in $systemPatterns) { if ($principal -match $p) { return $true } }
    return $false
}

# Splits 'DOMAIN\name' or 'name@domain' into its parts.
function Split-Principal([string]$principal) {
    if ($principal -match '^(?<d>[^\\]+)\\(?<n>.+)$') { return @{ Domain = $Matches.d; Name = $Matches.n } }
    if ($principal -match '^(?<n>[^@]+)@(?<d>.+)$')   { return @{ Domain = $Matches.d; Name = $Matches.n } }
    return @{ Domain = ''; Name = $principal }
}

$permRows  = [System.Collections.Generic.List[object]]::new()
$errorRows = [System.Collections.Generic.List[object]]::new()
$vcNames   = [System.Collections.Generic.List[string]]::new()

foreach ($vc in @($global:DefaultVIServers)) {
    $vcName = $vc.Name
    $vcNames.Add($vcName)
    Write-Host "[$vcName] Reading permissions ..." -ForegroundColor Cyan

    try {
        $perms = @(Get-VIPermission -Server $vc -ErrorAction Stop)
    } catch {
        Write-Warning "[$vcName] Get-VIPermission failed: $($_.Exception.Message)"
        $errorRows.Add([pscustomobject]@{ vCenter = $vcName; Principal = '(all)'; Error = "Get-VIPermission failed: $($_.Exception.Message)" })
        continue
    }

    # Role -> privilege count, so 'Admin'-like custom roles are easy to spot.
    $roleInfo = @{}
    try {
        foreach ($r in @(Get-VIRole -Server $vc -ErrorAction Stop)) {
            $roleInfo[$r.Name] = @{ System = $r.IsSystem; Privs = @($r.PrivilegeList).Count }
        }
    } catch {
        Write-Warning "[$vcName] Get-VIRole failed (privilege counts will be blank): $($_.Exception.Message)"
    }

    $hidden = 0
    foreach ($p in $perms) {
        $principal = "$($p.Principal)"
        if (-not $IncludeSystemAccounts -and (Test-SystemAccount $principal)) { $hidden++; continue }
        $parts = Split-Principal $principal
        $ri    = $roleInfo["$($p.Role)"]
        $permRows.Add([pscustomobject][ordered]@{
            vCenter        = $vcName
            Principal      = $principal
            Domain         = $parts.Domain
            Name           = $parts.Name
            Type           = if ($p.IsGroup) { 'Group' } else { 'User' }
            Role           = "$($p.Role)"
            RolePrivileges = if ($ri) { $ri.Privs } else { '' }
            EntityType     = if ($p.Entity) { ($p.Entity.GetType().Name -replace 'Impl$', '') } else { '' }
            Entity         = if ($p.Entity) { "$($p.Entity.Name)" } else { "$($p.EntityId)" }
            Propagate      = [bool]$p.Propagate
        })
    }
    $kept = @($permRows | Where-Object { $_.vCenter -eq $vcName }).Count
    Write-Host "[$vcName]   Permissions: $kept$(if ($hidden) { " (+$hidden built-in system accounts hidden)" })"
}

#region ---- Principals (one row per user/group per vCenter) ----
$principalRows = foreach ($g in ($permRows | Group-Object vCenter, Principal)) {
    $f = $g.Group[0]
    $roles = @($g.Group | ForEach-Object Role | Sort-Object -Unique)
    [pscustomobject][ordered]@{
        vCenter     = $f.vCenter
        Principal   = $f.Principal
        Domain      = $f.Domain
        Type        = $f.Type
        Roles       = $roles -join '; '
        IsAdmin     = [bool]($roles -contains 'Admin')
        ObjectCount = $g.Count
        Objects     = (@($g.Group | ForEach-Object { "$($_.EntityType):$($_.Entity) [$($_.Role)]" }) -join '; ')
    }
}
$principalRows = @($principalRows | Sort-Object vCenter, Type, Principal)
#endregion

#region ---- Matrix (one row per user/group, one column per vCenter) ----
$matrixRows = foreach ($g in ($principalRows | Group-Object Principal)) {
    $o = [ordered]@{ Principal = $g.Name; Type = $g.Group[0].Type }
    foreach ($n in $vcNames) {
        $hit = @($g.Group | Where-Object { $_.vCenter -eq $n })
        $o[$n] = if ($hit.Count -gt 0) { $hit[0].Roles }
                 elseif ($errorRows | Where-Object vCenter -eq $n) { '(unable to read)' }
                 else { '' }
    }
    $o['vCenterCount'] = $g.Count
    [pscustomobject]$o
}
$matrixRows = @($matrixRows | Sort-Object Type, Principal)
#endregion

#region ---- Optional: expand groups to users ----
$memberRows = [System.Collections.Generic.List[object]]::new()
if ($ExpandGroups) {
    $haveAD  = [bool](Get-Module -ListAvailable -Name ActiveDirectory)
    $haveSso = [bool](Get-Command Get-SsoGroup -ErrorAction SilentlyContinue) -and
               $global:DefaultSsoAdminServers -and @($global:DefaultSsoAdminServers).Count -gt 0
    if ($haveAD) { Import-Module ActiveDirectory -ErrorAction SilentlyContinue }

    $cache = @{}   # Principal -> list of members (same group often appears on several vCenters)
    foreach ($grp in @($principalRows | Where-Object Type -eq 'Group')) {
        if (-not $cache.ContainsKey($grp.Principal)) {
            $parts   = Split-Principal $grp.Principal
            $members = @()
            $note    = ''
            try {
                # SSO domain is vsphere.local by default; a custom SSO domain is detected by asking SSO.
                $sg = $null
                if ($haveSso) { $sg = Get-SsoGroup -Name $parts.Name -Domain $parts.Domain -ErrorAction SilentlyContinue | Select-Object -First 1 }
                if ($sg -or $parts.Domain -ieq 'vsphere.local') {
                    if (-not $haveSso) { throw 'SSO group: needs VMware.vSphere.SsoAdmin module and Connect-SsoAdminServer.' }
                    if (-not $sg) { throw "SSO group '$($parts.Name)' not found in $($parts.Domain)." }
                    $members = @(Get-SsoPersonUser -Group $sg -ErrorAction Stop |
                                 ForEach-Object { [pscustomobject]@{ User = "$($_.Name)@$($_.Domain)"; Display = "$($_.FirstName) $($_.LastName)".Trim(); Enabled = -not $_.Disabled; Source = 'SSO' } })
                } else {
                    if (-not $haveAD) { throw 'AD group: needs the RSAT ActiveDirectory module on this machine.' }
                    $srv = @{}
                    if ($parts.Domain) { $srv.Server = $parts.Domain }
                    $members = @(Get-ADGroupMember -Identity $parts.Name -Recursive -ErrorAction Stop @srv | Where-Object objectClass -eq 'user' | ForEach-Object {
                        $u = Get-ADUser $_ -Properties Enabled, DisplayName -ErrorAction SilentlyContinue @srv
                        [pscustomobject]@{ User = "$($parts.Domain)\$($_.SamAccountName)"; Display = "$($u.DisplayName)"; Enabled = $u.Enabled; Source = 'AD' }
                    })
                }
                if ($members.Count -eq 0) { $note = 'Group is empty.' }
            } catch {
                $note = "Unable to expand: $($_.Exception.Message)"
            }
            $cache[$grp.Principal] = @{ Members = $members; Note = $note }
        }

        $c = $cache[$grp.Principal]
        if ($c.Members.Count -eq 0) {
            $memberRows.Add([pscustomobject][ordered]@{ vCenter = $grp.vCenter; Group = $grp.Principal; Roles = $grp.Roles; MemberUser = ''; DisplayName = ''; Enabled = ''; Source = ''; Notes = $c.Note })
        }
        foreach ($m in $c.Members) {
            $memberRows.Add([pscustomobject][ordered]@{ vCenter = $grp.vCenter; Group = $grp.Principal; Roles = $grp.Roles; MemberUser = $m.User; DisplayName = $m.Display; Enabled = $m.Enabled; Source = $m.Source; Notes = '' })
        }
    }
    Write-Host "Group members expanded: $(@($memberRows | Where-Object MemberUser).Count) user row(s)."
}
#endregion

#region ---- Export ----
$stamp  = Get-Date -Format 'yyyyMMdd_HHmm'
$sheets = [ordered]@{
    Matrix      = $matrixRows
    Principals  = $principalRows
    Permissions = @($permRows | Sort-Object vCenter, Principal, Entity)
}
if ($ExpandGroups)        { $sheets.Add('GroupMembers', @($memberRows)) }
if ($errorRows.Count -gt 0) { $sheets.Add('Errors', @($errorRows)) }

foreach ($k in $sheets.Keys) {
    if (@($sheets[$k]).Count -eq 0) { continue }
    $csv = Join-Path $OutputFolder "vCenter_UserAccess_${k}_$stamp.csv"
    $sheets[$k] | Export-Csv -Path $csv -NoTypeInformation -Encoding UTF8
    Write-Host "CSV saved : $csv" -ForegroundColor Green
}

if (Get-Module -ListAvailable -Name ImportExcel) {
    $xlsxPath = Join-Path $OutputFolder "vCenter_UserAccess_$stamp.xlsx"
    foreach ($k in $sheets.Keys) {
        if (@($sheets[$k]).Count -eq 0) { continue }
        $sheets[$k] | Export-Excel -Path $xlsxPath -WorksheetName $k -AutoSize -AutoFilter -FreezeTopRow -BoldTopRow
    }
    Write-Host "Excel saved: $xlsxPath" -ForegroundColor Green
} else {
    Write-Host '(Install-Module ImportExcel to also get a single .xlsx with one sheet per view.)' -ForegroundColor DarkGray
}

Write-Host "`nReminder: Global Permissions are not visible to Get-VIPermission - check them in the vSphere Client." -ForegroundColor Yellow

# Summary
$failed = @($errorRows | ForEach-Object vCenter)
$summary = foreach ($n in $vcNames) {
    $rows = @($principalRows | Where-Object vCenter -eq $n)
    [pscustomobject][ordered]@{
        vCenter = $n
        Status  = if ($failed -contains $n) { 'FAILED - see Errors' } else { 'OK' }
        Users   = @($rows | Where-Object Type -eq 'User').Count
        Groups  = @($rows | Where-Object Type -eq 'Group').Count
        Admins  = @($rows | Where-Object IsAdmin).Count
    }
}
$summary | Format-Table -AutoSize | Out-String | Write-Host
#endregion
