# Azure → GCP AD Migration – Working Log

Running record of the real migration: environment facts, steps done, problems hit and how each
was solved. This is the source for the final runbook.

Last updated: 2026-11-05

---

## 1. Environment facts (verified)

| Item | Value | How verified |
|---|---|---|
| AD domain | `amccinemasksa.net` | `Resolve-DnsName`, `netdom` |
| Azure DC (healthy, writable) | `p-amcsccl-adc` = **10.7.30.8** | `repadmin /replsummary` 0 fails; `ipconfig` |
| FSMO (before any change) | Schema, Domain Naming, PDC, Infrastructure = `p-amcsccl-adc`; **RID = `KSA-HOS-AD`** | `netdom query fsmo` |
| Dead writable DCs | `KSA-HOS-AD`, `KSA-HOS-BD`: 100% replication failure, >60 days, error 1722 (RPC unavailable) | `repadmin /replsummary` |
| Domain A records also returned | 10.7.7.13, 10.7.7.14 (probably the dead KSA-HOS DCs, **not yet confirmed**) | `Resolve-DnsName amccinemasksa.net` |
| RODCs (to be decommissioned later) | KSA-KAF/HAB/ALK/AZZ/KAR/DAW-RODC: unreachable, error 58 | `repadmin /replsummary` |
| GCP server | `sncgd-aad-app01`, 172.18.51.3 /28, GW 172.18.51.1, DHCP from 169.254.169.254 | `ipconfig /all` |
| GCP server OS | Windows Server 2025 (build 10.0.26100) | cmd banner |
| GCP IP type | **Static** (reserved internal IP) | GCP Console → VPC network → IP addresses |
| GCP project / zone | `snc-aad-pj-p01` / `me-central2-a` | DNS suffix, console |
| Access method | IAP Desktop (RDP over IAP) | — |

## 2. Steps completed

| # | Step | Result |
|---|---|---|
| 1 | Health check of Azure DC (`repadmin /replsummary`, `netdom query fsmo`) | Azure DC healthy; found dead KSA-HOS DCs and RID on KSA-HOS-AD |
| 2 | Confirmed GCP IP is static in GCP Console | Static ✅ (Windows still shows DHCP; this is expected on GCP) |
| 3 | Set GCP server DNS = 10.7.30.8 (IP left on automatic) | ✅ `ipconfig /all` shows DNS 10.7.30.8 |
| 4 | Port tests → firewall request → re-test | TCP opened; UDP 389/123 added in a second request |
| 5 | Installed AD DS + DNS roles on GCP server | ✅ Success (static IP warning is expected) |
| 6 | Promotion attempts | Failed: (a) DC not contactable, then (b) needs Enterprise + Schema Admins. See section 3 |

## 3. Problems and solutions

| # | Symptom / error | Cause | Solution |
|---|---|---|---|
| P1 | `repadmin /replsummary`: KSA-HOS-AD/BD 100% fail, >60 days, **1722** | Those DCs are offline or decommissioned | Ignore for the migration; use `-ReplicationSourceDC p-amcsccl-adc`; clean up later (metadata cleanup) |
| P2 | RODCs show error **58** | Not reachable from the Azure DC | Not a blocker; decommission later |
| P3 | **RID master on dead KSA-HOS-AD** | Role holder offline, so it can't be transferred | Seize RID to `p-amcsccl-adc` (`Move-ADDirectoryServerOperationMasterRole ... RIDMaster -Force`) **after confirming KSA-HOS-AD is permanently gone**; get approval |
| P4 | IP shows in `ipconfig` but not in the IPv4 dialog | Windows IP is DHCP; GCP hands out the reserved address | Keep IP automatic; reserve the IP in GCP (verified Static) |
| P5 | How to set DNS while IP is DHCP | IP and DNS are separate settings | IPv4 dialog: keep "Obtain IP automatically", set "Use the following DNS" = 10.7.30.8; Alternate empty until promoted |
| P6 | `ping 10.7.30.8` timed out | ICMP is blocked; this doesn't prove the AD ports are blocked | Test the TCP ports instead |
| P7 | `Test-NetConnection` very slow | Waits ~20–30 s for each blocked port | Use the fast TcpClient test (1 s per port), section 5 |
| P8 | First port list was incomplete (missed UDP and RPC 49152–65535) | Oversimplified in the short steps | Use the full port matrix, section 4 |
| P9 | All ports BLOCKED | Firewall/tunnel rules missing | Firewall team opened TCP 53/88/135/389/445/464/636/3268/3269/9389/49152–65535, UDP 53/88/464 |
| P10 | `nltest /dsgetdc` → **1355 ERROR_NO_SUCH_DOMAIN**; promotion → *"domain controller … could not be contacted"* | **UDP 389** (LDAP ping / DC locator) blocked; TCP 389 alone isn't enough | Firewall added UDP 389 |
| P11 | `w32tm /stripchart` → **0x800705B4** (timeout) | **UDP 123** (NTP) blocked | Firewall added UDP 123 |
| P12 | `-SiteName '<same site as p-amcsccl-adc>'` | Placeholder text used literally | Get the real name: `Get-ADDomainController p-amcsccl-adc \| Select Site` |
| P13 | Promotion → *"user is not a member of Enterprise Admins and Schema Admins"* | New DC is **Windows Server 2025**, newer than the existing DCs, so promotion must run **adprep** (forest schema update) | Use an account in Domain Admins + Enterprise Admins + Schema Admins; take a system state backup of p-amcsccl-adc and get approval first (schema change is forest-wide and irreversible); remove the temporary group membership afterwards |

## 4. Required ports (GCP 172.18.51.3 ↔ Azure 10.7.30.8, both directions)

| Port | Protocol | Purpose |
|---|---|---|
| 53 | TCP + UDP | DNS |
| 88 | TCP + UDP | Kerberos |
| 123 | **UDP** | Time (NTP) |
| 135 | TCP | RPC endpoint mapper |
| 389 | TCP + **UDP** | LDAP + DC locator (UDP) |
| 445 | TCP | SMB (SYSVOL / GPO) |
| 464 | TCP + UDP | Kerberos password change |
| 636 | TCP | LDAPS |
| 3268 / 3269 | TCP | Global Catalog |
| 9389 | TCP | AD Web Services |
| 49152–65535 | TCP | RPC dynamic (replication, DFSR) |

## 5. Useful test commands

```powershell
# Fast TCP test (1 s per port). Add RPC ports found on the Azure DC to $rpc
$ip  = '10.7.30.8'
$rpc = @()
$ports = @{53='DNS';88='Kerberos';135='RPC-EPM';389='LDAP';445='SMB';464='Kpasswd';636='LDAPS';3268='GC';3269='GC-SSL';9389='ADWS'}
foreach ($r in $rpc) { $ports[[int]$r] = 'RPC-dynamic' }
foreach ($p in ($ports.Keys | Sort-Object)) {
  $c = New-Object Net.Sockets.TcpClient
  try { $ok = $c.ConnectAsync($ip,$p).Wait(1000) -and $c.Connected } catch { $ok = $false }
  '{0,-6} {1,-12} {2}' -f $p, $ports[$p], $(if ($ok) {'OPEN'} else {'BLOCKED'})
  $c.Close()
}

# RPC dynamic ports in use on the Azure DC (run ON the Azure DC)
$p = @((Get-Process lsass).Id); try { $p += (Get-Process dfsrs -ErrorAction Stop).Id } catch {}
(Get-NetTCPConnection -State Listen | ? { $p -contains $_.OwningProcess -and $_.LocalPort -ge 49152 }).LocalPort | Sort-Object -Unique

# UDP checks
Resolve-DnsName amccinemasksa.net -Server 10.7.30.8 -DnsOnly     # UDP 53
nltest /dsgetdc:amccinemasksa.net                                 # UDP 389, must return p-amcsccl-adc
w32tm /stripchart /computer:10.7.30.8 /samples:1 /dataonly        # UDP 123, must show an offset

# Look up an error code
net helpmsg 1355
certutil -error 0x800705B4
```

## 6. Open items / decisions

- [ ] Confirm `sncgd-aad-app01` is the server meant to become the DC (the name suggests an app server)
- [ ] Get the real site name of `p-amcsccl-adc`
- [ ] Account with Domain + Enterprise + Schema Admins (temporary)
- [ ] System state backup of `p-amcsccl-adc` before adprep
- [ ] Change approval for the schema update (adprep) and the RID seizure
- [ ] Confirm KSA-HOS-AD is permanently gone → seize RID to `p-amcsccl-adc` (status: **not yet confirmed done**)
- [ ] Re-run promotion → verify replication → transfer FSMO → time config
- [ ] Later: metadata cleanup of KSA-HOS-AD/BD, decommission RODCs, second GCP DC, retire Azure DC

## 7. Remaining planned steps (current version)

1. Promote: `Install-ADDSDomainController -DomainName amccinemasksa.net -InstallDns -SiteName '<REAL SITE>' -ReplicationSourceDC p-amcsccl-adc.amccinemasksa.net -Credential (Get-Credential) -SafeModeAdministratorPassword (Read-Host -AsSecureString 'DSRM')`
2. After reboot: DNS = 10.7.30.8, 127.0.0.1
3. Verify: `repadmin /showrepl`, `dcdiag /q`, `net share` (SYSVOL/NETLOGON), DFSR event 4604
4. FSMO: seize RID → p-amcsccl-adc (if not done), then transfer all 5 → GCP DC (no `-Force`)
5. Time: GCP DC → `metadata.google.internal`; Azure DC → `domhier`
6. Final check: `repadmin /showrepl ; dcdiag /q ; netdom query fsmo`
