# Get-FCHBAPathState

> Audit Fibre Channel HBA path states across your entire VMware vCenter environment — per cluster, color-coded, with CSV and dark-mode HTML export.

---

## Overview

`Get-FCHBAPathState.ps1` connects to a vCenter Server and checks the **Active**, **Dead**, and **Standby** path counts for every Fibre Channel HBA on every connected ESXi host. Results are grouped by cluster with a color-coded console table and optional CSV and HTML exports.

Key features:

- **Auto-detects** all FC HBAs in the environment — no hardcoded device names
- **Per-cluster output** with individual cluster summaries and an overall health total
- **Color-coded rows** — green for healthy, red for dead paths, yellow for standby-only
- **Dark-mode HTML report** — self-contained, auto-generated alongside CSV or via `-HTMLReport`
- **Credential store** — saves/loads vCenter credentials securely using DPAPI encryption
- **Optional HBA rescan** before collecting data
- **PowerCLI 13.x compatible** — runtime module check instead of `#Requires`
- **Single file** — no companion scripts required

---

## Requirements

| Requirement | Detail |
|---|---|
| PowerShell | 5.1 or 7+ |
| VMware PowerCLI | 12.x or 13.x (`VMware.VimAutomation.Core`) |
| vCenter | vCenter Server 7.x / 8.x |
| Permissions | Read access to hosts and storage |

Install PowerCLI if not already present:

```powershell
Install-Module VMware.PowerCLI -Scope CurrentUser
```

---

## Usage

### Basic — interactive, check all FC HBAs

```powershell
.\Get-FCHBAPathState.ps1
```

### Filter to specific HBAs

```powershell
.\Get-FCHBAPathState.ps1 -HBAFilter "vmhba4,vmhba5"
```

### Export CSV (HTML report auto-generated alongside it)

```powershell
.\Get-FCHBAPathState.ps1 -ExportPath "C:\Reports\FCHBAPathState.csv"
# → C:\Reports\FCHBAPathState.csv
# → C:\Reports\FCHBAPathState.html  (auto)
```

### Generate HTML report only

```powershell
.\Get-FCHBAPathState.ps1 -HTMLReport "C:\Reports\FCHBAPathState.html"
```

### Rescan HBAs before collecting

```powershell
.\Get-FCHBAPathState.ps1 -Rescan -ExportPath "C:\Reports\FCHBAPathState.csv"
```

### Fully automated (no prompts after first run)

```powershell
.\Get-FCHBAPathState.ps1 -HBAFilter "vmhba4,vmhba5" -ExportPath "C:\Reports\out.csv"
```

> **Note:** vCenter FQDN/IP and credentials are always prompted interactively unless credentials are already saved from a previous run.

---

## Parameters

| Parameter | Type | Required | Description |
|---|---|---|---|
| `-HBAFilter` | `string` | No | Comma-separated HBA device names to check. If omitted, ALL FC HBAs are checked. Can also be entered interactively when prompted. |
| `-ExportPath` | `string` | No | Full path for CSV output. A dark-mode HTML report is auto-generated at the same path with a `.html` extension. |
| `-HTMLReport` | `string` | No | Full path for a standalone dark-mode HTML report, independent of `-ExportPath`. |
| `-Rescan` | `switch` | No | Triggers `Get-VMHostStorage -RescanAllHba` on each host before data collection. |

---

## Console Output

```
  ╔═══════════════════════════════════════════╗
  ║     Get-FCHBAPathState  v2.6              ║
  ║     Paul van Dieen - hollebollevsan.nl    ║
  ╚═══════════════════════════════════════════╝

  Cluster: Cluster-Production
  +----------------------+--------+--------+------+---------+
  | VMHost               | HBA    | Active | Dead | Standby |
  +----------------------+--------+--------+------+---------+
  | esxi01.prod.local    | vmhba4 | 8      | 0    | 0       |   ← Green
  | esxi02.prod.local    | vmhba4 | 6      | 2    | 0       |   ← Red
  | esxi02.prod.local    | vmhba5 | 8      | 0    | 0       |   ← Green
  +----------------------+--------+--------+------+---------+
  [!] 1 HBAs with dead paths in this cluster.

  Legend: Green = Active   Yellow = Standby only   Red = Dead paths
```

| Color | Meaning |
|---|---|
| 🟢 Green | Active paths present, no issues |
| 🔴 Red | One or more dead paths detected |
| 🟡 Yellow | Standby paths only, no active paths |
| 🟠 Dark Yellow | No paths found in any state |

---

## HTML Report

When `-ExportPath` or `-HTMLReport` is used, the script generates a self-contained dark-mode HTML report with:

- Summary stat cards (clusters, hosts, HBAs checked, healthy, dead paths)
- Overall health status banner (green / red)
- Per-cluster tables with color-coded status chips per row
- Glowing host status indicators
- Legend and footer with author branding

The HTML file is fully standalone — no server or external runtime dependencies beyond Google Fonts for typography.

---

## Credential Store

On first run the script prompts for username and password directly in the console. You can choose to save them for future runs. Credentials are stored at:

```
%USERPROFILE%\.vcenter_creds
```

The password is encrypted using Windows DPAPI via `ConvertFrom-SecureString` — it is tied to your Windows user account and machine and cannot be decrypted by another user or on another machine.

To clear saved credentials:

```powershell
Remove-Item "$env:USERPROFILE\.vcenter_creds"
```

---

## Version History

| Version | Date | Changes |
|---|---|---|
| 2.6 | 2026-03-19 | PS 5.1 compatibility fixes (UTF-8 BOM, character escaping); console credential prompts; single-file deployment |
| 2.5 | 2026-03-19 | Added `-HTMLReport` parameter; dark-mode HTML report auto-generated alongside CSV |
| 2.4 | 2026-03-06 | Per-cluster output tables with cluster summaries and overall total |
| 2.3 | 2026-03-06 | Colored console table, dynamic column widths, summary line, legend |
| 2.2 | 2026-03-06 | Replaced `#Requires` with runtime PowerCLI check for 13.x compatibility |
| 2.1 | 2026-03-06 | Auto-detect all FC HBAs; added `-HBAFilter` and interactive filter prompt |
| 2.0 | 2026-03-06 | Full rewrite: vCenter prompt, credential save/load, dead-path alerting, CSV export, `-Rescan` |
| 1.0 | initial | Original: hardcoded vCenter, basic FC path check |

---

## Author

**Paul van Dieen**
[hollebollevsan.nl](https://www.hollebollevsan.nl)
