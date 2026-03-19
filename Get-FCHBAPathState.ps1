<#
.SYNOPSIS
    Get-FCHBAPathState — Reports Fibre Channel HBA path states across all connected ESXi hosts.

.DESCRIPTION
    Connects to a user-specified vCenter server and audits all Fibre Channel HBAs on every
    connected ESXi host. Auto-detects HBAs per host by default; optionally filter to specific
    devices via -HBAFilter. Reports Active, Dead, and Standby path counts per HBA.
    Supports credential save/load, optional HBA rescan, CSV export, and dark-mode HTML report.

.PARAMETER HBAFilter
    Optional. Comma-separated list of HBA device names to check (e.g. "vmhba4,vmhba5").
    If omitted, ALL Fibre Channel HBAs on each host are checked automatically.
    Can also be passed at runtime when prompted.

.PARAMETER ExportPath
    Optional. Full path to export results as a CSV file.
    An HTML report is automatically saved alongside the CSV (same path, .html extension).
    Example: -ExportPath "C:\Reports\FCHBAPathState.csv"

.PARAMETER HTMLReport
    Optional. Full path to save a standalone dark-mode HTML report.
    If -ExportPath is also specified, an HTML report is generated automatically
    alongside the CSV regardless of this parameter.
    Example: -HTMLReport "C:\Reports\FCHBAPathState.html"

.PARAMETER Rescan
    Optional switch. If specified, triggers a rescan of all HBAs before collecting data.

.EXAMPLE
    .\Get-FCHBAPathState.ps1
    .\Get-FCHBAPathState.ps1 -HBAFilter "vmhba4,vmhba5"
    .\Get-FCHBAPathState.ps1 -ExportPath "C:\Reports\FCHBAPathState.csv" -Rescan
    .\Get-FCHBAPathState.ps1 -HTMLReport "C:\Reports\FCHBAPathState.html"

.NOTES
    Author  : Paul van Dieen
    Blog    : https://www.hollebollevsan.nl
    Version : 2.5  (2026-03-06) — Added -HTMLReport parameter; auto-generates dark-mode
                                   HTML report alongside CSV when -ExportPath is used.
              2.4  (2026-03-06) — Output split into per-cluster tables with individual
                                   cluster summaries and an overall total summary.
              2.3  (2026-03-06) — Colored output table with dynamic column widths,
                                   per-row health coloring, summary line, and legend.
              2.2  (2026-03-06) — Replaced #Requires with runtime PowerCLI check for
                                   compatibility with PowerCLI 13.x module structure.
              2.1  (2026-03-06) — Auto-detect all FC HBAs; added -HBAFilter parameter
                                   and interactive filter prompt at runtime.
              2.0  (2026-03-06) — Rewrite: vCenter prompt, credential save/load,
                                   PSCustomObject collection, dead-path alerting,
                                   error handling, CSV export, -Rescan switch.
              1.0  (initial)    — Original: hardcoded vCenter, basic FC path check.
#>
param(
    [string]$HBAFilter   = "",
    [string]$ExportPath  = "",
    [string]$HTMLReport  = "",
    [switch]$Rescan
)

# ─────────────────────────────────────────────
#  POWERCLI VERSION CHECK
# ─────────────────────────────────────────────
$powercli = Get-Module -ListAvailable -Name VMware.VimAutomation.Core | Sort-Object Version -Descending | Select-Object -First 1
if (-not $powercli) {
    Write-Error "VMware PowerCLI does not appear to be installed. Please install it with: Install-Module VMware.PowerCLI"
    exit 1
}
Import-Module VMware.VimAutomation.Core -ErrorAction SilentlyContinue

# ─────────────────────────────────────────────
#  CREDENTIAL STORE HELPERS
# ─────────────────────────────────────────────
$credStorePath = "$env:USERPROFILE\.vcenter_creds"

function Save-VCenterCredential {
    param([PSCredential]$Credential)
    $export = [PSCustomObject]@{
        Username = $Credential.UserName
        Password = $Credential.Password | ConvertFrom-SecureString
    }
    $export | Export-Clixml -Path $credStorePath
    Write-Host "  Credentials saved to $credStorePath" -ForegroundColor DarkGray
}

function Load-VCenterCredential {
    if (Test-Path $credStorePath) {
        try {
            $import = Import-Clixml -Path $credStorePath
            $securePass = $import.Password | ConvertTo-SecureString
            return New-Object System.Management.Automation.PSCredential($import.Username, $securePass)
        } catch {
            Write-Warning "Saved credentials could not be loaded. You will be prompted."
            return $null
        }
    }
    return $null
}

# ─────────────────────────────────────────────
#  HTML REPORT GENERATOR
# ─────────────────────────────────────────────
function New-FCHBAHTMLReport {
    param(
        [System.Collections.Generic.List[PSObject]]$Results,
        [string]$VCenter,
        [string]$OutputPath
    )

    $generatedAt  = (Get-Date).ToString("dd MMM yyyy  HH:mm:ss")
    $totalClusters = ($Results | Select-Object -ExpandProperty Cluster -Unique).Count
    $totalHosts    = ($Results | Select-Object -ExpandProperty VMHost  -Unique).Count
    $totalHBAs     = $Results.Count
    $totalHealthy  = ($Results | Where-Object { $_.Dead -eq 0 }).Count
    $totalDead     = ($Results | Where-Object { $_.Dead -gt 0 }).Count

    if ($totalDead -gt 0) {
        $statusClass   = "crit"
        $statusMsg     = "&#9888; $totalDead HBA(s) have dead paths &mdash; investigate immediately."
    } else {
        $statusClass   = "ok"
        $statusMsg     = "&#10003; All HBAs across all clusters reporting healthy paths."
    }

    # ── Build per-cluster table HTML ──
    $clusterHTML = ""
    $clusters = $Results | Select-Object -ExpandProperty Cluster -Unique | Sort-Object

    foreach ($cluster in $clusters) {
        $rows = $Results | Where-Object { $_.Cluster -eq $cluster } | Sort-Object VMHost, HBA
        $clusterDead = ($rows | Where-Object { $_.Dead -gt 0 }).Count

        if ($clusterDead -gt 0) {
            $cpillClass = "crit"
            $cpillText  = "$clusterDead dead path(s)"
        } else {
            $cpillClass = "ok"
            $cpillText  = "All healthy"
        }

        $rowHTML = ""
        foreach ($row in $rows) {
            if ($row.Dead -gt 0) {
                $rowClass  = "row-dead"
                $dotClass  = "dead"
                $statusChip = '<span class="chip dead">&#10005; Dead paths</span>'
            } elseif ($row.Standby -gt 0 -and $row.Active -eq 0) {
                $rowClass  = "row-warn"
                $dotClass  = "warn"
                $statusChip = '<span class="chip standby">&#9873; Standby only</span>'
            } elseif ($row.Active -gt 0) {
                $rowClass  = "row-ok"
                $dotClass  = "ok"
                $statusChip = '<span class="chip active">&#10003; Healthy</span>'
            } else {
                $rowClass  = "row-nopath"
                $dotClass  = "warn"
                $statusChip = '<span class="chip standby">No paths</span>'
            }

            $activeChip  = if ($row.Active  -gt 0) { '<span class="chip active">'  + $row.Active  + '</span>' } else { '<span class="chip zero">0</span>' }
            $deadChip    = if ($row.Dead    -gt 0) { '<span class="chip dead">'    + $row.Dead    + '</span>' } else { '<span class="chip zero">0</span>' }
            $standbyChip = if ($row.Standby -gt 0) { '<span class="chip standby">' + $row.Standby + '</span>' } else { '<span class="chip zero">0</span>' }

            $rowHTML += @"
        <tr class="$rowClass">
          <td><div class="host-cell"><div class="host-dot $dotClass"></div>$($row.VMHost)</div></td>
          <td><span class="hba-device">$($row.HBA)</span></td>
          <td class="num">$activeChip</td>
          <td class="num">$deadChip</td>
          <td class="num">$standbyChip</td>
          <td>$statusChip</td>
        </tr>
"@
        }

        $clusterHTML += @"
  <div class="cluster-block">
    <div class="cluster-heading">
      <span class="cluster-name">&#9672; $cluster</span>
      <span class="cluster-pill $cpillClass">$cpillText</span>
    </div>
    <div class="table-wrap">
    <table class="hba-table">
      <thead><tr>
        <th>VMHost</th><th>HBA</th>
        <th class="num">Active</th><th class="num">Dead</th><th class="num">Standby</th>
        <th>Status</th>
      </tr></thead>
      <tbody>
$rowHTML
      </tbody>
    </table>
    </div>
  </div>
"@
    }

    # ── Assemble full HTML ──
    $html = @"
<!DOCTYPE html>
<html lang="en">
<head>
<meta charset="UTF-8">
<meta name="viewport" content="width=device-width, initial-scale=1.0">
<title>Get-FCHBAPathState &mdash; FC HBA Path Report</title>
<style>
  @import url('https://fonts.googleapis.com/css2?family=JetBrains+Mono:wght@400;600;700&family=Syne:wght@400;600;700;800&display=swap');
  :root {
    --bg:#0d0f14;--bg2:#12151c;--bg3:#181c26;--border:#252a38;--border2:#2e3448;
    --text:#c8cfe0;--text-dim:#5a6480;--text-head:#e8edf8;
    --accent:#3d9cf0;--accent2:#5ab4ff;
    --green:#3dd68c;--green-bg:rgba(61,214,140,.08);--green-bdr:rgba(61,214,140,.25);
    --yellow:#f0c040;--yellow-bg:rgba(240,192,64,.08);--yellow-bdr:rgba(240,192,64,.25);
    --red:#f05060;--red-bg:rgba(240,80,96,.08);--red-bdr:rgba(240,80,96,.30);
    --magenta:#c080f0;
    --mono:'JetBrains Mono',monospace;--sans:'Syne',sans-serif;
  }
  *,*::before,*::after{box-sizing:border-box;margin:0;padding:0}
  body{background:var(--bg);color:var(--text);font-family:var(--mono);font-size:13px;line-height:1.6;min-height:100vh;padding:0 0 60px}
  .page-header{background:var(--bg2);border-bottom:1px solid var(--border2);padding:28px 40px 24px;display:flex;align-items:flex-start;justify-content:space-between;gap:24px;position:relative;overflow:hidden}
  .page-header::before{content:'';position:absolute;top:0;left:0;right:0;height:2px;background:linear-gradient(90deg,var(--accent) 0%,var(--magenta) 50%,var(--green) 100%)}
  .header-left{display:flex;flex-direction:column;gap:6px}
  .header-tool{font-family:var(--sans);font-size:22px;font-weight:800;color:var(--text-head);letter-spacing:-0.5px}
  .header-tool span{color:var(--accent2)}
  .header-subtitle{color:var(--text-dim);font-size:11px;letter-spacing:0.5px}
  .header-meta{display:flex;flex-direction:column;align-items:flex-end;gap:4px;font-size:11px;color:var(--text-dim)}
  .badge{display:inline-flex;align-items:center;gap:5px;background:var(--bg3);border:1px solid var(--border2);border-radius:4px;padding:3px 8px;font-size:11px;color:var(--text-dim)}
  .badge.v{color:var(--accent);border-color:rgba(61,156,240,.3);background:rgba(61,156,240,.06)}
  .summary-bar{display:flex;gap:12px;padding:16px 40px;background:var(--bg2);border-bottom:1px solid var(--border);flex-wrap:wrap}
  .stat-card{display:flex;flex-direction:column;gap:2px;background:var(--bg3);border:1px solid var(--border);border-radius:6px;padding:10px 18px;min-width:110px}
  .stat-label{font-size:10px;color:var(--text-dim);letter-spacing:0.8px;text-transform:uppercase}
  .stat-value{font-size:22px;font-weight:700;font-family:var(--sans);color:var(--text-head)}
  .stat-card.ok .stat-value{color:var(--green)}.stat-card.crit .stat-value{color:var(--red)}
  .status-banner{margin:20px 40px 0;border-radius:6px;padding:10px 16px;font-size:12px;font-weight:600;display:flex;align-items:center;gap:8px}
  .status-banner.ok{background:var(--green-bg);border:1px solid var(--green-bdr);color:var(--green)}
  .status-banner.crit{background:var(--red-bg);border:1px solid var(--red-bdr);color:var(--red)}
  .main{padding:20px 40px 0}
  .cluster-block{margin-bottom:32px}
  .cluster-heading{display:flex;align-items:center;gap:10px;margin-bottom:10px}
  .cluster-name{font-family:var(--sans);font-size:14px;font-weight:700;color:var(--magenta);letter-spacing:0.3px}
  .cluster-pill{font-size:10px;background:rgba(192,128,240,.1);border:1px solid rgba(192,128,240,.25);color:var(--magenta);border-radius:20px;padding:2px 8px}
  .cluster-pill.ok{background:var(--green-bg);border-color:var(--green-bdr);color:var(--green)}
  .cluster-pill.crit{background:var(--red-bg);border-color:var(--red-bdr);color:var(--red)}
  .hba-table{width:100%;border-collapse:collapse;border:1px solid var(--border2);border-radius:6px;overflow:hidden}
  .hba-table thead tr{background:var(--bg3);border-bottom:1px solid var(--border2)}
  .hba-table th{padding:9px 14px;text-align:left;font-size:10px;font-weight:600;letter-spacing:0.8px;text-transform:uppercase;color:var(--accent)}
  .hba-table th.num{text-align:right}
  .hba-table tbody tr{border-bottom:1px solid var(--border);transition:background 0.15s}
  .hba-table tbody tr:last-child{border-bottom:none}
  .hba-table tbody tr:hover{background:rgba(255,255,255,.025)}
  .hba-table td{padding:9px 14px;font-size:12px;color:var(--text)}
  .hba-table td.num{text-align:right;font-weight:600}
  .row-ok{background:rgba(61,214,140,.03)}.row-warn{background:rgba(240,192,64,.04)}
  .row-dead{background:rgba(240,80,96,.05)}.row-nopath{background:rgba(240,160,64,.04)}
  .chip{display:inline-flex;align-items:center;gap:4px;border-radius:4px;padding:2px 7px;font-size:11px;font-weight:600}
  .chip.active{background:var(--green-bg);border:1px solid var(--green-bdr);color:var(--green)}
  .chip.dead{background:var(--red-bg);border:1px solid var(--red-bdr);color:var(--red)}
  .chip.standby{background:var(--yellow-bg);border:1px solid var(--yellow-bdr);color:var(--yellow)}
  .chip.zero{background:var(--bg3);border:1px solid var(--border);color:var(--text-dim)}
  .host-cell{display:flex;align-items:center;gap:8px}
  .host-dot{width:6px;height:6px;border-radius:50%;flex-shrink:0}
  .host-dot.ok{background:var(--green);box-shadow:0 0 6px var(--green)}
  .host-dot.warn{background:var(--yellow);box-shadow:0 0 6px var(--yellow)}
  .host-dot.dead{background:var(--red);box-shadow:0 0 6px var(--red)}
  .hba-device{font-family:var(--mono);font-size:11px;color:var(--accent2);background:rgba(61,156,240,.08);border:1px solid rgba(61,156,240,.2);border-radius:4px;padding:1px 6px}
  .legend{display:flex;gap:16px;flex-wrap:wrap;margin:24px 40px 0;padding:12px 16px;background:var(--bg2);border:1px solid var(--border);border-radius:6px;font-size:11px;color:var(--text-dim);align-items:center}
  .legend-title{font-weight:600;color:var(--text-dim);letter-spacing:0.5px;text-transform:uppercase;font-size:10px}
  .legend-item{display:flex;align-items:center;gap:6px}
  .page-footer{margin:32px 40px 0;padding-top:16px;border-top:1px solid var(--border);display:flex;justify-content:space-between;align-items:center;font-size:11px;color:var(--text-dim)}
  .page-footer a{color:var(--accent);text-decoration:none}
  .table-wrap{overflow-x:auto;border-radius:6px}
</style>
</head>
<body>
<header class="page-header">
  <div class="header-left">
    <div class="header-tool">Get-FC<span>HBAPathState</span></div>
    <div class="header-subtitle">Fibre Channel HBA Path State Report &nbsp;&middot;&nbsp; $VCenter</div>
  </div>
  <div class="header-meta">
    <span class="badge v">v2.5</span>
    <span class="badge">Generated: $generatedAt</span>
  </div>
</header>
<div class="summary-bar">
  <div class="stat-card"><span class="stat-label">Clusters</span><span class="stat-value">$totalClusters</span></div>
  <div class="stat-card"><span class="stat-label">Hosts</span><span class="stat-value">$totalHosts</span></div>
  <div class="stat-card"><span class="stat-label">HBAs checked</span><span class="stat-value">$totalHBAs</span></div>
  <div class="stat-card ok"><span class="stat-label">Healthy</span><span class="stat-value">$totalHealthy</span></div>
  <div class="stat-card crit"><span class="stat-label">Dead paths</span><span class="stat-value">$totalDead</span></div>
</div>
<div class="status-banner $statusClass">$statusMsg</div>
<div class="main">
$clusterHTML
</div>
<div class="legend">
  <span class="legend-title">Legend</span>
  <span class="legend-item"><span class="chip active">Active</span> Paths routing I/O normally</span>
  <span class="legend-item"><span class="chip dead">Dead</span> Paths lost &mdash; immediate action required</span>
  <span class="legend-item"><span class="chip standby">Standby</span> Paths available but not routing I/O</span>
  <span class="legend-item"><span class="chip zero">0</span> No paths in this state</span>
</div>
<footer class="page-footer">
  <span>Get-FCHBAPathState v2.5 &nbsp;&middot;&nbsp; <a href="https://www.hollebollevsan.nl" target="_blank">hollebollevsan.nl</a> &nbsp;&middot;&nbsp; Paul van Dieen</span>
  <span>$generatedAt</span>
</footer>
</body>
</html>
"@

    try {
        [System.IO.File]::WriteAllText($OutputPath, $html, [System.Text.UTF8Encoding]::new($false))
        Write-Host "  HTML report saved to:  $OutputPath" -ForegroundColor Green
    } catch {
        Write-Warning "  HTML report export failed: $_"
    }
}

# ─────────────────────────────────────────────
#  VCENTER SERVER PROMPT
# ─────────────────────────────────────────────
Write-Host ""
Write-Host "╔═══════════════════════════════════════════╗" -ForegroundColor Cyan
Write-Host "║     Get-FCHBAPathState  v2.5              ║" -ForegroundColor Cyan
Write-Host "║     Paul van Dieen - hollebollevsan.nl    ║" -ForegroundColor Cyan
Write-Host "╚═══════════════════════════════════════════╝" -ForegroundColor Cyan
Write-Host ""

$vCenter = Read-Host "  Enter vCenter FQDN or IP"
if ([string]::IsNullOrWhiteSpace($vCenter)) {
    Write-Error "No vCenter specified. Exiting."
    exit 1
}

# ─────────────────────────────────────────────
#  CREDENTIAL HANDLING
# ─────────────────────────────────────────────
$savedCred = Load-VCenterCredential

if ($savedCred) {
    Write-Host ""
    Write-Host "  Found saved credentials for: " -NoNewline
    Write-Host $savedCred.UserName -ForegroundColor Yellow
    $useSaved = Read-Host "  Use saved credentials? (Y/N) [Y]"

    if ($useSaved -eq "" -or $useSaved -match "^[Yy]") {
        $cred = $savedCred
        Write-Host "  Using saved credentials." -ForegroundColor Green
    } else {
        $cred = Get-Credential -Message "Enter vCenter credentials for $vCenter"
        $saveNew = Read-Host "  Save these credentials for next time? (Y/N) [N]"
        if ($saveNew -match "^[Yy]") { Save-VCenterCredential -Credential $cred }
    }
} else {
    Write-Host ""
    $cred = Get-Credential -Message "Enter vCenter credentials for $vCenter"
    $saveNew = Read-Host "  Save these credentials for next time? (Y/N) [N]"
    if ($saveNew -match "^[Yy]") { Save-VCenterCredential -Credential $cred }
}

# ─────────────────────────────────────────────
#  CONNECT TO VCENTER
# ─────────────────────────────────────────────
Write-Host ""
Write-Host "  Connecting to $vCenter ..." -ForegroundColor Cyan

try {
    Connect-VIServer -Server $vCenter -Credential $cred -ErrorAction Stop | Out-Null
    Write-Host "  Connected successfully." -ForegroundColor Green
} catch {
    Write-Error "Failed to connect to $vCenter : $_"
    exit 1
}

# ─────────────────────────────────────────────
#  MAIN LOGIC
# ─────────────────────────────────────────────
Write-Host ""
Write-Host "  Gathering ESXi hosts..." -ForegroundColor Cyan

$VMHosts = Get-VMHost | Where-Object { $_.ConnectionState -eq "Connected" } | Sort-Object Name

if (-not $VMHosts) {
    Write-Warning "No connected hosts found. Disconnecting."
    Disconnect-VIServer * -Confirm:$false
    exit 0
}

Write-Host "  Found $($VMHosts.Count) connected host(s). Checking HBAs...`n" -ForegroundColor Cyan

# ─────────────────────────────────────────────
#  HBA FILTER — auto-detect, then optionally narrow
# ─────────────────────────────────────────────

# Collect all unique FC HBA device names across all hosts
$allFCHBAs = $VMHosts | Get-VMHostHba -Type FibreChannel |
             Select-Object -ExpandProperty Device -Unique | Sort-Object

if (-not $allFCHBAs) {
    Write-Warning "No Fibre Channel HBAs found on any connected host. Disconnecting."
    Disconnect-VIServer * -Confirm:$false
    exit 0
}

Write-Host "  Fibre Channel HBAs detected in this environment:" -ForegroundColor Cyan
$allFCHBAs | ForEach-Object { Write-Host "    - $_" -ForegroundColor White }
Write-Host ""

# If -HBAFilter was not passed as a parameter, prompt interactively
if ([string]::IsNullOrWhiteSpace($HBAFilter)) {
    $filterInput = Read-Host "  Filter to specific HBAs? (comma-separated, e.g. vmhba4,vmhba5) [Leave blank for ALL]"
    $HBAFilter = $filterInput.Trim()
}

# Build the final list of HBAs to check
if ([string]::IsNullOrWhiteSpace($HBAFilter)) {
    $HBASelection = $allFCHBAs
    Write-Host "  Checking ALL $($HBASelection.Count) FC HBA(s)." -ForegroundColor Green
} else {
    $HBASelection = $HBAFilter -split "," | ForEach-Object { $_.Trim() } | Where-Object { $_ -ne "" }
    # Warn about any requested HBAs not found in the environment
    $notFound = $HBASelection | Where-Object { $_ -notin $allFCHBAs }
    if ($notFound) {
        Write-Warning "  The following HBAs were not detected in this environment: $($notFound -join ', ')"
    }
    $HBASelection = $HBASelection | Where-Object { $_ -in $allFCHBAs }
    if (-not $HBASelection) {
        Write-Error "  No valid HBAs remain after filtering. Exiting."
        Disconnect-VIServer * -Confirm:$false
        exit 1
    }
    Write-Host "  Checking HBA(s): $($HBASelection -join ', ')" -ForegroundColor Green
}
Write-Host ""

$results = [System.Collections.Generic.List[PSObject]]::new()

foreach ($VMHost in $VMHosts) {

    if ($Rescan) {
        Write-Host "  [$($VMHost.Name)] Rescanning HBAs..." -ForegroundColor DarkGray
        Get-VMHostStorage -RescanAllHba -VMHost $VMHost | Out-Null
    }

    $HBAs = $VMHost | Get-VMHostHba -Type FibreChannel |
            Where-Object { $_.Device -in $HBASelection }

    if (-not $HBAs) {
        Write-Host "  [$($VMHost.Name)] No matching HBAs found — skipping." -ForegroundColor DarkYellow
        continue
    }

    foreach ($HBA in $HBAs) {
        try {
            $pathGroups = $HBA | Get-ScsiLun | Get-ScsiLunPath | Group-Object -Property State

            $active  = [int]($pathGroups | Where-Object Name -eq "Active"  | Select-Object -ExpandProperty Count -ErrorAction SilentlyContinue)
            $dead    = [int]($pathGroups | Where-Object Name -eq "Dead"    | Select-Object -ExpandProperty Count -ErrorAction SilentlyContinue)
            $standby = [int]($pathGroups | Where-Object Name -eq "Standby" | Select-Object -ExpandProperty Count -ErrorAction SilentlyContinue)

            $results.Add([PSCustomObject]@{
                VMHost  = $VMHost.Name
                HBA     = $HBA.Device
                Cluster = [string]$VMHost.Parent
                Active  = $active
                Dead    = $dead
                Standby = $standby
            })

        } catch {
            Write-Warning "  [$($VMHost.Name)][$($HBA.Device)] Error reading paths: $_"
        }
    }
}

# ─────────────────────────────────────────────
#  OUTPUT
# ─────────────────────────────────────────────
Write-Host ""

if ($results.Count -eq 0) {
    Write-Warning "No results collected."
} else {

    # ── Dynamic column widths (global, so all cluster tables align) ──
    $w_host    = [Math]::Max(6,  ($results | ForEach-Object { $_.VMHost.Length } | Measure-Object -Maximum).Maximum)
    $w_hba     = [Math]::Max(3,  ($results | ForEach-Object { $_.HBA.Length   } | Measure-Object -Maximum).Maximum)
    $w_active  = 6
    $w_dead    = 4
    $w_standby = 7

    # ── Border / header helpers (no Cluster column — grouped by cluster instead) ──
    $div = "  +" + ("-" * ($w_host    + 2)) + "+" +
                   ("-" * ($w_hba     + 2)) + "+" +
                   ("-" * ($w_active  + 2)) + "+" +
                   ("-" * ($w_dead    + 2)) + "+" +
                   ("-" * ($w_standby + 2)) + "+"

    $header = "  | " + "VMHost".PadRight($w_host)    + " | " +
                        "HBA".PadRight($w_hba)         + " | " +
                        "Active".PadRight($w_active)   + " | " +
                        "Dead".PadRight($w_dead)       + " | " +
                        "Standby".PadRight($w_standby) + " |"

    # ── Group results by cluster and print one table per cluster ──
    $clusters = $results | Select-Object -ExpandProperty Cluster -Unique | Sort-Object

    foreach ($cluster in $clusters) {
        $clusterRows = $results | Where-Object { $_.Cluster -eq $cluster } | Sort-Object VMHost, HBA

        # Cluster header banner
        $clusterLabel = "  Cluster: $cluster"
        Write-Host ""
        Write-Host $clusterLabel -ForegroundColor Magenta

        Write-Host $div    -ForegroundColor DarkGray
        Write-Host $header -ForegroundColor Cyan
        Write-Host $div    -ForegroundColor DarkGray

        foreach ($row in $clusterRows) {
            $line = "  | " + $row.VMHost.PadRight($w_host)               + " | " +
                              $row.HBA.PadRight($w_hba)                   + " | " +
                              ([string]$row.Active).PadRight($w_active)   + " | " +
                              ([string]$row.Dead).PadRight($w_dead)       + " | " +
                              ([string]$row.Standby).PadRight($w_standby) + " |"

            if ($row.Dead -gt 0) {
                Write-Host $line -ForegroundColor Red
            } elseif ($row.Standby -gt 0 -and $row.Active -eq 0) {
                Write-Host $line -ForegroundColor Yellow
            } elseif ($row.Active -gt 0) {
                Write-Host $line -ForegroundColor Green
            } else {
                Write-Host $line -ForegroundColor DarkYellow
            }
        }

        Write-Host $div -ForegroundColor DarkGray

        # Per-cluster summary
        $clusterDead = ($clusterRows | Where-Object { $_.Dead -gt 0 }).Count
        if ($clusterDead -gt 0) {
            Write-Host ("  [!] {0} HBA(s) with dead paths in this cluster." -f $clusterDead) -ForegroundColor Red
        } else {
            Write-Host "  [OK] All HBAs healthy in this cluster." -ForegroundColor Green
        }
    }

    Write-Host ""

    # ── Overall summary ────────────────────────
    $totalDead = ($results | Where-Object { $_.Dead -gt 0 }).Count
    if ($totalDead -gt 0) {
        Write-Host ("  [!!] TOTAL: {0} HBA(s) across all clusters have dead paths." -f $totalDead) -ForegroundColor Red
    } else {
        Write-Host "  [OK] All HBAs across all clusters reporting healthy paths." -ForegroundColor Green
    }
    Write-Host ""

    # ── Legend ─────────────────────────────────
    Write-Host "  Legend: " -NoNewline
    Write-Host "Green" -ForegroundColor Green   -NoNewline; Write-Host " = Active paths OK   " -NoNewline
    Write-Host "Yellow" -ForegroundColor Yellow -NoNewline; Write-Host " = Standby only   " -NoNewline
    Write-Host "Red" -ForegroundColor Red       -NoNewline; Write-Host " = Dead paths detected"
    Write-Host ""

    if ($ExportPath -ne "") {
        try {
            $results | Export-Csv -Path $ExportPath -NoTypeInformation -Encoding UTF8
            Write-Host "  Results exported to:   $ExportPath" -ForegroundColor Green
        } catch {
            Write-Warning "  CSV export failed: $_"
        }
        # Auto-generate HTML alongside the CSV
        $autoHTMLPath = [System.IO.Path]::ChangeExtension($ExportPath, ".html")
        New-FCHBAHTMLReport -Results $results -VCenter $vCenter -OutputPath $autoHTMLPath
    }

    if ($HTMLReport -ne "" -and $HTMLReport -ne [System.IO.Path]::ChangeExtension($ExportPath, ".html")) {
        New-FCHBAHTMLReport -Results $results -VCenter $vCenter -OutputPath $HTMLReport
    }
}

# ─────────────────────────────────────────────
#  DISCONNECT
# ─────────────────────────────────────────────
Disconnect-VIServer * -Confirm:$false
Write-Host ""
Write-Host "  Disconnected from $vCenter." -ForegroundColor DarkGray
Write-Host ""
