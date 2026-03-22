<#
.SYNOPSIS
    Get-FCHBAPathState - Reports Fibre Channel HBA path states across all ESXi hosts.

.DESCRIPTION
    Connects to vCenter and audits all Fibre Channel HBAs on every connected ESXi host.
    Auto-detects HBAs per host by default; optionally filter via -HBAFilter.
    Reports Active, Dead, and Standby path counts per HBA, grouped by cluster.
    Supports credential save/load, optional HBA rescan, CSV export, and HTML report.

.PARAMETER HBAFilter
    Optional. Comma-separated HBA device names to check (e.g. "vmhba4,vmhba5").
    If omitted, ALL Fibre Channel HBAs on each host are checked.

.PARAMETER ExportPath
    Optional. Full path for CSV output. HTML report auto-saved alongside it.
    Example: -ExportPath "C:\Reports\FCHBAPathState.csv"

.PARAMETER HTMLReport
    Optional. Full path for a standalone dark-mode HTML report.
    Example: -HTMLReport "C:\Reports\FCHBAPathState.html"

.PARAMETER Rescan
    Optional switch. Rescans all HBAs before collecting data.

.EXAMPLE
    .\Get-FCHBAPathState.ps1
    .\Get-FCHBAPathState.ps1 -HBAFilter "vmhba4,vmhba5"
    .\Get-FCHBAPathState.ps1 -ExportPath "C:\Reports\out.csv" -Rescan
    .\Get-FCHBAPathState.ps1 -HTMLReport "C:\Reports\out.html"

.NOTES
    Script  : Get-FCHBAPathState.ps1
    Version : 2.6
    Author  : Paul van Dieen
    Blog    : https://www.hollebollevsan.nl
    Date    : 2026-03-19

    Changelog:
    2.6  (2026-03-19) - PS 5.1 compatibility fixes (UTF-8 BOM, [char] escaping);
                         Read-Host credential prompts replace Get-Credential dialog;
                         HTML template converted from base64 to here-string;
                         $ScriptMeta block added as single source of truth for branding.
    2.5  (2026-03-19) - Added -HTMLReport parameter; dark-mode HTML report
                         auto-generated alongside CSV when -ExportPath is used.
    2.4  (2026-03-06) - Per-cluster output tables with summaries.
    2.3  (2026-03-06) - Colored console table, dynamic widths, legend.
    2.2  (2026-03-06) - Runtime PowerCLI check for 13.x compatibility.
    2.1  (2026-03-06) - Auto-detect FC HBAs; -HBAFilter parameter.
    2.0  (2026-03-06) - Full rewrite with credential store, CSV export.
    1.0  (initial)    - Original hardcoded script.
#>
param(
    [string]$HBAFilter  = "",
    [string]$ExportPath = "",
    [string]$HTMLReport = "",
    [switch]$Rescan
)

# --- Script metadata ---
$ScriptMeta = @{
    Name    = "Get-FCHBAPathState.ps1"
    Version = "2.6"
    Author  = "Paul van Dieen"
    Blog    = "https://www.hollebollevsan.nl"
    Date    = "2026-03-19"
}

# --- PowerCLI check ---
$powercli = Get-Module -ListAvailable -Name VMware.VimAutomation.Core |
            Sort-Object Version -Descending | Select-Object -First 1
if (-not $powercli) {
    Write-Error ("VMware PowerCLI not found. Install with" + [char]58 + " Install-Module VMware.PowerCLI")
    exit 1
}
Import-Module VMware.VimAutomation.Core -ErrorAction SilentlyContinue

# --- Credential helpers ---
$credStorePath = $env:USERPROFILE + "\.vcenter_creds"

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
    if (-not (Test-Path $credStorePath)) { return $null }
    try {
        $import = Import-Clixml -Path $credStorePath
        $sec = $import.Password | ConvertTo-SecureString
        return New-Object System.Management.Automation.PSCredential($import.Username, $sec)
    } catch {
        Write-Warning "Saved credentials could not be loaded. You will be prompted."
        return $null
    }
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

    $generatedAt   = (Get-Date).ToString("dd MMM yyyy  HH:mm:ss")
    $totalClusters = ($Results | Select-Object -ExpandProperty Cluster -Unique).Count
    $totalHosts    = ($Results | Select-Object -ExpandProperty VMHost  -Unique).Count
    $totalHBAs     = $Results.Count
    $totalHealthy  = ($Results | Where-Object { $_.Dead -eq 0 }).Count
    $totalDead     = ($Results | Where-Object { $_.Dead -gt 0 }).Count

    $template = @"
<!DOCTYPE html>
<html lang="en">
<head>
<meta charset="UTF-8">
<meta name="viewport" content="width=device-width, initial-scale=1.0">
<title>Get-FCHBAPathState - FC HBA Path Report</title>
<style>
@import url('https://fonts.googleapis.com/css2?family=JetBrains+Mono:wght@400;600;700&family=Syne:wght@400;600;700;800&display=swap');
:root{--bg:#0d0f14;--bg2:#12151c;--bg3:#181c26;--border:#252a38;--border2:#2e3448;--text:#c8cfe0;--text-dim:#5a6480;--text-head:#e8edf8;--accent:#3d9cf0;--accent2:#5ab4ff;--green:#3dd68c;--green-bg:rgba(61,214,140,.08);--green-bdr:rgba(61,214,140,.25);--yellow:#f0c040;--yellow-bg:rgba(240,192,64,.08);--yellow-bdr:rgba(240,192,64,.25);--red:#f05060;--red-bg:rgba(240,80,96,.08);--red-bdr:rgba(240,80,96,.30);--magenta:#c080f0;--mono:'JetBrains Mono',monospace;--sans:'Syne',sans-serif}
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
<body>%%HEADER%%%%SUMMARY%%%%STATUSBANNER%%<div class="main">%%CLUSTERS%%</div>%%LEGEND%%<footer class="page-footer"><span>$($script:ScriptMeta.Name) v$($script:ScriptMeta.Version) &middot; <a href="$($script:ScriptMeta.Blog)" target="_blank">hollebollevsan.nl</a> &middot; $($script:ScriptMeta.Author)</span><span>%%GENERATED%%</span></footer>
</body>
</html>
"@

    if ($totalDead -gt 0) {
        $statusClass = "crit"
        $statusMsg   = "&#9888; $totalDead HBA(s) have dead paths &mdash; investigate immediately."
    } else {
        $statusClass = "ok"
        $statusMsg   = "&#10003; All HBAs across all clusters reporting healthy paths."
    }

    # Build per-cluster HTML
    $clusterHTML = ""
    $clusters = $Results | Select-Object -ExpandProperty Cluster -Unique | Sort-Object

    foreach ($cluster in $clusters) {
        $rows = $Results | Where-Object { $_.Cluster -eq $cluster } | Sort-Object VMHost, HBA
        $clusterDead = ($rows | Where-Object { $_.Dead -gt 0 }).Count
        if ($clusterDead -gt 0) { $cpillClass = "crit"; $cpillText = "$clusterDead dead path(s)" }
        else                     { $cpillClass = "ok";   $cpillText = "All healthy" }

        $rowHTML = ""
        foreach ($row in $rows) {
            if    ($row.Dead    -gt 0)                         { $rc = "row-dead";   $dc = "dead";    $sc = "chip dead";    $sl = "&#10005; Dead paths" }
            elseif($row.Standby -gt 0 -and $row.Active -eq 0) { $rc = "row-warn";   $dc = "warn";    $sc = "chip standby"; $sl = "&#9873; Standby only" }
            elseif($row.Active  -gt 0)                         { $rc = "row-ok";     $dc = "ok";      $sc = "chip active";  $sl = "&#10003; Healthy" }
            else                                               { $rc = "row-nopath"; $dc = "warn";    $sc = "chip standby"; $sl = "No paths" }

            if ($row.Active  -gt 0) { $ac = "<span class=`"chip active`">"  + $row.Active  + "</span>" } else { $ac = "<span class=`"chip zero`">0</span>" }
            if ($row.Dead    -gt 0) { $dc2 = "<span class=`"chip dead`">"   + $row.Dead    + "</span>" } else { $dc2 = "<span class=`"chip zero`">0</span>" }
            if ($row.Standby -gt 0) { $stc = "<span class=`"chip standby`">" + $row.Standby + "</span>" } else { $stc = "<span class=`"chip zero`">0</span>" }

            $rowHTML += "<tr class=`"$rc`"><td><div class=`"host-cell`"><div class=`"host-dot $dc`"></div>" + $row.VMHost + "</div></td>"
            $rowHTML += "<td><span class=`"hba-device`">" + $row.HBA + "</span></td>"
            $rowHTML += "<td class=`"num`">$ac</td><td class=`"num`">$dc2</td><td class=`"num`">$stc</td>"
            $rowHTML += "<td><span class=`"$sc`">$sl</span></td></tr>"
        }

        $clusterHTML += "<div class=`"cluster-block`"><div class=`"cluster-heading`">"
        $clusterHTML += "<span class=`"cluster-name`">&#9672; $cluster</span>"
        $clusterHTML += "<span class=`"cluster-pill $cpillClass`">$cpillText</span></div>"
        $clusterHTML += "<div class=`"table-wrap`"><table class=`"hba-table`">"
        $clusterHTML += "<thead><tr><th>VMHost</th><th>HBA</th><th class=`"num`">Active</th><th class=`"num`">Dead</th><th class=`"num`">Standby</th><th>Status</th></tr></thead>"
        $clusterHTML += "<tbody>$rowHTML</tbody></table></div></div>"
    }

    $headerHTML  = "<header class=`"page-header`"><div class=`"header-left`">"
    $headerHTML += "<div class=`"header-tool`">Get-FC<span>HBAPathState</span></div>"
    $headerHTML += "<div class=`"header-subtitle`">FC HBA Path State Report &middot; $VCenter</div></div>"
    $headerHTML += "<div class=`"header-meta`"><span class=`"badge v`">v$($script:ScriptMeta.Version)</span>"
    $headerHTML += "<span class=`"badge`">Generated: $generatedAt</span></div></header>"

    $summaryHTML  = "<div class=`"summary-bar`">"
    $summaryHTML += "<div class=`"stat-card`"><span class=`"stat-label`">Clusters</span><span class=`"stat-value`">$totalClusters</span></div>"
    $summaryHTML += "<div class=`"stat-card`"><span class=`"stat-label`">Hosts</span><span class=`"stat-value`">$totalHosts</span></div>"
    $summaryHTML += "<div class=`"stat-card`"><span class=`"stat-label`">HBAs checked</span><span class=`"stat-value`">$totalHBAs</span></div>"
    $summaryHTML += "<div class=`"stat-card ok`"><span class=`"stat-label`">Healthy</span><span class=`"stat-value`">$totalHealthy</span></div>"
    $summaryHTML += "<div class=`"stat-card crit`"><span class=`"stat-label`">Dead paths</span><span class=`"stat-value`">$totalDead</span></div></div>"

    $bannerHTML = "<div class=`"status-banner $statusClass`">$statusMsg</div>"

    $legendHTML  = "<div class=`"legend`"><span class=`"legend-title`">Legend</span>"
    $legendHTML += "<span class=`"legend-item`"><span class=`"chip active`">Active</span> Paths routing I/O normally</span>"
    $legendHTML += "<span class=`"legend-item`"><span class=`"chip dead`">Dead</span> Paths lost - immediate action required</span>"
    $legendHTML += "<span class=`"legend-item`"><span class=`"chip standby`">Standby</span> Paths available but not routing I/O</span>"
    $legendHTML += "<span class=`"legend-item`"><span class=`"chip zero`">0</span> No paths in this state</span></div>"

    $html = $template -replace "%%HEADER%%",       $headerHTML
    $html = $html     -replace "%%SUMMARY%%",      $summaryHTML
    $html = $html     -replace "%%STATUSBANNER%%", $bannerHTML
    $html = $html     -replace "%%CLUSTERS%%",     $clusterHTML
    $html = $html     -replace "%%LEGEND%%",       $legendHTML
    $html = $html     -replace "%%GENERATED%%",    $generatedAt

    try {
        [System.IO.File]::WriteAllText($OutputPath, $html, [System.Text.UTF8Encoding]::new($false))
        Write-Host "  HTML report saved to:  $OutputPath" -ForegroundColor Green
    } catch {
        Write-Warning "  HTML report export failed: $_"
    }
}


# --- Banner ---
Write-Host ""
Write-Host ("=" * 52) -ForegroundColor Cyan
Write-Host ("  " + $ScriptMeta.Name + "  v" + $ScriptMeta.Version) -ForegroundColor Cyan
Write-Host ("  " + $ScriptMeta.Author) -ForegroundColor Cyan
Write-Host ("  " + $ScriptMeta.Blog) -ForegroundColor Cyan
Write-Host ("  " + $ScriptMeta.Date) -ForegroundColor Cyan
Write-Host ("=" * 52) -ForegroundColor Cyan
Write-Host ""

# --- vCenter prompt ---
$vCenter = Read-Host "  Enter vCenter FQDN or IP"
if ([string]::IsNullOrWhiteSpace($vCenter)) {
    Write-Error "No vCenter specified. Exiting."
    exit 1
}

# --- Credentials ---
$savedCred = Load-VCenterCredential
if ($savedCred) {
    Write-Host ""
    Write-Host ("  Found saved credentials for" + [char]58 + " ") -NoNewline
    Write-Host $savedCred.UserName -ForegroundColor Yellow
    $useSaved = Read-Host "  Use saved credentials? Y/N [Y]"
    if ($useSaved -eq "" -or $useSaved -match "^[Yy]") {
        $cred = $savedCred
        Write-Host "  Using saved credentials." -ForegroundColor Green
    } else {
        $credUser = Read-Host "  Username"
        $credPass = Read-Host "  Password" -AsSecureString
        $cred = New-Object System.Management.Automation.PSCredential($credUser, $credPass)
        $saveNew = Read-Host "  Save these credentials? Y/N [N]"
        if ($saveNew -match "^[Yy]") { Save-VCenterCredential -Credential $cred }
    }
} else {
    Write-Host ""
    $credUser = Read-Host "  Username"
    $credPass = Read-Host "  Password" -AsSecureString
    $cred = New-Object System.Management.Automation.PSCredential($credUser, $credPass)
    $saveNew = Read-Host "  Save these credentials? Y/N [N]"
    if ($saveNew -match "^[Yy]") { Save-VCenterCredential -Credential $cred }
}

# --- Connect ---
Write-Host ""
Write-Host "  Connecting to $vCenter ..." -ForegroundColor Cyan
try {
    Connect-VIServer -Server $vCenter -Credential $cred -ErrorAction Stop | Out-Null
    Write-Host "  Connected successfully." -ForegroundColor Green
} catch {
    Write-Error ("Failed to connect to " + $vCenter + " " + [char]58 + " " + $_)
    exit 1
}

# --- Get hosts ---
Write-Host ""
Write-Host "  Gathering ESXi hosts..." -ForegroundColor Cyan
$VMHosts = Get-VMHost | Where-Object { $_.ConnectionState -eq "Connected" } | Sort-Object Name
if (-not $VMHosts) {
    Write-Warning "No connected hosts found."
    Disconnect-VIServer * -Confirm:$false
    exit 0
}
Write-Host ("  Found " + $VMHosts.Count + " connected hosts.") -ForegroundColor Cyan
Write-Host ""

# --- HBA detection ---
$allFCHBAs = $VMHosts | Get-VMHostHba -Type FibreChannel |
             Select-Object -ExpandProperty Device -Unique | Sort-Object
if (-not $allFCHBAs) {
    Write-Warning "No Fibre Channel HBAs found on any host."
    Disconnect-VIServer * -Confirm:$false
    exit 0
}

Write-Host ("  FC HBAs detected" + [char]58) -ForegroundColor Cyan
foreach ($h in $allFCHBAs) {
    Write-Host "    - $h" -ForegroundColor White
}
Write-Host ""

# --- HBA filter ---
if ([string]::IsNullOrWhiteSpace($HBAFilter)) {
    $filterInput = Read-Host "  Filter HBAs? Comma-separated, leave blank for ALL"
    $HBAFilter = $filterInput.Trim()
}

if ([string]::IsNullOrWhiteSpace($HBAFilter)) {
    $HBASelection = $allFCHBAs
    Write-Host ("  Checking ALL " + $HBASelection.Count + " FC HBAs.") -ForegroundColor Green
} else {
    $requested = $HBAFilter -split "," | ForEach-Object { $_.Trim() } | Where-Object { $_ -ne "" }
    $notFound = $requested | Where-Object { $_ -notin $allFCHBAs }
    if ($notFound) {
        Write-Warning ("HBAs not found in environment" + [char]58 + " " + ($notFound -join ", "))
    }
    $HBASelection = $requested | Where-Object { $_ -in $allFCHBAs }
    if (-not $HBASelection) {
        Write-Error "No valid HBAs after filtering. Exiting."
        Disconnect-VIServer * -Confirm:$false
        exit 1
    }
    Write-Host ("  Checking" + [char]58 + " " + ($HBASelection -join ", ")) -ForegroundColor Green
}
Write-Host ""

# --- Collect data ---
$results = [System.Collections.Generic.List[PSObject]]::new()

foreach ($VMHost in $VMHosts) {
    if ($Rescan) {
        Write-Host "  [$($VMHost.Name)] Rescanning..." -ForegroundColor DarkGray
        Get-VMHostStorage -RescanAllHba -VMHost $VMHost | Out-Null
    }

    $HBAs = $VMHost | Get-VMHostHba -Type FibreChannel |
            Where-Object { $_.Device -in $HBASelection }

    if (-not $HBAs) {
        Write-Host "  [$($VMHost.Name)] No matching HBAs - skipping." -ForegroundColor DarkYellow
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
            Write-Warning ("[" + $VMHost.Name + "][" + $HBA.Device + "] Error" + [char]58 + " " + $_)
        }
    }
}

# --- Output ---
Write-Host ""
if ($results.Count -eq 0) {
    Write-Warning "No results collected."
} else {
    $w_host    = [Math]::Max(6,  ($results | ForEach-Object { $_.VMHost.Length } | Measure-Object -Maximum).Maximum)
    $w_hba     = [Math]::Max(3,  ($results | ForEach-Object { $_.HBA.Length   } | Measure-Object -Maximum).Maximum)
    $w_active  = 6
    $w_dead    = 4
    $w_standby = 7

    $seg1 = [string]::new([char]'-', ($w_host    + 2))
    $seg2 = [string]::new([char]'-', ($w_hba     + 2))
    $seg3 = [string]::new([char]'-', ($w_active  + 2))
    $seg4 = [string]::new([char]'-', ($w_dead    + 2))
    $seg5 = [string]::new([char]'-', ($w_standby + 2))
    $div  = "  +" + $seg1 + "+" + $seg2 + "+" + $seg3 + "+" + $seg4 + "+" + $seg5 + "+"

    $h0 = "VMHost".PadRight($w_host)
    $h1 = "HBA".PadRight($w_hba)
    $h2 = "Active".PadRight($w_active)
    $h3 = "Dead".PadRight($w_dead)
    $h4 = "Standby".PadRight($w_standby)
    $header = "  " + [char]124 + " " + $h0 + " " + [char]124 + " " + $h1 + " " + [char]124 + " " + $h2 + " " + [char]124 + " " + $h3 + " " + [char]124 + " " + $h4 + " " + [char]124

    $clusters = $results | Select-Object -ExpandProperty Cluster -Unique | Sort-Object

    foreach ($cluster in $clusters) {
        $clusterRows = $results | Where-Object { $_.Cluster -eq $cluster } | Sort-Object VMHost, HBA
        Write-Host ""
        Write-Host ("  Cluster" + [char]58 + " " + $cluster) -ForegroundColor Magenta
        Write-Host $div    -ForegroundColor DarkGray
        Write-Host $header -ForegroundColor Cyan
        Write-Host $div    -ForegroundColor DarkGray

        foreach ($row in $clusterRows) {
            $c0   = $row.VMHost.PadRight($w_host)
            $c1   = $row.HBA.PadRight($w_hba)
            $c2   = ([string]$row.Active).PadRight($w_active)
            $c3   = ([string]$row.Dead).PadRight($w_dead)
            $c4   = ([string]$row.Standby).PadRight($w_standby)
            $pipe = [char]124
            $line = "  " + $pipe + " " + $c0 + " " + $pipe + " " + $c1 + " " + $pipe + " " + $c2 + " " + $pipe + " " + $c3 + " " + $pipe + " " + $c4 + " " + $pipe
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
        $clusterDead = ($clusterRows | Where-Object { $_.Dead -gt 0 }).Count
        if ($clusterDead -gt 0) {
            Write-Host ("  [!] " + $clusterDead + " HBAs with dead paths in this cluster.") -ForegroundColor Red
        } else {
            Write-Host "  [OK] All HBAs healthy in this cluster." -ForegroundColor Green
        }
    }

    Write-Host ""
    $totalDead = ($results | Where-Object { $_.Dead -gt 0 }).Count
    if ($totalDead -gt 0) {
        Write-Host ("  [!!] TOTAL" + [char]58 + " " + $totalDead + " HBAs across all clusters have dead paths.") -ForegroundColor Red
    } else {
        Write-Host "  [OK] All HBAs across all clusters reporting healthy paths." -ForegroundColor Green
    }
    Write-Host ""

    Write-Host "  Legend" -NoNewline
    Write-Host [char]58                             -NoNewline
    Write-Host " Green"  -ForegroundColor Green    -NoNewline
    Write-Host " = Active   "                      -NoNewline
    Write-Host " Yellow" -ForegroundColor Yellow   -NoNewline
    Write-Host " = Standby only   "                -NoNewline
    Write-Host " Red"    -ForegroundColor Red      -NoNewline
    Write-Host " = Dead paths"
    Write-Host ""

    if ($ExportPath -ne "") {
        try {
            $results | Export-Csv -Path $ExportPath -NoTypeInformation -Encoding UTF8
            Write-Host ("  CSV saved to" + [char]58 + " " + $ExportPath) -ForegroundColor Green
        } catch {
            Write-Warning ("CSV export failed" + [char]58 + " " + $_)
        }
        $autoHTML = [System.IO.Path]::ChangeExtension($ExportPath, ".html")
        New-FCHBAHTMLReport -Results $results -VCenter $vCenter -OutputPath $autoHTML
    }

    if ($HTMLReport -ne "") {
        $autoHTML2 = ""
        if ($ExportPath -ne "") {
            $autoHTML2 = [System.IO.Path]::ChangeExtension($ExportPath, ".html")
        }
        if ($HTMLReport -ne $autoHTML2) {
            New-FCHBAHTMLReport -Results $results -VCenter $vCenter -OutputPath $HTMLReport
        }
    }
}
# --- Disconnect ---
Disconnect-VIServer * -Confirm:$false
Write-Host ""
Write-Host ("  Disconnected from " + $vCenter + ".") -ForegroundColor DarkGray
Write-Host ("  " + $ScriptMeta.Blog) -ForegroundColor DarkGray
Write-Host ""
