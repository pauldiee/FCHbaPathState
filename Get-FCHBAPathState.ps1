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
    Author  : Paul van Dieen
    Blog    : https://www.hollebollevsan.nl
    Version : 2.5  (2026-03-19) - Added -HTMLReport; HTML auto-generated with -ExportPath.
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

    # Decode HTML template from base64 (avoids PS 5.1 parser issues with CSS braces)
    $templateB64 = "PCFET0NUWVBFIGh0bWw+CjxodG1sIGxhbmc9ImVuIj4KPGhlYWQ+CjxtZXRhIGNoYXJzZXQ9IlVURi04Ij4KPG1ldGEgbmFtZT0idmlld3BvcnQiIGNvbnRlbnQ9IndpZHRoPWRldmljZS13aWR0aCwgaW5pdGlhbC1zY2FsZT0xLjAiPgo8dGl0bGU+R2V0LUZDSEJBUGF0aFN0YXRlIC0gRkMgSEJBIFBhdGggUmVwb3J0PC90aXRsZT4KPHN0eWxlPgpAaW1wb3J0IHVybCgnaHR0cHM6Ly9mb250cy5nb29nbGVhcGlzLmNvbS9jc3MyP2ZhbWlseT1KZXRCcmFpbnMrTW9ubzp3Z2h0QDQwMDs2MDA7NzAwJmZhbWlseT1TeW5lOndnaHRANDAwOzYwMDs3MDA7ODAwJmRpc3BsYXk9c3dhcCcpOwo6cm9vdHstLWJnOiMwZDBmMTQ7LS1iZzI6IzEyMTUxYzstLWJnMzojMTgxYzI2Oy0tYm9yZGVyOiMyNTJhMzg7LS1ib3JkZXIyOiMyZTM0NDg7LS10ZXh0OiNjOGNmZTA7LS10ZXh0LWRpbTojNWE2NDgwOy0tdGV4dC1oZWFkOiNlOGVkZjg7LS1hY2NlbnQ6IzNkOWNmMDstLWFjY2VudDI6IzVhYjRmZjstLWdyZWVuOiMzZGQ2OGM7LS1ncmVlbi1iZzpyZ2JhKDYxLDIxNCwxNDAsLjA4KTstLWdyZWVuLWJkcjpyZ2JhKDYxLDIxNCwxNDAsLjI1KTstLXllbGxvdzojZjBjMDQwOy0teWVsbG93LWJnOnJnYmEoMjQwLDE5Miw2NCwuMDgpOy0teWVsbG93LWJkcjpyZ2JhKDI0MCwxOTIsNjQsLjI1KTstLXJlZDojZjA1MDYwOy0tcmVkLWJnOnJnYmEoMjQwLDgwLDk2LC4wOCk7LS1yZWQtYmRyOnJnYmEoMjQwLDgwLDk2LC4zMCk7LS1tYWdlbnRhOiNjMDgwZjA7LS1tb25vOidKZXRCcmFpbnMgTW9ubycsbW9ub3NwYWNlOy0tc2FuczonU3luZScsc2Fucy1zZXJpZn0KKiwqOjpiZWZvcmUsKjo6YWZ0ZXJ7Ym94LXNpemluZzpib3JkZXItYm94O21hcmdpbjowO3BhZGRpbmc6MH0KYm9keXtiYWNrZ3JvdW5kOnZhcigtLWJnKTtjb2xvcjp2YXIoLS10ZXh0KTtmb250LWZhbWlseTp2YXIoLS1tb25vKTtmb250LXNpemU6MTNweDtsaW5lLWhlaWdodDoxLjY7bWluLWhlaWdodDoxMDB2aDtwYWRkaW5nOjAgMCA2MHB4fQoucGFnZS1oZWFkZXJ7YmFja2dyb3VuZDp2YXIoLS1iZzIpO2JvcmRlci1ib3R0b206MXB4IHNvbGlkIHZhcigtLWJvcmRlcjIpO3BhZGRpbmc6MjhweCA0MHB4IDI0cHg7ZGlzcGxheTpmbGV4O2FsaWduLWl0ZW1zOmZsZXgtc3RhcnQ7anVzdGlmeS1jb250ZW50OnNwYWNlLWJldHdlZW47Z2FwOjI0cHg7cG9zaXRpb246cmVsYXRpdmU7b3ZlcmZsb3c6aGlkZGVufQoucGFnZS1oZWFkZXI6OmJlZm9yZXtjb250ZW50OicnO3Bvc2l0aW9uOmFic29sdXRlO3RvcDowO2xlZnQ6MDtyaWdodDowO2hlaWdodDoycHg7YmFja2dyb3VuZDpsaW5lYXItZ3JhZGllbnQoOTBkZWcsdmFyKC0tYWNjZW50KSAwJSx2YXIoLS1tYWdlbnRhKSA1MCUsdmFyKC0tZ3JlZW4pIDEwMCUpfQouaGVhZGVyLWxlZnR7ZGlzcGxheTpmbGV4O2ZsZXgtZGlyZWN0aW9uOmNvbHVtbjtnYXA6NnB4fQouaGVhZGVyLXRvb2x7Zm9udC1mYW1pbHk6dmFyKC0tc2Fucyk7Zm9udC1zaXplOjIycHg7Zm9udC13ZWlnaHQ6ODAwO2NvbG9yOnZhcigtLXRleHQtaGVhZCk7bGV0dGVyLXNwYWNpbmc6LTAuNXB4fQouaGVhZGVyLXRvb2wgc3Bhbntjb2xvcjp2YXIoLS1hY2NlbnQyKX0KLmhlYWRlci1zdWJ0aXRsZXtjb2xvcjp2YXIoLS10ZXh0LWRpbSk7Zm9udC1zaXplOjExcHg7bGV0dGVyLXNwYWNpbmc6MC41cHh9Ci5oZWFkZXItbWV0YXtkaXNwbGF5OmZsZXg7ZmxleC1kaXJlY3Rpb246Y29sdW1uO2FsaWduLWl0ZW1zOmZsZXgtZW5kO2dhcDo0cHg7Zm9udC1zaXplOjExcHg7Y29sb3I6dmFyKC0tdGV4dC1kaW0pfQouYmFkZ2V7ZGlzcGxheTppbmxpbmUtZmxleDthbGlnbi1pdGVtczpjZW50ZXI7Z2FwOjVweDtiYWNrZ3JvdW5kOnZhcigtLWJnMyk7Ym9yZGVyOjFweCBzb2xpZCB2YXIoLS1ib3JkZXIyKTtib3JkZXItcmFkaXVzOjRweDtwYWRkaW5nOjNweCA4cHg7Zm9udC1zaXplOjExcHg7Y29sb3I6dmFyKC0tdGV4dC1kaW0pfQouYmFkZ2Uudntjb2xvcjp2YXIoLS1hY2NlbnQpO2JvcmRlci1jb2xvcjpyZ2JhKDYxLDE1NiwyNDAsLjMpO2JhY2tncm91bmQ6cmdiYSg2MSwxNTYsMjQwLC4wNil9Ci5zdW1tYXJ5LWJhcntkaXNwbGF5OmZsZXg7Z2FwOjEycHg7cGFkZGluZzoxNnB4IDQwcHg7YmFja2dyb3VuZDp2YXIoLS1iZzIpO2JvcmRlci1ib3R0b206MXB4IHNvbGlkIHZhcigtLWJvcmRlcik7ZmxleC13cmFwOndyYXB9Ci5zdGF0LWNhcmR7ZGlzcGxheTpmbGV4O2ZsZXgtZGlyZWN0aW9uOmNvbHVtbjtnYXA6MnB4O2JhY2tncm91bmQ6dmFyKC0tYmczKTtib3JkZXI6MXB4IHNvbGlkIHZhcigtLWJvcmRlcik7Ym9yZGVyLXJhZGl1czo2cHg7cGFkZGluZzoxMHB4IDE4cHg7bWluLXdpZHRoOjExMHB4fQouc3RhdC1sYWJlbHtmb250LXNpemU6MTBweDtjb2xvcjp2YXIoLS10ZXh0LWRpbSk7bGV0dGVyLXNwYWNpbmc6MC44cHg7dGV4dC10cmFuc2Zvcm06dXBwZXJjYXNlfQouc3RhdC12YWx1ZXtmb250LXNpemU6MjJweDtmb250LXdlaWdodDo3MDA7Zm9udC1mYW1pbHk6dmFyKC0tc2Fucyk7Y29sb3I6dmFyKC0tdGV4dC1oZWFkKX0KLnN0YXQtY2FyZC5vayAuc3RhdC12YWx1ZXtjb2xvcjp2YXIoLS1ncmVlbil9LnN0YXQtY2FyZC5jcml0IC5zdGF0LXZhbHVle2NvbG9yOnZhcigtLXJlZCl9Ci5zdGF0dXMtYmFubmVye21hcmdpbjoyMHB4IDQwcHggMDtib3JkZXItcmFkaXVzOjZweDtwYWRkaW5nOjEwcHggMTZweDtmb250LXNpemU6MTJweDtmb250LXdlaWdodDo2MDA7ZGlzcGxheTpmbGV4O2FsaWduLWl0ZW1zOmNlbnRlcjtnYXA6OHB4fQouc3RhdHVzLWJhbm5lci5va3tiYWNrZ3JvdW5kOnZhcigtLWdyZWVuLWJnKTtib3JkZXI6MXB4IHNvbGlkIHZhcigtLWdyZWVuLWJkcik7Y29sb3I6dmFyKC0tZ3JlZW4pfQouc3RhdHVzLWJhbm5lci5jcml0e2JhY2tncm91bmQ6dmFyKC0tcmVkLWJnKTtib3JkZXI6MXB4IHNvbGlkIHZhcigtLXJlZC1iZHIpO2NvbG9yOnZhcigtLXJlZCl9Ci5tYWlue3BhZGRpbmc6MjBweCA0MHB4IDB9Ci5jbHVzdGVyLWJsb2Nre21hcmdpbi1ib3R0b206MzJweH0KLmNsdXN0ZXItaGVhZGluZ3tkaXNwbGF5OmZsZXg7YWxpZ24taXRlbXM6Y2VudGVyO2dhcDoxMHB4O21hcmdpbi1ib3R0b206MTBweH0KLmNsdXN0ZXItbmFtZXtmb250LWZhbWlseTp2YXIoLS1zYW5zKTtmb250LXNpemU6MTRweDtmb250LXdlaWdodDo3MDA7Y29sb3I6dmFyKC0tbWFnZW50YSk7bGV0dGVyLXNwYWNpbmc6MC4zcHh9Ci5jbHVzdGVyLXBpbGx7Zm9udC1zaXplOjEwcHg7YmFja2dyb3VuZDpyZ2JhKDE5MiwxMjgsMjQwLC4xKTtib3JkZXI6MXB4IHNvbGlkIHJnYmEoMTkyLDEyOCwyNDAsLjI1KTtjb2xvcjp2YXIoLS1tYWdlbnRhKTtib3JkZXItcmFkaXVzOjIwcHg7cGFkZGluZzoycHggOHB4fQouY2x1c3Rlci1waWxsLm9re2JhY2tncm91bmQ6dmFyKC0tZ3JlZW4tYmcpO2JvcmRlci1jb2xvcjp2YXIoLS1ncmVlbi1iZHIpO2NvbG9yOnZhcigtLWdyZWVuKX0KLmNsdXN0ZXItcGlsbC5jcml0e2JhY2tncm91bmQ6dmFyKC0tcmVkLWJnKTtib3JkZXItY29sb3I6dmFyKC0tcmVkLWJkcik7Y29sb3I6dmFyKC0tcmVkKX0KLmhiYS10YWJsZXt3aWR0aDoxMDAlO2JvcmRlci1jb2xsYXBzZTpjb2xsYXBzZTtib3JkZXI6MXB4IHNvbGlkIHZhcigtLWJvcmRlcjIpO2JvcmRlci1yYWRpdXM6NnB4O292ZXJmbG93OmhpZGRlbn0KLmhiYS10YWJsZSB0aGVhZCB0cntiYWNrZ3JvdW5kOnZhcigtLWJnMyk7Ym9yZGVyLWJvdHRvbToxcHggc29saWQgdmFyKC0tYm9yZGVyMil9Ci5oYmEtdGFibGUgdGh7cGFkZGluZzo5cHggMTRweDt0ZXh0LWFsaWduOmxlZnQ7Zm9udC1zaXplOjEwcHg7Zm9udC13ZWlnaHQ6NjAwO2xldHRlci1zcGFjaW5nOjAuOHB4O3RleHQtdHJhbnNmb3JtOnVwcGVyY2FzZTtjb2xvcjp2YXIoLS1hY2NlbnQpfQouaGJhLXRhYmxlIHRoLm51bXt0ZXh0LWFsaWduOnJpZ2h0fQouaGJhLXRhYmxlIHRib2R5IHRye2JvcmRlci1ib3R0b206MXB4IHNvbGlkIHZhcigtLWJvcmRlcik7dHJhbnNpdGlvbjpiYWNrZ3JvdW5kIDAuMTVzfQouaGJhLXRhYmxlIHRib2R5IHRyOmxhc3QtY2hpbGR7Ym9yZGVyLWJvdHRvbTpub25lfQouaGJhLXRhYmxlIHRib2R5IHRyOmhvdmVye2JhY2tncm91bmQ6cmdiYSgyNTUsMjU1LDI1NSwuMDI1KX0KLmhiYS10YWJsZSB0ZHtwYWRkaW5nOjlweCAxNHB4O2ZvbnQtc2l6ZToxMnB4O2NvbG9yOnZhcigtLXRleHQpfQouaGJhLXRhYmxlIHRkLm51bXt0ZXh0LWFsaWduOnJpZ2h0O2ZvbnQtd2VpZ2h0OjYwMH0KLnJvdy1va3tiYWNrZ3JvdW5kOnJnYmEoNjEsMjE0LDE0MCwuMDMpfS5yb3ctd2FybntiYWNrZ3JvdW5kOnJnYmEoMjQwLDE5Miw2NCwuMDQpfQoucm93LWRlYWR7YmFja2dyb3VuZDpyZ2JhKDI0MCw4MCw5NiwuMDUpfS5yb3ctbm9wYXRoe2JhY2tncm91bmQ6cmdiYSgyNDAsMTYwLDY0LC4wNCl9Ci5jaGlwe2Rpc3BsYXk6aW5saW5lLWZsZXg7YWxpZ24taXRlbXM6Y2VudGVyO2dhcDo0cHg7Ym9yZGVyLXJhZGl1czo0cHg7cGFkZGluZzoycHggN3B4O2ZvbnQtc2l6ZToxMXB4O2ZvbnQtd2VpZ2h0OjYwMH0KLmNoaXAuYWN0aXZle2JhY2tncm91bmQ6dmFyKC0tZ3JlZW4tYmcpO2JvcmRlcjoxcHggc29saWQgdmFyKC0tZ3JlZW4tYmRyKTtjb2xvcjp2YXIoLS1ncmVlbil9Ci5jaGlwLmRlYWR7YmFja2dyb3VuZDp2YXIoLS1yZWQtYmcpO2JvcmRlcjoxcHggc29saWQgdmFyKC0tcmVkLWJkcik7Y29sb3I6dmFyKC0tcmVkKX0KLmNoaXAuc3RhbmRieXtiYWNrZ3JvdW5kOnZhcigtLXllbGxvdy1iZyk7Ym9yZGVyOjFweCBzb2xpZCB2YXIoLS15ZWxsb3ctYmRyKTtjb2xvcjp2YXIoLS15ZWxsb3cpfQouY2hpcC56ZXJve2JhY2tncm91bmQ6dmFyKC0tYmczKTtib3JkZXI6MXB4IHNvbGlkIHZhcigtLWJvcmRlcik7Y29sb3I6dmFyKC0tdGV4dC1kaW0pfQouaG9zdC1jZWxse2Rpc3BsYXk6ZmxleDthbGlnbi1pdGVtczpjZW50ZXI7Z2FwOjhweH0KLmhvc3QtZG90e3dpZHRoOjZweDtoZWlnaHQ6NnB4O2JvcmRlci1yYWRpdXM6NTAlO2ZsZXgtc2hyaW5rOjB9Ci5ob3N0LWRvdC5va3tiYWNrZ3JvdW5kOnZhcigtLWdyZWVuKTtib3gtc2hhZG93OjAgMCA2cHggdmFyKC0tZ3JlZW4pfQouaG9zdC1kb3Qud2FybntiYWNrZ3JvdW5kOnZhcigtLXllbGxvdyk7Ym94LXNoYWRvdzowIDAgNnB4IHZhcigtLXllbGxvdyl9Ci5ob3N0LWRvdC5kZWFke2JhY2tncm91bmQ6dmFyKC0tcmVkKTtib3gtc2hhZG93OjAgMCA2cHggdmFyKC0tcmVkKX0KLmhiYS1kZXZpY2V7Zm9udC1mYW1pbHk6dmFyKC0tbW9ubyk7Zm9udC1zaXplOjExcHg7Y29sb3I6dmFyKC0tYWNjZW50Mik7YmFja2dyb3VuZDpyZ2JhKDYxLDE1NiwyNDAsLjA4KTtib3JkZXI6MXB4IHNvbGlkIHJnYmEoNjEsMTU2LDI0MCwuMik7Ym9yZGVyLXJhZGl1czo0cHg7cGFkZGluZzoxcHggNnB4fQoubGVnZW5ke2Rpc3BsYXk6ZmxleDtnYXA6MTZweDtmbGV4LXdyYXA6d3JhcDttYXJnaW46MjRweCA0MHB4IDA7cGFkZGluZzoxMnB4IDE2cHg7YmFja2dyb3VuZDp2YXIoLS1iZzIpO2JvcmRlcjoxcHggc29saWQgdmFyKC0tYm9yZGVyKTtib3JkZXItcmFkaXVzOjZweDtmb250LXNpemU6MTFweDtjb2xvcjp2YXIoLS10ZXh0LWRpbSk7YWxpZ24taXRlbXM6Y2VudGVyfQoubGVnZW5kLXRpdGxle2ZvbnQtd2VpZ2h0OjYwMDtjb2xvcjp2YXIoLS10ZXh0LWRpbSk7bGV0dGVyLXNwYWNpbmc6MC41cHg7dGV4dC10cmFuc2Zvcm06dXBwZXJjYXNlO2ZvbnQtc2l6ZToxMHB4fQoubGVnZW5kLWl0ZW17ZGlzcGxheTpmbGV4O2FsaWduLWl0ZW1zOmNlbnRlcjtnYXA6NnB4fQoucGFnZS1mb290ZXJ7bWFyZ2luOjMycHggNDBweCAwO3BhZGRpbmctdG9wOjE2cHg7Ym9yZGVyLXRvcDoxcHggc29saWQgdmFyKC0tYm9yZGVyKTtkaXNwbGF5OmZsZXg7anVzdGlmeS1jb250ZW50OnNwYWNlLWJldHdlZW47YWxpZ24taXRlbXM6Y2VudGVyO2ZvbnQtc2l6ZToxMXB4O2NvbG9yOnZhcigtLXRleHQtZGltKX0KLnBhZ2UtZm9vdGVyIGF7Y29sb3I6dmFyKC0tYWNjZW50KTt0ZXh0LWRlY29yYXRpb246bm9uZX0KLnRhYmxlLXdyYXB7b3ZlcmZsb3cteDphdXRvO2JvcmRlci1yYWRpdXM6NnB4fQo8L3N0eWxlPgo8L2hlYWQ+Cjxib2R5PiUlSEVBREVSJSUlJVNVTU1BUlklJSUlU1RBVFVTQkFOTkVSJSU8ZGl2IGNsYXNzPSJtYWluIj4lJUNMVVNURVJTJSU8L2Rpdj4lJUxFR0VORCUlPGZvb3RlciBjbGFzcz0icGFnZS1mb290ZXIiPjxzcGFuPkdldC1GQ0hCQVBhdGhTdGF0ZSB2Mi41ICZtaWRkb3Q7IDxhIGhyZWY9Imh0dHBzOi8vd3d3LmhvbGxlYm9sbGV2c2FuLm5sIiB0YXJnZXQ9Il9ibGFuayI+aG9sbGVib2xsZXZzYW4ubmw8L2E+ICZtaWRkb3Q7IFBhdWwgdmFuIERpZWVuPC9zcGFuPjxzcGFuPiUlR0VORVJBVEVEJSU8L3NwYW4+PC9mb290ZXI+CjwvYm9keT4KPC9odG1sPg=="
    $template = [System.Text.Encoding]::UTF8.GetString([System.Convert]::FromBase64String($templateB64))

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
    $headerHTML += "<div class=`"header-meta`"><span class=`"badge v`">v2.5</span>"
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
Write-Host "╔═══════════════════════════════════════════╗" -ForegroundColor Cyan
Write-Host "║     Get-FCHBAPathState  v2.5              ║" -ForegroundColor Cyan
Write-Host "║     Paul van Dieen - hollebollevsan.nl    ║" -ForegroundColor Cyan
Write-Host "╚═══════════════════════════════════════════╝" -ForegroundColor Cyan
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
Write-Host ""
