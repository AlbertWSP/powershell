<#
.SYNOPSIS
    PC Information Scanner - Collects hardware, software, battery, driver, and system info.
.DESCRIPTION
    Generates a comprehensive HTML report of the local or remote PC.
.NOTES
    Author : Albert Ng / IT Onsite Support
    Version: 1.0
    Date   : 2026-05-18
#>

# ========================== CONFIG ==========================
$ReportPath = "C:\Temp\PC_Report_$($env:COMPUTERNAME)_$(Get-Date -Format 'yyyyMMdd_HHmmss').html"

# ========================== FUNCTIONS ==========================
function Get-HtmlSection {
    param([string]$Title, [string]$Content)
    @"
    <div class="section">
        <h2>$Title</h2>
        <table>$Content</table>
    </div>
"@
}

function Get-HtmlRow {
    param([string]$Label, [string]$Value)
    "<tr><td class='label'>$Label</td><td>$Value</td></tr>"
}

function Get-HtmlTableFromObjects {
    param([array]$Objects)
    if (-not $Objects) { return "<tr><td>No data found.</td></tr>" }
    $headers = $Objects[0].PSObject.Properties.Name
    $html = "<tr>" + ($headers | ForEach-Object { "<th>$_</th>" }) -join "" + "</tr>"
    foreach ($obj in $Objects) {
        $html += "<tr>"
        foreach ($h in $headers) {
            $html += "<td>$($obj.$h)</td>"
        }
        $html += "</tr>"
    }
    return $html
}

# ========================== DATA COLLECTION ==========================
Write-Host "🔍 Scanning PC information... Please wait." -ForegroundColor Cyan

# ---------- 1. OS Information ----------
Write-Host "  [1/9] Operating System..."
$os = Get-CimInstance Win32_OperatingSystem
$osRows = Get-HtmlRow "Computer Name"      $os.CSName
$osRows += Get-HtmlRow "OS Name"             $os.Caption
$osRows += Get-HtmlRow "Version / Build"     "$($os.Version) (Build $($os.BuildNumber))"
$osRows += Get-HtmlRow "Architecture"        $os.OSArchitecture
$osRows += Get-HtmlRow "Install Date"        $os.InstallDate.ToString("yyyy-MM-dd HH:mm")
$osRows += Get-HtmlRow "Last Boot"           $os.LastBootUpTime.ToString("yyyy-MM-dd HH:mm")
$osRows += Get-HtmlRow "Registered User"     $os.RegisteredUser
$osRows += Get-HtmlRow "Windows Directory"   $os.WindowsDirectory

# ---------- 2. Hardware - CPU ----------
Write-Host "  [2/9] CPU..."
$cpu = Get-CimInstance Win32_Processor
$cpuRows = Get-HtmlRow "Processor"          $cpu.Name
$cpuRows += Get-HtmlRow "Cores / Threads"    "$($cpu.NumberOfCores) Cores / $($cpu.NumberOfLogicalProcessors) Threads"
$cpuRows += Get-HtmlRow "Max Clock Speed"    "$($cpu.MaxClockSpeed) MHz"
$cpuRows += Get-HtmlRow "Socket"             $cpu.SocketDesignation
$cpuRows += Get-HtmlRow "L2 Cache (KB)"      $cpu.L2CacheSize
$cpuRows += Get-HtmlRow "L3 Cache (KB)"      $cpu.L3CacheSize

# ---------- 3. Hardware - Memory ----------
Write-Host "  [3/9] Memory..."
$totalRAM = [math]::Round($os.TotalVisibleMemorySize / 1MB, 2)
$freeRAM = [math]::Round($os.FreePhysicalMemory / 1MB, 2)
$ramSticks = Get-CimInstance Win32_PhysicalMemory | Select-Object `
@{N = 'Slot'; E = { $_.DeviceLocator } },
@{N = 'Capacity (GB)'; E = { [math]::Round($_.Capacity / 1GB, 1) } },
@{N = 'Speed (MHz)'; E = { $_.Speed } },
@{N = 'Type'; E = { $_.MemoryType } },
@{N = 'Manufacturer'; E = { $_.Manufacturer } },
@{N = 'Part Number'; E = { $_.PartNumber.Trim() } }

$memRows = Get-HtmlRow "Total RAM"   "$totalRAM GB"
$memRows += Get-HtmlRow "Free RAM"    "$freeRAM GB"
$memTable = Get-HtmlTableFromObjects $ramSticks

# ---------- 4. Hardware - Motherboard & BIOS ----------
Write-Host "  [4/9] Motherboard & BIOS..."
$mb = Get-CimInstance Win32_BaseBoard
$bios = Get-CimInstance Win32_BIOS
$cs = Get-CimInstance Win32_ComputerSystem

$mbRows = Get-HtmlRow "System Manufacturer"  $cs.Manufacturer
$mbRows += Get-HtmlRow "System Model"         $cs.Model
$mbRows += Get-HtmlRow "Serial Number"        $bios.SerialNumber
$mbRows += Get-HtmlRow "Motherboard"          "$($mb.Manufacturer) - $($mb.Product)"
$mbRows += Get-HtmlRow "BIOS Version"         $bios.SMBIOSBIOSVersion
$mbRows += Get-HtmlRow "BIOS Date"            $bios.ReleaseDate.ToString("yyyy-MM-dd")

# ---------- 5. Disk / Storage ----------
Write-Host "  [5/9] Disk & Storage..."
$disks = Get-CimInstance Win32_DiskDrive | Select-Object `
@{N = 'Model'; E = { $_.Model } },
@{N = 'Interface'; E = { $_.InterfaceType } },
@{N = 'Size (GB)'; E = { [math]::Round($_.Size / 1GB, 1) } },
@{N = 'Partitions'; E = { $_.Partitions } },
@{N = 'Serial'; E = { $_.SerialNumber.Trim() } },
@{N = 'Status'; E = { $_.Status } }

$volumes = Get-CimInstance Win32_LogicalDisk -Filter "DriveType=3" | Select-Object `
@{N = 'Drive'; E = { $_.DeviceID } },
@{N = 'Label'; E = { $_.VolumeName } },
@{N = 'File System'; E = { $_.FileSystem } },
@{N = 'Total (GB)'; E = { [math]::Round($_.Size / 1GB, 1) } },
@{N = 'Free (GB)'; E = { [math]::Round($_.FreeSpace / 1GB, 1) } },
@{N = 'Free %'; E = { if ($_.Size -gt 0) { [math]::Round(($_.FreeSpace / $_.Size) * 100, 1) }else { 'N/A' } } }

# ---------- 6. GPU / Display ----------
Write-Host "  [6/9] GPU & Display..."
$gpus = Get-CimInstance Win32_VideoController | Select-Object `
@{N = 'GPU Name'; E = { $_.Name } },
@{N = 'Driver Version'; E = { $_.DriverVersion } },
@{N = 'VRAM (GB)'; E = { [math]::Round($_.AdapterRAM / 1GB, 1) } },
@{N = 'Resolution'; E = { "$($_.CurrentHorizontalResolution) x $($_.CurrentVerticalResolution)" } },
@{N = 'Status'; E = { $_.Status } }

# ---------- 7. Network Adapters ----------
Write-Host "  [7/9] Network Adapters..."
$nics = Get-CimInstance Win32_NetworkAdapterConfiguration -Filter "IPEnabled=True" | Select-Object `
@{N = 'Adapter'; E = { $_.Description } },
@{N = 'IP Address'; E = { ($_.IPAddress | Where-Object { $_ -match '\d+\.\d+\.\d+\.\d+' }) -join ', ' } },
@{N = 'Subnet'; E = { $_.IPSubnet -join ', ' } },
@{N = 'Gateway'; E = { $_.DefaultIPGateway -join ', ' } },
@{N = 'DNS'; E = { $_.DNSServerSearchOrder -join ', ' } },
@{N = 'MAC Address'; E = { $_.MACAddress } },
@{N = 'DHCP'; E = { $_.DHCPEnabled } }

# ---------- 8. Battery (Laptop) ----------
Write-Host "  [8/9] Battery..."
$battery = Get-CimInstance Win32_Battery -ErrorAction SilentlyContinue
if ($battery) {
    $batRows = Get-HtmlRow "Battery Name"           $battery.Name
    $batRows += Get-HtmlRow "Status"                  $battery.Status
    $batRows += Get-HtmlRow "Charge (%)"              "$($battery.EstimatedChargeRemaining)%"
    $batRows += Get-HtmlRow "Estimated Runtime (min)" $(if ($battery.EstimatedRunTime -eq 71582788) { "Plugged In / Charging" }else { $battery.EstimatedRunTime })
    $batRows += Get-HtmlRow "Chemistry"               $(switch ($battery.Chemistry) { 1 { "Other" }; 2 { "Unknown" }; 3 { "Lead Acid" }; 4 { "Nickel Cadmium" }; 5 { "Nickel Metal Hydride" }; 6 { "Lithium-ion" }; 7 { "Zinc Air" }; 8 { "Lithium Polymer" }; default { $battery.Chemistry } })
    $batRows += Get-HtmlRow "Design Voltage (mV)"     $battery.DesignVoltage

    # Battery health via powercfg (optional)
    try {
        $batReportPath = "$env:TEMP\battery-report.xml"
        & powercfg /batteryreport /xml /output $batReportPath 2>$null
        if (Test-Path $batReportPath) {
            [xml]$batXml = Get-Content $batReportPath
            $designCap = $batXml.BatteryReport.Batteries.Battery.DesignCapacity
            $fullCap = $batXml.BatteryReport.Batteries.Battery.FullChargeCapacity
            if ($designCap -and $fullCap -and [int]$designCap -gt 0) {
                $healthPct = [math]::Round(([int]$fullCap / [int]$designCap) * 100, 1)
                $batRows += Get-HtmlRow "Design Capacity (mWh)"      $designCap
                $batRows += Get-HtmlRow "Full Charge Capacity (mWh)" $fullCap
                $batRows += Get-HtmlRow "Battery Health"             "$healthPct%"
            }
            Remove-Item $batReportPath -Force -ErrorAction SilentlyContinue
        }
    }
    catch {}
}
else {
    $batRows = Get-HtmlRow "Battery" "No battery detected (Desktop PC)"
}

# ---------- 9. Installed Software ----------
Write-Host "  [9/9] Installed Software & Drivers..."
$software = @()
$regPaths = @(
    "HKLM:\SOFTWARE\Microsoft\Windows\CurrentVersion\Uninstall\*",
    "HKLM:\SOFTWARE\WOW6432Node\Microsoft\Windows\CurrentVersion\Uninstall\*"
)
$software = Get-ItemProperty $regPaths -ErrorAction SilentlyContinue |
Where-Object { $_.DisplayName } |
Select-Object `
@{N = 'Name'; E = { $_.DisplayName } },
@{N = 'Version'; E = { $_.DisplayVersion } },
@{N = 'Publisher'; E = { $_.Publisher } },
@{N = 'Install Date'; E = { $_.InstallDate } } |
Sort-Object Name -Unique

# ---------- 10. Drivers ----------
$drivers = Get-CimInstance Win32_PnPSignedDriver |
Where-Object { $_.DeviceName } |
Select-Object `
@{N = 'Device'; E = { $_.DeviceName } },
@{N = 'Driver Version'; E = { $_.DriverVersion } },
@{N = 'Manufacturer'; E = { $_.Manufacturer } },
@{N = 'Driver Date'; E = { if ($_.DriverDate) { $_.DriverDate.ToString("yyyy-MM-dd") }else { 'N/A' } } },
@{N = 'Signer'; E = { $_.Signer } } |
Sort-Object Device

# ========================== BUILD HTML REPORT ==========================
Write-Host "`n📝 Generating HTML report..." -ForegroundColor Yellow

$htmlReport = @"
<!DOCTYPE html>
<html>
<head>
<meta charset="UTF-8">
<title>PC Report - $($os.CSName)</title>
<style>
    body { font-family: 'Segoe UI', Arial, sans-serif; margin: 20px; background: #f5f5f5; color: #333; }
    h1 { color: #0078D4; border-bottom: 3px solid #0078D4; padding-bottom: 10px; }
    h2 { color: #fff; background: #0078D4; padding: 8px 15px; border-radius: 5px; margin-top: 25px; }
    .section { background: #fff; padding: 15px; border-radius: 8px; margin-bottom: 10px; box-shadow: 0 2px 5px rgba(0,0,0,0.1); }
    table { width: 100%; border-collapse: collapse; margin-top: 8px; }
    th { background: #e8e8e8; padding: 8px; text-align: left; border: 1px solid #ccc; }
    td { padding: 6px 10px; border: 1px solid #ddd; }
    tr:nth-child(even) { background: #f9f9f9; }
    .label { font-weight: bold; width: 250px; background: #f0f0f0; }
    .footer { text-align: center; color: #888; margin-top: 30px; font-size: 12px; }
    @media print { body { background: #fff; } .section { box-shadow: none; } }
</style>
</head>
<body>
<h1>PC Inventory Report - $($os.CSName)</h1>
<p><strong>Generated:</strong> $(Get-Date -Format 'yyyy-MM-dd HH:mm:ss') &nbsp;|&nbsp; <strong>Scanned by:</strong> $($env:USERNAME)</p>

$(Get-HtmlSection "Operating System"        $osRows)
$(Get-HtmlSection "Processor (CPU)"         $cpuRows)
$(Get-HtmlSection "Memory (RAM) — Summary"  $memRows)
$(Get-HtmlSection "Memory (RAM) — Slots"    (Get-HtmlTableFromObjects $ramSticks))
$(Get-HtmlSection "Motherboard & BIOS"      $mbRows)
$(Get-HtmlSection "Disk Drives"             (Get-HtmlTableFromObjects $disks))
$(Get-HtmlSection "Volumes / Partitions"    (Get-HtmlTableFromObjects $volumes))
$(Get-HtmlSection "GPU / Display"           (Get-HtmlTableFromObjects $gpus))
$(Get-HtmlSection "Network Adapters"        (Get-HtmlTableFromObjects $nics))
$(Get-HtmlSection "Battery"                 $batRows)
$(Get-HtmlSection "Installed Software ($($software.Count) apps)" (Get-HtmlTableFromObjects $software))
$(Get-HtmlSection "Drivers ($($drivers.Count) drivers)"         (Get-HtmlTableFromObjects $drivers))

<div class="footer">PC Inventory Report — Generated by PowerShell | WSP IT Support</div>
</body>
</html>
"@

# ========================== EXPORT ==========================
$htmlReport | Out-File -FilePath $ReportPath -Encoding UTF8
Write-Host "Report saved to: $ReportPath" -ForegroundColor Green
Start-Process $ReportPath  # Auto-open in browser