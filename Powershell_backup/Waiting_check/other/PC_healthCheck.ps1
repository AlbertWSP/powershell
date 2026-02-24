# 設定日誌文件路徑
$LogFilePath = "C:\Temp\PC_Health_Check_Log.txt"
$date = Get-Date -Format "yyyy-MM-dd HH:mm:ss"

# 初始化日誌文件
"PC Health Check Report - $date" | Out-File -FilePath $LogFilePath -Encoding UTF8
"---------------------------------------------" | Out-File -FilePath $LogFilePath -Append

# 定義函數來記錄日誌
function Write-Log {
    param (
        [string]$Message
    )
    Write-Host $Message
    $Message | Out-File -FilePath $LogFilePath -Append
}

# 1. 檢查 CPU 使用率
Write-Log "Checking CPU Usage..."
$cpuUsage = (Get-Counter '\Processor(_Total)\% Processor Time').CounterSamples.CookedValue
Write-Log "CPU Usage: $([math]::round($cpuUsage, 2))%"

# 2. 檢查記憶體使用情況
Write-Log "Checking Memory Usage..."
$totalMemory = (Get-CimInstance Win32_ComputerSystem).TotalPhysicalMemory / 1GB
$freeMemory = (Get-Counter '\Memory\Available MBytes').CounterSamples.CookedValue / 1024
$usedMemory = $totalMemory - $freeMemory
Write-Log "Total Memory: $([math]::round($totalMemory, 2)) GB"
Write-Log "Used Memory: $([math]::round($usedMemory, 2)) GB"
Write-Log "Free Memory: $([math]::round($freeMemory, 2)) GB"

# 3. 檢查磁碟狀態
Write-Log "Checking Disk Health..."
$drives = Get-WmiObject Win32_LogicalDisk -Filter "DriveType=3"
foreach ($drive in $drives) {
    $freeSpaceGB = [math]::round($drive.FreeSpace / 1GB, 2)
    $totalSpaceGB = [math]::round($drive.Size / 1GB, 2)
    Write-Log "Drive $($drive.DeviceID): Free Space: $freeSpaceGB GB / Total Space: $totalSpaceGB GB"
}

# 4. 檢查網絡連接狀態
Write-Log "Checking Network Connectivity..."
$pingResult = Test-Connection -ComputerName google.com -Count 1 -ErrorAction SilentlyContinue
if ($pingResult) {
    Write-Log "Network Connectivity: Online"
} else {
    Write-Log "Network Connectivity: Offline"
}

# 5. 檢查系統文件完整性 (SFC)
Write-Log "Running System File Checker (SFC)..."
$sfcResult = sfc /scannow | Out-String
if ($sfcResult -match "Windows Resource Protection did not find any integrity violations") {
    Write-Log "SFC Result: No integrity violations found."
} elseif ($sfcResult -match "Windows Resource Protection found corrupt files and successfully repaired them") {
    Write-Log "SFC Result: Corrupt files found and repaired."
} else {
    Write-Log "SFC Result: Issues detected. Please check the logs for details."
}

# 6. 檢查磁碟錯誤 (CHKDSK)
Write-Log "Running Disk Check (CHKDSK)..."
$systemDrive = $env:SystemDrive.TrimEnd(':')
$chkdskResult = chkdsk "${systemDrive}:" /scan | Out-String
if ($chkdskResult -match "Windows has scanned the file system and found no problems") {
    Write-Log "CHKDSK Result: No file system errors found."
} else {
    Write-Log "CHKDSK Result: Issues detected. Please check the logs for details."
}

# 7. 總結健康檢查結果
Write-Log ""
Write-Log "Health Check Completed. Detailed results are saved in: $LogFilePath"
