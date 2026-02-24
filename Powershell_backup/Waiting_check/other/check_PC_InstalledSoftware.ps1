# 確保目標目錄存在
$directory = "C:\temp"
if (-Not (Test-Path -Path $directory)) {
    New-Item -ItemType Directory -Path $directory
}


# 或者，導出到 CSV 文件
Get-WmiObject -Class Win32_Product | Select-Object -Property Name, Version | Export-Csv -Path "$directory\InstalledSoftware.csv" -NoTypeInformation