param(
    [string]$remoteComputer
)
$credential = Get-Credential

$scriptBlock = {
    # 查詢電腦名稱
    $computerName = $env:COMPUTERNAME

    # 查詢 32 位和 64 位應用程式
    $installedSoftware = @()

    # 查詢 64 位應用程式
    $installedSoftware += Get-ItemProperty -Path "HKLM:\SOFTWARE\Microsoft\Windows\CurrentVersion\Uninstall\*" | Select-Object DisplayName, DisplayVersion, Publisher, InstallDate

    # 查詢 32 位應用程式
    $installedSoftware += Get-ItemProperty -Path "HKLM:\SOFTWARE\Wow6432Node\Microsoft\Windows\CurrentVersion\Uninstall\*" | Select-Object DisplayName, DisplayVersion, Publisher, InstallDate

    # 將電腦名稱添加到每個軟體物件中
    $installedSoftware = $installedSoftware | ForEach-Object {
        $_ | Add-Member -MemberType NoteProperty -Name ComputerName -Value $computerName -PassThru
    }

    # 返回安裝的軟體列表
    return $installedSoftware
}

# 執行遠端命令並獲取結果
$installedSoftware = Invoke-Command -ComputerName $remoteComputer -ScriptBlock $scriptBlock -Credential $credential

# 確保目標目錄存在
$directory = "C:\temp"
if (-Not (Test-Path -Path $directory)) {
    New-Item -ItemType Directory -Path $directory
}

# 設定文件名為電腦名稱+InstalledSoftware.csv
$computerName = $remoteComputer
$filePath = "$directory\$computerName-InstalledSoftware.csv"

# 將安裝的軟體列表導出到 CSV 文件
$installedSoftware | Export-Csv -Path $filePath -NoTypeInformation

Write-Host "CSV file saved locally at: $filePath"