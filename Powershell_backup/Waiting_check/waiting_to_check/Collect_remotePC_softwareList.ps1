# Prompt for the remote computer name
$remoteComputer = Read-Host -Prompt "Enter the remote computer name"

$scriptBlock = {
    try {
        # 確保目標目錄存在
        $directory = "C:\temp"
        if (-Not (Test-Path -Path $directory)) {
            New-Item -ItemType Directory -Path $directory
            Write-Host "Directory created: $directory"
        } else {
            Write-Host "Directory already exists: $directory"
        }

        # 查詢電腦名稱
        $computerName = $env:COMPUTERNAME
        Write-Host "Computer name: $computerName"

        # 查詢 32 位和 64 位應用程式
        $installedSoftware = @()

        # 查詢 64 位應用程式
        $installedSoftware += Get-ItemProperty -Path "HKLM:\SOFTWARE\Microsoft\Windows\CurrentVersion\Uninstall\*" | Select-Object DisplayName, DisplayVersion, Publisher, InstallDate
        Write-Host "64-bit applications found: $($installedSoftware.Count)"

        # 查詢 32 位應用程式
        $installedSoftware += Get-ItemProperty -Path "HKLM:\SOFTWARE\Wow6432Node\Microsoft\Windows\CurrentVersion\Uninstall\*" | Select-Object DisplayName, DisplayVersion, Publisher, InstallDate
        Write-Host "Total applications found: $($installedSoftware.Count)"

        # 將電腦名稱添加到每個軟體物件中
        $installedSoftware = $installedSoftware | ForEach-Object {
            $_ | Add-Member -MemberType NoteProperty -Name ComputerName -Value $computerName -PassThru
        }

        # 設定文件名為電腦名稱+InstalledSoftware.csv
        $filePath = "$directory\$computerName-InstalledSoftware.csv"
        Write-Host "File path: $filePath"

        # 將安裝的軟體列表導出到 CSV 文件
        $installedSoftware | Export-Csv -Path $filePath -NoTypeInformation
        Write-Host "CSV file created: $filePath"

        # 返回檔案路徑，以便後續處理
        return $filePath
    }
    catch {
        Write-Error "An error occurred: $_"
        return $null
    }
}

# 執行遠端命令
$result = Invoke-Command -ComputerName $remoteComputer -ScriptBlock $scriptBlock

# 顯示結果
if ($result) {
    Write-Host "CSV file created on remote computer: $result"
} else {
    Write-Host "Failed to create CSV file on remote computer."
}