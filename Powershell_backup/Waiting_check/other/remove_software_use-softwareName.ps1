# 指定要卸載的軟體名稱
$softwareName = "Java 8 Update 421"

# 查找 64 位和 32 位系統中的卸載字符串
$uninstallKeyPaths = @(
    "HKLM:\SOFTWARE\Microsoft\Windows\CurrentVersion\Uninstall\*",
    "HKLM:\SOFTWARE\Wow6432Node\Microsoft\Windows\CurrentVersion\Uninstall\*"
)

# 查找卸載字符串
$uninstallString = $null
foreach ($keyPath in $uninstallKeyPaths) {
    $uninstallString = Get-ItemProperty -Path $keyPath | Where-Object { $_.DisplayName -eq $softwareName } | Select-Object -ExpandProperty UninstallString -ErrorAction SilentlyContinue
    if ($uninstallString) {
        break
    }
}

# 檢查是否找到卸載字符串
if ($uninstallString) {
    Write-Output "找到卸載字符串: $uninstallString"
    # 執行卸載命令
    Start-Process -FilePath "cmd.exe" -ArgumentList "/c $uninstallString /quiet" -Wait
    Write-Output "$softwareName 已成功卸載。"
} else {
    Write-Output "未找到名為 $softwareName 的軟體。"
}

Get-Package -Name $softwareName