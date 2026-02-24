# 定義遠端電腦名稱或 IP
$RemotePC = "3BQZ9Y3"

# 建立遠端執行區塊
Invoke-Command -ComputerName $RemotePC -ScriptBlock {
    $uninstaller = "C:\Program Files\Bulk Rename Utility\unins000.exe"
    
    # 檢查路徑是否存在
    if (Test-Path $uninstaller) {
        Write-Host "正在遠端解除安裝 Bulk Rename Utility..."
        
        # 使用 Start-Process 執行靜默參數
        Start-Process -FilePath $uninstaller -ArgumentList "/VERYSILENT", "/NORESTART", "/SUPPRESSMSGBOXES" -Wait
        
        Write-Host "解除安裝程序已完成。"
    } else {
        Write-Warning "在遠端路徑中找不到卸載程式。"
    }
}
