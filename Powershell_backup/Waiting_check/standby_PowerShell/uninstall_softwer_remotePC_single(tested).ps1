$computerName = "DNXF4M3"
$softwareName = "PVsyst"

Invoke-Command -ComputerName $computerName -ScriptBlock {
    param($name)
    $app = Get-WmiObject -Class Win32_Product | Where-Object { $_.Name -eq $name }
    if ($app) {
        $app.Uninstall()
        Write-Host "$name 已成功卸載。"
    } else {
        Write-Host "$name 未找到。"
    }
} -ArgumentList $softwareName