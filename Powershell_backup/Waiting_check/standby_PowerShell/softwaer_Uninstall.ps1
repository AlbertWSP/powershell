$hostnameFile = "C:\temp\hostname.txt"
$hostnames = Get-Content -Path $hostnameFile
$logFile = "C:\temp\uninstall_log.txt"

foreach ($hostname in $hostnames) {
    try {
        Invoke-Command -ComputerName $hostname -ScriptBlock {
            $uninstallKey = "HKLM:\Software\Microsoft\Windows\CurrentVersion\Uninstall\PVsyst 7.4.8"
            if (Test-Path $uninstallKey) {
                $uninstallString = (Get-ItemProperty $uninstallKey).UninstallString
                
                # 構建 UNC 路徑
                if ($uninstallString) {
                    $uncPath = "\\$using:hostname\C$\Windows\System32\msiexec.exe" + [System.IO.Path]::GetFileName($uninstallString)
                    
                    # 確認檔案存在
                    if (Test-Path $uncPath) {
                        Start-Process -FilePath $uncPath -ArgumentList "/qn" -Wait
                        Write-Output "[$(Get-Date)] Software uninstalled successfully on ${using:hostname}"
                    } else {
                        Write-Output "[$(Get-Date)] Uninstall executable not found at $uncPath for ${using:hostname}"
                    }
                } else {
                    Write-Output "[$(Get-Date)] Uninstall string not found for ${using:hostname}"
                }
            } else {
                Write-Output "[$(Get-Date)] Uninstall key not found for ${using:hostname}"
            }
        } | Out-File -FilePath $logFile -Append
    } catch {
        $errorMessage = $_.Exception.Message
        $logMessage = "[${Get-Date}] Failed to uninstall software on ${hostname}: ${errorMessage}"
        Write-Output $logMessage | Out-File -FilePath $logFile -Append
    }
}