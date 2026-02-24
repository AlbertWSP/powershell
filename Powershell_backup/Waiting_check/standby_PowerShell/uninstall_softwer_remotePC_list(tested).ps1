$hostnames = Get-Content -Path "C:\temp\hostname.txt"
$softwareName = "PVsyst"

foreach ($computerName in $hostnames) {
    Invoke-Command -ComputerName $computerName -ScriptBlock {
        param($name)
        $app = Get-WmiObject -Class Win32_Product | Where-Object { $_.Name -eq $name }
        if ($app) {
            $app.Uninstall()
            Write-Host "$name success uninstall。"
        } else {
            Write-Host "$name cannot find。"
        }
    } -ArgumentList $softwareName
}