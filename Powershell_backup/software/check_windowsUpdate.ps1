# Check if Windows Update service is running
$service = Get-Service -Name wuauserv -ErrorAction SilentlyContinue

if ($null -eq $service) {
    Write-Host "Windows Update service (wuauserv) not found." -ForegroundColor Red
} elseif ($service.Status -ne 'Running') {
    Write-Host "Windows Update service is not running. Attempting to start..." -ForegroundColor Yellow
    try {
        Start-Service -Name wuauserv
        Write-Host "Service started successfully." -ForegroundColor Green
    } catch {
        Write-Host "Failed to start the service: $_" -ForegroundColor Red
    }
} else {
    Write-Host "Windows Update service is running." -ForegroundColor Green
}

# Check for Windows Update module
if (Get-Module -ListAvailable -Name PSWindowsUpdate) {
    Import-Module PSWindowsUpdate
    Write-Host "PSWindowsUpdate module loaded." -ForegroundColor Green

    # Check update history
    $history = Get-WUHistory | Select-Object -First 5
    Write-Host "Recent Windows Update History:" -ForegroundColor Cyan
    $history | Format-Table -AutoSize

    # Try to initiate a scan
    Write-Host "Initiating update scan..." -ForegroundColor Cyan
    Get-WindowsUpdate -MicrosoftUpdate -AcceptAll -IgnoreReboot
} else {
    Write-Host "PSWindowsUpdate module not found. Please install it using:" -ForegroundColor Yellow
    Write-Host "Install-Module -Name PSWindowsUpdate -Force"
}
