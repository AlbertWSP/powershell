# Check if running as administrator
$isAdmin = ([Security.Principal.WindowsPrincipal] [Security.Principal.WindowsIdentity]::GetCurrent()).IsInRole([Security.Principal.WindowsBuiltInRole]::Administrator)
if (-not $isAdmin) {
    Write-Host "This script requires administrator privileges." -ForegroundColor Red
    exit 1
}

try {
    # Create installation directory
    $FolderPath = "C:\WindowsUpgrade"
    if (!(Test-Path -Path $FolderPath)) {
        New-Item -Path $FolderPath -ItemType Directory -Force
    }

    # Download Windows 11 Installation Assistant
    $DownloadUrl = "https://go.microsoft.com/fwlink/?linkid=2171764"
    $InstallerPath = "$FolderPath\Windows11InstallationAssistant.exe"
        
    Write-Host "Downloading Windows 11 Installation Assistant..." -ForegroundColor Yellow
    Invoke-WebRequest -Uri $DownloadUrl -OutFile $InstallerPath -ErrorAction Stop

    # Start installation
    Write-Host "Starting Windows 11 installation..." -ForegroundColor Yellow
    $processStartInfo = @{
        FilePath = $InstallerPath
        ArgumentList = "/auto upgrade /quiet /noreboot /Compat IgnoreWarning"
        Wait = $true
        PassThru = $true
    }
    
    $process = Start-Process @processStartInfo
    
    if ($process.ExitCode -ne 0) {
        throw "Installation failed with exit code: $($process.ExitCode)"
    }

    Write-Host "Installation completed successfully!" -ForegroundColor Green
}
catch {
    Write-Host "Error: $_" -ForegroundColor Red
    exit 1
}
