param (
    [string]$InstallerPath,      # Path to the installer on the remote computer
    [string]$Arguments = ""      # Optional arguments for the installer
)

# Prompt the user to input the remote computer name
$RemoteComputerName = Read-Host "Enter the name of the remote computer"

try {
    # Use Invoke-Command to run the installer on the remote computer
    Invoke-Command -ComputerName $RemoteComputerName -ScriptBlock {
        param ($InstallerPath, $Arguments)
        Start-Process -FilePath $InstallerPath -ArgumentList $Arguments -Wait
    } -ArgumentList $InstallerPath, $Arguments

    Write-Host "Installation started successfully on $RemoteComputerName."
}
catch {
    Write-Host "Failed to start the installation on $RemoteComputerName. Error: $_"
}