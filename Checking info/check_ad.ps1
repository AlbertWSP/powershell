# Import the Active Directory module
Import-Module ActiveDirectory

# Define the path to the file containing the asset names
$assetFilePath = "C:\temp\powershell\hostname.txt"

# Read the asset names from the file
$assetNames = Get-Content -Path $assetFilePath

# Loop through each asset name and check if it exists in Active Directory
foreach ($assetName in $assetNames) {
    try {
        $asset = Get-ADComputer -Filter { Name -eq $assetName } -ErrorAction Stop
        if ($asset) {
            Write-Host "PC '$assetName' exists in Active Directory." -ForegroundColor Green
        } else {
            Write-Host "PC '$assetName' does not exist in Active Directory." -ForegroundColor Red
        }
    } catch {
        Write-Host "An error occurred while checking asset '$assetName': $_ " -ForegroundColor Red
    }
}