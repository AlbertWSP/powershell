# Define the path to the file containing the asset names
$assetFilePath = Join-Path $PSScriptRoot "assets.txt"
$checkassetsFilePath = Join-Path $PSScriptRoot "CheckAssetAndCapture.ps1"

# Read the asset names from the file
$assetNames = Get-Content -Path $assetFilePath

# Loop through each asset name and call the second script
foreach ($assetName in $assetNames) {
    # Call the second script and pass the asset name as an argument
    Start-Process powershell -ArgumentList "-File `"$checkassetsFilePath`" -AssetName `"$assetName`""

    # Wait for the second script to complete
    Start-Sleep -Seconds 3
}

