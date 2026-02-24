param (
    [string]$AssetName
)

# Import the Active Directory module
Import-Module ActiveDirectory

# Function to capture the screen
function Capture-Screen {
    Add-Type -AssemblyName System.Windows.Forms
    Add-Type -AssemblyName System.Drawing

    $bounds = [System.Windows.Forms.Screen]::PrimaryScreen.Bounds
    $bitmap = New-Object System.Drawing.Bitmap $bounds.width, $bounds.height
    $graphics = [System.Drawing.Graphics]::FromImage($bitmap)
    $graphics.CopyFromScreen($bounds.Location, [System.Drawing.Point]::Empty, $bounds.Size)

    $filePath = Join-Path $PSScriptRoot "$AssetName.png"
    $bitmap.Save($filePath, "PNG")
    $graphics.Dispose()
    $bitmap.Dispose()
}

# Check if the asset exists in Active Directory
try {
    $asset = Get-ADComputer -Filter { Name -eq $AssetName } -ErrorAction Stop
    if ($asset) {
        Write-Host "**** Asset '$AssetName' exists in Active Directory. ****" -ForegroundColor Green
    } else {
        Write-Host "**** Asset '$AssetName' does not exist in Active Directory. ****" -ForegroundColor Red
    }
} catch {
    Write-Host "**** An error occurred while checking asset '$AssetName': $_ ****" -ForegroundColor Red
}




# Capture the screen
Capture-Screen

# Wait for the second script to complete
Start-Sleep -Seconds 1

# Close the script
exit
