param (
    [string]$AssetName
)

# Import the Active Directory module
Import-Module ActiveDirectory

# Import the necessary .NET assemblies
Add-Type -AssemblyName System.Windows.Forms
Add-Type -AssemblyName System.Drawing

# Function to capture the active window
function Capture-ActiveWindow {
    # Get the handle of the active window
    $activeWindowHandle = [System.Windows.Forms.Form]::ActiveForm.Handle

    # Get the bounds of the active window
    $bounds = [System.Windows.Forms.Screen]::FromHandle($activeWindowHandle).Bounds
    $bitmap = New-Object System.Drawing.Bitmap $bounds.Width, $bounds.Height
    $graphics = [System.Drawing.Graphics]::FromImage($bitmap)
    $graphics.CopyFromScreen($bounds.Location, [System.Drawing.Point]::Empty, $bounds.Size)

    # Save the screenshot
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

# Capture the active window
Capture-ActiveWindow

# Wait for the second script to complete
Start-Sleep -Seconds 1

# Close the script
exit