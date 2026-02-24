# Function to convert bytes to MB or GB
function Format-FileSize {
    param (
        [long]$Size
    )
    
    if ($Size -ge 1GB) {
        return "{0:N2} GB" -f ($Size / 1GB)
    }
    else {
        return "{0:N2} MB" -f ($Size / 1MB)
    }
}

# Function to get color based on size
function Get-SizeColor {
    param (
        [long]$Size
    )
    
    if ($Size -ge 1GB) {
        return "Red"
    }
    elseif ($Size -ge 500MB) {
        return "Yellow"
    }
    else {
        return "Green"
    }
}

# Read the list of folders from the text file
$folders = Get-Content -Path "C:\temp\folders.txt"

# Initialize an array to store the results
$results = @()

Write-Host "`nFolder Size Analysis" -ForegroundColor Cyan
Write-Host "==================" -ForegroundColor Cyan
Write-Host ""

foreach ($folder in $folders) {
    # Get the size of the folder and the number of files
    $folderSize = (Get-ChildItem -Path $folder -Recurse | Measure-Object -Property Length -Sum).Sum
    $fileCount = (Get-ChildItem -Path $folder -Recurse | Measure-Object).Count

    # Create a custom object to store the results
    $result = [PSCustomObject]@{
        FolderPath = $folder
        FolderSize = Format-FileSize -Size $folderSize
        FileCount  = $fileCount
    }

    # Add the result to the array
    $results += $result

    # Display results with color
    Write-Host "Folder: " -NoNewline
    Write-Host $folder -ForegroundColor White
    Write-Host "Size: " -NoNewline
    Write-Host (Format-FileSize -Size $folderSize) -ForegroundColor (Get-SizeColor -Size $folderSize)
    Write-Host "Files: " -NoNewline
    Write-Host $fileCount -ForegroundColor Magenta
    Write-Host "-------------------"
}

Write-Host "`nExporting results to CSV..." -ForegroundColor Cyan

# Export the results to a CSV file
$results | Export-Csv -Path "C:\temp\Folders_output.csv" -NoTypeInformation

Write-Host "Results exported to C:\temp\Folders_output.csv" -ForegroundColor Green