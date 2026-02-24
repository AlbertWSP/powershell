# Prompt user for the base path
$basePath = "\\dcasi300dat20\Projects2"

# Create the main project folder
$projectFolderName = Read-Host "Enter the name of the main project folder:"

# Check if the main project folder already exists
if (Test-Path -Path (Join-Path -Path $basePath -ChildPath $projectFolderName)) {
    Write-Host "Main project folder already exists."
} else {
    # Create the main project folder
    New-Item -Path $basePath -Name $projectFolderName -ItemType Directory
    Write-Host "Main project folder created successfully."
}

# Create subfolders within the main project folder
$folders = @(
    "00_PM (Restricted)",
    "01 Admin",
    "02 Correspondence",
    "03 Deliverables",
    "04 Working",
    "05 Reference",
    "06 Project Brief",
    "07 Meeting"
)

foreach ($folderName in $folders) {
    $subFolderPath = Join-Path -Path $basePath -ChildPath $projectFolderName
    if (Test-Path -Path (Join-Path -Path $subFolderPath -ChildPath $folderName)) {
        Write-Host "Subfolder '$folderName' already exists."
    } else {
        New-Item -Path $subFolderPath -Name $folderName -ItemType Directory
        Write-Host "Subfolder '$folderName' created successfully."
    }
}
