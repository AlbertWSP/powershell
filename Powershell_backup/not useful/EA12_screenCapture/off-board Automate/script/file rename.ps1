# Define the path to the CSV file
$csvPath = "C:\temp\off-board\off-board_assets.csv"

# Import the CSV file
$csv = Import-Csv -Path $csvPath

# Loop through each row in the CSV
foreach ($row in $csv) {
    # Define the current and new file names
    $currentName = "C:\temp\off-board\" + $row.OldName
    $newName = "C:\temp\off-board\DONE" + $row.NewName

    # Rename the file
    Rename-Item -Path $currentName -NewName $newName
}

