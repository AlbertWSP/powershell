# Set the folder path containing the CSV files
$csvFolder = "C:\temp\csv"

# Set the output file path
$mergedCsv = Join-Path $csvFolder "merged_software_list.csv"

# Get all CSV files in the folder
$csvFiles = Get-ChildItem -Path $csvFolder -Filter *.csv

if ($csvFiles.Count -eq 0) {
    Write-Host "No CSV files found in $csvFolder" -ForegroundColor Red
    exit
}

# Initialize a flag to track the first file
$firstFile = $true

# Merge files
foreach ($file in $csvFiles) {
    if ($firstFile) {
        # Include header for the first file
        Get-Content $file.FullName | Out-File -FilePath $mergedCsv -Encoding utf8
        $firstFile = $false
    } else {
        # Skip header for subsequent files
        Get-Content $file.FullName | Select-Object -Skip 1 | Out-File -FilePath $mergedCsv -Append -Encoding utf8
    }
}

Write-Host "Merged CSV created at: $mergedCsv" -ForegroundColor Green
