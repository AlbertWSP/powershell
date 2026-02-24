# Load Windows Forms assembly for file dialog
Add-Type -AssemblyName System.Windows.Forms

# Date and time for output file
$today = Get-Date -Format "yyyyMMddHHmmss"
$outputFile = "c:\temp\OSVersion_$today.csv"

# Create OpenFileDialog
$openFileDialog = New-Object System.Windows.Forms.OpenFileDialog
$openFileDialog.Filter = "Text files (*.txt)|*.txt|All files (*.*)|*.*"
$openFileDialog.Title = "Select computers list file"
$openFileDialog.InitialDirectory = "C:\"

# Check if output file exists, if so, delete it
if (Test-Path $outputFile) {
    Remove-Item $outputFile
}

# Create CSV file with headers
"Computer,OSVersion,ServicePack,OSArchitecture,Status" | Out-File $outputFile -Encoding utf8

# Show file dialog and process the selected file
if ($openFileDialog.ShowDialog() -eq 'OK') {
    $computers = Get-Content $openFileDialog.FileName
    
    foreach ($computer in $computers) {
        if (Test-Connection -ComputerName $computer -Count 1 -Quiet) {
            try {
                $os = Get-WmiObject -Class Win32_OperatingSystem -ComputerName $computer -ErrorAction Stop
                "$computer,$($os.Caption),$($os.ServicePackMajorVersion),$($os.OSArchitecture),Online" |
                Out-File $outputFile -Append -Encoding utf8
            }
            catch {
                "$computer,Error accessing WMI,N/A,N/A,Error" |
                Out-File $outputFile -Append -Encoding utf8
            }
        }
        else {
            "$computer,N/A,N/A,N/A,Offline" |
            Out-File $outputFile -Append -Encoding utf8
        }
    }
    Write-Host "Scan complete. Results saved to: $outputFile"
}
else {
    Write-Host "No file selected. Exiting script."
    exit
}
