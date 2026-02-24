# Load Windows Forms assembly for file dialog
Add-Type -AssemblyName System.Windows.Forms

# Date and time
$today = Get-Date -Format "yyyyMMddHHmmss"

# Define the output file
$outputFile = "c:\temp\software$today.csv"

# Check if the output file exists, if so, delete it
if (Test-Path $outputFile) {
   Remove-Item $outputFile
}

# Create the CSV file with headers
"Computer,Software,Version" | Out-File $outputFile -Encoding utf8

# Create OpenFileDialog
$openFileDialog = New-Object System.Windows.Forms.OpenFileDialog
$openFileDialog.Filter = "Text files (*.txt)|*.txt|All files (*.*)|*.*"
$openFileDialog.Title = "Select computers list file"
$openFileDialog.InitialDirectory = "C:\"

# Show the dialog and get the selected file
if ($openFileDialog.ShowDialog() -eq 'OK') {
    $computers = Get-Content $openFileDialog.FileName
    
    foreach ($computer in $computers) {
       if (Test-Connection -ComputerName $computer -Count 1 -Quiet) {
           # Get the list of installed software
           $software = Get-WmiObject -Class Win32_Product -ComputerName $computer

           foreach ($program in $software) {
               # Output the software details to the CSV file
               "$computer,$($program.Name -replace ',',''),$($program.Version)" | 
               Out-File $outputFile -Append -Encoding utf8
           }
       } else {
           # The computer is offline or not responding
           "$computer,Offline,N/A" | Out-File $outputFile -Append -Encoding utf8
       }
    }
} else {
    Write-Host "No file selected. Exiting script."
    exit
}