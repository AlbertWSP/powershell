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

# Import the list of computers from a text file
$computers = Get-Content "c:\temp\hostname.txt"

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