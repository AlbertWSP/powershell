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

# Prompt for the computer name or IP address
$computer = Read-Host "Enter the PC name"

try {
    # Check if the computer is online using Test-Connection
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
} catch {
    # Handle any errors
    "$computer,Error,$($_.Exception.Message)" | Out-File $outputFile -Append -Encoding utf8
}
