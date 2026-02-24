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

# Prompt for the IP address
$ipAddress = Read-Host "Enter the IP address"

try {
    # Check if the computer is online using Test-Connection
    if (Test-Connection -ComputerName $ipAddress -Count 1 -Quiet) {
        # Get the list of installed software using the IP address
        $software = Get-WmiObject -Class Win32_Product -ComputerName $ipAddress

        foreach ($program in $software) {
            # Output the software details to the CSV file
            "$ipAddress,$($program.Name -replace ',',''),$($program.Version)" | 
            Out-File $outputFile -Append -Encoding utf8
        }
    } else {
        # The computer is offline or not responding
        "$ipAddress,Offline,N/A" | Out-File $outputFile -Append -Encoding utf8
    }
} catch {
    # Handle any errors
    "$ipAddress,Error,$($_.Exception.Message)" | Out-File $outputFile -Append -Encoding utf8
}
