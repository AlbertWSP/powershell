sfc /scannow# Define the path to the text file containing hostnames
$hostFilePath = "C:\temp\hostnames.txt"

# Check if the file exists
if (-not (Test-Path $hostFilePath)) {
    Write-Host "Hostnames file not found at $hostFilePath" -ForegroundColor Red
    exit
}

# Read all hostnames from the file
$hostnames = Get-Content $hostFilePath

# Create output directory if it doesn't exist
$outputDir = "C:\temp"
if (-not (Test-Path $outputDir)) {
    New-Item -ItemType Directory -Path $outputDir | Out-Null
}

foreach ($computerName in $hostnames) {
    $computerName = $computerName.Trim()
    if ([string]::IsNullOrWhiteSpace($computerName)) {
        continue
    }

    Write-Host "`nProcessing ${computerName}..." -ForegroundColor Cyan

    # Test connectivity
    if (-not (Test-Connection -ComputerName $computerName -Count 1 -Quiet)) {
        Write-Host "Cannot reach ${computerName}. Skipping..." -ForegroundColor Red
        continue
    }

    try {
        # Generate output file name
        $timestamp = Get-Date -Format "yyyyMMddHHmmss"
        $outputFile = Join-Path $outputDir "${computerName}_software_${timestamp}.csv"

        # Write header
        "Computer,Software,Version,Vendor" | Out-File $outputFile -Encoding utf8

        # Get installed software
        $software = Get-WmiObject -Class Win32_Product -ComputerName $computerName

        if ($software) {
            foreach ($program in $software) {
                "${computerName},$($program.Name -replace ',',''),$($program.Version),$($program.Vendor -replace ',')" |
                Out-File $outputFile -Append -Encoding utf8
            }

            Write-Host "Software list saved to ${outputFile}" -ForegroundColor Green
        } else {
            Write-Host "No software found on ${computerName}" -ForegroundColor Yellow
            "${computerName},No Software Found,N/A,N/A" | Out-File $outputFile -Append -Encoding utf8
        }
    }
    catch {
        Write-Host "Error retrieving software from ${computerName}: $($_.Exception.Message)" -ForegroundColor Red
        "${computerName},Error,$($_.Exception.Message),N/A" | Out-File $outputFile -Append -Encoding utf8
    }
}

Write-Host "`nAll hosts processed." -ForegroundColor Cyan
