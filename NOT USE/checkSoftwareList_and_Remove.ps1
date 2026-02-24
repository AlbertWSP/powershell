# Remote Software Management Script

# Function to get the remote PC name and verify connectivity
function Get-RemotePCName {
    do {
        $computerName = Read-Host "Enter the PC name"
        if ([string]::IsNullOrWhiteSpace($computerName)) {
            Write-Host "Hostname cannot be empty. Please try again." -ForegroundColor Red
            continue
        }
        
        # Test if the computer is reachable
        if (Test-Connection -ComputerName $computerName -Count 1 -Quiet) {
            return $computerName
        } else {
            Write-Host "Cannot reach the computer '$computerName'. Please verify the hostname and network connection." -ForegroundColor Red
            $retry = Read-Host "Would you like to try another hostname? (Y/N)"
            if ($retry -notmatch '^[Yy]') {
                exit
            }
        }
    } while ($true)
}

function Get-InstalledSoftware {
    param (
        [string]$ComputerName
    )
    
    try {
        # Create output directory if it doesn't exist
        $outputDir = "c:\temp"
        if (-not (Test-Path $outputDir)) {
            New-Item -ItemType Directory -Path $outputDir | Out-Null
        }

        # Generate filename with timestamp
        $timestamp = Get-Date -Format "yyyyMMddHHmmss"
        $outputFile = Join-Path $outputDir "software_$timestamp.csv"

        # Create the CSV file with headers
        "Computer,Software,Version,Vendor" | Out-File $outputFile -Encoding utf8

        Write-Host "Retrieving software list from $ComputerName..." -ForegroundColor Yellow
        $software = Get-WmiObject -Class Win32_Product -ComputerName $ComputerName

        if ($software) {
            foreach ($program in $software) {
                # Output to CSV file
                "$ComputerName,$($program.Name -replace ',',''),$($program.Version),$($program.Vendor -replace ',')" |
                Out-File $outputFile -Append -Encoding utf8
            }

            # Display the software list in console
            Write-Host "`nInstalled Software List:" -ForegroundColor Green
            $software | Select-Object Name, Version, Vendor | Format-Table -AutoSize

            Write-Host "`nSoftware list has been saved to: $outputFile" -ForegroundColor Cyan
            return $software
        } else {
            Write-Host "No software found on the remote computer." -ForegroundColor Yellow
            "$ComputerName,No Software Found,N/A,N/A" | Out-File $outputFile -Append -Encoding utf8
            return $null
        }
    }
    catch {
        Write-Host "Error retrieving software list: $_" -ForegroundColor Red
        "$ComputerName,Error,$($_.Exception.Message),N/A" | Out-File $outputFile -Append -Encoding utf8
        return $null
    }
}

# Main script execution
try {
    # Step 1: Get remote PC hostname
    $remotePCName = Get-RemotePCName

    do {
        Write-Host "`nRemote Software Management Menu:" -ForegroundColor Cyan
        Write-Host "1. Scan and list installed software" -ForegroundColor Yellow
        Write-Host "2. Exit" -ForegroundColor Yellow

        $choice = Read-Host "`nEnter your choice (1-2)"

        switch ($choice) {
            "1" {
                # Scan and list software
                $softwareList = Get-InstalledSoftware -ComputerName $remotePCName
            }
            "2" {
                Write-Host "Exiting..." -ForegroundColor Cyan
                return
            }
            default {
                Write-Host "Invalid choice. Please enter 1 or 2." -ForegroundColor Red
            }
        }

        $continue = Read-Host "`nWould you like to check another PC? (Y/N)"
        if ($continue -notmatch '^[Yy]') {
            break
        }
    } while ($true)
}
catch {
    Write-Host "An error occurred: $_" -ForegroundColor Red
}
finally {
    Write-Host "`nScript execution completed." -ForegroundColor Cyan
}