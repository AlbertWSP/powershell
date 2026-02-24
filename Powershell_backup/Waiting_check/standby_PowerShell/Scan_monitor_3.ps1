# Import the list of computer names from a text file
$computers = Get-Content -Path "C:\temp\PCList.txt"

# Create an empty array to hold the results
$results = @()

foreach ($computer in $computers) {
    try {
        # Use WMI to get monitor information
        $monitors = Get-WmiObject -Namespace "root\WMI" -Class "WmiMonitorID" -ComputerName $computer -ErrorAction Stop

        foreach ($monitor in $monitors) {
            # Convert the model and serial number from byte array to string, removing null bytes
            $model = ([System.Text.Encoding]::ASCII.GetString($monitor.UserFriendlyName) -replace '\x00', '').Trim()
            $serial = ([System.Text.Encoding]::ASCII.GetString($monitor.SerialNumberID) -replace '\x00', '').Trim()

            # Create a result object
            $result = [PSCustomObject]@{
                "Computer" = $computer
                "Model" = $model
                "SerialNumber" = $serial
            }

            # Output the result to the console
            Write-Output $result

            # Add the result to the results array
            $results += $result
        }
    } catch {
        Write-Warning "Failed to query ${computer}: $_"
        continue
    }
}

# Export the results to a CSV file
$results | Export-Csv -Path "C:\temp\monitorlist.csv" -NoTypeInformation