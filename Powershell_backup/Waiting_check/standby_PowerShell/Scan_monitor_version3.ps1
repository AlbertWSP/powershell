# Import the list of computer names from a text file
$computers = Get-Content -Path "C:\temp\PCList.txt"

# Create an empty array to hold the results
$results = @()

foreach ($computer in $computers) {
    try {
        # Use WMI to get monitor information
        $monitors = Get-WmiObject -Namespace "root\WMI" -Class "WmiMonitorID" -ComputerName $computer -ErrorAction Stop

        foreach ($monitor in $monitors) {
            if ($monitor.UserFriendlyName -ne $null) {
                $model = ([System.Text.Encoding]::ASCII.GetString($monitor.UserFriendlyName) -replace '\x00', '').Trim()
            } else {
                $model = "Unknown"
            }

            if ($monitor.SerialNumberID -ne $null) {
                $serial = ([System.Text.Encoding]::ASCII.GetString($monitor.SerialNumberID) -replace '\x00', '').Trim()
            } else {
                $serial = "Unknown"
            }

            $result = [PSCustomObject]@{
                "Computer" = $computer
                "Model" = $model
                "SerialNumber" = $serial
            }

            Write-Output $result
            $results += $result
        }
    } catch {
        Write-Warning "Failed to query ${computer}: $_.Exception.Message"
        continue
    }
}

# Get the current date and time
$currentDateTime = Get-Date -Format "yyyyMMdd_HHmmss"

# Export the results to a CSV file with date and time in the filename
$results | Export-Csv -Path "C:\temp\monitorlist_$currentDateTime.csv" -NoTypeInformation
