# Import the list of computer names from a text file
$computers = Get-Content -Path "C:\temp\PCList.txt"
 
# Create an empty array to hold the results
$results = @()
 
foreach ($computer in $computers) {
   # Use WMI to get monitor information
   $monitors = Get-WmiObject -Namespace "root\WMI" -Class "WmiMonitorID" -ComputerName $computer -ErrorAction SilentlyContinue
 
   foreach ($monitor in $monitors) {
       # Convert the model and serial number from byte array to string
       $model = [System.Text.Encoding]::ASCII.GetString($monitor.UserFriendlyName -ne 00)
       $serial = [System.Text.Encoding]::ASCII.GetString($monitor.SerialNumberID -ne 00)
 
       # Add the result to the results array
       $results += New-Object PSObject -Property @{
           "Computer" = $computer
           "Model" = $model
           "SerialNumber" = $serial
       }
   }
}
 
# Export the results to a CSV file
$results | Export-Csv -Path "C:\temp\monitorlist.csv" -NoTypeInformation