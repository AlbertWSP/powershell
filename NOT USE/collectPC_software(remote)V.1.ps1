#date and time
$today=Get-Date -Format "yyyyMMddHHmmss"

# Define the output file
$outputFile = "c:\temp\software" + $today+ ".csv"
 
# Check if the output file exists, if so, delete it
if (Test-Path $outputFile) {
   Remove-Item $outputFile
}
 
# Import the list of computers from a text file
$computers = Get-Content "c:\temp\PCList.txt"
 
foreach ($computer in $computers) {
   if (Test-Connection -ComputerName $computer -Count 1 -Quiet) {
       # Get the list of installed software
       $software = Get-WmiObject -Class Win32_Product -ComputerName $computer
 
       foreach ($program in $software) {
           # Output the software details to the CSV file
           $output = New-Object -TypeName PSObject
           $output | Add-Member -MemberType NoteProperty -Name "Computer" -Value $computer
           $output | Add-Member -MemberType NoteProperty -Name "Software" -Value $program.Name
           $output | Add-Member -MemberType NoteProperty -Name "Version" -Value $program.Version
           $output | Export-Csv $outputFile -Append -NoTypeInformation
       }
   } else {
       # The computer is offline or not responding
       $output = New-Object -TypeName PSObject
       $output | Add-Member -MemberType NoteProperty -Name "Computer" -Value $computer
       $output | Add-Member -MemberType NoteProperty -Name "Software" -Value "Offline"
       $output | Export-Csv $outputFile -Append -NoTypeInformation
   }
}