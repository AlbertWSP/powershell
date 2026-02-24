$PathToPCsTextFile = "C:\temp\PCList.txt"
$PathToPasswordsTextFile = "C:\temp\PWD.txt"
$PathToCSVFile = "C:\temp\monitorlist.csv"
 
$PCs = Get-Content $PathToPCsTextFile
$Passwords = Get-Content $PathToPasswordsTextFile
 
if ($PCs.Count -ne $Passwords.Count) {
   #Write-Error "The number of PCs and passwords do not match."
   #return
}
 
$monitorInfo = @()
 
for ($i = 0; $i -lt $PCs.Count; $i++) {
   $RemotePCName = $PCs[$i]
   $Username = "hkan739291it5"
   $Password = ConvertTo-SecureString $Passwords[0] -AsPlainText -Force
   $cred = New-Object System.Management.Automation.PSCredential($Username, $Password)
 
   try {
       $session = New-PSSession -ComputerName $RemotePCName -Credential $cred -ErrorAction Stop
 
       $result = Invoke-Command -Session $session -ScriptBlock {
           $monitors = Get-WmiObject -Namespace root\wmi -Class WmiMonitorID
           if ($monitors) {
               $monitors | ForEach-Object {
                   $serialNumber = ($_.SerialNumberID -ne 0 | ForEach-Object { [char]$_ }) -join ""
                   $userFriendlyName = ($_.UserFriendlyName -ne 0 | ForEach-Object { [char]$_ }) -join ""
                   New-Object PSObject -Property @{
                       "PCName" = $env:COMPUTERNAME
                      "MonitorModel" = $userFriendlyName
                      "SerialNumber" = $serialNumber
                   }
               }
           } else {
               New-Object PSObject -Property @{
                   "PCName" = $env:COMPUTERNAME
                   "MonitorModel" = "None"
                   "SerialNumber" = "None"
               }
           }
       }
 
       $monitorInfo += $result
 
       Remove-PSSession $session
   } catch {
       $monitorInfo += New-Object PSObject -Property @{
           "PCName" = $RemotePCName
           "MonitorModel" = "Offline"
           "SerialNumber" = "Offline"
       }
   }
}
 
$monitorInfo | Sort-Object MonitorModel, SerialNumber -Unique | Export-Csv -Path $PathToCSVFile -NoTypeInformation