$computerListPath = "C:\temp\Powershell\pc_list.txt"

$scriptPath = "C:\temp\powershell\collectPC_software.ps1"

$credentialPath = "C:\temp\powershell\credential.xml"

$Credential = Import-Clixml -Path $credentialPath

$computerNames = Get-Content -Path $computerListPath

foreach($computerName in $computerNames){
    Write-Host "Running script for Computer : $computerName"
    & $scriptPath -remoteComputer $computerName -Credential $credential
    }