#Test the remote PC network connection
Test-Connection -ComputerName "4qsgc54" -Count 1
Test-WsMan -ComputerName "4qsgc54"
$LocalMachine = $env:COMPUTERNAME
$TargetMachine = "4qsgc54"

if ($LocalMachine -eq $TargetMachine) {
    Write-Output "This is the local computer."
} else {
    Write-Output "This is a remote computer."
}

Invoke-Command -ComputerName "4qsgc54" -ScriptBlock { Get-Culture }
