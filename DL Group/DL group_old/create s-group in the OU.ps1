# Prompt the user to enter the group name
$groupName = Read-Host -Prompt "Enter the AD group name"

# The OU path
$ouPath = "OU=Security,OU=Groups,OU=HK,OU=WSPObjects,DC=corp,DC=pbwan,DC=net"

# Create the DL group in the specified OU
New-ADGroup -Name $groupName -GroupScope Global -GroupCategory Security -Path $ouPath

Write-Output "Distribution List group '$groupName' has been created in OU '$ouPath'."