# Prompt the user to enter the group name
$groupName = Read-Host -Prompt "Enter the AD group name"

# Prompt the user to enter the OU path
$ouPath = Read-Host -Prompt "Enter the OU path (e.g., OU=Groups,DC=example,DC=com)"

# Create the group in the specified OU
New-ADGroup -Name $groupName -GroupScope Global -GroupCategory Security -Path $ouPath

Write-Output "Group '$groupName' has been created in OU '$ouPath'."