# Function to get the OU path of an AD object
function GetOUPath {
    param (
        [string]$DistinguishedName
    )
    $ouPath = ($DistinguishedName -split ",")[1..($DistinguishedName.Length - 1)] -join ","
    return $ouPath
}

# Prompt the user to enter the group name
$groupName = Read-Host -Prompt "Enter the AD group name"

# Get the group information
$group = Get-ADGroup -Identity $groupName

# Display the group CN and OU path
Write-Output "Group CN: $($group.CN)"
Write-Output "Group OU Path: $(GetOUPath -DistinguishedName $group.DistinguishedName)"
