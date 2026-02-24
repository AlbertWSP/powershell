# Prompt the user to enter the group name
$groupName = Read-Host -Prompt "Enter the AD group name"

# Get the group information
$group = Get-ADGroup -Identity $groupName

# Display the group information
Write-Output "Group Name: $($group.Name)"
Write-Output "Group Description: $($group.Description)"
Write-Output "Group Distinguished Name: $($group.DistinguishedName)"
Write-Output "Group Object GUID: $($group.ObjectGUID)"
Write-Output "Group Created Date: $($group.WhenCreated)"
Write-Output "Group Modified Date: $($group.WhenChanged)"
Write-Output ""

# Get the group members
$groupMembers = Get-ADGroupMember -Identity $groupName

# Display the group members
Write-Output "Group Members:"
$groupMembers | ForEach-Object {
    Write-Output $_.Name
}
