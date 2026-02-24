function Get-GroupMembersRecursively {
    param (
        [string]$groupName
    )

    # Get the group members
    $groupMembers = Get-ADGroupMember -Identity $groupName

    # Display the group members
    foreach ($member in $groupMembers) {
        Write-Output $member.Name

        # If the member is a group, call the function recursively
        if ($member.objectClass -eq "group") {
            Write-Output "Checking members of group: $($member.Name)"
            Get-GroupMembersRecursively -groupName $member.SamAccountName
        }
    }
}

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

# Display the group members recursively
Write-Output "Group Members:"
Get-GroupMembersRecursively -groupName $groupName