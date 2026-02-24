# Function to get the SAM account name from an email address
function GetSamAccountName {
    param (
        [string]$Email
    )
    $user = Get-ADUser -Filter {Mail -eq $Email} -Properties SamAccountName
    return $user.SamAccountName
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

# Get the group members
$groupMembers = Get-ADGroupMember -Identity $groupName

# Display the group members in alphabetical order
Write-Output "Group Members:"
$groupMembers | Sort-Object Name | ForEach-Object {
    Write-Output $_.Name
}

# Prompt the user to confirm if they want to add new members
$confirmAddMembers = Read-Host -Prompt "Do you want to add new members to the group? (Y/n)"

if ($confirmAddMembers -eq "Y" -or $confirmAddMembers -eq "y") {
    # Prompt the user to enter the new members' email addresses or SAM account names (comma-separated)
    $newMembersInput = Read-Host -Prompt "Enter the email addresses or SAM account names of the new members to add (comma-separated)"

    # Split the input into an array
    $newMembersArray = $newMembersInput -split ","

    # Initialize an array to hold the SAM account names
    $newMembersSamAccountNames = @()

    # Process each item in the list
    foreach ($member in $newMembersArray) {
        $member = $member.Trim()
        if ($member -match "^[\w\.-]+@[\w\.-]+\.\w+$") {
            # Input is an email address, get the SAM account name
            $samAccountName = GetSamAccountName -Email $member
        } else {
            # Input is a SAM account name
            $samAccountName = $member
        }
        $newMembersSamAccountNames += $samAccountName
    }

    # Add the new members to the group
    Add-ADGroupMember -Identity $groupName -Members $newMembersSamAccountNames

    # Confirm the new members have been added
    $updatedGroupMembers = Get-ADGroupMember -Identity $groupName

    # Display the updated group members in alphabetical order
    Write-Output "Updated Group Members:"
    $updatedGroupMembers | Sort-Object Name | ForEach-Object {
        Write-Output $_.Name
    }
} elseif ($confirmAddMembers -eq "N" -or $confirmAddMembers -eq "n") {
    Write-Output "No members were added to the group."
} else {
    Write-Output "Invalid input. No members were added to the group."
}