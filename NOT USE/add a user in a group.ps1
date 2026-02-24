# Specify the user and group names
$groupName = "HK-GRP-HKKWN200-CS"
$userName = "HKAC731108"

# Get the group object
$group = Get-ADGroup -Identity $groupName

# Get the user object
$user = Get-ADUser -Identity $userName

# Add the user to the group
Add-ADGroupMember -Identity $group -Members $user