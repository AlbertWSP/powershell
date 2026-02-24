# Replace "GroupName" with the actual name of your AD group
$UserEmails = "user1@example.com", "user2@example.com", "user3@example.com"
$GroupName = "GroupName"

foreach ($Email in $UserEmails) {
    Add-ADGroupMember -Identity $GroupName -Members (Get-ADUser -Filter {EmailAddress -eq $Email})
}
