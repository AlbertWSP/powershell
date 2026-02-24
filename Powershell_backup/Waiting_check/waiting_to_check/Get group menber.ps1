Import-Module ActiveDirectory

# Prompt for group name
$groupName = Read-Host "Please enter the AD Group name"

try {
    $ADGroup = Get-ADGroup -Filter {Name -eq $groupName} -ErrorAction Stop
    if ($null -eq $ADGroup) {
        Write-Host "Group not found."
        exit
    }

    Write-Host "Retrieving group members..."
    $members = Get-ADGroupMember -Identity $ADGroup.DistinguishedName -Recursive

    $results = @()
    $total = $members.Count
    $current = 0

    foreach ($member in $members) {
        $current++
        Write-Progress -Activity "Processing Users" -Status "$current of $total" -PercentComplete (($current / $total) * 100)

        if ($member.objectClass -eq "user") {
            $user = Get-ADUser -Identity $member.DistinguishedName -Properties Name, UserPrincipalName, Department, Title
            $results += [PSCustomObject]@{
                Name = $user.Name
                Email = $user.UserPrincipalName
                Department = $user.Department
                Title = $user.Title
            }
        }
    }

    $fileName = "C:\temp\ADGroupMembers_$groupName.csv"
    $results | Export-Csv -Path $fileName -NoTypeInformation

    Write-Host "Export completed. File saved as: $fileName"
}
catch {
    Write-Host "An error occurred: $_"
}