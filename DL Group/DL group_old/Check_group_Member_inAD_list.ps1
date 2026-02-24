# Ensure ActiveDirectory module is loaded
Import-Module ActiveDirectory -ErrorAction SilentlyContinue

function Get-GroupMembersRecursively {
    param (
        [string]$groupName,
        [System.Collections.ArrayList]$membersList,
        [System.Collections.Generic.HashSet[string]]$seenMembers
    )

    try {
        $groupMembers = Get-ADGroupMember -Identity $groupName -ErrorAction Stop
    } catch {
        Write-Warning "Failed to get members of group: $groupName. Error: $_"
        return
    }

    foreach ($member in $groupMembers) {
        if ($seenMembers.Contains($member.DistinguishedName)) {
            continue
        }
        $seenMembers.Add($member.DistinguishedName) | Out-Null

        $email = ""
        if ($member.objectClass -eq "user") {
            $user = Get-ADUser -Identity $member.DistinguishedName -Properties EmailAddress
            $email = $user.EmailAddress
        } elseif ($member.objectClass -eq "group") {
            Get-GroupMembersRecursively -groupName $member.SamAccountName -membersList $membersList -seenMembers $seenMembers
        }

        $membersList.Add([PSCustomObject]@{
            Name        = $member.Name
            ObjectClass = $member.objectClass
            Email       = $email
            ParentGroup = $groupName
        }) | Out-Null
    }
}

# Path to the text file containing the list of group names
$groupListPath = "C:\Temp\groups.txt"

# Read the group names
$groupNames = Get-Content -Path $groupListPath

# Global list to collect all members from all groups
$allMembers = New-Object System.Collections.ArrayList

foreach ($groupName in $groupNames) {
    try {
        $group = Get-ADGroup -Identity $groupName -ErrorAction Stop
    } catch {
        Write-Warning "Group '$groupName' not found. Skipping."
        continue
    }

    Write-Output "Processing Group: $groupName"

    $membersList = New-Object System.Collections.ArrayList
    $seenMembers = New-Object 'System.Collections.Generic.HashSet[string]'

    Get-GroupMembersRecursively -groupName $groupName -membersList $membersList -seenMembers $seenMembers

    # Add this group's members to the global list
    $allMembers.AddRange($membersList)
}

# Export all collected members to a single CSV
$timestamp = Get-Date -Format "yyyyMMdd_HHmmss"
$csvPath = "C:\Temp\AllGroups_Members_$timestamp.csv"
$allMembers | Export-Csv -Path $csvPath -NoTypeInformation -Encoding UTF8

Write-Output "All group member data exported to: $csvPath"
