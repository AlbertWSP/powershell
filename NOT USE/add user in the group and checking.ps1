# Specify the group names and CSV file paths
$group1Name = "HK-GRP-HKKWN200-CS"
$group2Name = "GRP-FSS-HKKWN200-STR"
$csvFilePath = "C:\temp\STR_user_list.csv"
$outputFilePath = "C:\temp\STR_user_list_checking_result.csv"

# Get the group objects
$group1 = Get-ADGroup -Identity $group1Name
$group2 = Get-ADGroup -Identity $group2Name

# Import the CSV file
$users = Import-Csv -Path $csvFilePath

$results = @()

# Add each user to both groups and check membership
foreach ($user in $users) {
    $userName = $user.UserName
    $userObj = Get-ADUser -Identity $userName

    # Check if user is already a member of Group 1
    $isMemberOfGroup1 = $group1 | Get-ADGroupMember | Where-Object { $_.SamAccountName -eq $userName }

    # Check if user is already a member of Group 2
    $isMemberOfGroup2 = $group2 | Get-ADGroupMember | Where-Object { $_.SamAccountName -eq $userName }

    if ($isMemberOfGroup1 -and $isMemberOfGroup2) {
        $result = [PSCustomObject]@{
            UserName = $userName
            Group1 = $true
            Group2 = $true
        }
        Write-Host "User '$userName' is already a member of both groups."
    } else {
        # Add user to Group 1
        if (-not $isMemberOfGroup1) {
            Add-ADGroupMember -Identity $group1 -Members $userObj
            Write-Host "User '$userName' added to Group 1."
        }

        # Add user to Group 2
        if (-not $isMemberOfGroup2) {
            Add-ADGroupMember -Identity $group2 -Members $userObj
            Write-Host "User '$userName' added to Group 2."
        }

        $result = [PSCustomObject]@{
            UserName = $userName
            Group1 = (-not $isMemberOfGroup1)
            Group2 = (-not $isMemberOfGroup2)
        }
    }

    $results += $result
}

# Export the results to a CSV file
$results | Export-Csv -Path $outputFilePath -NoTypeInformation

Write-Host "Results exported to $outputFilePath."
