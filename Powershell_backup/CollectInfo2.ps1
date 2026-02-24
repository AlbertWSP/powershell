write-host "`n####### AD Object(DL/SG/User) Info Collection Tool V1.5 #######`n" -ForegroundColor Cyan

$timestamp = Get-Date -Format "yyyyMMdd_HHmmss"
$scriptPath = Split-Path -Parent $MyInvocation.MyCommand.Path

$ObjectID=Read-Host "Please input object ID(samAccountName or email address of user account/DL/security group)"
$obj=get-adobject -filter {mail -eq $ObjectID -or samaccountname -eq $ObjectID} -Properties *
if($obj.objectClass -eq "user")
{
    write-host "###### Basic Info ###### " -ForegroundColor Cyan
    $userInfo = get-aduser -identity $obj.samaccountname -Properties * | Select-Object DisplayName, @{Label="LastName";Expression={$_.sn}}, @{Label="FirstName";Expression={$_.GivenName}}, sAMAccountName, @{Label="UPN Logon";Expression={$_.userPrincipalName}},@{Label="FullName";Expression={$_.Name}}, Mail,Description, @{Label="ExpireDate";Expression={[datetime]::FromFileTime($_.accountExpires)}},Enabled,@{Label="EmployeeType";Expression={$_.EmployeeType}},homePhone, extensionAttribute1, extensionAttribute4,extensionAttribute5,extensionAttribute9, extensionAttribute11, extensionAttribute12,@{Label="LogonScript";Expression={$_.scriptPath}}, @{Label="OU";Expression={$_.distinguishedName}},@{Label="msExchUMDtmfMap";Expression={$_.msExchUMDtmfMap}},proxyAddresses
    $userInfo | format-list

    # Export user info to CSV
    $csvPath = Join-Path $scriptPath "UserInfo_${ObjectID}_${timestamp}.csv"
    $userInfo | Export-Csv -Path $csvPath -NoTypeInformation
    Write-Host "User information exported to: $csvPath" -ForegroundColor Green

    write-host "###### Member of ###### " -ForegroundColor Cyan
    #Get-ADPrincipalGroupMembership is not working normally in Win11
    #Get-ADPrincipalGroupMembership -Identity $obj.samaccountname | Select-Object @{Label="Member of";Expression={$_.name}} | format-wide -autosize
    $Groupmembers=(Get-Aduser $obj.samaccountname -Properties MemberOf | Select-Object MemberOf).MemberOf.Replace('\','')
    $Groupmembers = $Groupmembers | Sort-Object
    
    $GroupArray = New-Object -TypeName 'system.collections.ArrayList'
    foreach($GroupMember in $Groupmembers)
    {
        $GroupArray.Add($GroupMember.Substring(3, $GroupMember.IndexOf(',OU=')-3)) | Out-Null
    }
    $GroupArray | format-wide -Property {$_} -column 3 -Force

    # Export group membership to CSV
    $groupCsvPath = Join-Path $scriptPath "UserGroups_${ObjectID}_${timestamp}.csv"
    $GroupArray | Select-Object @{Name='GroupMembership';Expression={$_}} | Export-Csv -Path $groupCsvPath -NoTypeInformation
    Write-Host "Group membership exported to: $groupCsvPath" -ForegroundColor Green

    write-host "`n####### End #######`n" -ForegroundColor Cyan
}
elseif($obj.objectClass -eq "group")
{
    $groupInfo = $obj | Select-Object samaccountname,mail,displayname,description,extensionAttribute9,distinguishedName,grouptype,managedby
    $groupInfo | format-list
    
    # Export group info to CSV
    $groupCsvPath = Join-Path $scriptPath "GroupInfo_${ObjectID}_${timestamp}.csv"
    $groupInfo | Export-Csv -Path $groupCsvPath -NoTypeInformation
    Write-Host "Group information exported to: $groupCsvPath" -ForegroundColor Green

    write-host "`n####### Group member(s) of ", $obj.samaccountname ," #######`n" -ForegroundColor Cyan
    $groupMembers = get-adgroupmember -identity $obj.samaccountname | Sort-Object
    $groupMembers | format-wide -column 3

    # Export group members to CSV
    $membersCsvPath = Join-Path $scriptPath "GroupMembers_${ObjectID}_${timestamp}.csv"
    $groupMembers | Select-Object name,samAccountName,distinguishedName | Export-Csv -Path $membersCsvPath -NoTypeInformation
    Write-Host "Group members exported to: $membersCsvPath" -ForegroundColor Green

    write-host "`n####### End #######`n" -ForegroundColor Cyan
}
else
{
    write-host "`n####### No Record!!! #######`n" -ForegroundColor Red
}

pause

