#####################################Shell Start#########################################
write-host "`n####### AD Object(DL/SG/User) Info Collection Tool V1.5 #######`n" -ForegroundColor Cyan
#Get user input for object ID and search AD for the object. If it's a user, display basic info and group membership. If not found, show an error message.
$ObjectID=Read-Host "Please input object ID(samAccountName or email address of user account/DL/security group)"
$obj=get-adobject -filter {mail -eq $ObjectID -or samaccountname -eq $ObjectID} -Properties *
if($obj.objectClass -eq "user")
{
    write-host "###### Basic Info ###### " -ForegroundColor Cyan

    $user = Get-ADUser -Identity $obj.samaccountname -Properties *

    $user | Select-Object `
        DisplayName,
        @{Label="LastName";Expression={$_.sn}},
        @{Label="FirstName";Expression={$_.GivenName}},
        sAMAccountName,
        @{Label="UPN Logon";Expression={$_.userPrincipalName}},
        @{Label="FullName";Expression={$_.Name}},
        Mail,
        @{Label="TelephoneNumber";Expression={$_.telephoneNumber}},
        @{Label="JobTitle";Expression={$_.title}},
        @{Label="Department";Expression={$_.department}},
        Description,
        @{Label="ExpireDate";Expression={[datetime]::FromFileTime($_.accountExpires)}},Enabled,
        @{Label="EmployeeType";Expression={$_.EmployeeType}},
        homePhone,
        extensionAttribute1,
        extensionAttribute4,
        extensionAttribute5,
        extensionAttribute9,
        extensionAttribute11,
        extensionAttribute12,
        @{Label="LogonScript";Expression={$_.scriptPath}},
        @{Label="OU";Expression={$_.distinguishedName}},
        @{Label="msExchUMDtmfMap";Expression={$_.msExchUMDtmfMap}},
        proxyAddresses | Format-List

    write-host "###### Member of ###### " -ForegroundColor Cyan

    $Groupmembers = ($user | Select-Object MemberOf).MemberOf.Replace('\','')
    $Groupmembers = $Groupmembers | Sort-Object

    $GroupArray = New-Object -TypeName 'System.Collections.ArrayList'
    foreach($GroupMember in $Groupmembers)
    {
        $GroupArray.Add($GroupMember.Substring(3, $GroupMember.IndexOf(',OU=')-3)) | Out-Null
    }
    $GroupArray | Format-Wide -Property {$_} -Column 3 -Force

    write-host "`n####### End #######`n" -ForegroundColor Cyan
}
else
{
    write-host "`n####### No Record!!! #######`n" -ForegroundColor Red
}

pause

