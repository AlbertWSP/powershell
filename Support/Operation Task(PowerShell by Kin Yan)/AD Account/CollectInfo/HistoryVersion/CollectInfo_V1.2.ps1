#author: kin.yan@wsp.com

#last update:
#V1.2: 2024/5/16
#New: 

#V1.1: 2024/4/15
#New: this new version is able to collect information from DL&SG

#V1.0: 2024/4/9
#initial release

write-host "`n####### AD Object(DL/SG/User) Info Collection Tool V1.2 #######`n"

$ObjectID=Read-Host "Please input object ID(samAccountName or email address of user account/DL/security group)"
$obj=get-adobject -filter {mail -eq $ObjectID -or samaccountname -eq $ObjectID} -Properties *
if($obj.objectClass -eq "user")
{
    get-aduser -identity $obj.samaccountname -Properties * | select DisplayName, @{Label="LastName";Expression={$_.sn}}, @{Label="FirstName";Expression={$_.GivenName}}, sAMAccountName, @{Label="UPN Logon";Expression={$_.userPrincipalName}},@{Label="FullName";Expression={$_.Name}}, Mail,Description, @{Label="ExpireDate";Expression={[datetime]::FromFileTime($_.accountExpires)}},Enabled,@{Label="EmployeeType";Expression={$_.EmployeeType}},homePhone,extensionAttribute1, extensionAttribute9, extensionAttribute11, @{Label="LogonScript";Expression={$_.scriptPath}}, @{Label="OU";Expression={$_.distinguishedName}} | format-list
    #$obj | select DisplayName, @{Label="LastName";Expression={$_.sn}}, @{Label="FirstName";Expression={$_.GivenName}}, sAMAccountName,@{Label="FullName";Expression={$_.Name}}, Mail,Description, @{Label="ExpireDate";Expression={[datetime]::FromFileTime($_.accountExpires)}},Enabled,extensionAttribute1, extensionAttribute9, extensionAttribute11, @{Label="LogonScript";Expression={$_.scriptPath}}, @{Label="OU";Expression={$_.distinguishedName}} | format-list
    Get-ADPrincipalGroupMembership -Identity $obj.samaccountname | select @{Label="Member of";Expression={$_.name}} | out-host
    write-host "`n####### End #######`n"
}
elseif($obj.objectClass -eq "group")
{
    $cmdfilter="(Name -eq '" + $obj.samaccountname +"')"
    $obj | select samaccountname,mail,displayname,description,extensionAttribute9,distinguishedName,grouptype,managedby | format-list
    write-host "`n####### Group member(s) of ", $obj.samaccountname ," #######`n"
    get-adgroupmember -identity $obj.samaccountname | format-wide -column 3
    write-host "`n####### End #######`n"
}
else
{
    write-host "`n####### No Record!!! #######`n"
}


pause