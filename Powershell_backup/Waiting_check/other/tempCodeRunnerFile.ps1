# Step 1: Main Script Logic
# The main script prompts for an AD object (user/group) ID and displays detailed information based on the object type.
write-host "`n####### AD Object(DL/SG/User) Info Collection Tool V1.4 #######`n" -ForegroundColor Cyan

$ObjectID=Read-Host "Please input object ID(samAccountName or email address of user account/DL/security group)"
$obj=get-adobject -filter {mail -eq $ObjectID -or samaccountname -eq $ObjectID} -Properties *

#Step 2: Handling User Objects
if($obj.objectClass -eq "user")
{
	write-host "###### Basic Info ###### " -ForegroundColor Cyan
	get-aduser -identity $obj.samaccountname -Properties * | Select-Object DisplayName, @{Label="LastName";Expression={$_.sn}}, @{Label="FirstName";Expression={$_.GivenName}}, sAMAccountName, @{Label="UPN Logon";Expression={$_.userPrincipalName}},@{Label="FullName";Expression={$_.Name}}, Mail,Description, @{Label="ExpireDate";Expression={[datetime]::FromFileTime($_.accountExpires)}},Enabled,@{Label="EmployeeType";Expression={$_.EmployeeType}},homePhone, extensionAttribute1, extensionAttribute4,extensionAttribute5,extensionAttribute9, extensionAttribute11, extensionAttribute12,@{Label="LogonScript";Expression={$_.scriptPath}}, @{Label="OU";Expression={$_.distinguishedName}},@{Label="msExchUMDtmfMap";Expression={$_.msExchUMDtmfMap}},proxyAddresses  | format-list
	
	write-host "###### Member of ###### " -ForegroundColor Cyan
	Get-ADPrincipalGroupMembership -Identity $obj.samaccountname | Select-Object @{Label="Member of";Expression={$_.name}} | format-wide -autosize
	write-host "`n####### End #######`n" -ForegroundColor Cyan

    $today=Get-Date -Format "yyyyMMddHHmm"
    

}
# Step 3: Handling Group Objects
elseif($obj.objectClass -eq "group")
{
    #$cmdfilter="(Name -eq '" + $obj.samaccountname +"')"
    $obj | Select-Object samaccountname,mail,displayname,description,extensionAttribute9,distinguishedName,grouptype,managedby | format-list
    write-host "`n####### Group member(s) of ", $obj.samaccountname ," #######`n" -ForegroundColor Cyan
    get-adgroupmember -identity $obj.samaccountname | format-wide -column 3
    write-host "`n####### End #######`n" -ForegroundColor Cyan
}
#Step 5: Handling Invalid Input
else
{
    write-host "`n####### No Record!!! #######`n" -ForegroundColor Red
}


