#author: kin.yan@wsp.com

#last update:
#V1.0: 2024/4/9


write-host "`n####### AD Account Info Collection Tool V1.0 #######`n"

$useraccount=Read-Host "Please input user ID(SamAccountName or Email address)"
$Sam=(get-aduser -filter {emailaddress -eq $useraccount -or samaccountname -eq $useraccount}).sAMAccountName
get-aduser -identity $Sam -Properties * | select DisplayName, @{Label="LastName";Expression={$_.sn}}, @{Label="FirstName";Expression={$_.GivenName}}, sAMAccountName,@{Label="FullName";Expression={$_.Name}}, Mail,Description, @{Label="PasswordExpiryDate";Expression={[datetime]::FromFileTime($_.accountExpires)}},Enabled,extensionAttribute1, extensionAttribute9, extensionAttribute11, @{Label="LogonScript";Expression={$_.scriptPath}}, @{Label="OU";Expression={$_.distinguishedName}} | format-list
Get-ADPrincipalGroupMembership -Identity $Sam | select @{Label="Member of";Expression={$_.name}} | out-host
write-host "`n####### End #######`n"

pause
