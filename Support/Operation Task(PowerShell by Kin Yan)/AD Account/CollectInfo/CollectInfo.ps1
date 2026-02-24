function ExportToCSV
{
    Param([string] $FilePath, [PSObject] $Data)

    try {
        $Data | Export-Csv -Path $FilePath -NoTypeInformation -Encoding UTF8
        Write-Host "Successfully exported to CSV: $FilePath" -ForegroundColor Green
    } catch {
        Write-Error "Failed to export to CSV. More Info: $_"
    }
}



write-host "`n####### AD Object(DL/SG/User) Info Collection Tool V1.5 #######`n" -ForegroundColor Cyan

$ObjectID=Read-Host "Please input object ID(samAccountName or email address of user account/DL/security group)"
$obj=get-adobject -filter {mail -eq $ObjectID -or samaccountname -eq $ObjectID} -Properties *
if($obj.objectClass -eq "user")
{
    write-host "###### Basic Info ###### " -ForegroundColor Cyan
    $UserInfo = get-aduser -identity $obj.samaccountname -Properties * | Select-Object DisplayName, @{Label="LastName";Expression={$_.sn}}, @{Label="FirstName";Expression={$_.GivenName}}, sAMAccountName, @{Label="UPN Logon";Expression={$_.userPrincipalName}},@{Label="FullName";Expression={$_.Name}}, Mail,Description, @{Label="ExpireDate";Expression={[datetime]::FromFileTime($_.accountExpires)}},Enabled,@{Label="EmployeeType";Expression={$_.EmployeeType}},homePhone, extensionAttribute1, extensionAttribute4,extensionAttribute5,extensionAttribute9, extensionAttribute11, extensionAttribute12,@{Label="LogonScript";Expression={$_.scriptPath}}, @{Label="OU";Expression={$_.distinguishedName}},@{Label="msExchUMDtmfMap";Expression={$_.msExchUMDtmfMap}},proxyAddresses
    $UserInfo | format-list
	#$obj | select DisplayName, @{Label="LastName";Expression={$_.sn}}, @{Label="FirstName";Expression={$_.GivenName}}, sAMAccountName,@{Label="FullName";Expression={$_.Name}}, Mail,Description, @{Label="ExpireDate";Expression={[datetime]::FromFileTime($_.accountExpires)}},Enabled,extensionAttribute1, extensionAttribute9, extensionAttribute11, @{Label="LogonScript";Expression={$_.scriptPath}}, @{Label="OU";Expression={$_.distinguishedName}} | format-list 
	#Get-ADPrincipalGroupMembership -Identity $obj.samaccountname | select @{Label="Member of";Expression={$_.name}} | out-host
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

	write-host "`n####### End #######`n" -ForegroundColor Cyan

    $today=Get-Date -Format "yyyyMMddHHmm"

    [System.Reflection.Assembly]::LoadWithPartialName("System.Windows.Forms") | Out-Null
    $SaveFileDialog = New-Object System.Windows.Forms.SaveFileDialog
    $SaveFileDialog.initialDirectory = [Environment]::GetFolderPath('Desktop')
    $SaveFileDialog.filter = "CSV file (*.CSV)| *.CSV"
    $SaveFileDialog.title = "Save AD object information to a CSV file..."
    $SaveFileDialog.filename = $ObjectID + "_" + $today
    $R=$SaveFileDialog.ShowDialog()

    if($R -eq "OK")
    {
        ExportToCSV -FilePath $SaveFileDialog.filename -Data $UserInfo
    }
}
elseif($obj.objectClass -eq "group")
{
    $GroupInfo = $obj | Select-Object samaccountname,mail,displayname,description,extensionAttribute9,distinguishedName,grouptype,managedby
    $GroupInfo | format-list
    write-host "`n####### Group member(s) of ", $obj.samaccountname ," #######`n" -ForegroundColor Cyan
    $GroupMembers = get-adgroupmember -identity $obj.samaccountname | Sort-Object
    $GroupMembers | format-wide -column 3
    write-host "`n####### End #######`n" -ForegroundColor Cyan

    $today=Get-Date -Format "yyyyMMddHHmm"
    [System.Reflection.Assembly]::LoadWithPartialName("System.Windows.Forms") | Out-Null
    $SaveFileDialog = New-Object System.Windows.Forms.SaveFileDialog
    $SaveFileDialog.initialDirectory = [Environment]::GetFolderPath('Desktop')
    $SaveFileDialog.filter = "CSV file (*.CSV)| *.CSV"
    $SaveFileDialog.title = "Save AD group members to a CSV file..."
    $SaveFileDialog.filename = $ObjectID + "_" + $today
    $R=$SaveFileDialog.ShowDialog()
    if($R -eq "OK")
    {
        ExportToCSV -FilePath $SaveFileDialog.filename -Data $GroupMembers
    }
}
else
{
    write-host "`n####### No Record!!! #######`n" -ForegroundColor Red
}



