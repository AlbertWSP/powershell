#author: kin.yan@wsp.com

#last update:
#V1.5: 2024/10/17
#Fix: Get-ADPrincipalGroupMembership is not working normally in Win11, change to a workaround solution
#New: Add sorting to the membership

#V1.4: 2024/6/18
#New: More property added for validation
#New: Generate screenshot automatically

#V1.3: 2024/06/05
#New: show one more property(msExchUMDtmfMap) for validation, if Teams phone number assigned to a user, you can find the number from this property(reverse number format)

#V1.2: 2024/5/16
#New: shit, forget what's new in this version

#V1.1: 2024/4/15
#New: this new version is able to collect information from DL&SG

#V1.0: 2024/4/9
#initial release



#####################################Shell Start#########################################

function TakeScreenshot
{
    Param([string] $ObjectID)

Add-Type -AssemblyName System.Windows.Forms
Add-Type -AssemblyName System.Drawing

Add-Type @"
        using System;
        using System.Runtime.InteropServices;
        public class UserWindow {
        [DllImport("user32.dll")]
        [return: MarshalAs(UnmanagedType.Bool)]
        public static extern bool GetWindowRect(
            IntPtr hWnd, out RECT lpRect);

        [DllImport("user32.dll")]
        public static extern IntPtr GetForegroundWindow();
        }

        public struct RECT
        {
        public int Left;
        public int Top;
        public int Right;
        public int Bottom;
        }
"@
try {
    $ActiveHandle = [UserWindow]::GetForegroundWindow()
    $Rectangle = New-Object RECT
    [UserWindow]::GetWindowRect($ActiveHandle,[ref]$Rectangle)
} catch {            
    Write-Error "Failed to get active Window details. More Info: $_"
}

$Width  = $Rectangle.Right-$Rectangle.Left
$Height = $Rectangle.Bottom-$Rectangle.Top
$Left   = $Rectangle.Left
$Top    = $Rectangle.Top

$bitmap  = New-Object System.Drawing.Bitmap $Width, $Height
$graphic = [System.Drawing.Graphics]::FromImage($bitmap)
$graphic.CopyFromScreen($Left, $Top, 0, 0, $bitmap.Size)

#$today=Get-Date -Format "yyyyMMddHHmm"
$FileName=$ObjectID
$bitmap.Save($FileName,"PNG")
}



write-host "`n####### AD Object(DL/SG/User) Info Collection Tool V1.5 #######`n" -ForegroundColor Cyan

$ObjectID=Read-Host "Please input object ID(samAccountName or email address of user account/DL/security group)"
$obj=get-adobject -filter {mail -eq $ObjectID -or samaccountname -eq $ObjectID} -Properties *
if($obj.objectClass -eq "user")
{
	write-host "###### Basic Info ###### " -ForegroundColor Cyan
	get-aduser -identity $obj.samaccountname -Properties * | Select-Object DisplayName, @{Label="LastName";Expression={$_.sn}}, @{Label="FirstName";Expression={$_.GivenName}}, sAMAccountName, @{Label="UPN Logon";Expression={$_.userPrincipalName}},@{Label="FullName";Expression={$_.Name}}, Mail,Description, @{Label="ExpireDate";Expression={[datetime]::FromFileTime($_.accountExpires)}},Enabled,@{Label="EmployeeType";Expression={$_.EmployeeType}},homePhone, extensionAttribute1, extensionAttribute4,extensionAttribute5,extensionAttribute9, extensionAttribute11, extensionAttribute12,@{Label="LogonScript";Expression={$_.scriptPath}}, @{Label="OU";Expression={$_.distinguishedName}},@{Label="msExchUMDtmfMap";Expression={$_.msExchUMDtmfMap}},proxyAddresses  | format-list
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
    
    [System.Reflection.Assembly]::LoadWithPartialName(“System.windows.forms”) | Out-Null
    $SaveFileDialog = New-Object System.Windows.Forms.SaveFileDialog
    $SaveFileDialog.initialDirectory = [Environment]::GetFolderPath('Desktop')
    $SaveFileDialog.filter = “PNG file (*.PNG)| *.PNG”
    $SaveFileDialog.title = "Save screenshot to a PNG file..."
    $SaveFileDialog.filename = $ObjectID + "_" + $today
    $R=$SaveFileDialog.ShowDialog()
    
    if($R -eq "OK")
    {
        TakeScreenshot -ObjectID $SaveFileDialog.filename
    }
}
elseif($obj.objectClass -eq "group")
{
    #$cmdfilter="(Name -eq '" + $obj.samaccountname +"')"
    $obj | Select-Object samaccountname,mail,displayname,description,extensionAttribute9,distinguishedName,grouptype,managedby | format-list
    write-host "`n####### Group member(s) of ", $obj.samaccountname ," #######`n" -ForegroundColor Cyan
    get-adgroupmember -identity $obj.samaccountname | Sort-Object | format-wide -column 3
    write-host "`n####### End #######`n" -ForegroundColor Cyan
}
else
{
    write-host "`n####### No Record!!! #######`n" -ForegroundColor Red
}

pause

