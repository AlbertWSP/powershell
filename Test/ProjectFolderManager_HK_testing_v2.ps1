###
#Requirement:
#1) PowerShell 5.x & 7.x;
#2) RSAT-AD-PowerShell feature installed;

#Initial Setup:
#1. Copy this shell and the XML profile to--->"C:\Support\ProjectFolderManager"
#1. Enable PS: Set-ExecutionPolicy RemoteSigned -Scope CurrentUser
#2. Install RSAT-AD-PowerShell feature: Install-WindowsFeature RSAT-AD-PowerShell
#    (you can check whether RSAT is installed on your PC by this PS command: Get-Module -Name ActiveDirectory -Listavailable)

#Author: Kin.Yan@wsp.com, Albert.ng@wsp.com
#Description: This is the most powerful tool for project folder management in the planet, i'm not joking, you can feel how powerful it is.......from the window title :)

#What's new in V4.0.0(2025/12/16)
#修改了 RemoveFolderPermission
    #添加验证逻辑，检查权限是否真的被删除
    #区分"权限不存在"和"删除失败"两种情况
    #异常处理中包含错误消息
#修改了 GrantFolderPermission
    #添加 AD 身份验证，确保组存在（包含重试机制）
    #等待 2 秒后验证权限是否真的被设置
    #支持多种查找 AD 组的方法（备用方案）
    #详细的错误报告
#新增了 存在检查
    #CreateADGroup：如果组已存在，直接返回成功，不创建重复组
    #CreateFolder：如果文件夹已存在，直接返回成功，跳过创建


#What's new in V4.0.0(2025/12/15)
#0 To add a wait time after group creation, you would need to insert a Start-Sleep command after line 106 (after New-ADGroup).  the code start at line 112

#What's new in V4.0.0(2025/12/11)
#0. Add a new HK team MEPH in projectFolders.xml line 1879 to 1889

#What's new in V4.0.0(2024/5/30)
#0. Starting from this version, this shell is compatible with Powershell 7
#1. New: Validate target security groups and project folder before creation, if yes, alert and stop the whole process
#2. New: Only when profile version is higher than minimum version the Shell can be opened
#3. Change: Project name restriction changed, only these characters are not allow: \ / : * ? "" < > |
#4. Change: Description is mandatory
#5. Change: Remove Window icon(PowerShell icon)
#6. Change: Reorganize UI
#7. Profile update(4240530): 1)Remove BIM from CNSNZ100; 2)Combine removepermission&CreateADGroup actions in the XML file; 3)Add tasks for HK and enable(unhide) for use;

#What's new in V3.4(2024/3/27):
#1. New: Give you a preview of the new project folder path
#2. New: Verify if the target project folder is existed before folder creation

#What's new in V3.3(2024/3/12):
#1. Update some wordings
#2. Profile Update, version 3240312: remove tasks of "remove permission of SDL2"

#What's new in V3.2(2023/7/6):
#1. Add PowerShell version and AD command status in the log file for troubleshoot purpose

#What's new in V3.13(2022/1/27):
#1. BUG fixed: project path invalid if the project folder contains project name
#2. Use SetAccessControl instead of set-acl to set permission, try to solve the weird case that set-acl not able to grant permission to some folders.

#What's new in V3.12(2022/1/11):
#1. In this version, there is a new Action----"SetFolderOwner"
#2. Add shell version and profile version in the report

#What's new in V3(2021/11/15):
#1. Not use NTFSSecurity module anymore, use pure Powershell
#2. 'CreateFolder', 'GrantFolderPermission', 'CreateADGroup' now support define multiple targets in one line by adding a separator "|"
#3. Add parameter 'Inheritance' & 'PermissionType' to 'GrantFolderPermission' action so that you can control permission apply to this folder only or subfolders, and grant "Allow" or "Deny" permission.
#4. Create "Report" folder automatically if the folder is not exist
#5. Show success and fail task count in report and prompt window after tasks completed
#one more thing......bug fixed!!!
###

$ShellVersion="4.0.0"
$ProfileMinVer="4240530"
$FormTitle="Project Folder Manager Pro Plus Ultimate - V" + $ShellVersion + " - Kin.Yan@wsp.com"
$PowershellPath = "C:\Support\ProjectFolderManager"
$ProfilePath = Join-Path $PowershellPath "ProjectFolders.xml"
#Import-Module Microsoft.PowerShell.Security

if((Get-Location).Path -ne $PowershellPath)
{
    [System.Windows.Forms.Messagebox]::Show("Please run this shell in below path:`n C:\Support\ProjectFolderManager")
    exit
}
if((Test-Path $ProfilePath) -eq $True)
{
    $global:xmldata = New-Object -TypeName XML
    $global:xmldata.Load($ProfilePath)

    $ProfileVersion=$xmldata.Offices.Attributes["ProfileVersion"].Value
    if($ProfileVersion -lt $ProfileMinVer)
    {
        [System.Windows.Forms.Messagebox]::Show("Profile outdate, please download the latest profile(xml)")
        exit
    }
}
else {
    write-host ("XML Profile not found![" + $ProfilePath + "]")
    Exit
}

$global:exitcode = 1
$global:RootFolder=""
#$SuccessTaskCount = 0
#$FailTaskCount = 0
# success=1, fail=9

########################################Action Start##########################################
function CreateADGroup
{
    Param(  [string] $GroupName,
	        [string] $GroupScope,
            [string] $Description,
            [string] $OUPath,
            [int] $WaitSeconds = 15)
    
    #GroupScope: Universal, Global, Domain Local
    $Groups=$GroupName.Split("|")
    foreach($Group in $Groups)
    {

        $ResultDescription=" - Create AD group: <AD Group>: " + $Group
        try {
            # Check if group already exists
            $existingGroup = Get-ADGroup -Filter {Name -eq $Group} -ErrorAction SilentlyContinue
            if($null -ne $existingGroup)
            {
                AddResultItem -ResultStatus "Success" -ResultDescription $ResultDescription + " (Group already exists - skipped)"
                # Continue to next step without waiting
                continue
            }
            
            # Group does not exist, create it
            New-ADGroup -name $Group -GroupScope $GroupScope -Description $Description -path $OUPath -ErrorAction Stop
            # Wait for AD replication before continuing
            Start-Sleep -Seconds $WaitSeconds
            
            # Verify group was created
            $verifyCount = 0
            $groupVerified = $false
            while($verifyCount -lt 5 -and -not $groupVerified)
            {
                $checkGroup = Get-ADGroup -Filter {Name -eq $Group} -ErrorAction SilentlyContinue
                if($null -ne $checkGroup)
                {
                    $groupVerified = $true
                    break
                }
                Start-Sleep -Seconds 2
                $verifyCount++
            }
            
            if($groupVerified)
            {
                AddResultItem -ResultStatus "Success" -ResultDescription $ResultDescription
            }
            else
            {
                AddResultItem -ResultStatus "Fail" -ResultDescription $ResultDescription + " (Verification failed)"
            }
        }
        catch
        {
            AddResultItem -ResultStatus "Fail" -ResultDescription $ResultDescription + " (Exception: " + $_.Exception.Message + ")"
        }

    }
}

function CreateFolder
{
    Param([string] $FolderPath,
          [string] $FolderName
    )

    $Folders=$FolderName.Split("|")
    foreach($folder in $Folders)
    {
        $fullpath=Join-Path $FolderPath $folder
        $ResultDescription=" - Create folder: <Folder path>:" + $fullpath
        try {
            # Check if folder already exists
            if((Test-Path $fullpath) -eq $True)
            {
                AddResultItem -ResultStatus "Success" -ResultDescription $ResultDescription + " (Folder already exists)"
            }
            else
            {
                New-Item -Path $fullpath -ItemType Directory -ErrorAction Stop
                # Verify folder was created
                if((Test-Path $fullpath) -eq $True)
                {
                    AddResultItem -ResultStatus "Success" -ResultDescription $ResultDescription
                }
                else
                {
                    AddResultItem -ResultStatus "Fail" -ResultDescription $ResultDescription + " (Verification failed)"
                }
            }
            }
        catch
        {
            AddResultItem -ResultStatus "Fail" -ResultDescription $ResultDescription + " (Exception: " + $_.Exception.Message + ")"
        }

    }    
}

function AddADGroupMember
{
    Param([string] $ADGroupName,
          [string] $Member
    )

    $ResultDescription=" - Add AD group member: <Member>: " + $Member + "; <ADGroup>: " + $ADGroupName
    try {
        Get-ADGroup $ADGroupName | Add-ADGroupMember -Members $Member -ErrorAction Stop
        AddResultItem -ResultStatus "Success" -ResultDescription $ResultDescription
    }
    catch
    {
        AddResultItem -ResultStatus "Fail" -ResultDescription $ResultDescription
    }
}

function GrantFolderPermission
{
    Param([string] $FolderPath,
          [string] $UserOrGroups,
          [string] $Permission,
          [string] $Inheritance,
          [string] $IsNew,
          [string] $PermissionType
    )
    #Permission set in PowerShell: ReadAndExecute, FullControl, Modify, Write, Read, CreateFiles, CreateDirectories, Delete
    #Reference: https://docs.microsoft.com/en-us/dotnet/api/system.security.accesscontrol.filesystemrights?view=windowsdesktop-5.0
    #$Inheritance: ThisFolderOnly, ThisAndSubFolders, ThisAndSubFoldersAndFiles
    #$PermissionType: Allow, Deny
    $Identities=$UserOrGroups.Split("|")
    foreach($Identity in $Identities)
    {
        $ResultDescription=" - Grant folder permission: <Folder>: " + $FolderPath + "; <Identity>: " + $Identity + "; <Permission>: " + $Permission + "; <Inheritance>: " + $Inheritance + "; <Type>: " + $PermissionType

        if((Test-Path $FolderPath) -eq $True)
        {
            try {
                # Verify identity exists in AD (with retries for newly created groups)
                $identityExists = $false
                $retryCount = 0
                $maxRetries = 5
                
                while($retryCount -lt $maxRetries -and -not $identityExists)
                {
                    try {
                        $checkIdentity = Get-ADGroup -Filter {Name -eq $Identity} -ErrorAction Stop
                        if($null -ne $checkIdentity)
                        {
                            $identityExists = $true
                            break
                        }
                    }
                    catch {
                        # Try alternative method if first fails
                        try {
                            $checkIdentity = Get-ADGroup $Identity -ErrorAction Stop
                            $identityExists = $true
                            break
                        }
                        catch {
                            $retryCount++
                            if($retryCount -lt $maxRetries)
                            {
                                Start-Sleep -Seconds 2
                            }
                        }
                    }
                }
                
                if(-not $identityExists)
                {
                    AddResultItem -ResultStatus "Fail" -ResultDescription $ResultDescription + " (Identity not found in AD)"
                    continue
                }
                
                $CurrentAcl = Get-Acl $FolderPath
                if($Inheritance -eq "ThisFolderOnly")
                {
                    $InheritanceFlag='None'
                    $PropagationFlag='InheritOnly'
                }
                elseif($Inheritance -eq "ThisAndSubFolders")
                {
                    $InheritanceFlag='ContainerInherit'
                    $PropagationFlag='NoPropagateInherit'
                }
                elseif($Inheritance -eq "ThisAndSubFoldersAndFiles")
                {   #'NoPropagateInherit'
                    $InheritanceFlag='ContainerInherit, ObjectInherit'
                    $PropagationFlag='None'
                }

                $NewAcl = New-Object -TypeName System.Security.AccessControl.FileSystemAccessRule($Identity,$Permission,$InheritanceFlag,$PropagationFlag,$PermissionType)
                if($IsNew -eq "Yes")
                {
                    $CurrentAcl.SetAccessRule($NewAcl)
                }
                else { #Same ID different permission
                    $CurrentAcl.AddAccessRule($NewAcl)
                }
                #(Get-Item $FolderPath).SetAccessControl($CurrentAcl)
                Set-Acl -path $FolderPath -AclObject $CurrentAcl -ErrorAction Stop
                
                # Verify permission was actually applied
                Start-Sleep -Milliseconds 500
                $verifyAcl = Get-Acl $FolderPath
                $permissionVerified = $false
                
                foreach ($access in $verifyAcl.Access)
                {
                    $accessvalue=$access.IdentityReference.Value.split("\") | Select-Object -Last 1
                    if($Identity -eq $accessvalue -and $access.FileSystemRights -like "*$Permission*")
                    {
                        $permissionVerified = $true
                        break
                    }
                }
                
                if($permissionVerified)
                {
                    AddResultItem -ResultStatus "Success" -ResultDescription $ResultDescription
                }
                else
                {
                    AddResultItem -ResultStatus "Fail" -ResultDescription $ResultDescription + " (Verification failed - permission not applied)"
                }
            }
            catch
            {
                AddResultItem -ResultStatus "Fail" -ResultDescription $ResultDescription + " (Exception: " + $_.Exception.Message + ")"
            }
        }
        else
        {
            AddResultItem -ResultStatus "Fail" -ResultDescription $ResultDescription + " (Folder not found)"
        }

    }
}

function RemoveFolderPermission
{
    Param([string] $FolderPath,
          [string] $UserOrGroups
    )


    $Identities=$UserOrGroups.Split("|")
    foreach($Identity in $Identities)
    {
        $ResultDescription=" - Remove folder permission: <Folder>: " + $FolderPath + "; <Identity>: " + $Identity
        if((Test-Path $FolderPath) -eq $True)
        {
            try {
                    $CurrentAcl = Get-Acl -Path $FolderPath
                    $permissionFound = $false
                    
                    foreach ($access in $CurrentAcl.Access)
                    {
                        $accessvalue=$access.IdentityReference.Value.split("\")
                        #$Identity2=$Identity.Split("\")
                        if($Identity -eq $accessvalue[$accessvalue.length-1])
                        {
                            $permissionFound = $true
                            if($access.IsInherited -eq $True)
                            {
                                #DisablePermissionInherit -FolderPath $FolderPath
                                $CurrentAcl.SetAccessRuleProtection($true, $true)
                                Set-Acl -path $FolderPath -AclObject $CurrentAcl
                            }
                            $CurrentAcl = $null
                            #RemoveNTFSPermission -FolderPath $FolderPath -Identity $Identity
                            $CurrentAcl = Get-Acl -Path $FolderPath
                            $AccessRule=New-Object -TypeName System.Security.AccessControl.FileSystemAccessRule($Identity,'FullControl','ContainerInherit, ObjectInherit','None','Allow')
                            $CurrentAcl.RemoveAccessRule($AccessRule)
                            Set-Acl -path $FolderPath -AclObject $CurrentAcl
                            $CurrentAcl = $null

                            $CurrentAcl = Get-Acl -Path $FolderPath
                            #$CurrentAcl3 = Get-Acl -Path $FolderPath
                            $SID=New-Object System.Security.Principal.NTAccount($Identity)
                            $CurrentAcl.PurgeAccessRules($SID)
                            Set-Acl -path $FolderPath -AclObject $CurrentAcl
                        }
                    }
                    
                    # Verify permission was actually removed
                    Start-Sleep -Milliseconds 500
                    $verifyAcl = Get-Acl -Path $FolderPath
                    $stillExists = $false
                    
                    foreach ($access in $verifyAcl.Access)
                    {
                        $accessvalue=$access.IdentityReference.Value.split("\")
                        if($Identity -eq $accessvalue[$accessvalue.length-1])
                        {
                            $stillExists = $true
                            break
                        }
                    }
                    
                    if(-not $permissionFound)
                    {
                        AddResultItem -ResultStatus "Success" -ResultDescription $ResultDescription + " (Permission not found in ACL)"
                    }
                    elseif($stillExists)
                    {
                        AddResultItem -ResultStatus "Fail" -ResultDescription $ResultDescription + " (Verification failed - permission still exists)"
                    }
                    else
                    {
                        AddResultItem -ResultStatus "Success" -ResultDescription $ResultDescription
                    }
                }
            catch{
                AddResultItem -ResultStatus "Fail" -ResultDescription $ResultDescription + " (Exception: " + $_.Exception.Message + ")"
                }
        }
        else
        {
            AddResultItem -ResultStatus "Fail" -ResultDescription $ResultDescription + " (Folder not found)"
        }
    }
}

function SetFolderOwner
{
    Param([string] $RootFolderPath,
          [string] $Folders,
          [string] $FolderOwner
    )
    #$folderowner=(get-acl $folder).Owner

    $SubFolders=$Folders.Split("|")

    foreach($SubFolder in $SubFolders)
    {
        $TargetFolder = Join-Path $RootFolderPath $SubFolder
        $ResultDescription=" - Set Folder Owner: <Folder>: " + $TargetFolder + "; <Owner>: " + $FolderOwner
        try {
            $acl = Get-Acl $TargetFolder
            $object = New-Object -TypeName System.Security.Principal.Ntaccount($FolderOwner)
            $acl.SetOwner($object)
            #$acl | set-acl $TargetFolder
            Set-Acl -Path $TargetFolder -AclObject $acl
            #(Get-Item $TargetFolder).SetAccessControl($acl)
            
            AddResultItem -ResultStatus "Success" -ResultDescription $ResultDescription
        }
        catch {
            AddResultItem -ResultStatus "Fail" -ResultDescription $ResultDescription
        }
    }    
}

function AddResultItem
{
    Param([string] $ResultStatus,
          [string] $ResultDescription
    )

    if($ResultStatus -eq "Success")
    {
        $Script:SuccessTaskCount=$Script:SuccessTaskCount+1
        $R="[Success] " + $ResultDescription
    }
    elseif($ResultStatus -eq "Fail") {
        $Script:FailTaskCount=$Script:FailTaskCount+1
        $R="[Fail] " + $ResultDescription
    }

    if($ResultList)
    {
        $ResultList.Items.Add($R)
    }
    else {
        write-host ($R)
    }  
}

function GenerateReport{
    Param([string] $ReportType
    )

    $today=Get-Date -Format "yyyyMMddHHmmss"
    $ReportPath= Join-Path $PowershellPath "Report"
    if((Test-Path $ReportPath) -ne $True)
    {
        New-Item -Path $ReportPath -ItemType Directory -ErrorAction Stop
    }
    $reportname=$ReportPath + "\PFMResult_" + $today + ".log"

    $R=""
    for($j=$ResultList.Items.Count-1;$j -gt -1; $j--)
    {
        $R=$ResultList.Items[$j] + "`r`n" + $R
    
    }
    $ADCommandStatus = CheckADCommand
    $ReportHeader="PFM Shell Version: " + $ShellVersion + "`r`n"
    $ReportHeader= $ReportHeader + "Profile Version: " + $ProfileVersion + "`r`n"
    $ReportHeader= $ReportHeader + "Host Name: " + $env:COMPUTERNAME + "`r`n"
    $ReportHeader= $ReportHeader + "PowerShell Version: " + $PSVersiontable.PSVersion + "`r`n"
    $ReportHeader= $ReportHeader + "AD Command Status: " + $ADCommandStatus + "`r`n"
    $ReportHeader= $ReportHeader + "Executed By: " + $env:UserName + "`r`n"
    $ReportHeader= $ReportHeader + "Office: " + $OfficeList.SelectedItem.ToString() + "`r`n" 
    $ReportHeader= $ReportHeader + "Team: " + $TeamList.SelectedItem.ToString() + "`r`n"
    $ReportHeader= $ReportHeader + "Date/Time: " + $today.Substring(0,4) + "/" + $today.Substring(4,2) + "/" + $today.Substring(6,2) + " " + $today.Substring(8,2) + ":"+ $today.Substring(10,2) + ":" + $today.Substring(12,2) + "`r`n"
    $ReportHeader= $ReportHeader + "Project No.: " + $ProjectNumber.Text + "`r`n"
    $ReportHeader= $ReportHeader + "Project Name: " + $ProjectName.Text + "`r`n"
    $ReportHeader= $ReportHeader + "Description: " + $ProjectDescription.Text + "`r`n"
    $ReportHeader= $ReportHeader + "Tasks(Success:" + $script:SuccessTaskCount + " ; Fail:" + $script:FailTaskCount + "): " + "`r`n"

    
    $R=$ReportHeader + $R + "*********Report End*********"
    #Write-Host "Selected File and Location:"  -ForegroundColor Green
    if($ReportType -eq "Clipboard")
    {
        $R | Set-Clipboard
        [System.Windows.Forms.Messagebox]::Show("Result copied to clipboard")
    }
    else {
        $R | Out-File -FilePath $reportname        
    }

}

function GetProjectFolderPathPreview
{
    Param([string] $aoffice,
          [string] $ateam
    )

    $node=getTeamNode -aoffice $aoffice -ateam $ateam
    $global:RootFolder=$node.Node.Attributes["RootFolder"].Value
    for($i=0;$i -lt $node.Node.ChildNodes.Count; $i++)
    {
        if($node.Node.ChildNodes[$i].Attributes["Type"].Value -eq 'CreateFolder')
        {
            $folderpath=$node.Node.ChildNodes[$i].Attributes["Path"].Value
            If($null -ne $RootFolder)
            {
                $folderpath=$folderpath.Replace("{RootFolder}",$RootFolder)
            }
            
            $foldername=$node.Node.ChildNodes[$i].InnerText
            $foldername=$foldername.Replace("{ProjectNumber}", $ProjectNumber.Text.trim())
            if($foldername.Contains("{ProjectName}") -eq $True)
            {
                if($ProjectName.Text.Trim() -eq "")
                {
                    $foldername=$foldername.Replace(" {ProjectName}", "")
                    $foldername=$foldername.Replace("-{ProjectName}", "")
                    $foldername=$foldername.Replace("{ProjectName}", "")
                }
                else {
                    $foldername=$foldername.Replace("{ProjectName}", $ProjectName.Text.Trim())
                }
            }

            $foldername=$foldername.Replace("{ProjectNumber}", $ProjectNumber.Text.trim())
            $foldername=$foldername.Replace("{ProjectName}", $ProjectName.Text.Trim())
            $fullpath = Join-Path -Path $folderpath.trim() $foldername.trim()
            $LbFullPathPreview.Text= $fullpath

            break
        }
    }
}

function PreValidate
{
    Param([string] $aoffice,
          [string] $ateam
    )

    $node=getTeamNode -aoffice $aoffice -ateam $ateam
    $global:RootFolder=$node.Node.Attributes["RootFolder"].Value



    $duplicatefolder= [System.Collections.Generic.List[string]]::new()
    $duplicategroup= [System.Collections.Generic.List[string]]::new()
    for($i=0;$i -lt $node.Node.ChildNodes.Count; $i++)
    {
        if($node.Node.ChildNodes[$i].Attributes["Type"].Value -eq 'CreateFolder')
        {
            $folderpath=$node.Node.ChildNodes[$i].Attributes["Path"].Value
            if($folderpath.contains("|"))
            {
                If($null -ne $RootFolder)
                {
                    $folderpath=$folderpath.Replace("{RootFolder}",$RootFolder)
                }
                
                $foldername=$node.Node.ChildNodes[$i].InnerText
                $foldername=$foldername.Replace("{ProjectNumber}", $ProjectNumber.Text.trim())
                if($foldername.Contains("{ProjectName}") -eq $True)
                {
                    if($ProjectName.Text.Trim() -eq "")
                    {
                        $foldername=$foldername.Replace(" {ProjectName}", "")
                        $foldername=$foldername.Replace("-{ProjectName}", "")
                        $foldername=$foldername.Replace("{ProjectName}", "")
                    }
                    else {
                        $foldername=$foldername.Replace("{ProjectName}", $ProjectName.Text.Trim())
                    }
                }

                $foldername=$foldername.Replace("{ProjectNumber}", $ProjectNumber.Text.trim())
                $foldername=$foldername.Replace("{ProjectName}", $ProjectName.Text.Trim())
                $fullpath = Join-Path -Path $folderpath.trim() $foldername.trim()

                if((Test-Path -path $fullpath) -eq $true)
                {
                    $duplicatefolder.Add($fullpath)
                }
            }
        }
        elseif($node.Node.ChildNodes[$i].Attributes["Type"].Value -eq 'CreateADGroup')
        {
            #<Action Type="CreateADGroup" OU="OU=Security,OU=Groups,OU=SG,OU=WSPObjects,DC=corp,DC=pbwan,DC=net" Scope="Universal" Description="">GRP-FSU-SGSIN100-MEP-{ProjectNumber}-Adm|GRP-FSU-SGSIN100-MEP-{ProjectNumber}-AU|GRP-FSU-SGSIN100-MEP-{ProjectNumber}-CFMsg|GRP-FSU-SGSIN100-MEP-{ProjectNumber}-CMMsg|GRP-FSU-SGSIN100-MEP-{ProjectNumber}-Doctif|GRP-FSU-SGSIN100-MEP-{ProjectNumber}-ENG|GRP-FSU-SGSIN100-MEP-{ProjectNumber}-PM</Action>
            $ADGroups=$node.Node.ChildNodes[$i].InnerText.Split("|")
            foreach($ADGroup in $ADGroups)
            {
                $ADGroup=$ADGroup.Replace("{ProjectNumber}", $ProjectNumber.Text.trim())
                $r=Get-ADGroup $ADGroup
                if(($null -eq $r) -eq $false)
                {
                    $duplicategroup.Add($ADGroup)
                }
            }
        }
    }
    if($duplicatefolder.Count -gt 0)
    {
        [System.Windows.Forms.Messagebox]::Show("Following target folder(s) already exist:`n" + $duplicatefolder + "`n`nPlease double check your input.")
        $validateOK=$false
    }
    if($duplicategroup.Count -gt 0)
    {
        [System.Windows.Forms.Messagebox]::Show("Following target group(s) already exist:`n" + $duplicategroup + "`n`nPlease double check your input.")
        $validateOK=$false
    }
    return $validateOK
}
##############################Action End################################################
function getTeamNode()
{
    Param(  [string] $aoffice,
            [string] $ateam
    )
    $xp="//Office[@Name='" + $aoffice + "']/Team[@Name='" + $ateam + "']"
    $node=$global:xmldata | Select-Xml -XPath $xp
    return $node
}

function CheckADCommand()
{
$ADCommands=@("New-ADGroup","Get-ADGroup")
$ADCommandResult=""
foreach($ADCommand in $ADCommands)
{
	try{
		$r=Get-Command -Name $ADCommand
		$ADCommandResult = $ADCommandResult + "[" + $ADCommand + ":Good]"
	}
	Catch{
		$ADCommandResult = $ADCommandResult + "[" + $ADCommand + ":NotGood]"
	}
}
return $ADCommandResult
}


function ExecuteActions()
{
    Param(  [string] $aoffice,
            [string] $ateam,
            [string] $aprojectnumber,
            [string] $aprojectname,
            [string] $aprojectdescription
    )

    $node=getTeamNode -aoffice $aoffice -ateam $ateam
    $global:RootFolder=$node.Node.Attributes["RootFolder"].Value
    for($i=0;$i -lt $node.Node.ChildNodes.Count; $i++)
    {
        Switch($node.Node.ChildNodes[$i].Attributes["Type"].Value)
        {
            'CreateADGroup' {
                                $adpath=$node.Node.ChildNodes[$i].Attributes["OU"].Value
                                $adgroupname=$node.Node.ChildNodes[$i].InnerText
				                $adgroupname=$adgroupname.Replace("{ProjectNumber}",$aprojectnumber)
                                $adgroupscope=$node.Node.ChildNodes[$i].Attributes["Scope"].Value
                                if($null -ne $node.Node.ChildNodes[$i].Attributes["Description"].Value)
                                {
                                    $aprojectdescription=$node.Node.ChildNodes[$i].Attributes["Description"].Value
                                }
                                CreateADGroup -OUPath $adpath -GroupScope $adgroupscope -GroupName $adgroupname.trim() -Description $aprojectdescription
                            }
            'CreateFolder' {
                                $folderpath=$node.Node.ChildNodes[$i].Attributes["Path"].Value
                                If($null -ne $RootFolder)
                                {
                                    $folderpath=$folderpath.Replace("{RootFolder}",$RootFolder)
                                }

                                $folderpath=$folderpath.Replace("{ProjectNumber}", $aprojectnumber)
                                if($folderpath.Contains("{ProjectName}") -eq $True)
                                {
                                    if($aprojectname -eq "")
                                    {
                                        $folderpath=$folderpath.Replace(" {ProjectName}", "")
                                        $folderpath=$folderpath.Replace("-{ProjectName}", "")
                                        $folderpath=$folderpath.Replace("{ProjectName}", "")
                                    }
                                    else {
                                        $folderpath=$folderpath.Replace("{ProjectName}", $aprojectname)
                                    }
                                }

                                $foldername=$node.Node.ChildNodes[$i].InnerText
				                $foldername=$foldername.Replace("{ProjectNumber}", $aprojectnumber)
				                $foldername=$foldername.Replace("{ProjectName}", $aprojectname)
                                CreateFolder -FolderPath $folderpath.trim() -FolderName $foldername.trim()
                            }
            'AddADGroupMember'{
                                $adgroupname=$node.Node.ChildNodes[$i].Attributes["Group"].Value
				                $adgroupname=$adgroupname.Replace("{ProjectNumber}",$aprojectnumber)
                                $adgroupmember=$node.Node.ChildNodes[$i].InnerText
				                $adgroupmember=$adgroupmember.Replace("{ProjectNumber}",$aprojectnumber)
                                AddADGroupMember -ADGroupName $adgroupname.trim() -Member $adgroupmember.trim()
                                }
            'GrantFolderPermission' {
                                $folderpath=$node.Node.ChildNodes[$i].Attributes["Path"].Value
                                If($null -ne $RootFolder)
                                {
                                    $folderpath=$folderpath.Replace("{RootFolder}",$RootFolder)
                                }
                                $folderpath=$folderpath.Replace("{ProjectNumber}", $aprojectnumber)
                                if($folderpath.Contains("{ProjectName}") -eq $True)
                                {
                                    if($aprojectname -eq "")
                                    {
                                        $folderpath=$folderpath.Replace(" {ProjectName}", "")
                                        $folderpath=$folderpath.Replace("-{ProjectName}", "")
                                        $folderpath=$folderpath.Replace("{ProjectName}", "")
                                    }
                                    else {
                                        $folderpath=$folderpath.Replace("{ProjectName}", $aprojectname)
                                    }
                                }
                                $permission=$node.Node.ChildNodes[$i].Attributes["Permission"].Value
                                $inheritance=$node.Node.ChildNodes[$i].Attributes["Inheritance"].Value
                                $permissiontype=$node.Node.ChildNodes[$i].Attributes["PermissionType"].Value
                                $isnew=$node.Node.ChildNodes[$i].Attributes["New"].Value
                                $userorgroup=$node.Node.ChildNodes[$i].InnerText
				                $userorgroup=$userorgroup.Replace("{ProjectNumber}",$aprojectnumber)
                                GrantFolderPermission -FolderPath $folderpath.trim() -UserOrGroups $userorgroup.trim() -Permission $permission -inheritance $inheritance -IsNew $isnew -PermissionType $permissiontype
                                }
            'RemoveFolderPermission'{
                                $folderpath=$node.Node.ChildNodes[$i].Attributes["Path"].Value
                                If($null -ne $RootFolder)
                                {
                                    $folderpath=$folderpath.Replace("{RootFolder}",$RootFolder)
                                }
                                $folderpath=$folderpath.Replace("{ProjectNumber}", $aprojectnumber)
                                if($folderpath.Contains("{ProjectName}") -eq $True)
                                {
                                    if($aprojectname -eq "")
                                    {
                                        $folderpath=$folderpath.Replace(" {ProjectName}", "")
                                        $folderpath=$folderpath.Replace("-{ProjectName}", "")
                                        $folderpath=$folderpath.Replace("{ProjectName}", "")
                                    }
                                    else {
                                        $folderpath=$folderpath.Replace("{ProjectName}", $aprojectname)
                                    }
                                }
                                $userorgroup=$node.Node.ChildNodes[$i].InnerText
				                $userorgroup=$userorgroup.Replace("{ProjectNumber}",$aprojectnumber)
                                #$permissiontype=$node.Node.ChildNodes[$i].Attributes["PermissionType"].Value
                                RemoveFolderPermission -FolderPath $folderpath.trim() -UserOrGroups $userorgroup.trim()
                                }
            'SetFolderOwner'{
                                $folderpath=$node.Node.ChildNodes[$i].Attributes["Path"].Value
                                $folderowner=$node.Node.ChildNodes[$i].Attributes["Owner"].Value
                                If($null -ne $RootFolder)
                                {
                                    $folderpath=$folderpath.Replace("{RootFolder}",$RootFolder)
                                }
                                $folderpath=$folderpath.Replace("{ProjectNumber}", $aprojectnumber)
                                if($folderpath.Contains("{ProjectName}") -eq $True)
                                {
                                    if($aprojectname -eq "")
                                    {
                                        $folderpath=$folderpath.Replace(" {ProjectName}", "")
                                        $folderpath=$folderpath.Replace("-{ProjectName}", "")
                                        $folderpath=$folderpath.Replace("{ProjectName}", "")
                                    }
                                    else {
                                        $folderpath=$folderpath.Replace("{ProjectName}", $aprojectname)
                                    }
                                }

                                $folders=$node.Node.ChildNodes[$i].InnerText
                                $folders=$folders.Replace("{ProjectNumber}", $aprojectnumber)
                                if($folders.Contains("{ProjectName}") -eq $True)
                                {
                                    if($aprojectname -eq "")
                                    {
                                        $folders=$folders.Replace(" {ProjectName}", "")
                                        $folders=$folders.Replace("-{ProjectName}", "")
                                        $folders=$folders.Replace("{ProjectName}", "")
                                    }
                                    else {
                                        $folders=$folders.Replace("{ProjectName}", $aprojectname)
                                    }
                                }
                                SetFolderOwner -RootFolderPath $folderpath -Folders $folders -FolderOwner $folderowner
                            }
        }
    }
    GenerateReport -ReportType "LogFile"

    [System.Windows.Forms.Messagebox]::Show("Task(s) completed!`nSuccess:" + $Script:SuccessTaskCount + " ; Fail:" + $Script:FailTaskCount)

}


## Command line mode or GUI mode
if(($NULL -ne $args[0]) -and ($NULL -ne $args[1]) -and ($NULL -ne $args[2]) -and ($NULL -ne $args[3]))
{
    if($null -ne $args[4])
    {
    $aprojectdescription=$args[4]
    }
    else {
        $aprojectdescription=""
    }
    ExecuteActions -aoffice $args[0] -ateam $args[1] -aprojectnumber $args[2] -aprojectname $args[3] -aprojectdescription $aprojectdescription
    Exit
}

function EnvironmentCheck
{
<#
$keywords='New-Object','Join-Path','Test-Path','New-Item2','New-ADGroup','Get-ADGroup','Add-ADGroupMember','Get-ADGroupMember','Add-NTFSAccess','Remove-NTFSAccess','Get-ACL','Add-Type','Select-Xml'
$r=$true
foreach($keyword in $keywords)
    {
        $r=$r -and (get-command -Name $keyword -ErrorAction Ignore)
    }
return $r
#>
    $OSversion=(get-ciminstance Win32_OperatingSystem).caption
    
    if($OSversion.Contains("Server") -eq $true)
    {
        $Module2=Get-WindowsFeature -Name RSAT-AD-PowerShell | where-object {$_.InstallState -eq "Installed"}
    }
    else
    {
        #$Module2=Get-Module -Name Activedirectory
        $Module2=get-command -name get-adgroup -ErrorAction Ignore
    }
    
    if($Module2.length -eq 0)
    {
        write-host ("Please install ActiveDirectory module before create folder. `nPS to install this module: Install-WindowsFeature RSAT-AD-PowerShell")
        [System.Windows.Forms.Messagebox]::Show("Please install ActiveDirectory module before create folder. `nPS to install this module: Install-WindowsFeature RSAT-AD-PowerShell")
        Exit
    }
}


function isFolderNameValid
{
    Param([string] $foldername
    )
    $r=!($foldername.Contains("\") -or $foldername.Contains("/") -or $foldername.Contains(":") -or $foldername.Contains("*") -or $foldername.Contains("?") -or $foldername.Contains("""") -or $foldername.Contains("<") -or $foldername.Contains(">") -or $foldername.Contains("|"))
    return $r
}


function RefreshTeam
{
    Param([string] $Office
    )
    $TeamList.Items.Clear()
    for($o=0;$o -le $xmldata.Offices.ChildNodes.Count-1;$o++)
    {
        if($xmldata.Offices.ChildNodes[$o].Name -eq $Office)
        {
            foreach($TeamName in $xmldata.Offices.ChildNodes[$o].Team)
            {
                if($TeamName.Attributes["Status"].Value -eq "Active")
                {
                    $TeamList.Items.Add($TeamName.Attributes["Name"].Value)
                }
            }
        }
    }
}


#### create GUI#######
Add-Type -AssemblyName System.Windows.Forms
$PowerShellForms = New-Object system.Windows.Forms.Form
$PowerShellForms.Text=$FormTitle
$PowerShellForms.Size = New-Object System.Drawing.Size(490,580)
$PowerShellForms.MinimizeBox = $False
$PowerShellForms.MaximizeBox = $False
$PowerShellForms.FormBorderStyle = [System.Windows.Forms.FormBorderStyle]::FixedDialog

$PowerShellForms.SizeGripStyle = "Hide"
#$Icons = [system.drawing.icon]::ExtractAssociatedIcon($PSHOME + "\powershell.exe")
#$PowerShellForms.Icon = $Icons
$PowerShellForms.StartPosition = "CenterScreen"
$PowerShellForms.Font = New-Object System.Drawing.Font("Tahoma",10,[System.Drawing.FontStyle]::Regular)
$PowerShellForms.Add_Load({EnvironmentCheck})


$Lboffice = New-Object System.Windows.Forms.Label
$Lboffice.Text = "Office"
$Lboffice.AutoSize = $True
$Lboffice.BackColor = "Transparent"
$Lboffice.Font = New-Object System.Drawing.Font("Tahoma",10,[System.Drawing.FontStyle]::Regular)
$Lboffice.ForeColor = "Red"
$Lboffice.Location = New-Object System.Drawing.Point(10,10)

$OfficeList = New-Object System.Windows.Forms.ListBox
$OfficeList.Font = New-Object System.Drawing.Font("Tahoma",9,[System.Drawing.FontStyle]::Regular)
$OfficeList.Location = New-Object System.Drawing.Point(10,35)
$OfficeList.Size = New-Object System.Drawing.Size(200,120)
$OfficeList.Height = 120
for($o=0;$o -le $xmldata.Offices.ChildNodes.Count-1;$o++)
{
    if($xmldata.Offices.ChildNodes[$o].Attributes["Status"].Value -eq "Active")
    {
        $OfficeList.items.Add($xmldata.Offices.ChildNodes[$o].Name)
    }
}
$OfficeList.TabIndex = 1
$OfficeList.Add_SelectedIndexChanged(
    {
        RefreshTeam -Office $OfficeList.SelectedItem.ToString()
    }
)

$LbTeam = New-Object System.Windows.Forms.Label
$LbTeam.Text = "Team"
$LbTeam.AutoSize = $True
$LbTeam.BackColor = "Transparent"
$LbTeam.Font = New-Object System.Drawing.Font("Tahoma",10,[System.Drawing.FontStyle]::Regular)
$LbTeam.ForeColor = "Red"
$LbTeam.Location = New-Object System.Drawing.Point(260,10)

$TeamList = New-Object System.Windows.Forms.ListBox
$TeamList.Font = New-Object System.Drawing.Font("Tahoma",10,[System.Drawing.FontStyle]::Regular)
$TeamList.Location = New-Object System.Drawing.Point(260,35)
$TeamList.Size = New-Object System.Drawing.Size(200,120)
$TeamList.Height=120
#Group name length limit is 64, group name example: 
$TeamList.TabIndex = 2


$LbProjectNumber = New-Object System.Windows.Forms.Label
$LbProjectNumber.Text = "Project No."
$LbProjectNumber.AutoSize = $True
$LbProjectNumber.BackColor = "Transparent"
$LbProjectNumber.Font = New-Object System.Drawing.Font("Tahoma",10,[System.Drawing.FontStyle]::Regular)
$LbProjectNumber.ForeColor = "Red"
$LbProjectNumber.Location = New-Object System.Drawing.Point(10,165)

$ProjectNumber = New-Object System.Windows.Forms.TextBox
$ProjectNumber.Font = New-Object System.Drawing.Font("Tahoma",10,[System.Drawing.FontStyle]::Regular)
$ProjectNumber.Location = New-Object System.Drawing.Point(10,190)
$ProjectNumber.Size = New-Object System.Drawing.Size(450,100)
#Group name length limit is 64, group name example: GRP-FSU-CNGGZ100-123456 or GRP-FSU-CNGGZ100-MEP-123456-M
$ProjectNumber.MaxLength = 15
$ProjectNumber.TabIndex = 3
$global:ProjectNumberStatus=$False
$global:ProjectNameStatus=$True
$ProjectNumber.add_TextChanged({
    #if($ProjectNumber.Text -match '[^-\w\.]')
    if($ProjectNumber.Text -match '[^-\s\w\.]')
    {
        [System.Windows.Forms.Messagebox]::Show("Project number only allow: A-Z a-z 0-9 -_")
        $global:ProjectNumberStatus=$False
        #$LbProjectNumber.ForeColor="Red"
    }
    else
    {
        $global:ProjectNumberStatus=$True
        #$LbProjectNumber.ForeColor="Black"
        #[System.Windows.Forms.Messagebox]::Show($OfficeList.SelectedIndex)
        if(($OfficeList.SelectedIndex -gt -1) -and ($TeamList.SelectedIndex -gt -1)) #selected
        {
            GetProjectFolderPathPreview -aoffice $OfficeList.SelectedItem -ateam $TeamList.SelectedItem
        }
    }
    $BtnOK.enabled=($global:ProjectNameStatus -and $global:ProjectNumberStatus)
})


$LbProjectName = New-Object System.Windows.Forms.Label
$LbProjectName.Text = "Project Name"
$LbProjectName.AutoSize = $True
$LbProjectName.BackColor = "Transparent"
$LbProjectName.Font = New-Object System.Drawing.Font("Tahoma",10,[System.Drawing.FontStyle]::Regular)
$LbProjectName.ForeColor = "Black"
$LbProjectName.Location = New-Object System.Drawing.Point(10,220)

$ProjectName = New-Object System.Windows.Forms.TextBox
$ProjectName.Font = New-Object System.Drawing.Font("Tahoma",10,[System.Drawing.FontStyle]::Regular)
$ProjectName.Location = New-Object System.Drawing.Point(10,245)
$ProjectName.Size = New-Object System.Drawing.Size(450,100)
$ProjectName.MaxLength = 32
#Project name length limit is 32
$ProjectName.TabIndex = 4
$ProjectName.add_TextChanged({
    if($ProjectName.Text -match '[\\/:*?"<>|]')
    {
	#[^-\w\. ]

        [System.Windows.Forms.Messagebox]::Show("These characters are not allowed: \ / : * ? "" < > |")
        $global:ProjectNameStatus=$False
        $LbProjectName.ForeColor="Red"
    }
    else
    {
        $global:ProjectNameStatus=$True
        $LbProjectName.ForeColor="Black"

        #[System.Windows.Forms.Messagebox]::Show($OfficeList.SelectedIndex)
        if(($OfficeList.SelectedIndex -gt -1) -and ($TeamList.SelectedIndex -gt -1)) #selected
        {
            GetProjectFolderPathPreview -aoffice $OfficeList.SelectedItem -ateam $TeamList.SelectedItem
        }        
    }
    $BtnOK.enabled=($global:ProjectNameStatus -and $global:ProjectNumberStatus)
})

$LbFullPathPreview = New-Object System.Windows.Forms.Label
$LbFullPathPreview.Text = "(Full Path Preview)"
#$LbFullPathPreview.AutoSize = $True
$LbFullPathPreview.BackColor = "Transparent"
$LbFullPathPreview.Font = New-Object System.Drawing.Font("Calibri",8,[System.Drawing.FontStyle]::Italic)
$LbFullPathPreview.ForeColor = "Black"
$LbFullPathPreview.Location = New-Object System.Drawing.Point(100,220)
$LbFullPathPreview.Height = 25
$LbFullPathPreview.Width = 380


$LbProjectDescription = New-Object System.Windows.Forms.Label
$LbProjectDescription.Text = "Description - RITM"
$LbProjectDescription.AutoSize = $True
$LbProjectDescription.BackColor = "Transparent"
$LbProjectDescription.Font = New-Object System.Drawing.Font("Tahoma",10,[System.Drawing.FontStyle]::Regular)
$LbProjectDescription.ForeColor = "Red"
$LbProjectDescription.Location = New-Object System.Drawing.Point(10,275)

$ProjectDescription = New-Object System.Windows.Forms.TextBox
$ProjectDescription.Font = New-Object System.Drawing.Font("Tahoma",10,[System.Drawing.FontStyle]::Regular)
$ProjectDescription.Location = New-Object System.Drawing.Point(10,300)
$ProjectDescription.Size = New-Object System.Drawing.Size(450,100)
$ProjectDescription.TabIndex = 5


$LbResultList = New-Object System.Windows.Forms.Label
$LbResultList.Text = "Result"
$LbResultList.AutoSize = $True
$LbResultList.BackColor = "Transparent"
$LbResultList.Font = New-Object System.Drawing.Font("Tahoma",10,[System.Drawing.FontStyle]::Regular)
$LbResultList.ForeColor = "Black"
$LbResultList.Location = New-Object System.Drawing.Point(10,330)

$ResultList = New-Object System.Windows.Forms.ListBox
$ResultList.Font = New-Object System.Drawing.Font("Tahoma",10,[System.Drawing.FontStyle]::Regular)
$ResultList.Location = New-Object System.Drawing.Point(10,355)
$ResultList.Size = New-Object System.Drawing.Size(450,100)
$ResultList.HorizontalScrollbar = $True
$ResultList.BackColor = [System.Drawing.Color]::Gray


$BtnResult2Clipboard = New-Object System.Windows.Forms.Button
$BtnResult2Clipboard.Size = New-Object System.Drawing.Size(30,30)
$BtnResult2Clipboard.Location = New-Object System.Drawing.Point(10,460)
$BtnResult2Clipboard.Text = "C"
$tt1=New-Object System.Windows.Forms.ToolTip
$tt1.SetToolTip($BtnResult2Clipboard,"Copy result to clipboard")
$BtnResult2Clipboard.Font = New-Object System.Drawing.Font("Tahoma",10,[System.Drawing.FontStyle]::Regular)
$BtnResult2Clipboard.Add_Click(
    {
        GenerateReport -ReportType "Clipboard"
    }
)

$BtnResult2File = New-Object System.Windows.Forms.Button
$BtnResult2File.Size = New-Object System.Drawing.Size(30,30)
$BtnResult2File.Location = New-Object System.Drawing.Point(45,460)
$BtnResult2File.Text = "E"
$tt2=New-Object System.Windows.Forms.ToolTip
$tt2.SetToolTip($BtnResult2File,"Export result to log file")
$BtnResult2File.Font = New-Object System.Drawing.Font("Tahoma",10,[System.Drawing.FontStyle]::Regular)
$BtnResult2File.Add_Click(
    {
        GenerateReport -ReportType "LogFile"
    }
)

$BtnOK = New-Object System.Windows.Forms.Button
$btnOK.Size = New-Object System.Drawing.Size(100,40)
$BtnOK.Location = New-Object System.Drawing.Point(355,465)
$BtnOK.Text = "Create"
$BtnOK.Font = New-Object System.Drawing.Font("Tahoma",10,[System.Drawing.FontStyle]::Regular)
$BtnOK.Add_Click(
    {
        $R=PreValidate -aoffice $OfficeList.SelectedItem -ateam $TeamList.SelectedItem
        if($R -eq $null)
        {
            $pathexist = Test-Path $LbFullPathPreview.Text
            if($pathexist -eq $true)
            {
                [System.Windows.Forms.Messagebox]::Show("Target project folder already exist: `n" + $LbFullPathPreview.Text + "`n Please validate project no and name.")
            }
            else {
                $ResultList.Items.Clear()
                $script:SuccessTaskCount=0;
                $script:FailTaskCount=0;
                
                $txtprojectnumber=$ProjectNumber.Text.Trim()
                $txtprojectname=$ProjectName.Text.Trim()
                if(($null -ne $OfficeList.SelectedItem) -and ($null -ne $TeamList.SelectedItem) -and ($txtprojectnumber.Length -ne 0) -and ($ProjectDescription.Text.Trim() -ne ""))
                {
                    $r=(isFolderNameValid -foldername $txtprojectnumber) -and (isFolderNameValid -foldername $txtprojectname)
                    if($r -eq $true)
                    {
                        ExecuteActions -aoffice $OfficeList.SelectedItem -ateam $TeamList.SelectedItem -aprojectnumber $txtprojectnumber -aprojectname $txtprojectname -aprojectdescription $ProjectDescription.Text.Trim()
                    }
                    else
                    {
                        [System.Windows.Forms.Messagebox]::Show("Project number or name invalid")
                    }
                }
                else
                {
                    [System.Windows.Forms.Messagebox]::Show("Following items are mandatory: `nOffice`nTeam`nProject No.`nDescription")
                }
            }
        }
    }
)

$PFMLinkLabel = New-Object System.Windows.Forms.LinkLabel
$PFMLinkLabel.Size = New-Object System.Drawing.Size(300,40)
$PFMLinkLabel.Location = New-Object System.Drawing.Point(10,495)
$PFMLinkLabel.Text = "Get latest version@Asia OSS SharePoint"
$PFMLinkLabel.Font = New-Object System.Drawing.Font("Tahoma",10,[System.Drawing.FontStyle]::Regular)
$PFMLinkLabel.Add_LinkClicked(
    {
        $PFMURL =  "https://wsponline.sharepoint.com/:f:/r/sites/SG-ASIAOSSTEAM/Shared%20Documents/General/Operation%20Task/Tools/Project%20Folder/ProjectFolderManager/Latest%20Version?csf=1&web=1&e=hGcuEJ"
        Start-Process $PFMURL
    }
)

$TabControl1 = New-Object System.Windows.Forms.TabControl
$tabPage1=New-Object System.Windows.Forms.TabPage
$tabPage1.text = "Create Project"
$tabPage1.TabIndex=0
$tabPage1.Controls.AddRange(($Lboffice, $OfficeList, $LbTeam, $TeamList, $LbProjectNumber, $ProjectNumber, $LbProjectName, $ProjectName, $LbFullPathPreview, $LbProjectDescription, $ProjectDescription, $LbResultList,$ResultList,$BtnResult2Clipboard,$BtnResult2File, $PFMLinkLabel,$BtnOK))

$TabControl1.Controls.Add($tabPage1)
$TabControl1.Location = New-Object System.Drawing.Point(0,0)
$TabControl1.Size=$PowerShellForms.Size
$Tabcontrol1.Add_SelectedIndexChanged(
    {
        $TabControl1.TabPages[$TabControl1.SelectedIndex].Controls.AddRange(($Lboffice, $OfficeList, $LbTeam, $TeamList))
    }
)

$PowerShellForms.Controls.Add($TabControl1)

$PowerShellForms.ShowDialog()