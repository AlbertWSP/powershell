#Region ShellInfo
<#
Author: kin.yan@wsp.com
Description: This shell is designed for user AD account creation and finalize configuration for WSP Asia region only. Given that it's lighter, faster, and more powerful than AD Manager, I give it a name ---> AD Director :)
---------------------------------------------------------------------------------
Requirement:
1) PowerShell 5.x/7.x;
2) RSAT-AD-PowerShell feature installed;
3) Run this shell with elevated account (IT5)

Initial Setup:
1. Enable PS: Set-ExecutionPolicy RemoteSigned -Scope CurrentUser
2. Install RSAT-AD-PowerShell feature: Get-WindowsCapability -Name RSAT* -Online | Add-WindowsCapability -Online
---------------------------------------------------------------------------------
Version History
Version 0.2 (2024/09/06)
New:
1. In case duplicate default groups are preset in Teams.xml file, only unique group will be added to the member group list
2. Validate if target groups are in AD or not before adding to member group list
3. Support multiple target groups seperate by ";" to be added to group list
4. Add user Title and Tel(Home) filed. Tel(Home) is specificly for TW only.
5. Group name is now support both email and samaccountname 
6. The AD property value supplimented or changed by this shell will be highlighted in RED in the textbox
Bug fix: 
1. fix an issue failed to add groups into group list box when default group of the site is blank
2. fix a bug that user setting not updated when some fields(logon script/department) are blank

Version 0.1 (2024/08/16)
1. Initial release for pilot testing

---------------------------------------------------------------------------------
#>
#EndRegion ShellInfo

#Region Variable
$ADDirectorVersion = "0.2"
$FormTitle = "AD Director V" + $ADDirectorVersion + " - Kin.Yan@wsp.com"

$ADDirectorPath = "C:\support\ADDirector"
$ProfilePath = Join-Path $ADDirectorPath "Team.xml"

$global:xmldata = New-Object -TypeName XML
$global:xmldata.Load($ProfilePath)

$global:SiteCodeList=@()
#$global:CorporateDefaultGroup=@()
$CorporateDefaultGroups = New-Object -TypeName 'System.Collections.ArrayList'
$ProfileVersion=$xmldata.Sites.Attributes["Version"].Value
$ProfileLatestUpdate=$xmldata.Sites.Attributes["LastestUpdate"].Value

#EndRegion Variable



#Region Function
function PopulateBasicConfig
{
    #$Fn=$Page1_TBJoinerFirstName.Text.ToUpper().Substring(0,1) + $Page1_TBJoinerFirstName.Text.ToLower().Substring(1,$Page1_TBJoinerFirstName.Text.Length-1)
    if($Page1_RBFinalizeADAccount.Checked -eq $false)
    {
        $Fn1=$Page1_TBJoinerFirstName.Text
        $Fns=$Fn1.Split('-')
        foreach($Fnmember in $Fns)
        {
            if($Fn -eq $null)
            {
                $Fn=$Fnmember.ToUpper().Substring(0,1) + $Fnmember.ToLower().Substring(1,$Fnmember.Length-1)
            }
            else {
                $Fn=$Fn + "-" + $Fnmember.ToUpper().Substring(0,1) + $Fnmember.ToLower().Substring(1,$Fnmember.Length-1)
            }
        }
        $Page1_TBJoinerFirstName.Text=$Fn
        $Ln=$Page1_TBJoinerLastName.Text.ToUpper().Substring(0,1) + $Page1_TBJoinerLastName.Text.ToLower().Substring(1,$Page1_TBJoinerLastName.Text.Length-1)
        $Page1_TBJoinerLastName.Text=$Ln
        $Page1_TBJoinerDisplayName.Text = $Ln + ", " + $Fn
        $Page1_TBJoinerSamAccountName.Text = $Page1_OfficeList.SelectedItem.Substring(0,2) + $Fn.Substring(0,1) + $Ln.Substring(0,1) + $Page1_TBJoinerEmployeeID.Text.ToUpper()
        $Page1_TBJoinerFullName.Text = $Ln + ", " + $Fn + " (" + $Page1_TBJoinerSamAccountName.Text + ")"
        $Page1_TBEmail.Text = $Fn + "." + $Ln +"@wsp.com"
        
    }
}

function JoinerSamAccountExist
{
    Param(  [string] $SamAccountName
            )    
    $pscmd="Get-aduser -identity '" + $SamAccountName + "'"
    $R=@(Invoke-Expression -Command $pscmd).count
    if($R -eq 0) #AD account not found
    {
        return $false
    }
    else {
            return $true
    }
}

function JoinerEmailExist
{
    Param(  [string] $EmailAddress
            )
    $pscmd="Get-aduser -identity '" + $SamAccountName + "'"
    $R=@(Invoke-Expression -Command $pscmd).count
    if($R -eq 0) #AD account not found
    {
        return $false
    }
    else {
            return $true
    }
}

function ProposeNewEmail
{
    $CurrentEmail = $Page1_TBEmail.Text
    #$p="mail -eq '" + $CurrentEmail + "'"
    #$R=@(get-aduser -filter $p).count

    $EmailPrefix = $CurrentEmail.Substring(0,$CurrentEmail.indexof("@"))
    $NewMailaddress=$EmailPrefix
    $i=1
    $pscmd="get-aduser -filter {mail -eq '" + $Page1_TBEmail.Text +"'}"
    $usercount=@(Invoke-Expression -Command $pscmd).count
    while ($usercount -eq 1) {
        $NewMailaddress = $EmailPrefix + $i.tostring()
        $Page1_TBEmail.text = $NewMailaddress +"@wsp.com"
        $i=$i+1
        $pscmd="get-aduser -filter {mail -eq '" + $NewMailaddress +"@wsp.com'}"
        $usercount=@(Invoke-Expression -Command $pscmd).count
    }

    return $NewMailaddress +"@wsp.com"
}

function GenerateInitialPassword
{
    $P1=-Join("AEFHMNQRTY".tochararray() | Get-Random -Count 1)
    $P2=-Join("aefhmnrty".tochararray() | Get-Random -Count 1)
    $P3=-Join("2345678".tochararray() | Get-Random -Count 1)
    $P4=-Join("@#$%&*()+<>?".tochararray() | Get-Random -Count 1)
    $P5=-Join("AEFHMNQRTYaefhmnrty2345678@#$%&*()+<>?".tochararray() | Get-Random -Count 4)
    $PW=$P1+$P2+$P3+$P4+$P5
    return $PW
}


function InitialDataColect
{
    foreach($site in $xmldata.Sites.ChildNodes)
    {
        $global:SiteCodeList +=$site.Code
    }

    $DefaultGroups=$xmldata.Sites.Attributes["CorporateDefaultGroup"].Value.Split('|')
    if($DefaultGroups -ne "")
    {
        foreach($DefaultGroup in $DefaultGroups)
        {
            #$global:CorporateDefaultGroup +=$DefaultGroup
            $CorporateDefaultGroups.Add($DefaultGroup)
        }
    }
}
InitialDataColect

function GenerateLog{
    Param(  [string] $ActionType, #NewADAccount -or- FinalizeADAccount
            [string] $ADUser,
            [string] $GroupMember
            )   

    $today=Get-Date -Format "yyyy/MM/dd HH:mm:ss"
    $LogPath= Join-Path $ADDirectorPath "Log"
    if((Test-Path $LogPath) -ne $True)
    {
        New-Item -Path $LogPath -ItemType Directory -ErrorAction Stop
    }
    $LogFileName=$LogPath + "\ADDirector_" + $Page1_TBJoinerSamAccountName.Text + "_" + (Get-Date -Format "yyyyMMddHHmmss").ToString() + ".log"
    $LogItems= $LogItems + "******************Report Start******************`r`n"
    $LogItems="AD Director Version: V" + $ADDirectorVersion + "`r`n"
    $LogItems= $LogItems + "XML Profile Version: " + $ProfileVersion + "`r`n"
    $LogItems= $LogItems + "Profile Latest Update: " + $ProfileLatestUpdate + "`r`n"
    $LogItems= $LogItems + "PowerShell Version: " + $PSVersiontable.PSVersion + "`r`n"
    $LogItems= $LogItems + "Host Name: " + $env:COMPUTERNAME + "`r`n"
    $LogItems= $LogItems + "Executed By: " + $env:UserName + "`r`n"
    $LogItems= $LogItems + "Date Time: " + $today + "`r`n"
    $LogItems= $LogItems + "Join Ticket: " + $Page1_TBJoinerTicket.Text + "`r`n" 
    $LogItems= $LogItems + "Action Type: " + $ActionType + "`r`n"
    $LogItems= $LogItems + "AD User Task: " + $ADUser + "`r`n"
    $LogItems= $LogItems + "Group Task: " + $GroupMember + "`r`n"
    $LogItems= $LogItems + "************Account Info************`r`n"
    if($ActionType -eq "NewADAccount")
    {
        $LogItems= $LogItems + "Initial Password: " + $Page1_TBInitialPassword.Text + "`r`n"
    }
    $LogItems= $LogItems + "First Name: " + $Page1_TBJoinerFirstName.Text + "`r`n"
    $LogItems= $LogItems + "Last Name: " + $Page1_TBJoinerLastName.Text + "`r`n"
    $LogItems= $LogItems + "Full Name: " + $Page1_TBJoinerFullName.Text + "`r`n" 
    $LogItems= $LogItems + "Display Name: " + $Page1_TBJoinerDisplayName.Text + "`r`n"
    $LogItems= $LogItems + "SamAccountName: " + $Page1_TBJoinerSamAccountName.Text + "`r`n"
    $LogItems= $LogItems + "Office: " + $Page1_OfficeList.SelectedItem + "`r`n" 
    $LogItems= $LogItems + "Department: " + $Page1_DepartmentList.SelectedItem + "`r`n"
    $LogItems= $LogItems + "Email(UPN): " + $Page1_TBEmail.Text + "`r`n"
    $LogItems= $LogItems + "EA1: " + $Page1_TBJoinerEmployeeID.Text + "`r`n"
    $LogItems= $LogItems + "EA9: " + $Page1_TBJoinerEA9.Text + "`r`n"
    $LogItems= $LogItems + "EA11: " + $Page1_TBJoinerEA11.Text + "`r`n"
    $LogItems= $LogItems + "Logon Script: " + $Page1_TBJoinerLogonScript.Text + "`r`n"
    $LogItems= $LogItems + "EmployeeType: " + $Page1_TBJoinerEmployeeType.Text + "`r`n"
    $LogItems= $LogItems + "OU: " + $Page1_TBJoinerOU.Text + "`r`n"
    foreach($ADgroup in $Page1_DefaultGroupList.Items)
    {
        $LogItems= $LogItems + "Member Groups: " + $ADgroup + "`r`n"
    }
    $LogItems= $LogItems + "******************Report End******************"
    #Write-Host "Selected File and Location:"  -ForegroundColor Green
    $LogItems | Out-File -FilePath $LogFileName

}
#EndRegion Function

#Region GUI
###################### Create GUI############################
Add-Type -AssemblyName System.Windows.Forms
$PowerShellForms = New-Object system.Windows.Forms.Form
$PowerShellForms.Text=$FormTitle
$PowerShellForms.Size = New-Object System.Drawing.Size(640,560)
#$PowerShellForms.MinimizeBox = $False
$PowerShellForms.MaximizeBox = $False
$PowerShellForms.FormBorderStyle = [System.Windows.Forms.FormBorderStyle]::FixedDialog
$PowerShellForms.SizeGripStyle = "Hide"
$Icons = [system.drawing.icon]::ExtractAssociatedIcon($PSHOME + "\powershell.exe")
$PowerShellForms.Icon = $Icons
$PowerShellForms.StartPosition = "CenterScreen"
$PowerShellForms.Font = New-Object System.Drawing.Font("Tahoma",10,[System.Drawing.FontStyle]::Regular)


$Page1_GBObjectType = New-Object System.Windows.Forms.GroupBox
$Page1_GBObjectType.Location = New-Object System.Drawing.Point(10,5)
$Page1_GBObjectType.Size = New-Object System.Drawing.Size(310,30)
$Page1_GBObjectType.Text = "Object type"


$Page1_RBNewADAccount = New-Object System.Windows.Forms.RadioButton
$Page1_RBNewADAccount.AutoSize = $True
$Page1_RBNewADAccount.Location = New-Object System.Drawing.Point(10,10)
$Page1_RBNewADAccount.Text = "New AD Account"
$Page1_RBNewADAccount.Checked = $False
#$Page1_RBNewADAccount.Visible = $False
$Page1_GBObjectType.Controls.AddRange($Page1_RBNewADAccount)
$Page1_RBNewADAccount.Add_CheckedChanged(
    {
        if($Page1_RBNewADAccount.Checked -eq $true)
        {
            $Page1_TBJoinerFirstName.Enabled=$True
            $Page1_TBJoinerLastName.Enabled=$True
            $Page1_TBJoinerDisplayName.Enabled=$True
            $Page1_TBJoinerFullName.Enabled=$True
            #$Page1_TBJoinerTicket.Enabled=$True            
            #$Page1_TBJoinerOU.Enabled=$True
            $Page1_TBInitialPassword.Enabled=$true
            $Page1_BtnGeneratePassword.Enabled=$true
            $Page1_TBEmail.Enabled=$false
            $Page1_TBJoinerSamAccountName.Enabled=$false
            $Page1_TBJoinerEA11.Text="PERSONAL"
            $Page1_TBJoinerFirstName.Text=""
            $Page1_TBJoinerLastName.Text=""
            $Page1_TBJoinerEmployeeID.Text=""
            $Page1_TBJoinerDisplayName.Text=""
            $Page1_TBJoinerFullName.Text=""
            $Page1_TBEmail.Text=""
            $Page1_TBJoinerOU.Text=""
            $Page1_TBJoinerLogonScript.Text=""
            $Page1_TBJoinerEA9.Text=""
            $Page1_TBJoinerSamAccountName.Text=""
            $Page1_DepartmentList.SelectedIndex=-1
            $Page1_OfficeList.SelectedIndex=-1

            $HighlightObjects=@($Page1_LBJoinerFirstName,$Page1_LBJoinerLastName,$Page1_LbJoinerEmployeeID,$Page1_LbJoinerOffice, $Page1_LbDepartment, $Page1_LbJoinerDisplayName, $Page1_LbJoinerTicket)
            for($i=0;$i -lt $HighlightObjects.Length;$i=$i+1)
            {
                $HighlightObjects[$i].ForeColor="Red"
                $HighlightObjects[$i].Text=($i+1).ToString() + "." + $HighlightObjects[$i].Tag
            }
            $RestoreObjects=@($Page1_LBJoinerSamAccountName,$Page1_LbJoinerEA9,$Page1_LbJoinerLogonScript)
            for($j=0;$j -lt $RestoreObjects.Length;$j=$j+1)
            {
                $RestoreObjects[$j].ForeColor="Black"
                $RestoreObjects[$j].Text=$RestoreObjects[$j].Tag
            }
            $global:ActionLogs.Items.Clear
        }
    }
)


$Page1_RBFinalizeADAccount = New-Object System.Windows.Forms.RadioButton
$Page1_RBFinalizeADAccount.AutoSize = $True
$Page1_RBFinalizeADAccount.Location = New-Object System.Drawing.Point(200,10)
$Page1_RBFinalizeADAccount.Text = "Finalize AD Account"
$Page1_RBFinalizeADAccount.Checked = $True
#$Page1_RBFinalizeADAccount.Visible = $False
$Page1_GBObjectType.Controls.AddRange($Page1_RBFinalizeADAccount)
$Page1_RBFinalizeADAccount.Add_CheckedChanged(
    {
        if($Page1_RBFinalizeADAccount.Checked -eq $true)
        {
            $Page1_TBJoinerFirstName.Enabled=$False
            $Page1_TBJoinerLastName.Enabled=$False
            #$Page1_TBJoinerDisplayName.Enabled=$False
            #$Page1_TBJoinerFullName.Enabled=$False
            $Page1_TBEmail.Enabled=$false
            #$Page1_TBJoinerTicket.Enabled=$False
            #$Page1_TBJoinerOU.Enabled=$False
            $Page1_TBJoinerSamAccountName.Enabled=$true
            $Page1_TBJoinerEA11.Text="PERSONAL"
            $Page1_TBJoinerFirstName.Text=""
            $Page1_TBJoinerLastName.Text=""
            $Page1_TBJoinerEmployeeID.Text=""
            $Page1_TBJoinerDisplayName.Text=""
            $Page1_TBJoinerFullName.Text=""
            $Page1_TBEmail.Text=""
            $Page1_TBJoinerOU.Text=""
            $Page1_TBJoinerLogonScript.Text=""
            $Page1_TBJoinerEA9.Text=""
            $Page1_TBJoinerSamAccountName.Text=""
            $Page1_TBJoinerTicket.Text=""
            $Page1_DepartmentList.SelectedIndex=-1
            $Page1_OfficeList.SelectedIndex=-1
            $Page1_TBInitialPassword.Enabled=$false
            $Page1_BtnGeneratePassword.Enabled=$false

            $HighlightObjects=@($Page1_LBJoinerSamAccountName,$Page1_LbJoinerOffice, $Page1_LbDepartment, $Page1_LbJoinerDisplayName,$Page1_LbJoinerEmployeeID,$Page1_LbJoinerEA9,$Page1_LbJoinerTicket)
            for($i=0;$i -lt $HighlightObjects.Length;$i=$i+1)
            {
                $HighlightObjects[$i].ForeColor="Red"
                $HighlightObjects[$i].Text=($i+1).ToString() + "." + $HighlightObjects[$i].Tag
            }
            $RestoreObjects=@($Page1_LBJoinerFirstName,$Page1_LBJoinerLastName)
            for($j=0;$j -lt $RestoreObjects.Length;$j=$j+1)
            {
                $RestoreObjects[$j].ForeColor="Black"
                $RestoreObjects[$j].Text=$RestoreObjects[$j].Tag
            }
            $global:ActionLogs.Items.Clear
        }
    }
)



$Page1_LbJoinerFirstName = New-Object System.Windows.Forms.Label
$Page1_LbJoinerFirstName.Text = "First Name"
$Page1_LbJoinerFirstName.AutoSize = $True
$Page1_LbJoinerFirstName.BackColor = "Transparent"
$Page1_LbJoinerFirstName.Font = New-Object System.Drawing.Font("Tahoma",10,[System.Drawing.FontStyle]::Regular)
$Page1_LbJoinerFirstName.ForeColor = "Black"
$Page1_LbJoinerFirstName.Location = New-Object System.Drawing.Point(10,40)
$Page1_LbJoinerFirstName.Tag = "First Name"

$Page1_TBJoinerFirstName = New-Object System.Windows.Forms.TextBox
$Page1_TBJoinerFirstName.Font = New-Object System.Drawing.Font("Tahoma",10,[System.Drawing.FontStyle]::Regular)
$Page1_TBJoinerFirstName.Location = New-Object System.Drawing.Point(105,40)
$Page1_TBJoinerFirstName.Size = New-Object System.Drawing.Size(180,20)
$Page1_TBJoinerFirstName.Tag = "First Name"
$Page1_TBJoinerFirstName.TabIndex = 1
$Page1_TBJoinerFirstName.Enabled = $false
$Page1_TBJoinerFirstName.Add_LostFocus(
    {
    if(($Page1_TBJoinerFirstName.Text.Length -gt 0) -and ($Page1_TBJoinerLastName.Text.Length -gt 0) -and ($Page1_TBJoinerEmployeeID.Text.Length -gt 0) -and ($Page1_OfficeList.SelectedItem.Length -gt 0))
    {
        PopulateBasicConfig
    }
}
)
#$Page1_TBJoinerFirstName.TabIndex = 5

$Page1_LbJoinerLastName = New-Object System.Windows.Forms.Label
$Page1_LbJoinerLastName.Text = "Last Name"
$Page1_LbJoinerLastName.AutoSize = $True
$Page1_LbJoinerLastName.BackColor = "Transparent"
$Page1_LbJoinerLastName.Font = New-Object System.Drawing.Font("Tahoma",10,[System.Drawing.FontStyle]::Regular)
$Page1_LbJoinerLastName.ForeColor = "Black"
$Page1_LbJoinerLastName.Location = New-Object System.Drawing.Point(10,70)
$Page1_LbJoinerLastName.Tag = "Last Name"

$Page1_TBJoinerLastName = New-Object System.Windows.Forms.TextBox
$Page1_TBJoinerLastName.Font = New-Object System.Drawing.Font("Tahoma",10,[System.Drawing.FontStyle]::Regular)
$Page1_TBJoinerLastName.Location = New-Object System.Drawing.Point(105,70)
$Page1_TBJoinerLastName.Size = New-Object System.Drawing.Size(180,20)
$Page1_TBJoinerLastName.Tag = "Last Name"
$Page1_TBJoinerLastName.TabIndex = 2
$Page1_TBJoinerLastName.Enabled = $false
$Page1_TBJoinerLastName.Add_LostFocus(
    {
    if(($Page1_TBJoinerFirstName.Text.Length -gt 0) -and ($Page1_TBJoinerLastName.Text.Length -gt 0) -and ($Page1_TBJoinerEmployeeID.Text.Length -gt 0) -and ($Page1_OfficeList.SelectedItem.Length -gt 0))
    {
        PopulateBasicConfig
    }
}
)

$Page1_LbJoinerDisplayName = New-Object System.Windows.Forms.Label
$Page1_LbJoinerDisplayName.Text = "Display Name"
$Page1_LbJoinerDisplayName.AutoSize = $True
$Page1_LbJoinerDisplayName.BackColor = "Transparent"
$Page1_LbJoinerDisplayName.Font = New-Object System.Drawing.Font("Tahoma",10,[System.Drawing.FontStyle]::Regular)
$Page1_LbJoinerDisplayName.ForeColor = "Black"
$Page1_LbJoinerDisplayName.Location = New-Object System.Drawing.Point(300,40)
$Page1_LbJoinerDisplayName.Tag = "Display Name"

$Page1_TBJoinerDisplayName = New-Object System.Windows.Forms.TextBox
$Page1_TBJoinerDisplayName.Font = New-Object System.Drawing.Font("Tahoma",10,[System.Drawing.FontStyle]::Regular)
$Page1_TBJoinerDisplayName.Location = New-Object System.Drawing.Point(420,40)
$Page1_TBJoinerDisplayName.Size = New-Object System.Drawing.Size(180,20)
$Page1_TBJoinerDisplayName.TabIndex = 7
$Page1_TBJoinerDisplayName.Tag = "Display Name"

$Page1_LbJoinerFullName = New-Object System.Windows.Forms.Label
$Page1_LbJoinerFullName.Text = "Full Name"
$Page1_LbJoinerFullName.AutoSize = $True
$Page1_LbJoinerFullName.BackColor = "Transparent"
$Page1_LbJoinerFullName.Font = New-Object System.Drawing.Font("Tahoma",10,[System.Drawing.FontStyle]::Regular)
$Page1_LbJoinerFullName.ForeColor = "Black"
$Page1_LbJoinerFullName.Location = New-Object System.Drawing.Point(300,70)

$Page1_TBJoinerFullName = New-Object System.Windows.Forms.TextBox
$Page1_TBJoinerFullName.Font = New-Object System.Drawing.Font("Tahoma",10,[System.Drawing.FontStyle]::Regular)
$Page1_TBJoinerFullName.Location = New-Object System.Drawing.Point(420,70)
$Page1_TBJoinerFullName.Size = New-Object System.Drawing.Size(180,20)
$Page1_TBJoinerFullName.TabIndex = 8
$Page1_TBJoinerFullName.Tag = "Full Name"

$Page1_LbJoinerTicket = New-Object System.Windows.Forms.Label
$Page1_LbJoinerTicket.Text = "Join Ticket"
$Page1_LbJoinerTicket.AutoSize = $True
$Page1_LbJoinerTicket.BackColor = "Transparent"
$Page1_LbJoinerTicket.Font = New-Object System.Drawing.Font("Tahoma",10,[System.Drawing.FontStyle]::Regular)
$Page1_LbJoinerTicket.ForeColor = "Black"
$Page1_LbJoinerTicket.Location = New-Object System.Drawing.Point(10,100)
$Page1_LbJoinerTicket.Tag="Join Ticket"

$Page1_TBJoinerTicket = New-Object System.Windows.Forms.TextBox
$Page1_TBJoinerTicket.Font = New-Object System.Drawing.Font("Tahoma",10,[System.Drawing.FontStyle]::Regular)
$Page1_TBJoinerTicket.Location = New-Object System.Drawing.Point(105,100)
$Page1_TBJoinerTicket.Size = New-Object System.Drawing.Size(180,20)
$Page1_TBJoinerTicket.TabIndex = 9
$Page1_TBJoinerTicket.Tag = "Join Ticket number"

$Page1_LbJoinerSamAccountName = New-Object System.Windows.Forms.Label
$Page1_LbJoinerSamAccountName.Text = "SamAccount"
$Page1_LbJoinerSamAccountName.AutoSize = $True
$Page1_LbJoinerSamAccountName.BackColor = "Transparent"
$Page1_LbJoinerSamAccountName.Font = New-Object System.Drawing.Font("Tahoma",10,[System.Drawing.FontStyle]::Regular)
$Page1_LbJoinerSamAccountName.ForeColor = "Black"
$Page1_LbJoinerSamAccountName.Location = New-Object System.Drawing.Point(300,100)
$Page1_LbJoinerSamAccountName.Tag = "SamAccount"

$Page1_TBJoinerSamAccountName = New-Object System.Windows.Forms.TextBox
$Page1_TBJoinerSamAccountName.Font = New-Object System.Drawing.Font("Tahoma",10,[System.Drawing.FontStyle]::Regular)
$Page1_TBJoinerSamAccountName.Location = New-Object System.Drawing.Point(420,100)
$Page1_TBJoinerSamAccountName.Size = New-Object System.Drawing.Size(180,20)
$Page1_TBJoinerSamAccountName.TabIndex = 9
$Page1_TBJoinerSamAccountName.Tag = "SamAccount"
#$Page1_TBJoinerSamAccountName.Enabled = $false
$Page1_TBJoinerSamAccountName.Add_LostFocus(
    {
        $Page1_TBJoinerEA11.ForeColor = "Black"
        $Page1_TbJoinerEmployeeType.ForeColor = "Black"

        if(($Page1_RBFinalizeADAccount.Checked -eq $True) -and ($Page1_TBJoinerSamAccountName.Text -ne ""))
        {
                #$userfilter =" SamAccountName -eq '" + $Page1_TBJoinerSamAccountName.Text + "')"
                try {
                    #$userInfo=Get-Aduser -LDAPFilter $userfilter -Property CN,Displayname,extensionAttribute1,extensionAttribute9,extensionAttribute11,Department
                    $userInfo=Get-Aduser -Identity $Page1_TBJoinerSamAccountName.Text -Properties *
                    $Page1_TBEmail.Text =$userInfo.mail
                    $Page1_TBJoinerDisplayName.Text = $userInfo.DisplayName
                    $Page1_TBJoinerEmployeeID.Text = $userInfo.extensionAttribute1
                    $Page1_TBTitle.Text =$userInfo.Title
                    $Page1_TBJoinerLogonScript.Text = $userInfo.scriptPath
                    $Page1_TBJoinerFullName.Text = $userInfo.cn
                    $Page1_TBJoinerFirstName.Text = $userInfo.GivenName
                    $Page1_TBJoinerLastName.Text = $userInfo.sn
                    $Page1_TBHomeTel.Text = $userInfo.HomePhone
                    if(($userInfo.extensionAttribute11 -eq "") -or ($userInfo.extensionAttribute11 -ne "PERSONAL"))
                    {
                        $Page1_TBJoinerEA11.Text = "PERSONAL"
                        $Page1_TBJoinerEA11.ForeColor = "Red"
                    }

                    if(($userInfo.employeeType -eq "") -or ($userInfo.employeeType -ne "Employee"))
                    {
                        $Page1_TbJoinerEmployeeType.Text = "Employee"
                        $Page1_TbJoinerEmployeeType.ForeColor = "Red"
                    }
                    else {
                        $Page1_TbJoinerEmployeeType.Text = $userInfo.employeeType
                    }

                    if($userInfo.extensionAttribute9 -ne "")
                    {
                        $Page1_OfficeList.SelectedItem = $userInfo.extensionAttribute9
                        #$Page1_TBJoinerEA9.Text = $userInfo.extensionAttribute9
                    }

                    if($userInfo.Department -ne "")
                    {
                        $Page1_DepartmentList.SelectedItem = $userInfo.Department
                    }
                }
                catch {
                    $Page1_TBEmail.Text =""
                    $Page1_TBJoinerDisplayName.Text = ""
                    $Page1_TBJoinerEmployeeID.Text = ""
                    #$Page1_TBJoinerEA11.Text = ""
                    $Page1_TBJoinerFullName.Text = ""
                    $Page1_TBJoinerFirstName.Text = ""
                    $Page1_TBJoinerLastName.Text = ""
                    $Page1_TBJoinerEA9.Text = ""
                    $Page1_TBJoinerLogonScript.Text = ""
                    $Page1_TBJoinerOU.Text=""
                    $Page1_OfficeList.SelectedIndex=-1
                    #$Page1_DepartmentList.SelectedItem=""
                    $Page1_DefaultGroupList.Items.Clear()
                }
        }
    }
)

$Page1_LbJoinerOffice = New-Object System.Windows.Forms.Label
$Page1_LbJoinerOffice.Text = "Office"
$Page1_LbJoinerOffice.AutoSize = $True
$Page1_LbJoinerOffice.BackColor = "Transparent"
$Page1_LbJoinerOffice.Font = New-Object System.Drawing.Font("Tahoma",10,[System.Drawing.FontStyle]::Regular)
$Page1_LbJoinerOffice.ForeColor = "Black"
$Page1_LbJoinerOffice.Location = New-Object System.Drawing.Point(10,130)
$Page1_LbJoinerOffice.Tag="Office"

$Page1_OfficeList = New-Object System.Windows.Forms.ComboBox
$Page1_OfficeList.DropDownStyle = [System.Windows.Forms.ComboBoxStyle]::DropDownList
$Page1_OfficeList.Font = New-Object System.Drawing.Font("Tahoma",10,[System.Drawing.FontStyle]::Regular)
$Page1_OfficeList.Location = New-Object System.Drawing.Point(105,130)
$Page1_OfficeList.Size = New-Object System.Drawing.Size(180,20)
$Page1_OfficeList.Height=60
$Page1_OfficeList.Tag = "Office"
$Page1_OfficeList.TabIndex = 5
foreach($sc in $SiteCodeList)
{
    $Page1_OfficeList.Items.Add($sc)
}

$Page1_OfficeList.Add_SelectedIndexChanged(
    {
        <#
        if(($Page1_RBFinalizeADAccount.Checked -eq $True) -AND ($Page1_TBJoinerEA9.Text -ne $Page1_OfficeList.SelectedItem))
        {
            $Page1_TBJoinerEA9.ForeColor="Red"
        }
            #>
        $Page1_TBJoinerEA9.Text = $Page1_OfficeList.SelectedItem
        
        if($Page1_OfficeList.SelectedItem -eq "TWNTC100")
        {
            $Page1_TBHomeTel.Enabled=$True
        }
        else {
            $Page1_TBHomeTel.Enabled=$False
        }

        if(($Page1_TBJoinerFirstName.Text.Length -gt 0) -and ($Page1_TBJoinerLastName.Text.Length -gt 0) -and ($Page1_TBJoinerEmployeeID.Text.Length -gt 0) -and ($Page1_OfficeList.SelectedItem.Length -gt 0))
        {
            PopulateBasicConfig
        }
        foreach($sc in $SiteCodeList)
        {
            if($Page1_OfficeList.SelectedItem -eq $sc)
            {
                $Page1_TBJoinerOU.Text = "OU=Active,OU=Users,OU=" + $sc.Substring(0,2) + ",OU=WSPObjects,DC=corp,DC=pbwan,DC=net"
            }
        }

        $Page1_DepartmentList.Items.Clear()
        $Page1_DefaultGroupList.Items.Clear()
        #$Page1_TBJoinerLogonScript.Text=""
        for($site=0;$site -le $xmldata.Sites.ChildNodes.Count-1;$site++)
        {
            if($xmldata.Sites.ChildNodes[$site].Code -eq $Page1_OfficeList.SelectedItem)
            {
                foreach($Department in $xmldata.Sites.ChildNodes[$site].Department)
                {
                    $Page1_DepartmentList.Items.Add($Department.Attributes["Name"].Value) 
                }
            }
        }
    }

)

$Page1_LbDepartment = New-Object System.Windows.Forms.Label
$Page1_LbDepartment.Text = "Department"
$Page1_LbDepartment.AutoSize = $True
$Page1_LbDepartment.BackColor = "Transparent"
$Page1_LbDepartment.Font = New-Object System.Drawing.Font("Tahoma",10,[System.Drawing.FontStyle]::Regular)
$Page1_LbDepartment.ForeColor = "Black"
$Page1_LbDepartment.Location = New-Object System.Drawing.Point(300,130)
$Page1_LbDepartment.Tag="Department"

$Page1_DepartmentList = New-Object System.Windows.Forms.ComboBox
$Page1_DepartmentList.DropDownStyle = [System.Windows.Forms.ComboBoxStyle]::DropDownList
$Page1_DepartmentList.Sorted = $True
$Page1_DepartmentList.Font = New-Object System.Drawing.Font("Tahoma",8,[System.Drawing.FontStyle]::Regular)
$Page1_DepartmentList.Location = New-Object System.Drawing.Point(420,130)
$Page1_DepartmentList.Size = New-Object System.Drawing.Size(180,20)
$Page1_DepartmentList.Height = 60
$Page1_DepartmentList.Tag = "Department"
$Page1_DepartmentList.TabIndex = 6
$Page1_DepartmentList.Add_SelectedIndexChanged(
    {
        $Page1_DefaultGroupList.Items.Clear()
        $NewJoinerDefaultGroups = New-Object -TypeName 'System.Collections.ArrayList'
        $NewJoinerDefaultGroups.Clear()
        $Page1_TBJoinerLogonScript.ForeColor="Black"
        #Add Corporate default groups
        if($CorporateDefaultGroups.count -gt 0)
        {
            foreach($CorporateDefaultGroup in $CorporateDefaultGroups)
            {
                #$Page1_DefaultGroupList.Items.Add($CorporateDefaultGroup)
                $NewJoinerDefaultGroups.Add($CorporateDefaultGroup)
            }
        }
        #Add Site & Team default groups
        for($site=0;$site -le $xmldata.Sites.ChildNodes.Count-1;$site++)
        {
            if($xmldata.Sites.ChildNodes[$site].Code -eq $Page1_OfficeList.SelectedItem)
            {
                $SiteDefaultGroups=$xmldata.Sites.ChildNodes[$site].SiteDefaultGroup
                if($SiteDefaultGroups -ne "")
                {
                    $SiteDefaultGroups=$SiteDefaultGroups.Split('|')
                    foreach($SiteDefaultGroup in $SiteDefaultGroups)
                    {
                        #$Page1_DefaultGroupList.Items.Add($SiteDefaultGroup)
                        $NewJoinerDefaultGroups.Add($SiteDefaultGroup)
                    }
                }

                foreach($Department in $xmldata.Sites.ChildNodes[$site].ChildNodes)
                {
                    if($Department.Name -eq $Page1_DepartmentList.SelectedItem)
                    {
                        $TeamDefaultGroups=$Department.InnerText.Split('|')
                        foreach($TeamDefaultGroup in $TeamDefaultGroups)
                        {
                            #$Page1_DefaultGroupList.Items.Add($TeamDefaultGroup)
                            $NewJoinerDefaultGroups.Add($TeamDefaultGroup)
                        }
                        #Add logon script
                        if((($Page1_TBJoinerLogonScript.Text -ne $Department.LogonScript) -or ($Page1_TBJoinerLogonScript.Text -eq "")) -and ($Page1_RBFinalizeADAccount.Checked -eq $True))
                        {
                            $Page1_TBJoinerLogonScript.ForeColor="Red"
                        }
                        $Page1_TBJoinerLogonScript.Text=$Department.LogonScript
                    }
                }

            }
        }

        foreach($group in ($NewJoinerDefaultGroups|get-unique))
        {
            $Page1_DefaultGroupList.Items.Add($group)
        }
    }
)

$Page1_LbEmail = New-Object System.Windows.Forms.Label
$Page1_LbEmail.Text = "Email(UPN)"
$Page1_LbEmail.AutoSize = $True
$Page1_LbEmail.BackColor = "Transparent"
$Page1_LbEmail.Font = New-Object System.Drawing.Font("Tahoma",10,[System.Drawing.FontStyle]::Regular)
$Page1_LbEmail.ForeColor = "Black"
$Page1_LbEmail.Location = New-Object System.Drawing.Point(10,160)

$Page1_TBEmail = New-Object System.Windows.Forms.TextBox
$Page1_TBEmail.Font = New-Object System.Drawing.Font("Tahoma",10,[System.Drawing.FontStyle]::Regular)
$Page1_TBEmail.Location = New-Object System.Drawing.Point(105,160)
$Page1_TBEmail.Size = New-Object System.Drawing.Size(180,20)
$Page1_TBEmail.Enabled=$false
$Page1_TBEmail.TabIndex = 3
$Page1_TBEmail.Tag="Email"

$Page1_LbTitle = New-Object System.Windows.Forms.Label
$Page1_LbTitle.Text = "Title"
$Page1_LbTitle.AutoSize = $True
$Page1_LbTitle.BackColor = "Transparent"
$Page1_LbTitle.Font = New-Object System.Drawing.Font("Tahoma",10,[System.Drawing.FontStyle]::Regular)
$Page1_LbTitle.ForeColor = "Black"
$Page1_LbTitle.Location = New-Object System.Drawing.Point(300,160)

$Page1_TBTitle = New-Object System.Windows.Forms.TextBox
$Page1_TBTitle.Font = New-Object System.Drawing.Font("Tahoma",10,[System.Drawing.FontStyle]::Regular)
$Page1_TBTitle.Location = New-Object System.Drawing.Point(420,160)
$Page1_TBTitle.Size = New-Object System.Drawing.Size(180,20)
$Page1_TBTitle.TabIndex = 3
$Page1_TBTitle.Tag="Title"


$Page1_LbJoinerLogonScript = New-Object System.Windows.Forms.Label
$Page1_LbJoinerLogonScript.Text = "Logon Script"
$Page1_LbJoinerLogonScript.AutoSize = $True
$Page1_LbJoinerLogonScript.BackColor = "Transparent"
$Page1_LbJoinerLogonScript.Font = New-Object System.Drawing.Font("Tahoma",10,[System.Drawing.FontStyle]::Regular)
$Page1_LbJoinerLogonScript.ForeColor = "Black"
$Page1_LbJoinerLogonScript.Location = New-Object System.Drawing.Point(10,190)
$Page1_LbJoinerLogonScript.Tag="Logon Script"

$Page1_TBJoinerLogonScript = New-Object System.Windows.Forms.TextBox
$Page1_TBJoinerLogonScript.Font = New-Object System.Drawing.Font("Tahoma",10,[System.Drawing.FontStyle]::Regular)
$Page1_TBJoinerLogonScript.Location = New-Object System.Drawing.Point(105,190)
$Page1_TBJoinerLogonScript.Size = New-Object System.Drawing.Size(180,20)

$Page1_LbHomeTel = New-Object System.Windows.Forms.Label
$Page1_LbHomeTel.Text = "Tel(Home)"
$Page1_LbHomeTel.AutoSize = $True
$Page1_LbHomeTel.BackColor = "Transparent"
$Page1_LbHomeTel.Font = New-Object System.Drawing.Font("Tahoma",10,[System.Drawing.FontStyle]::Regular)
$Page1_LbHomeTel.ForeColor = "Black"
$Page1_LbHomeTel.Location = New-Object System.Drawing.Point(300,190)


$Page1_TBHomeTel = New-Object System.Windows.Forms.TextBox
$Page1_TBHomeTel.Font = New-Object System.Drawing.Font("Tahoma",10,[System.Drawing.FontStyle]::Regular)
$Page1_TBHomeTel.Location = New-Object System.Drawing.Point(420,190)
$Page1_TBHomeTel.Size = New-Object System.Drawing.Size(180,20)
$Page1_TBHomeTel.Enabled=$false
$Page1_TBHomeTel.TabIndex = 3
$Page1_TBHomeTel.Tag="Tel(Home)"
$Page1_TBHomeTel.Enabled=$False
$Page1_TBHomeTel.MaxLength=5

$Page1_LbJoinerEmployeeID = New-Object System.Windows.Forms.Label
$Page1_LbJoinerEmployeeID.Text = "EA1\/Emp.ID"
$Page1_LbJoinerEmployeeID.AutoSize = $True
$Page1_LbJoinerEmployeeID.BackColor = "Transparent"
$Page1_LbJoinerEmployeeID.Font = New-Object System.Drawing.Font("Tahoma",10,[System.Drawing.FontStyle]::Regular)
$Page1_LbJoinerEmployeeID.ForeColor = "Black"
$Page1_LbJoinerEmployeeID.Location = New-Object System.Drawing.Point(10,220)
$Page1_LbJoinerEmployeeID.Tag="EA1/Emp.ID"

$Page1_TBJoinerEmployeeID = New-Object System.Windows.Forms.TextBox
$Page1_TBJoinerEmployeeID.Font = New-Object System.Drawing.Font("Tahoma",10,[System.Drawing.FontStyle]::Regular)
$Page1_TBJoinerEmployeeID.Location = New-Object System.Drawing.Point(105,220)
$Page1_TBJoinerEmployeeID.Size = New-Object System.Drawing.Size(180,20)
$Page1_TBJoinerEmployeeID.Tag="EmployeeID"
$Page1_TBJoinerEmployeeID.TabIndex = 4
$Page1_TBJoinerEmployeeID.Add_TextChanged(
    {
    if(($Page1_TBJoinerFirstName.Text.Length -gt 0) -and ($Page1_TBJoinerLastName.Text.Length -gt 0) -and ($Page1_TBJoinerEmployeeID.Text.Length -gt 0) -and ($Page1_OfficeList.SelectedItem.Length -gt 0))
    {
        PopulateBasicConfig
    }
}
)

$Page1_LbJoinerEA9 = New-Object System.Windows.Forms.Label
$Page1_LbJoinerEA9.Text = "EA9"
$Page1_LbJoinerEA9.AutoSize = $True
$Page1_LbJoinerEA9.BackColor = "Transparent"
$Page1_LbJoinerEA9.Font = New-Object System.Drawing.Font("Tahoma",10,[System.Drawing.FontStyle]::Regular)
$Page1_LbJoinerEA9.ForeColor = "Black"
$Page1_LbJoinerEA9.Location = New-Object System.Drawing.Point(300,220)
$Page1_LbJoinerEA9.Tag="EA9"

$Page1_TBJoinerEA9 = New-Object System.Windows.Forms.TextBox
$Page1_TBJoinerEA9.Font = New-Object System.Drawing.Font("Tahoma",10,[System.Drawing.FontStyle]::Regular)
$Page1_TBJoinerEA9.Location = New-Object System.Drawing.Point(420,220)
$Page1_TBJoinerEA9.Size = New-Object System.Drawing.Size(180,20)

$Page1_LbJoinerEA11 = New-Object System.Windows.Forms.Label
$Page1_LbJoinerEA11.Text = "EA11"
$Page1_LbJoinerEA11.AutoSize = $True
$Page1_LbJoinerEA11.BackColor = "Transparent"
$Page1_LbJoinerEA11.Font = New-Object System.Drawing.Font("Tahoma",10,[System.Drawing.FontStyle]::Regular)
$Page1_LbJoinerEA11.ForeColor = "Black"
$Page1_LbJoinerEA11.Location = New-Object System.Drawing.Point(10,250)

$Page1_TBJoinerEA11 = New-Object System.Windows.Forms.TextBox
$Page1_TBJoinerEA11.Font = New-Object System.Drawing.Font("Tahoma",10,[System.Drawing.FontStyle]::Regular)
$Page1_TBJoinerEA11.Location = New-Object System.Drawing.Point(105,250)
$Page1_TBJoinerEA11.Size = New-Object System.Drawing.Size(180,20)
$Page1_TBJoinerEA11.Text = "PERSONAL"


$Page1_LbJoinerEmployeeType = New-Object System.Windows.Forms.Label
$Page1_LbJoinerEmployeeType.Text = "EmployeeType"
$Page1_LbJoinerEmployeeType.AutoSize = $True
$Page1_LbJoinerEmployeeType.BackColor = "Transparent"
$Page1_LbJoinerEmployeeType.Font = New-Object System.Drawing.Font("Tahoma",10,[System.Drawing.FontStyle]::Regular)
$Page1_LbJoinerEmployeeType.ForeColor = "Black"
$Page1_LbJoinerEmployeeType.Location = New-Object System.Drawing.Point(300,250)

$Page1_TBJoinerEmployeeType = New-Object System.Windows.Forms.TextBox
$Page1_TBJoinerEmployeeType.Font = New-Object System.Drawing.Font("Tahoma",10,[System.Drawing.FontStyle]::Regular)
$Page1_TBJoinerEmployeeType.Location = New-Object System.Drawing.Point(420,250)
$Page1_TBJoinerEmployeeType.Size = New-Object System.Drawing.Size(180,20)
$Page1_TBJoinerEmployeeType.Text = "Employee"


$Page1_LbJoinerOU = New-Object System.Windows.Forms.Label
$Page1_LbJoinerOU.Text = "OU"
$Page1_LbJoinerOU.AutoSize = $True
$Page1_LbJoinerOU.BackColor = "Transparent"
$Page1_LbJoinerOU.Font = New-Object System.Drawing.Font("Tahoma",10,[System.Drawing.FontStyle]::Regular)
$Page1_LbJoinerOU.ForeColor = "Black"
$Page1_LbJoinerOU.Location = New-Object System.Drawing.Point(10,280)

$Page1_TBJoinerOU = New-Object System.Windows.Forms.TextBox
$Page1_TBJoinerOU.Font = New-Object System.Drawing.Font("Tahoma",8,[System.Drawing.FontStyle]::Regular)
$Page1_TBJoinerOU.Location = New-Object System.Drawing.Point(105,280)
$Page1_TBJoinerOU.Size = New-Object System.Drawing.Size(495,20)
$Page1_TBJoinerOU.Tag = "OU"
$Page1_TBJoinerOU.Enabled=$false
#$Page1_TBJoinerDepartment.TabIndex = 5


$Page1_LbDefaultGroups = New-Object System.Windows.Forms.Label
$Page1_LbDefaultGroups.Text = "Groups`nTo be added"
$Page1_LbDefaultGroups.AutoSize = $True
$Page1_LbDefaultGroups.BackColor = "Transparent"
$Page1_LbDefaultGroups.Font = New-Object System.Drawing.Font("Tahoma",10,[System.Drawing.FontStyle]::Regular)
$Page1_LbDefaultGroups.ForeColor = "Black"
$Page1_LbDefaultGroups.Location = New-Object System.Drawing.Point(10,310)

$Page1_DefaultGroupList = New-Object System.Windows.Forms.ListBox
$Page1_DefaultGroupList.Font = New-Object System.Drawing.Font("Tahoma",10,[System.Drawing.FontStyle]::Regular)
$Page1_DefaultGroupList.Location = New-Object System.Drawing.Point(105,310)
$Page1_DefaultGroupList.Size = New-Object System.Drawing.Size(495,90)

$Page1_BtnAddGroup = New-Object System.Windows.Forms.Button
$Page1_BtnAddGroup.Size = New-Object System.Drawing.Size(25,25)
$Page1_BtnAddGroup.Location = New-Object System.Drawing.Point(540,400)
$Page1_BtnAddGroup.Text = "+"
$Page1_BtnAddGroup.Font = New-Object System.Drawing.Font("Tahoma",10,[System.Drawing.FontStyle]::Regular)
$Page1_BtnAddGroup.Add_Click(
    {
        $TargetGroups=$Page1_TBNewGroup.Text.split(";")
        foreach($TargetGroup in $TargetGroups)
        {
            if($TargetGroup.Trim() -ne "")
            {
                $groupcount=@(Get-ADGroup -Filter {(mail -eq $TargetGroup) -or (sAMAccountName -eq $TargetGroup)}).Count
                if($groupcount -gt 0)
                {
                    $Page1_DefaultGroupList.Items.Add($TargetGroup.Trim())
                    $Page1_TBNewGroup.Text=""
                }
                else {
                    [System.Windows.Forms.Messagebox]::Show("Group:[" + $TargetGroup + "] not found!")
                }
            }
        }
    }
)

$Page1_TBNewGroup = New-Object System.Windows.Forms.TextBox
$Page1_TBNewGroup.Font = New-Object System.Drawing.Font("Tahoma",10,[System.Drawing.FontStyle]::Regular)
$Page1_TBNewGroup.Location = New-Object System.Drawing.Point(105,400)
$Page1_TBNewGroup.Size = New-Object System.Drawing.Size(420,20)


$Page1_BtnRemoveGroup = New-Object System.Windows.Forms.Button
$Page1_BtnRemoveGroup.Size = New-Object System.Drawing.Size(25,25)
$Page1_BtnRemoveGroup.Location = New-Object System.Drawing.Point(575,400)
$Page1_BtnRemoveGroup.Text = "-"
$Page1_BtnRemoveGroup.Font = New-Object System.Drawing.Font("Tahoma",10,[System.Drawing.FontStyle]::Regular)
$Page1_BtnRemoveGroup.Add_Click(
    {
        if($Page1_DefaultGroupList.SelectedIndex -gt -1)
        {
            $Page1_TBNewGroup.Text = $Page1_DefaultGroupList.SelectedItem
            $Page1_DefaultGroupList.Items.RemoveAt($Page1_DefaultGroupList.SelectedIndex)
        }
    }
)


$Page1_LbInitialPassword = New-Object System.Windows.Forms.Label
$Page1_LbInitialPassword.Text = "Initial PW"
$Page1_LbInitialPassword.AutoSize = $True
$Page1_LbInitialPassword.BackColor = "Transparent"
$Page1_LbInitialPassword.Font = New-Object System.Drawing.Font("Tahoma",10,[System.Drawing.FontStyle]::Regular)
$Page1_LbInitialPassword.ForeColor = "Black"
$Page1_LbInitialPassword.Location = New-Object System.Drawing.Point(10,445)
#$Page1_LbInitialPassword.Visible=$false

$Page1_TBInitialPassword = New-Object System.Windows.Forms.TextBox
$Page1_TBInitialPassword.Font = New-Object System.Drawing.Font("Tahoma",10,[System.Drawing.FontStyle]::Regular)
$Page1_TBInitialPassword.Location = New-Object System.Drawing.Point(105,445)
$Page1_TBInitialPassword.Size = New-Object System.Drawing.Size(180,20)
$Page1_TBInitialPassword.Tag = "Initial Password"
$Page1_TBInitialPassword.text=GenerateInitialPassword
#$Page1_TBInitialPassword.Visible = $false

$Page1_BtnGeneratePassword = New-Object System.Windows.Forms.Button
$Page1_BtnGeneratePassword.Size = New-Object System.Drawing.Size(30,25)
$Page1_BtnGeneratePassword.Location = New-Object System.Drawing.Point(300,445)
$Page1_BtnGeneratePassword.Text = "G"
$Page1_BtnGeneratePassword.Font = New-Object System.Drawing.Font("Tahoma",10,[System.Drawing.FontStyle]::Regular)
$Page1_BtnGeneratePassword.TabIndex=10
#$Page1_BtnGeneratePassword.Visible=$false
$Page1_BtnGeneratePassword.Add_Click(
    {
        $Page1_TBInitialPassword.Text=GenerateInitialPassword
    }
)




$Page1_BtnGo = New-Object System.Windows.Forms.Button
$Page1_BtnGo.Size = New-Object System.Drawing.Size(90,35)
$Page1_BtnGo.Location = New-Object System.Drawing.Point(510,445)
$Page1_BtnGo.Text = "Go"
$Page1_BtnGo.Font = New-Object System.Drawing.Font("Tahoma",10,[System.Drawing.FontStyle]::Regular)
$Page1_BtnGo.TabIndex = 11
$Page1_BtnGo.Add_Click(
    {
        if($Page1_RBNewADAccount.Checked -eq $true) #create new ad accout
        {
            $inputvalid=$true
            $mandatorycontrols=@($Page1_TBJoinerDisplayName,$Page1_TBJoinerFirstName,$Page1_TBJoinerLastName,$Page1_TBJoinerSamAccountName,$Page1_TBJoinerFullName,$Page1_OfficeList,$Page1_DepartmentList,$Page1_TBJoinerTicket,$Page1_TBJoinerEmployeeID, $Page1_TBJoinerOU,$Page1_TBInitialPassword)
            foreach($c in $mandatorycontrols)
            {
                if(($c.Text -eq "") -or ($c.SelectedItem -eq ""))
                {
                    [System.Windows.Forms.Messagebox]::Show($c.Tag + " cannot be blank!!!")
                    $c.Focus()
                    $inputvalid=$false
                }
            }

            if($inputvalid -eq $true)
            {
                $usercount=@(Get-ADUser -Identity $Page1_TBJoinerSamAccountName.Text).Count
                if($usercount -eq 0) #cannot find target samaccount in AD
                {
                    $TargetEmail=$Page1_TBEmail.Text
                    $NewProposeEmail=ProposeNewEmail
                    $EmailConflict = $false
                    if($TargetEmail -ne $NewProposeEmail) #email(UPN) conflicts
                    {
                        $EmailConflict = $true
                        Add-Type -AssemblyName PresentationCore,PresentationFramework
                        $ButtonType = [System.Windows.MessageBoxButton]::OKCancel
                        $MessageIcon = [System.Windows.MessageBoxImage]::Warning
                        $MessageBody = "Target email address conflict with existing user, new proposed address is:`n`n"+ $NewProposeEmail + "`n`nClick OK to create AD account with new proposed email address, Cancel to stop AD account creation."
                        $MessageTitle = "Please choose..."
                        $Result = [System.Windows.MessageBox]::Show($MessageBody,$MessageTitle,$ButtonType,$MessageIcon)
                    }

                    if(($Result -eq "OK") -or ($EmailConflict -eq $false))
                    {
                        $Page1_TBEmail.Text = $NewProposeEmail
                        $Secure_Pwd = ConvertTo-SecureString $Page1_TBInitialPassword.Text -AsPlainText -Force

                        $EAs=@{}
                        $EAs.Clear()
                        $EAs.Add('userPrincipalName',$Page1_TBEmail.Text)
                        $EAs.Add('extensionAttribute1',$Page1_TBJoinerEmployeeID.Text)
                        $EAs.Add('extensionAttribute9',$Page1_TBJoinerEA9.Text)
                        $EAs.Add('extensionAttribute11',$Page1_TBJoinerEA11.Text)
                        $EAs.Add('employeeType',$Page1_TBJoinerEmployeeType.Text)

                        $Result_ADUser = "Success"
                        $Result_GroupMember = "Success"
                        try {
                            New-ADUser -Name $Page1_TBJoinerFullName.Text -Description $Page1_TBJoinerTicket.Text -DisplayName $Page1_TBJoinerDisplayName.Text -GivenName $Page1_TBJoinerFirstName.Text -Department $Page1_DepartmentList.SelectedItem -ScriptPath $Page1_TBJoinerLogonScript.Text -SamAccountName $Page1_TBJoinerSamAccountName.Text -EmailAddress $Page1_TBEmail.Text -Surname $Page1_TBJoinerLastName.Text -Path $Page1_TBJoinerOU.Text -Enabled $true -AccountPassword $Secure_Pwd -HomePhone $Page1_TBHomeTel.Text -Title $Page1_TBTitle.Text -OtherAttributes $EAs

                            try {
                                foreach($ADGroup in $Page1_DefaultGroupList.Items)
                                {
                                    $G=Get-ADGroup -Filter {(mail -eq $ADGroup) -or (sAMAccountName -eq $ADGroup)}
                                    get-aduser -Identity $Page1_TBJoinerSamAccountName.Text | %{Add-ADGroupMember -Identity $G.Name -Members $_.samaccountname} -ErrorAction Continue
                                }
                            }
                            catch {
                                $Result_GroupMember = "Failed"
                            }

                            If($Result_GroupMember -eq "Success")
                            {
                                [System.Windows.Forms.Messagebox]::Show("Good job! AD account created successfully! Congratulation!")
                            }
                            else {
                                [System.Windows.Forms.Messagebox]::Show("Good and bad news for you: `n Good news, AD account created successfully! `n Bad news, not all the groups are added to this account.")
                            }
                        }
                        catch {
                            $Result_ADUser = "Failed"
                            $Result_GroupMember = "Failed"
                            [System.Windows.Forms.Messagebox]::Show("Oops! Failed to create AD account, but don't worry, it's not the end of the world. Please check the following: `n 1. Are you running this PS with your priviledge account; `n 2. SamAccount conflict? `n 3. DC server unavailable? `n 4. RSAT not installed?")
                        }
                        GenerateLog -ActionType "NewADAccount" -ADUser $Result_ADUser -GroupMember $Result_GroupMember

                    }
                    $Page1_TBJoinerLogonScript.ForeColor="Black"
                    $Page1_TBJoinerEA9.ForeColor="Black"
                    $Page1_TBJoinerEA11.ForeColor="Black"
                    $Page1_TBJoinerEmployeeType.ForeColor="Black"
                }
                else {
                    [System.Windows.Forms.Messagebox]::Show("SamAccount : " + $Page1_TBJoinerSamAccountName.Text + " conflict with existing user, please double check FirstName/LastName and Employee ID")
                }
            }
        }
        elseif($Page1_RBFinalizeADAccount.Checked -eq $true) #Finalize ad account
        {
            $inputvalid=$true
            $mandatorycontrols=@($Page1_TBJoinerSamAccountName,$Page1_TBJoinerFullName,$Page1_TBJoinerLastName,$Page1_TBJoinerDisplayName,$Page1_TBJoinerFirstName, $Page1_TBJoinerTicket)
            foreach($c in $mandatorycontrols)
            {
                if($c.Text -eq "")
                {
                    [System.Windows.Forms.Messagebox]::Show($c.Tag + " cannot be blank!!!")
                    $c.Focus()
                    $inputvalid=$false
                }
            }
            
            if($inputvalid -eq $true)
            {
                $EAs=@{}
                $EAs.Clear()
                #$EAs.Add('cn',$Page1_TBJoinerFullName.Text) #Full Name(CN) can't be changed by this way, need to use rename-adobject
                $EAs.Add('extensionAttribute1',$Page1_TBJoinerEmployeeID.Text)
                $EAs.Add('extensionAttribute9',$Page1_TBJoinerEA9.Text)
                $EAs.Add('extensionAttribute11',$Page1_TBJoinerEA11.Text)
                $EAs.Add('employeeType',$Page1_TBJoinerEmployeeType.Text)
                

                $Result_ADUser = "Success"
                $Result_GroupMember = "Success"
                try {
                    #get-aduser -identity $Page1_TBJoinerSamAccountName.Text | set-aduser -DisplayName $Page1_TBJoinerDisplayName.Text -Department $Page1_DepartmentList.SelectedItem -HomePhone $Page1_TBHomeTel.Text -Title $Page1_TBTitle.Text -Replace $EAs
                    if($Page1_TBHomeTel.Text -eq "")
                    {$Page1_TBHomeTel.Text=" "}
                    if($Page1_TBJoinerLogonScript.Text -eq "") 
                    {$Page1_TBJoinerLogonScript.Text=" "}
                    if($Page1_TBTitle.Text -eq "")
                    {$Page1_TBTitle.Text=" "}

                    set-aduser -identity $Page1_TBJoinerSamAccountName.Text -DisplayName $Page1_TBJoinerDisplayName.Text -Department $Page1_DepartmentList.SelectedItem -Title $Page1_TBTitle.Text -HomePhone $Page1_TBHomeTel.Text -ScriptPath $Page1_TBJoinerLogonScript.Text -Replace $EAs
                    get-aduser -identity $Page1_TBJoinerSamAccountName.Text | Rename-ADObject -NewName $Page1_TBJoinerFullName.Text
                    try {
                        foreach($ADGroup in $Page1_DefaultGroupList.Items)
                        {
                            $G=Get-ADGroup -Filter {(mail -eq $ADGroup) -or (sAMAccountName -eq $ADGroup)}
                            get-aduser -Identity $Page1_TBJoinerSamAccountName.Text | %{Add-ADGroupMember -Identity $G.Name -Members $_.samaccountname} -ErrorAction Continue
                        }
                    }
                    catch {
                        $Result_GroupMember = "Failed"
                    }

                    If($Result_GroupMember -eq "Success")
                    {
                        [System.Windows.Forms.Messagebox]::Show("Excellent! AD account updated successfully! Congratulation!")
                    }
                    else {
                        [System.Windows.Forms.Messagebox]::Show("Good and bad news for you: `n Good news, AD account updated successfully! `n Bad news, not all the groups are added to this account.")
                    }

                }
                catch {
                    $Result_ADUser = "Failed"
                    $Result_GroupMember = "Failed"
                    [System.Windows.Forms.Messagebox]::Show("Oops! Failed to update AD account, but don't worry, it's not the end of the world. Please check the following: `n 1. Are you running this PS with your priviledge account; `n 2. Incorrect SamAccount? `n 3. DC server unavailable? `n 4. RSAT not installed?")
                }
                    $Page1_TBJoinerLogonScript.ForeColor="Black"
                    $Page1_TBJoinerEA9.ForeColor="Black"
                    $Page1_TBJoinerEA11.ForeColor="Black"
                    $Page1_TBJoinerEmployeeType.ForeColor="Black"
                GenerateLog -ActionType "FinalizeADAccount" -ADUser $Result_ADUser -GroupMember $Result_GroupMember
            }
        }
    }
)


$TabControl1 = New-Object System.Windows.Forms.TabControl
$tabPage1=New-Object System.Windows.Forms.TabPage
$tabPage1.text = "Account Info"
$tabPage1.TabIndex=0
$tabPage1.Controls.AddRange(($Page1_LbJoinerFirstName, $Page1_TBJoinerFirstName, 
                                $Page1_RBNewADAccount,$Page1_RBFinalizeADAccount,
                                $Page1_LbJoinerLastName, $Page1_TBJoinerLastName,
                                $Page1_LbJoinerEmployeeID, $Page1_TBJoinerEmployeeID, 
                                $Page1_LbJoinerSamAccountName, $Page1_TBJoinerSamAccountName,
                                $Page1_LbDepartment, $Page1_DepartmentList,
                                $Page1_LbHomeTel,$Page1_TBHomeTel,
                                $Page1_LbTitle, $Page1_TBTitle,
                                $Page1_LbJoinerEA9,$Page1_TBJoinerEA9,
                                $Page1_LbJoinerEA11,$Page1_TBJoinerEA11,
                                $Page1_LbJoinerEmployeeType,$Page1_TBJoinerEmployeeType,
                                $Page1_LbDefaultGroups,$Page1_DefaultGroupList,
                                $Page1_BtnAddGroup,$Page1_TBNewGroup,$Page1_BtnRemoveGroup,
                                $Page1_LbJoinerFullName, $Page1_TBJoinerFullName, 
                                $Page1_LbJoinerDisplayName, $Page1_TBJoinerDisplayName,
                                $Page1_LbEmail, $Page1_TBEmail, 
                                $Page1_LbJoinerOffice,$Page1_OfficeList,
                                $Page1_LbJoinerLogonScript,$Page1_TBJoinerLogonScript,
                                $Page1_LbJoinerTicket,$Page1_TBJoinerTicket,
                                $Page1_LbJoinerOU,$Page1_TBJoinerOU,
                                $Page1_LbInitialPassword,$Page1_TBInitialPassword,$Page1_BtnGeneratePassword,
                                $Page1_BtnGo))


$Page1_RBNewADAccount.Checked=$true
$Page1_RBFinalizeADAccount.Checked=$true


$TabControl1.Controls.Add($tabPage1)
$TabControl1.Location = New-Object System.Drawing.Point(0,0)
$TabControl1.Size=$PowerShellForms.Size

$PowerShellForms.Controls.Add($TabControl1)
$PowerShellForms.ShowDialog()
#EndRegion GUI
