<#
Author: Kin.Yan@wsp.com
Description: This is the most powerful tool for you to collect PC&User information from AD, if you can find a better one, please delete me :)

Latest update:
Ver 1.56789 (2024/5/17)
1. There is an option(enabled by default) to convert timestamp value to date value(only support AccountExpires,LastLogonTimeStamp,pwdlastSet)

Ver 1.45678 (2024/3/22)
1. Add 2 more object types: Security group & Distribution list. Amazing!

Ver 1.34567 (2024/2/27)
1. You are able to generate report for a specific list of PCs or users listed in a CSV file, amazing!

Ver 1.23456789 (2023/5/26)
1. Initial release
#>

$ShellVersion= "1.56789"
$FormTitle = "AD Reaper - v" + $ShellVersion + " - Kin.Yan@wsp.com"
$PCDefaultProperty ="extensionAttribute9,Name,OperatingSystemVersion,OperatingSystem,lastlogondate,whencreated,enabled,distinguishedName"
$UserDefaultProperty ="extensionAttribute7,extensionAttribute9,SamAccountName,displayName,mail,title,department,extensionAttribute12,Enabled,employeeType,whenCreated,lastLogonDate,scriptPath,distinguishedName"
$SecurityGroupDefaultProperty = "sAMAccountName,mail,cn,managedBy,description,displayName,groupType,whenCreated"
$DistributionListDefaultProperty = "extensionAttribute9,sAMAccountName,mail,cn,managedBy,description,displayName,groupType,whenCreated"

$WorkstationOUs=@(
	("CN","OU=Clients,OU=CN,OU=WSPObjects,DC=corp,DC=pbwan,DC=net"),
	("HK","OU=Clients,OU=HK,OU=WSPObjects,DC=corp,DC=pbwan,DC=net"),
	("SG","OU=Clients,OU=SG,OU=WSPObjects,DC=corp,DC=pbwan,DC=net"),
	("TW","OU=Clients,OU=TW,OU=WSPObjects,DC=corp,DC=pbwan,DC=net"),
	("TH","OU=Clients,OU=TH,OU=WSPObjects,DC=corp,DC=pbwan,DC=net"),
	("KR","OU=Clients,OU=KR,OU=WSPObjects,DC=corp,DC=pbwan,DC=net"),
	("MY","OU=Clients,OU=MY,OU=WSPObjects,DC=corp,DC=pbwan,DC=net"),
	("PH","OU=Clients,OU=PH,OU=WSPObjects,DC=corp,DC=pbwan,DC=net"),
	("Customize","")
)

$UserOUs=@(
	("CN","OU=Users,OU=CN,OU=WSPObjects,DC=corp,DC=pbwan,DC=net"),
	("HK","OU=Users,OU=HK,OU=WSPObjects,DC=corp,DC=pbwan,DC=net"),
	("SG","OU=Users,OU=SG,OU=WSPObjects,DC=corp,DC=pbwan,DC=net"),
	("TW","OU=Users,OU=TW,OU=WSPObjects,DC=corp,DC=pbwan,DC=net"),
	("TH","OU=Users,OU=TH,OU=WSPObjects,DC=corp,DC=pbwan,DC=net"),
	("KR","OU=Users,OU=KR,OU=WSPObjects,DC=corp,DC=pbwan,DC=net"),
	("MY","OU=Users,OU=MY,OU=WSPObjects,DC=corp,DC=pbwan,DC=net"),
	("PH","OU=Users,OU=PH,OU=WSPObjects,DC=corp,DC=pbwan,DC=net"),
	("Customize","")
	)

$SecurityGroupOUs=@(
	("CN","OU=Security,OU=Groups,OU=CN,OU=WSPObjects,DC=corp,DC=pbwan,DC=net"),
	("HK","OU=Security,OU=Groups,OU=HK,OU=WSPObjects,DC=corp,DC=pbwan,DC=net"),
	("SG","OU=Security,OU=Groups,OU=SG,OU=WSPObjects,DC=corp,DC=pbwan,DC=net"),
	("TW","OU=Security,OU=Groups,OU=TW,OU=WSPObjects,DC=corp,DC=pbwan,DC=net"),
	("TH","OU=Security,OU=Groups,OU=TH,OU=WSPObjects,DC=corp,DC=pbwan,DC=net"),
	("KR","OU=Security,OU=Groups,OU=KR,OU=WSPObjects,DC=corp,DC=pbwan,DC=net"),
	("MY","OU=Security,OU=Groups,OU=MY,OU=WSPObjects,DC=corp,DC=pbwan,DC=net"),
	("PH","OU=Security,OU=Groups,OU=PH,OU=WSPObjects,DC=corp,DC=pbwan,DC=net"),
	("Customize","")
	)

$DistributionListOUs=@(
	("CN","OU=Messaging,OU=Groups,OU=CN,OU=WSPObjects,DC=corp,DC=pbwan,DC=net"),
	("HK","OU=Messaging,OU=Groups,OU=HK,OU=WSPObjects,DC=corp,DC=pbwan,DC=net"),
	("SG","OU=Messaging,OU=Groups,OU=SG,OU=WSPObjects,DC=corp,DC=pbwan,DC=net"),
	("TW","OU=Messaging,OU=Groups,OU=TW,OU=WSPObjects,DC=corp,DC=pbwan,DC=net"),
	("TH","OU=Messaging,OU=Groups,OU=TH,OU=WSPObjects,DC=corp,DC=pbwan,DC=net"),
	("KR","OU=Messaging,OU=Groups,OU=KR,OU=WSPObjects,DC=corp,DC=pbwan,DC=net"),
	("MY","OU=Messaging,OU=Groups,OU=MY,OU=WSPObjects,DC=corp,DC=pbwan,DC=net"),
	("PH","OU=Messaging,OU=Groups,OU=PH,OU=WSPObjects,DC=corp,DC=pbwan,DC=net"),
	("Customize","")
	)

<#
param([Parameter(Mandatory)]$Scope
	)
#>


#### create GUI#######
Add-Type -AssemblyName System.Windows.Forms
$PowerShellForms = New-Object system.Windows.Forms.Form
$PowerShellForms.Text= $FormTitle
$PowerShellForms.Size = New-Object System.Drawing.Size(470,620)
$PowerShellForms.MinimizeBox = $False
$PowerShellForms.MaximizeBox = $False
$PowerShellForms.FormBorderStyle = [System.Windows.Forms.FormBorderStyle]::FixedDialog

$PowerShellForms.SizeGripStyle = "Hide"
#$Icons = [system.drawing.icon]::ExtractAssociatedIcon($PSHOME + "\powershell.exe")
$PowerShellForms.Icon = $Icons
$PowerShellForms.StartPosition = "CenterScreen"
$PowerShellForms.Font = New-Object System.Drawing.Font("Tahoma",10,[System.Drawing.FontStyle]::Regular)

$TabControl1 = New-Object System.Windows.Forms.TabControl




########################tabPage Start#######################

$tabADObject=New-Object System.Windows.Forms.TabPage
$tabADObject.text = "AD Object"
$tabADObject.TabIndex=0

$Page1_tbCustomizeOUPath = New-Object System.Windows.Forms.Textbox
$Page1_tbCustomizeOUPath.Size = New-Object System.Drawing.Size(290,40)
$Page1_tbCustomizeOUPath.AutoSize = $True
$Page1_tbCustomizeOUPath.Font = New-Object System.Drawing.Font("Tahoma",10,[System.Drawing.FontStyle]::Regular)
$Page1_tbCustomizeOUPath.ForeColor = "Black"
$Page1_tbCustomizeOUPath.Location = New-Object System.Drawing.Point(10,115)
$Page1_tbCustomizeOUPath.TabIndex = 10
$Page1_tbCustomizeOUPath.Enabled=$False
$tabADObject.Controls.AddRange($Page1_tbCustomizeOUPath)

$Page1_CbCustomizeOU = New-Object System.Windows.Forms.Checkbox
$Page1_CbCustomizeOU.Text = "Customize OU (distinguishedName)"
$Page1_CbCustomizeOU.AutoSize = $True
$Page1_CbCustomizeOU.BackColor = "Transparent"
$Page1_CbCustomizeOU.Font = New-Object System.Drawing.Font("Tahoma",10,[System.Drawing.FontStyle]::Regular)
$Page1_CbCustomizeOU.ForeColor = "Black"
$Page1_CbCustomizeOU.Location = New-Object System.Drawing.Point(10,90)
$Page1_CbCustomizeOU.TabIndex = 7
$tabADObject.Controls.AddRange($Page1_CbCustomizeOU)
$Page1_CbCustomizeOU.Add_CheckedChanged(
    {
        if($Page1_CbCustomizeOU.Checked -eq $True)
        {
            $Page1_CbAsia.Checked=$false
            $Page1_CbCN.Checked=$false
            $Page1_CbHK.Checked=$false
            $Page1_CbSG.Checked=$false
            $Page1_CbTW.Checked=$false
            $Page1_CbTH.Checked=$false
            $Page1_CbKR.Checked=$false
            $Page1_CbMY.Checked=$false
            $Page1_CbPH.Checked=$false
            $Page1_tbImportFromCSVPath.Enabled=$false
            $Page1_CbImportFromCSV.Checked=$false
            $Page1_tbCustomizeOUPath.Enabled=$true
        }
        else
        {
            $Page1_tbCustomizeOUPath.Enabled=$false
        }
    }
)

$Page1_GBObjectType = New-Object System.Windows.Forms.GroupBox
$Page1_GBObjectType.Location = New-Object System.Drawing.Point(110,5)
$Page1_GBObjectType.Size = New-Object System.Drawing.Size(310,80)
$Page1_GBObjectType.Text = "Object type"

$Page1_RBPC = New-Object System.Windows.Forms.RadioButton
$Page1_RBPC.AutoSize = $True
$Page1_RBPC.Location = New-Object System.Drawing.Point(10,20)
$Page1_RBPC.Text = "PC"
$Page1_RBPC.Checked = $True
$Page1_GBObjectType.Controls.AddRange($Page1_RBPC)
$Page1_RBPC.Add_CheckedChanged(
    {
        if($Page1_RBPC.Checked -eq $true)
        {
            $Page1_tbADProperty.Text = $PCDefaultProperty
            $Page1_lbCSVDescription.Text = "Format: each PC name a line, no header is required"
        }
    }
)


$Page1_RBUser = New-Object System.Windows.Forms.RadioButton
$Page1_RBUser.AutoSize = $True
$Page1_RBUser.Location = New-Object System.Drawing.Point(150,20)
$Page1_RBUser.Text = "User"
$Page1_GBObjectType.Controls.AddRange($Page1_RBUser)
$Page1_RBUser.Add_CheckedChanged(
    {
        if($Page1_RBUser.Checked -eq $true)
        {
            $Page1_tbADProperty.Text = $UserDefaultProperty
            $Page1_lbCSVDescription.Text = "Format: each sAMAccountName or email address a line, no header is required"
        }
    }
)


$Page1_RBSG = New-Object System.Windows.Forms.RadioButton
$Page1_RBSG.AutoSize = $True
$Page1_RBSG.Location = New-Object System.Drawing.Point(10,50)
$Page1_RBSG.Text = "Security Group"
$Page1_GBObjectType.Controls.AddRange($Page1_RBSG)
$Page1_RBSG.Add_CheckedChanged(
    {
        if($Page1_RBSG.Checked -eq $true)
        {
            $Page1_tbADProperty.Text = $SecurityGroupDefaultProperty
            $Page1_lbCSVDescription.Text = "Format: each sAMAccountName a line, no header is required"
        }
    }
)


$Page1_RBDL = New-Object System.Windows.Forms.RadioButton
$Page1_RBDL.AutoSize = $True
$Page1_RBDL.Location = New-Object System.Drawing.Point(150,50)
$Page1_RBDL.Text = "Distribution List"
$Page1_GBObjectType.Controls.AddRange($Page1_RBDL)
$Page1_RBDL.Add_CheckedChanged(
    {
        if($Page1_RBDL.Checked -eq $true)
        {
            $Page1_tbADProperty.Text = $DistributionListDefaultProperty
            $Page1_lbCSVDescription.Text = "Format: each sAMAccountName or email address a line, no header is required"
        }
    }
)


$tabADObject.Controls.AddRange($Page1_GBObjectType)


$Page1_GBScope = New-Object System.Windows.Forms.GroupBox
$Page1_GBScope.Location = New-Object System.Drawing.Point(120,115)
$Page1_GBScope.Size = New-Object System.Drawing.Size(310,235)
$Page1_GBScope.Text = "Scope"
$Page1_GBScope.Controls.AddRange($Page1_CbCustomizeOU)
$Page1_GBScope.Controls.AddRange($Page1_tbCustomizeOUPath)

$Page1_CbAsia	= New-Object System.Windows.Forms.Checkbox
$Page1_CbAsia.Text = "Asia Offices"
$Page1_CbAsia.AutoSize = $True
$Page1_CbAsia.BackColor = "Transparent"
$Page1_CbAsia.Font = New-Object System.Drawing.Font("Tahoma",10,[System.Drawing.FontStyle]::Regular)
$Page1_CbAsia.ForeColor = "Black"
$Page1_CbAsia.Location = New-Object System.Drawing.Point(10,25)
$Page1_CbAsia.TabIndex = 1
$Page1_GBScope.Controls.AddRange($Page1_CbAsia)
#$tabADObject.Controls.AddRange($Page1_CbAsia)
$Page1_CbAsia.Add_CheckedChanged(
    {
        #[System.Windows.Forms.Messagebox]::Show($Page1_CbAsia.Checked)
        if($Page1_CbAsia.Checked -eq $true)
        {
            $Page1_CbCN.Checked=$true
            $Page1_CbHK.Checked=$true
            $Page1_CbSG.Checked=$true
            $Page1_CbTW.Checked=$true
            $Page1_CbTH.Checked=$true
            $Page1_CbKR.Checked=$true
            $Page1_CbMY.Checked=$true
            $Page1_CbPH.Checked=$true
            $Page1_CbCustomizeOU.Checked=$false
            $Page1_tbCustomizeOUPath.Enabled=$false
            $Page1_CbImportFromCSV.Checked=$false
            $Page1_tbImportFromCSVPath.Enabled=$false
            #$Page1_CbImportFromCSV.Checked=$false
        }
        else
        {
            $Page1_tbCustomizeOUPath.Enabled=$false
            $Page1_tbImportFromCSVPath.Enabled=$false
        }
    }
)


$Page1_CbCN	= New-Object System.Windows.Forms.Checkbox
$Page1_CbCN.Text = "CN"
$Page1_CbCN.AutoSize = $True
$Page1_CbCN.BackColor = "Transparent"
$Page1_CbCN.Font = New-Object System.Drawing.Font("Tahoma",10,[System.Drawing.FontStyle]::Regular)
$Page1_CbCN.ForeColor = "Black"
$Page1_CbCN.Location = New-Object System.Drawing.Point(30,45)
$Page1_CbCN.TabIndex = 2
#$tabADObject.Controls.AddRange($Page1_CbCN)
$Page1_GBScope.Controls.AddRange($Page1_CbCN)
$Page1_CbCN.Add_CheckedChanged(
    {
            if($Page1_CbCN.Checked -eq $false)
            {
                $Page1_CbAsia.Checked=$false
                $Page1_tbImportFromCSVPath.Enabled=$false
                $Page1_tbCustomizeOUPath.Enabled=$false
            }
            else
            {
                $Page1_CbCustomizeOU.Checked=$false
                $Page1_tbCustomizeOUPath.Enabled=$false
                $Page1_CbImportFromCSV.Checked=$false
                $Page1_tbImportFromCSVPath.Enabled=$false
            }
    }
)

$Page1_CbHK	= New-Object System.Windows.Forms.Checkbox
$Page1_CbHK.Text = "HK"
$Page1_CbHK.AutoSize = $True
$Page1_CbHK.BackColor = "Transparent"
$Page1_CbHK.Font = New-Object System.Drawing.Font("Tahoma",10,[System.Drawing.FontStyle]::Regular)
$Page1_CbHK.ForeColor = "Black"
$Page1_CbHK.Location = New-Object System.Drawing.Point(30,65)
$Page1_CbHK.TabIndex = 3
$Page1_GBScope.Controls.AddRange($Page1_CbHK)
#$tabADObject.Controls.AddRange($Page1_CbHK)
$Page1_CbHK.Add_CheckedChanged(
    {
            if($Page1_CbHK.Checked -eq $false)
            {
            $Page1_CbAsia.Checked=$false
            $Page1_tbImportFromCSVPath.Enabled=$false
            $Page1_tbCustomizeOUPath.Enabled=$false
            }
            else
            {
                $Page1_CbCustomizeOU.Checked=$false
                $Page1_tbCustomizeOUPath.Enabled=$false
                $Page1_CbImportFromCSV.Checked=$false
                $Page1_tbImportFromCSVPath.Enabled=$false
            }
    }
)

$Page1_CbSG	= New-Object System.Windows.Forms.Checkbox
$Page1_CbSG.Text = "SG"
$Page1_CbSG.AutoSize = $True
$Page1_CbSG.BackColor = "Transparent"
$Page1_CbSG.Font = New-Object System.Drawing.Font("Tahoma",10,[System.Drawing.FontStyle]::Regular)
$Page1_CbSG.ForeColor = "Black"
$Page1_CbSG.Location = New-Object System.Drawing.Point(100,45)
$Page1_CbSG.TabIndex = 4
$Page1_GBScope.Controls.AddRange($Page1_CbSG)
#$tabADObject.Controls.AddRange($Page1_CbSG)
$Page1_CbSG.Add_CheckedChanged(
    {
            if($Page1_CbSG.Checked -eq $false)
            {
            $Page1_CbAsia.Checked=$false
            $Page1_tbImportFromCSVPath.Enabled=$false
            $Page1_tbCustomizeOUPath.Enabled=$false
            }
            else
            {
                $Page1_CbCustomizeOU.Checked=$false
                $Page1_tbCustomizeOUPath.Enabled=$false
                $Page1_CbImportFromCSV.Checked=$false
                $Page1_tbImportFromCSVPath.Enabled=$false
            }
    }
)

$Page1_CbTW	= New-Object System.Windows.Forms.Checkbox
$Page1_CbTW.Text = "TW"
$Page1_CbTW.AutoSize = $True
$Page1_CbTW.BackColor = "Transparent"
$Page1_CbTW.Font = New-Object System.Drawing.Font("Tahoma",10,[System.Drawing.FontStyle]::Regular)
$Page1_CbTW.ForeColor = "Black"
$Page1_CbTW.Location = New-Object System.Drawing.Point(100,65)
$Page1_CbTW.TabIndex = 5
$Page1_GBScope.Controls.AddRange($Page1_CbTW)
#$tabADObject.Controls.AddRange($Page1_CbTW)
$Page1_CbTW.Add_CheckedChanged(
    {
            if($Page1_CbTW.Checked -eq $false)
            {
            $Page1_CbAsia.Checked=$false
            $Page1_tbImportFromCSVPath.Enabled=$false
            $Page1_tbCustomizeOUPath.Enabled=$false
            }
            else
            {
                $Page1_CbCustomizeOU.Checked=$false
                $Page1_tbCustomizeOUPath.Enabled=$false
                $Page1_CbImportFromCSV.Checked=$false
                $Page1_tbImportFromCSVPath.Enabled=$false
            }
    }
)

$Page1_CbTH	= New-Object System.Windows.Forms.Checkbox
$Page1_CbTH.Text = "TH"
$Page1_CbTH.AutoSize = $True
$Page1_CbTH.BackColor = "Transparent"
$Page1_CbTH.Font = New-Object System.Drawing.Font("Tahoma",10,[System.Drawing.FontStyle]::Regular)
$Page1_CbTH.ForeColor = "Black"
$Page1_CbTH.Location = New-Object System.Drawing.Point(170,45)
$Page1_CbTH.TabIndex = 6
$Page1_GBScope.Controls.AddRange($Page1_CbTH)
#$tabADObject.Controls.AddRange($Page1_CbTH)
$Page1_CbTH.Add_CheckedChanged(
    {
            if($Page1_CbTH.Checked -eq $false)
            {
            $Page1_CbAsia.Checked=$false
            $Page1_tbImportFromCSVPath.Enabled=$false
            $Page1_tbCustomizeOUPath.Enabled=$false
            }
            else
            {
                $Page1_CbCustomizeOU.Checked=$false
                $Page1_tbCustomizeOUPath.Enabled=$false
                $Page1_CbImportFromCSV.Checked=$false
                $Page1_tbImportFromCSVPath.Enabled=$false
            }
    }
)

$Page1_CbKR	= New-Object System.Windows.Forms.Checkbox
$Page1_CbKR.Text = "KR"
$Page1_CbKR.AutoSize = $True
$Page1_CbKR.BackColor = "Transparent"
$Page1_CbKR.Font = New-Object System.Drawing.Font("Tahoma",10,[System.Drawing.FontStyle]::Regular)
$Page1_CbKR.ForeColor = "Black"
$Page1_CbKR.Location = New-Object System.Drawing.Point(170,65)
$Page1_CbKR.TabIndex = 7
$Page1_GBScope.Controls.AddRange($Page1_CbKR)
#$tabADObject.Controls.AddRange($Page1_CbKR)
$Page1_CbKR.Add_CheckedChanged(
    {
            if($Page1_CbKR.Checked -eq $false)
            {
            $Page1_CbAsia.Checked=$false
            $Page1_tbImportFromCSVPath.Enabled=$false
            $Page1_tbCustomizeOUPath.Enabled=$false
            }
            else
            {
                $Page1_CbCustomizeOU.Checked=$false
                $Page1_tbCustomizeOUPath.Enabled=$false
                $Page1_CbImportFromCSV.Checked=$false
                $Page1_tbImportFromCSVPath.Enabled=$false
            }
    }
)

$Page1_CbMY	= New-Object System.Windows.Forms.Checkbox
$Page1_CbMY.Text = "MY"
$Page1_CbMY.AutoSize = $True
$Page1_CbMY.BackColor = "Transparent"
$Page1_CbMY.Font = New-Object System.Drawing.Font("Tahoma",10,[System.Drawing.FontStyle]::Regular)
$Page1_CbMY.ForeColor = "Black"
$Page1_CbMY.Location = New-Object System.Drawing.Point(250,45)
$Page1_CbMY.TabIndex = 8
$Page1_GBScope.Controls.AddRange($Page1_CbMY)
#$tabADObject.Controls.AddRange($Page1_CbMY)
$Page1_CbMY.Add_CheckedChanged(
    {
            if($Page1_CbMY.Checked -eq $false)
            {
            $Page1_CbAsia.Checked=$false
            $Page1_tbImportFromCSVPath.Enabled=$false
            $Page1_tbCustomizeOUPath.Enabled=$false
            }
            else
            {
                $Page1_CbCustomizeOU.Checked=$false
                $Page1_tbCustomizeOUPath.Enabled=$false
                $Page1_CbImportFromCSV.Checked=$false
                $Page1_tbImportFromCSVPath.Enabled=$false
            }
    }
)

$Page1_CbPH	= New-Object System.Windows.Forms.Checkbox
$Page1_CbPH.Text = "PH"
$Page1_CbPH.AutoSize = $True
$Page1_CbPH.BackColor = "Transparent"
$Page1_CbPH.Font = New-Object System.Drawing.Font("Tahoma",10,[System.Drawing.FontStyle]::Regular)
$Page1_CbPH.ForeColor = "Black"
$Page1_CbPH.Location = New-Object System.Drawing.Point(250,65)
$Page1_CbPH.TabIndex = 9
$Page1_GBScope.Controls.AddRange($Page1_CbPH)
#$tabADObject.Controls.AddRange($Page1_CbPH)
$Page1_CbPH.Add_CheckedChanged(
    {
            if($Page1_CbPH.Checked -eq $false)
            {
            $Page1_CbAsia.Checked=$false
            $Page1_tbImportFromCSVPath.Enabled=$false
            $Page1_tbCustomizeOUPath.Enabled=$false
            }
            else
            {
                $Page1_CbCustomizeOU.Checked=$false
                $Page1_tbCustomizeOUPath.Enabled=$false
                $Page1_CbImportFromCSV.Checked=$false
                $Page1_tbImportFromCSVPath.Enabled=$false
            }
    }
)

$Page1_CbImportFromCSV	= New-Object System.Windows.Forms.Checkbox
$Page1_CbImportFromCSV.Text = "From CSV List"
$Page1_CbImportFromCSV.AutoSize = $True
$Page1_CbImportFromCSV.BackColor = "Transparent"
$Page1_CbImportFromCSV.Font = New-Object System.Drawing.Font("Tahoma",10,[System.Drawing.FontStyle]::Regular)
$Page1_CbImportFromCSV.ForeColor = "Black"
$Page1_CbImportFromCSV.Location = New-Object System.Drawing.Point(10,145)
$Page1_CbImportFromCSV.TabIndex = 9
$Page1_GBScope.Controls.AddRange($Page1_CbImportFromCSV)
$Page1_CbImportFromCSV.Add_CheckedChanged(
    {
            if($Page1_CbImportFromCSV.Checked -eq $true)
            {
                $Page1_CbAsia.Checked=$false
                $Page1_CbCN.Checked=$false
                $Page1_CbHK.Checked=$false
                $Page1_CbSG.Checked=$false
                $Page1_CbTW.Checked=$false
                $Page1_CbTH.Checked=$false
                $Page1_CbKR.Checked=$false
                $Page1_CbMY.Checked=$false
                $Page1_CbPH.Checked=$false
                $Page1_tbImportFromCSVPath.Enabled=$true
                $Page1_CbCustomizeOU.Checked=$false
                $Page1_tbCustomizeOUPath.Enabled=$false
            }
    }
)

$Page1_tbImportFromCSVPath	= New-Object System.Windows.Forms.Textbox
$Page1_tbImportFromCSVPath.Size = New-Object System.Drawing.Size(210,40)
$Page1_tbImportFromCSVPath.AutoSize = $True
$Page1_tbImportFromCSVPath.Font = New-Object System.Drawing.Font("Tahoma",10,[System.Drawing.FontStyle]::Regular)
$Page1_tbImportFromCSVPath.ForeColor = "Black"
$Page1_tbImportFromCSVPath.Location = New-Object System.Drawing.Point(10,170)
$Page1_tbImportFromCSVPath.TabIndex = 10
$Page1_tbImportFromCSVPath.Enabled=$False
#$Page1_tbImportFromCSVPath.Text="C:\Support\ADReaper\NAUsers.csv"
$Page1_GBScope.Controls.AddRange($Page1_tbImportFromCSVPath)

$Page1_lbCSVDescription	= New-Object System.Windows.Forms.Label
$Page1_lbCSVDescription.Text = "Format: each PC name a line, no header is required"
#$Page1_lbADProperty.AutoSize = $True
$Page1_lbCSVDescription.Size = New-Object System.Drawing.Size(280,30)
$Page1_lbCSVDescription.BackColor = "Transparent"
$Page1_lbCSVDescription.Font = New-Object System.Drawing.Font("Tahoma",8,[System.Drawing.FontStyle]::Italic)
$Page1_lbCSVDescription.ForeColor = "Black"
$Page1_lbCSVDescription.Location = New-Object System.Drawing.Point(10,200)
$Page1_lbCSVDescription.TabIndex = 10
$Page1_GBScope.Controls.AddRange($Page1_lbCSVDescription)

$Page1_BtnSelectCSVFile = New-Object System.Windows.Forms.Button
$Page1_BtnSelectCSVFile.Size = New-Object System.Drawing.Size(70,25)
$Page1_BtnSelectCSVFile.Location = New-Object System.Drawing.Point(230,170)
$Page1_BtnSelectCSVFile.Text = "CSV File"
$Page1_BtnSelectCSVFile.Font = New-Object System.Drawing.Font("Tahoma",10,[System.Drawing.FontStyle]::Regular)
$Page1_BtnSelectCSVFile.Add_Click(
    {
        $CurrentLocation = Get-Location
        $CSVBrowser = New-Object System.Windows.Forms.OpenFileDialog
        $CSVBrowser.InitialDirectory = $CurrentLocation
        $CSVBrowser.Filter = "CSV (*.csv)|*.csv"
        $CSVBrowser.ShowDialog() | Out-Null
        if($CSVBrowser.FileName -ne $null)
        {
            $Page1_tbImportFromCSVPath.text = $CSVBrowser.FileName
        }
    }
)
$Page1_GBScope.Controls.AddRange($Page1_BtnSelectCSVFile)

$Page1_CbAsia.Checked=$true
$PowerShellForms.Controls.Add($Page1_GBScope)


$Page1_lbADProperty	= New-Object System.Windows.Forms.Label
$Page1_lbADProperty.Text = "Report column`n(AD property)"
#$Page1_lbADProperty.AutoSize = $True
$Page1_lbADProperty.Size = New-Object System.Drawing.Size(100,130)
$Page1_lbADProperty.BackColor = "Transparent"
$Page1_lbADProperty.Font = New-Object System.Drawing.Font("Tahoma",10,[System.Drawing.FontStyle]::Regular)
$Page1_lbADProperty.ForeColor = "Black"
$Page1_lbADProperty.Location = New-Object System.Drawing.Point(5,335)
$Page1_lbADProperty.TabIndex = 10
$tabADObject.Controls.AddRange($Page1_lbADProperty)

$Page1_tbADProperty	= New-Object System.Windows.Forms.Textbox
$Page1_tbADProperty.Text = $PCDefaultProperty
$Page1_tbADProperty.Size = New-Object System.Drawing.Size(310,80)
$Page1_tbADProperty.AutoSize = $True
$Page1_tbADProperty.Multiline=$True
$Page1_tbADProperty.Font = New-Object System.Drawing.Font("Tahoma",10,[System.Drawing.FontStyle]::Regular)
$Page1_tbADProperty.ForeColor = "Black"
$Page1_tbADProperty.Location = New-Object System.Drawing.Point(110,335)
$Page1_tbADProperty.TabIndex = 10
$tabADObject.Controls.AddRange($Page1_tbADProperty)


$lbReportPath	= New-Object System.Windows.Forms.Label
$lbReportPath.Text = "Report Path"
#$lbReportPath.Size = New-Object System.Drawing.Size(300,50)
$lbReportPath.AutoSize = $True
$lbReportPath.BackColor = "Transparent"
$lbReportPath.Font = New-Object System.Drawing.Font("Tahoma",10,[System.Drawing.FontStyle]::Regular)
$lbReportPath.ForeColor = "Black"
$lbReportPath.Location = New-Object System.Drawing.Point(15,485)
$lbReportPath.TabIndex = 10
$PowerShellForms.Controls.Add($lbReportPath)

$tbReportPath	= New-Object System.Windows.Forms.Textbox
$tbReportPath.Text = $PSScriptRoot
$tbReportPath.Size = New-Object System.Drawing.Size(310,50)
$tbReportPath.AutoSize = $True
$tbReportPath.Font = New-Object System.Drawing.Font("Tahoma",10,[System.Drawing.FontStyle]::Regular)
$tbReportPath.ForeColor = "Black"
$tbReportPath.Location = New-Object System.Drawing.Point(120,485)
$tbReportPath.TabIndex = 10
$PowerShellForms.Controls.Add($tbReportPath)

$BtnGenerateReport = New-Object System.Windows.Forms.Button
$BtnGenerateReport.Size = New-Object System.Drawing.Size(140,40)
$BtnGenerateReport.Location = New-Object System.Drawing.Point(290,520)
$BtnGenerateReport.Text = "Show me the report"
$BtnGenerateReport.Font = New-Object System.Drawing.Font("Tahoma",10,[System.Drawing.FontStyle]::Regular)
$BtnGenerateReport.Add_Click(
    {
        $selectedoffice=$false
        if($Page1_CbAsia.Checked -eq $true) {$selectedoffice=$true}
        if($Page1_CbCN.Checked -eq $true) {$selectedoffice=$true}
        if($Page1_CbHK.Checked -eq $true) {$selectedoffice=$true}
        if($Page1_CbSG.Checked -eq $true) {$selectedoffice=$true}
        if($Page1_CbTW.Checked -eq $true) {$selectedoffice=$true}
        if($Page1_CbTH.Checked -eq $true) {$selectedoffice=$true}
        if($Page1_CbKR.Checked -eq $true) {$selectedoffice=$true}
        if($Page1_CbMY.Checked -eq $true) {$selectedoffice=$true}
        if($Page1_CbPH.Checked -eq $true) {$selectedoffice=$true}
        if((($selectedoffice -eq $false) -and (($Page1_CbCustomizeOU.Checked -eq $false) -or ($Page1_tbCustomizeOUPath.Text.Length -lt 10))) -and (($Page1_CbImportFromCSV.Checked -eq $false) -or ($Page1_tbImportFromCSVPath.Text.Length -lt 8)))
        {
            [System.Windows.Forms.Messagebox]::Show("Please choose a OU or specified a CSV file!")
        }
        else
        {
            GenerateList
        }
    }
)
$PowerShellForms.Controls.Add($BtnGenerateReport)


$Page1_CbConvertTimeStampToDate = New-Object System.Windows.Forms.Checkbox
$Page1_CbConvertTimeStampToDate.Text = "Convert Timestamp value to Date value"
$Page1_CbConvertTimeStampToDate.AutoSize = $True
$Page1_CbConvertTimeStampToDate.BackColor = "Transparent"
$Page1_CbConvertTimeStampToDate.Font = New-Object System.Drawing.Font("Tahoma",10,[System.Drawing.FontStyle]::Regular)
$Page1_CbConvertTimeStampToDate.ForeColor = "Black"
$Page1_CbConvertTimeStampToDate.Location = New-Object System.Drawing.Point(170,420)
$Page1_CbConvertTimeStampToDate.TabIndex = 7
$Page1_CbConvertTimeStampToDate.Checked=$true
$tabADObject.Controls.AddRange($Page1_CbConvertTimeStampToDate)



function GenerateList
{
	Import-Module ActiveDirectory
	$today=Get-Date -Format "yyyyMMddHHmm"
    $ScopeArrayList = [System.Collections.ArrayList]::new()

    if($Page1_CbCN.Checked -eq $true) {$ScopeArrayList.Add("CN")}
    if($Page1_CbHK.Checked -eq $true) {$ScopeArrayList.Add("HK")}
    if($Page1_CbSG.Checked -eq $true) {$ScopeArrayList.Add("SG")}
    if($Page1_CbTW.Checked -eq $true) {$ScopeArrayList.Add("TW")}
    if($Page1_CbTH.Checked -eq $true) {$ScopeArrayList.Add("TH")}
    if($Page1_CbKR.Checked -eq $true) {$ScopeArrayList.Add("KR")}
    if($Page1_CbMY.Checked -eq $true) {$ScopeArrayList.Add("MY")}
    if($Page1_CbPH.Checked -eq $true) {$ScopeArrayList.Add("PH")}
    if($Page1_CbCustomizeOU.Checked -eq $true) {$ScopeArrayList.Add("Customize")}

    $filename=""
    foreach($s in $ScopeArrayList)
    {
        $filename=$filename + $s
    }

    $ReportColumn=$Page1_tbADProperty.Text
    if($Page1_CbConvertTimeStampToDate.Checked -eq $true) 
    {
        $ReportColumn=$ReportColumn -Replace "accountexpires","AccountExpirationDate"
        $ReportColumn=$ReportColumn -Replace "lastLogonTimestamp","LastLogonDate"
        $ReportColumn=$ReportColumn -Replace "pwdLastSet","PasswordLastSet"
    }
    #$ReportColumn=$ReportColumn.Replace("lastLogonTimestamp","@{Label='lastLogonTimestamp';Expression={[datetime]::FromFileTime($_.lastLogonTimestamp)}}")
    $ProgressBar.Value=0
    $i=0

    if($Page1_RBPC.Checked -eq $true) #PC
    {

        if($Page1_CbImportFromCSV.Checked -eq $true)
        {
            $filename=$tbReportPath.Text + "\PCList_" + $today + ".csv"
            $List = import-csv -path $Page1_tbImportFromCSVPath.Text -header "Name"
            $progresscount = $List.Count
            ForEach($PCList in $List){
                $i++
                [int]$percentage=($i/$progresscount)*100
                $ProgressBar.Value=$percentage
                #$PCList = """" + $PCList.Name + """"
                $cmdline = "Get-ADComputer -identity """ + $PCList.Name + """ -Properties " + $ReportColumn + " | select " + $ReportColumn + " | Export-Csv -Encoding UTF8 -NoTypeInformation -append -Force -path " + $filename
                Invoke-Expression -Command $cmdline
                }
        }
        else {
            $filename=$tbReportPath.Text + "\PCList_" + $filename + "_" + $today + ".csv"
            $WorkstationOUs[$WorkstationOUs.Count-1][1]=$Page1_tbCustomizeOUPath.Text
            $progresscount=$ScopeArrayList.Count*$WorkstationOUs.Count
            foreach($Scope in $ScopeArrayList)
            {
                foreach($WorkstationOU in $WorkstationOUs)
                {
                    $i++
                    [int]$percentage=($i/$progresscount)*100
                    $ProgressBar.Value=$percentage
                    if($Scope -eq $WorkstationOU[0])
                    {
                        $cmdline = "Get-ADComputer -SearchBase """ + $WorkstationOU[1] + """ -Filter * -Properties " + $ReportColumn + " | select " + $ReportColumn + " | Export-Csv -Encoding UTF8 -NoTypeInformation -append -Force -path " + $filename
                        Invoke-Expression -Command $cmdline
                    }
                }
            }
        }

    }
    elseif($Page1_RBUser.Checked -eq $true) #User
    {
        #$ReportColumn=$ReportColumn.ToLower()
        #$SelectColumn=$ReportColumn -Replace "accountExpires","@{Label='accountExpires';Expression={[datetime]::FromFileTime($_.accountExpires)}}"
        #$SelectColumn=$ReportColumn -Replace "accountExpires","@{Label='accountExpires';Expression={accountExpires}}"
        if($Page1_CbImportFromCSV.Checked -eq $true)
        {
            $filename=$tbReportPath.Text + "\UserList_" + $today + ".csv"
            $List = import-csv -path $Page1_tbImportFromCSVPath.Text -header "Name"
            $progresscount = $List.Count
		    
            ForEach($UserList in $List){
                $i++
                [int]$percentage=($i/$progresscount)*100
                $ProgressBar.Value=$percentage

                if($UserList.Name.contains("@") -eq $true)
                {
                    $cmdline = "Get-ADUser -filter {EmailAddress -eq """ + $UserList.Name + """} -Properties " + $ReportColumn + "| select " + $ReportColumn + "| export-csv -encoding UTF8 -NoTypeInformation -Force -Append -Path " + $filename
                    #$cmdline = "Get-ADUser -filter {EmailAddress -eq ""Bei-Chen.Zheng@wsp.com""} -Properties " + $ReportColumn + "| select " + $ReportColumn + "| export-csv -encoding UTF8 -NoTypeInformation -Force -Append -Path " + $filename
                }
                else {
                    $cmdline = "Get-ADUser -Identity """ + $UserList.Name + """ -Properties " + $ReportColumn + "| select " + $ReportColumn + "| export-csv -encoding UTF8 -NoTypeInformation -Force -Append -Path " + $filename
                }
                #Get-ADUser -filter {EmailAddress -eq $UserList} -Properties $ReportColumn | select $SelectColumn | export-csv -encoding UTF8 -NoTypeInformation -Force -Append -Path $filename
                Invoke-Expression -Command $cmdline
                }
        }
        else {
            $filename=$tbReportPath.Text + "\UserList_" + $filename + "_" + $today + ".csv"
            $UserOUs[$UserOUs.Count-1][1]=$Page1_tbCustomizeOUPath.Text
            $progresscount=$ScopeArrayList.Count*$UserOUs.Count
            foreach($Scope in $ScopeArrayList)
            {
                foreach($UserOU in $UserOUs)
                {
                    $i++
                    [int]$percentage=($i/$progresscount)*100
                    $ProgressBar.Value=$percentage
                    if($Scope -eq $UserOU[0])
                    {
                        
                        $cmdline = "Get-ADUser -SearchBase """ + $UserOU[1] + """ -Filter * -Properties " + $ReportColumn + " | select " + $ReportColumn + " | export-csv -encoding UTF8 -NoTypeInformation -Force -Append -Path " + $filename
                        #Get-ADUser -SearchBase $UserOU[1] -Filter * -Properties * | select $ReportColumn.Replace("accountExpires","@{Label='accountExpires';Expression={$_.accountExpires}}") | export-csv -encoding UTF8 -NoTypeInformation -Force -Append -Path $filename
                        Invoke-Expression -Command $cmdline

                    }
                }
            }
        }
    }
    elseif($Page1_RBSG.Checked -eq $true) #Security Group
    {
        if($Page1_CbImportFromCSV.Checked -eq $true)
        {
            $filename=$tbReportPath.Text + "\SGList_" + $today + ".csv"
            $List = import-csv -path $Page1_tbImportFromCSVPath.Text -header "Name"
            $progresscount = $List.Count
            ForEach($DL in $List){
                $i++
                [int]$percentage=($i/$progresscount)*100
                $ProgressBar.Value=$percentage
                $cmdline = "Get-ADObject -filter {(objectClass -eq ""group"") -and (samAccountName -eq """ + $DL.Name + """)} -property " + $ReportColumn + " | select-object " + $ReportColumn + " | export-csv -Encoding UTF8 -NoTypeInformation -Force -Append -Path " + $filename
                Invoke-Expression -Command $cmdline
                }
        }
        else {
            $filename=$tbReportPath.Text + "\SGList_" + $filename + "_" + $today + ".csv"
            $SecurityGroupOUs[$SecurityGroupOUs.Count-1][1]=$Page1_tbCustomizeOUPath.Text
            $progresscount=$ScopeArrayList.Count*$UserOUs.Count
            foreach($Scope in $ScopeArrayList)
            {
                foreach($SGOU in $SecurityGroupOUs)
                {
                    $i++
                    [int]$percentage=($i/$progresscount)*100
                    $ProgressBar.Value=$percentage
                    if($Scope -eq $SGOU[0])
                    {
                        #$cmdline = "Get-ADUser -SearchBase """ + $UserOU[1] + """ -Filter * -Properties " + $ReportColumn + " | select " + $ReportColumn + " | export-csv -encoding UTF8 -NoTypeInformation -Force -Append -Path " + $filename
                        $cmdline = "Get-ADObject -SearchBase """ + $SGOU[1] + """ -filter {(objectClass -eq ""group"")} -property " + $ReportColumn + " | select-object " + $ReportColumn + " | export-csv -Encoding UTF8 -NoTypeInformation -Force -Append -Path " + $filename
                
                        Invoke-Expression -Command $cmdline
                    }
                }
            }
        }
    }
    elseif($Page1_RBDL.Checked -eq $true) #Distribution List
    {
        if($Page1_CbImportFromCSV.Checked -eq $true)
        {
            $filename=$tbReportPath.Text + "\DLList_" + $today + ".csv"
            $List = import-csv -path $Page1_tbImportFromCSVPath.Text -header "Name"
            $progresscount = $List.Count
            ForEach($DL in $List){
                $i++
                [int]$percentage=($i/$progresscount)*100
                $ProgressBar.Value=$percentage
                if($DL.Name.contains("@") -eq $true) #email
                {
                    $cmdline = "Get-ADObject -filter {(objectClass -eq ""group"") -and (mail -eq """ + $DL.Name + """)} -property " + $ReportColumn + " | select-object " + $ReportColumn + " | export-csv -Encoding UTF8 -NoTypeInformation -Force -Append -Path " + $filename

                }
                else { #samAccountName
                    $cmdline = "Get-ADObject -filter {(objectClass -eq ""group"") -and (samAccountName -eq """ + $DL.Name + """)} -property " + $ReportColumn + " | select-object " + $ReportColumn + " | export-csv -Encoding UTF8 -NoTypeInformation -Force -Append -Path " + $filename
                }
                Invoke-Expression -Command $cmdline
                }
        }
        else {
            $filename=$tbReportPath.Text + "\DLList_" + $filename + "_" + $today + ".csv"
            $DistributionListOUs[$DistributionListOUs.Count-1][1]=$Page1_tbCustomizeOUPath.Text
            $progresscount=$ScopeArrayList.Count*$UserOUs.Count
            foreach($Scope in $ScopeArrayList)
            {
                foreach($DLOU in $DistributionListOUs)
                {
                    $i++
                    [int]$percentage=($i/$progresscount)*100
                    $ProgressBar.Value=$percentage
                    if($Scope -eq $DLOU[0])
                    {
                        #$cmdline = "Get-ADUser -SearchBase """ + $UserOU[1] + """ -Filter * -Properties " + $ReportColumn + " | select " + $ReportColumn + " | export-csv -encoding UTF8 -NoTypeInformation -Force -Append -Path " + $filename
                        $cmdline = "Get-ADObject -SearchBase """ + $DLOU[1] + """ -filter {(objectClass -eq ""group"")} -property " + $ReportColumn + " | select-object " + $ReportColumn + " | export-csv -Encoding UTF8 -NoTypeInformation -Force -Append -Path " + $filename
                
                        Invoke-Expression -Command $cmdline
                    }
                }
            }
        }
    }   
    [System.Windows.Forms.Messagebox]::Show("Done, go and get your report, no thanks ^_^")
    
}

########################tabPageGetPCList End#######################


$ProgressBar    = New-Object System.Windows.Forms.ProgressBar
#$ProgressBar.Size         = New-Object System.Drawing.Size(460,40)
$ProgressBar.Width = $PowerShellForms.Width-7
$ProgressBar.Height = 20
$ProgressBar.Location = New-Object System.Drawing.Point(10,35)
$ProgressBar.Left = 0
$ProgressBar.Top = $PowerShellForms.Height-50
$ProgressBar.Value=0
$ProgressBar.Style="Continuous"
#$ProgressBar.Style = "Marquee"
#$ProgressBar.MarqueeAnimationSpeed = 20
#$ProgressBar.Hide()
$PowerShellForms.Controls.Add($ProgressBar)


$tabPageGetUserList=New-Object System.Windows.Forms.TabPage
$tabPageGetUserList.text = "User"
$tabPageGetUserList.TabIndex=0

$TabControl1.Controls.Add($tabADObject)
$TabControl1.Location = New-Object System.Drawing.Point(0,0)
#$TabControl1.Size=$PowerShellForms.Size
$TabControl1.Left = 5
$TabControl1.Width=$PowerShellForms.Width-25
$TabControl1.Height=470

$PowerShellForms.Controls.Add($TabControl1)
$PowerShellForms.ShowDialog()



