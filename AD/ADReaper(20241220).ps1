<#
Author: Kin.Yan@wsp.com
Description: This is the most powerful tool for you to export object list from AD, such as user/computer/group, if you can find a better one, please delete me :)

Latest update:
Ver 2.0 (2024/12/20)
1. Rewrite most of the logic which enable you to control predefine OU easily
2. You are able to add multiple OUs for one Object type, just split OUs with "!!"
3. Re-design UI

Ver 1.6 (2024/11/05)
1. DL exists in multiple OU, add one more OU for DL ("OU=Messaging,OU=Users,OU=XX,OU=WSPObjects,DC=corp,DC=pbwan,DC=net")
2. Fix the incorrect progress bar percentage
3. To avoid confusion, rename "Distribution List" to "DL&SharedMail", since the report contains DL & Shared Mailbox

Ver 1.56789 (2024/5/17)
1. There is an option(enabled by default) to convert timestamp value to date value(only support AccountExpires,LastLogonTimeStamp,pwdlastSet)

Ver 1.45678 (2024/3/22)
1. Add 2 more object types: Security group & Distribution list. Amazing!

Ver 1.34567 (2024/2/27)
1. You are able to generate report for a specific list of PCs or users listed in a CSV file, amazing!

Ver 1.23456789 (2023/5/26)
1. Initial release
#>

#Region definition
$ShellVersion= "2.0"
$FormTitle = "AD Reaper - V" + $ShellVersion + " - Kin.Yan@wsp.com"
$WorkstationDefaultProperty ="extensionAttribute9,Name,OperatingSystemVersion,OperatingSystem,lastlogondate,whencreated,enabled,distinguishedName"
$UserDefaultProperty ="extensionAttribute7,extensionAttribute9,SamAccountName,displayName,mail,title,department,extensionAttribute12,Enabled,employeeType,whenCreated,lastLogonDate,scriptPath,distinguishedName"
$SecurityGroupDefaultProperty = "sAMAccountName,mail,cn,managedBy,description,displayName,groupType,whenCreated"
$MailboxDefaultProperty = "extensionAttribute9,sAMAccountName,mail,cn,managedBy,description,displayName,groupType,whenCreated,distinguishedName"

$PredefineOUs=@(
    [pscustomobject]@{
        Name="CN";
        Workstation="OU=Clients,OU=CN,OU=WSPObjects,DC=corp,DC=pbwan,DC=net";
        User="OU=Users,OU=CN,OU=WSPObjects,DC=corp,DC=pbwan,DC=net"
        Security="OU=Security,OU=Groups,OU=CN,OU=WSPObjects,DC=corp,DC=pbwan,DC=net";
        Mailbox="OU=Messaging,OU=Groups,OU=CN,OU=WSPObjects,DC=corp,DC=pbwan,DC=net!!OU=Messaging,OU=Users,OU=CN,OU=WSPObjects,DC=corp,DC=pbwan,DC=net"
    }
    [pscustomobject]@{
        Name="HK";
        Workstation="OU=Clients,OU=HK,OU=WSPObjects,DC=corp,DC=pbwan,DC=net";
        User="OU=Users,OU=HK,OU=WSPObjects,DC=corp,DC=pbwan,DC=net"
        Security="OU=Security,OU=Groups,OU=HK,OU=WSPObjects,DC=corp,DC=pbwan,DC=net";
        Mailbox="OU=Messaging,OU=Groups,OU=HK,OU=WSPObjects,DC=corp,DC=pbwan,DC=net!!OU=Messaging,OU=Users,OU=HK,OU=WSPObjects,DC=corp,DC=pbwan,DC=net"
    }
    [pscustomobject]@{
        Name="SG";
        Workstation="OU=Clients,OU=SG,OU=WSPObjects,DC=corp,DC=pbwan,DC=net";
        User="OU=Users,OU=SG,OU=WSPObjects,DC=corp,DC=pbwan,DC=net"
        Security="OU=Security,OU=Groups,OU=SG,OU=WSPObjects,DC=corp,DC=pbwan,DC=net";
        Mailbox="OU=Messaging,OU=Groups,OU=SG,OU=WSPObjects,DC=corp,DC=pbwan,DC=net!!OU=Messaging,OU=Users,OU=SG,OU=WSPObjects,DC=corp,DC=pbwan,DC=net"
    }
    [pscustomobject]@{
        Name="TW";
        Workstation="OU=Clients,OU=TW,OU=WSPObjects,DC=corp,DC=pbwan,DC=net";
        User="OU=Users,OU=TW,OU=WSPObjects,DC=corp,DC=pbwan,DC=net"
        Security="OU=Security,OU=Groups,OU=TW,OU=WSPObjects,DC=corp,DC=pbwan,DC=net";
        Mailbox="OU=Messaging,OU=Groups,OU=TW,OU=WSPObjects,DC=corp,DC=pbwan,DC=net!!OU=Messaging,OU=Users,OU=TW,OU=WSPObjects,DC=corp,DC=pbwan,DC=net"
    }
    [pscustomobject]@{
        Name="TH";
        Workstation="OU=Clients,OU=TH,OU=WSPObjects,DC=corp,DC=pbwan,DC=net";
        User="OU=Users,OU=TH,OU=WSPObjects,DC=corp,DC=pbwan,DC=net"
        Security="OU=Security,OU=Groups,OU=TH,OU=WSPObjects,DC=corp,DC=pbwan,DC=net";
        Mailbox="OU=Messaging,OU=Groups,OU=TH,OU=WSPObjects,DC=corp,DC=pbwan,DC=net!!OU=Messaging,OU=Users,OU=TH,OU=WSPObjects,DC=corp,DC=pbwan,DC=net"
    }
    [pscustomobject]@{
        Name="KR";
        Workstation="OU=Clients,OU=KR,OU=WSPObjects,DC=corp,DC=pbwan,DC=net";
        User="OU=Users,OU=KR,OU=WSPObjects,DC=corp,DC=pbwan,DC=net"
        Security="OU=Security,OU=Groups,OU=KR,OU=WSPObjects,DC=corp,DC=pbwan,DC=net";
        Mailbox="OU=Messaging,OU=Groups,OU=KR,OU=WSPObjects,DC=corp,DC=pbwan,DC=net!!OU=Messaging,OU=Users,OU=KR,OU=WSPObjects,DC=corp,DC=pbwan,DC=net"
    }
    [pscustomobject]@{
        Name="MY";
        Workstation="OU=Clients,OU=MY,OU=WSPObjects,DC=corp,DC=pbwan,DC=net";
        User="OU=Users,OU=MY,OU=WSPObjects,DC=corp,DC=pbwan,DC=net"
        Security="OU=Security,OU=Groups,OU=MY,OU=WSPObjects,DC=corp,DC=pbwan,DC=net";
        Mailbox="OU=Messaging,OU=Groups,OU=MY,OU=WSPObjects,DC=corp,DC=pbwan,DC=net!!OU=Messaging,OU=Users,OU=MY,OU=WSPObjects,DC=corp,DC=pbwan,DC=net"
    }
    [pscustomobject]@{
        Name="PH";
        Workstation="OU=Clients,OU=PH,OU=WSPObjects,DC=corp,DC=pbwan,DC=net";
        User="OU=Users,OU=PH,OU=WSPObjects,DC=corp,DC=pbwan,DC=net"
        Security="OU=Security,OU=Groups,OU=PH,OU=WSPObjects,DC=corp,DC=pbwan,DC=net";
        Mailbox="OU=Messaging,OU=Groups,OU=PH,OU=WSPObjects,DC=corp,DC=pbwan,DC=net!!OU=Messaging,OU=Users,OU=PH,OU=WSPObjects,DC=corp,DC=pbwan,DC=net"
    }
    [pscustomobject]@{
        Name="AU";
        Workstation="OU=Clients,OU=AU,OU=WSPObjects,DC=corp,DC=pbwan,DC=net";
        User="OU=Users,OU=AU,OU=WSPObjects,DC=corp,DC=pbwan,DC=net"
        Security="OU=Security,OU=Groups,OU=AU,OU=WSPObjects,DC=corp,DC=pbwan,DC=net";
        Mailbox="OU=Messaging,OU=Groups,OU=AU,OU=WSPObjects,DC=corp,DC=pbwan,DC=net!!OU=Messaging,OU=Users,OU=AU,OU=WSPObjects,DC=corp,DC=pbwan,DC=net"
    }
    [pscustomobject]@{
        Name="NZ";
        Workstation="OU=Clients,OU=NZ,OU=WSPObjects,DC=corp,DC=pbwan,DC=net";
        User="OU=Users,OU=NZ,OU=WSPObjects,DC=corp,DC=pbwan,DC=net"
        Security="OU=Security,OU=Groups,OU=NZ,OU=WSPObjects,DC=corp,DC=pbwan,DC=net";
        Mailbox="OU=Messaging,OU=Groups,OU=NZ,OU=WSPObjects,DC=corp,DC=pbwan,DC=net!!OU=Messaging,OU=Users,OU=NZ,OU=WSPObjects,DC=corp,DC=pbwan,DC=net"
    }
)

#Endregion Definition

#Region Function
function CreatePredefineOUCheckbox
{
    Param(  [string] $CText,
            [Int16] $PosLeft,
            [Int16] $PosTop
)

$PredefineOU_Checkbox	= New-Object System.Windows.Forms.Checkbox
$PredefineOU_Checkbox.Text = $CText
$PredefineOU_Checkbox.Tag = "PredefineOUCB_" + $CText
$PredefineOU_Checkbox.AutoSize = $True
$PredefineOU_Checkbox.BackColor = "Transparent"
$PredefineOU_Checkbox.Font = New-Object System.Drawing.Font("Tahoma",10,[System.Drawing.FontStyle]::Regular)
$PredefineOU_Checkbox.ForeColor = "Black"
$PredefineOU_Checkbox.Location = New-Object System.Drawing.Point($PosLeft,$PosTop)
#$Page1_Checkbox.TabIndex = 2
$GroupBox_Source.Controls.AddRange($PredefineOU_Checkbox)
}

function GenerateList
{
	Import-Module ActiveDirectory
	$today=Get-Date -Format "yyyyMMddHHmm"
    $SourceArrayList = [System.Collections.ArrayList]::new()

    #add source into SourceArrayList
    if($Source_CbPredefineOU.Checked -eq $True)
    {
        foreach ($FormObject in $PowerShellForms.Controls)
        {
            if($FormObject -match "GroupBox")
            {
                foreach($GBO in $FormObject.Controls)
                {
                    if(($GBO.Tag -match "PredefineOUCB_") -and ($GBO.Checked -eq $true))
                    {
                        if($ObjectType_User.Checked -eq $true)
                        {
                            $users=($PredefineOUs | Where-Object {$_.Name -eq $GBO.Text}).User -split "!!"

                            foreach($user in $users)
                            {
                            $SourceArrayList.Add($user)
                            }
                        }
                        elseif($ObjectType_PC.Checked -eq $true)
                        {
                            $workstations=($PredefineOUs | Where-Object {$_.Name -eq $GBO.Text}).Workstation -split "!!"
                            foreach($workstation in $workstations)
                            {
                            $SourceArrayList.Add($workstation)
                            }
                        }
                        elseif($ObjectType_Mailbox.Checked -eq $true)
                        {
                            $mailboxs=($PredefineOUs | Where-Object {$_.Name -eq $GBO.Text}).Mailbox -split "!!"
                            foreach($mailbox in $mailboxs)
                            {
                            $SourceArrayList.Add($mailbox)
                            }
                        }
                        elseif($ObjectType_SecurityGroup.Checked -eq $true)
                        {
                            $securities=($PredefineOUs | Where-Object {$_.Name -eq $GBO.Text}).Security -split "!!"
                            foreach($security in $securities)
                            {
                            $SourceArrayList.Add($security)
                            }
                        }
                    }
                }
            }
        }
    }    
    elseif($Source_CbSpecificOU.Checked -eq $true)
    {
        $SourceArrayList.Add($Source_tbSpecificOUPath.Text)
    }
    elseif($Source_CbImportFromCSV.checked -eq $true)
    {
        $CSVObjects=import-csv -path $Source_tbImportFromCSVPath.Text -header "Name"
        foreach($CSVObject in $CSVObjects)
        {
            $SourceArrayList.Add($CSVObject.Name)
        }
    }

    $ReportColumn=$tbADProperty.Text
    if($CbConvertTimeStampToDate.Checked -eq $true) 
    {
        #https://learn.microsoft.com/en-us/archive/technet-wiki/22461.understanding-the-ad-account-attributes-lastlogon-lastlogontimestamp-and-lastlogondate
        $ReportColumn=$ReportColumn -Replace "accountexpires","AccountExpirationDate"
        $ReportColumn=$ReportColumn -Replace "lastLogonTimestamp","LastLogonDate"
        $ReportColumn=$ReportColumn -Replace "pwdLastSet","PasswordLastSet"
    }
    $RColumns=$ReportColumn -Split ","
    $ReportColumnList=[System.Collections.ArrayList]@()
    Foreach($rcolumn in $RColumns)
    {
        $ReportColumnList.Add($rcolumn)
    }

    $ProgressBar.Value=0
    $progresscount = $SourceArrayList.Count
    $i=0

    if($ObjectType_PC.Checked -eq $true) #PC
    {
        $filename=$tbReportPath.Text + "\PCList_" + $today + ".csv"
    }
    elseif ($ObjectType_User.Checked -eq $true) #User
    {
        $filename=$tbReportPath.Text + "\UserList_" + $today + ".csv"
    }
    elseif($ObjectType_SecurityGroup.Checked -eq $true) #Security Group
    {
        $filename=$tbReportPath.Text + "\SecurityGroupList_" + $today + ".csv"
    }
    elseif($ObjectType_Mailbox.Checked -eq $true) #Mailbox
    {
        $filename=$tbReportPath.Text + "\MailboxList_" + $today + ".csv"
    }

    foreach($Source in $SourceArrayList)
    {
        $i++
        [int]$percentage=($i/$progresscount)*100
        $ProgressBar.Value=$percentage
        if($Source_CbImportFromCSV.Checked -eq $true)
        {
            if($ObjectType_PC.Checked -eq $true) #PC
            {
                Get-ADComputer -filter {Name -eq $Source} -Properties $ReportColumnList | Select-Object $ReportColumnList | Export-Csv -Encoding UTF8 -NoTypeInformation -append -Force -path $filename
            }
            elseif ($ObjectType_User.Checked -eq $true) #User
            {
                Get-ADUser -filter {(EmailAddress -eq $Source) -or (SamAccountName -eq $Source)} -Properties $ReportColumnList | Select-Object $SelectColumnList | export-csv -encoding UTF8 -NoTypeInformation -Force -Append -Path $filename
            }
            elseif($ObjectType_SecurityGroup.Checked -eq $true) #Security Group
            {
                Get-ADObject -filter {(objectClass -eq "group") -and (samAccountName -eq $Source)} -Properties $ReportColumnList | select-object $ReportColumnList | export-csv -Encoding UTF8 -NoTypeInformation -Force -Append -Path $filename
            }
            elseif($ObjectType_Mailbox.Checked -eq $true) #Mailbox
            {
                Get-ADObject -filter {(samAccountName -eq $Source) -or (mail -eq $Source)} -Properties $ReportColumnList | select-object $ReportColumnList | export-csv -Encoding UTF8 -NoTypeInformation -Force -Append -Path $filename
            }
        }
        else {
            if($ObjectType_PC.Checked -eq $true) #PC
            {
                Get-ADComputer -SearchBase $Source -Filter * -Properties $ReportColumnList | Select-Object $ReportColumnList | Export-Csv -Encoding UTF8 -NoTypeInformation -append -Force -path $filename
                #Get-ADObject -SearchBase $Source -Filter {(objectClass -eq "computer")} -Properties $ReportColumnList | Select-Object $ReportColumnList | Export-Csv -Encoding UTF8 -NoTypeInformation -append -Force -path $filename
            }
            elseif ($ObjectType_User.Checked -eq $true) #User
            {
                Get-ADUser -SearchBase $Source -Filter * -Properties $ReportColumnList | Select-Object $ReportColumnList | export-csv -encoding UTF8 -NoTypeInformation -Force -Append -Path $filename
            }
            elseif($ObjectType_SecurityGroup.Checked -eq $true) #Security Group
            {
                Get-ADObject -SearchBase $Source -filter {(objectClass -eq "group")} -Properties $ReportColumnList | select-object $ReportColumnList | export-csv -Encoding UTF8 -NoTypeInformation -Force -Append -Path $filename
            }
            elseif($ObjectType_Mailbox.Checked -eq $true) #Mailbox
            {
                Get-ADObject -SearchBase $Source -filter {(objectClass -eq "group") -or (objectClass -eq "user")} -Properties $ReportColumnList | select-object $ReportColumnList | export-csv -Encoding UTF8 -NoTypeInformation -Force -Append -Path $filename
            }
        }
    }

    [System.Windows.Forms.Messagebox]::Show("Done, go and get your report, no thanks ^_^")
    
}
#Endregion Function

#Region GUI
############## WinForm ##############
Add-Type -AssemblyName System.Windows.Forms
$PowerShellForms = New-Object system.Windows.Forms.Form
$PowerShellForms.Text= $FormTitle
$PowerShellForms.Size = New-Object System.Drawing.Size(470,620)
$PowerShellForms.MinimizeBox = $False
$PowerShellForms.MaximizeBox = $False
$PowerShellForms.FormBorderStyle = [System.Windows.Forms.FormBorderStyle]::FixedDialog

$PowerShellForms.SizeGripStyle = "Hide"
$PowerShellForms.Icon = $Icons
$PowerShellForms.StartPosition = "CenterScreen"
$PowerShellForms.Font = New-Object System.Drawing.Font("Tahoma",10,[System.Drawing.FontStyle]::Regular)
############## End WinForm ##############

############## Object Type ##############
$GroupBox_ObjectType = New-Object System.Windows.Forms.GroupBox
$GroupBox_ObjectType.Location = New-Object System.Drawing.Point(10,10)
$GroupBox_ObjectType.Width =$PowerShellForms.Width-40
$GroupBox_ObjectType.Height =80
$GroupBox_ObjectType.Text = "Object Type"

$ObjectType_PC = New-Object System.Windows.Forms.RadioButton
$ObjectType_PC.AutoSize = $True
$ObjectType_PC.Location = New-Object System.Drawing.Point(10,20)
$ObjectType_PC.Text = "PC"
$ObjectType_PC.Checked = $True
$GroupBox_ObjectType.Controls.AddRange($ObjectType_PC)
$ObjectType_PC.Add_CheckedChanged(
    {
        if($ObjectType_PC.Checked -eq $true)
        {
            $tbADProperty.Text = $WorkstationDefaultProperty
            $ObjectType_lbCSVDescription.Text = "Format: one PC name per line, no header is required"
        }
    }
)

$ObjectType_User = New-Object System.Windows.Forms.RadioButton
$ObjectType_User.AutoSize = $True
$ObjectType_User.Location = New-Object System.Drawing.Point(145,20)
$ObjectType_User.Text = "User"
$GroupBox_ObjectType.Controls.AddRange($ObjectType_User)
$ObjectType_User.Add_CheckedChanged(
    {
        if($ObjectType_User.Checked -eq $true)
        {
            $tbADProperty.Text = $UserDefaultProperty
            $ObjectType_lbCSVDescription.Text = "Format: one sAMAccountName or email address per line, no header is required"
        }
    }
)

$ObjectType_SecurityGroup = New-Object System.Windows.Forms.RadioButton
$ObjectType_SecurityGroup.AutoSize = $True
$ObjectType_SecurityGroup.Location = New-Object System.Drawing.Point(10,50)
$ObjectType_SecurityGroup.Text = "Security Group"
$GroupBox_ObjectType.Controls.AddRange($ObjectType_SecurityGroup)
$ObjectType_SecurityGroup.Add_CheckedChanged(
    {
        if($ObjectType_SecurityGroup.Checked -eq $true)
        {
            $tbADProperty.Text = $SecurityGroupDefaultProperty
            $ObjectType_lbCSVDescription.Text = "Format: one sAMAccountName per line, no header is required"
        }
    }
)

$ObjectType_Mailbox = New-Object System.Windows.Forms.RadioButton
$ObjectType_Mailbox.AutoSize = $True
$ObjectType_Mailbox.Location = New-Object System.Drawing.Point(145,50)
$ObjectType_Mailbox.Text = "Mailbox"
$GroupBox_ObjectType.Controls.AddRange($ObjectType_Mailbox)
$ObjectType_Mailbox.Add_CheckedChanged(
    {
        if($ObjectType_Mailbox.Checked -eq $true)
        {
            $tbADProperty.Text = $MailboxDefaultProperty
            $ObjectType_lbCSVDescription.Text = "Format: one sAMAccountName or email address per line, no header is required"
        }
    }
)
$PowerShellForms.Controls.Add($GroupBox_ObjectType)
############## End Object Type ##############

############## Source ##############
$GroupBox_Source = New-Object System.Windows.Forms.GroupBox
#$GroupBox_Source.Location = New-Object System.Drawing.Point(10,115)
#$GroupBox_Source.Size = New-Object System.Drawing.Size(320,235)
$GroupBox_Source.Width =$PowerShellForms.Width-40
$GroupBox_Source.Top = $GroupBox_ObjectType.Height+$GroupBox_ObjectType.Top+10
$GroupBox_Source.Left = 10
$GroupBox_Source.Height = 245
$GroupBox_Source.Text = "Source"

$Source_CbPredefineOU = New-Object System.Windows.Forms.Checkbox
$Source_CbPredefineOU.Text = "From Predefine AD OU"
$Source_CbPredefineOU.AutoSize = $True
$Source_CbPredefineOU.BackColor = "Transparent"
$Source_CbPredefineOU.Font = New-Object System.Drawing.Font("Tahoma",10,[System.Drawing.FontStyle]::Regular)
$Source_CbPredefineOU.ForeColor = "Black"
$Source_CbPredefineOU.Location = New-Object System.Drawing.Point(10,15)
$Source_CbPredefineOU.TabIndex = 7
$Source_CbPredefineOU.Checked = $true
$GroupBox_Source.Controls.AddRange($Source_CbPredefineOU)
$Source_CbPredefineOU.Add_CheckedChanged(
    {
        foreach ($FormObject in $PowerShellForms.Controls)
        {
            if($FormObject -match "GroupBox")
            {
                foreach($GBO in $FormObject.Controls)
                {
                    if($GBO.Tag -match "PredefineOUCB_")
                    {
                        $GBO.enabled = $Source_CbPredefineOU.Checked
                    }
                }
            }
        }
    }
)
$Source_CbPredefineOU.Add_Click(
    {
        if($Source_CbPredefineOU.Checked -eq $true)
        {
            $Source_CbSpecificOU.Checked = -not $Source_CbPredefineOU.Checked
            $Source_CbImportFromCSV.Checked = -not $Source_CbPredefineOU.Checked
        }
    }
)
############## Source -> PredefineOU ##############
$CheckboxEachRow=5
for($i=0;$i -lt $PredefineOUs.Count;$i++)
{
    $pl=30+70*($i % $CheckboxEachRow)
    $pt=15+20*([Math]::Floor($i/$CheckboxEachRow)+1)
    CreatePredefineOUCheckbox -CText $PredefineOUs[$i].Name -PosLeft $pl -PosTop $pt
}

############## End Source -> PredefineOU ##############

$Source_CbSpecificOU = New-Object System.Windows.Forms.Checkbox
$Source_CbSpecificOU.Text = "From Specific AD OU (distinguishedName)"
$Source_CbSpecificOU.AutoSize = $True
$Source_CbSpecificOU.BackColor = "Transparent"
$Source_CbSpecificOU.Font = New-Object System.Drawing.Font("Tahoma",10,[System.Drawing.FontStyle]::Regular)
$Source_CbSpecificOU.ForeColor = "Black"
$Source_CbSpecificOU.Location = New-Object System.Drawing.Point(10,100)
$Source_CbSpecificOU.TabIndex = 7
$GroupBox_Source.Controls.AddRange($Source_CbSpecificOU)
$Source_CbSpecificOU.Add_CheckedChanged(
    {
        $Source_tbSpecificOUPath.enabled = $Source_CbSpecificOU.Checked
    }
)
$Source_CbSpecificOU.Add_Click(
    {
        if($Source_CbSpecificOU.Checked -eq $true)
        {
            $Source_CbPredefineOU.Checked = -not $Source_CbSpecificOU.Checked
            $Source_CbImportFromCSV.Checked = -not $Source_CbSpecificOU.Checked
        }
    }
)
$Source_tbSpecificOUPath = New-Object System.Windows.Forms.Textbox
$Source_tbSpecificOUPath.Size = New-Object System.Drawing.Size(390,40)
#$Source_tbSpecificOUPath.Width = 400
#$Source_tbSpecificOUPath.Height = 40
$Source_tbSpecificOUPath.AutoSize = $True
$Source_tbSpecificOUPath.Font = New-Object System.Drawing.Font("Tahoma",10,[System.Drawing.FontStyle]::Regular)
$Source_tbSpecificOUPath.ForeColor = "Black"
$Source_tbSpecificOUPath.Location = New-Object System.Drawing.Point(30,125)
$Source_tbSpecificOUPath.TabIndex = 10
$Source_tbSpecificOUPath.Enabled=$False
$GroupBox_Source.Controls.AddRange($Source_tbSpecificOUPath)

$Source_CbImportFromCSV	= New-Object System.Windows.Forms.Checkbox
$Source_CbImportFromCSV.Text = "From CSV List"
$Source_CbImportFromCSV.AutoSize = $True
$Source_CbImportFromCSV.BackColor = "Transparent"
$Source_CbImportFromCSV.Font = New-Object System.Drawing.Font("Tahoma",10,[System.Drawing.FontStyle]::Regular)
$Source_CbImportFromCSV.ForeColor = "Black"
$Source_CbImportFromCSV.Location = New-Object System.Drawing.Point(10,155)
$Source_CbImportFromCSV.TabIndex = 9
$GroupBox_Source.Controls.AddRange($Source_CbImportFromCSV)
$Source_CbImportFromCSV.Add_CheckedChanged(
    {
        $Source_tbImportFromCSVPath.enabled = $Source_CbImportFromCSV.Checked
        $Source_BtnSelectCSVFile.enabled=$Source_CbImportFromCSV.Checked
    }
)
$Source_CbImportFromCSV.Add_Click(
    {
        if($Source_CbImportFromCSV.Checked -eq $true)
        {
            $Source_CbPredefineOU.Checked = -not $Source_CbImportFromCSV.Checked
            $Source_CbSpecificOU.Checked = -not $Source_CbImportFromCSV.Checked
        }
    }
)
$Source_tbImportFromCSVPath	= New-Object System.Windows.Forms.Textbox
$Source_tbImportFromCSVPath.Size = New-Object System.Drawing.Size(390,40)
$Source_tbImportFromCSVPath.AutoSize = $True
$Source_tbImportFromCSVPath.Font = New-Object System.Drawing.Font("Tahoma",10,[System.Drawing.FontStyle]::Regular)
$Source_tbImportFromCSVPath.ForeColor = "Black"
$Source_tbImportFromCSVPath.Location = New-Object System.Drawing.Point(30,180)
$Source_tbImportFromCSVPath.TabIndex = 10
$Source_tbImportFromCSVPath.Enabled=$False
#$Source_tbImportFromCSVPath.Text="C:\Support\ADReaper\NAUsers.csv"
$GroupBox_Source.Controls.AddRange($Source_tbImportFromCSVPath)

$ObjectType_lbCSVDescription	= New-Object System.Windows.Forms.Label
$ObjectType_lbCSVDescription.Text = "Format: one PC name per line, no header is required"
#$lbADProperty.AutoSize = $True
$ObjectType_lbCSVDescription.Size = New-Object System.Drawing.Size(280,30)
$ObjectType_lbCSVDescription.BackColor = "Transparent"
$ObjectType_lbCSVDescription.Font = New-Object System.Drawing.Font("Tahoma",8,[System.Drawing.FontStyle]::Italic)
$ObjectType_lbCSVDescription.ForeColor = "Black"
$ObjectType_lbCSVDescription.Location = New-Object System.Drawing.Point(30,210)
$ObjectType_lbCSVDescription.TabIndex = 10
$GroupBox_Source.Controls.AddRange($ObjectType_lbCSVDescription)

$Source_BtnSelectCSVFile = New-Object System.Windows.Forms.Button
$Source_BtnSelectCSVFile.Size = New-Object System.Drawing.Size(70,25)
$Source_BtnSelectCSVFile.Location = New-Object System.Drawing.Point(350,210)
$Source_BtnSelectCSVFile.Text = "CSV File"
$Source_BtnSelectCSVFile.Font = New-Object System.Drawing.Font("Tahoma",10,[System.Drawing.FontStyle]::Regular)
$Source_BtnSelectCSVFile.Add_Click(
    {
        $CurrentLocation = Get-Location
        $CSVBrowser = New-Object System.Windows.Forms.OpenFileDialog
        $CSVBrowser.InitialDirectory = $CurrentLocation
        $CSVBrowser.Filter = "CSV (*.csv)|*.csv"
        $CSVBrowser.ShowDialog() | Out-Null
        if($CSVBrowser.FileName -ne $null)
        {
            $Source_tbImportFromCSVPath.text = $CSVBrowser.FileName
        }
    }
)
$GroupBox_Source.Controls.AddRange($Source_BtnSelectCSVFile)
$PowerShellForms.Controls.Add($GroupBox_Source)
############## End Source ##############

$lbADProperty	= New-Object System.Windows.Forms.Label
$lbADProperty.Text = "Report column(AD property)"
#$lbADProperty.AutoSize = $True
$lbADProperty.Size = New-Object System.Drawing.Size(300,20)
$lbADProperty.BackColor = "Transparent"
$lbADProperty.Font = New-Object System.Drawing.Font("Tahoma",10,[System.Drawing.FontStyle]::Regular)
$lbADProperty.ForeColor = "Black"
$lbADProperty.Location = New-Object System.Drawing.Point(10,355)
$lbADProperty.TabIndex = 10
$PowerShellForms.Controls.Add($lbADProperty)

$tbADProperty	= New-Object System.Windows.Forms.Textbox
$tbADProperty.Text = $WorkstationDefaultProperty
#$tbADProperty.Size = New-Object System.Drawing.Size(320,80)
$tbADProperty.Width = $PowerShellForms.Width - 40
$tbADProperty.Height = 60
$tbADProperty.AutoSize = $True
$tbADProperty.Multiline=$True
$tbADProperty.Font = New-Object System.Drawing.Font("Tahoma",10,[System.Drawing.FontStyle]::Regular)
$tbADProperty.ForeColor = "Black"
$tbADProperty.Location = New-Object System.Drawing.Point(10,375)
$tbADProperty.TabIndex = 10
$PowerShellForms.Controls.Add($tbADProperty)

$lbReportPath	= New-Object System.Windows.Forms.Label
$lbReportPath.Text = "Report Path"
#$lbReportPath.Size = New-Object System.Drawing.Size(300,50)
$lbReportPath.AutoSize = $True
$lbReportPath.BackColor = "Transparent"
$lbReportPath.Font = New-Object System.Drawing.Font("Tahoma",10,[System.Drawing.FontStyle]::Regular)
$lbReportPath.ForeColor = "Black"
$lbReportPath.Location = New-Object System.Drawing.Point(10,465)
$lbReportPath.TabIndex = 10
$PowerShellForms.Controls.Add($lbReportPath)

$tbReportPath	= New-Object System.Windows.Forms.Textbox
$tbReportPath.Text = $PSScriptRoot
$tbReportPath.Size = New-Object System.Drawing.Size(430,50)
$tbReportPath.AutoSize = $True
$tbReportPath.Font = New-Object System.Drawing.Font("Tahoma",10,[System.Drawing.FontStyle]::Regular)
$tbReportPath.ForeColor = "Black"
$tbReportPath.Location = New-Object System.Drawing.Point(10,490)
$tbReportPath.TabIndex = 10
$PowerShellForms.Controls.Add($tbReportPath)

$BtnGenerateReport = New-Object System.Windows.Forms.Button
$BtnGenerateReport.Size = New-Object System.Drawing.Size(140,40)
$BtnGenerateReport.Location = New-Object System.Drawing.Point(300,520)
$BtnGenerateReport.Text = "Show me the report"
$BtnGenerateReport.Font = New-Object System.Drawing.Font("Tahoma",10,[System.Drawing.FontStyle]::Regular)
$BtnGenerateReport.Add_Click(
    {
        $selectedoffice=$false

        foreach ($FormObject in $PowerShellForms.Controls)
        {
            if($FormObject -match "GroupBox")
            {
                foreach($GBO in $FormObject.Controls)
                {
                    if(($GBO.Tag -match "PredefineOUCB_") -and ($GBO.Checked -eq $true))
                    {
                        $selectedoffice=$true
                    }
                }
            }
        }

        if((($selectedoffice -eq $false) -and (($Source_CbSpecificOU.Checked -eq $false) -or ($Source_tbSpecificOUPath.Text.Length -lt 10))) -and (($Source_CbImportFromCSV.Checked -eq $false) -or ($Source_tbImportFromCSVPath.Text.Length -lt 8)))
        {
            [System.Windows.Forms.Messagebox]::Show("Please choose a OU or specify a CSV file!")
        }
        else
        {
            GenerateList
        }
    }
)
$PowerShellForms.Controls.Add($BtnGenerateReport)

$CbConvertTimeStampToDate = New-Object System.Windows.Forms.Checkbox
$CbConvertTimeStampToDate.Text = "Use PowerShell converted attribute"
$CbConvertTimeStampToDate.AutoSize = $True
$CbConvertTimeStampToDate.BackColor = "Transparent"
$CbConvertTimeStampToDate.Font = New-Object System.Drawing.Font("Tahoma",10,[System.Drawing.FontStyle]::Regular)
$CbConvertTimeStampToDate.ForeColor = "Black"
$CbConvertTimeStampToDate.Location = New-Object System.Drawing.Point(10,435)
$CbConvertTimeStampToDate.TabIndex = 7
$CbConvertTimeStampToDate.Checked=$true
$PowerShellForms.Controls.Add($CbConvertTimeStampToDate)

$ProgressBar    = New-Object System.Windows.Forms.ProgressBar
#$ProgressBar.Size         = New-Object System.Drawing.Size(460,40)
$ProgressBar.Width = $PowerShellForms.Width-7
$ProgressBar.Height = 20
$ProgressBar.Location = New-Object System.Drawing.Point(10,35)
$ProgressBar.Left = 0
$ProgressBar.Top = $PowerShellForms.Height-50
$ProgressBar.Value=0
$ProgressBar.Style="Continuous"
$PowerShellForms.Controls.Add($ProgressBar)

#Endregion GUI

$PowerShellForms.ShowDialog()
