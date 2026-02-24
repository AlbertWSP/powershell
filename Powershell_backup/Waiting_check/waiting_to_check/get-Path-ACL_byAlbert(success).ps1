Add-Type -AssemblyName System.Windows.Forms

$PowershellPath = "C:\temp\Powershell"
$ReportPath = Join-Path $PowershellPath "Report"
if((Test-Path $ReportPath) -eq $False)
{
	New-Item -Path $ReportPath -ItemType Directory
}

function Get-ACLReport ($folder) {
    $report = @()
    try
    {
    $accessresult=(get-acl -LiteralPath $folder).access
    $folderowner=(get-acl -LiteralPath $folder).Owner

    $Identitys=$accessresult.IdentityReference
    $FileSystemRights=$accessresult.FileSystemRights
    #$accesscontroltypes=$accessresult.AccessControlType

    $Identity=""
    #$accesscontroltype=""
    for($i=0;$i -lt $Identitys.count;$i++)
        {
            $Identity=$Identitys[$i]

            $FileSystemRight=""
            for($j=0;$j -lt $FileSystemRights[$i].count;$j++)
            {
                $FileSystemRight=$FileSystemRight+";"+$FileSystemRights[$i]
            }
            #$accesscontroltype=$accesscontroltypes[$i]
            $FileSystemRight=$FileSystemRight.Substring(1,$FileSystemRight.Length-1)

            #New-Object psobject -Property @{'Folder path'=$folder;'Identity'=$Identity;'AccessRight'=$FileSystemRight;'Control Type'=$accesscontroltype}

            $isADAccount=$false
            if($Identity.Value.Substring(0,5) -eq "CORP\")
            {
                $IdentityName=$Identity.Value -replace "CORP\\"
                $IdentityObject=Get-ADObject -Filter "SamAccountName -eq '$IdentityName'"

                if(($IdentityObject.objectClass -ne "user") -and ($IdentityName -ne "Domain Users")) #security group, ignore "Domain Users"
                {
                    $GroupMembers=Get-ADGroupMember $IdentityName -recursive | select-object SamAccountName
                }
                else #AD user account or "domain users" group
                {
                    $GroupMembers=$IdentityName
                }
                $isADAccount=$true
            }
            else #local account
            {
                $GroupMembers = $Identity.Value
            }

            foreach($GroupMember in $GroupMembers)
            {
                $obj = New-Object -TypeName psobject
                #$DisplayName=""
                if($isADAccount -eq $true)
                {
                    #$ADO=Get-ADObject -Filter "SamAccountName -eq '$IdentityName'"
                    if($IdentityObject.objectClass -ne "group") #user account
                    {
                        $GM="-"
                        $DisplayName=(get-aduser $GroupMember -Properties *).Displayname
                    }
                    else #AD group(include domain users)
                    {
                        if($GroupMember -eq "Domain Users")
                        {
                            $GM=$GroupMember
                            $DisplayName=$GroupMember
                        }
                        else
                        {
                            $GM=$GroupMember.SamAccountName
                            $DisplayName=(get-aduser $GroupMember.SamAccountName -Properties *).Displayname
                        }
                    }
                    $ID=$IdentityName
                }
                else #local id
                {
                    $ID=$GroupMember
                    $GM="-"
                    $DisplayName = $GroupMember
                }                  
                $obj | Add-Member -MemberType NoteProperty -Name FolderPath -Value $folder
                $obj | Add-Member -MemberType NoteProperty -Name Identity -Value $ID
                $obj | Add-Member -MemberType NoteProperty -Name Member -Value $GM
                $obj | Add-Member -MemberType NoteProperty -Name DisplayName -Value $DisplayName
                $obj | Add-Member -MemberType NoteProperty -Name AccessRight -Value $FileSystemRight
                #$obj | Add-Member -MemberType NoteProperty -Name AccessType -Value $accesscontroltype
                $obj | Add-Member -MemberType NoteProperty -Name Owner -Value $folderowner

                $report +=$obj
            }
        }
    }
    catch
    {
        $obj = New-Object -TypeName psobject

        $obj | Add-Member -MemberType NoteProperty -Name FolderPath -Value $folder
        $obj | Add-Member -MemberType NoteProperty -Name Identity -Value "error"
        $obj | Add-Member -MemberType NoteProperty -Name Member -Value "error"
        $obj | Add-Member -MemberType NoteProperty -Name AccessRight -Value "error"
        #$obj | Add-Member -MemberType NoteProperty -Name AccessType -Value "error"
        $obj | Add-Member -MemberType NoteProperty -Name Owner -Value "error"

        $report +=$obj
    }

    return $report
    }

function Show-InputForm {
    $form = New-Object System.Windows.Forms.Form
    $form.Text = "ACL Report Input"
    $form.Size = New-Object System.Drawing.Size(300, 150)
    $form.StartPosition = "CenterScreen"

    $label = New-Object System.Windows.Forms.Label
    $label.Text = "Enter Folder Path:"
    $label.AutoSize = $true
    $label.Location = New-Object System.Drawing.Point(10, 20)
    $form.Controls.Add($label)

    $textbox = New-Object System.Windows.Forms.TextBox
    $textbox.Size = New-Object System.Drawing.Size(260, 20)
    $textbox.Location = New-Object System.Drawing.Point(10, 50)
    $form.Controls.Add($textbox)

    $button = New-Object System.Windows.Forms.Button
    $button.Text = "Generate Report"
    $button.Location = New-Object System.Drawing.Point(10, 80)
    $button.Add_Click({
        $folderPath = $textbox.Text
        $R=Get-ACLReport $folderPath
        $today=Get-Date -Format "yyyyMMddHHmm"
        $reportname = "ACLReport_" + "_" + $today + ".csv"
        $reportname = Join-Path $ReportPath $reportname
        $R | export-csv $reportname -NoTypeInformation -Force -Append -Encoding UTF8
        $form.Close()
    })
    $form.Controls.Add($button)

    $form.ShowDialog()
}

Show-InputForm