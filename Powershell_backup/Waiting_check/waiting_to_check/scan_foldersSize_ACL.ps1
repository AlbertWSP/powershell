# Add required assembly for Windows Forms
Add-Type -AssemblyName System.Windows.Forms

# Set up report paths
$PowershellPath = "C:\temp\Powershell"
$ReportPath = Join-Path $PowershellPath "Report"
if((Test-Path $ReportPath) -eq $False) {
    New-Item -Path $ReportPath -ItemType Directory
}

# Function to convert bytes to MB or GB
function Format-FileSize {
    param (
        [long]$Size
    )
    
    if ($Size -ge 1GB) {
        return "{0:N2} GB" -f ($Size / 1GB)
    }
    else {
        return "{0:N2} MB" -f ($Size / 1MB)
    }
}

# Function to get color based on size
function Get-SizeColor {
    param (
        [long]$Size
    )
    
    if ($Size -ge 1GB) {
        return "Red"
    }
    elseif ($Size -ge 500MB) {
        return "Yellow"
    }
    else {
        return "Green"
    }
}

# Function to get ACL report
function Get-ACLReport ($folder) {
    $report = @()
    try {
        $accessresult = (get-acl -LiteralPath $folder).access
        $folderowner = (get-acl -LiteralPath $folder).Owner

        $Identitys = $accessresult.IdentityReference
        $FileSystemRights = $accessresult.FileSystemRights

        $Identity = ""
        for($i=0; $i -lt $Identitys.count; $i++) {
            $Identity = $Identitys[$i]

            $FileSystemRight = ""
            for($j=0; $j -lt $FileSystemRights[$i].count; $j++) {
                $FileSystemRight = $FileSystemRight + ";" + $FileSystemRights[$i]
            }
            $FileSystemRight = $FileSystemRight.Substring(1, $FileSystemRight.Length-1)

            $isADAccount = $false
            if($Identity.Value.Substring(0,5) -eq "CORP\") {
                $IdentityName = $Identity.Value -replace "CORP\\"
                $IdentityObject = Get-ADObject -Filter "SamAccountName -eq '$IdentityName'"

                if(($IdentityObject.objectClass -ne "user") -and ($IdentityName -ne "Domain Users")) {
                    $GroupMembers = Get-ADGroupMember $IdentityName -recursive | select-object SamAccountName
                }
                else {
                    $GroupMembers = $IdentityName
                }
                $isADAccount = $true
            }
            else {
                $GroupMembers = $Identity.Value
            }

            foreach($GroupMember in $GroupMembers) {
                $obj = New-Object -TypeName psobject
                if($isADAccount -eq $true) {
                    if($IdentityObject.objectClass -ne "group") {
                        $GM = "-"
                        $DisplayName = (get-aduser $GroupMember -Properties *).Displayname
                    }
                    else {
                        if($GroupMember -eq "Domain Users") {
                            $GM = $GroupMember
                            $DisplayName = $GroupMember
                        }
                        else {
                            $GM = $GroupMember.SamAccountName
                            $DisplayName = (get-aduser $GroupMember.SamAccountName -Properties *).Displayname
                        }
                    }
                    $ID = $IdentityName
                }
                else {
                    $ID = $GroupMember
                    $GM = "-"
                    $DisplayName = $GroupMember
                }                  
                $obj | Add-Member -MemberType NoteProperty -Name FolderPath -Value $folder
                $obj | Add-Member -MemberType NoteProperty -Name Identity -Value $ID
                $obj | Add-Member -MemberType NoteProperty -Name Member -Value $GM
                $obj | Add-Member -MemberType NoteProperty -Name DisplayName -Value $DisplayName
                $obj | Add-Member -MemberType NoteProperty -Name AccessRight -Value $FileSystemRight
                $obj | Add-Member -MemberType NoteProperty -Name Owner -Value $folderowner

                $report += $obj
            }
        }
    }
    catch {
        $obj = New-Object -TypeName psobject
        $obj | Add-Member -MemberType NoteProperty -Name FolderPath -Value $folder
        $obj | Add-Member -MemberType NoteProperty -Name Identity -Value "error"
        $obj | Add-Member -MemberType NoteProperty -Name Member -Value "error"
        $obj | Add-Member -MemberType NoteProperty -Name AccessRight -Value "error"
        $obj | Add-Member -MemberType NoteProperty -Name Owner -Value "error"
        $report += $obj
    }
    return $report
}

# Function to show file selection dialog
function Show-FileSelectionDialog {
    $form = New-Object System.Windows.Forms.Form
    $form.Text = "Select Paths File"
    $form.Size = New-Object System.Drawing.Size(500, 150)
    $form.StartPosition = "CenterScreen"

    $label = New-Object System.Windows.Forms.Label
    $label.Text = "Select the text file containing folder paths:"
    $label.AutoSize = $true
    $label.Location = New-Object System.Drawing.Point(10, 20)
    $form.Controls.Add($label)

    $textbox = New-Object System.Windows.Forms.TextBox
    $textbox.Size = New-Object System.Drawing.Size(360, 20)
    $textbox.Location = New-Object System.Drawing.Point(10, 50)
    $form.Controls.Add($textbox)

    $browseButton = New-Object System.Windows.Forms.Button
    $browseButton.Text = "Browse"
    $browseButton.Location = New-Object System.Drawing.Point(380, 48)
    $browseButton.Add_Click({
        $openFileDialog = New-Object System.Windows.Forms.OpenFileDialog
        $openFileDialog.Filter = "Text files (*.txt)|*.txt"
        $openFileDialog.Title = "Select Paths File"
        if ($openFileDialog.ShowDialog() -eq [System.Windows.Forms.DialogResult]::OK) {
            $textbox.Text = $openFileDialog.FileName
        }
    })
    $form.Controls.Add($browseButton)

    $button = New-Object System.Windows.Forms.Button
    $button.Text = "Process Folders"
    $button.Location = New-Object System.Drawing.Point(10, 80)
    $button.Add_Click({
        $pathsFile = $textbox.Text
        if (Test-Path $pathsFile) {
            Process-Folders $pathsFile
            $form.Close()
        }
        else {
            [System.Windows.Forms.MessageBox]::Show("Please select a valid file path.", "Error", [System.Windows.Forms.MessageBoxButtons]::OK, [System.Windows.Forms.MessageBoxIcon]::Error)
        }
    })
    $form.Controls.Add($button)

    $form.ShowDialog()
}

# Function to process folders from file
function Process-Folders ($pathsFile) {
    # Read the list of folders from the text file
    $folders = Get-Content -Path $pathsFile
    
    Write-Host "`nStarting Folder Analysis" -ForegroundColor Cyan
    Write-Host "=====================" -ForegroundColor Cyan
    Write-Host ""
    
    # Initialize arrays to store results
    $sizeResults = @()
    $aclResults = @()
    
    foreach ($folder in $folders) {
        if (Test-Path $folder) {
            Write-Host "Processing folder: $folder" -ForegroundColor Yellow
            
            # Generate size report
            $folderSize = (Get-ChildItem -Path $folder -Recurse | Measure-Object -Property Length -Sum).Sum
            $fileCount = (Get-ChildItem -Path $folder -Recurse | Measure-Object).Count
            
            Write-Host "Size: " -NoNewline
            Write-Host (Format-FileSize -Size $folderSize) -ForegroundColor (Get-SizeColor -Size $folderSize)
            Write-Host "Files: " -NoNewline
            Write-Host $fileCount -ForegroundColor Magenta
            Write-Host "-------------------"
            
            # Add to size results
            $sizeResults += [PSCustomObject]@{
                FolderPath = $folder
                FolderSize = Format-FileSize -Size $folderSize
                FileCount = $fileCount
            }
            
            # Generate ACL report
            Write-Host "Generating ACL Report..." -ForegroundColor Cyan
            $aclResults += Get-ACLReport $folder
        }
        else {
            Write-Host "Folder not found: $folder" -ForegroundColor Red
        }
    }
    
    # Export results
    $today = Get-Date -Format "yyyyMMddHHmm"
    
    # Export size report
    $sizeReportName = Join-Path $ReportPath "SizeReport_$today.csv"
    $sizeResults | Export-Csv -Path $sizeReportName -NoTypeInformation
    
    # Export ACL report
    $aclReportName = Join-Path $ReportPath "ACLReport_$today.csv"
    $aclResults | Export-Csv -Path $aclReportName -NoTypeInformation -Force -Encoding UTF8
    
    Write-Host "`nReports generated successfully:" -ForegroundColor Green
    Write-Host "Size Report: $sizeReportName" -ForegroundColor Green
    Write-Host "ACL Report: $aclReportName" -ForegroundColor Green
}

# Show the file selection dialog
Show-FileSelectionDialog