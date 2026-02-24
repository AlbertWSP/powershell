# 1. Check and Install Active Directory Module
if (-not (Get-Module -ListAvailable ActiveDirectory)) {
    Write-Host "AD Module missing. Attempting installation..." -ForegroundColor Yellow
    try {
        Add-WindowsCapability -Online -Name Rsat.ActiveDirectory.DS-LDS.Tools~~~~0.0.1.0
        Import-Module ActiveDirectory
    } catch {
        [System.Windows.Forms.MessageBox]::Show("Installation failed. Please install RSAT manually.")
        exit
    }
} else { Import-Module ActiveDirectory }

Add-Type -AssemblyName System.Windows.Forms
Add-Type -AssemblyName System.Drawing

# --- UI Setup ---
$form = New-Object System.Windows.Forms.Form
$form.Text = "AD Group Manager (Multi-Type & CSV Export)"
$form.Size = New-Object System.Drawing.Size(620,920)
$form.StartPosition = "CenterScreen"

# --- UI Elements (Search Section) ---
$labelUser = New-Object System.Windows.Forms.Label
$labelUser.Text = "Current Owner Email:"
$labelUser.Location = New-Object System.Drawing.Point(20,20)
$labelUser.AutoSize = $true
$form.Controls.Add($labelUser)

$inputUser = New-Object System.Windows.Forms.TextBox
$inputUser.Location = New-Object System.Drawing.Point(20,45)
$inputUser.Width = 220
$form.Controls.Add($inputUser)

$btnCheck = New-Object System.Windows.Forms.Button
$btnCheck.Text = "Search"
$btnCheck.Location = New-Object System.Drawing.Point(250,43)
$btnCheck.Width = 70
$form.Controls.Add($btnCheck)

$btnRefresh = New-Object System.Windows.Forms.Button
$btnRefresh.Text = "Refresh"
$btnRefresh.Location = New-Object System.Drawing.Point(330,43)
$btnRefresh.Width = 70
$form.Controls.Add($btnRefresh)

$btnReset = New-Object System.Windows.Forms.Button
$btnReset.Text = "Reset"
$btnReset.Location = New-Object System.Drawing.Point(410,43)
$btnReset.Width = 70
$form.Controls.Add($btnReset)

# --- Selection List ---
$labelList = New-Object System.Windows.Forms.Label
$labelList.Text = "Groups Found: 0"
$labelList.Location = New-Object System.Drawing.Point(20,85)
$labelList.AutoSize = $true
$form.Controls.Add($labelList)

$chkSelectAll = New-Object System.Windows.Forms.CheckBox
$chkSelectAll.Text = "Select All"
$chkSelectAll.Location = New-Object System.Drawing.Point(470,85)
$form.Controls.Add($chkSelectAll)

$outputList = New-Object System.Windows.Forms.CheckedListBox
$outputList.Location = New-Object System.Drawing.Point(20,110)
$outputList.Width = 560
$outputList.Height = 250
$outputList.CheckOnClick = $true
$outputList.DisplayMember = "DisplayText" 
$form.Controls.Add($outputList)

$btnExportPre = New-Object System.Windows.Forms.Button
$btnExportPre.Text = "Export Current List to CSV (Pre-Update)"
$btnExportPre.Location = New-Object System.Drawing.Point(20,365)
$btnExportPre.Width = 250
$form.Controls.Add($btnExportPre)

$labelLegend = New-Object System.Windows.Forms.Label
$labelLegend.Text = "[YES/NO] = Manager Can Update Membership Status"
$labelLegend.Location = New-Object System.Drawing.Point(20,400)
$labelLegend.ForeColor = [System.Drawing.Color]::DimGray
$labelLegend.AutoSize = $true
$form.Controls.Add($labelLegend)

# --- New Owner Section ---
$labelNewOwner = New-Object System.Windows.Forms.Label
$labelNewOwner.Text = "New Owner Email (Can be same as current):"
$labelNewOwner.Location = New-Object System.Drawing.Point(20,435)
$labelNewOwner.AutoSize = $true
$form.Controls.Add($labelNewOwner)

$inputNewOwner = New-Object System.Windows.Forms.TextBox
$inputNewOwner.Location = New-Object System.Drawing.Point(20,460)
$inputNewOwner.Width = 300
$form.Controls.Add($inputNewOwner)

$chkCanUpdate = New-Object System.Windows.Forms.CheckBox
$chkCanUpdate.Text = "Grant 'Update membership' to New Owner"
$chkCanUpdate.Location = New-Object System.Drawing.Point(20,495)
$chkCanUpdate.AutoSize = $true
$form.Controls.Add($chkCanUpdate)

$chkRemovePrev = New-Object System.Windows.Forms.CheckBox
$chkRemovePrev.Text = "REMOVE 'Update membership' from PREVIOUS owner"
$chkRemovePrev.Location = New-Object System.Drawing.Point(20,520)
$chkRemovePrev.ForeColor = [System.Drawing.Color]::DarkRed
$chkRemovePrev.AutoSize = $true
$form.Controls.Add($chkRemovePrev)

$btnUpdate = New-Object System.Windows.Forms.Button
$btnUpdate.Text = "Transfer Selected"
$btnUpdate.Location = New-Object System.Drawing.Point(330, 555)
$btnUpdate.Width = 125
$btnUpdate.Enabled = $false
$form.Controls.Add($btnUpdate)

$btnExportPost = New-Object System.Windows.Forms.Button
$btnExportPost.Text = "Export Final List (Post-Update)"
$btnExportPost.Location = New-Object System.Drawing.Point(20, 555)
$btnExportPost.Width = 200
$btnExportPost.Enabled = $false
$form.Controls.Add($btnExportPost)

# --- Progress Section ---
$labelStatus = New-Object System.Windows.Forms.Label
$currentDate = Get-Date -Format "yyyy-MM-dd"
$labelStatus.Text = "[$currentDate] Status: Ready"
$labelStatus.Location = New-Object System.Drawing.Point(20,600)
$labelStatus.Width = 540
$form.Controls.Add($labelStatus)

$progressBar = New-Object System.Windows.Forms.ProgressBar
$progressBar.Location = New-Object System.Drawing.Point(20,625)
$progressBar.Width = 560
$progressBar.Height = 25
$form.Controls.Add($progressBar)

# --- Logic: Export Function ---
$ExportToCSV = {
    param($suffix)
    $saveDialog = New-Object System.Windows.Forms.SaveFileDialog
    $saveDialog.Filter = "CSV files (*.csv)|*.csv"
    $saveDialog.FileName = "AD_Group_Export_$($suffix)_$(Get-Date -Format 'yyyyMMdd_HHmm').csv"
    if ($saveDialog.ShowDialog() -eq "OK") {
        $exportData = foreach ($item in $outputList.Items) {
            [PSCustomObject]@{
                PermissionStatus = ($item.DisplayText -split '\]')[0].Trim('[')
                Type             = ($item.DisplayText -split '\]')[1].Trim('[ ')
                GroupName        = $item.Name
                Email            = ($item.DisplayText -split '\(')[-1].Trim(')')
                DistinguishedName = $item.DN
            }
        }
        $exportData | Export-Csv -Path $saveDialog.FileName -NoTypeInformation -Encoding UTF8
        [System.Windows.Forms.MessageBox]::Show("File exported successfully!")
    }
}

# --- Logic: Search ---
$DoSearch = {
    $email = $inputUser.Text
    $currentDate = Get-Date -Format "yyyy-MM-dd"
    if ([string]::IsNullOrWhiteSpace($email)) { [System.Windows.Forms.MessageBox]::Show("Please enter an email.") ; return }
    try {
        $user = Get-ADUser -Filter "mail -eq '$email'" -Properties SID, DistinguishedName
        if (-not $user) { [System.Windows.Forms.MessageBox]::Show("User not found.") ; return }
        $userSID = $user.SID
        
        $outputList.Items.Clear()
        $labelStatus.Text = "[$currentDate] Status: Scanning Permissions..."
        [System.Windows.Forms.Application]::DoEvents()

        $groups = Get-ADGroup -Filter "managedBy -eq '$($user.DistinguishedName)'" -Properties Name, mail, DistinguishedName, GroupCategory
        
        if ($groups) {
            $memberAttrGuid = "bf9679c0-0de6-11d0-a285-00aa003049e2"
            foreach ($group in $groups) { 
                $canUpdate = "NO"
                try {
                    $acl = Get-Acl "AD:\$($group.DistinguishedName)"
                    foreach($access in $acl.Access) {
                        try {
                            $ruleSID = $access.IdentityReference.Translate([System.Security.Principal.SecurityIdentifier])
                            if ($ruleSID -eq $userSID) {
                                if ($access.ActiveDirectoryRights -match "WriteProperty" -and 
                                    ($access.ObjectType.ToString() -eq $memberAttrGuid -or $access.ObjectType.ToString() -eq "00000000-0000-0000-0000-000000000000")) {
                                    $canUpdate = "YES"
                                    break
                                }
                            }
                        } catch {}
                    }
                } catch { $canUpdate = "ERR" }

                $groupMail = if ($group.mail) { $group.mail } else { "No Email" }
                $typeTag = if ($group.GroupCategory -eq "Security") { "Sec" } else { "Dist" }
                
                $item = New-Object PSObject -Property @{ 
                    DisplayText = "[$canUpdate] [$typeTag] $($group.Name) ($groupMail)"
                    DN          = $group.DistinguishedName
                    Name        = $group.Name 
                }
                $null = $outputList.Items.Add($item) 
            }
            $labelList.Text = "Groups Found: $($groups.Count)"
            $btnUpdate.Enabled = $true
            $btnExportPost.Enabled = $false
            $labelStatus.Text = "[$currentDate] Status: Scan Complete."
        } else {
            $labelStatus.Text = "[$currentDate] Status: User is not a manager of any groups."
        }
    } catch { [System.Windows.Forms.MessageBox]::Show("Error: $($_.Exception.Message)") }
}

# --- Logic: Update ---
$btnUpdate.Add_Click({
    $prevOwnerEmail = $inputUser.Text
    $newOwnerEmail = $inputNewOwner.Text
    $checkedItems = $outputList.CheckedItems
    
    if ($checkedItems.Count -eq 0) { return }

    try {
        $prevOwner = Get-ADUser -Filter "mail -eq '$prevOwnerEmail'" -Properties SID, DistinguishedName
        $newOwner = Get-ADUser -Filter "mail -eq '$newOwnerEmail'" -Properties SID, DistinguishedName
        
        $msg = "Process $($checkedItems.Count) groups?`nNote: If emails are the same, only permissions will be updated."
        if ([System.Windows.Forms.MessageBox]::Show($msg, "Confirm", "YesNo") -eq "Yes") {
            $progressBar.Maximum = $checkedItems.Count; $progressBar.Value = 0; $counter = 0
            
            foreach ($item in $checkedItems) {
                $counter++; $progressBar.Value = $counter
                $groupDN = $item.DN
                $memberAttrGuid = New-Object Guid "bf9679c0-0de6-11d0-a285-00aa003049e2"
                
                if ($prevOwner.DistinguishedName -ne $newOwner.DistinguishedName) {
                    Set-ADGroup -Identity $groupDN -ManagedBy $newOwner.DistinguishedName
                }

                $acl = Get-Acl "AD:\$groupDN"
                if ($chkRemovePrev.Checked) {
                    $rules = $acl.Access | Where-Object { 
                        try { $_.IdentityReference.Translate([System.Security.Principal.SecurityIdentifier]) -eq $prevOwner.SID } catch { $false }
                    }
                    foreach ($rule in $rules) { $acl.RemoveAccessRule($rule) }
                }

                if ($chkCanUpdate.Checked) {
                    $adRights = [System.DirectoryServices.ActiveDirectoryRights]::WriteProperty
                    $type = [System.Security.AccessControl.AccessControlType]::Allow
                    $rule = New-Object System.DirectoryServices.ActiveDirectoryAccessRule($newOwner.SID, $adRights, $type, $memberAttrGuid)
                    $acl.AddAccessRule($rule)
                }
                Set-Acl -Path "AD:\$groupDN" -AclObject $acl
            }
            [System.Windows.Forms.MessageBox]::Show("Update Process Completed.")
            # Trigger refresh to see new status and enable post-export
            $DoSearch.Invoke()
            $btnExportPost.Enabled = $true
        }
    } catch { [System.Windows.Forms.MessageBox]::Show("Update Failed: $($_.Exception.Message)") }
})

# --- Event Handlers ---
$btnCheck.Add_Click($DoSearch)
$btnRefresh.Add_Click($DoSearch)
$btnReset.Add_Click({ 
    $inputUser.Text = ""; $inputNewOwner.Text = ""; 
    $outputList.Items.Clear(); $labelList.Text = "Groups Found: 0"
    $progressBar.Value = 0; $btnExportPost.Enabled = $false
})
$chkSelectAll.Add_CheckedChanged({
    for($i=0; $i -lt $outputList.Items.Count; $i++) { $outputList.SetItemChecked($i, $chkSelectAll.Checked) }
})

# CSV Export Actions
$btnExportPre.Add_Click({ &$ExportToCSV "PreUpdate" })
$btnExportPost.Add_Click({ &$ExportToCSV "PostUpdate" })

$form.ShowDialog()
