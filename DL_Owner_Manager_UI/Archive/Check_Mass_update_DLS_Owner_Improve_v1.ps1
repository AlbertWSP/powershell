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
$form.Text = "AD Group Manager (Permission Force Update)"
$form.Size = New-Object System.Drawing.Size(600,880)
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
$outputList.Width = 540
$outputList.Height = 250
$outputList.CheckOnClick = $true
$outputList.DisplayMember = "DisplayText" 
$form.Controls.Add($outputList)

$labelLegend = New-Object System.Windows.Forms.Label
$labelLegend.Text = "[YES/NO] = Manager Can Update Membership Status"
$labelLegend.Location = New-Object System.Drawing.Point(20,360)
$labelLegend.ForeColor = [System.Drawing.Color]::DimGray
$labelLegend.AutoSize = $true
$form.Controls.Add($labelLegend)

$labelSelected = New-Object System.Windows.Forms.Label
$labelSelected.Text = "Selected: 0"
$labelSelected.Location = New-Object System.Drawing.Point(20,380)
$labelSelected.ForeColor = [System.Drawing.Color]::Blue
$labelSelected.AutoSize = $true
$form.Controls.Add($labelSelected)

# --- New Owner Section ---
$labelNewOwner = New-Object System.Windows.Forms.Label
$labelNewOwner.Text = "New Owner Email (Can be same as current):"
$labelNewOwner.Location = New-Object System.Drawing.Point(20,415)
$labelNewOwner.AutoSize = $true
$form.Controls.Add($labelNewOwner)

$inputNewOwner = New-Object System.Windows.Forms.TextBox
$inputNewOwner.Location = New-Object System.Drawing.Point(20,440)
$inputNewOwner.Width = 300
$form.Controls.Add($inputNewOwner)

$chkCanUpdate = New-Object System.Windows.Forms.CheckBox
$chkCanUpdate.Text = "Grant 'Update membership' to New Owner"
$chkCanUpdate.Location = New-Object System.Drawing.Point(20,475)
$chkCanUpdate.AutoSize = $true
$form.Controls.Add($chkCanUpdate)

$chkRemovePrev = New-Object System.Windows.Forms.CheckBox
$chkRemovePrev.Text = "REMOVE 'Update membership' from PREVIOUS owner"
$chkRemovePrev.Location = New-Object System.Drawing.Point(20,500)
$chkRemovePrev.ForeColor = [System.Drawing.Color]::DarkRed
$chkRemovePrev.AutoSize = $true
$form.Controls.Add($chkRemovePrev)

$btnUpdate = New-Object System.Windows.Forms.Button
$btnUpdate.Text = "Transfer Selected"
$btnUpdate.Location = New-Object System.Drawing.Point(330, 535)
$btnUpdate.Width = 125
$btnUpdate.Enabled = $false
$form.Controls.Add($btnUpdate)

# --- Progress Section ---
$labelStatus = New-Object System.Windows.Forms.Label
$currentDate = Get-Date -Format "yyyy-MM-dd"
$labelStatus.Text = "[$currentDate] Status: Ready"
$labelStatus.Location = New-Object System.Drawing.Point(20,580)
$labelStatus.Width = 540
$form.Controls.Add($labelStatus)

$progressBar = New-Object System.Windows.Forms.ProgressBar
$progressBar.Location = New-Object System.Drawing.Point(20,605)
$progressBar.Width = 540
$progressBar.Height = 25
$form.Controls.Add($progressBar)

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
        $labelStatus.Text = "[$currentDate] Status: Scanning Permissions (Security & Distribution)..."
        [System.Windows.Forms.Application]::DoEvents()

        # Search all groups where user is the primary manager (regardless of Category)
        $groups = Get-ADGroup -Filter "managedBy -eq '$($user.DistinguishedName)'" -Properties Name, mail, DistinguishedName, GroupCategory
        
        if ($groups) {
            $memberAttrGuid = "bf9679c0-0de6-11d0-a285-00aa003049e2"
            foreach ($group in $groups) { 
                $canUpdate = "NO"
                try {
                    $acl = Get-Acl "AD:\$($group.DistinguishedName)"
                    foreach($access in $acl.Access) {
                        try {
                            # Translate Identity to SID for accurate matching
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
                
                # 1. Update Primary Manager
                if ($prevOwner.DistinguishedName -ne $newOwner.DistinguishedName) {
                    Set-ADGroup -Identity $groupDN -ManagedBy $newOwner.DistinguishedName
                }

                # 2. Update Permissions (ACL)
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
            $DoSearch.Invoke()
        }
    } catch { [System.Windows.Forms.MessageBox]::Show("Update Failed: $($_.Exception.Message)") }
})

# --- Event Handlers ---
$btnCheck.Add_Click($DoSearch)
$btnRefresh.Add_Click($DoSearch)
$btnReset.Add_Click({ 
    $inputUser.Text = ""; $inputNewOwner.Text = ""; 
    $outputList.Items.Clear(); $labelList.Text = "Groups Found: 0"
    $progressBar.Value = 0
})
$chkSelectAll.Add_CheckedChanged({
    for($i=0; $i -lt $outputList.Items.Count; $i++) { $outputList.SetItemChecked($i, $chkSelectAll.Checked) }
})

$form.ShowDialog()
