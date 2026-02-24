# 1. Ensure AD Module is loaded
if (-not (Get-Module -ListAvailable ActiveDirectory)) {
    [System.Windows.Forms.MessageBox]::Show("Active Directory module is required. Please install RSAT.")
    exit
}
Import-Module ActiveDirectory

Add-Type -AssemblyName System.Windows.Forms
Add-Type -AssemblyName System.Drawing

# --- Main Form Setup ---
$form = New-Object System.Windows.Forms.Form
$form.Text = "AD Security Group Checker - Version 1.4"
$form.Size = New-Object System.Drawing.Size(850, 950)
$form.StartPosition = "CenterScreen"

$tabControl = New-Object System.Windows.Forms.TabControl
$tabControl.Dock = "Fill"
$form.Controls.Add($tabControl)

# ==========================================
# TAB 1: CHECK MEMBERS
# ==========================================
$tabView = New-Object System.Windows.Forms.TabPage
$tabView.Text = "Check Members"
$tabControl.TabPages.Add($tabView)

$panelHint1 = New-Object System.Windows.Forms.Panel
$panelHint1.Location = "20, 10"; $panelHint1.Size = "790, 132"; $panelHint1.BackColor = "Info"
$tabView.Controls.Add($panelHint1)

$lblHint1 = New-Object System.Windows.Forms.Label
$lblHint1.Text = "FUNCTION: Audit Group Membership.`n`nGUIDELINES:`n1. Enter Group Name or Email in the box below and click 'Fetch'.`n2. Use 'Export CSV' to save results to your preferred folder.`n3. QUICK ACTION: You can select one or multiple users in the search result table below, then RIGHT-CLICK and select 'Add selected to Modify list'.`n4. This will automatically transfer their SAM Account Names to the 'Modify Members' tab for bulk processing."
$lblHint1.Dock = "Fill"; $lblHint1.Padding = "10, 10, 10, 10"; $lblHint1.Font = New-Object System.Drawing.Font("Segoe UI", 9, [System.Drawing.FontStyle]::Bold)
$panelHint1.Controls.Add($lblHint1)

$lblGroupSearch = New-Object System.Windows.Forms.Label
$lblGroupSearch.Text = "Group Name/Email:"; $lblGroupSearch.Location = "20, 150"; $lblGroupSearch.AutoSize = $true
$tabView.Controls.Add($lblGroupSearch)

$txtGroupSearch = New-Object System.Windows.Forms.TextBox
$txtGroupSearch.Location = "20, 170"; $txtGroupSearch.Width = 350
$tabView.Controls.Add($txtGroupSearch)

$btnFetch = New-Object System.Windows.Forms.Button
$btnFetch.Text = "Fetch"; $btnFetch.Location = "380, 168"; $btnFetch.Width = 80
$tabView.Controls.Add($btnFetch)

$btnExport = New-Object System.Windows.Forms.Button
$btnExport.Text = "Export CSV"; $btnExport.Location = "470, 168"; $btnExport.Width = 100; $btnExport.Enabled = $false
$tabView.Controls.Add($btnExport)

$btnClearTab1 = New-Object System.Windows.Forms.Button
$btnClearTab1.Text = "Clear All"; $btnClearTab1.Location = "580, 168"; $btnClearTab1.Width = 100
$tabView.Controls.Add($btnClearTab1)

$lblCount = New-Object System.Windows.Forms.Label
$lblCount.Text = "Total Members: 0"; $lblCount.Location = "20, 200"; $lblCount.ForeColor = "Blue"; $lblCount.AutoSize = $true
$tabView.Controls.Add($lblCount)

$dgvMembers = New-Object System.Windows.Forms.DataGridView
$dgvMembers.Location = "20, 225"; $dgvMembers.Size = "790, 640"
$dgvMembers.ReadOnly = $true; $dgvMembers.AutoSizeColumnsMode = "Fill"
$dgvMembers.SelectionMode = "FullRowSelect"
$tabView.Controls.Add($dgvMembers)

# Context Menu
$contextMenu = New-Object System.Windows.Forms.ContextMenuStrip
$menuItemAdd = $contextMenu.Items.Add("Add selected to Modify list")
$menuItemAdd.Add_Click({
    if ($dgvMembers.SelectedRows.Count -gt 0) {
        $count = 0
        foreach ($row in $dgvMembers.SelectedRows) {
            $sam = $row.Cells["SamAccountName"].Value.ToString()
            if ($sam) { $null = $dgvInput.Rows.Add($sam); $count++ }
        }
        $tabControl.SelectedTab = $tabUpdate
        $txtUpdateLog.AppendText("`r`n[INFO] Transferred $count users from Check tab.")
    }
})
$dgvMembers.ContextMenuStrip = $contextMenu

# Fetch Logic
$btnFetch.Add_Click({
    try {
        $group = Get-ADGroup -Filter "Name -eq '$($txtGroupSearch.Text)' -or mail -eq '$($txtGroupSearch.Text)'" -ErrorAction Stop
        $members = Get-ADGroupMember -Identity $group | Get-ADUser -Properties mail, DisplayName, Department, Title | 
                   Select-Object DisplayName, SamAccountName, mail, Department, Title
        $dgvMembers.DataSource = [System.Collections.ArrayList]$members
        $lblCount.Text = "Total Members: $($members.Count)"
        $btnExport.Enabled = ($members.Count -gt 0)
        $script:currentData = $members
    } catch { [System.Windows.Forms.MessageBox]::Show("Group not found.") }
})

$btnClearTab1.Add_Click({ $txtGroupSearch.Text = ""; $dgvMembers.DataSource = $null; $lblCount.Text = "Total Members: 0"; $btnExport.Enabled = $false })

$btnExport.Add_Click({
    $save = New-Object System.Windows.Forms.SaveFileDialog
    $save.Filter = "CSV Files (*.csv)|*.csv"
    if ($save.ShowDialog() -eq [System.Windows.Forms.DialogResult]::OK) { $script:currentData | Export-Csv -Path $save.FileName -NoTypeInformation -Encoding UTF8 }
})

# ==========================================
# TAB 2: MODIFY MEMBERS
# ==========================================
$tabUpdate = New-Object System.Windows.Forms.TabPage
$tabUpdate.Text = "Modify Members"
$tabControl.TabPages.Add($tabUpdate)

$panelHint2 = New-Object System.Windows.Forms.Panel
$panelHint2.Location = "20, 10"; $panelHint2.Size = "790, 150"; $panelHint2.BackColor = "Info"
$tabUpdate.Controls.Add($panelHint2)

$lblHint2 = New-Object System.Windows.Forms.Label
$lblHint2.Text = "FUNCTION: Bulk Membership Updates.`n`nGUIDELINES:`n1. Enter the exact 'Target Security Group' name or email.`n2. Populate the input table via: Manual entry, uploading a CSV/TXT file (Header-less), or using the Right-Click transfer from the previous tab.`n3. Important: If importing CSV, ensure there are no column headers.`n4. Click 'Bulk ADD' or 'Bulk REMOVE'. You will be prompted to choose a save location for the Modification Log.`n5. Review the 'Update Log' at the bottom for real-time success/failure status."
$lblHint2.Dock = "Fill"; $lblHint2.Padding = "10, 10, 10, 10"; $lblHint2.Font = New-Object System.Drawing.Font("Segoe UI", 9, [System.Drawing.FontStyle]::Bold)
$panelHint2.Controls.Add($lblHint2)

$lblTargetGroup = New-Object System.Windows.Forms.Label
$lblTargetGroup.Text = "Target Group:"; $lblTargetGroup.Location = "20, 175"; $lblTargetGroup.AutoSize = $true
$tabUpdate.Controls.Add($lblTargetGroup)

$txtTargetGroup = New-Object System.Windows.Forms.TextBox
$txtTargetGroup.Location = "20, 195"; $txtTargetGroup.Width = 400
$tabUpdate.Controls.Add($txtTargetGroup)

$btnImportFile = New-Object System.Windows.Forms.Button
$btnImportFile.Text = "Upload CSV/TXT"; $btnImportFile.Location = "430, 193"; $btnImportFile.Width = 130
$tabUpdate.Controls.Add($btnImportFile)

# --- FIX: CORRECT GRID INITIALIZATION ---
$dgvInput = New-Object System.Windows.Forms.DataGridView
$dgvInput.Location = "20, 230"; $dgvInput.Size = "790, 220"
$dgvInput.ColumnCount = 1
$dgvInput.Columns[0].Name = "User Identifier (Email or SAM)"
$dgvInput.Columns[0].Width = 700
$tabUpdate.Controls.Add($dgvInput)

$btnAdd = New-Object System.Windows.Forms.Button
$btnAdd.Text = "Bulk ADD"; $btnAdd.Location = "20, 465"; $btnAdd.Width = 120; $btnAdd.BackColor = "LightGreen"
$tabUpdate.Controls.Add($btnAdd)

$btnRemove = New-Object System.Windows.Forms.Button
$btnRemove.Text = "Bulk REMOVE"; $btnRemove.Location = "150, 465"; $btnRemove.Width = 120; $btnRemove.BackColor = "LightCoral"
$tabUpdate.Controls.Add($btnRemove)

$btnClearTab2 = New-Object System.Windows.Forms.Button
$btnClearTab2.Text = "Clear All"; $btnClearTab2.Location = "690, 465"; $btnClearTab2.Width = 120
$tabUpdate.Controls.Add($btnClearTab2)

$txtUpdateLog = New-Object System.Windows.Forms.TextBox
$txtUpdateLog.Location = "20, 500"; $txtUpdateLog.Size = "790, 360"; $txtUpdateLog.Multiline = $true; $txtUpdateLog.ReadOnly = $true; $txtUpdateLog.ScrollBars = "Vertical"
$tabUpdate.Controls.Add($txtUpdateLog)

# Import/Process Logic
$btnImportFile.Add_Click({
    $open = New-Object System.Windows.Forms.OpenFileDialog
    $open.Filter = "CSV or Text Files (*.csv;*.txt)|*.csv;*.txt"
    if ($open.ShowDialog() -eq [System.Windows.Forms.DialogResult]::OK) {
        $content = Get-Content $open.FileName | Where-Object { $_.Trim() -ne "" }
        $headerWords = @("email", "samaccountname", "sam", "user", "identifier", "username")
        $startIndex = if ($headerWords -contains $content.Trim().ToLower()) { 1 } else { 0 }
        for ($i = $startIndex; $i -lt $content.Count; $i++) { $null = $dgvInput.Rows.Add($content[$i].Trim()) }
    }
})

$ProcessGrid = {
    param($Action)
    $groupName = $txtTargetGroup.Text
    $userList = New-Object System.Collections.Generic.List[string]
    foreach($row in $dgvInput.Rows) { if ($row.Cells.Value) { $userList.Add($row.Cells.Value.ToString().Trim()) } }
    $users = $userList | Select-Object -Unique
    if (!$groupName -or $users.Count -eq 0) { [System.Windows.Forms.MessageBox]::Show("Enter group and users."); return }
    $saveLog = New-Object System.Windows.Forms.SaveFileDialog
    $saveLog.Filter = "Text Files (*.txt)|*.txt"; $saveLog.FileName = "BulkUpdate_Log_$(Get-Date -Format 'yyyyMMdd_HHmm').txt"
    if ($saveLog.ShowDialog() -eq [System.Windows.Forms.DialogResult]::OK) {
        $logPath = $saveLog.FileName
        try {
            $group = Get-ADGroup -Filter "Name -eq '$groupName' -or mail -eq '$groupName'" -ErrorAction Stop
            foreach($u in $users) {
                try {
                    $adUser = Get-ADUser -Filter "mail -eq '$u' -or SamAccountName -eq '$u'" -ErrorAction Stop
                    if ($Action -eq "Add") { Add-ADGroupMember $group $adUser -ErrorAction Stop; $s = "[SUCCESS] Added: $u" }
                    else { Remove-ADGroupMember $group $adUser -Confirm:$false -ErrorAction Stop; $s = "[SUCCESS] Removed: $u" }
                } catch { $s = "[ERROR] Not Found: $u" }
                $txtUpdateLog.AppendText("`r`n$s"); $s | Out-File $logPath -Append
                [System.Windows.Forms.Application]::DoEvents()
            }
            [System.Windows.Forms.MessageBox]::Show("Process Finished.")
        } catch { [System.Windows.Forms.MessageBox]::Show("Group not found.") }
    }
}

$btnAdd.Add_Click({ &$ProcessGrid "Add" })
$btnRemove.Add_Click({ &$ProcessGrid "Remove" })
$btnClearTab2.Add_Click({ $txtTargetGroup.Text=""; $dgvInput.Rows.Clear(); $txtUpdateLog.Clear() })

$form.ShowDialog()
