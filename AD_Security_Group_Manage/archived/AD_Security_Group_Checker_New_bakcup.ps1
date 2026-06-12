# =================================================================================
# 1. Self-Elevation (Request Administrator Privileges)
# =================================================================================
if (-not ([Security.Principal.WindowsPrincipal][Security.Principal.WindowsIdentity]::GetCurrent()).IsInRole([Security.Principal.WindowsBuiltInRole]::Administrator)) {
    $newProcess = New-Object System.Diagnostics.ProcessStartInfo "PowerShell"
    $newProcess.Arguments = "-NoProfile -ExecutionPolicy Bypass -File `"$PSCommandPath`""
    $newProcess.Verb = "runas"
    try {
        [System.Diagnostics.Process]::Start($newProcess) | Out-Null
    } catch {
        # User cancelled UAC
    }
    exit
}

# =================================================================================
# 2. Module & Group Membership Check (GRP-RBC-R-ASI-WSP-ClientSideSupport)
# =================================================================================
Add-Type -AssemblyName System.Windows.Forms
Add-Type -AssemblyName System.Drawing

if (-not (Get-Module -ListAvailable ActiveDirectory)) {
    [System.Windows.Forms.MessageBox]::Show("Active Directory module is required. Please install RSAT.")
    exit
}
Import-Module ActiveDirectory

$supportGroupName = "GRP-RBC-R-ASI-WSP-ClientSideSupport"
$isSupportMember = $false

try {
    $currentUserSam = [Security.Principal.WindowsIdentity]::GetCurrent().Name.Split('\')[-1]
    $groupMembers = Get-ADGroupMember -Identity $supportGroupName -Recursive | Select-Object -ExpandProperty SamAccountName
    if ($groupMembers -contains $currentUserSam) { $isSupportMember = $true }
} catch {}

if (-not $isSupportMember) {
    $msgText = "PERMISSION WARNING:`n`nYou are currently NOT a member of '$supportGroupName'.`n`nYou may 'Search' groups, but 'Modify' operations will likely fail due to restricted access.`n`nDo you want to continue?"
    if ([System.Windows.Forms.MessageBox]::Show($msgText, "Access Check", "YesNo", "Warning") -eq "No") { exit }
}

# =================================================================================
# 3. UI Interface Code (Version 1.7 - English Edition)
# =================================================================================
$form = New-Object System.Windows.Forms.Form
$form.Text = "AD Security Group Manager - Version 1.7"
$form.Size = New-Object System.Drawing.Size(850, 1000)
$form.StartPosition = "CenterScreen"

# Main Container
$tabControl = New-Object System.Windows.Forms.TabControl
$tabControl.Size = New-Object System.Drawing.Size(830, 880)
$tabControl.Location = New-Object System.Drawing.Point(5, 5)
$form.Controls.Add($tabControl)

# --- TAB 1: AUDIT MEMBERS ---
$tabView = New-Object System.Windows.Forms.TabPage
$tabView.Text = "Audit Members"
$tabControl.TabPages.Add($tabView)

$panelHint1 = New-Object System.Windows.Forms.Panel
$panelHint1.Location = "20, 10"; $panelHint1.Size = "790, 130"; $panelHint1.BackColor = "Info"
$tabView.Controls.Add($panelHint1)

$lblHint1 = New-Object System.Windows.Forms.Label
$lblHint1.Text = "FUNCTION: Review Group Membership.`n`nGUIDELINES:`n1. Enter Group Name or Email and click 'Fetch'.`n2. Use 'Export CSV' to save the current list.`n3. QUICK ACTION: Select row(s), RIGHT-CLICK, and choose 'Add selected to Modify list' to transfer users to the next tab."
$lblHint1.Dock = "Fill"; $lblHint1.Padding = "10, 10, 10, 10"; $lblHint1.Font = New-Object System.Drawing.Font("Segoe UI", 9, [System.Drawing.FontStyle]::Bold)
$panelHint1.Controls.Add($lblHint1)

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
$dgvMembers.Location = "20, 225"; $dgvMembers.Size = "785, 620"
$dgvMembers.ReadOnly = $true; $dgvMembers.AutoSizeColumnsMode = "Fill"
$dgvMembers.SelectionMode = "FullRowSelect"
$tabView.Controls.Add($dgvMembers)

# --- TAB 2: BULK UPDATES ---
$tabUpdate = New-Object System.Windows.Forms.TabPage
$tabUpdate.Text = "Bulk Updates"
$tabControl.TabPages.Add($tabUpdate)

$panelHint2 = New-Object System.Windows.Forms.Panel
$panelHint2.Location = "20, 10"; $panelHint2.Size = "790, 150"; $panelHint2.BackColor = "Info"
$tabUpdate.Controls.Add($panelHint2)

$lblHint2 = New-Object System.Windows.Forms.Label
$lblHint2.Text = "FUNCTION: Bulk Membership Management.`n`nGUIDELINES:`n1. Enter the 'Target Security Group'.`n2. Populate identifiers (Email/SAM) via Import or Right-Click from the Audit tab.`n3. Select 'Bulk ADD' or 'Bulk REMOVE'.`n4. A log file will be generated for all transactions."
$lblHint2.Dock = "Fill"; $lblHint2.Padding = "10, 10, 10, 10"; $lblHint2.Font = New-Object System.Drawing.Font("Segoe UI", 9, [System.Drawing.FontStyle]::Bold)
$panelHint2.Controls.Add($lblHint2)

$txtTargetGroup = New-Object System.Windows.Forms.TextBox
$txtTargetGroup.Location = "20, 195"; $txtTargetGroup.Width = 400
$tabUpdate.Controls.Add($txtTargetGroup)

$btnImportFile = New-Object System.Windows.Forms.Button
$btnImportFile.Text = "Upload CSV/TXT"; $btnImportFile.Location = "430, 193"; $btnImportFile.Width = 130
$tabUpdate.Controls.Add($btnImportFile)

$dgvInput = New-Object System.Windows.Forms.DataGridView
$dgvInput.Location = "20, 230"; $dgvInput.Size = "785, 220"
$col1 = New-Object System.Windows.Forms.DataGridViewTextBoxColumn
$col1.HeaderText = "User Identifier (Email or SAM)"; $col1.Name = "colIdentifier"; $col1.Width = 700
$dgvInput.Columns.Add($col1) | Out-Null
$tabUpdate.Controls.Add($dgvInput)

$btnAdd = New-Object System.Windows.Forms.Button
$btnAdd.Text = "Bulk ADD"; $btnAdd.Location = "20, 465"; $btnAdd.Width = 120; $btnAdd.BackColor = "LightGreen"
$tabUpdate.Controls.Add($btnAdd)

$btnRemove = New-Object System.Windows.Forms.Button
$btnRemove.Text = "Bulk REMOVE"; $btnRemove.Location = "150, 465"; $btnRemove.Width = 120; $btnRemove.BackColor = "LightCoral"
$tabUpdate.Controls.Add($btnRemove)

$btnClearTab2 = New-Object System.Windows.Forms.Button
$btnClearTab2.Text = "Clear List"; $btnClearTab2.Location = "685, 465"; $btnClearTab2.Width = 120
$tabUpdate.Controls.Add($btnClearTab2)

$txtUpdateLog = New-Object System.Windows.Forms.TextBox
$txtUpdateLog.Location = "20, 500"; $txtUpdateLog.Size = "785, 330"; $txtUpdateLog.Multiline = $true; $txtUpdateLog.ReadOnly = $true; $txtUpdateLog.ScrollBars = "Vertical"
$tabUpdate.Controls.Add($txtUpdateLog)

# --- TAB 3: ACL VIEWER ---
$tabACL = New-Object System.Windows.Forms.TabPage
$tabACL.Text = "ACL Viewer"
$tabControl.TabPages.Add($tabACL)

$lblPath = New-Object System.Windows.Forms.Label
$lblPath.Location = "20, 10"; $lblPath.Size = "100, 20"; $lblPath.Text = "Folder Path:"
$tabACL.Controls.Add($lblPath)

$txtPath = New-Object System.Windows.Forms.TextBox
$txtPath.Location = "20, 35"; $txtPath.Size = "680, 25"
$tabACL.Controls.Add($txtPath)

$btnBrowsePath = New-Object System.Windows.Forms.Button
$btnBrowsePath.Location = "710, 33"; $btnBrowsePath.Size = "100, 28"; $btnBrowsePath.Text = "Browse..."
$btnBrowsePath.Add_Click({
    $browser = New-Object System.Windows.Forms.FolderBrowserDialog
    if ($browser.ShowDialog() -eq [System.Windows.Forms.DialogResult]::OK) { $txtPath.Text = $browser.SelectedPath }
})
$tabACL.Controls.Add($btnBrowsePath)

$btnCheckAcl = New-Object System.Windows.Forms.Button
$btnCheckAcl.Location = "20, 70"; $btnCheckAcl.Size = "150, 35"; $btnCheckAcl.Text = "Check Permissions"; $btnCheckAcl.BackColor = "LightGreen"
$tabACL.Controls.Add($btnCheckAcl)

$dgvAcl = New-Object System.Windows.Forms.DataGridView
$dgvAcl.Location = "20, 120"; $dgvAcl.Size = "785, 360"
$dgvAcl.ReadOnly = $true; $dgvAcl.AutoSizeColumnsMode = "AllCells"; $dgvAcl.BackgroundColor = "White"; $dgvAcl.RowHeadersVisible = $false
$tabACL.Controls.Add($dgvAcl)

# Enable copying from DataGridView
$dgvAcl.MultiSelect = $true
$dgvAcl.SelectionMode = 'FullRowSelect'
$dgvAcl.AllowUserToAddRows = $false
$dgvAcl.ClipboardCopyMode = [System.Windows.Forms.DataGridViewClipboardCopyMode]::EnableAlwaysIncludeHeaderText

# Context menu for copy actions
$aclContext = New-Object System.Windows.Forms.ContextMenuStrip
$miCopySel = $aclContext.Items.Add("Copy Selected")
$miCopyAll = $aclContext.Items.Add("Copy All")
$miCopyOwner = $aclContext.Items.Add("Copy Owner")

$miCopySel.Add_Click({
    $clip = $dgvAcl.GetClipboardContent()
    if ($clip) { [System.Windows.Forms.Clipboard]::SetDataObject($clip) }
})

$miCopyAll.Add_Click({
    $dgvAcl.SelectAll()
    $clip = $dgvAcl.GetClipboardContent()
    if ($clip) { [System.Windows.Forms.Clipboard]::SetDataObject($clip) }
    $dgvAcl.ClearSelection()
})

$miCopyOwner.Add_Click({
    if ($lblAclOwner.Text) { [System.Windows.Forms.Clipboard]::SetText($lblAclOwner.Text) }
})

$dgvAcl.ContextMenuStrip = $aclContext

# Support Ctrl+C to copy selected cells
$dgvAcl.Add_KeyDown({ param($s,$e) if ($e.Control -and $e.KeyCode -eq [System.Windows.Forms.Keys]::C) { $clip = $dgvAcl.GetClipboardContent(); if ($clip) { [System.Windows.Forms.Clipboard]::SetDataObject($clip) } } })

$lblAclOwner = New-Object System.Windows.Forms.Label
$lblAclOwner.Location = "20, 500"; $lblAclOwner.Size = "800, 25"; $lblAclOwner.ForeColor = "DarkBlue"
$tabACL.Controls.Add($lblAclOwner)

$btnCheckAcl.Add_Click({
    $path = $txtPath.Text
    if (Test-Path $path) {
        try {
            $acl = Get-Acl -Path $path
            $lblAclOwner.Text = "Owner: $($acl.Owner)"
            $results = $acl.Access | Select-Object IdentityReference, FileSystemRights, AccessControlType
            $dt = New-Object System.Data.DataTable
            $dt.Columns.Add("IdentityReference") | Out-Null
            $dt.Columns.Add("FileSystemRights") | Out-Null
            $dt.Columns.Add("AccessControlType") | Out-Null
            foreach ($item in $results) {
                $row = $dt.NewRow()
                $row.IdentityReference = $item.IdentityReference.ToString()
                $row.FileSystemRights = $item.FileSystemRights.ToString()
                $row.AccessControlType = $item.AccessControlType.ToString()
                $dt.Rows.Add($row)
            }
            $dgvAcl.DataSource = $dt
        } catch {
            [System.Windows.Forms.MessageBox]::Show("Access Denied or Error: $($_.Exception.Message)")
        }
    } else {
        [System.Windows.Forms.MessageBox]::Show("Invalid Path!")
    }
})

# =================================================================================
# GLOBAL EXIT BUTTON
# =================================================================================
$btnExit = New-Object System.Windows.Forms.Button
$btnExit.Text = "EXIT PROGRAM"
$btnExit.Location = New-Object System.Drawing.Point(30, 900)
$btnExit.Size = New-Object System.Drawing.Size(770, 45)
$btnExit.BackColor = "Gainsboro"
$btnExit.Font = New-Object System.Drawing.Font("Segoe UI", 10, [System.Drawing.FontStyle]::Bold)
$btnExit.Add_Click({
    if ([System.Windows.Forms.MessageBox]::Show("Are you sure you want to exit?", "Exit Confirmation", "YesNo", "Question") -eq "Yes") {
        $form.Close()
    }
})
$form.Controls.Add($btnExit)

# =================================================================================
# 4. Logic Handling
# =================================================================================

# Context Menu
$contextMenu = New-Object System.Windows.Forms.ContextMenuStrip
$menuItemAdd = $contextMenu.Items.Add("Add selected to Modify list")
$menuItemAdd.Add_Click({
    if ($dgvMembers.SelectedRows.Count -gt 0) {
        foreach ($row in $dgvMembers.SelectedRows) {
            $sam = $row.Cells["SamAccountName"].Value.ToString()
            if ($sam) { $null = $dgvInput.Rows.Add($sam) }
        }
        $tabControl.SelectedTab = $tabUpdate
    }
})
$dgvMembers.ContextMenuStrip = $contextMenu

# Tab 1 Fetch
$btnFetch.Add_Click({
    try {
        $group = Get-ADGroup -Filter "Name -eq '$($txtGroupSearch.Text)' -or mail -eq '$($txtGroupSearch.Text)'" -ErrorAction Stop
        $members = Get-ADGroupMember -Identity $group | Get-ADUser -Properties mail, DisplayName, Department, Title | 
                   Select-Object DisplayName, SamAccountName, mail, Department, Title
        $dgvMembers.DataSource = [System.Collections.ArrayList]$members
        $lblCount.Text = "Total Members: $($members.Count)"
        $btnExport.Enabled = ($members.Count -gt 0)
        $script:currentData = $members
    } catch { [System.Windows.Forms.MessageBox]::Show("Group not found or AD access error.", "Error") }
})

# Tab 2 Process
$ProcessGrid = {
    param($Action)
    $groupName = $txtTargetGroup.Text
    $userList = New-Object System.Collections.Generic.List[string]
    foreach($row in $dgvInput.Rows) { if ($row.Cells["colIdentifier"].Value) { $userList.Add($row.Cells["colIdentifier"].Value.ToString().Trim()) } }
    $users = $userList | Select-Object -Unique
    if (!$groupName -or $users.Count -eq 0) { [System.Windows.Forms.MessageBox]::Show("Please provide both Target Group and User identifiers.", "Validation Error") ; return }
    
    $saveLog = New-Object System.Windows.Forms.SaveFileDialog
    $saveLog.Filter = "Text Files (*.txt)|*.txt"; $saveLog.FileName = "BulkUpdate_Log_$(Get-Date -Format 'yyyyMMdd').txt"
    
    if ($saveLog.ShowDialog() -eq [System.Windows.Forms.DialogResult]::OK) {
        $logPath = $saveLog.FileName
        try {
            $group = Get-ADGroup -Filter "Name -eq '$groupName' -or mail -eq '$groupName'" -ErrorAction Stop
            foreach($u in $users) {
                try {
                    $adUser = Get-ADUser -Filter "mail -eq '$u' -or SamAccountName -eq '$u'" -ErrorAction Stop
                    if ($Action -eq "Add") { Add-ADGroupMember $group $adUser -ErrorAction Stop; $s = "[SUCCESS] Added: $u" }
                    else { Remove-ADGroupMember $group $adUser -Confirm:$false -ErrorAction Stop; $s = "[SUCCESS] Removed: $u" }
                } catch { $s = "[ERROR] $u : $($_.Exception.Message)" }
                $txtUpdateLog.AppendText("`r`n$s"); $s | Out-File $logPath -Append
                $txtUpdateLog.ScrollToCaret()
                [System.Windows.Forms.Application]::DoEvents()
            }
            [System.Windows.Forms.MessageBox]::Show("Processing Finished.", "Complete")
        } catch { [System.Windows.Forms.MessageBox]::Show("Target Group resolution failed.", "Error") }
    }
}

$btnAdd.Add_Click({ &$ProcessGrid "Add" })
$btnRemove.Add_Click({ &$ProcessGrid "Remove" })
$btnClearTab1.Add_Click({ $txtGroupSearch.Text = ""; $dgvMembers.DataSource = $null; $lblCount.Text = "Total Members: 0" })
$btnClearTab2.Add_Click({ $txtTargetGroup.Text=""; $dgvInput.Rows.Clear(); $txtUpdateLog.Clear() })

$form.ShowDialog()
