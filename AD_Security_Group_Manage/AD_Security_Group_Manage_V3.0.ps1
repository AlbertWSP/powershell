# 2026-03-31
# @Albert Ng
# AD Security Group Manager - Version 3.0

# =================================================================================
# 1. 自動提升為管理員權限 (Self-Elevation)
# 若目前非管理員，則以 runas 重新啟動腳本，確保能完整讀取 ACL 或異動 AD 群組
# =================================================================================
if (-not ([Security.Principal.WindowsPrincipal][Security.Principal.WindowsIdentity]::GetCurrent()).IsInRole([Security.Principal.WindowsBuiltInRole]::Administrator)) {
    $newProcess = New-Object System.Diagnostics.ProcessStartInfo "PowerShell"
    $newProcess.Arguments = "-NoProfile -ExecutionPolicy Bypass -File `"$PSCommandPath`""
    $newProcess.Verb = "runas"
    try {
        [System.Diagnostics.Process]::Start($newProcess) | Out-Null
    }
    catch {
        # 使用者取消 UAC 提示
    }
    exit
}

# =================================================================================
# 2. 模組檢查與身份驗證 (Module & Group Membership Check)
# =================================================================================
Add-Type -AssemblyName System.Windows.Forms
Add-Type -AssemblyName System.Drawing

# 檢查是否具備 ActiveDirectory 模組 (通常需要安裝 RSAT)
if (-not (Get-Module -ListAvailable ActiveDirectory)) {
    [System.Windows.Forms.MessageBox]::Show("Active Directory module is required. Please install RSAT.")
    exit
}
Import-Module ActiveDirectory

# --- Helper Functions (v2.0) 輔助函式區 ---

# [輔助函式] 依據名稱或 Email 解析 AD 群組
function Resolve-ADGroupByNameOrMail {
    param([string]$SearchValue)
    $val = $SearchValue.Trim()
    Get-ADGroup -Filter { Name -eq $val -or mail -eq $val } -ErrorAction Stop
}

# [輔助函式] 依據 Email 或 SAM 帳號解析 AD 使用者
function Resolve-ADUserByMailOrSam {
    param([string]$SearchValue)
    $val = $SearchValue.Trim()
    Get-ADUser -Filter { mail -eq $val -or SamAccountName -eq $val } -ErrorAction Stop
}

# [輔助函式] 更新底部狀態列文字並強制 UI 重繪
function Update-Status {
    param([string]$Message)
    $statusLabel.Text = $Message
    [System.Windows.Forms.Application]::DoEvents()
}


# 驗證執行者是否在指定的 Support 群組中
$supportGroupName = "GRP-RBC-R-ASI-WSP-ClientSideSupport"
$isSupportMember = $false
try {
    $currentUserSam = [Security.Principal.WindowsIdentity]::GetCurrent().Name.Split('\')[-1]
    $groupMembers = Get-ADGroupMember -Identity $supportGroupName -Recursive | Select-Object -ExpandProperty SamAccountName
    if ($groupMembers -contains $currentUserSam) { $isSupportMember = $true }
}
catch {
    [System.Windows.Forms.MessageBox]::Show("Warning: Could not verify membership in '$supportGroupName'.`n`nError: $($_.Exception.Message)", "Membership Check Warning", "OK", "Warning")
}

if (-not $isSupportMember) {
    $msgText = "PERMISSION WARNING:`n`nYou are currently NOT a member of '$supportGroupName'.`n`nYou may 'Search' groups, but 'Modify' operations will likely fail due to restricted access.`n`nDo you want to continue?"
    if ([System.Windows.Forms.MessageBox]::Show($msgText, "Access Check", "YesNo", "Warning") -eq "No") { exit }
}

# =================================================================================
# 3. 建立主視窗 (UI Interface Code)
# =================================================================================
$form = New-Object System.Windows.Forms.Form
$form.Text = "AD Security Group Manager"
$form.Size = New-Object System.Drawing.Size(850, 1000)
$form.StartPosition = "CenterScreen"

# -----------------------
# Theme support
# -----------------------
function Apply-Theme {
    param([string]$Theme = "Light")

    switch ($Theme) {
        "Dark" {
            $form.BackColor = [System.Drawing.Color]::FromArgb(30,30,30)
            $fore = [System.Drawing.Color]::White
            $panelBack = [System.Drawing.Color]::FromArgb(45,45,45)
            $buttonBack = [System.Drawing.Color]::FromArgb(60,60,60)
            $dgvBack = [System.Drawing.Color]::FromArgb(25,25,25)
            break
        }
        "Color" {
            $form.BackColor = [System.Drawing.Color]::FromArgb(15,52,96)
            $fore = [System.Drawing.Color]::White
            $panelBack = [System.Drawing.Color]::FromArgb(30,80,140)
            $buttonBack = [System.Drawing.Color]::FromArgb(0,120,215)
            $dgvBack = [System.Drawing.Color]::FromArgb(15,52,96)
            break
        }
        default {
            # Light
            $form.BackColor = [System.Drawing.SystemColors]::Control
            $fore = [System.Drawing.Color]::Black
            $panelBack = [System.Drawing.SystemColors]::Info
            $buttonBack = [System.Drawing.Color]::Gainsboro
            $dgvBack = [System.Drawing.Color]::White
            break
        }
    }

    function Set-ColorsRec {
        param($ctrl)
        try { $ctrl.BackColor = $panelBack } catch {}
        try { $ctrl.ForeColor = $fore } catch {}

        foreach ($c in $ctrl.Controls) { Set-ColorsRec -ctrl $c }

        if ($ctrl -is [System.Windows.Forms.DataGridView]) {
            try {
                $ctrl.BackgroundColor = $dgvBack
                $ctrl.ColumnHeadersDefaultCellStyle.BackColor = $buttonBack
                $ctrl.ColumnHeadersDefaultCellStyle.ForeColor = $fore
                $ctrl.DefaultCellStyle.BackColor = $dgvBack
                $ctrl.DefaultCellStyle.ForeColor = $fore
                $ctrl.RowHeadersDefaultCellStyle.BackColor = $panelBack
                $ctrl.EnableHeadersVisualStyles = $false
            } catch {}
        }
        if ($ctrl -is [System.Windows.Forms.Button]) {
            try { $ctrl.BackColor = $buttonBack; $ctrl.FlatStyle = 'Standard' } catch {}
        }
        if ($ctrl -is [System.Windows.Forms.Panel]) { try { $ctrl.BackColor = $panelBack } catch {} }
        if ($ctrl -is [System.Windows.Forms.StatusStrip]) {
            try { $ctrl.BackColor = $panelBack } catch {}
            try { foreach ($item in $ctrl.Items) { $item.ForeColor = $fore } } catch {}
        }
    }

    Set-ColorsRec -ctrl $form

    try { $statusLabel.ForeColor = $fore } catch {}
}

# Theme selector (ComboBox)
$cmbTheme = New-Object System.Windows.Forms.ComboBox
$cmbTheme.DropDownStyle = 'DropDownList'
$cmbTheme.Items.AddRange(@("Light","Dark","Color")) | Out-Null
$cmbTheme.SelectedIndex = 0
$cmbTheme.Location = New-Object System.Drawing.Point(700,8)
$cmbTheme.Size = New-Object System.Drawing.Size(120,24)
$cmbTheme.Add_SelectedIndexChanged({ Apply-Theme -Theme $cmbTheme.SelectedItem.ToString() })
$form.Controls.Add($cmbTheme)

# End Theme support

$form.Text = "AD Security Group Manager"
$form.Size = New-Object System.Drawing.Size(850, 1000)
$form.StartPosition = "CenterScreen"

# --- 建立狀態列 ---
$statusStrip = New-Object System.Windows.Forms.StatusStrip
$statusLabel = New-Object System.Windows.Forms.ToolStripStatusLabel
$statusLabel.Text = "Ready"
$statusStrip.Items.Add($statusLabel) | Out-Null
$form.Controls.Add($statusStrip)

# 主要分頁容器
$tabControl = New-Object System.Windows.Forms.TabControl
$tabControl.Size = New-Object System.Drawing.Size(830, 880)
$tabControl.Location = New-Object System.Drawing.Point(5, 5)
$form.Controls.Add($tabControl)

# --- [分頁 1] ACL Viewer：檢視與匯出資料夾權限 ---
$tabACL = New-Object System.Windows.Forms.TabPage
$tabACL.Text = "ACL Viewer"
$tabControl.TabPages.Add($tabACL)
$lblPath = New-Object System.Windows.Forms.Label
$lblPath.Location = "20, 10"; $lblPath.Size = "100, 20"; $lblPath.Text = "Folder Path:"
$tabACL.Controls.Add($lblPath)
$txtPath = New-Object System.Windows.Forms.TextBox
$txtPath.Location = "20, 35"; $txtPath.Size = "790, 25"
$txtPath.Add_KeyDown({ param($s, $e) if ($e.Control -and $e.KeyCode -eq [System.Windows.Forms.Keys]::A) { $s.SelectAll(); $e.SuppressKeyPress = $true; $e.Handled = $true } })
$tabACL.Controls.Add($txtPath)
$btnCheckAcl = New-Object System.Windows.Forms.Button
$btnCheckAcl.Location = "20, 70"; $btnCheckAcl.Size = "150, 35"; $btnCheckAcl.Text = "Check Permissions"; $btnCheckAcl.BackColor = "LightGreen"
$tabACL.Controls.Add($btnCheckAcl)
$btnExportAclCsv = New-Object System.Windows.Forms.Button
$btnExportAclCsv.Location = "180, 70"
$btnExportAclCsv.Size = "150, 35"
$btnExportAclCsv.Text = "Export CSV"
$btnExportAclCsv.Enabled = $false
$tabACL.Controls.Add($btnExportAclCsv)
$dgvAcl = New-Object System.Windows.Forms.DataGridView
$dgvAcl.Location = "20, 120"; $dgvAcl.Size = "785, 360"
$dgvAcl.ReadOnly = $true; $dgvAcl.AutoSizeColumnsMode = "AllCells"; $dgvAcl.BackgroundColor = "White"; $dgvAcl.RowHeadersVisible = $false
$tabACL.Controls.Add($dgvAcl)

# 啟用 DataGridView 複製功能與自訂右鍵選單
$dgvAcl.MultiSelect = $true
$dgvAcl.SelectionMode = 'FullRowSelect'
$dgvAcl.AllowUserToAddRows = $false
$dgvAcl.ClipboardCopyMode = [System.Windows.Forms.DataGridViewClipboardCopyMode]::EnableAlwaysIncludeHeaderText
$aclContext = New-Object System.Windows.Forms.ContextMenuStrip
$miCopyGroup = $aclContext.Items.Add("Copy Group Name (no domain)")
$miCopyGroup.Add_Click({
        $selected = $dgvAcl.SelectedRows
        if ($selected.Count -gt 0) {
            $identity = $selected[0].Cells["IdentityReference"].Value
            $groupOnly = ($identity -split '\\')[-1]
            [System.Windows.Forms.Clipboard]::SetText($groupOnly)
            Update-Status "Copied group name: $groupOnly"
        }
    })
$dgvAcl.ContextMenuStrip = $aclContext
$dgvAcl.Add_KeyDown({ param($s, $e) if ($e.Control -and $e.KeyCode -eq [System.Windows.Forms.Keys]::C) { $clip = $dgvAcl.GetClipboardContent(); if ($clip) { [System.Windows.Forms.Clipboard]::SetDataObject($clip) } } })

$lblAclOwner = New-Object System.Windows.Forms.Label
$lblAclOwner.Location = "20, 500"; $lblAclOwner.Size = "800, 25"; $lblAclOwner.ForeColor = "DarkBlue"
$tabACL.Controls.Add($lblAclOwner)

# [事件] 檢查資料夾權限按鈕點擊
$btnCheckAcl.Add_Click({
        $path = $txtPath.Text.Trim()
        if ([string]::IsNullOrWhiteSpace($path)) {
            [System.Windows.Forms.MessageBox]::Show("Please enter a folder path.", "Validation", "OK", "Warning")
            return
        }
        if (Test-Path $path) {
            $form.Cursor = [System.Windows.Forms.Cursors]::WaitCursor
            Update-Status "Reading ACL for '$path'..."
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
                $script:currentAclData = $dt
                $script:currentAclPath = $path
                $script:currentAclOwner = $acl.Owner
                $btnExportAclCsv.Enabled = ($dt.Rows.Count -gt 0)
                Update-Status "ACL loaded: $($dt.Rows.Count) entries. Owner: $($acl.Owner)"
            }
            catch {
                $btnExportAclCsv.Enabled = $false
                $script:currentAclData = $null
                $script:currentAclPath = $null
                [System.Windows.Forms.MessageBox]::Show("Access Denied or Error: $($_.Exception.Message)")
                Update-Status "ACL check failed."
            }
            finally {
                $form.Cursor = [System.Windows.Forms.Cursors]::Default
            }
        }
        else {
            [System.Windows.Forms.MessageBox]::Show("Invalid Path: '$path'")
        }
    })

# [事件] 匯出 ACL 結果至 CSV 按鈕點擊
$btnExportAclCsv.Add_Click({
        if (-not $script:currentAclData -or $script:currentAclData.Rows.Count -eq 0) {
            [System.Windows.Forms.MessageBox]::Show("No ACL data available to export.", "Export CSV", "OK", "Warning")
            return
        }
        $saveFile = New-Object System.Windows.Forms.SaveFileDialog
        $saveFile.Filter = "CSV Files (*.csv)|*.csv"
        $saveFile.FileName = "ACL_Viewer_$(Get-Date -Format 'yyyyMMdd_HHmmss').csv"
        if ($saveFile.ShowDialog() -eq [System.Windows.Forms.DialogResult]::OK) {
            $form.Cursor = [System.Windows.Forms.Cursors]::WaitCursor
            Update-Status "Exporting ACL data to CSV..."
            try {
                $script:currentAclData | Export-Csv -Path $saveFile.FileName -NoTypeInformation -Encoding UTF8 -Force
                [System.Windows.Forms.MessageBox]::Show(
                    "ACL data exported successfully to:`n`n$($saveFile.FileName)", "Export Complete", "OK", "Information"
                )
                Update-Status "ACL CSV export completed: $($saveFile.FileName)"
            }
            catch {
                [System.Windows.Forms.MessageBox]::Show("CSV export failed.`n`n$($_.Exception.Message)", "Export Error", "OK", "Error")
                Update-Status "ACL CSV export failed."
            }
            finally {
                $form.Cursor = [System.Windows.Forms.Cursors]::Default
            }
        }
    })

# --- [分頁 2] Audit Members：查閱群組成員 (支援遞迴查詢) ---
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
$txtGroupSearch.Add_KeyDown({ param($s, $e) if ($e.Control -and $e.KeyCode -eq [System.Windows.Forms.Keys]::A) { $s.SelectAll(); $e.SuppressKeyPress = $true; $e.Handled = $true } })
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

$lblSubGroups = New-Object System.Windows.Forms.Label
$lblSubGroups.Location = "20, 225"; $lblSubGroups.Size = "400, 18"; $lblSubGroups.Text = "Sub-Groups (nested groups inside this group):"
$lblSubGroups.Font = New-Object System.Drawing.Font("Segoe UI", 9, [System.Drawing.FontStyle]::Bold)
$tabView.Controls.Add($lblSubGroups)
$dgvSubGroups = New-Object System.Windows.Forms.DataGridView
$dgvSubGroups.Location = "20, 245"; $dgvSubGroups.Size = "785, 130"
$dgvSubGroups.ReadOnly = $true; $dgvSubGroups.AutoSizeColumnsMode = "Fill"; $dgvSubGroups.BackgroundColor = "White"
$dgvSubGroups.SelectionMode = "FullRowSelect"; $dgvSubGroups.AllowUserToAddRows = $false; $dgvSubGroups.RowHeadersVisible = $false
$tabView.Controls.Add($dgvSubGroups)

$lblMembersHeader = New-Object System.Windows.Forms.Label
$lblMembersHeader.Location = "20, 380"; $lblMembersHeader.Size = "300, 18"; $lblMembersHeader.Text = "All Members (recursive):"
$lblMembersHeader.Font = New-Object System.Drawing.Font("Segoe UI", 9, [System.Drawing.FontStyle]::Bold)
$tabView.Controls.Add($lblMembersHeader)
$dgvMembers = New-Object System.Windows.Forms.DataGridView
$dgvMembers.Location = "20, 400"; $dgvMembers.Size = "785, 445"
$dgvMembers.ReadOnly = $true; $dgvMembers.AutoSizeColumnsMode = "Fill"
$dgvMembers.SelectionMode = "FullRowSelect"; $dgvMembers.AllowUserToAddRows = $false
$tabView.Controls.Add($dgvMembers)

# [事件] 獲取群組成員按鈕點擊 (先查找子群組，再遞迴尋找最終使用者)
$btnFetch.Add_Click({
        if ([string]::IsNullOrWhiteSpace($txtGroupSearch.Text)) {
            [System.Windows.Forms.MessageBox]::Show("Please enter a group name or email address.", "Validation", "OK", "Warning")
            return
        }
        $form.Cursor = [System.Windows.Forms.Cursors]::WaitCursor
        Update-Status "Fetching group members..."
        try {
            $group = Resolve-ADGroupByNameOrMail -SearchValue $txtGroupSearch.Text
            $directMembers = Get-ADGroupMember -Identity $group -ErrorAction Stop
            $subGroupObjects = $directMembers | Where-Object { $_.objectClass -eq 'group' }
        
            $dtSubGroups = New-Object System.Data.DataTable
            $dtSubGroups.Columns.Add("Sub-Group Name") | Out-Null
            $dtSubGroups.Columns.Add("Members (recursive)") | Out-Null
            foreach ($sg in $subGroupObjects) {
                $sgMemberCount = (Get-ADGroupMember -Identity $sg -Recursive -ErrorAction SilentlyContinue | Where-Object { $_.objectClass -eq 'user' }).Count
                $sgRow = $dtSubGroups.NewRow()
                $sgRow."Sub-Group Name" = $sg.Name
                $sgRow."Members (recursive)" = $sgMemberCount
                $dtSubGroups.Rows.Add($sgRow)
            }
            $dgvSubGroups.DataSource = $dtSubGroups
        
            $members = Get-ADGroupMember -Identity $group -Recursive |
            Where-Object { $_.objectClass -eq 'user' } |
            Get-ADUser -Properties mail, DisplayName, Department, Title |
            Select-Object DisplayName, SamAccountName, mail, Department, Title
        
            $dt = New-Object System.Data.DataTable
            $dt.Columns.Add("DisplayName") | Out-Null
            $dt.Columns.Add("SamAccountName") | Out-Null
            $dt.Columns.Add("mail") | Out-Null
            $dt.Columns.Add("Department") | Out-Null
            $dt.Columns.Add("Title") | Out-Null
            foreach ($m in $members) {
                $row = $dt.NewRow()
                $row.DisplayName = $m.DisplayName
                $row.SamAccountName = $m.SamAccountName
                $row.mail = $m.mail
                $row.Department = $m.Department
                $row.Title = $m.Title
                $dt.Rows.Add($row)
            }
            $dgvMembers.DataSource = $dt
            $lblCount.Text = "Sub-Groups: $($dtSubGroups.Rows.Count) | Total Members: $($dt.Rows.Count)"
            $btnExport.Enabled = ($dt.Rows.Count -gt 0)
            $script:currentData = $members
            Update-Status "Loaded $($dtSubGroups.Rows.Count) sub-group(s) and $($dt.Rows.Count) member(s) from '$($group.Name)'."
        }
        catch {
            [System.Windows.Forms.MessageBox]::Show("Group not found or AD access error.`n`n$($_.Exception.Message)", "Error")
            Update-Status "Fetch failed."
        }
        finally {
            $form.Cursor = [System.Windows.Forms.Cursors]::Default
        }
    })

# [事件] 匯出群組成員為 CSV
$btnExport.Add_Click({
        $saveFile = New-Object System.Windows.Forms.SaveFileDialog
        $saveFile.Filter = "CSV Files (*.csv)|*.csv"
        $saveFile.FileName = "GroupMembers_$(Get-Date -Format 'yyyyMMdd_HHmmss').csv"
        if ($saveFile.ShowDialog() -eq [System.Windows.Forms.DialogResult]::OK) {
            try {
                if ($script:currentData -and $script:currentData.Count -gt 0) {
                    $script:currentData | Export-Csv -Path $saveFile.FileName -NoTypeInformation -Encoding UTF8 -Force
                    [System.Windows.Forms.MessageBox]::Show("Data exported successfully to:`n`n$($saveFile.FileName)", "Export Complete", "OK", "Information")
                    Update-Status "Exported to $($saveFile.FileName)"
                }
                else {
                    [System.Windows.Forms.MessageBox]::Show("No data to export.", "Warning")
                }
            }
            catch {
                [System.Windows.Forms.MessageBox]::Show("Export failed: $($_.Exception.Message)", "Error")
            }
        }
    })

# 快捷傳送名單的右鍵選單邏輯
$contextMenu = New-Object System.Windows.Forms.ContextMenuStrip
$menuItemAdd = $contextMenu.Items.Add("Add selected to Modify list")
$menuItemAdd.Add_Click({
        if ($dgvMembers.SelectedRows.Count -gt 0) {
            foreach ($row in $dgvMembers.SelectedRows) {
                $sam = $row.Cells["SamAccountName"].Value.ToString()
                if ($sam) { $null = $dgvInput.Rows.Add($sam) }
            }
            $tabControl.SelectedTab = $tabUpdate
            Update-Status "Transferred $($dgvMembers.SelectedRows.Count) user(s) to Bulk Updates tab."
        }
    })
$dgvMembers.ContextMenuStrip = $contextMenu

# --- [分頁 3] Bulk Updates：批次新增或移除群組成員 ---
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
$txtTargetGroup.Add_KeyDown({ param($s, $e) if ($e.Control -and $e.KeyCode -eq [System.Windows.Forms.Keys]::A) { $s.SelectAll(); $e.SuppressKeyPress = $true; $e.Handled = $true } })
$tabUpdate.Controls.Add($txtTargetGroup)
$btnImportFile = New-Object System.Windows.Forms.Button
$btnImportFile.Text = "Upload CSV/TXT"; $btnImportFile.Location = "430, 193"; $btnImportFile.Width = 130
$tabUpdate.Controls.Add($btnImportFile)

# 檔案匯入處理
$btnImportFile.Add_Click({
        $openFile = New-Object System.Windows.Forms.OpenFileDialog
        $openFile.Filter = "CSV/TXT Files (*.csv;*.txt)|*.csv;*.txt"
        if ($openFile.ShowDialog() -eq [System.Windows.Forms.DialogResult]::OK) {
            $lines = Get-Content $openFile.FileName | Where-Object { $_.Trim() -ne "" }
            foreach ($line in $lines) { $dgvInput.Rows.Add($line.Trim()) | Out-Null }
            Update-Status "Imported $($lines.Count) identifier(s) from file."
        }
    })

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

# [核心邏輯] 處理批次新增或移除群組成員
$ProcessGrid = {
    param($Action)
    $groupName = $txtTargetGroup.Text.Trim()
    $userList = New-Object System.Collections.Generic.List[string]
    foreach ($row in $dgvInput.Rows) { if ($row.Cells["colIdentifier"].Value) { $userList.Add($row.Cells["colIdentifier"].Value.ToString().Trim()) } }
    $users = $userList | Select-Object -Unique
    if ([string]::IsNullOrWhiteSpace($groupName) -or $users.Count -eq 0) { [System.Windows.Forms.MessageBox]::Show("Please provide both Target Group and User identifiers.", "Validation Error", "OK", "Warning") ; return }
    
    $confirmMsg = "You are about to $($Action.ToUpper()) $($users.Count) user(s) to/from group:`n`n$groupName`n`nDo you want to proceed?"
    if ([System.Windows.Forms.MessageBox]::Show($confirmMsg, "Confirm Bulk $Action", "YesNo", "Question") -eq "No") { Update-Status "Bulk $Action cancelled."; return }
    
    $saveLog = New-Object System.Windows.Forms.SaveFileDialog
    $saveLog.Filter = "Text Files (*.txt)|*.txt"; $saveLog.FileName = "BulkUpdate_Log_$(Get-Date -Format 'yyyyMMdd').txt"
    if ($saveLog.ShowDialog() -eq [System.Windows.Forms.DialogResult]::OK) {
        $logPath = $saveLog.FileName
        $form.Cursor = [System.Windows.Forms.Cursors]::WaitCursor
        $successCount = 0; $errorCount = 0
        try {
            $group = Resolve-ADGroupByNameOrMail -SearchValue $groupName
            foreach ($u in $users) {
                try {
                    $adUser = Resolve-ADUserByMailOrSam -SearchValue $u
                    if ($Action -eq "Add") { Add-ADGroupMember $group $adUser -ErrorAction Stop; $s = "[SUCCESS] Added: $u" }
                    else { Remove-ADGroupMember $group $adUser -Confirm:$false -ErrorAction Stop; $s = "[SUCCESS] Removed: $u" }
                    $successCount++
                }
                catch { $s = "[ERROR] $u : $($_.Exception.Message)"; $errorCount++ }
                $txtUpdateLog.AppendText("`r`n$s"); $s | Out-File $logPath -Append
                $txtUpdateLog.ScrollToCaret()
                [System.Windows.Forms.Application]::DoEvents()
            }
            $summary = "Processing Finished.`n`nSuccess: $successCount | Errors: $errorCount`nLog saved to: $logPath"
            [System.Windows.Forms.MessageBox]::Show($summary, "Complete", "OK", "Information")
            Update-Status "Bulk $Action complete. Success: $successCount, Errors: $errorCount"
        }
        catch {
            [System.Windows.Forms.MessageBox]::Show("Target Group resolution failed.`n`n$($_.Exception.Message)", "Error")
            Update-Status "Target group resolution failed."
        }
        finally {
            $form.Cursor = [System.Windows.Forms.Cursors]::Default
        }
    }
}
$btnAdd.Add_Click({ &$ProcessGrid "Add" })
$btnRemove.Add_Click({ &$ProcessGrid "Remove" })

# --- [分頁 4] Access Checker：檢查單一使用者的目錄權限 ---
$tabAccessChk = New-Object System.Windows.Forms.TabPage
$tabAccessChk.Text = "Access Checker"
$tabControl.TabPages.Add($tabAccessChk)

$panelHint4 = New-Object System.Windows.Forms.Panel
$panelHint4.Location = "20, 10"; $panelHint4.Size = "790, 110"; $panelHint4.BackColor = "Info"
$tabAccessChk.Controls.Add($panelHint4)
$lblHint4 = New-Object System.Windows.Forms.Label
$lblHint4.Text = "FUNCTION: Check Individual Access Rights.`n`nGUIDELINES:`n1. Enter or Browse to a folder path.`n2. Enter the user's Email or SAM account name.`n3. Click 'Check Access' to see their effective permissions (direct + via group memberships)."
$lblHint4.Dock = "Fill"; $lblHint4.Padding = "10, 10, 10, 10"; $lblHint4.Font = New-Object System.Drawing.Font("Segoe UI", 9, [System.Drawing.FontStyle]::Bold)
$panelHint4.Controls.Add($lblHint4)

$lblAccessPath = New-Object System.Windows.Forms.Label
$lblAccessPath.Location = "20, 130"; $lblAccessPath.Size = "100, 20"; $lblAccessPath.Text = "Folder Path:"
$tabAccessChk.Controls.Add($lblAccessPath)
$txtAccessPath = New-Object System.Windows.Forms.TextBox
$txtAccessPath.Location = "20, 155"; $txtAccessPath.Size = "790, 25"
$txtAccessPath.Add_KeyDown({ param($s, $e) if ($e.Control -and $e.KeyCode -eq [System.Windows.Forms.Keys]::A) { $s.SelectAll(); $e.SuppressKeyPress = $true; $e.Handled = $true } })
$tabAccessChk.Controls.Add($txtAccessPath)

$lblAccessUser = New-Object System.Windows.Forms.Label
$lblAccessUser.Location = "20, 195"; $lblAccessUser.Size = "150, 20"; $lblAccessUser.Text = "User (Email or SAM):"
$tabAccessChk.Controls.Add($lblAccessUser)
$txtAccessUser = New-Object System.Windows.Forms.TextBox
$txtAccessUser.Location = "20, 220"; $txtAccessUser.Size = "400, 25"
$txtAccessUser.Add_KeyDown({ param($s, $e) if ($e.Control -and $e.KeyCode -eq [System.Windows.Forms.Keys]::A) { $s.SelectAll(); $e.SuppressKeyPress = $true; $e.Handled = $true } })
$tabAccessChk.Controls.Add($txtAccessUser)

$btnCheckAccess = New-Object System.Windows.Forms.Button
$btnCheckAccess.Location = "430, 218"; $btnCheckAccess.Size = "150, 30"; $btnCheckAccess.Text = "Check Access"; $btnCheckAccess.BackColor = "LightGreen"
$tabAccessChk.Controls.Add($btnCheckAccess)
$btnClearTab4 = New-Object System.Windows.Forms.Button
$btnClearTab4.Location = "590, 218"; $btnClearTab4.Size = "100, 30"; $btnClearTab4.Text = "Clear"
$tabAccessChk.Controls.Add($btnClearTab4)

$lblAccessInfo = New-Object System.Windows.Forms.Label
$lblAccessInfo.Location = "20, 260"; $lblAccessInfo.Size = "790, 20"; $lblAccessInfo.ForeColor = "DarkBlue"; $lblAccessInfo.AutoSize = $true
$tabAccessChk.Controls.Add($lblAccessInfo)

$lblMatchedGroups = New-Object System.Windows.Forms.Label
$lblMatchedGroups.Location = "20, 285"; $lblMatchedGroups.Size = "400, 18"; $lblMatchedGroups.Text = "Matched Groups (ACL Group → User Membership Path):"
$lblMatchedGroups.Font = New-Object System.Drawing.Font("Segoe UI", 9, [System.Drawing.FontStyle]::Bold)
$tabAccessChk.Controls.Add($lblMatchedGroups)
$dgvMatchedGroups = New-Object System.Windows.Forms.DataGridView
$dgvMatchedGroups.Location = "20, 305"; $dgvMatchedGroups.Size = "785, 130"; $dgvMatchedGroups.BackColor = "White"
$dgvMatchedGroups.ReadOnly = $true; $dgvMatchedGroups.AutoSizeColumnsMode = "Fill"
$dgvMatchedGroups.SelectionMode = "FullRowSelect"; $dgvMatchedGroups.AllowUserToAddRows = $false; $dgvMatchedGroups.RowHeadersVisible = $false
$tabAccessChk.Controls.Add($dgvMatchedGroups)

$lblPermHeader = New-Object System.Windows.Forms.Label
$lblPermHeader.Location = "20, 440"; $lblPermHeader.Size = "300, 18"; $lblPermHeader.Text = "Effective Permissions:"
$lblPermHeader.Font = New-Object System.Drawing.Font("Segoe UI", 9, [System.Drawing.FontStyle]::Bold)
$tabAccessChk.Controls.Add($lblPermHeader)
$dgvAccess = New-Object System.Windows.Forms.DataGridView
$dgvAccess.Location = "20, 460"; $dgvAccess.Size = "785, 375"
$dgvAccess.ReadOnly = $true; $dgvAccess.AutoSizeColumnsMode = "Fill"; $dgvAccess.BackgroundColor = "White"
$dgvAccess.SelectionMode = "FullRowSelect"; $dgvAccess.AllowUserToAddRows = $false; $dgvAccess.RowHeadersVisible = $false
$tabAccessChk.Controls.Add($dgvAccess)

# [事件] 開始分析使用者對指定目錄的有效權限
$btnCheckAccess.Add_Click({
        $path = $txtAccessPath.Text.Trim()
        $userSearch = $txtAccessUser.Text.Trim()
        if ([string]::IsNullOrWhiteSpace($path) -or [string]::IsNullOrWhiteSpace($userSearch)) {
            [System.Windows.Forms.MessageBox]::Show("Please enter both a folder path and a user identifier.", "Validation", "OK", "Warning")
            return
        }
        if (-not (Test-Path $path)) {
            [System.Windows.Forms.MessageBox]::Show("Invalid Path: '$path'", "Error")
            return
        }
        $form.Cursor = [System.Windows.Forms.Cursors]::WaitCursor
        Update-Status "Checking access for '$userSearch' on '$path'..."
        try {
            $adUser = Resolve-ADUserByMailOrSam -SearchValue $userSearch
            $userSam = $adUser.SamAccountName
            $userDisplay = (Get-ADUser $adUser -Properties DisplayName).DisplayName
        
            $userGroups = @()
            try {
                # 利用 ADSI 的 tokenGroups 屬性精準獲取使用者的所有隸屬群組 (包含遞迴與巢狀群組 SID)
                $userDN = (Get-ADUser $adUser).DistinguishedName
                $userEntry = [ADSI]"LDAP://$userDN"
                $userEntry.RefreshCache("tokenGroups")
                $userGroups = $userEntry.Properties["tokenGroups"] | ForEach-Object {
                    $sid = New-Object System.Security.Principal.SecurityIdentifier($_, 0)
                    try { (Get-ADGroup $sid -ErrorAction SilentlyContinue).Name } catch {}
                } | Where-Object { $_ }
            }
            catch {
                # 發生例外時的備案處理 (回退使用未遞迴的基本 API)
                try {
                    $userGroups = Get-ADPrincipalGroupMembership $adUser -ErrorAction SilentlyContinue | Select-Object -ExpandProperty Name
                }
                catch {
                    $userObj = Get-ADUser $adUser -Properties MemberOf
                    $userGroups = $userObj.MemberOf | ForEach-Object { (Get-ADGroup $_).Name }
                }
            }
        
            $acl = Get-Acl -Path $path
            $dt = New-Object System.Data.DataTable
            $dt.Columns.Add("IdentityReference") | Out-Null
            $dt.Columns.Add("FileSystemRights") | Out-Null
            $dt.Columns.Add("AccessControlType") | Out-Null
            $dt.Columns.Add("IsInherited") | Out-Null
            $dt.Columns.Add("AccessType") | Out-Null
            foreach ($ace in $acl.Access) {
                $identity = $ace.IdentityReference.Value
                $identityName = ($identity -split '\\')[-1]
                $matchType = $null
                if ($identityName -eq $userSam) {
                    $matchType = "Direct"
                }
                elseif ($identityName -in $userGroups) {
                    $matchType = "Via Group"
                }
                if ($matchType) {
                    $row = $dt.NewRow()
                    $row.IdentityReference = $identity
                    $row.FileSystemRights = $ace.FileSystemRights.ToString()
                    $row.AccessControlType = $ace.AccessControlType.ToString()
                    $row.IsInherited = $ace.IsInherited.ToString()
                    $row.AccessType = $matchType
                    $dt.Rows.Add($row)
                }
            }
        
            $dtGroups = New-Object System.Data.DataTable
            $dtGroups.Columns.Add("ACL Group") | Out-Null
            $dtGroups.Columns.Add("Membership") | Out-Null
            $matchedGroupNames = @()
            foreach ($ace in $acl.Access) {
                $gName = ($ace.IdentityReference.Value -split '\\')[-1]
                if ($gName -in $userGroups -and $gName -notin $matchedGroupNames) {
                    $matchedGroupNames += $gName
                    $membershipPath = "Direct Member"
                    try {
                        $directMembers = Get-ADGroupMember -Identity $gName -ErrorAction SilentlyContinue
                        $directSams = $directMembers | Where-Object { $_.objectClass -eq 'user' } | Select-Object -ExpandProperty SamAccountName
                        if ($userSam -notin $directSams) {
                            $subGroups = $directMembers | Where-Object { $_.objectClass -eq 'group' }
                            foreach ($sg in $subGroups) {
                                $sgMembers = Get-ADGroupMember -Identity $sg -Recursive -ErrorAction SilentlyContinue | Select-Object -ExpandProperty SamAccountName
                                if ($userSam -in $sgMembers) {
                                    $membershipPath = "Via: $($sg.Name)"
                                    break
                                }
                            }
                        }
                    }
                    catch { $membershipPath = "(Unable to resolve path)" }
                    $gRow = $dtGroups.NewRow()
                    $gRow."ACL Group" = $gName
                    $gRow."Membership" = $membershipPath
                    $dtGroups.Rows.Add($gRow)
                }
            }
            $dgvMatchedGroups.DataSource = $dtGroups
            $dgvAccess.DataSource = $dt
            $lblAccessInfo.Text = "User: $userDisplay ($userSam) | Groups: $($userGroups.Count) | Matched Groups on ACL: $($matchedGroupNames.Count) | ACL Entries: $($dt.Rows.Count)"
            if ($dt.Rows.Count -eq 0) {
                Update-Status "No matching ACL entries found for '$userDisplay'. User may have no explicit access."
            }
            else {
                Update-Status "Found $($dt.Rows.Count) ACL entries for '$userDisplay' on '$path'."
            }
        }
        catch {
            [System.Windows.Forms.MessageBox]::Show("Error: $($_.Exception.Message)", "Access Check Failed")
            Update-Status "Access check failed."
        }
        finally {
            $form.Cursor = [System.Windows.Forms.Cursors]::Default
        }
    })

# 清除所有分頁狀態
$btnClearTab1.Add_Click({ $txtGroupSearch.Text = ""; $dgvMembers.DataSource = $null; $dgvSubGroups.DataSource = $null; $lblCount.Text = "Sub-Groups: 0 | Total Members: 0"; $btnExport.Enabled = $false; Update-Status "Audit tab cleared." })
$btnClearTab2.Add_Click({ $txtTargetGroup.Text = ""; $dgvInput.Rows.Clear(); $txtUpdateLog.Clear(); Update-Status "Bulk Updates tab cleared." })
$btnClearTab4.Add_Click({ $txtAccessPath.Text = ""; $txtAccessUser.Text = ""; $dgvAccess.DataSource = $null; $dgvMatchedGroups.DataSource = $null; $lblAccessInfo.Text = ""; Update-Status "Access Checker tab cleared." })

# 全域結束按鈕
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

# 顯示主視窗
try { Apply-Theme -Theme $cmbTheme.SelectedItem.ToString() } catch { Apply-Theme -Theme 'Light' }
$form.ShowDialog()