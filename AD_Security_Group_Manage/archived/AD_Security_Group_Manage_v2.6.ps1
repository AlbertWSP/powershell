
# 2026-03-12
# @Albert Ng
# AD Security Group Manager - Version 2.6
# =================================================================================

# 1. Self-Elevation (Request Administrator Privileges)

# =================================================================================

if (-not ([Security.Principal.WindowsPrincipal][Security.Principal.WindowsIdentity]::GetCurrent()).IsInRole([Security.Principal.WindowsBuiltInRole]::Administrator)) {

    $newProcess = New-Object System.Diagnostics.ProcessStartInfo "PowerShell"

    $newProcess.Arguments = "-NoProfile -ExecutionPolicy Bypass -File `"$PSCommandPath`""

    $newProcess.Verb = "runas"

    try {

        [System.Diagnostics.Process]::Start($newProcess) | Out-Null

    }
    catch {

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


# --- Helper Functions (v2.0) ---

function Resolve-ADGroupByNameOrMail {

    param([string]$SearchValue)

    $val = $SearchValue.Trim()

    Get-ADGroup -Filter { Name -eq $val -or mail -eq $val } -ErrorAction Stop

}

  

function Resolve-ADUserByMailOrSam {

    param([string]$SearchValue)

    $val = $SearchValue.Trim()

    Get-ADUser -Filter { mail -eq $val -or SamAccountName -eq $val } -ErrorAction Stop

}

  

function Update-Status {

    param([string]$Message)

    $statusLabel.Text = $Message

    [System.Windows.Forms.Application]::DoEvents()

}

# BEGIN CHANGE - Replace Export-DataTableToExcel with safer COM version

function Export-DataTableToExcel {
    param(
        [Parameter(Mandatory = $true)]
        [System.Data.DataTable]$DataTable,

        [Parameter(Mandatory = $true)]
        [string]$FilePath,

        [string]$WorksheetName = "ACL Viewer",

        [string]$FolderPath = "",

        [string]$Owner = ""
    )

    $excel = $null
    $workbook = $null
    $worksheet = $null
    $saveSucceeded = $false

    try {
        $excel = New-Object -ComObject Excel.Application
        $excel.Visible = $false
        $excel.DisplayAlerts = $false

        $workbook = $excel.Workbooks.Add()
        $worksheet = $workbook.Worksheets.Item(1)

        # Safe worksheet naming
        if ($WorksheetName.Length -gt 31) {
            $WorksheetName = $WorksheetName.Substring(0, 31)
        }
        $worksheet.Name = $WorksheetName

        # Header section
        $worksheet.Cells.Item(1, 1) = "ACL Viewer Export"
        $worksheet.Cells.Item(2, 1) = "Folder Path"
        $worksheet.Cells.Item(2, 2) = $FolderPath
        $worksheet.Cells.Item(3, 1) = "Owner"
        $worksheet.Cells.Item(3, 2) = $Owner
        $worksheet.Cells.Item(4, 1) = "Export Time"
        $worksheet.Cells.Item(4, 2) = (Get-Date).ToString("yyyy-MM-dd HH:mm:ss")

        $worksheet.Range("A1:B1").Merge() | Out-Null
        $worksheet.Range("A1").Font.Bold = $true
        $worksheet.Range("A1").Font.Size = 14

        # Table header
        $startRow = 6
        for ($c = 0; $c -lt $DataTable.Columns.Count; $c++) {
            $worksheet.Cells.Item($startRow, $c + 1) = $DataTable.Columns[$c].ColumnName
        }

        $headerRange = $worksheet.Range(
            $worksheet.Cells.Item($startRow, 1),
            $worksheet.Cells.Item($startRow, $DataTable.Columns.Count)
        )
        $headerRange.Font.Bold = $true
        $headerRange.Interior.ColorIndex = 15

        # Table data
        for ($r = 0; $r -lt $DataTable.Rows.Count; $r++) {
            for ($c = 0; $c -lt $DataTable.Columns.Count; $c++) {
                $value = $DataTable.Rows[$r][$c]
                if ($null -eq $value) { $value = "" }
                $worksheet.Cells.Item($startRow + 1 + $r, $c + 1) = $value.ToString()
            }
        }

        # Borders
        $lastRow = $startRow + $DataTable.Rows.Count
        $lastCol = $DataTable.Columns.Count
        $usedRange = $worksheet.Range(
            $worksheet.Cells.Item($startRow, 1),
            $worksheet.Cells.Item($lastRow, $lastCol)
        )
        $usedRange.Borders.LineStyle = 1

        # AutoFit
        $worksheet.Columns.AutoFit() | Out-Null

        # Save
        # 51 = xlOpenXMLWorkbook (.xlsx)
        $workbook.SaveAs($FilePath, 51)
        $saveSucceeded = $true
    }
    catch {
        throw "Excel SaveAs failed: $($_.Exception.Message)"
    }
    finally {
        # IMPORTANT:
        # Cleanup errors should NOT cause the export to be marked as failed
        # if the file was already saved successfully.

        if ($worksheet) {
            try { [void][System.Runtime.InteropServices.Marshal]::FinalReleaseComObject($worksheet) } catch {}
            $worksheet = $null
        }

        if ($workbook) {
            try { $workbook.Close($false) } catch {}
            try { [void][System.Runtime.InteropServices.Marshal]::FinalReleaseComObject($workbook) } catch {}
            $workbook = $null
        }

        if ($excel) {
            try { $excel.DisplayAlerts = $false } catch {}
            try { $excel.Quit() } catch {}
            try { [void][System.Runtime.InteropServices.Marshal]::FinalReleaseComObject($excel) } catch {}
            $excel = $null
        }

        [GC]::Collect()
        [GC]::WaitForPendingFinalizers()
        [GC]::Collect()
        [GC]::WaitForPendingFinalizers()
    }

    # Final validation:
    if (-not $saveSucceeded -or -not (Test-Path $FilePath)) {
        throw "Excel export did not complete successfully. File was not created."
    }
}

# END CHANGE

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

# 3. UI Interface Code (Version 2.5- Improved Edition)

# =================================================================================

$form = New-Object System.Windows.Forms.Form

$form.Text = "AD Security Group Manager - Version 2.5"

$form.Size = New-Object System.Drawing.Size(850, 1000)

$form.StartPosition = "CenterScreen"

  

# --- Status Bar ---

$statusStrip = New-Object System.Windows.Forms.StatusStrip

$statusLabel = New-Object System.Windows.Forms.ToolStripStatusLabel

$statusLabel.Text = "Ready"

$statusStrip.Items.Add($statusLabel) | Out-Null

$form.Controls.Add($statusStrip)

  

# Main Container

$tabControl = New-Object System.Windows.Forms.TabControl

$tabControl.Size = New-Object System.Drawing.Size(830, 880)

$tabControl.Location = New-Object System.Drawing.Point(5, 5)

$form.Controls.Add($tabControl)

  

# --- TAB 1: ACL VIEWER (moved to first position in v2.2) ---

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

  
# BEGIN CHANGE - Add Export Excel button to ACL Viewer

$btnExportAclExcel = New-Object System.Windows.Forms.Button
$btnExportAclExcel.Location = "180, 70"
$btnExportAclExcel.Size = "150, 35"
$btnExportAclExcel.Text = "Export Excel"
$btnExportAclExcel.Enabled = $false
$tabACL.Controls.Add($btnExportAclExcel)

# END CHANGE


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

  

# Support Ctrl+C to copy selected cells

$dgvAcl.Add_KeyDown({ param($s, $e) if ($e.Control -and $e.KeyCode -eq [System.Windows.Forms.Keys]::C) { $clip = $dgvAcl.GetClipboardContent(); if ($clip) { [System.Windows.Forms.Clipboard]::SetDataObject($clip) } } })

  

$lblAclOwner = New-Object System.Windows.Forms.Label

$lblAclOwner.Location = "20, 500"; $lblAclOwner.Size = "800, 25"; $lblAclOwner.ForeColor = "DarkBlue"

$tabACL.Controls.Add($lblAclOwner)

  

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

                #$dgvAcl.DataSource = $dt

                #Update-Status "ACL loaded: $($dt.Rows.Count) entries. Owner: $($acl.Owner)"
            
                # BEGIN CHANGE - Enable Excel export after ACL load

                $dgvAcl.DataSource = $dt
                $script:currentAclData = $dt
                $script:currentAclPath = $path
                $script:currentAclOwner = $acl.Owner
                $btnExportAclExcel.Enabled = ($dt.Rows.Count -gt 0)

                Update-Status "ACL loaded: $($dt.Rows.Count) entries. Owner: $($acl.Owner)"

                # END CHANGE


            }
            catch {
            
                # BEGIN CHANGE - Reset ACL export state on failure

                $btnExportAclExcel.Enabled = $false
                $script:currentAclData = $null
                $script:currentAclPath = $null
                $sc


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

# BEGIN CHANGE - ACL Viewer Export Excel event

$btnExportAclExcel.Add_Click({

        if (-not $script:currentAclData -or $script:currentAclData.Rows.Count -eq 0) {
            [System.Windows.Forms.MessageBox]::Show("No ACL data available to export.", "Export Excel", "OK", "Warning")
            return
        }

        $saveFile = New-Object System.Windows.Forms.SaveFileDialog
        $saveFile.Filter = "Excel Workbook (*.xlsx)|*.xlsx"
        $saveFile.FileName = "ACL_Viewer_$(Get-Date -Format 'yyyyMMdd_HHmmss').xlsx"

        if ($saveFile.ShowDialog() -eq [System.Windows.Forms.DialogResult]::OK) {
            $form.Cursor = [System.Windows.Forms.Cursors]::WaitCursor
            Update-Status "Exporting ACL data to Excel..."

            try {
                Export-DataTableToExcel `
                    -DataTable $script:currentAclData `
                    -FilePath $saveFile.FileName `
                    -WorksheetName "ACL Viewer" `
                    -FolderPath $script:currentAclPath `
                    -Owner $script:currentAclOwner

                [System.Windows.Forms.MessageBox]::Show(
                    "ACL data exported successfully to:`n`n$($saveFile.FileName)",
                    "Export Complete",
                    "OK",
                    "Information"
                )

                Update-Status "ACL Excel export completed: $($saveFile.FileName)"
            }
            catch {
                [System.Windows.Forms.MessageBox]::Show(
                    "Excel export failed.`n`n$($_.Exception.Message)",
                    "Export Error",
                    "OK",
                    "Error"
                )
                Update-Status "ACL Excel export failed."
            }
            finally {
                $form.Cursor = [System.Windows.Forms.Cursors]::Default
            }
        }
    })

# END CHANGE

# --- TAB 2: ACCESS CHECKER ---

$tabAccessChk = New-Object System.Windows.Forms.TabPage

$tabAccessChk.Text = "Access Checker"

# $tabControl.TabPages.Add($tabAccessChk) # Moved to end

  

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

  

# --- TAB 3: AUDIT MEMBERS ---

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

  

# --- TAB 4: BULK UPDATES ---

$tabUpdate = New-Object System.Windows.Forms.TabPage

$tabUpdate.Text = "Bulk Updates"

$tabControl.TabPages.Add($tabUpdate)

$tabControl.TabPages.Add($tabAccessChk) # Moved from Tab 2 section to be the last tab

  

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

            Update-Status "Transferred $($dgvMembers.SelectedRows.Count) user(s) to Bulk Updates tab."

        }

    })

$dgvMembers.ContextMenuStrip = $contextMenu

  

# Tab 1 Fetch (v2.0: input validation, script-block filter, nested group handling, DataTable binding, wait cursor)

$btnFetch.Add_Click({

        if ([string]::IsNullOrWhiteSpace($txtGroupSearch.Text)) {

            [System.Windows.Forms.MessageBox]::Show("Please enter a group name or email address.", "Validation", "OK", "Warning")

            return

        }

        $form.Cursor = [System.Windows.Forms.Cursors]::WaitCursor

        Update-Status "Fetching group members..."

        try {

            $group = Resolve-ADGroupByNameOrMail -SearchValue $txtGroupSearch.Text

            # First: get direct members to find sub-groups

            $directMembers = Get-ADGroupMember -Identity $group -ErrorAction Stop

            $subGroupObjects = $directMembers | Where-Object { $_.objectClass -eq 'group' }

  

            # Populate sub-groups DataGridView

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

  

            # Then: get all recursive user members

            $members = Get-ADGroupMember -Identity $group -Recursive |

            Where-Object { $_.objectClass -eq 'user' } |

            Get-ADUser -Properties mail, DisplayName, Department, Title |

            Select-Object DisplayName, SamAccountName, mail, Department, Title

            # DataTable for reliable DataGridView binding

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

  

# Tab 2 Export CSV

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

  

# Import File handler

$btnImportFile.Add_Click({

        $openFile = New-Object System.Windows.Forms.OpenFileDialog

        $openFile.Filter = "CSV/TXT Files (*.csv;*.txt)|*.csv;*.txt"

        if ($openFile.ShowDialog() -eq [System.Windows.Forms.DialogResult]::OK) {

            $lines = Get-Content $openFile.FileName | Where-Object { $_.Trim() -ne "" }

            foreach ($line in $lines) { $dgvInput.Rows.Add($line.Trim()) | Out-Null }

            Update-Status "Imported $($lines.Count) identifier(s) from file."

        }

    })

  

# Bulk Process (v2.0: script-block filters, input validation, confirmation dialog, wait cursor, success/error counts)

$ProcessGrid = {

    param($Action)

    $groupName = $txtTargetGroup.Text.Trim()

    $userList = New-Object System.Collections.Generic.List[string]

    foreach ($row in $dgvInput.Rows) { if ($row.Cells["colIdentifier"].Value) { $userList.Add($row.Cells["colIdentifier"].Value.ToString().Trim()) } }

    $users = $userList | Select-Object -Unique

    if ([string]::IsNullOrWhiteSpace($groupName) -or $users.Count -eq 0) { [System.Windows.Forms.MessageBox]::Show("Please provide both Target Group and User identifiers.", "Validation Error", "OK", "Warning") ; return }

    # Confirmation dialog before bulk operation

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

$btnClearTab1.Add_Click({ $txtGroupSearch.Text = ""; $dgvMembers.DataSource = $null; $dgvSubGroups.DataSource = $null; $lblCount.Text = "Sub-Groups: 0 | Total Members: 0"; $btnExport.Enabled = $false; Update-Status "Audit tab cleared." })

$btnClearTab2.Add_Click({ $txtTargetGroup.Text = ""; $dgvInput.Rows.Clear(); $txtUpdateLog.Clear(); Update-Status "Bulk Updates tab cleared." })

  

# Tab 4 Access Checker Logic

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

            # Resolve the user

            $adUser = Resolve-ADUserByMailOrSam -SearchValue $userSearch

            $userSam = $adUser.SamAccountName

            $userDisplay = (Get-ADUser $adUser -Properties DisplayName).DisplayName

  

            # Get all group memberships (recursive, including nested groups)

            $userGroups = @()

            try {

                # Use tokenGroups attribute for true recursive membership (includes nested groups)

                $userDN = (Get-ADUser $adUser).DistinguishedName

                $userEntry = [ADSI]"LDAP://$userDN"

                $userEntry.RefreshCache("tokenGroups")

                $userGroups = $userEntry.Properties["tokenGroups"] | ForEach-Object {

                    $sid = New-Object System.Security.Principal.SecurityIdentifier($_, 0)

                    try { (Get-ADGroup $sid -ErrorAction SilentlyContinue).Name } catch {}

                } | Where-Object { $_ }

            }
            catch {

                # Fallback: use Get-ADPrincipalGroupMembership (non-recursive)

                try {

                    $userGroups = Get-ADPrincipalGroupMembership $adUser -ErrorAction SilentlyContinue |

                    Select-Object -ExpandProperty Name

                }
                catch {

                    $userObj = Get-ADUser $adUser -Properties MemberOf

                    $userGroups = $userObj.MemberOf | ForEach-Object { (Get-ADGroup $_).Name }

                }

            }

  

            # Get ACL

            $acl = Get-Acl -Path $path

  

            # Build results DataTable

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

  

            # Populate matched groups DataGridView with membership path

            $dtGroups = New-Object System.Data.DataTable

            $dtGroups.Columns.Add("ACL Group") | Out-Null

            $dtGroups.Columns.Add("Membership") | Out-Null

            $matchedGroupNames = @()

  

            foreach ($ace in $acl.Access) {

                $gName = ($ace.IdentityReference.Value -split '\\')[-1]

                if ($gName -in $userGroups -and $gName -notin $matchedGroupNames) {

                    $matchedGroupNames += $gName

  

                    # Check if user is a direct member or via nested group

                    $membershipPath = "Direct Member"

                    try {

                        $directMembers = Get-ADGroupMember -Identity $gName -ErrorAction SilentlyContinue

                        $directSams = $directMembers | Where-Object { $_.objectClass -eq 'user' } | Select-Object -ExpandProperty SamAccountName

                        if ($userSam -notin $directSams) {

                            # User is not a direct member, find which sub-group they're in

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

  

$btnClearTab4.Add_Click({ $txtAccessPath.Text = ""; $txtAccessUser.Text = ""; $dgvAccess.DataSource = $null; $dgvMatchedGroups.DataSource = $null; $lblAccessInfo.Text = ""; Update-Status "Access Checker tab cleared." })

  

$form.ShowDialog()
