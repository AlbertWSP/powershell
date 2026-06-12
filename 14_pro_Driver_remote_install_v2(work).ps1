Add-Type -AssemblyName System.Windows.Forms
Add-Type -AssemblyName System.Drawing

# ══════════════════════════════════════════════════════════════════════════════
#  Setup File Deployment v5 – Live Progress & In-Form Deployment
#  Supports Dell Update Packages (DUP) with /s /l /r /p switches
#  Changelog v5:
#    - [NEW] Live progress bar inside the GUI
#    - [NEW] Status label showing current operation in real time
#    - [NEW] Color-coded log panel (RichTextBox) for deployment output
#    - [NEW] Deployment runs inside the form – no more blind console wait
#    - [NEW] Stop button to cancel mid-deployment
#    - [NEW] Export CSV button appears after deployment
#    - [NEW] All controls disabled during deployment to prevent conflicts
#    - Carries forward all v4 features (folder scan, BIOS pwd, exit codes, etc.)
# ══════════════════════════════════════════════════════════════════════════════

# ── Script-level variables ─────────────────────────────────────────────────────
$script:exeFileMap = @{}
$script:cancelRequested = $false
$script:deployResults = @()
$script:isDeploying = $false

# ── Dell DUP Exit Codes Reference ─────────────────────────────────────────────
$script:exitCodeMap = @{
    -1 = "CANCELLED / DCU TIMEOUT"
    0  = "SUCCESS (no reboot needed)"
    1  = "FAILURE (general error)"
    2  = "SUCCESS (reboot required)"
    3  = "SOFT DEPENDENCY ERROR (same version or downgrade)"
    4  = "HARD DEPENDENCY ERROR (prerequisites not met)"
    5  = "QUALIFICATION ERROR (not applicable to system)"
    6  = "REBOOTING SYSTEM"
    7  = "PASSWORD ERROR (BIOS password not provided or incorrect)"
    8  = "DOWNGRADE NOT ALLOWED"
    9  = "RPM VERIFY FAILED (Linux only)"
    10 = "UNSPECIFIED ERROR (battery/EC/HW failure)"
}

# ── Helper: Append colored text to RichTextBox ────────────────────────────────
function Write-Log {
    param(
        [System.Windows.Forms.RichTextBox]$LogBox,
        [string]$Message,
        [System.Drawing.Color]$Color = [System.Drawing.Color]::White
    )
    $LogBox.SelectionStart = $LogBox.TextLength
    $LogBox.SelectionLength = 0
    $LogBox.SelectionColor = $Color
    $LogBox.AppendText("$Message`r`n")
    $LogBox.ScrollToCaret()
    [System.Windows.Forms.Application]::DoEvents()
}

# ── Helper: Update progress bar & status label ────────────────────────────────
function Update-Progress {
    param(
        [System.Windows.Forms.ProgressBar]$Bar,
        [System.Windows.Forms.Label]$Label,
        [int]$Value,
        [int]$Maximum,
        [string]$Text
    )
    if ($Maximum -gt 0) { $Bar.Maximum = $Maximum }
    $Bar.Value = [Math]::Min($Value, $Bar.Maximum)
    $Label.Text = $Text
    [System.Windows.Forms.Application]::DoEvents()
}

# ── Helper: Toggle controls enabled/disabled ──────────────────────────────────
function Set-ControlsEnabled {
    param([bool]$Enabled)
    $textBoxFolder.Enabled = $Enabled
    $btnBrowse.Enabled = $Enabled
    $btnScan.Enabled = $Enabled
    $checkedListBox.Enabled = $Enabled
    $btnSelectAll.Enabled = $Enabled
    $btnDeselectAll.Enabled = $Enabled
    $textBoxHostName.Enabled = $Enabled
    $radioCopyOnly.Enabled = $Enabled
    $radioCopyInstall.Enabled = $Enabled
    $textBoxBiosPassword.Enabled = $Enabled
    $chkShowPassword.Enabled = $Enabled
    $chkReboot.Enabled = $Enabled
    $okButton.Enabled = $Enabled
    $btnExportCsv.Visible = (-not $Enabled) -and ($script:deployResults.Count -gt 0)
}

# ══════════════════════════════════════════════════════════════════════════════
#  GUI SETUP
# ══════════════════════════════════════════════════════════════════════════════

$form = New-Object System.Windows.Forms.Form
$form.Text = "Setup File Deployment v5"
$form.Size = New-Object System.Drawing.Size(580, 980)
$form.StartPosition = "CenterScreen"
$form.FormBorderStyle = "FixedDialog"
$form.MaximizeBox = $false
$form.BackColor = [System.Drawing.Color]::FromArgb(30, 30, 30)
$form.ForeColor = [System.Drawing.Color]::White

$yPos = 12

# ── Source Folder ──────────────────────────────────────────────────────────────

$labelFolder = New-Object System.Windows.Forms.Label
$labelFolder.Text = "Source Folder (contains .exe files):"
$labelFolder.Location = New-Object System.Drawing.Point(10, $yPos)
$labelFolder.Size = New-Object System.Drawing.Size(540, 20)
$form.Controls.Add($labelFolder)
$yPos += 22

$textBoxFolder = New-Object System.Windows.Forms.TextBox
$textBoxFolder.Location = New-Object System.Drawing.Point(10, $yPos)
$textBoxFolder.Size = New-Object System.Drawing.Size(370, 22)
$textBoxFolder.Font = New-Object System.Drawing.Font("Consolas", 9)
$textBoxFolder.Text = "\\corp\hk\Transition\DAT18\Kits\Software\Engineering"
$form.Controls.Add($textBoxFolder)

$btnBrowse = New-Object System.Windows.Forms.Button
$btnBrowse.Text = "Browse..."
$btnBrowse.Location = New-Object System.Drawing.Point(385, $yPos)
$btnBrowse.Size = New-Object System.Drawing.Size(80, 24)
$form.Controls.Add($btnBrowse)

$btnScan = New-Object System.Windows.Forms.Button
$btnScan.Text = "Scan"
$btnScan.Location = New-Object System.Drawing.Point(470, $yPos)
$btnScan.Size = New-Object System.Drawing.Size(80, 24)
$btnScan.Font = New-Object System.Drawing.Font("Segoe UI", 9, [System.Drawing.FontStyle]::Bold)
$btnScan.BackColor = [System.Drawing.Color]::FromArgb(0, 120, 212)
$btnScan.ForeColor = [System.Drawing.Color]::White
$btnScan.FlatStyle = "Flat"
$form.Controls.Add($btnScan)
$yPos += 30

# Browse button handler
$btnBrowse.Add_Click({
        $folderDialog = New-Object System.Windows.Forms.FolderBrowserDialog
        $folderDialog.Description = "Select the folder containing .exe setup files"
        $folderDialog.ShowNewFolderButton = $false
        if ($textBoxFolder.Text -and (Test-Path -Path $textBoxFolder.Text -ErrorAction SilentlyContinue)) {
            $folderDialog.SelectedPath = $textBoxFolder.Text
        }
        if ($folderDialog.ShowDialog() -eq [System.Windows.Forms.DialogResult]::OK) {
            $textBoxFolder.Text = $folderDialog.SelectedPath
        }
    })

# ── Scanned .exe List ─────────────────────────────────────────────────────────

$labelExeList = New-Object System.Windows.Forms.Label
$labelExeList.Text = "Available .exe files (check to deploy):"
$labelExeList.Location = New-Object System.Drawing.Point(10, $yPos)
$labelExeList.Size = New-Object System.Drawing.Size(400, 20)
$form.Controls.Add($labelExeList)

$labelExeCount = New-Object System.Windows.Forms.Label
$labelExeCount.Text = ""
$labelExeCount.Location = New-Object System.Drawing.Point(400, $yPos)
$labelExeCount.Size = New-Object System.Drawing.Size(150, 20)
$labelExeCount.ForeColor = [System.Drawing.Color]::Gray
$labelExeCount.TextAlign = [System.Drawing.ContentAlignment]::MiddleRight
$form.Controls.Add($labelExeCount)
$yPos += 22

$checkedListBox = New-Object System.Windows.Forms.CheckedListBox
$checkedListBox.Location = New-Object System.Drawing.Point(10, $yPos)
$checkedListBox.Size = New-Object System.Drawing.Size(540, 130)
$checkedListBox.Font = New-Object System.Drawing.Font("Consolas", 9)
$checkedListBox.CheckOnClick = $true
$form.Controls.Add($checkedListBox)
$yPos += 135

$btnSelectAll = New-Object System.Windows.Forms.Button
$btnSelectAll.Text = "Select All"
$btnSelectAll.Location = New-Object System.Drawing.Point(10, $yPos)
$btnSelectAll.Size = New-Object System.Drawing.Size(90, 25)
$form.Controls.Add($btnSelectAll)

$btnDeselectAll = New-Object System.Windows.Forms.Button
$btnDeselectAll.Text = "Deselect All"
$btnDeselectAll.Location = New-Object System.Drawing.Point(105, $yPos)
$btnDeselectAll.Size = New-Object System.Drawing.Size(90, 25)
$form.Controls.Add($btnDeselectAll)

$btnSelectAll.Add_Click({
        for ($i = 0; $i -lt $checkedListBox.Items.Count; $i++) {
            $checkedListBox.SetItemChecked($i, $true)
        }
    })
$btnDeselectAll.Add_Click({
        for ($i = 0; $i -lt $checkedListBox.Items.Count; $i++) {
            $checkedListBox.SetItemChecked($i, $false)
        }
    })
$yPos += 32

# ── Scan button handler ───────────────────────────────────────────────────────

$btnScan.Add_Click({
        $checkedListBox.Items.Clear()
        $script:exeFileMap = @{}
        $folder = $textBoxFolder.Text.Trim()

        if (-not $folder -or -not (Test-Path -Path $folder -ErrorAction SilentlyContinue)) {
            [System.Windows.Forms.MessageBox]::Show(
                "Folder not found or inaccessible:`n$folder",
                "Scan Error",
                [System.Windows.Forms.MessageBoxButtons]::OK,
                [System.Windows.Forms.MessageBoxIcon]::Warning
            )
            $labelExeCount.Text = ""
            return
        }

        $form.Cursor = [System.Windows.Forms.Cursors]::WaitCursor
        try {
            $exeFiles = Get-ChildItem -Path $folder -Filter "*.exe" -Recurse -Depth 1 -File -ErrorAction SilentlyContinue |
            Sort-Object Name

            if ($exeFiles.Count -eq 0) {
                [System.Windows.Forms.MessageBox]::Show(
                    "No .exe files found in:`n$folder`n`n(Scanned up to 1 subfolder deep)",
                    "No Files Found",
                    [System.Windows.Forms.MessageBoxButtons]::OK,
                    [System.Windows.Forms.MessageBoxIcon]::Information
                )
                $labelExeCount.Text = "0 files"
                return
            }

            foreach ($exe in $exeFiles) {
                $relativePath = $exe.FullName.Substring($folder.TrimEnd('\').Length + 1)
                $displayName = $relativePath
                $checkedListBox.Items.Add($displayName)
                $script:exeFileMap[$displayName] = $exe.FullName
            }
            $labelExeCount.Text = "$($exeFiles.Count) file(s) found"
        }
        catch {
            [System.Windows.Forms.MessageBox]::Show(
                "Error scanning folder:`n$($_.Exception.Message)",
                "Scan Error",
                [System.Windows.Forms.MessageBoxButtons]::OK,
                [System.Windows.Forms.MessageBoxIcon]::Error
            )
        }
        finally {
            $form.Cursor = [System.Windows.Forms.Cursors]::Default
        }
    })

# ── Host Name Input ────────────────────────────────────────────────────────────

$labelHostName = New-Object System.Windows.Forms.Label
$labelHostName.Text = "Target Host Names (one per line):"
$labelHostName.Location = New-Object System.Drawing.Point(10, $yPos)
$labelHostName.Size = New-Object System.Drawing.Size(540, 20)
$form.Controls.Add($labelHostName)
$yPos += 22

$textBoxHostName = New-Object System.Windows.Forms.TextBox
$textBoxHostName.Location = New-Object System.Drawing.Point(10, $yPos)
$textBoxHostName.Size = New-Object System.Drawing.Size(540, 70)
$textBoxHostName.Multiline = $true
$textBoxHostName.ScrollBars = "Vertical"
$textBoxHostName.Font = New-Object System.Drawing.Font("Consolas", 10)
$form.Controls.Add($textBoxHostName)
$yPos += 78

# ── Action Mode ────────────────────────────────────────────────────────────────

$grpAction = New-Object System.Windows.Forms.GroupBox
$grpAction.Text = "Action"
$grpAction.Location = New-Object System.Drawing.Point(10, $yPos)
$grpAction.Size = New-Object System.Drawing.Size(540, 50)
$grpAction.ForeColor = [System.Drawing.Color]::White
$form.Controls.Add($grpAction)

$radioCopyOnly = New-Object System.Windows.Forms.RadioButton
$radioCopyOnly.Text = "Copy Only"
$radioCopyOnly.Location = New-Object System.Drawing.Point(15, 20)
$radioCopyOnly.Size = New-Object System.Drawing.Size(150, 25)
$radioCopyOnly.Checked = $true
$radioCopyOnly.ForeColor = [System.Drawing.Color]::White
$grpAction.Controls.Add($radioCopyOnly)

$radioCopyInstall = New-Object System.Windows.Forms.RadioButton
$radioCopyInstall.Text = "Copy + Silent Install"
$radioCopyInstall.Location = New-Object System.Drawing.Point(220, 20)
$radioCopyInstall.Size = New-Object System.Drawing.Size(200, 25)
$radioCopyInstall.ForeColor = [System.Drawing.Color]::White
$grpAction.Controls.Add($radioCopyInstall)
$yPos += 56

# ── BIOS Password ─────────────────────────────────────────────────────────────

$grpBios = New-Object System.Windows.Forms.GroupBox
$grpBios.Text = "BIOS Password (for BIOS/Firmware DUPs)"
$grpBios.Location = New-Object System.Drawing.Point(10, $yPos)
$grpBios.Size = New-Object System.Drawing.Size(540, 70)
$grpBios.ForeColor = [System.Drawing.Color]::White
$form.Controls.Add($grpBios)

$labelBios = New-Object System.Windows.Forms.Label
$labelBios.Text = "Password:"
$labelBios.Location = New-Object System.Drawing.Point(10, 28)
$labelBios.Size = New-Object System.Drawing.Size(70, 20)
$grpBios.Controls.Add($labelBios)

$textBoxBiosPassword = New-Object System.Windows.Forms.TextBox
$textBoxBiosPassword.Location = New-Object System.Drawing.Point(85, 25)
$textBoxBiosPassword.Size = New-Object System.Drawing.Size(200, 22)
$textBoxBiosPassword.PasswordChar = '*'
$textBoxBiosPassword.Font = New-Object System.Drawing.Font("Consolas", 10)
$grpBios.Controls.Add($textBoxBiosPassword)

$chkShowPassword = New-Object System.Windows.Forms.CheckBox
$chkShowPassword.Text = "Show"
$chkShowPassword.Location = New-Object System.Drawing.Point(295, 26)
$chkShowPassword.Size = New-Object System.Drawing.Size(70, 22)
$chkShowPassword.ForeColor = [System.Drawing.Color]::White
$chkShowPassword.Add_CheckedChanged({
        if ($chkShowPassword.Checked) { $textBoxBiosPassword.PasswordChar = [char]0 }
        else { $textBoxBiosPassword.PasswordChar = '*' }
    })
$grpBios.Controls.Add($chkShowPassword)

$labelBiosHint = New-Object System.Windows.Forms.Label
$labelBiosHint.Text = "Leave blank if no BIOS password is set on the target PCs."
$labelBiosHint.Location = New-Object System.Drawing.Point(10, 50)
$labelBiosHint.Size = New-Object System.Drawing.Size(520, 16)
$labelBiosHint.ForeColor = [System.Drawing.Color]::Gray
$labelBiosHint.Font = New-Object System.Drawing.Font("Segoe UI", 8)
$grpBios.Controls.Add($labelBiosHint)
$yPos += 76

# ── Post-Install Options ──────────────────────────────────────────────────────

$grpPost = New-Object System.Windows.Forms.GroupBox
$grpPost.Text = "Post-Install Options"
$grpPost.Location = New-Object System.Drawing.Point(10, $yPos)
$grpPost.Size = New-Object System.Drawing.Size(540, 50)
$grpPost.ForeColor = [System.Drawing.Color]::White
$form.Controls.Add($grpPost)

$chkReboot = New-Object System.Windows.Forms.CheckBox
$chkReboot.Text = "Auto-reboot after install (if required by firmware)"
$chkReboot.Location = New-Object System.Drawing.Point(15, 20)
$chkReboot.Size = New-Object System.Drawing.Size(500, 25)
$chkReboot.ForeColor = [System.Drawing.Color]::White
$grpPost.Controls.Add($chkReboot)
$yPos += 56

# ── Progress Section ───────────────────────────────────────────────────────────

$grpProgress = New-Object System.Windows.Forms.GroupBox
$grpProgress.Text = "Deployment Progress"
$grpProgress.Location = New-Object System.Drawing.Point(10, $yPos)
$grpProgress.Size = New-Object System.Drawing.Size(540, 50)
$grpProgress.ForeColor = [System.Drawing.Color]::White
$form.Controls.Add($grpProgress)

$progressBar = New-Object System.Windows.Forms.ProgressBar
$progressBar.Location = New-Object System.Drawing.Point(10, 20)
$progressBar.Size = New-Object System.Drawing.Size(420, 22)
$progressBar.Minimum = 0
$progressBar.Maximum = 100
$progressBar.Value = 0
$progressBar.Style = "Continuous"
$grpProgress.Controls.Add($progressBar)

$labelPercent = New-Object System.Windows.Forms.Label
$labelPercent.Text = "0%"
$labelPercent.Location = New-Object System.Drawing.Point(438, 22)
$labelPercent.Size = New-Object System.Drawing.Size(90, 20)
$labelPercent.ForeColor = [System.Drawing.Color]::FromArgb(0, 200, 0)
$labelPercent.Font = New-Object System.Drawing.Font("Consolas", 9, [System.Drawing.FontStyle]::Bold)
$grpProgress.Controls.Add($labelPercent)
$yPos += 55

$labelStatus = New-Object System.Windows.Forms.Label
$labelStatus.Text = "Ready."
$labelStatus.Location = New-Object System.Drawing.Point(10, $yPos)
$labelStatus.Size = New-Object System.Drawing.Size(540, 18)
$labelStatus.ForeColor = [System.Drawing.Color]::FromArgb(180, 180, 180)
$labelStatus.Font = New-Object System.Drawing.Font("Segoe UI", 8.5)
$form.Controls.Add($labelStatus)
$yPos += 22

# ── Log Panel (RichTextBox) ───────────────────────────────────────────────────

$richLog = New-Object System.Windows.Forms.RichTextBox
$richLog.Location = New-Object System.Drawing.Point(10, $yPos)
$richLog.Size = New-Object System.Drawing.Size(540, 180)
$richLog.Font = New-Object System.Drawing.Font("Consolas", 8.5)
$richLog.BackColor = [System.Drawing.Color]::FromArgb(20, 20, 20)
$richLog.ForeColor = [System.Drawing.Color]::White
$richLog.ReadOnly = $true
$richLog.WordWrap = $false
$richLog.ScrollBars = [System.Windows.Forms.RichTextBoxScrollBars]::Both
$form.Controls.Add($richLog)
$yPos += 188

# ── Buttons ────────────────────────────────────────────────────────────────────

$okButton = New-Object System.Windows.Forms.Button
$okButton.Text = [char]0x25B6 + "  Deploy"
$okButton.Size = New-Object System.Drawing.Size(120, 34)
$okButton.Location = New-Object System.Drawing.Point(140, $yPos)
$okButton.Font = New-Object System.Drawing.Font("Segoe UI", 9.5, [System.Drawing.FontStyle]::Bold)
$okButton.BackColor = [System.Drawing.Color]::FromArgb(0, 120, 212)
$okButton.ForeColor = [System.Drawing.Color]::White
$okButton.FlatStyle = "Flat"
$form.Controls.Add($okButton)

$cancelButton = New-Object System.Windows.Forms.Button
$cancelButton.Text = "Close"
$cancelButton.Size = New-Object System.Drawing.Size(100, 34)
$cancelButton.Location = New-Object System.Drawing.Point(270, $yPos)
$cancelButton.FlatStyle = "Flat"
$form.Controls.Add($cancelButton)

$btnExportCsv = New-Object System.Windows.Forms.Button
$btnExportCsv.Text = "Export CSV"
$btnExportCsv.Size = New-Object System.Drawing.Size(100, 34)
$btnExportCsv.Location = New-Object System.Drawing.Point(380, $yPos)
$btnExportCsv.FlatStyle = "Flat"
$btnExportCsv.BackColor = [System.Drawing.Color]::FromArgb(50, 50, 50)
$btnExportCsv.ForeColor = [System.Drawing.Color]::White
$btnExportCsv.Visible = $false
$form.Controls.Add($btnExportCsv)

$cancelButton.Add_Click({
        if ($script:isDeploying) {
            $script:cancelRequested = $true
            $cancelButton.Enabled = $false
            $labelStatus.Text = "Stopping after current task..."
            $labelStatus.ForeColor = [System.Drawing.Color]::Orange
        }
        else {
            $form.Close()
        }
    })

# ── Export CSV handler ─────────────────────────────────────────────────────────

$btnExportCsv.Add_Click({
        if ($script:deployResults.Count -eq 0) {
            [System.Windows.Forms.MessageBox]::Show("No results to export.", "Export",
                [System.Windows.Forms.MessageBoxButtons]::OK,
                [System.Windows.Forms.MessageBoxIcon]::Information)
            return
        }
        $saveDialog = New-Object System.Windows.Forms.SaveFileDialog
        $saveDialog.Filter = "CSV files (*.csv)|*.csv"
        $saveDialog.FileName = "DeploymentLog_$(Get-Date -Format 'yyyyMMdd_HHmmss').csv"
        $saveDialog.InitialDirectory = "C:\temp"
        if ($saveDialog.ShowDialog() -eq [System.Windows.Forms.DialogResult]::OK) {
            try {
                $script:deployResults | Export-Csv -Path $saveDialog.FileName -NoTypeInformation -Encoding UTF8
                Write-Log $richLog "CSV exported: $($saveDialog.FileName)" ([System.Drawing.Color]::FromArgb(0, 200, 0))
                [System.Windows.Forms.MessageBox]::Show("Exported to:`n$($saveDialog.FileName)", "Export Successful",
                    [System.Windows.Forms.MessageBoxButtons]::OK,
                    [System.Windows.Forms.MessageBoxIcon]::Information)
            }
            catch {
                [System.Windows.Forms.MessageBox]::Show("Export failed:`n$($_.Exception.Message)", "Export Error",
                    [System.Windows.Forms.MessageBoxButtons]::OK,
                    [System.Windows.Forms.MessageBoxIcon]::Error)
            }
        }
    })

# ══════════════════════════════════════════════════════════════════════════════
#  DEPLOY BUTTON HANDLER – Runs deployment inside the form
# ══════════════════════════════════════════════════════════════════════════════

$okButton.Add_Click({

        # ── Collect user selections ────────────────────────────────────────────
        $doInstall = $radioCopyInstall.Checked
        $doReboot = $chkReboot.Checked
        $biosPassword = $textBoxBiosPassword.Text.Trim()
        $hostInput = $textBoxHostName.Text
        $folderPath = $textBoxFolder.Text.Trim()
        $modeLabel = $(if ($doInstall) { "Copy + Silent Install" } else { "Copy Only" })

        # Collect checked .exe items
        $selectedExes = @()
        for ($i = 0; $i -lt $checkedListBox.Items.Count; $i++) {
            if ($checkedListBox.GetItemChecked($i)) {
                $displayName = $checkedListBox.Items[$i]
                $fullPath = $script:exeFileMap[$displayName]
                $selectedExes += [PSCustomObject]@{
                    DisplayName = $displayName
                    FullPath    = $fullPath
                    FileName    = [System.IO.Path]::GetFileName($fullPath)
                    LogName     = [System.IO.Path]::GetFileNameWithoutExtension($fullPath) + ".log"
                }
            }
        }

        # Validate
        $hostNames = $hostInput -split "`r`n" | ForEach-Object { $_.Trim() } | Where-Object { $_ -match '\S' }
        if ($hostNames.Count -eq 0) {
            [System.Windows.Forms.MessageBox]::Show("No host names entered.", "Validation",
                [System.Windows.Forms.MessageBoxButtons]::OK, [System.Windows.Forms.MessageBoxIcon]::Warning)
            return
        }
        if ($selectedExes.Count -eq 0) {
            [System.Windows.Forms.MessageBox]::Show("No .exe files selected.", "Validation",
                [System.Windows.Forms.MessageBoxButtons]::OK, [System.Windows.Forms.MessageBoxIcon]::Warning)
            return
        }

        # Build deployment map
        $deployments = [ordered]@{}
        foreach ($exe in $selectedExes) {
            $deployments[$exe.DisplayName] = @{
                Source   = $exe.FullPath
                FileName = $exe.FileName
                LogName  = $exe.LogName
            }
        }

        # Pre-check source files
        foreach ($label in $deployments.Keys) {
            $src = $deployments[$label].Source
            if (!(Test-Path -Path $src -ErrorAction SilentlyContinue)) {
                [System.Windows.Forms.MessageBox]::Show("Source file not found for [$label]:`n$src", "File Missing",
                    [System.Windows.Forms.MessageBoxButtons]::OK, [System.Windows.Forms.MessageBoxIcon]::Error)
                return
            }
        }

        # ── Begin deployment ───────────────────────────────────────────────────
        $script:isDeploying = $true
        $script:cancelRequested = $false
        $script:deployResults = @()
        $successCount = 0
        $failCount = 0
        $totalTasks = $hostNames.Count * $deployments.Count
        $taskNum = 0

        # Lock controls
        Set-ControlsEnabled -Enabled $false
        $cancelButton.Text = [char]0x23F9 + "  Stop"
        $cancelButton.Enabled = $true
        $cancelButton.BackColor = [System.Drawing.Color]::FromArgb(200, 50, 50)
        $cancelButton.ForeColor = [System.Drawing.Color]::White
        $btnExportCsv.Visible = $false

        # Reset progress
        $progressBar.Value = 0
        $progressBar.Maximum = $totalTasks
        $labelPercent.Text = "0%"
        $richLog.Clear()

        # Banner
        Write-Log $richLog "=================================================" ([System.Drawing.Color]::Cyan)
        Write-Log $richLog "  Setup File Deployment v5" ([System.Drawing.Color]::Cyan)
        Write-Log $richLog "=================================================" ([System.Drawing.Color]::Cyan)
        Write-Log $richLog "  Source : $folderPath" ([System.Drawing.Color]::Cyan)
        Write-Log $richLog "  Mode   : $modeLabel" ([System.Drawing.Color]::Cyan)
        Write-Log $richLog "  Hosts  : $($hostNames.Count) | Packages: $($deployments.Count)" ([System.Drawing.Color]::Cyan)
        Write-Log $richLog "  Total  : $totalTasks task(s)" ([System.Drawing.Color]::Cyan)
        if ($doInstall -and $biosPassword) {
            Write-Log $richLog "  BIOS   : Password provided" ([System.Drawing.Color]::Gray)
        }
        if ($doInstall -and $doReboot) {
            Write-Log $richLog "  Reboot : Enabled (if required)" ([System.Drawing.Color]::Magenta)
        }
        Write-Log $richLog "=================================================" ([System.Drawing.Color]::Cyan)
        Write-Log $richLog "" ([System.Drawing.Color]::White)

        # ── Main deployment loop ───────────────────────────────────────────────
        foreach ($name in $hostNames) {

            if ($script:cancelRequested) {
                Write-Log $richLog "Deployment cancelled by user." ([System.Drawing.Color]::Orange)
                break
            }

            Write-Log $richLog "--- Processing: $name ---" ([System.Drawing.Color]::Cyan)
            Update-Progress $progressBar $labelStatus $taskNum $totalTasks "Pinging $name..."

            # Connectivity check
            if (!(Test-Connection -ComputerName $name -Count 1 -Quiet)) {
                Write-Log $richLog "  [!] $name is OFFLINE or unreachable." ([System.Drawing.Color]::Red)
                $failCount += $deployments.Count
                $taskNum += $deployments.Count
                foreach ($label in $deployments.Keys) {
                    $script:deployResults += [PSCustomObject]@{ Host = $name; Package = $label; Step = "Ping"; Status = "OFFLINE" }
                }
                $pct = [math]::Round(($taskNum / $totalTasks) * 100)
                $progressBar.Value = [Math]::Min($taskNum, $totalTasks)
                $labelPercent.Text = "$pct%"
                [System.Windows.Forms.Application]::DoEvents()
                continue
            }

            Write-Log $richLog "  [OK] $name is online." ([System.Drawing.Color]::FromArgb(0, 200, 0))

            # WinRM pre-check
            $winrmOk = $true
            if ($doInstall) {
                Update-Progress $progressBar $labelStatus $taskNum $totalTasks "Checking WinRM on $name..."
                $winrmTest = Test-WSMan -ComputerName $name -ErrorAction SilentlyContinue
                if (-not $winrmTest) {
                    Write-Log $richLog "  [!] WinRM unavailable on $name. Install will be skipped." ([System.Drawing.Color]::Yellow)
                    $winrmOk = $false
                }
                else {
                    Write-Log $richLog "  [OK] WinRM OK." ([System.Drawing.Color]::FromArgb(0, 200, 0))
                }
            }

            # Ensure C$\temp exists
            $destinationPath = "\\$name\C$\temp\"
            try {
                if (!(Test-Path -Path $destinationPath)) {
                    New-Item -ItemType Directory -Path $destinationPath -Force | Out-Null
                    Write-Log $richLog "  Created $destinationPath" ([System.Drawing.Color]::Gray)
                }
            }
            catch {
                Write-Log $richLog "  [X] Cannot create ${destinationPath}: $($_.Exception.Message)" ([System.Drawing.Color]::Red)
                $failCount += $deployments.Count
                $taskNum += $deployments.Count
                foreach ($label in $deployments.Keys) {
                    $script:deployResults += [PSCustomObject]@{ Host = $name; Package = $label; Step = "Mkdir"; Status = "FAILED" }
                }
                $pct = [math]::Round(($taskNum / $totalTasks) * 100)
                $progressBar.Value = [Math]::Min($taskNum, $totalTasks)
                $labelPercent.Text = "$pct%"
                [System.Windows.Forms.Application]::DoEvents()
                continue
            }

            # Process each package
            foreach ($label in $deployments.Keys) {

                if ($script:cancelRequested) {
                    Write-Log $richLog "Deployment cancelled by user." ([System.Drawing.Color]::Orange)
                    break
                }

                $pkg = $deployments[$label]
                $taskNum++
                $pct = [math]::Round(($taskNum / $totalTasks) * 100)

                # ── Copy ──────────────────────────────────────────────────────
                Update-Progress $progressBar $labelStatus $taskNum $totalTasks "[$taskNum/$totalTasks] Copying $($pkg.FileName) to $name..."
                try {
                    Copy-Item -Path $pkg.Source -Destination $destinationPath -Force
                    Write-Log $richLog "  [$taskNum/$totalTasks][$label] Copy... OK" ([System.Drawing.Color]::FromArgb(0, 200, 0))
                }
                catch {
                    Write-Log $richLog "  [$taskNum/$totalTasks][$label] Copy... FAILED: $($_.Exception.Message)" ([System.Drawing.Color]::Red)
                    $failCount++
                    $script:deployResults += [PSCustomObject]@{ Host = $name; Package = $label; Step = "Copy"; Status = "FAILED" }
                    $progressBar.Value = [Math]::Min($taskNum, $totalTasks)
                    $labelPercent.Text = "$pct%"
                    [System.Windows.Forms.Application]::DoEvents()
                    continue
                }

                # ── Install ───────────────────────────────────────────────────
                if ($doInstall) {
                    if (-not $winrmOk) {
                        Write-Log $richLog "  [$taskNum/$totalTasks][$label] Install SKIPPED (WinRM unavailable)" ([System.Drawing.Color]::Yellow)
                        $failCount++
                        $script:deployResults += [PSCustomObject]@{ Host = $name; Package = $label; Step = "Install"; Status = "SKIPPED (WinRM unavailable)" }
                        $progressBar.Value = [Math]::Min($taskNum, $totalTasks)
                        $labelPercent.Text = "$pct%"
                        [System.Windows.Forms.Application]::DoEvents()
                        continue
                    }

                    Update-Progress $progressBar $labelStatus $taskNum $totalTasks "[$taskNum/$totalTasks] Installing $($pkg.FileName) on $name..."
                    try {
                        $remoteExe = "C:\temp\$($pkg.FileName)"
                        $remoteLog = "C:\temp\$($pkg.LogName)"
                        $installArgs = "/s /l=$remoteLog"
                        if ($biosPassword) { $installArgs += " /p=$biosPassword" }
                        if ($doReboot) { $installArgs += " /r" }

                        $exitCode = Invoke-Command -ComputerName $name -ScriptBlock {
                            param($exe, $arguments)
                            $process = Start-Process -FilePath $exe -ArgumentList $arguments -Wait -PassThru -NoNewWindow
                            return $process.ExitCode
                        } -ArgumentList $remoteExe, $installArgs -ErrorAction Stop

                        $exitDesc = if ($script:exitCodeMap.ContainsKey([int]$exitCode)) {
                            $script:exitCodeMap[[int]$exitCode]
                        }
                        elseif ($exitCode -lt 0) {
                            $hexCode = "0x{0:X8}" -f ([uint32]$exitCode)
                            "APPLICATION CRASH (Exit Code: $exitCode / $hexCode)"
                        }
                        else {
                            "UNKNOWN (Exit Code: $exitCode)"
                        }

                        if ($exitCode -eq 0 -or $exitCode -eq 2 -or $exitCode -eq 6) {
                            $clr = if ($exitCode -eq 0) { [System.Drawing.Color]::FromArgb(0, 200, 0) } else { [System.Drawing.Color]::Yellow }
                            Write-Log $richLog "  [$taskNum/$totalTasks][$label] $exitDesc" $clr
                            $successCount++
                        }
                        else {
                            Write-Log $richLog "  [$taskNum/$totalTasks][$label] $exitDesc" ([System.Drawing.Color]::Red)
                            $failCount++
                        }
                        $script:deployResults += [PSCustomObject]@{ Host = $name; Package = $label; Step = "Install"; Status = $exitDesc }
                    }
                    catch {
                        Write-Log $richLog "  [$taskNum/$totalTasks][$label] ERROR: $($_.Exception.Message)" ([System.Drawing.Color]::Red)
                        $failCount++
                        $script:deployResults += [PSCustomObject]@{ Host = $name; Package = $label; Step = "Install"; Status = "ERROR: $($_.Exception.Message)" }
                    }
                }
                else {
                    $successCount++
                    $script:deployResults += [PSCustomObject]@{ Host = $name; Package = $label; Step = "Copy"; Status = "COPIED" }
                }

                $progressBar.Value = [Math]::Min($taskNum, $totalTasks)
                $labelPercent.Text = "$pct%"
                [System.Windows.Forms.Application]::DoEvents()
            }
        }

        # ── Deployment finished ────────────────────────────────────────────────
        $progressBar.Value = $progressBar.Maximum
        $labelPercent.Text = "100%"

        Write-Log $richLog "" ([System.Drawing.Color]::White)
        Write-Log $richLog "=================================================" ([System.Drawing.Color]::Yellow)
        Write-Log $richLog "  DEPLOYMENT COMPLETE" ([System.Drawing.Color]::Yellow)
        Write-Log $richLog "=================================================" ([System.Drawing.Color]::Yellow)
        Write-Log $richLog "  Mode    : $modeLabel" ([System.Drawing.Color]::White)
        $clrS = [System.Drawing.Color]::FromArgb(0, 200, 0)
        $clrF = if ($failCount -gt 0) { [System.Drawing.Color]::Red } else { [System.Drawing.Color]::FromArgb(0, 200, 0) }
        Write-Log $richLog "  Success : $successCount" $clrS
        Write-Log $richLog "  Failed  : $failCount" $clrF
        Write-Log $richLog "  Total   : $totalTasks" ([System.Drawing.Color]::Yellow)
        if ($script:cancelRequested) {
            Write-Log $richLog "  [!] Deployment was stopped early by user." ([System.Drawing.Color]::Orange)
        }
        Write-Log $richLog "=================================================" ([System.Drawing.Color]::Yellow)

        if ($doInstall) {
            Write-Log $richLog "" ([System.Drawing.Color]::White)
            Write-Log $richLog "Tip: Check remote logs at C:\temp\<LogName> on each host." ([System.Drawing.Color]::Gray)
        }

        # Restore controls
        $script:isDeploying = $false
        Set-ControlsEnabled -Enabled $true
        $cancelButton.Text = "Close"
        $cancelButton.Enabled = $true
        $cancelButton.BackColor = [System.Drawing.Color]::FromArgb(30, 30, 30)
        $cancelButton.ForeColor = [System.Drawing.Color]::White
        $btnExportCsv.Visible = ($script:deployResults.Count -gt 0)

        if ($failCount -eq 0) {
            $labelStatus.Text = "[OK] All $totalTasks task(s) completed successfully."
            $labelStatus.ForeColor = [System.Drawing.Color]::FromArgb(0, 200, 0)
        }
        else {
            $labelStatus.Text = "[!] Completed with $failCount failure(s) out of $totalTasks task(s)."
            $labelStatus.ForeColor = [System.Drawing.Color]::Orange
        }
    })

# ── Show Form ──────────────────────────────────────────────────────────────────
$form.Add_FormClosing({
        if ($script:isDeploying) {
            $result = [System.Windows.Forms.MessageBox]::Show(
                "Deployment is still running. Are you sure you want to close?",
                "Confirm Close",
                [System.Windows.Forms.MessageBoxButtons]::YesNo,
                [System.Windows.Forms.MessageBoxIcon]::Warning
            )
            if ($result -eq [System.Windows.Forms.DialogResult]::No) {
                $_.Cancel = $true
            }
            else {
                $script:cancelRequested = $true
            }
        }
    })

[void]$form.ShowDialog()
$form.Dispose()