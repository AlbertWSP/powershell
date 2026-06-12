Add-Type -AssemblyName System.Windows.Forms
Add-Type -AssemblyName System.Drawing

# ══════════════════════════════════════════════════════════════════════════════
#  Setup File Deployment v6 – Dynamic Folder Scan + Remote Silent Install + Pre-Copied
#  Supports Dell Update Packages (DUP) with /s /l /r /p switches
#  Changelog v6:
#    - [NEW] Pre-Copied mode – install drivers already on remote PC
#    - [NEW] Remote path field for pre-copied driver location
#  Changelog v5:
#    - [NEW] Dynamic folder scanning – browse any folder for .exe files
#    - [NEW] CheckedListBox with Select All / Deselect All
#    - [NEW] No more hardcoded package checkboxes – fully dynamic
#    - Carries forward all v4 fixes and features
#  Changelog v4:
#    - [FIX] Capture all GUI control states before Form.Dispose()
#    - [FIX] Added -ErrorAction Stop to Copy-Item / New-Item in try/catch
#    - [FIX] Null-safety guard for $exitCode from Invoke-Command
#    - [FIX] Reboot flag (/r) now only applied to the LAST package per host
#    - [NEW] PSSession timeout (10 min) to prevent hung installs
#    - [NEW] Elapsed time shown in deployment summary
# ══════════════════════════════════════════════════════════════════════════════

# ── GUI Setup ──────────────────────────────────────────────────────────────────

$form = New-Object System.Windows.Forms.Form
$form.Text = "Setup File Deployment v6"
$form.Size = New-Object System.Drawing.Size(520, 850)
$form.StartPosition = "CenterScreen"
$form.FormBorderStyle = "FixedDialog"
$form.MaximizeBox = $false

$yPos = 10

# ── Source Folder ──────────────────────────────────────────────────────────────

$labelFolder = New-Object System.Windows.Forms.Label
$labelFolder.Text = "Source Folder (contains .exe files):"
$labelFolder.Location = New-Object System.Drawing.Point(10, $yPos)
$labelFolder.Size = New-Object System.Drawing.Size(480, 20)
$form.Controls.Add($labelFolder)
$yPos += 22

$textBoxFolder = New-Object System.Windows.Forms.TextBox
$textBoxFolder.Location = New-Object System.Drawing.Point(10, $yPos)
$textBoxFolder.Size = New-Object System.Drawing.Size(330, 22)
$textBoxFolder.Font = New-Object System.Drawing.Font("Consolas", 9)
$textBoxFolder.Text = "\\corp\hk\Transition\DAT18\Kits\Software\Engineering"
$form.Controls.Add($textBoxFolder)

$btnBrowse = New-Object System.Windows.Forms.Button
$btnBrowse.Text = "Browse..."
$btnBrowse.Location = New-Object System.Drawing.Point(345, $yPos)
$btnBrowse.Size = New-Object System.Drawing.Size(75, 24)
$form.Controls.Add($btnBrowse)

$btnScan = New-Object System.Windows.Forms.Button
$btnScan.Text = "Scan"
$btnScan.Location = New-Object System.Drawing.Point(425, $yPos)
$btnScan.Size = New-Object System.Drawing.Size(65, 24)
$btnScan.Font = New-Object System.Drawing.Font("Segoe UI", 9, [System.Drawing.FontStyle]::Bold)
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
$labelExeCount.Location = New-Object System.Drawing.Point(350, $yPos)
$labelExeCount.Size = New-Object System.Drawing.Size(140, 20)
$labelExeCount.ForeColor = [System.Drawing.Color]::Gray
$labelExeCount.TextAlign = [System.Drawing.ContentAlignment]::MiddleRight
$form.Controls.Add($labelExeCount)
$yPos += 22

$checkedListBox = New-Object System.Windows.Forms.CheckedListBox
$checkedListBox.Location = New-Object System.Drawing.Point(10, $yPos)
$checkedListBox.Size = New-Object System.Drawing.Size(480, 150)
$checkedListBox.Font = New-Object System.Drawing.Font("Consolas", 9)
$checkedListBox.CheckOnClick = $true
$form.Controls.Add($checkedListBox)
$yPos += 155

# Select All / Deselect All buttons
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
# Stores full path mapping: display name -> full path

$script:exeFileMap = @{}

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
            # Scan for .exe files (top-level and one level deep)
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
                # Show relative path from the scanned folder for clarity
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
$labelHostName.Size = New-Object System.Drawing.Size(480, 20)
$form.Controls.Add($labelHostName)
$yPos += 22

$textBoxHostName = New-Object System.Windows.Forms.TextBox
$textBoxHostName.Location = New-Object System.Drawing.Point(10, $yPos)
$textBoxHostName.Size = New-Object System.Drawing.Size(480, 100)
$textBoxHostName.Multiline = $true
$textBoxHostName.ScrollBars = "Vertical"
$textBoxHostName.Font = New-Object System.Drawing.Font("Consolas", 10)
$form.Controls.Add($textBoxHostName)
$yPos += 108

# ── Action Mode (GroupBox) ─────────────────────────────────────────────────────

$grpAction = New-Object System.Windows.Forms.GroupBox
$grpAction.Text = "Action"
$grpAction.Location = New-Object System.Drawing.Point(10, $yPos)
$grpAction.Size = New-Object System.Drawing.Size(480, 90)
$form.Controls.Add($grpAction)

$radioCopyOnly = New-Object System.Windows.Forms.RadioButton
$radioCopyOnly.Text = "Copy Only"
$radioCopyOnly.Location = New-Object System.Drawing.Point(10, 22)
$radioCopyOnly.Size = New-Object System.Drawing.Size(150, 25)
$radioCopyOnly.Checked = $true
$grpAction.Controls.Add($radioCopyOnly)

$radioCopyInstall = New-Object System.Windows.Forms.RadioButton
$radioCopyInstall.Text = "Copy + Silent Install"
$radioCopyInstall.Location = New-Object System.Drawing.Point(200, 22)
$radioCopyInstall.Size = New-Object System.Drawing.Size(200, 25)
$grpAction.Controls.Add($radioCopyInstall)

$radioInstallPreCopied = New-Object System.Windows.Forms.RadioButton
$radioInstallPreCopied.Text = "Install Pre-Copied"
$radioInstallPreCopied.Location = New-Object System.Drawing.Point(10, 48)
$radioInstallPreCopied.Size = New-Object System.Drawing.Size(150, 25)
$grpAction.Controls.Add($radioInstallPreCopied)
$yPos += 62

# ── Remote Path (for pre-copied drivers) ────────────────────────────────────────

$labelRemotePath = New-Object System.Windows.Forms.Label
$labelRemotePath.Text = "Remote Driver Path (for Pre-Copied mode):"
$labelRemotePath.Location = New-Object System.Drawing.Point(10, $yPos)
$labelRemotePath.Size = New-Object System.Drawing.Size(480, 20)
$form.Controls.Add($labelRemotePath)
$yPos += 22

$textBoxRemotePath = New-Object System.Windows.Forms.TextBox
$textBoxRemotePath.Location = New-Object System.Drawing.Point(10, $yPos)
$textBoxRemotePath.Size = New-Object System.Drawing.Size(480, 22)
$textBoxRemotePath.Font = New-Object System.Drawing.Font("Consolas", 9)
$textBoxRemotePath.Text = "C:\temp\"
$textBoxRemotePath.ForeColor = [System.Drawing.Color]::Gray
$textBoxRemotePath.Add_GotFocus({
    if ($textBoxRemotePath.ForeColor -eq [System.Drawing.Color]::Gray) {
        $textBoxRemotePath.Text = ""
        $textBoxRemotePath.ForeColor = [System.Drawing.Color]::Black
    }
})
$textBoxRemotePath.Add_LostFocus({
    if ($textBoxRemotePath.Text -eq "") {
        $textBoxRemotePath.Text = "C:\temp\"
        $textBoxRemotePath.ForeColor = [System.Drawing.Color]::Gray
    }
})
$form.Controls.Add($textBoxRemotePath)
$yPos += 30

# ── BIOS Password (GroupBox) ──────────────────────────────────────────────────

$grpBios = New-Object System.Windows.Forms.GroupBox
$grpBios.Text = "BIOS Password (for BIOS/Firmware DUPs)"
$grpBios.Location = New-Object System.Drawing.Point(10, $yPos)
$grpBios.Size = New-Object System.Drawing.Size(480, 75)
$form.Controls.Add($grpBios)

$labelBios = New-Object System.Windows.Forms.Label
$labelBios.Text = "Password:"
$labelBios.Location = New-Object System.Drawing.Point(10, 28)
$labelBios.Size = New-Object System.Drawing.Size(70, 20)
$grpBios.Controls.Add($labelBios)

$textBoxBiosPassword = New-Object System.Windows.Forms.TextBox
$textBoxBiosPassword.Location = New-Object System.Drawing.Point(85, 25)
$textBoxBiosPassword.Size = New-Object System.Drawing.Size(240, 22)
$textBoxBiosPassword.PasswordChar = '*'
$textBoxBiosPassword.Font = New-Object System.Drawing.Font("Consolas", 10)
$grpBios.Controls.Add($textBoxBiosPassword)

$chkShowPassword = New-Object System.Windows.Forms.CheckBox
$chkShowPassword.Text = "Show"
$chkShowPassword.Location = New-Object System.Drawing.Point(335, 26)
$chkShowPassword.Size = New-Object System.Drawing.Size(90, 22)
$chkShowPassword.Add_CheckedChanged({
        if ($chkShowPassword.Checked) {
            $textBoxBiosPassword.PasswordChar = [char]0
        }
        else {
            $textBoxBiosPassword.PasswordChar = '*'
        }
    })
$grpBios.Controls.Add($chkShowPassword)

$labelBiosHint = New-Object System.Windows.Forms.Label
$labelBiosHint.Text = "Leave blank if no BIOS password is set on the target PCs."
$labelBiosHint.Location = New-Object System.Drawing.Point(10, 52)
$labelBiosHint.Size = New-Object System.Drawing.Size(460, 18)
$labelBiosHint.ForeColor = [System.Drawing.Color]::Gray
$labelBiosHint.Font = New-Object System.Drawing.Font("Segoe UI", 8)
$grpBios.Controls.Add($labelBiosHint)
$yPos += 82

# ── Post-Install Options (GroupBox) ────────────────────────────────────────────

$grpPost = New-Object System.Windows.Forms.GroupBox
$grpPost.Text = "Post-Install Options"
$grpPost.Location = New-Object System.Drawing.Point(10, $yPos)
$grpPost.Size = New-Object System.Drawing.Size(480, 55)
$form.Controls.Add($grpPost)

$chkReboot = New-Object System.Windows.Forms.CheckBox
$chkReboot.Text = "Auto-reboot after install (if required by firmware)"
$chkReboot.Location = New-Object System.Drawing.Point(10, 22)
$chkReboot.Size = New-Object System.Drawing.Size(450, 25)
$grpPost.Controls.Add($chkReboot)
$yPos += 62

# ── Buttons ────────────────────────────────────────────────────────────────────

$okButton = New-Object System.Windows.Forms.Button
$okButton.Text = "Deploy"
$okButton.Size = New-Object System.Drawing.Size(100, 32)
$okButton.Location = New-Object System.Drawing.Point(150, ($yPos + 5))
$okButton.Font = New-Object System.Drawing.Font("Segoe UI", 9, [System.Drawing.FontStyle]::Bold)
$okButton.Add_Click({ $form.Tag = "OK"; $form.Close() })
$form.Controls.Add($okButton)

$cancelButton = New-Object System.Windows.Forms.Button
$cancelButton.Text = "Cancel"
$cancelButton.Size = New-Object System.Drawing.Size(100, 32)
$cancelButton.Location = New-Object System.Drawing.Point(260, ($yPos + 5))
$cancelButton.Add_Click({ $form.Tag = "Cancel"; $form.Close() })
$form.Controls.Add($cancelButton)

$form.AcceptButton = $okButton
$form.CancelButton = $cancelButton

# ── Show Form ──────────────────────────────────────────────────────────────────

[void]$form.ShowDialog()

if ($form.Tag -ne "OK") {
    Write-Host "Cancelled by user." -ForegroundColor Yellow
    $form.Dispose()
    return
}

# ── Collect ALL GUI State BEFORE Dispose (v4+ fix) ────────────────────────────

$doInstall = $radioCopyInstall.Checked
$doInstallPreCopied = $radioInstallPreCopied.Checked
$doReboot = $chkReboot.Checked
$biosPassword = $textBoxBiosPassword.Text.Trim()
$hostInput = $textBoxHostName.Text
$folderPath = $textBoxFolder.Text.Trim()
$remotePath = $textBoxRemotePath.Text.Trim()
if ($remotePath -eq "C:\temp\") { $remotePath = "" }

# Collect checked .exe items and their full paths
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

$form.Dispose()

# ── Validate ───────────────────────────────────────────────────────────────────

$hostNames = $hostInput -split "`r`n" | ForEach-Object { $_.Trim() } | Where-Object { $_ -match '\S' }

if ($hostNames.Count -eq 0) {
    Write-Warning "No host names entered. Exiting."
    return
}

if ($selectedExes.Count -eq 0) {
    Write-Warning "No .exe files selected. Exiting."
    return
}

if ($doInstallPreCopied -and -not $remotePath) {
    Write-Warning "Pre-Copied mode selected but no remote path provided. Exiting."
    return
}

# Pre-check: verify every source file still exists
foreach ($exe in $selectedExes) {
    if (!(Test-Path -Path $exe.FullPath -ErrorAction SilentlyContinue)) {
        Write-Error "Source file not found: $($exe.FullPath)"
        return
    }
}

# ── Dell DUP Exit Codes Reference ─────────────────────────────────────────────
#  -1 = CANCELLED / DCU TIMEOUT
#   0 = SUCCESS  - Update completed, no reboot required
#   1 = FAILURE  - General error
#   2 = SUCCESS  - Update completed, reboot required to finish
#   3 = SOFT DEPENDENCY ERROR  - Same version already installed or downgrade
#   4 = HARD DEPENDENCY ERROR  - Prerequisites not met
#   5 = QUALIFICATION ERROR    - Not applicable to target system
#   6 = REBOOTING SYSTEM       - System is rebooting now
#   7 = PASSWORD ERROR         - BIOS password not provided or incorrect
#   8 = DOWNGRADE NOT ALLOWED  - Downgrade blocked by policy
#   9 = RPM VERIFY FAILED      - Linux only
#  10 = UNSPECIFIED ERROR       - Battery / EC / other HW failure
# ───────────────────────────────────────────────────────────────────────────────

$exitCodeMap = @{
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

# ── Deploy ─────────────────────────────────────────────────────────────────────

$successCount = 0
$failCount = 0
$results = @()
$pkgCount = $selectedExes.Count
$totalTasks = $hostNames.Count * $pkgCount
$taskNum = 0
$modeLabel = $(if ($doInstallPreCopied) { "Install Pre-Copied" } elseif ($doInstall) { "Copy + Silent Install" } else { "Copy Only" })
$stopwatch = [System.Diagnostics.Stopwatch]::StartNew()

Write-Host "`n========================================" -ForegroundColor Cyan
Write-Host "  Setup File Deployment v6" -ForegroundColor Cyan
Write-Host "========================================" -ForegroundColor Cyan
Write-Host "  Source : $folderPath" -ForegroundColor Cyan
Write-Host "  Mode   : $modeLabel" -ForegroundColor Cyan
if ($doInstallPreCopied) {
    Write-Host "  Remote : $remotePath" -ForegroundColor Cyan
}
Write-Host "  Hosts  : $($hostNames.Count) | Packages: $pkgCount" -ForegroundColor Cyan
Write-Host "  Total  : $totalTasks task(s)" -ForegroundColor Cyan
if (($doInstall -or $doInstallPreCopied) -and $biosPassword) {
    Write-Host "  BIOS   : Password provided" -ForegroundColor DarkGray
}
if (($doInstall -or $doInstallPreCopied) -and $doReboot) {
    Write-Host "  Reboot : Enabled (applied to last package per host)" -ForegroundColor Magenta
}
Write-Host "========================================" -ForegroundColor Cyan
Write-Host "  Selected packages:" -ForegroundColor DarkGray
foreach ($exe in $selectedExes) {
    Write-Host "    - $($exe.FileName)" -ForegroundColor DarkGray
}
Write-Host "========================================`n" -ForegroundColor Cyan

# PSSession options: 30 s connect timeout, 10 min operation timeout
$sessionOpts = New-PSSessionOption -OpenTimeout 30000 -OperationTimeout 600000

foreach ($name in $hostNames) {
    Write-Host "--- Processing: $name ---" -ForegroundColor Cyan

    # Connectivity check
    if (!(Test-Connection -ComputerName $name -Count 1 -Quiet)) {
        Write-Warning "  $name is offline or unreachable."
        $failCount += $pkgCount
        $taskNum += $pkgCount
        foreach ($exe in $selectedExes) {
            $results += [PSCustomObject]@{ Host = $name; Package = $exe.FileName; Step = "Ping"; Status = "OFFLINE" }
        }
        continue
    }

    # WinRM pre-check (only relevant when install mode is selected)
    $winrmOk = $true
    if ($doInstall) {
        $winrmTest = Test-WSMan -ComputerName $name -ErrorAction SilentlyContinue
        if (-not $winrmTest) {
            Write-Warning "  $name - WinRM not available. Files will be copied but install will be skipped."
            $winrmOk = $false
        }
    }

    # Ensure destination folder exists (skip for pre-copied mode)
    if (-not $doInstallPreCopied) {
        $destinationPath = "\\$name\C$\temp\"
        try {
            if (!(Test-Path -Path $destinationPath -ErrorAction Stop)) {
                New-Item -ItemType Directory -Path $destinationPath -Force -ErrorAction Stop | Out-Null
                Write-Host "  Created $destinationPath" -ForegroundColor DarkGray
            }
        }
        catch {
            Write-Error "  Cannot create $destinationPath - $($_.Exception.Message)"
            $failCount += $pkgCount
            $taskNum += $pkgCount
            foreach ($exe in $selectedExes) {
                $results += [PSCustomObject]@{ Host = $name; Package = $exe.FileName; Step = "Mkdir"; Status = "FAILED" }
            }
            continue
        }
    }

    # Process each selected .exe
    for ($pkgIdx = 0; $pkgIdx -lt $pkgCount; $pkgIdx++) {
        $exe = $selectedExes[$pkgIdx]
        $isLastPkg = ($pkgIdx -eq ($pkgCount - 1))
        $taskNum++

        # ── Step 1: Copy (skip for pre-copied mode) ──────────────────────────
        if (-not $doInstallPreCopied) {
            try {
                Write-Host "  [$taskNum/$totalTasks][$($exe.FileName)] Copying..." -NoNewline
                Copy-Item -Path $exe.FullPath -Destination $destinationPath -Force -ErrorAction Stop
                Write-Host " OK" -ForegroundColor Green
            }
            catch {
                Write-Host " FAILED" -ForegroundColor Red
                Write-Error "    Copy error: $($_.Exception.Message)"
                $failCount++
                $results += [PSCustomObject]@{ Host = $name; Package = $exe.FileName; Step = "Copy"; Status = "FAILED" }
                continue
            }
        }

        # ── Step 2: Remote Silent Install ──────────────────────────────────────
        if ($doInstall -or $doInstallPreCopied) {
            if (-not $winrmOk) {
                Write-Warning "  [$taskNum/$totalTasks][$($exe.FileName)] Skipped install (WinRM unavailable)."
                $failCount++
                $results += [PSCustomObject]@{ Host = $name; Package = $exe.FileName; Step = "Install"; Status = "SKIPPED (WinRM unavailable)" }
                continue
            }

            try {
                if ($doInstallPreCopied) {
                    $remoteExe = Join-Path $remotePath $exe.FileName
                    $remoteLog = Join-Path $remotePath $exe.LogName
                }
                else {
                    $remoteExe = "C:\temp\$($exe.FileName)"
                    $remoteLog = "C:\temp\$($exe.LogName)"
                }

                # Build Dell DUP silent install arguments
                $installArgs = "/s /l=$remoteLog"
                if ($biosPassword) { $installArgs += " /p=$biosPassword" }
                # Only append /r to the LAST package to avoid mid-deploy reboots
                if ($doReboot -and $isLastPkg) { $installArgs += " /r" }

                Write-Host "  [$taskNum/$totalTasks][$($exe.FileName)] Installing on $name..." -NoNewline

                $exitCode = Invoke-Command -ComputerName $name -SessionOption $sessionOpts -ScriptBlock {
                    param($exePath, $arguments)
                    $process = Start-Process -FilePath $exePath -ArgumentList $arguments -Wait -PassThru -NoNewWindow
                    return $process.ExitCode
                } -ArgumentList $remoteExe, $installArgs -ErrorAction Stop

                # Null-safety guard
                if ($null -eq $exitCode) {
                    Write-Host " NO EXIT CODE (session may have dropped)" -ForegroundColor Red
                    $failCount++
                    $results += [PSCustomObject]@{ Host = $name; Package = $exe.FileName; Step = "Install"; Status = "NO EXIT CODE (session dropped?)" }
                    continue
                }

                # Interpret exit code
                $exitDesc = $(if ($exitCodeMap.ContainsKey([int]$exitCode)) {
                        $exitCodeMap[[int]$exitCode]
                    }
                    elseif ($exitCode -lt 0) {
                        $hexCode = "0x{0:X8}" -f ([uint32]$exitCode)
                        "APPLICATION CRASH (Exit Code: $exitCode / $hexCode)"
                    }
                    else {
                        "UNKNOWN (Exit Code: $exitCode)"
                    })

                if ($exitCode -eq 0 -or $exitCode -eq 2 -or $exitCode -eq 6) {
                    $color = $(if ($exitCode -eq 0) { "Green" } else { "Yellow" })
                    Write-Host " $exitDesc" -ForegroundColor $color
                    $successCount++
                    $results += [PSCustomObject]@{ Host = $name; Package = $exe.FileName; Step = "Install"; Status = $exitDesc }
                }
                else {
                    Write-Host " $exitDesc" -ForegroundColor Red
                    $failCount++
                    $results += [PSCustomObject]@{ Host = $name; Package = $exe.FileName; Step = "Install"; Status = $exitDesc }
                }
            }
            catch {
                Write-Host " FAILED" -ForegroundColor Red
                Write-Error "    Remote install error: $($_.Exception.Message)"
                $failCount++
                $results += [PSCustomObject]@{ Host = $name; Package = $exe.FileName; Step = "Install"; Status = "ERROR: $($_.Exception.Message)" }
            }
        }
        else {
            # Copy-only mode -> mark as success
            $successCount++
            $results += [PSCustomObject]@{ Host = $name; Package = $exe.FileName; Step = "Copy"; Status = "COPIED" }
        }
    }
}

$stopwatch.Stop()
$elapsed = $stopwatch.Elapsed.ToString("hh\:mm\:ss")

# ── Summary ────────────────────────────────────────────────────────────────────

Write-Host "`n========================================" -ForegroundColor Yellow
Write-Host "  Deployment Summary" -ForegroundColor Yellow
Write-Host "========================================" -ForegroundColor Yellow
Write-Host "  Source  : $folderPath"
Write-Host "  Mode    : $modeLabel"
Write-Host "  Success : $successCount" -ForegroundColor Green
Write-Host "  Failed  : $failCount" -ForegroundColor $(if ($failCount -gt 0) { "Red" } else { "Green" })
Write-Host "  Total   : $totalTasks" -ForegroundColor Yellow
Write-Host "  Elapsed : $elapsed" -ForegroundColor DarkGray
Write-Host "========================================`n" -ForegroundColor Yellow

if ($results.Count -gt 0) {
    Write-Host "Detailed Results:" -ForegroundColor Cyan
    $results | Format-Table Host, Package, Step, Status -AutoSize
}

# ── CSV Export ─────────────────────────────────────────────────────────────────

if ($results.Count -gt 0) {
    $timestamp = Get-Date -Format "yyyyMMdd_HHmmss"
    $csvPath = "C:\temp\DeploymentLog_$timestamp.csv"
    try {
        $results | Export-Csv -Path $csvPath -NoTypeInformation -Encoding UTF8
        Write-Host "Results exported to $csvPath" -ForegroundColor Green
    }
    catch {
        Write-Warning "Could not export CSV: $($_.Exception.Message)"
    }
}

if ($doInstall) {
    Write-Host "Tip: Check remote logs at C:\temp\<LogName> on each host for details." -ForegroundColor DarkGray
}