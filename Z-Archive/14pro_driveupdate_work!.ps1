Add-Type -AssemblyName System.Windows.Forms
Add-Type -AssemblyName System.Drawing

# ══════════════════════════════════════════════════════════════════════════════
#  Setup File Deployment v3 – Copy + Optional Remote Silent Install
#  Supports Dell Update Packages (DUP) with /s /l /r /p switches
#  Changelog v3:
#    - Added BIOS password GUI field with show/hide toggle
#    - Expanded Dell DUP exit code map (codes -1 through 10)
#    - Crash detection for non-DUP exit codes (hex display)
#    - WinRM pre-check before remote install
#    - Task progress counter during deployment
#    - CSV export of results
#    - PS 5.1 compatibility fixes
# ══════════════════════════════════════════════════════════════════════════════

# ── GUI Setup ──────────────────────────────────────────────────────────────────

$form = New-Object System.Windows.Forms.Form
$form.Text = "Setup File Deployment v3"
$form.Size = New-Object System.Drawing.Size(440, 700)
$form.StartPosition = "CenterScreen"
$form.FormBorderStyle = "FixedDialog"
$form.MaximizeBox = $false

# ── Host Name Input ────────────────────────────────────────────────────────────

$labelHostName = New-Object System.Windows.Forms.Label
$labelHostName.Text = "Enter Host Names (one per line):"
$labelHostName.Location = New-Object System.Drawing.Point(10, 10)
$labelHostName.Size = New-Object System.Drawing.Size(400, 20)
$form.Controls.Add($labelHostName)

$textBoxHostName = New-Object System.Windows.Forms.TextBox
$textBoxHostName.Location = New-Object System.Drawing.Point(10, 35)
$textBoxHostName.Size = New-Object System.Drawing.Size(400, 130)
$textBoxHostName.Multiline = $true
$textBoxHostName.ScrollBars = "Vertical"
$textBoxHostName.Font = New-Object System.Drawing.Font("Consolas", 10)
$form.Controls.Add($textBoxHostName)

# ── Deployment Options (GroupBox) ──────────────────────────────────────────────
# To add a new package: 1) create a checkbox   2) add an entry in $deployments

$grpDeploy = New-Object System.Windows.Forms.GroupBox
$grpDeploy.Text = "Deployment Options"
$grpDeploy.Location = New-Object System.Drawing.Point(10, 175)
$grpDeploy.Size = New-Object System.Drawing.Size(400, 105)
$form.Controls.Add($grpDeploy)

$checkBox1 = New-Object System.Windows.Forms.CheckBox
$checkBox1.Text = "SanDisk SN8000s Firmware Update"
$checkBox1.Location = New-Object System.Drawing.Point(10, 20)
$checkBox1.Size = New-Object System.Drawing.Size(370, 20)
$grpDeploy.Controls.Add($checkBox1)

$checkBox2 = New-Object System.Windows.Forms.CheckBox
$checkBox2.Text = "Dell Pro BIOS Driver v2.12.4"
$checkBox2.Location = New-Object System.Drawing.Point(10, 45)
$checkBox2.Size = New-Object System.Drawing.Size(370, 20)
$grpDeploy.Controls.Add($checkBox2)

$checkBox3 = New-Object System.Windows.Forms.CheckBox
$checkBox3.Text = "Intel Arc Graphics Driver v32.0.101.8508"
$checkBox3.Location = New-Object System.Drawing.Point(10, 70)
$checkBox3.Size = New-Object System.Drawing.Size(370, 20)
$grpDeploy.Controls.Add($checkBox3)

$checkBox4 = New-Object System.Windows.Forms.CheckBox
$checkBox4.Text = "Dell_Pro_Max_MA14250_MA16250_1.9.0"
$checkBox4.Location = New-Object System.Drawing.Point(10, 95)
$checkBox4.Size = New-Object System.Drawing.Size(370, 20)
$grpDeploy.Controls.Add($checkBox4)
# ── Action Mode (GroupBox) ─────────────────────────────────────────────────────

$grpAction = New-Object System.Windows.Forms.GroupBox
$grpAction.Text = "Action"
$grpAction.Location = New-Object System.Drawing.Point(10, 290)
$grpAction.Size = New-Object System.Drawing.Size(400, 60)
$form.Controls.Add($grpAction)

$radioCopyOnly = New-Object System.Windows.Forms.RadioButton
$radioCopyOnly.Text = "Copy Only"
$radioCopyOnly.Location = New-Object System.Drawing.Point(10, 25)
$radioCopyOnly.Size = New-Object System.Drawing.Size(150, 25)
$radioCopyOnly.Checked = $true
$grpAction.Controls.Add($radioCopyOnly)

$radioCopyInstall = New-Object System.Windows.Forms.RadioButton
$radioCopyInstall.Text = "Copy + Silent Install"
$radioCopyInstall.Location = New-Object System.Drawing.Point(180, 25)
$radioCopyInstall.Size = New-Object System.Drawing.Size(200, 25)
$grpAction.Controls.Add($radioCopyInstall)

# ── BIOS Password (GroupBox) ──────────────────────────────────────────────────

$grpBios = New-Object System.Windows.Forms.GroupBox
$grpBios.Text = "BIOS Password (for BIOS/Firmware DUPs)"
$grpBios.Location = New-Object System.Drawing.Point(10, 360)
$grpBios.Size = New-Object System.Drawing.Size(400, 80)
$form.Controls.Add($grpBios)

$labelBios = New-Object System.Windows.Forms.Label
$labelBios.Text = "Password:"
$labelBios.Location = New-Object System.Drawing.Point(10, 30)
$labelBios.Size = New-Object System.Drawing.Size(70, 20)
$grpBios.Controls.Add($labelBios)

$textBoxBiosPassword = New-Object System.Windows.Forms.TextBox
$textBoxBiosPassword.Location = New-Object System.Drawing.Point(85, 27)
$textBoxBiosPassword.Size = New-Object System.Drawing.Size(200, 22)
$textBoxBiosPassword.PasswordChar = '*'
$textBoxBiosPassword.Font = New-Object System.Drawing.Font("Consolas", 10)
$grpBios.Controls.Add($textBoxBiosPassword)

$chkShowPassword = New-Object System.Windows.Forms.CheckBox
$chkShowPassword.Text = "Show"
$chkShowPassword.Location = New-Object System.Drawing.Point(295, 27)
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
$labelBiosHint.Location = New-Object System.Drawing.Point(10, 55)
$labelBiosHint.Size = New-Object System.Drawing.Size(380, 18)
$labelBiosHint.ForeColor = [System.Drawing.Color]::Gray
$labelBiosHint.Font = New-Object System.Drawing.Font("Segoe UI", 8)
$grpBios.Controls.Add($labelBiosHint)

# ── Post-Install Options (GroupBox) ────────────────────────────────────────────

$grpPost = New-Object System.Windows.Forms.GroupBox
$grpPost.Text = "Post-Install Options"
$grpPost.Location = New-Object System.Drawing.Point(10, 450)
$grpPost.Size = New-Object System.Drawing.Size(400, 60)
$form.Controls.Add($grpPost)

$chkReboot = New-Object System.Windows.Forms.CheckBox
$chkReboot.Text = "Auto-reboot after install (if required by firmware)"
$chkReboot.Location = New-Object System.Drawing.Point(10, 25)
$chkReboot.Size = New-Object System.Drawing.Size(370, 25)
$grpPost.Controls.Add($chkReboot)

# ── Buttons ────────────────────────────────────────────────────────────────────

$okButton = New-Object System.Windows.Forms.Button
$okButton.Text = "OK"
$okButton.Size = New-Object System.Drawing.Size(90, 30)
$okButton.Location = New-Object System.Drawing.Point(130, 620)
$okButton.Add_Click({ $form.Tag = "OK"; $form.Close() })
$form.Controls.Add($okButton)

$cancelButton = New-Object System.Windows.Forms.Button
$cancelButton.Text = "Cancel"
$cancelButton.Size = New-Object System.Drawing.Size(90, 30)
$cancelButton.Location = New-Object System.Drawing.Point(230, 620)
$cancelButton.Add_Click({ $form.Tag = "Cancel"; $form.Close() })
$form.Controls.Add($cancelButton)

# Keyboard shortcuts: Enter -> OK, Escape -> Cancel
$form.AcceptButton = $okButton
$form.CancelButton = $cancelButton

# ── Show Form ──────────────────────────────────────────────────────────────────

[void]$form.ShowDialog()

if ($form.Tag -ne "OK") {
    Write-Host "Cancelled by user." -ForegroundColor Yellow
    $form.Dispose()
    return
}

# ── Collect User Selections ────────────────────────────────────────────────────

$doInstall = $radioCopyInstall.Checked
$doReboot = $chkReboot.Checked
$biosPassword = $textBoxBiosPassword.Text.Trim()

# Dispose the form now that we have all values
$form.Dispose()

# ── Build Deployment Map ───────────────────────────────────────────────────────
# Each entry:  Label -> @{ Source = UNC path; FileName = exe name; LogName = log }

$deployments = [ordered]@{}

if ($checkBox1.Checked) {
    $deployments["SanDisk SN8000s FW"] = @{
        Source   = "\\corp\hk\Transition\DAT18\Kits\Software\Engineering\sharePC_driver\SanDisk-SN8000s-SED-Solid-State-Drive-Firmware-Update_0NWH2_WIN64_6311.2104_A04.exe"
        FileName = "SanDisk-SN8000s-SED-Solid-State-Drive-Firmware-Update_0NWH2_WIN64_6311.2104_A04.exe"
        LogName  = "SanDisk_FW_Update.log"
    }
}

if ($checkBox2.Checked) {
    $deployments["Dell Pro BIOS Driver"] = @{
        Source   = "\\corp\hk\Transition\DAT18\Kits\Software\Engineering\Dell\Dell_Pro_14_Premium_Driver\Dell_Pro_PA13250_PA14250_PB13250_PB14250_PB16250_2.12.4.exe"
        FileName = "Dell_Pro_PA13250_PA14250_PB13250_PB14250_PB16250_2.12.4.exe"
        LogName  = "Dell_Pro_BIOS_Driver.log"
    }
}

if ($checkBox3.Checked) {
    $deployments["Intel Arc Graphics Driver"] = @{
        Source   = "\\corp\hk\Transition\DAT18\Kits\Software\Engineering\Dell\Dell_Pro_14_Premium_Driver\Intel-Arc-Graphics-Driver_CH4KW_WIN64_32.0.101.8508_A08.exe"
        FileName = "Intel-Arc-Graphics-Driver_CH4KW_WIN64_32.0.101.8508_A08.exe"
        LogName  = "Intel_Arc_Graphics.log"
    }
}

if ($checkBox4.Checked) {
    $deployments["Intel Arc Graphics Driver"] = @{
        Source   = "C:\Temp2\14_Premium_Driver\Dell_Pro_Max_MA14250_MA16250_1.9.0.exe"
        FileName = "Dell_Pro_Max_MA14250_MA16250_1.9.0.exe"
        LogName  = "Dell_Pro_Max_MA14250_MA16250_1.9.0.log"
    }
}

# ── Validate ───────────────────────────────────────────────────────────────────

$hostNames = $textBoxHostName.Text -split "`r`n" | ForEach-Object { $_.Trim() } | Where-Object { $_ -match '\S' }

if ($hostNames.Count -eq 0) {
    Write-Warning "No host names entered. Exiting."
    return
}

if ($deployments.Count -eq 0) {
    Write-Warning "No deployment options selected. Exiting."
    return
}

# Pre-check: verify every source file exists before we start
foreach ($label in $deployments.Keys) {
    $src = $deployments[$label].Source
    if (!(Test-Path -Path $src)) {
        Write-Error "Source file not found for [$label]: $src"
        return
    }
}

# ── Dell DUP Exit Codes Reference (Complete) ──────────────────────────────────
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
$totalTasks = $hostNames.Count * $deployments.Count
$taskNum = 0
$modeLabel = $(if ($doInstall) { "Copy + Silent Install" } else { "Copy Only" })

Write-Host "`n========================================" -ForegroundColor Cyan
Write-Host "  Mode  : $modeLabel" -ForegroundColor Cyan
Write-Host "  Hosts : $($hostNames.Count) | Packages: $($deployments.Count)" -ForegroundColor Cyan
Write-Host "  Total : $totalTasks task(s)" -ForegroundColor Cyan
if ($doInstall -and $biosPassword) {
    Write-Host "  BIOS  : Password provided" -ForegroundColor DarkGray
}
if ($doInstall -and $doReboot) {
    Write-Host "  Reboot: Enabled (if required)" -ForegroundColor Magenta
}
Write-Host "========================================`n" -ForegroundColor Cyan

foreach ($name in $hostNames) {
    Write-Host "--- Processing: $name ---" -ForegroundColor Cyan

    # Connectivity check
    if (!(Test-Connection -ComputerName $name -Count 1 -Quiet)) {
        Write-Warning "  $name is offline or unreachable."
        $failCount += $deployments.Count
        $taskNum += $deployments.Count
        foreach ($label in $deployments.Keys) {
            $results += [PSCustomObject]@{ Host = $name; Package = $label; Step = "Ping"; Status = "OFFLINE" }
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

    # Ensure destination folder exists
    $destinationPath = "\\$name\C$\temp\"
    try {
        if (!(Test-Path -Path $destinationPath)) {
            New-Item -ItemType Directory -Path $destinationPath -Force | Out-Null
            Write-Host "  Created $destinationPath" -ForegroundColor DarkGray
        }
    }
    catch {
        Write-Error "  Cannot create $destinationPath - $($_.Exception.Message)"
        $failCount += $deployments.Count
        $taskNum += $deployments.Count
        foreach ($label in $deployments.Keys) {
            $results += [PSCustomObject]@{ Host = $name; Package = $label; Step = "Mkdir"; Status = "FAILED" }
        }
        continue
    }

    # Process each deployment package
    foreach ($label in $deployments.Keys) {
        $pkg = $deployments[$label]
        $taskNum++

        # ── Step 1: Copy ──────────────────────────────────────────────────
        try {
            Write-Host "  [$taskNum/$totalTasks][$label] Copying..." -NoNewline
            Copy-Item -Path $pkg.Source -Destination $destinationPath -Force
            Write-Host " OK" -ForegroundColor Green
        }
        catch {
            Write-Host " FAILED" -ForegroundColor Red
            Write-Error "    Copy error: $($_.Exception.Message)"
            $failCount++
            $results += [PSCustomObject]@{ Host = $name; Package = $label; Step = "Copy"; Status = "FAILED" }
            continue
        }

        # ── Step 2: Remote Silent Install (if selected) ───────────────────
        if ($doInstall) {
            # Skip install if WinRM is unavailable
            if (-not $winrmOk) {
                Write-Warning "  [$taskNum/$totalTasks][$label] Skipped install (WinRM unavailable)."
                $failCount++
                $results += [PSCustomObject]@{ Host = $name; Package = $label; Step = "Install"; Status = "SKIPPED (WinRM unavailable)" }
                continue
            }

            try {
                $remoteExe = "C:\temp\$($pkg.FileName)"
                $remoteLog = "C:\temp\$($pkg.LogName)"

                # Build Dell DUP silent install arguments
                $installArgs = "/s /l=$remoteLog"
                if ($biosPassword) { $installArgs += " /p=$biosPassword" }
                if ($doReboot) { $installArgs += " /r" }

                Write-Host "  [$taskNum/$totalTasks][$label] Installing silently on $name..." -NoNewline

                # Execute the installer remotely via WinRM (Invoke-Command)
                $exitCode = Invoke-Command -ComputerName $name -ScriptBlock {
                    param($exe, $arguments)
                    $process = Start-Process -FilePath $exe -ArgumentList $arguments -Wait -PassThru -NoNewWindow
                    return $process.ExitCode
                } -ArgumentList $remoteExe, $installArgs -ErrorAction Stop

                # Interpret the exit code
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
                    $results += [PSCustomObject]@{ Host = $name; Package = $label; Step = "Install"; Status = $exitDesc }
                }
                else {
                    Write-Host " $exitDesc" -ForegroundColor Red
                    $failCount++
                    $results += [PSCustomObject]@{ Host = $name; Package = $label; Step = "Install"; Status = $exitDesc }
                }
            }
            catch {
                Write-Host " FAILED" -ForegroundColor Red
                Write-Error "    Remote install error: $($_.Exception.Message)"
                $failCount++
                $results += [PSCustomObject]@{ Host = $name; Package = $label; Step = "Install"; Status = "ERROR: $($_.Exception.Message)" }
            }
        }
        else {
            # Copy-only mode -> mark as success
            $successCount++
            $results += [PSCustomObject]@{ Host = $name; Package = $label; Step = "Copy"; Status = "COPIED" }
        }
    }
}

# ── Summary ────────────────────────────────────────────────────────────────────

Write-Host "`n========================================" -ForegroundColor Yellow
Write-Host "  Deployment Summary" -ForegroundColor Yellow
Write-Host "========================================" -ForegroundColor Yellow
Write-Host "  Mode    : $modeLabel"
Write-Host "  Success : $successCount" -ForegroundColor Green
Write-Host "  Failed  : $failCount" -ForegroundColor $(if ($failCount -gt 0) { "Red" } else { "Green" })
Write-Host "  Total   : $totalTasks" -ForegroundColor Yellow
Write-Host "========================================`n" -ForegroundColor Yellow

# Detailed results table
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

# Tip for checking remote logs
if ($doInstall) {
    Write-Host "Tip: Check remote logs at C:\temp\<LogName> on each host for details." -ForegroundColor DarkGray
}
