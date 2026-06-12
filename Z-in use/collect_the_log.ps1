Add-Type -AssemblyName System.Windows.Forms
Add-Type -AssemblyName System.Drawing

# ══════════════════════════════════════════════════════════════════════════════
#  Setup File Deployment v3 – Copy + Silent Install + Log Collection
#  Supports Dell Update Packages (DUP) with /s /l /r switches
# ══════════════════════════════════════════════════════════════════════════════

# ── GUI Setup ──────────────────────────────────────────────────────────────────

$form = New-Object System.Windows.Forms.Form
$form.Text = "Setup File Deployment v3"
$form.Size = New-Object System.Drawing.Size(440, 580)
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
$grpDeploy.Size = New-Object System.Drawing.Size(400, 60)
$form.Controls.Add($grpDeploy)

$checkBox1 = New-Object System.Windows.Forms.CheckBox
$checkBox1.Text = "SanDisk SN8000s Firmware Update"
$checkBox1.Location = New-Object System.Drawing.Point(10, 25)
$checkBox1.Size = New-Object System.Drawing.Size(370, 25)
$grpDeploy.Controls.Add($checkBox1)

# $checkBox2 = New-Object System.Windows.Forms.CheckBox
# $checkBox2.Text = "Example - Another Deployment"
# $checkBox2.Location = New-Object System.Drawing.Point(10, 50)
# $checkBox2.Size = New-Object System.Drawing.Size(370, 25)
# $grpDeploy.Controls.Add($checkBox2)

# ── Action Mode (GroupBox) ─────────────────────────────────────────────────────

$grpAction = New-Object System.Windows.Forms.GroupBox
$grpAction.Text = "Action"
$grpAction.Location = New-Object System.Drawing.Point(10, 245)
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

# ── Post-Install Options (GroupBox) ────────────────────────────────────────────

$grpPost = New-Object System.Windows.Forms.GroupBox
$grpPost.Text = "Post-Install Options"
$grpPost.Location = New-Object System.Drawing.Point(10, 315)
$grpPost.Size = New-Object System.Drawing.Size(400, 80)
$form.Controls.Add($grpPost)

$chkReboot = New-Object System.Windows.Forms.CheckBox
$chkReboot.Text = "Auto-reboot after install (if required by firmware)"
$chkReboot.Location = New-Object System.Drawing.Point(10, 25)
$chkReboot.Size = New-Object System.Drawing.Size(370, 25)
$grpPost.Controls.Add($chkReboot)

$chkCollectLogs = New-Object System.Windows.Forms.CheckBox
$chkCollectLogs.Text = "Collect remote logs after installation"
$chkCollectLogs.Location = New-Object System.Drawing.Point(10, 50)
$chkCollectLogs.Size = New-Object System.Drawing.Size(370, 25)
$chkCollectLogs.Checked = $true
$grpPost.Controls.Add($chkCollectLogs)

# ── Buttons ────────────────────────────────────────────────────────────────────

$okButton = New-Object System.Windows.Forms.Button
$okButton.Text = "OK"
$okButton.Size = New-Object System.Drawing.Size(90, 30)
$okButton.Location = New-Object System.Drawing.Point(130, 500)
$okButton.Add_Click({ $form.Tag = "OK"; $form.Close() })
$form.Controls.Add($okButton)

$cancelButton = New-Object System.Windows.Forms.Button
$cancelButton.Text = "Cancel"
$cancelButton.Size = New-Object System.Drawing.Size(90, 30)
$cancelButton.Location = New-Object System.Drawing.Point(230, 500)
$cancelButton.Add_Click({ $form.Tag = "Cancel"; $form.Close() })
$form.Controls.Add($cancelButton)

# Keyboard shortcuts: Enter -> OK, Escape -> Cancel
$form.AcceptButton = $okButton
$form.CancelButton = $cancelButton

# ── Show Form ──────────────────────────────────────────────────────────────────

[void]$form.ShowDialog()

if ($form.Tag -ne "OK") {
    Write-Host "Cancelled by user." -ForegroundColor Yellow
    return
}

# ── Collect User Selections ────────────────────────────────────────────────────

$doInstall = $radioCopyInstall.Checked
$doReboot = $chkReboot.Checked
$doCollectLog = $chkCollectLogs.Checked

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

# if ($checkBox2.Checked) {
#     $deployments["Another Tool"] = @{
#         Source   = "\\server\share\path\to\installer.exe"
#         FileName = "installer.exe"
#         LogName  = "AnotherTool_Install.log"
#     }
# }

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

# ── Dell DUP Exit Codes Reference ─────────────────────────────────────────────
#  0 = SUCCESS  - Update completed, no reboot required
#  1 = FAILURE  - General error
#  2 = SUCCESS  - Update completed, reboot required to finish
#  3 = SOFT DEPENDENCY ERROR  - Prerequisites not met
#  4 = HARD DEPENDENCY ERROR  - Hardware mismatch
#  5 = QUALIFICATION ERROR    - Not applicable to target system
#  6 = REBOOTING SYSTEM       - System is rebooting now
# ───────────────────────────────────────────────────────────────────────────────

$exitCodeMap = @{
    0 = "SUCCESS (no reboot needed)"
    1 = "FAILURE (general error)"
    2 = "SUCCESS (reboot required)"
    3 = "SOFT DEPENDENCY ERROR"
    4 = "HARD DEPENDENCY ERROR"
    5 = "QUALIFICATION ERROR (not applicable)"
    6 = "REBOOTING SYSTEM"
}

# ── Deploy ─────────────────────────────────────────────────────────────────────

$successCount = 0
$failCount = 0
$results = @()
$totalTasks = $hostNames.Count * $deployments.Count
$modeLabel = if ($doInstall) { "Copy + Silent Install" } else { "Copy Only" }

Write-Host "`n========================================" -ForegroundColor Cyan
Write-Host "  Mode  : $modeLabel" -ForegroundColor Cyan
Write-Host "  Hosts : $($hostNames.Count) | Packages: $($deployments.Count)" -ForegroundColor Cyan
Write-Host "  Total : $totalTasks task(s)" -ForegroundColor Cyan
if ($doInstall -and $doReboot) {
    Write-Host "  Reboot: Enabled (if required)" -ForegroundColor Magenta
}
if ($doInstall -and $doCollectLog) {
    Write-Host "  Logs  : Will collect after install" -ForegroundColor Magenta
}
Write-Host "========================================`n" -ForegroundColor Cyan

foreach ($name in $hostNames) {
    Write-Host "--- Processing: $name ---" -ForegroundColor Cyan

    # Connectivity check
    if (!(Test-Connection -ComputerName $name -Count 1 -Quiet)) {
        Write-Warning "  $name is offline or unreachable."
        $failCount += $deployments.Count
        foreach ($label in $deployments.Keys) {
            $results += [PSCustomObject]@{ Host = $name; Package = $label; Step = "Ping"; Status = "OFFLINE" }
        }
        continue
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
        foreach ($label in $deployments.Keys) {
            $results += [PSCustomObject]@{ Host = $name; Package = $label; Step = "Mkdir"; Status = "FAILED" }
        }
        continue
    }

    # Process each deployment package
    foreach ($label in $deployments.Keys) {
        $pkg = $deployments[$label]

        # ── Step 1: Copy ──────────────────────────────────────────────────
        try {
            Write-Host "  [$label] Copying..." -NoNewline
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
            try {
                $remoteExe = "C:\temp\$($pkg.FileName)"
                $remoteLog = "C:\temp\$($pkg.LogName)"

                # Build Dell DUP silent install arguments
                $installArgs = "/s /l=$remoteLog"
                if ($doReboot) { $installArgs += " /r" }

                Write-Host "  [$label] Installing silently on $name..." -NoNewline

                # Execute the installer remotely via WinRM (Invoke-Command)
                $exitCode = Invoke-Command -ComputerName $name -ScriptBlock {
                    param($exe, $arguments)
                    $process = Start-Process -FilePath $exe -ArgumentList $arguments -Wait -PassThru -NoNewWindow
                    return $process.ExitCode
                } -ArgumentList $remoteExe, $installArgs -ErrorAction Stop

                # Interpret the DUP exit code
                $exitDesc = if ($exitCodeMap.ContainsKey([int]$exitCode)) {
                    $exitCodeMap[[int]$exitCode]
                }
                else {
                    "UNKNOWN (Exit Code: $exitCode)"
                }

                if ($exitCode -eq 0 -or $exitCode -eq 2 -or $exitCode -eq 6) {
                    $color = if ($exitCode -eq 0) { "Green" } elseif ($exitCode -eq 2) { "Yellow" } else { "Yellow" }
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
            # Copy-only mode - mark as success
            $successCount++
            $results += [PSCustomObject]@{ Host = $name; Package = $label; Step = "Copy"; Status = "COPIED" }
        }
    }
}

# ── Collect Remote Logs ────────────────────────────────────────────────────────

if ($doInstall -and $doCollectLog) {

    $timestamp = Get-Date -Format "yyyy-MM-dd_HHmmss"
    $localLogFolder = "C:\temp\DeploymentLogs\$timestamp"
    New-Item -ItemType Directory -Path $localLogFolder -Force | Out-Null

    Write-Host "`n========================================" -ForegroundColor Magenta
    Write-Host "  Collecting Remote Logs" -ForegroundColor Magenta
    Write-Host "  Local folder: $localLogFolder" -ForegroundColor Magenta
    Write-Host "========================================`n" -ForegroundColor Magenta

    $logCollected = 0
    $logFailed = 0

    # Get unique hosts where an install was attempted (regardless of outcome)
    $installHosts = $results | Where-Object { $_.Step -eq "Install" } | Select-Object -ExpandProperty Host -Unique

    foreach ($name in $installHosts) {
        Write-Host "  Collecting from $name..." -NoNewline

        # Host may have rebooted - check connectivity first
        if (!(Test-Connection -ComputerName $name -Count 1 -Quiet)) {
            Write-Host " SKIPPED (offline / rebooting)" -ForegroundColor Yellow
            $logFailed++
            continue
        }

        $hostLogCollected = $false

        foreach ($label in $deployments.Keys) {
            $pkg = $deployments[$label]
            $remoteLogPath = "\\$name\C$\temp\$($pkg.LogName)"
            $localLogName = "${name}_$($pkg.LogName)"
            $localLogPath = Join-Path $localLogFolder $localLogName

            try {
                if (Test-Path -Path $remoteLogPath) {
                    Copy-Item -Path $remoteLogPath -Destination $localLogPath -Force
                    $hostLogCollected = $true
                }
                else {
                    Write-Host ""
                    Write-Warning "    Log not found: $remoteLogPath"
                }
            }
            catch {
                Write-Host ""
                Write-Warning "    Failed to collect from ${name}: $($_.Exception.Message)"
            }
        }

        if ($hostLogCollected) {
            Write-Host " OK" -ForegroundColor Green
            $logCollected++
        }
        else {
            $logFailed++
        }
    }

    Write-Host "`n  Logs collected : $logCollected" -ForegroundColor Green
    Write-Host "  Logs failed    : $logFailed" -ForegroundColor $(if ($logFailed -gt 0) { "Red" } else { "Green" })
    Write-Host "  Saved to       : $localLogFolder`n" -ForegroundColor Magenta
}

# ── Deployment Summary ─────────────────────────────────────────────────────────

Write-Host "========================================" -ForegroundColor Yellow
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

# Auto-open the log folder in Explorer
if ($doInstall -and $doCollectLog -and (Test-Path $localLogFolder)) {
    Write-Host "Opening log folder..." -ForegroundColor DarkGray
    Invoke-Item $localLogFolder
}
