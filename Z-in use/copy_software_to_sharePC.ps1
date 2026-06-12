Add-Type -AssemblyName System.Windows.Forms
Add-Type -AssemblyName System.Drawing

# ── GUI Setup ──────────────────────────────────────────────────────────────────

$form = New-Object System.Windows.Forms.Form
$form.Text = "Setup File Deployment"
$form.Size = New-Object System.Drawing.Size(420, 480)
$form.StartPosition = "CenterScreen"
$form.FormBorderStyle = "FixedDialog"
$form.MaximizeBox = $false

# Label
$labelHostName = New-Object System.Windows.Forms.Label
$labelHostName.Text = "Enter Host Names (one per line):"
$labelHostName.Location = New-Object System.Drawing.Point(10, 10)
$labelHostName.Size = New-Object System.Drawing.Size(380, 20)
$form.Controls.Add($labelHostName)

# Multiline textbox
$textBoxHostName = New-Object System.Windows.Forms.TextBox
$textBoxHostName.Location = New-Object System.Drawing.Point(10, 35)
$textBoxHostName.Size = New-Object System.Drawing.Size(380, 150)
$textBoxHostName.Multiline = $true
$textBoxHostName.ScrollBars = "Vertical"
$textBoxHostName.Font = New-Object System.Drawing.Font("Consolas", 10)
$form.Controls.Add($textBoxHostName)

# ── Deployment Options (Checkboxes) ───────────────────────────────────────────
# To add a new option: 1) create a checkbox  2) add an entry in $deployments below

$checkBox1 = New-Object System.Windows.Forms.CheckBox
$checkBox1.Text = "SanDisk SN8000s Firmware Update"
$checkBox1.Location = New-Object System.Drawing.Point(10, 200)
$checkBox1.Size = New-Object System.Drawing.Size(380, 25)
$form.Controls.Add($checkBox1)

# $checkBox2 = New-Object System.Windows.Forms.CheckBox
# $checkBox2.Text = "Example - Another Deployment"
# $checkBox2.Location = New-Object System.Drawing.Point(10, 230)
# $checkBox2.Size = New-Object System.Drawing.Size(380, 25)
# $form.Controls.Add($checkBox2)

# ── Buttons ────────────────────────────────────────────────────────────────────

$okButton = New-Object System.Windows.Forms.Button
$okButton.Text = "OK"
$okButton.Size = New-Object System.Drawing.Size(90, 30)
$okButton.Location = New-Object System.Drawing.Point(120, 390)
$okButton.Add_Click({ $form.Tag = "OK"; $form.Close() })
$form.Controls.Add($okButton)

$cancelButton = New-Object System.Windows.Forms.Button
$cancelButton.Text = "Cancel"
$cancelButton.Size = New-Object System.Drawing.Size(90, 30)
$cancelButton.Location = New-Object System.Drawing.Point(220, 390)
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

# ── Build Deployment Map ───────────────────────────────────────────────────────
# Maps a friendly label -> source file path for every checked option

$deployments = [ordered]@{}

if ($checkBox1.Checked) {
    $deployments["SanDisk SN8000s FW"] = "\\corp\hk\Transition\DAT18\Kits\Software\Engineering\sharePC_driver\SanDisk-SN8000s-SED-Solid-State-Drive-Firmware-Update_0NWH2_WIN64_6311.2104_A04.exe"
}

# if ($checkBox2.Checked) {
#     $deployments["Another Tool"] = "\\server\share\path\to\installer.exe"
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

# Pre-check: make sure every source file exists before we start
foreach ($label in $deployments.Keys) {
    if (!(Test-Path -Path $deployments[$label])) {
        Write-Error "Source file not found for [$label]: $($deployments[$label])"
        return
    }
}

# ── Deploy ─────────────────────────────────────────────────────────────────────

$successCount = 0
$failCount = 0
$totalTasks = $hostNames.Count * $deployments.Count

Write-Host "`n========================================" -ForegroundColor Cyan
Write-Host "  Starting Deployment ($totalTasks task(s))" -ForegroundColor Cyan
Write-Host "========================================`n" -ForegroundColor Cyan

foreach ($name in $hostNames) {
    Write-Host "--- Processing: $name ---" -ForegroundColor Cyan

    try {
        # Connectivity check
        if (!(Test-Connection -ComputerName $name -Count 1 -Quiet)) {
            Write-Warning "  $name is offline or unreachable."
            $failCount += $deployments.Count   # count each skipped item
            continue
        }

        # Ensure destination folder exists
        $destinationPath = "\\$name\C$\temp\"
        if (!(Test-Path -Path $destinationPath)) {
            New-Item -ItemType Directory -Path $destinationPath -Force | Out-Null
            Write-Host "  Created $destinationPath" -ForegroundColor DarkGray
        }

        # Copy each selected deployment file
        foreach ($label in $deployments.Keys) {
            Write-Host "  Copying [$label] ..." -NoNewline
            Copy-Item -Path $deployments[$label] -Destination $destinationPath -Force
            Write-Host " Done" -ForegroundColor Green
            $successCount++
        }
    }
    catch {
        Write-Error "  Failed on $name - $($_.Exception.Message)"
        $failCount++
    }
}

# ── Summary ────────────────────────────────────────────────────────────────────

Write-Host "`n========================================" -ForegroundColor Yellow
Write-Host "  Deployment Summary" -ForegroundColor Yellow
Write-Host "  Success : $successCount" -ForegroundColor Green
Write-Host "  Failed  : $failCount" -ForegroundColor $(if ($failCount -gt 0) { "Red" } else { "Green" })
Write-Host "  Total   : $totalTasks" -ForegroundColor Yellow
Write-Host "========================================`n" -ForegroundColor Yellow
