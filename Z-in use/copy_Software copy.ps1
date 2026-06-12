Add-Type -AssemblyName System.Windows.Forms
Add-Type -AssemblyName System.Drawing

# Create a form
$form = New-Object System.Windows.Forms.Form
$form.Text = "Setup File Deployment"
$form.Size = New-Object System.Drawing.Size(400, 450)
$form.StartPosition = "CenterScreen"

# Create a label for the host name
$labelHostName = New-Object System.Windows.Forms.Label
$labelHostName.Text = "Enter Host Names (One per line):"
$labelHostName.Location = New-Object System.Drawing.Point(10, 10)
$labelHostName.Size = New-Object System.Drawing.Size(200, 20)
$form.Controls.Add($labelHostName)

# Create a MULTILINE textbox for the host names
$textBoxHostName = New-Object System.Windows.Forms.TextBox
$textBoxHostName.Location = New-Object System.Drawing.Point(10, 35)
$textBoxHostName.Size = New-Object System.Drawing.Size(360, 150)
$textBoxHostName.Multiline = $true
$textBoxHostName.ScrollBars = "Vertical"
$form.Controls.Add($textBoxHostName)

# Create checkbox
$checkBox1 = New-Object System.Windows.Forms.CheckBox
$checkBox1.Text = "SanDisk SN8000s Firmware Update"
$checkBox1.Location = New-Object System.Drawing.Point(10, 200)
$checkBox1.Size = New-Object System.Drawing.Size(360, 30)
$form.Controls.Add($checkBox1)

# Create an OK button
$okButton = New-Object System.Windows.Forms.Button
$okButton.Text = "OK"
$okButton.Location = New-Object System.Drawing.Point(140, 350)
$okButton.DialogResult = [System.Windows.Forms.DialogResult]::OK
$okButton.Add_Click({ $form.Tag = "OK"; $form.Close() })
$form.Controls.Add($okButton)

# Show the form
$result = $form.ShowDialog()

# Process only if OK was clicked
if ($form.Tag -eq "OK") {
    # Get the list of host names, split by new line, and remove empty entries
    $hostNames = $textBoxHostName.Text -split "`r`n" | Where-Object { $_ -match '\S' }
    $option1Enabled = $checkBox1.Checked

    $setupFilePath = "\\corp\hk\Transition\DAT18\Kits\Software\Engineering\sharePC_driver\SanDisk-SN8000s-SED-Solid-State-Drive-Firmware-Update_0NWH2_WIN64_6311.2104_A04.exe"

    if ($option1Enabled) {
        foreach ($name in $hostNames) {
            $cleanName = $name.Trim()
            $destinationPath = "\\$cleanName\C$\temp\"

            Write-Host "--- Processing: $cleanName ---" -ForegroundColor Cyan

            try {
                # Test connectivity first
                if (Test-Connection -ComputerName $cleanName -Count 1 -Quiet) {
                    
                    # Create destination if it doesn't exist
                    if (!(Test-Path -Path $destinationPath)) {
                        New-Item -ItemType Directory -Path $destinationPath -Force | Out-Null
                    }

                    # Copy the file
                    Write-Host "Copying setup file..."
                    Copy-Item -Path $setupFilePath -Destination $destinationPath -Force
                    Write-Host "Successfully copied to $cleanName" -ForegroundColor Green
                }
                else {
                    Write-Warning "Target $cleanName is offline or unreachable."
                }
            }
            catch {
                Write-Error "Failed to process $cleanName. Error: $($_.Exception.Message)"
            }
        }
    }
    else {
        Write-Host "No options selected. Script exiting."
    }
}
