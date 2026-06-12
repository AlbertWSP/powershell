# 1. Load required assemblies at the start
Add-Type -AssemblyName System.Windows.Forms
Add-Type -AssemblyName System.Drawing

# Create the form object
$form = New-Object System.Windows.Forms.Form
$form.Text = "Setup File Deployment"
$form.Size = New-Object System.Drawing.Size(350, 300)

# Label for Host Name
$labelHostName = New-Object System.Windows.Forms.Label
$labelHostName.Text = "Enter Host Name:"
$labelHostName.Location = New-Object System.Drawing.Point(10, 20)
$form.Controls.Add($labelHostName)

# TextBox for Host Name
$textBoxHostName = New-Object System.Windows.Forms.TextBox
$textBoxHostName.Location = New-Object System.Drawing.Point(120, 20)
$form.Controls.Add($textBoxHostName)

# Checkbox for Firmware Update
$checkBox1 = New-Object System.Windows.Forms.CheckBox
$checkBox1.Text = "SanDisk Firmware Update"
$checkBox1.AutoSize = $true
$checkBox1.Location = New-Object System.Drawing.Point(10, 60)
$form.Controls.Add($checkBox1)

# OK Button
$okButton = New-Object System.Windows.Forms.Button
$okButton.Text = "OK"
$okButton.Location = New-Object System.Drawing.Point(100, 180)
$okButton.Add_Click({ $form.Tag = "OK"; $form.Close() })
$form.Controls.Add($okButton)

# 2. Show the form and capture the result
$null = $form.ShowDialog()

# 3. Only run if OK was clicked
if ($form.Tag -eq "OK") {
    $hostName = $textBoxHostName.Text
    $option1Enabled = $checkBox1.Checked

    if ($option1Enabled -and $hostName) {
        $setupFilePath = "\\corp\hk\Transition\DAT18\Kits\Software\Engineering\sharePC_driver\SanDisk-SN8000s-SED-Solid-State-Drive-Firmware-Update_0NWH2_WIN64_6311.2104_A04.exe"
        $destinationPath = "\\$hostName\C$\temp\"

        # Check if remote folder exists, if not, create it
        if (!(Test-Path $destinationPath)) {
            New-Item -ItemType Directory -Path $destinationPath -Force | Out-Null
        }

        Write-Host "Copying to $hostName..."
        Copy-Item -Path $setupFilePath -Destination $destinationPath -Force
        Write-Host "Done!"
    }
}
