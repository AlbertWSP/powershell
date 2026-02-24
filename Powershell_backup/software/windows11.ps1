Add-Type -AssemblyName System.Windows.Forms

# Create a form
$form = New-Object System.Windows.Forms.Form
$form.Text = "Setup File Deployment"
$form.Size = New-Object System.Drawing.Size(350, 400)

# Create a label for the host name
$labelHostName = New-Object System.Windows.Forms.Label
$labelHostName.Text = "Enter Host Name:"
$labelHostName.Location = New-Object System.Drawing.Point(10, 20)
$form.Controls.Add($labelHostName)

# Create a textbox for the host name
$textBoxHostName = New-Object System.Windows.Forms.TextBox
$textBoxHostName.Location = New-Object System.Drawing.Point(120, 20)
$form.Controls.Add($textBoxHostName)

# Create checkboxes
$checkBox1 = New-Object System.Windows.Forms.CheckBox
$checkBox1.Text = "DSignerV1.3.6"
$checkBox1.Location = New-Object System.Drawing.Point(10, 60)
$form.Controls.Add($checkBox1)

# Add new checkbox for PDF-Xchange Editor
$checkBox2 = New-Object System.Windows.Forms.CheckBox
$checkBox2.Text = "PDF-Xchange Editor V10"
$checkBox2.Location = New-Object System.Drawing.Point(10, 90)
$form.Controls.Add($checkBox2)

# Add new checkbox for Prokon
$checkBox3 = New-Object System.Windows.Forms.CheckBox
$checkBox3.Text = "Prokon 5.2"
$checkBox3.Location = New-Object System.Drawing.Point(10, 120)
$form.Controls.Add($checkBox3)

# Create an OK button
$okButton = New-Object System.Windows.Forms.Button
$okButton.Text = "OK"
$okButton.Location = New-Object System.Drawing.Point(100, 240)
$okButton.Add_Click({
    $form.Close()
})
$form.Controls.Add($okButton)

# Show the form
$form.ShowDialog()

# Get the host name and checkbox states
$hostName = $textBoxHostName.Text
$option1Enabled = $checkBox1.Checked
$option2Enabled = $checkBox2.Checked
$option3Enabled = $checkBox3.Checked

# Define the setup file paths and destination paths
$setupFilePath1 = "\\corp.pbwan.net\hk\Transition\DAT18\kits\Software\Engineering\Setup (DSigner) v1.3.6.exe"
$setupFilePath2 = "\\corp\hk\Transition\DAT18\Kits\Software\Engineering\PDF Xchange Editor\PDF-Xchange Editor V10.msi"
$setupFilePath3 = "\\corp\hk\Transition\DAT18\Kits\Software\Engineering\Prokon\5.2\ProkonSetup5 (1).exe"

$destinationPath1 = "\\$hostName\C$\temp\Setup (DSigner) v1.3.6.exe"
$destinationPath2 = "\\$hostName\C$\temp\PDF-Xchange Editor V10.msi"
$destinationPath3 = "\\$hostName\C$\temp\ProkonSetup5.exe"

# Function to perform the setup actions
function Perform-Setup {
    param (
        [string]$hostName,
        [string]$setupFilePath,
        [string]$destinationPath,
        $setpath,
        $commandLine
    )

    try {
        # Copy the setup file to the host
        Copy-Item -Path $setupFilePath -Destination $destinationPath -ErrorAction Stop

        # Run the setup file on the remote host
        Invoke-Command -ComputerName $hostName -ScriptBlock {
            param ($setpath, $commandLine)
            Start-Process -FilePath $setpath -ArgumentList $commandLine -Wait -ErrorAction Stop
        } -ArgumentList $setpath, "/VERYSILENT /NORESTART" # Silent install arguments added

        # Delete the setup file from the host
        Remove-Item -Path $destinationPath -ErrorAction SilentlyContinue
    }
    catch {
        [System.Windows.Forms.MessageBox]::Show("Error during installation: $_", "Error", [System.Windows.Forms.MessageBoxButtons]::OK, [System.Windows.Forms.MessageBoxIcon]::Error)
    }
}

# Perform setup actions based on checkbox states
if ($option1Enabled) {
    Perform-Setup -hostName $hostName `
                 -setupFilePath $setupFilePath1 `
                 -destinationPath $destinationPath1 `
                 -setpath "C:\temp\Setup (DSigner) v1.3.6.exe" `
                 -commandLine "/VERYSILENT /NORESTART"
}

if ($option2Enabled) {
    Perform-Setup -hostName $hostName `
                 -setupFilePath $setupFilePath2 `
                 -destinationPath $destinationPath2 `
                 -setpath "C:\temp\PDF-Xchange Editor V10.msi" `
                 -commandLine "/quiet /norestart"
}

if ($option3Enabled) {
    Perform-Setup -hostName $hostName `
                 -setupFilePath $setupFilePath3 `
                 -destinationPath $destinationPath3 `
                 -setpath "C:\temp\ProkonSetup5.exe" `
                 -commandLine "/VERYSILENT /NORESTART"
}

[System.Windows.Forms.MessageBox]::Show("Done")