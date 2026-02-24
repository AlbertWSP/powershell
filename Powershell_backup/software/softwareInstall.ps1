Add-Type -AssemblyName System.Windows.Forms

# Create a form
$form = New-Object System.Windows.Forms.Form
$form.Text = "Setup File Deployment"
$form.Size = New-Object System.Drawing.Size(350, 330)

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
$checkBox1.Text = "Sads Setup"
$checkBox1.Location = New-Object System.Drawing.Point(10, 60)
$form.Controls.Add($checkBox1)

# Add second checkbox
$checkBox2 = New-Object System.Windows.Forms.CheckBox
$checkBox2.Text = "BDE Setup"
$checkBox2.Location = New-Object System.Drawing.Point(10, 90)
$form.Controls.Add($checkBox2)

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

# Define the setup file paths and destination paths
$setupFilePath1 = "\\corp\hk\Transition\DAT18\Kits\Software\Engineering\SADS\SADS ( for new Dongle on HKKWN200lic01)\SADS 16\SadsSetup.exe"
$setupFilePath2 = "\\corp\hk\Transition\DAT18\Kits\Software\Engineering\SADS\SADS ( for new Dongle on HKKWN200lic01)\SADS 16\BdeSetup.exe"

$destinationPath1 = "\\$hostName\C$\temp\SADS 16\SadsSetup.exe"
$destinationPath2 = "\\$hostName\C$\temp\SADS 16\BdeSetup.exe"



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
                 -setpath "C:\temp\SADS 16\SadsSetup.exe" `
                 -commandLine "/VERYSILENT /NORESTART"
}

if ($option2Enabled) {
    Perform-Setup -hostName $hostName `
                 -setupFilePath $setupFilePath2 `
                 -destinationPath $destinationPath2 `
                 -setpath "C:\temp\SADS 16\BdeSetup.exe" `
                 -commandLine "/VERYSILENT /NORESTART"
}

[System.Windows.Forms.MessageBox]::Show("Done")