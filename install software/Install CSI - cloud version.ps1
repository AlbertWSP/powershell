Add-Type -AssemblyName System.Windows.Forms

# Create a form
$form = New-Object System.Windows.Forms.Form
$form.Text = "Setup File Deployment"
$form.Size = New-Object System.Drawing.Size(350, 300)

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
$checkBox1.Text = "ETABS 22.1.0"
$checkBox1.Location = New-Object System.Drawing.Point(10, 60)
$form.Controls.Add($checkBox1)

$checkBox2 = New-Object System.Windows.Forms.CheckBox
$checkBox2.Text = "SAFE 22.1.0"
$checkBox2.Location = New-Object System.Drawing.Point(10, 90)
$form.Controls.Add($checkBox2)

$checkBox3 = New-Object System.Windows.Forms.CheckBox
$checkBox3.Text = "SAP2K 25.1.0"
$checkBox3.Location = New-Object System.Drawing.Point(10, 120)
$form.Controls.Add($checkBox3)

# Create an OK button
$okButton = New-Object System.Windows.Forms.Button
$okButton.Text = "OK"
$okButton.Location = New-Object System.Drawing.Point(100, 160)
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
$setupFilePath1 = "\\hkkwn200dat18\kits\Software\Engineering\ETABS\ETABSv2210Setup.exe"
$setupFilePath2 = "\\hkkwn200dat18\kits\Software\Engineering\SAFE\SAFEv2210Setup.exe"
$setupFilePath3 = "\\hkkwn200dat18\kits\Software\Engineering\SAP2000\SAP2000v2510Setup.exe"
$destinationPath1 = "\\$hostName\C$\temp\ETABSv2210Setup.exe"
$destinationPath2 = "\\$hostName\C$\temp\SAFEv2210Setup.exe"
$destinationPath3 = "\\$hostName\C$\temp\SAP2000v2510Setup.exe"

# Function to perform the setup actions
function Perform-Setup {
    param (
        [string]$hostName,
        [string]$setupFilePath,
        [string]$destinationPath,
        $setpath,
        $commandLine
    )

    # Copy the setup file to the host
    Copy-Item -Path $setupFilePath -Destination $destinationPath

    # Run the setup file on the remote host
    Invoke-Command -ComputerName $hostName -ScriptBlock {
        param ($setpath, $commandLine)
        Start-Process -FilePath $setpath -ArgumentList $commandLine -Wait
    } -ArgumentList $setpath, $commandLine

    # Delete the setup file from the host
    Remove-Item -Path $destinationPath
}

# Perform setup actions based on checkbox states
if ($option1Enabled) {
    Perform-Setup -hostName $hostName -setupFilePath $setupFilePath1 -destinationPath $destinationPath1 -setpath "C:\temp\ETABSv2210Setup.exe" -commandLine '/s /v"/qn LicenseMode=Login"'
}
if ($option2Enabled) {
    Perform-Setup -hostName $hostName -setupFilePath $setupFilePath2 -destinationPath $destinationPath2 -setpath "C:\temp\SAFEv2210Setup.exe" -commandLine '/s /v"/qn LicenseMode=Login"'
}
if ($option3Enabled) {
    Perform-Setup -hostName $hostName -setupFilePath $setupFilePath3 -destinationPath $destinationPath3 -setpath "C:\temp\SAP2000v2510Setup.exe" -commandLine '/s /v"/qn LicenseMode=Login"'
}

[System.Windows.Forms.MessageBox]::Show("Done")
