Add-Type -AssemblyName System.Windows.Forms

# Create a form
$form = New-Object System.Windows.Forms.Form
$form.Text = "Setup File Deployment"
$form.Size = New-Object System.Drawing.Size(350, 350)

# Create a label for the host name
$labelHostName = New-Object System.Windows.Forms.Label
$labelHostName.Text = "Enter Host Names (one per line):"
$labelHostName.Location = New-Object System.Drawing.Point(10, 20)
$form.Controls.Add($labelHostName)

# Replace textbox with multiline textbox for host names
$textBoxHostName = New-Object System.Windows.Forms.TextBox
$textBoxHostName.Location = New-Object System.Drawing.Point(120, 20)
$textBoxHostName.Size = New-Object System.Drawing.Size(200, 100)
$textBoxHostName.Multiline = $true
$textBoxHostName.ScrollBars = "Vertical"
$form.Controls.Add($textBoxHostName)

# Create a "Load from File" button
$loadButton = New-Object System.Windows.Forms.Button
$loadButton.Text = "Load from File"
$loadButton.Location = New-Object System.Drawing.Point(120, 125)
$loadButton.Size = New-Object System.Drawing.Size(100, 23)
$loadButton.Add_Click({
    $openFileDialog = New-Object System.Windows.Forms.OpenFileDialog
    $openFileDialog.Filter = "Text files (*.txt)|*.txt|All files (*.*)|*.*"
    $openFileDialog.InitialDirectory = [Environment]::GetFolderPath("Desktop")
    
    if ($openFileDialog.ShowDialog() -eq 'OK') {
        try {
            $textBoxHostName.Text = Get-Content -Path $openFileDialog.FileName -Raw
        }
        catch {
            [System.Windows.Forms.MessageBox]::Show("Error reading file: $_", "Error", [System.Windows.Forms.MessageBoxButtons]::OK, [System.Windows.Forms.MessageBoxIcon]::Error)
        }
    }
})
$form.Controls.Add($loadButton)

# Create checkboxes
$checkBox1 = New-Object System.Windows.Forms.CheckBox
$checkBox1.Text = "DSignerV1.3.6"
$checkBox1.Location = New-Object System.Drawing.Point(10, 160)
$form.Controls.Add($checkBox1)

# Add new checkbox for PDF-Xchange Editor
$checkBox2 = New-Object System.Windows.Forms.CheckBox
$checkBox2.Text = "PDF-Xchange Editor V10"
$checkBox2.Location = New-Object System.Drawing.Point(10, 190)
$form.Controls.Add($checkBox2)

# Create an OK button
$okButton = New-Object System.Windows.Forms.Button
$okButton.Text = "OK"
$okButton.Location = New-Object System.Drawing.Point(100, 280)
$okButton.Add_Click({
    $form.Close()
})
$form.Controls.Add($okButton)

# Show the form
$form.ShowDialog()

# Get the host names as array and checkbox states
$hostNames = $textBoxHostName.Text -split "`r`n" | Where-Object { $_ -match '\S' }
$option1Enabled = $checkBox1.Checked
$option2Enabled = $checkBox2.Checked

# Define the setup file paths and destination paths
$setupFilePath1 = "\\bbp2g43\share_g$\software\Setup (DSigner) v1.3.7.exe"
$setupFilePath2 = "\\corp\hk\Transition\DAT18\Kits\Software\Engineering\PDF Xchange Editor\PDF-Xchange Editor V10.msi"

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
        Write-Host "Installing on $hostName..."
        Copy-Item -Path $setupFilePath -Destination $destinationPath -ErrorAction Stop

        Invoke-Command -ComputerName $hostName -ScriptBlock {
            param ($setpath, $commandLine)
            Start-Process -FilePath $setpath -ArgumentList $commandLine -Wait -ErrorAction Stop
        } -ArgumentList $setpath, "/VERYSILENT /NORESTART"

        Remove-Item -Path $destinationPath -ErrorAction SilentlyContinue
        return $true
    }
    catch {
        Write-Host "Failed to install on $hostName : $_"
        return $false
    }
}

# Process installation for each PC
$results = @()
if ($option1Enabled) {
    foreach ($hostName in $hostNames) {
        $destinationPath1 = "\\$hostName\C$\temp\Setup (DSigner) v1.3.7.exe"
        $success = Perform-Setup -hostName $hostName `
                               -setupFilePath $setupFilePath1 `
                               -destinationPath $destinationPath1 `
                               -setpath "C:\temp\Setup (DSigner) v1.3.7.exe" `
                               -commandLine "/VERYSILENT /NORESTART"
        
        $results += "PC $hostName : DSigner $(if ($success) {'Success'} else {'Failed'})"
    }
}

if ($option2Enabled) {
    foreach ($hostName in $hostNames) {
        $destinationPath2 = "\\$hostName\C$\temp\PDF-Xchange Editor V10.msi"
        $success = Perform-Setup -hostName $hostName `
                               -setupFilePath $setupFilePath2 `
                               -destinationPath $destinationPath2 `
                               -setpath "C:\temp\PDF-Xchange Editor V10.msi" `
                               -commandLine "/quiet /norestart"
        
        $results += "PC $hostName : PDF-Xchange $(if ($success) {'Success'} else {'Failed'})"
    }
}

# Show results
[System.Windows.Forms.MessageBox]::Show(($results -join "`n"), "Installation Results")