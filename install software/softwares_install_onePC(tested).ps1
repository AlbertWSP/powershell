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
        # Get file size for progress calculation
        $fileSize = (Get-Item -Path $setupFilePath).Length
        $reader = [System.IO.File]::OpenRead($setupFilePath)
        $writer = [System.IO.File]::Create($destinationPath)
        $buffer = New-Object byte[] 1048576  # 1MB buffer
        $totalBytesRead = 0

        Write-Host "Copying $setupFilePath to $destinationPath"
        
        while ($bytesRead = $reader.Read($buffer, 0, $buffer.Length)) {
            $writer.Write($buffer, 0, $bytesRead)
            $totalBytesRead += $bytesRead
            $percentComplete = [math]::Round(($totalBytesRead / $fileSize) * 100, 2)
            
            Write-Progress -Activity "Copying File" `
                         -Status "Progress: $percentComplete%" `
                         -PercentComplete $percentComplete
        }

        $reader.Close()
        $writer.Close()

        # Run the setup file on the remote host
        Invoke-Command -ComputerName $hostName -ScriptBlock {
            param ($setpath, $commandLine)
            Start-Process -FilePath $setpath -ArgumentList $commandLine -Wait -ErrorAction Stop
        } -ArgumentList $setpath, "/VERYSILENT /NORESTART"

        # Delete the setup file from the host
        Remove-Item -Path $destinationPath -ErrorAction SilentlyContinue
    }
    catch {
        [System.Windows.Forms.MessageBox]::Show("Error during installation: $_", "Error", [System.Windows.Forms.MessageBoxButtons]::OK, [System.Windows.Forms.MessageBoxIcon]::Error)
    }
    finally {
        if ($reader) { $reader.Dispose() }
        if ($writer) { $writer.Dispose() }
        Write-Progress -Activity "Copying File" -Completed
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
                 -commandLine "/VERYSILENT /NORESTART"
}

if ($option3Enabled) {
    try {
        # First copy the file using robocopy to show progress
        $sourceDir = Split-Path $setupFilePath3 -Parent
        $sourceFile = Split-Path $setupFilePath3 -Leaf
        $destDir = "\\$hostName\C$\temp"
        
        Write-Host "Copying Prokon setup file to $hostName..."
        robocopy $sourceDir $destDir $sourceFile /W:1 /R:1 /V /NP
        
        # Launch the installer directly on the remote machine
        Write-Host "Launching Prokon installer on $hostName..."
        Invoke-Command -ComputerName $hostName -ScriptBlock {
            Start-Process -FilePath "C:\temp\ProkonSetup5.exe" -Wait
        }
        
        # Cleanup after installation
        Write-Host "Cleaning up..."
        Remove-Item -Path "$destDir\$sourceFile" -Force -ErrorAction SilentlyContinue
    }
    catch {
        [System.Windows.Forms.MessageBox]::Show("Error with Prokon installation: $_", "Error", 
            [System.Windows.Forms.MessageBoxButtons]::OK, 
            [System.Windows.Forms.MessageBoxIcon]::Error)
    }
}

[System.Windows.Forms.MessageBox]::Show("Done")
