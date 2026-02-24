Add-Type -AssemblyName System.Windows.Forms

# Create and display the input box
$inputBox = New-Object System.Windows.Forms.Form
$inputBox.Text = "Network Path ACL Viewer"
$inputBox.Size = New-Object System.Drawing.Size(600, 800)
$inputBox.StartPosition = "CenterScreen"

# Label for input
$label = New-Object System.Windows.Forms.Label
$label.Text = "Please enter the network path:"
$label.AutoSize = $true
$label.Location = New-Object System.Drawing.Point(10, 20)
$inputBox.Controls.Add($label)

# TextBox for input
$textBox = New-Object System.Windows.Forms.TextBox
$textBox.Size = New-Object System.Drawing.Size(560, 20)
$textBox.Location = New-Object System.Drawing.Point(10, 50)
$inputBox.Controls.Add($textBox)



# RichTextBox for displaying results
$resultBox = New-Object System.Windows.Forms.RichTextBox
$resultBox.Size = New-Object System.Drawing.Size(560, 550)
$resultBox.Location = New-Object System.Drawing.Point(10, 130)
$resultBox.ReadOnly = $true
$resultBox.BackColor = [System.Drawing.Color]::White
$resultBox.ScrollBars = "Vertical"
$resultBox.Text = "Results will be displayed here..."
$inputBox.Controls.Add($resultBox)

# OK Button
$okButton = New-Object System.Windows.Forms.Button
$okButton.Text = "OK"
$okButton.Location = New-Object System.Drawing.Point(10, 90)
$okButton.Add_Click({
    # Get the path from the text box
    $path = $textBox.Text.Trim()
    
    # Validate input
    if ([string]::IsNullOrWhiteSpace($path)) {
        $resultBox.Text = "Error: No path entered. Please provide a valid network path."
        return
    }

    try {
        # Check ACL for the specified path
        $acl = Get-Acl -Path $path

        # Build result string for ACL details
        $resultText = "Access Control List (ACL) for path: $path`n`n"
        $resultText += ($acl | Format-List | Out-String)

        # Add permissions details
        $resultText += "`nPermissions:`n"
        $acl.Access | ForEach-Object {
            $resultText += "Identity: $_.IdentityReference`n"
            $resultText += "Access Type: $_.AccessControlType`n"
            $resultText += "Rights: $_.FileSystemRights`n"
            $resultText += "Inherited: $_.IsInherited`n"
            $resultText += "`n"
        }

        # Display results in RichTextBox
        $resultBox.Text = $resultText

    } catch {
        # Handle errors (e.g., invalid path or insufficient permissions)
        $resultBox.Text = "Error: Unable to retrieve ACL for the specified path.`n"
        $resultBox.Text += $_.Exception.Message
    }
})
$inputBox.Controls.Add($okButton)

# Cancel Button (optional for better UX)
$cancelButton = New-Object System.Windows.Forms.Button
$cancelButton.Text = "Cancel"
$cancelButton.Location = New-Object System.Drawing.Point(90, 90)
$cancelButton.Add_Click({
    $inputBox.Close()
})
$inputBox.Controls.Add($cancelButton)

# Show Dialog
[void] $inputBox.ShowDialog()
