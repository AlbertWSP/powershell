Add-Type -AssemblyName System.Windows.Forms

# Create and display the input box
$inputBox = New-Object System.Windows.Forms.Form
$inputBox.Text = "Enter Network Path"
$inputBox.Size = New-Object System.Drawing.Size(400, 150)

$label = New-Object System.Windows.Forms.Label
$label.Text = "Please enter the network path:"
$label.AutoSize = $true
$label.Location = New-Object System.Drawing.Point(10, 20)
$inputBox.Controls.Add($label)

$textBox = New-Object System.Windows.Forms.TextBox
$textBox.Size = New-Object System.Drawing.Size(360, 20)
$textBox.Location = New-Object System.Drawing.Point(10, 50)
$inputBox.Controls.Add($textBox)

$okButton = New-Object System.Windows.Forms.Button
$okButton.Text = "OK"
$okButton.Location = New-Object System.Drawing.Point(150, 80)
$okButton.Add_Click({ $inputBox.Close() })
$inputBox.Controls.Add($okButton)

$inputBox.ShowDialog()

# Get the path from the text box
$path = $textBox.Text

# Check the ACL for the specified path
$acl = Get-Acl -Path $path

# Display the ACL
$acl | Format-List

# Display the members and their permissions
$acl.Access | ForEach-Object {
    [PSCustomObject]@{
        IdentityReference = $_.IdentityReference
        AccessControlType = $_.AccessControlType
        FileSystemRights  = $_.FileSystemRights
        InheritanceFlags  = $_.InheritanceFlags
        PropagationFlags  = $_.PropagationFlags
        IsInherited       = $_.IsInherited
    }
} | Format-Table -AutoSize