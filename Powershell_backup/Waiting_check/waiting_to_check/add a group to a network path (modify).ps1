Add-Type -AssemblyName System.Windows.Forms

# Create the form
$form = New-Object System.Windows.Forms.Form
$form.Text = "Input Group Name and Network Path"
$form.Size = New-Object System.Drawing.Size(800, 200)
$form.TopMost = $true  # Make the form always on top

# Create the label and text box for the group name
$groupLabel = New-Object System.Windows.Forms.Label
$groupLabel.Text = "Group Name:"
$groupLabel.Location = New-Object System.Drawing.Point(10, 20)
$form.Controls.Add($groupLabel)

$groupTextBox = New-Object System.Windows.Forms.TextBox
$groupTextBox.Location = New-Object System.Drawing.Point(150, 20)  # Shifted to the right
$groupTextBox.Size = New-Object System.Drawing.Size(600, 20)
$form.Controls.Add($groupTextBox)

# Create the label and text box for the network path
$pathLabel = New-Object System.Windows.Forms.Label
$pathLabel.Text = "Network Path:"
$pathLabel.Location = New-Object System.Drawing.Point(10, 60)
$form.Controls.Add($pathLabel)

$pathTextBox = New-Object System.Windows.Forms.TextBox
$pathTextBox.Location = New-Object System.Drawing.Point(150, 60)  # Shifted to the right
$pathTextBox.Size = New-Object System.Drawing.Size(600, 20)
$form.Controls.Add($pathTextBox)

# Create the submit button
$submitButton = New-Object System.Windows.Forms.Button
$submitButton.Text = "Submit"
$submitButton.Location = New-Object System.Drawing.Point(150, 100)
$form.Controls.Add($submitButton)

# Define the button click event
$submitButton.Add_Click({
    $groupName = $groupTextBox.Text
    $networkPath = $pathTextBox.Text

    # Get the current ACL for the network path
    $acl = Get-Acl -Path $networkPath

    # Create a new access rule for the group with "Modify" rights
    $accessRule = New-Object System.Security.AccessControl.FileSystemAccessRule($groupName, "Modify", "ContainerInherit,ObjectInherit", "None", "Allow")

    # Add the access rule to the ACL
    $acl.SetAccessRule($accessRule)

    # Apply the updated ACL to the network path
    Set-Acl -Path $networkPath -AclObject $acl

    [System.Windows.Forms.MessageBox]::Show("Group '$groupName' has been added to '$networkPath' with 'Modify' access rights.")
})

# Show the form
$form.ShowDialog()