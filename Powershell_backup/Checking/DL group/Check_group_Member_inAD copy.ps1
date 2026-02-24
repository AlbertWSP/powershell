Add-Type -AssemblyName System.Windows.Forms

function Get-GroupMembersRecursively {
    param (
        [string]$groupName
    )

    # Get the group members
    $groupMembers = Get-ADGroupMember -Identity $groupName

    # Display the group members
    foreach ($member in $groupMembers) {
        Write-Output $member.Name

        # If the member is a group, call the function recursively
        if ($member.objectClass -eq "group") {
            Write-Output "Checking members of group: $($member.Name)"
            Get-GroupMembersRecursively -groupName $member.SamAccountName
        }
    }
}

# Create a Windows Forms input box
[System.Windows.Forms.Application]::EnableVisualStyles()
$form = New-Object System.Windows.Forms.Form
$form.Text = "Enter AD Group Name"
$form.Size = New-Object System.Drawing.Size(300,150)

$label = New-Object System.Windows.Forms.Label
$label.Text = "Group Name:"
$label.AutoSize = $true
$label.Location = New-Object System.Drawing.Point(10,20)
$form.Controls.Add($label)

$textBox = New-Object System.Windows.Forms.TextBox
$textBox.Size = New-Object System.Drawing.Size(200,20)
$textBox.Location = New-Object System.Drawing.Point(100,20)
$form.Controls.Add($textBox)

$okButton = New-Object System.Windows.Forms.Button
$okButton.Text = "OK"
$okButton.Location = New-Object System.Drawing.Point(100,60)
$okButton.Add_Click({
    $form.Close()
})
$form.Controls.Add($okButton)

$form.ShowDialog()

# Get the group name from the input box
$groupName = $textBox.Text

# Get the group information
$group = Get-ADGroup -Identity $groupName

# Display the group information
Write-Output "Group Name: $($group.Name)"
Write-Output "Group Description: $($group.Description)"
Write-Output "Group Distinguished Name: $($group.DistinguishedName)"
Write-Output "Group Object GUID: $($group.ObjectGUID)"
Write-Output "Group Created Date: $($group.WhenCreated)"
Write-Output "Group Modified Date: $($group.WhenChanged)"
Write-Output ""

# Display the group members recursively
Write-Output "Group Members:"
Get-GroupMembersRecursively -groupName $groupName