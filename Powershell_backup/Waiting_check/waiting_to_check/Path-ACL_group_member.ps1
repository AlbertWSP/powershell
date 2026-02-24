Add-Type -AssemblyName System.Windows.Forms

# Function to get the SAM account name from an email address
function GetSamAccountName {
    param (
        [string]$Email
    )
    $user = Get-ADUser -Filter {Mail -eq $Email} -Properties SamAccountName
    return $user.SamAccountName
}

# Function to get all members of a group, including nested groups
function GetGroupMembers {
    param (
        [string]$GroupName
    )
    $group = Get-ADGroup -Identity $GroupName -Properties Members
    $members = @()

    foreach ($member in $group.Members) {
        $memberObject = Get-ADObject -Identity $member
        if ($memberObject.objectClass -eq 'group') {
            $members += GetGroupMembers -GroupName $memberObject.SamAccountName
        } else {
            $members += $memberObject
        }
    }
    return $members
}

# Create the form
$form = New-Object System.Windows.Forms.Form
$form.Text = "Input Network Path and Member Emails"
$form.Size = New-Object System.Drawing.Size(800, 300)
$form.TopMost = $true  # Make the form always on top

# Create the label and text box for the network path
$pathLabel = New-Object System.Windows.Forms.Label
$pathLabel.Text = "Network Path:"
$pathLabel.Location = New-Object System.Drawing.Point(10, 20)
$form.Controls.Add($pathLabel)

$pathTextBox = New-Object System.Windows.Forms.TextBox
$pathTextBox.Location = New-Object System.Drawing.Point(150, 20)  # Shifted to the right
$pathTextBox.Size = New-Object System.Drawing.Size(600, 20)
$form.Controls.Add($pathTextBox)

# Create the label and text box for the member emails
$emailLabel = New-Object System.Windows.Forms.Label
$emailLabel.Text = "Member Emails (comma-separated):"
$emailLabel.Location = New-Object System.Drawing.Point(10, 60)
$form.Controls.Add($emailLabel)

$emailTextBox = New-Object System.Windows.Forms.TextBox
$emailTextBox.Location = New-Object System.Drawing.Point(150, 60)  # Shifted to the right
$emailTextBox.Size = New-Object System.Drawing.Size(600, 20)
$form.Controls.Add($emailTextBox)

# Create the submit button
$submitButton = New-Object System.Windows.Forms.Button
$submitButton.Text = "Submit"
$submitButton.Location = New-Object System.Drawing.Point(150, 100)
$form.Controls.Add($submitButton)

# Define the button click event
$submitButton.Add_Click({
    $networkPath = $pathTextBox.Text
    $emailInput = $emailTextBox.Text

    # Split the input into an array
    $emailArray = $emailInput -split ","

    # Initialize a list to hold the results
    $results = @()

    # Check the effective access for each member
    foreach ($email in $emailArray) {
        $email = $email.Trim()
        $samAccountName = GetSamAccountName -Email $email

        # Get the effective access for the member
        $effectiveAccess = Get-EffectiveAccess -Path $networkPath -Principal $samAccountName

        $results += [PSCustomObject]@{
            Email = $email
            SAMAccountName = $samAccountName
            Permissions = $effectiveAccess.AccessRights
            AccessLimitedBy = $effectiveAccess.AccessLimitedBy
        }

        # Check if the member is a group and get its members
        $groupMembers = GetGroupMembers -GroupName $samAccountName
        foreach ($member in $groupMembers) {
            $effectiveAccess = Get-EffectiveAccess -Path $networkPath -Principal $member.SamAccountName

            $results += [PSCustomObject]@{
                Email = $email
                SAMAccountName = $member.SamAccountName
                Permissions = $effectiveAccess.AccessRights
                AccessLimitedBy = $effectiveAccess.AccessLimitedBy
            }
        }
    }

    # Display the results
    $resultsForm = New-Object System.Windows.Forms.Form
    $resultsForm.Text = "Effective Access Results"
    $resultsForm.Size = New-Object System.Drawing.Size(800, 400)
    $resultsForm.TopMost = $true

    $resultsTextBox = New-Object System.Windows.Forms.TextBox
    $resultsTextBox.Multiline = $true
    $resultsTextBox.ScrollBars = "Vertical"
    $resultsTextBox.Size = New-Object System.Drawing.Size(760, 360)
    $resultsTextBox.Location = New-Object System.Drawing.Point(10, 10)
    $resultsForm.Controls.Add($resultsTextBox)

    $results | ForEach-Object {
        $resultsTextBox.AppendText("Email: $($_.Email)`r`n")
        $resultsTextBox.AppendText("SAM Account Name: $($_.SAMAccountName)`r`n")
        $resultsTextBox.AppendText("Permissions: $($_.Permissions)`r`n")
        $resultsTextBox.AppendText("Access Limited By: $($_.AccessLimitedBy)`r`n")
        $resultsTextBox.AppendText("`r`n")
    }

    $resultsForm.ShowDialog()
})

# Show the form
$form.ShowDialog()