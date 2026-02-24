Add-Type -AssemblyName System.Windows.Forms

Add-Type -AssemblyName System.Drawing

  

# Check if ActiveDirectory module is available

try {

    Import-Module ActiveDirectory -ErrorAction Stop

}

catch {

    [System.Windows.Forms.MessageBox]::Show("Active Directory module is required but not found or could not be loaded. Some features may not work.", "Warning", [System.Windows.Forms.MessageBoxButtons]::OK, [System.Windows.Forms.MessageBoxIcon]::Warning)

}

  

# Create Main Form

$form = New-Object System.Windows.Forms.Form

$form.Text = "Folder Access Rights Checker"

$form.Size = New-Object System.Drawing.Size(800, 600)

$form.StartPosition = "CenterScreen"

$form.FormBorderStyle = "FixedDialog"

$form.MaximizeBox = $false

  

# Fonts

$fontRegular = New-Object System.Drawing.Font("Segoe UI", 9)

$fontBold = New-Object System.Drawing.Font("Segoe UI", 9, [System.Drawing.FontStyle]::Bold)

  

# --- Path Selection Section ---

$lblPath = New-Object System.Windows.Forms.Label

$lblPath.Text = "Folder Path:"

$lblPath.Location = New-Object System.Drawing.Point(20, 20)

$lblPath.AutoSize = $true

$lblPath.Font = $fontBold

$form.Controls.Add($lblPath)

  

$txtPath = New-Object System.Windows.Forms.TextBox

$txtPath.Location = New-Object System.Drawing.Point(20, 45)

$txtPath.Size = New-Object System.Drawing.Size(550, 25)

$txtPath.Font = $fontRegular

$form.Controls.Add($txtPath)

  

$btnBrowse = New-Object System.Windows.Forms.Button

$btnBrowse.Text = "Browse..."

$btnBrowse.Location = New-Object System.Drawing.Point(580, 43)

$btnBrowse.Size = New-Object System.Drawing.Size(80, 27)

$btnBrowse.Font = $fontRegular

$form.Controls.Add($btnBrowse)

  

$btnGetGroups = New-Object System.Windows.Forms.Button

$btnGetGroups.Text = "Get Access Rights"

$btnGetGroups.Location = New-Object System.Drawing.Point(670, 43)

$btnGetGroups.Size = New-Object System.Drawing.Size(100, 27)

$btnGetGroups.Font = $fontRegular

$form.Controls.Add($btnGetGroups)

  

# --- Group Access Grid ---

$gridGroups = New-Object System.Windows.Forms.DataGridView

$gridGroups.Location = New-Object System.Drawing.Point(20, 90)

$gridGroups.Size = New-Object System.Drawing.Size(750, 300)

$gridGroups.ReadOnly = $true

$gridGroups.AllowUserToAddRows = $false

$gridGroups.RowHeadersVisible = $false

$gridGroups.SelectionMode = "FullRowSelect"

$gridGroups.AutoSizeColumnsMode = "Fill"

$gridGroups.Font = $fontRegular

  

$gridGroups.Columns.Add("Identity", "Identity") | Out-Null

$gridGroups.Columns.Add("Type", "Access Type") | Out-Null

$gridGroups.Columns.Add("Rights", "Rights") | Out-Null

$gridGroups.Columns.Add("Inherited", "Inherited") | Out-Null

$gridGroups.Columns.Add("ADType", "AD Object Type") | Out-Null

  

$form.Controls.Add($gridGroups)

  

# --- User Check Section ---

$grpUserCheck = New-Object System.Windows.Forms.GroupBox

$grpUserCheck.Text = "Check Individual User Access"

$grpUserCheck.Location = New-Object System.Drawing.Point(20, 410)

$grpUserCheck.Size = New-Object System.Drawing.Size(750, 130)

$grpUserCheck.Font = $fontBold

$form.Controls.Add($grpUserCheck)

  

$lblEmail = New-Object System.Windows.Forms.Label

$lblEmail.Text = "User Email:"

$lblEmail.Location = New-Object System.Drawing.Point(20, 30)

$lblEmail.AutoSize = $true

$lblEmail.Font = $fontRegular

$grpUserCheck.Controls.Add($lblEmail)

  

$txtEmail = New-Object System.Windows.Forms.TextBox

$txtEmail.Location = New-Object System.Drawing.Point(100, 27)

$txtEmail.Size = New-Object System.Drawing.Size(300, 25)

$txtEmail.Font = $fontRegular

$grpUserCheck.Controls.Add($txtEmail)

  

$btnCheckUser = New-Object System.Windows.Forms.Button

$btnCheckUser.Text = "Check User Access"

$btnCheckUser.Location = New-Object System.Drawing.Point(420, 25)

$btnCheckUser.Size = New-Object System.Drawing.Size(150, 27)

$btnCheckUser.Font = $fontRegular

$grpUserCheck.Controls.Add($btnCheckUser)

  

$lblResult = New-Object System.Windows.Forms.Label

$lblResult.Text = "Result:"

$lblResult.Location = New-Object System.Drawing.Point(20, 70)

$lblResult.AutoSize = $true

$lblResult.Font = $fontRegular

$grpUserCheck.Controls.Add($lblResult)

  

$txtResult = New-Object System.Windows.Forms.TextBox

$txtResult.Location = New-Object System.Drawing.Point(100, 67)

$txtResult.Size = New-Object System.Drawing.Size(630, 40)

$txtResult.Multiline = $true

$txtResult.ScrollBars = "Vertical"

$txtResult.ReadOnly = $true

$txtResult.Font = $fontRegular

$grpUserCheck.Controls.Add($txtResult)

  

# --- Events & Logic ---

  

$btnBrowse.Add_Click({

        $dialog = New-Object System.Windows.Forms.FolderBrowserDialog

        if ($dialog.ShowDialog() -eq [System.Windows.Forms.DialogResult]::OK) {

            $txtPath.Text = $dialog.SelectedPath

        }

    })

  

$btnGetGroups.Add_Click({

        $path = $txtPath.Text

        if ([string]::IsNullOrWhiteSpace($path) -or -not (Test-Path $path)) {

            [System.Windows.Forms.MessageBox]::Show("Please enter a valid folder path.", "Error", [System.Windows.Forms.MessageBoxButtons]::OK, [System.Windows.Forms.MessageBoxIcon]::Error)

            return

        }

  

        $gridGroups.Rows.Clear()

        $form.Cursor = [System.Windows.Forms.Cursors]::WaitCursor

  

        try {

            $acl = Get-Acl -LiteralPath $path

            foreach ($access in $acl.Access) {

                $identity = $access.IdentityReference.Value

                $rights = $access.FileSystemRights

                $type = $access.AccessControlType

                $inherited = $access.IsInherited

  

                # Try to get AD Object Type

                $adType = "Unknown/Local"

                try {

                    if ($identity -match "\\") {

                        $samAccountName = $identity.Split("\")[1]

  

                        # Try Group first

                        try {

                            $adGroup = Get-ADGroup -Identity $samAccountName -ErrorAction Stop

                            $adType = "AD Group"

                        }

                        catch {

                            # Try User

                            try {

                                $adUser = Get-ADUser -Identity $samAccountName -ErrorAction Stop

                                $adType = "AD User"

                            }

                            catch {

                                $adType = "Not Found in AD"

                            }

                        }

                    }

                }

                catch {

                    $adType = "Error Checking AD"

                }

  

                $gridGroups.Rows.Add($identity, $type, $rights, $inherited, $adType) | Out-Null

            }

        }

        catch {

            [System.Windows.Forms.MessageBox]::Show("Error reading ACL: $_", "Error", [System.Windows.Forms.MessageBoxButtons]::OK, [System.Windows.Forms.MessageBoxIcon]::Error)

        }

        finally {

            $form.Cursor = [System.Windows.Forms.Cursors]::Default

        }

    })

  

$btnCheckUser.Add_Click({

        $email = $txtEmail.Text

        $path = $txtPath.Text

  

        if ([string]::IsNullOrWhiteSpace($email)) {

            [System.Windows.Forms.MessageBox]::Show("Please enter a user email.", "Error", [System.Windows.Forms.MessageBoxButtons]::OK, [System.Windows.Forms.MessageBoxIcon]::Error)

            return

        }

        if ([string]::IsNullOrWhiteSpace($path) -or -not (Test-Path $path)) {

            [System.Windows.Forms.MessageBox]::Show("Please enter a valid folder path first.", "Error", [System.Windows.Forms.MessageBoxButtons]::OK, [System.Windows.Forms.MessageBoxIcon]::Error)

            return

        }

  

        $form.Cursor = [System.Windows.Forms.Cursors]::WaitCursor

        $txtResult.Text = "Checking..."

  

        try {

            # 1. Get User

            try {

                $adUser = Get-ADUser -Filter "EmailAddress -eq '$email'" -Properties MemberOf, EmailAddress -ErrorAction Stop

            }

            catch {

                $txtResult.Text = "Error querying AD for user: $_"

                $form.Cursor = [System.Windows.Forms.Cursors]::Default

                return

            }

  

            if (-not $adUser) {

                $txtResult.Text = "User not found in AD with email: $email"

                $form.Cursor = [System.Windows.Forms.Cursors]::Default

                return

            }

  

            # 2. Get User's Groups (Recursive)

            try {

                $userGroups = Get-ADPrincipalGroupMembership -Identity $adUser.DistinguishedName | Select-Object -ExpandProperty SID

                $userSid = $adUser.SID.Value

            }

            catch {

                $txtResult.Text = "Error retrieving group membership: $_"

                $form.Cursor = [System.Windows.Forms.Cursors]::Default

                return

            }

  

            # 3. Get ACL

            $acl = Get-Acl -LiteralPath $path

            $effectiveRights = @()

            $denied = $false

  

            foreach ($access in $acl.Access) {

                $entrySid = $null

                try {

                    $account = $access.IdentityReference.Translate([System.Security.Principal.SecurityIdentifier])

                    $entrySid = $account.Value

                }

                catch {

                    # Could not translate to SID, execute logic based on name matching if necessary,

                    # but relies on SID for AD comparison usually.

                    continue

                }

  

                if ($entrySid -eq $userSid -or $userGroups -contains $entrySid) {

                    if ($access.AccessControlType -eq "Deny") {

                        $denied = $true

                        # Technically should check if Deny overrides specific Allow, but for simple check, Deny is strong.

                    }

                    elseif ($access.AccessControlType -eq "Allow") {

                        $effectiveRights += $access.FileSystemRights

                    }

                }

            }

  

            if ($denied) {

                $txtResult.Text = "Access DENIED explicitly."

            }

            elseif ($effectiveRights.Count -eq 0) {

                $txtResult.Text = "No access rights found for this user."

            }

            else {

                # Consolidate rights string

                $rightsString = $effectiveRights -join ", "

                $txtResult.Text = "User '$($adUser.Name)' has rights: $rightsString"

            }

        }

        catch {

            $txtResult.Text = "Error: $_"

        }

        finally {

            $form.Cursor = [System.Windows.Forms.Cursors]::Default

        }

    })

  

# Show Form

$form.ShowDialog() | Out-Null

$form.Dispose()