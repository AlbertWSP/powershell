Add-Type -AssemblyName System.Windows.Forms, System.Drawing
Import-Module ActiveDirectory -ErrorAction Stop

$script:LastReportPath = $null
$script:ProcessingResults = @()

# ==================== HELPER FUNCTIONS ====================
function Resolve-ADPrincipal {
    param([string]$Identity)
    try {
        $principal = Get-ADUser -Identity $Identity -Properties SID, DisplayName -ErrorAction Stop
        return @{ Identity = $principal.DistinguishedName; SID = $principal.SID; DisplayName = $principal.DisplayName; Type = "User" }
    }
    catch {
        try {
            $principal = Get-ADComputer -Identity $Identity -Properties ObjectSid -ErrorAction Stop
            return @{ Identity = $principal.DistinguishedName; SID = $principal.ObjectSid; DisplayName = $principal.Name; Type = "Computer" }
        }
        catch { throw "Cannot resolve: $Identity" }
    }
}

function Add-GroupOwnerPermission {
    param([string]$GroupDN, [System.Security.Principal.SecurityIdentifier]$OwnerSID, [switch]$WhatIf)
    try {
        $aclPath = "AD:\$GroupDN"
        $acl = Get-Acl $aclPath
        $guidMemberAttr = [Guid]'bf9679c0-0de6-11d0-a285-00aa003049e2'
        $rights = [System.DirectoryServices.ActiveDirectoryRights]::WriteProperty
        $rule = New-Object System.DirectoryServices.ActiveDirectoryAccessRule ($OwnerSID, $rights, "Allow", $guidMemberAttr, "None")
        
        $exists = $acl.Access | Where-Object { $_.IdentityReference -eq $OwnerSID -and $_.ObjectType -eq $guidMemberAttr }
        if ($exists) { return @{ Added = $false; Reason = "ACE already exists" } }
        
        if (-not $WhatIf) { $acl.AddAccessRule($rule); Set-Acl -Path $aclPath -AclObject $acl }
        return @{ Added = $true; Reason = "Success" }
    }
    catch { return @{ Added = $false; Reason = $_.Exception.Message } }
}

function Process-GroupOwnerChange {
    param([string]$GroupEmail, [string]$OwnerIdentity, [switch]$WhatIf)
    try {
        $group = Get-ADGroup -Filter "Mail -eq '$GroupEmail'" -Properties ManagedBy -ErrorAction Stop
        if (-not $group) {
            return @{
                GroupEmail = $GroupEmail; GroupName = "N/A"; Owner = "N/A"; Status = "Error"
                Error = "Group not found"; ManagedBySet = $false; AceAdded = $false; Timestamp = Get-Date -Format "yyyy-MM-dd HH:mm:ss"
            }
        }
        
        $owner = Resolve-ADPrincipal -Identity $OwnerIdentity
        if (-not $WhatIf) { Set-ADGroup -Identity $group.DistinguishedName -Replace @{ managedBy = $owner.Identity } -ErrorAction Stop }
        
        $aceResult = Add-GroupOwnerPermission -GroupDN $group.DistinguishedName -OwnerSID $owner.SID -WhatIf:$WhatIf
        $status = if ($aceResult.Added -or $aceResult.Reason -eq "ACE already exists") { "Success" } else { "Partial" }
        
        return @{
            GroupEmail = $GroupEmail; GroupName = $group.Name; Owner = $owner.DisplayName
            Status = $status; Error = if ($status -eq "Partial") { $aceResult.Reason } else { "" }
            ManagedBySet = $true; AceAdded = $aceResult.Added; Timestamp = Get-Date -Format "yyyy-MM-dd HH:mm:ss"
        }
    }
    catch {
        return @{
            GroupEmail = $GroupEmail; GroupName = "N/A"; Owner = "N/A"; Status = "Error"
            Error = $_.Exception.Message; ManagedBySet = $false; AceAdded = $false; Timestamp = Get-Date -Format "yyyy-MM-dd HH:mm:ss"
        }
    }
}

function Search-DLGroupsByOwner {
    param([string]$OwnerSAM)
    try {
        $owner = Resolve-ADPrincipal -Identity $OwnerSAM
        $ownerDN = $owner.Identity
        
        $groups = Get-ADGroup -Filter "ManagedBy -eq '$ownerDN'" -Properties Mail, ManagedBy -ErrorAction Stop
        
        $results = @()
        if ($groups) {
            $groupArray = @($groups)
            foreach ($group in $groupArray) {
                $managedByUser = if ($group.ManagedBy) {
                    try { Get-ADUser -Identity $group.ManagedBy -Properties DisplayName -ErrorAction Stop | Select-Object -ExpandProperty DisplayName }
                    catch { "N/A" }
                } else { "N/A" }
                
                $results += [PSCustomObject]@{
                    GroupEmail = if ($group.Mail) { $group.Mail } else { "N/A" }
                    GroupName = $group.Name
                    Owner = $managedByUser
                    Status = "Found"
                    Error = ""
                }
            }
        }
        
        return $results
    }
    catch {
        return @([PSCustomObject]@{
            GroupEmail = "N/A"
            GroupName = "N/A"
            Owner = "N/A"
            Status = "Error"
            Error = $_.Exception.Message
        })
    }
}

function Export-OperationReport {
    param([string]$FilePath, [array]$Results)
    try {
        $reportContent = @()
        $reportContent += "DL Group Owner Manager - Operation Report"
        $reportContent += "Generated: $(Get-Date -Format 'yyyy-MM-dd HH:mm:ss')"
        $reportContent += "=" * 80
        $reportContent += ""
        
        foreach ($result in $Results) {
            $reportContent += "Group Email: $($result.GroupEmail)"
            $reportContent += "Group Name: $($result.GroupName)"
            $reportContent += "Owner: $($result.Owner)"
            $reportContent += "Status: $($result.Status)"
            if ($result.Error) { $reportContent += "Error: $($result.Error)" }
            $reportContent += "-" * 40
        }
        
        $reportContent += ""
        $successCount = ($Results | Where-Object { $_.Status -eq "Success" }).Count
        $errorCount = ($Results | Where-Object { $_.Status -eq "Error" }).Count
        $reportContent += "Summary: Success=$successCount, Error=$errorCount, Total=$($Results.Count)"
        
        $reportContent | Out-File -FilePath $FilePath -Encoding UTF8 -Force
        return $true
    }
    catch {
        return $false
    }
}

# ==================== UI ====================
$form = New-Object System.Windows.Forms.Form
$form.Text = "DL Group Owner Manager"
$form.Size = New-Object System.Drawing.Size(1150, 950)
$form.StartPosition = "CenterScreen"
$form.MaximizeBox = $false
$form.Font = New-Object System.Drawing.Font("Segoe UI", 10)

# Tab Control
$tabControl = New-Object System.Windows.Forms.TabControl
$tabControl.Dock = [System.Windows.Forms.DockStyle]::Fill
$form.Controls.Add($tabControl)

# ==================== TAB 1: BATCH CHANGE ====================
$tabBatchChange = New-Object System.Windows.Forms.TabPage
$tabBatchChange.Text = "Batch Change"
$tabControl.TabPages.Add($tabBatchChange)

# Input Panel
$panelInput = New-Object System.Windows.Forms.Panel
$panelInput.BackColor = [System.Drawing.Color]::LightGray
$panelInput.Dock = [System.Windows.Forms.DockStyle]::Top
$panelInput.Height = 180
$tabBatchChange.Controls.Add($panelInput)

# Text File Path
$labelTextFile = New-Object System.Windows.Forms.Label
$labelTextFile.Text = "Text File (group emails, one per line):"
$labelTextFile.Location = New-Object System.Drawing.Point(10, 15)
$labelTextFile.Size = New-Object System.Drawing.Size(300, 20)
$labelTextFile.Font = New-Object System.Drawing.Font("Segoe UI", 10, [System.Drawing.FontStyle]::Bold)
$panelInput.Controls.Add($labelTextFile)

$textFilePath = New-Object System.Windows.Forms.TextBox
$textFilePath.Location = New-Object System.Drawing.Point(10, 40)
$textFilePath.Size = New-Object System.Drawing.Size(600, 25)
$textFilePath.Text = "C:\Temp\GroupEmails.txt"
$panelInput.Controls.Add($textFilePath)

$btnBrowseFile = New-Object System.Windows.Forms.Button
$btnBrowseFile.Text = "Browse"
$btnBrowseFile.Location = New-Object System.Drawing.Point(620, 40)
$btnBrowseFile.Size = New-Object System.Drawing.Size(80, 25)
$panelInput.Controls.Add($btnBrowseFile)

$btnLoadFile = New-Object System.Windows.Forms.Button
$btnLoadFile.Text = "Load"
$btnLoadFile.Location = New-Object System.Drawing.Point(710, 40)
$btnLoadFile.Size = New-Object System.Drawing.Size(80, 25)
$panelInput.Controls.Add($btnLoadFile)

# Owner sAMAccountName
$labelOwner = New-Object System.Windows.Forms.Label
$labelOwner.Text = "Owner sAMAccountName:"
$labelOwner.Location = New-Object System.Drawing.Point(10, 80)
$labelOwner.Size = New-Object System.Drawing.Size(300, 20)
$labelOwner.Font = New-Object System.Drawing.Font("Segoe UI", 10, [System.Drawing.FontStyle]::Bold)
$panelInput.Controls.Add($labelOwner)

$textOwnerSam = New-Object System.Windows.Forms.TextBox
$textOwnerSam.Location = New-Object System.Drawing.Point(10, 105)
$textOwnerSam.Size = New-Object System.Drawing.Size(600, 25)
$textOwnerSam.Text = "hkxxxxxxxx"
$panelInput.Controls.Add($textOwnerSam)

$btnValidateOwner = New-Object System.Windows.Forms.Button
$btnValidateOwner.Text = "Validate"
$btnValidateOwner.Location = New-Object System.Drawing.Point(620, 105)
$btnValidateOwner.Size = New-Object System.Drawing.Size(80, 25)
$panelInput.Controls.Add($btnValidateOwner)

# Status labels
$labelFileStatus = New-Object System.Windows.Forms.Label
$labelFileStatus.Text = ""
$labelFileStatus.Location = New-Object System.Drawing.Point(800, 40)
$labelFileStatus.Size = New-Object System.Drawing.Size(120, 25)
$panelInput.Controls.Add($labelFileStatus)

$labelOwnerStatus = New-Object System.Windows.Forms.Label
$labelOwnerStatus.Text = ""
$labelOwnerStatus.Location = New-Object System.Drawing.Point(800, 105)
$labelOwnerStatus.Size = New-Object System.Drawing.Size(120, 25)
$panelInput.Controls.Add($labelOwnerStatus)

# Results Output TextBox (Simple Column Display)
$labelResults = New-Object System.Windows.Forms.Label
$labelResults.Text = "Results:"
$labelResults.Location = New-Object System.Drawing.Point(10, 200)
$labelResults.Size = New-Object System.Drawing.Size(100, 20)
$labelResults.Font = New-Object System.Drawing.Font("Segoe UI", 10, [System.Drawing.FontStyle]::Bold)
$tabBatchChange.Controls.Add($labelResults)

$textResults = New-Object System.Windows.Forms.TextBox
$textResults.Multiline = $true
$textResults.ScrollBars = [System.Windows.Forms.ScrollBars]::Both
$textResults.ReadOnly = $true
$textResults.Location = New-Object System.Drawing.Point(10, 225)
$textResults.Size = New-Object System.Drawing.Size(920, 450)
$textResults.Font = New-Object System.Drawing.Font("Courier New", 9)
$tabBatchChange.Controls.Add($textResults)

# Buttons Panel for Batch Change
$panelBatchButtons = New-Object System.Windows.Forms.Panel
$panelBatchButtons.Dock = [System.Windows.Forms.DockStyle]::Bottom
$panelBatchButtons.Height = 50
$panelBatchButtons.BackColor = [System.Drawing.Color]::LightGray
$tabBatchChange.Controls.Add($panelBatchButtons)

$btnPreview = New-Object System.Windows.Forms.Button
$btnPreview.Text = "Preview (WhatIf)"
$btnPreview.Location = New-Object System.Drawing.Point(10, 10)
$btnPreview.Size = New-Object System.Drawing.Size(120, 30)
$panelBatchButtons.Controls.Add($btnPreview)

$btnExecute = New-Object System.Windows.Forms.Button
$btnExecute.Text = "Execute"
$btnExecute.Location = New-Object System.Drawing.Point(140, 10)
$btnExecute.Size = New-Object System.Drawing.Size(100, 30)
$panelBatchButtons.Controls.Add($btnExecute)

$btnExportReport = New-Object System.Windows.Forms.Button
$btnExportReport.Text = "Export Report"
$btnExportReport.Location = New-Object System.Drawing.Point(250, 10)
$btnExportReport.Size = New-Object System.Drawing.Size(120, 30)
$btnExportReport.Enabled = $false
$panelBatchButtons.Controls.Add($btnExportReport)

# ==================== TAB 2: SEARCH GROUPS ====================
$tabSearchGroups = New-Object System.Windows.Forms.TabPage
$tabSearchGroups.Text = "Search Groups"
$tabControl.TabPages.Add($tabSearchGroups)

# Search Input Panel
$panelSearchInput = New-Object System.Windows.Forms.Panel
$panelSearchInput.BackColor = [System.Drawing.Color]::LightGray
$panelSearchInput.Dock = [System.Windows.Forms.DockStyle]::Top
$panelSearchInput.Height = 100
$tabSearchGroups.Controls.Add($panelSearchInput)

$labelSearchOwner = New-Object System.Windows.Forms.Label
$labelSearchOwner.Text = "Owner sAMAccountName:"
$labelSearchOwner.Location = New-Object System.Drawing.Point(10, 15)
$labelSearchOwner.Size = New-Object System.Drawing.Size(300, 20)
$labelSearchOwner.Font = New-Object System.Drawing.Font("Segoe UI", 10, [System.Drawing.FontStyle]::Bold)
$panelSearchInput.Controls.Add($labelSearchOwner)

$textSearchOwnerSam = New-Object System.Windows.Forms.TextBox
$textSearchOwnerSam.Location = New-Object System.Drawing.Point(10, 40)
$textSearchOwnerSam.Size = New-Object System.Drawing.Size(600, 25)
$panelSearchInput.Controls.Add($textSearchOwnerSam)

$btnSearch = New-Object System.Windows.Forms.Button
$btnSearch.Text = "Search"
$btnSearch.Location = New-Object System.Drawing.Point(620, 40)
$btnSearch.Size = New-Object System.Drawing.Size(80, 25)
$panelSearchInput.Controls.Add($btnSearch)

$labelSearchStatus = New-Object System.Windows.Forms.Label
$labelSearchStatus.Text = ""
$labelSearchStatus.Location = New-Object System.Drawing.Point(710, 40)
$labelSearchStatus.Size = New-Object System.Drawing.Size(200, 25)
$panelSearchInput.Controls.Add($labelSearchStatus)

# Search Results TextBox (Simple Column Display)
$labelSearchResults = New-Object System.Windows.Forms.Label
$labelSearchResults.Text = "Found Groups:"
$labelSearchResults.Location = New-Object System.Drawing.Point(10, 110)
$labelSearchResults.Size = New-Object System.Drawing.Size(200, 20)
$labelSearchResults.Font = New-Object System.Drawing.Font("Segoe UI", 10, [System.Drawing.FontStyle]::Bold)
$tabSearchGroups.Controls.Add($labelSearchResults)

$textSearchResults = New-Object System.Windows.Forms.TextBox
$textSearchResults.Multiline = $true
$textSearchResults.ScrollBars = [System.Windows.Forms.ScrollBars]::Both
$textSearchResults.ReadOnly = $true
$textSearchResults.Location = New-Object System.Drawing.Point(10, 135)
$textSearchResults.Size = New-Object System.Drawing.Size(920, 550)
$textSearchResults.Font = New-Object System.Drawing.Font("Courier New", 9)
$tabSearchGroups.Controls.Add($textSearchResults)

# ==================== FORM BUTTONS ====================
$panelFormButtons = New-Object System.Windows.Forms.Panel
$panelFormButtons.Dock = [System.Windows.Forms.DockStyle]::Bottom
$panelFormButtons.Height = 50
$panelFormButtons.BackColor = [System.Drawing.Color]::LightGray

$btnExit = New-Object System.Windows.Forms.Button
$btnExit.Text = "Exit"
$btnExit.Location = New-Object System.Drawing.Point(860, 10)
$btnExit.Size = New-Object System.Drawing.Size(70, 30)
$panelFormButtons.Controls.Add($btnExit)

$form.Controls.Add($panelFormButtons)

# ==================== EVENT HANDLERS ====================

$btnBrowseFile.Add_Click({
    $openFileDialog = New-Object System.Windows.Forms.OpenFileDialog
    $openFileDialog.Filter = "Text Files (*.txt)|*.txt|All Files (*.*)|*.*"
    $openFileDialog.InitialDirectory = "C:\Temp"
    if ($openFileDialog.ShowDialog() -eq [System.Windows.Forms.DialogResult]::OK) {
        $textFilePath.Text = $openFileDialog.FileName
    }
})

$btnLoadFile.Add_Click({
    $filePath = $textFilePath.Text.Trim()
    if (-not (Test-Path $filePath)) {
        [System.Windows.Forms.MessageBox]::Show("File not found: $filePath", "Error", "OK", "Error")
        return
    }
    try {
        $emails = Get-Content -Path $filePath | Where-Object { $_ -match '\S' } | ForEach-Object { $_.Trim() }
        $labelFileStatus.Text = "✓ Loaded"; $labelFileStatus.ForeColor = [System.Drawing.Color]::Green
        [System.Windows.Forms.MessageBox]::Show("Loaded: $($emails.Count) emails", "Success", "OK", "Information")
    }
    catch {
        [System.Windows.Forms.MessageBox]::Show("Error: $_", "Error", "OK", "Error")
        $labelFileStatus.Text = "✗ Error"; $labelFileStatus.ForeColor = [System.Drawing.Color]::Red
    }
})

$btnValidateOwner.Add_Click({
    $ownerInput = $textOwnerSam.Text.Trim()
    try {
        $owner = Resolve-ADPrincipal -Identity $ownerInput
        $labelOwnerStatus.Text = "✓ Valid"; $labelOwnerStatus.ForeColor = [System.Drawing.Color]::Green
        [System.Windows.Forms.MessageBox]::Show("Resolved: $($owner.DisplayName)", "Valid", "OK", "Information")
    }
    catch {
        $labelOwnerStatus.Text = "✗ Invalid"; $labelOwnerStatus.ForeColor = [System.Drawing.Color]::Red
        [System.Windows.Forms.MessageBox]::Show("Error: $_", "Invalid", "OK", "Error")
    }
})

$btnPreview.Add_Click({
    if ($labelFileStatus.Text -ne "✓ Loaded") {
        [System.Windows.Forms.MessageBox]::Show("Load file first", "Validation", "OK", "Warning"); return
    }
    if ($labelOwnerStatus.Text -ne "✓ Valid") {
        [System.Windows.Forms.MessageBox]::Show("Validate owner first", "Validation", "OK", "Warning"); return
    }
    try {
        $filePath = $textFilePath.Text.Trim()
        $ownerInput = $textOwnerSam.Text.Trim()
        $emails = Get-Content -Path $filePath | Where-Object { $_ -match '\S' } | ForEach-Object { $_.Trim() }
        $results = @()
        foreach ($email in $emails) {
            $results += Process-GroupOwnerChange -GroupEmail $email -OwnerIdentity $ownerInput -WhatIf
        }
        $script:ProcessingResults = $results
        $textResults.Text = $results | Format-Table -AutoSize | Out-String
        [System.Windows.Forms.MessageBox]::Show("Preview: $($results.Count) groups", "Complete", "OK", "Information")
    }
    catch {
        [System.Windows.Forms.MessageBox]::Show("Error: $_", "Error", "OK", "Error")
    }
})

$btnExecute.Add_Click({
    if ($labelFileStatus.Text -ne "✓ Loaded") {
        [System.Windows.Forms.MessageBox]::Show("Load file first", "Validation", "OK", "Warning"); return
    }
    if ($labelOwnerStatus.Text -ne "✓ Valid") {
        [System.Windows.Forms.MessageBox]::Show("Validate owner first", "Validation", "OK", "Warning"); return
    }
    
    $saveDialog = New-Object System.Windows.Forms.SaveFileDialog
    $saveDialog.Filter = "Text Files (*.txt)|*.txt"
    $saveDialog.FileName = "DL_Owner_Change_$(Get-Date -Format 'yyyyMMdd_HHmmss').txt"
    $saveDialog.InitialDirectory = $env:USERPROFILE
    
    if ($saveDialog.ShowDialog() -ne "OK") {
        [System.Windows.Forms.MessageBox]::Show("Export cancelled", "Cancelled", "OK", "Information")
        return
    }
    
    if ([System.Windows.Forms.MessageBox]::Show("Execute changes?", "Confirm", "YesNo", "Question") -ne "Yes") { return }
    
    try {
        $filePath = $textFilePath.Text.Trim()
        $ownerInput = $textOwnerSam.Text.Trim()
        $emails = Get-Content -Path $filePath | Where-Object { $_ -match '\S' } | ForEach-Object { $_.Trim() }
        $script:ProcessingResults = @()
        foreach ($email in $emails) {
            $script:ProcessingResults += Process-GroupOwnerChange -GroupEmail $email -OwnerIdentity $ownerInput
        }
        $textResults.Text = $script:ProcessingResults | Format-Table -AutoSize | Out-String
        $btnExportReport.Enabled = $true
        
        Export-OperationReport -FilePath $saveDialog.FileName -Results $script:ProcessingResults
        $script:LastReportPath = $saveDialog.FileName
        
        $successCount = ($script:ProcessingResults | Where-Object { $_.Status -eq "Success" }).Count
        [System.Windows.Forms.MessageBox]::Show("Done!`nSuccess: $successCount / $($script:ProcessingResults.Count)`nReport: $($saveDialog.FileName)", "Complete", "OK", "Information")
    }
    catch {
        [System.Windows.Forms.MessageBox]::Show("Error: $_", "Error", "OK", "Error")
    }
})

$btnExportReport.Add_Click({
    if (-not $script:ProcessingResults -or $script:ProcessingResults.Count -eq 0) {
        [System.Windows.Forms.MessageBox]::Show("No results to export", "Warning", "OK", "Warning")
        return
    }
    
    $saveDialog = New-Object System.Windows.Forms.SaveFileDialog
    $saveDialog.Filter = "Text Files (*.txt)|*.txt"
    $saveDialog.FileName = "DL_Owner_Change_$(Get-Date -Format 'yyyyMMdd_HHmmss').txt"
    $saveDialog.InitialDirectory = $env:USERPROFILE
    
    if ($saveDialog.ShowDialog() -eq "OK") {
        Export-OperationReport -FilePath $saveDialog.FileName -Results $script:ProcessingResults
        $script:LastReportPath = $saveDialog.FileName
        [System.Windows.Forms.MessageBox]::Show("Exported: $($saveDialog.FileName)", "Success", "OK", "Information")
    }
})

$btnSearch.Add_Click({
    $searchOwner = $textSearchOwnerSam.Text.Trim()
    if (-not $searchOwner) {
        [System.Windows.Forms.MessageBox]::Show("Enter Owner sAMAccountName", "Validation", "OK", "Warning")
        return
    }
    
    try {
        $labelSearchStatus.Text = "Searching..."
        $labelSearchStatus.ForeColor = [System.Drawing.Color]::Blue
        $form.Refresh()
        
        $results = Search-DLGroupsByOwner -OwnerSAM $searchOwner
        
        if ($results -and $results.Count -gt 0) {
            $textSearchResults.Text = $results | Format-Table -AutoSize | Out-String
            $labelSearchStatus.Text = "✓ Found: $($results.Count) groups"
            $labelSearchStatus.ForeColor = [System.Drawing.Color]::Green
        } else {
            $textSearchResults.Text = ""
            $labelSearchStatus.Text = "No groups found"
            $labelSearchStatus.ForeColor = [System.Drawing.Color]::Orange
        }
    }
    catch {
        [System.Windows.Forms.MessageBox]::Show("Error: $_", "Error", "OK", "Error")
        $labelSearchStatus.Text = "✗ Error"
        $labelSearchStatus.ForeColor = [System.Drawing.Color]::Red
    }
})

$btnExit.Add_Click({ $form.Close() })

$null = $form.ShowDialog()
