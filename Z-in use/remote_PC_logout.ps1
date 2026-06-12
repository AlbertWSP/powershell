Add-Type -AssemblyName System.Windows.Forms
Add-Type -AssemblyName System.Drawing

# ==============================================================================
#  Remote Session Manager v3 - WMI/CIM Edition
#  Uses WMI/CIM to scan sessions & logoff users (No query.exe / qwinsta needed)
# ==============================================================================

function Get-RemoteSessions {
    param([string[]]$ComputerNames)

    $allSessions = @()
    $dash = [string][char]0x2014

    foreach ($pc in $ComputerNames) {
        $pc = $pc.Trim()
        if ($pc -eq "") { continue }

        if (!(Test-Connection -ComputerName $pc -Count 1 -Quiet)) {
            $allSessions += [PSCustomObject]@{
                Host      = $pc
                Username  = $dash
                SessionID = $dash
                LogonType = $dash
                State     = "OFFLINE"
                LogonTime = $dash
            }
            continue
        }

        try {
            $sessions = Get-CimInstance -ClassName Win32_LogonSession -ComputerName $pc |
            Where-Object { $_.LogonType -in 2, 10 }

            if (!$sessions) {
                $allSessions += [PSCustomObject]@{
                    Host      = $pc
                    Username  = $dash
                    SessionID = $dash
                    LogonType = $dash
                    State     = "No active sessions"
                    LogonTime = $dash
                }
                continue
            }

            $seenUsers = @{}

            foreach ($session in $sessions) {
                $userLogged = Get-CimInstance -ClassName Win32_LoggedOnUser -ComputerName $pc |
                Where-Object { $_.Dependent.LogonId -eq $session.LogonId }
                $username = if ($userLogged) { $userLogged.Antecedent.Name } else { "Unknown" }

                if ($username -in "SYSTEM", "LOCAL SERVICE", "NETWORK SERVICE", "UMFD-0", "UMFD-1", "DWM-1", "DWM-2", "Unknown", "") { continue }

                $key = "$pc|$username"
                if ($seenUsers.ContainsKey($key)) { continue }
                $seenUsers[$key] = $true

                $logonTypeMap = @{ 2 = "Console"; 10 = "RDP" }
                $logonTypeStr = if ($logonTypeMap.ContainsKey([int]$session.LogonType)) {
                    $logonTypeMap[[int]$session.LogonType]
                }
                else { "Type $($session.LogonType)" }

                $logonTime = if ($session.StartTime) {
                    $session.StartTime.ToString("yyyy-MM-dd HH:mm:ss")
                }
                else { $dash }

                $allSessions += [PSCustomObject]@{
                    Host      = $pc
                    Username  = $username
                    SessionID = $session.LogonId
                    LogonType = $logonTypeStr
                    State     = "Active"
                    LogonTime = $logonTime
                }
            }

            if ($seenUsers.Count -eq 0) {
                $allSessions += [PSCustomObject]@{
                    Host      = $pc
                    Username  = $dash
                    SessionID = $dash
                    LogonType = $dash
                    State     = "No active sessions"
                    LogonTime = $dash
                }
            }
        }
        catch {
            $allSessions += [PSCustomObject]@{
                Host      = $pc
                Username  = $dash
                SessionID = $dash
                LogonType = $dash
                State     = "ERROR: $($_.Exception.Message)"
                LogonTime = $dash
            }
        }
    }

    return , $allSessions
}

function Invoke-RemoteLogoff {
    param(
        [string]$ComputerName,
        [string]$Username
    )
    try {
        $processes = Get-CimInstance -ClassName Win32_Process -ComputerName $ComputerName -Filter "Name='explorer.exe'"
        $found = $false
        foreach ($proc in $processes) {
            $owner = Invoke-CimMethod -InputObject $proc -MethodName GetOwner
            if ($owner.User -eq $Username) {
                Invoke-CimMethod -InputObject $proc -MethodName Terminate | Out-Null
                $found = $true
            }
        }
        return $found
    }
    catch {
        return $false
    }
}

$form = New-Object System.Windows.Forms.Form
$form.Text = "Remote Session Manager v3 (WMI Edition)"
$form.Size = New-Object System.Drawing.Size(780, 640)
$form.StartPosition = "CenterScreen"
$form.FormBorderStyle = "FixedDialog"
$form.MaximizeBox = $false

$labelHost = New-Object System.Windows.Forms.Label
$labelHost.Text = "Enter Host Names (one per line):"
$labelHost.Location = New-Object System.Drawing.Point(10, 10)
$labelHost.Size = New-Object System.Drawing.Size(400, 20)
$form.Controls.Add($labelHost)

$textBoxHosts = New-Object System.Windows.Forms.TextBox
$textBoxHosts.Location = New-Object System.Drawing.Point(10, 35)
$textBoxHosts.Size = New-Object System.Drawing.Size(560, 100)
$textBoxHosts.Multiline = $true
$textBoxHosts.ScrollBars = "Vertical"
$textBoxHosts.Font = New-Object System.Drawing.Font("Consolas", 10)
$form.Controls.Add($textBoxHosts)

$btnScan = New-Object System.Windows.Forms.Button
$btnScan.Text = "Scan"
$btnScan.Size = New-Object System.Drawing.Size(170, 35)
$btnScan.Location = New-Object System.Drawing.Point(580, 35)
$btnScan.Font = New-Object System.Drawing.Font("Segoe UI", 10, [System.Drawing.FontStyle]::Bold)
$form.Controls.Add($btnScan)

$btnRefresh = New-Object System.Windows.Forms.Button
$btnRefresh.Text = "Refresh"
$btnRefresh.Size = New-Object System.Drawing.Size(170, 30)
$btnRefresh.Location = New-Object System.Drawing.Point(580, 75)
$form.Controls.Add($btnRefresh)

$dataGrid = New-Object System.Windows.Forms.DataGridView
$dataGrid.Location = New-Object System.Drawing.Point(10, 145)
$dataGrid.Size = New-Object System.Drawing.Size(740, 350)
$dataGrid.AllowUserToAddRows = $false
$dataGrid.AllowUserToDeleteRows = $false
$dataGrid.ReadOnly = $false
$dataGrid.SelectionMode = "FullRowSelect"
$dataGrid.RowHeadersVisible = $false
$dataGrid.AutoSizeColumnsMode = "Fill"
$dataGrid.Font = New-Object System.Drawing.Font("Segoe UI", 9)
$dataGrid.BackgroundColor = [System.Drawing.Color]::White
$dataGrid.DefaultCellStyle.SelectionBackColor = [System.Drawing.Color]::LightSteelBlue
$dataGrid.DefaultCellStyle.SelectionForeColor = [System.Drawing.Color]::Black

$colSelect = New-Object System.Windows.Forms.DataGridViewCheckBoxColumn
$colSelect.Name = "Select"
$colSelect.HeaderText = [string][char]0x2713
$colSelect.Width = 35
$colSelect.ReadOnly = $false
$dataGrid.Columns.Add($colSelect) | Out-Null

foreach ($col in @(
        @{N = "Host"; W = 18 },
        @{N = "Username"; W = 18 },
        @{N = "SessionID"; W = 10; H = "Session ID" },
        @{N = "LogonType"; W = 10; H = "Type" },
        @{N = "State"; W = 12 },
        @{N = "LogonTime"; W = 20; H = "Logon Time" }
    )) {
    $c = New-Object System.Windows.Forms.DataGridViewTextBoxColumn
    $c.Name = $col.N
    $c.HeaderText = if ($col.H) { $col.H } else { $col.N }
    $c.ReadOnly = $true
    $c.FillWeight = $col.W
    $dataGrid.Columns.Add($c) | Out-Null
}

$form.Controls.Add($dataGrid)

$btnSelectAll = New-Object System.Windows.Forms.Button
$btnSelectAll.Text = "Select All"
$btnSelectAll.Size = New-Object System.Drawing.Size(100, 30)
$btnSelectAll.Location = New-Object System.Drawing.Point(10, 505)
$form.Controls.Add($btnSelectAll)

$btnDeselectAll = New-Object System.Windows.Forms.Button
$btnDeselectAll.Text = "Deselect All"
$btnDeselectAll.Size = New-Object System.Drawing.Size(100, 30)
$btnDeselectAll.Location = New-Object System.Drawing.Point(115, 505)
$form.Controls.Add($btnDeselectAll)

$btnLogoff = New-Object System.Windows.Forms.Button
$btnLogoff.Text = "Logoff Selected"
$btnLogoff.Size = New-Object System.Drawing.Size(160, 35)
$btnLogoff.Location = New-Object System.Drawing.Point(430, 503)
$btnLogoff.Font = New-Object System.Drawing.Font("Segoe UI", 10, [System.Drawing.FontStyle]::Bold)
$btnLogoff.BackColor = [System.Drawing.Color]::MistyRose
$form.Controls.Add($btnLogoff)

$btnClose = New-Object System.Windows.Forms.Button
$btnClose.Text = "Close"
$btnClose.Size = New-Object System.Drawing.Size(100, 35)
$btnClose.Location = New-Object System.Drawing.Point(600, 503)
$btnClose.Add_Click({ $form.Close() })
$form.Controls.Add($btnClose)

$statusBar = New-Object System.Windows.Forms.StatusStrip
$statusLabel = New-Object System.Windows.Forms.ToolStripStatusLabel
$statusLabel.Text = "Ready  |  Method: WMI/CIM  |  Enter host names and click Scan"
$statusBar.Items.Add($statusLabel) | Out-Null
$form.Controls.Add($statusBar)

function Populate-Grid {
    param($sessions)
    $dataGrid.Rows.Clear()

    foreach ($s in $sessions) {
        $rowIndex = $dataGrid.Rows.Add()
        $row = $dataGrid.Rows[$rowIndex]

        $row.Cells["Select"].Value = $false
        $row.Cells["Host"].Value = $s.Host
        $row.Cells["Username"].Value = $s.Username
        $row.Cells["SessionID"].Value = $s.SessionID
        $row.Cells["LogonType"].Value = $s.LogonType
        $row.Cells["State"].Value = $s.State
        $row.Cells["LogonTime"].Value = $s.LogonTime

        if ($s.State -eq "OFFLINE") {
            $row.DefaultCellStyle.BackColor = [System.Drawing.Color]::LightGray
            $row.DefaultCellStyle.ForeColor = [System.Drawing.Color]::DarkGray
            $row.Cells["Select"].ReadOnly = $true
        }
        elseif ($s.State -eq "No active sessions") {
            $row.DefaultCellStyle.BackColor = [System.Drawing.Color]::WhiteSmoke
            $row.Cells["Select"].ReadOnly = $true
        }
        elseif ($s.State -eq "Active") {
            $row.DefaultCellStyle.BackColor = [System.Drawing.Color]::Honeydew
        }
        elseif ($s.State -match "ERROR") {
            $row.DefaultCellStyle.BackColor = [System.Drawing.Color]::MistyRose
            $row.Cells["Select"].ReadOnly = $true
        }
    }
}

$scanAction = {
    $hosts = $textBoxHosts.Text -split "`r`n" | ForEach-Object { $_.Trim() } | Where-Object { $_ -match '\S' }

    if ($hosts.Count -eq 0) {
        [System.Windows.Forms.MessageBox]::Show("Please enter at least one host name.", "No Hosts", "OK", "Warning")
        return
    }

    $statusLabel.Text = "Scanning $($hosts.Count) host(s) via WMI/CIM... Please wait."
    $btnScan.Enabled = $false
    $btnScan.Text = "Scanning..."
    $form.Cursor = [System.Windows.Forms.Cursors]::WaitCursor
    $form.Refresh()

    $script:sessionData = Get-RemoteSessions -ComputerNames $hosts
    Populate-Grid -sessions $script:sessionData

    $totalSessions = ($script:sessionData | Where-Object { $_.State -eq "Active" }).Count
    $offlineCount = ($script:sessionData | Where-Object { $_.State -eq "OFFLINE" }).Count
    $errorCount = ($script:sessionData | Where-Object { $_.State -match "ERROR" }).Count
    $statusLabel.Text = "Found $totalSessions active session(s) across $($hosts.Count) host(s)  |  Offline: $offlineCount  |  Errors: $errorCount"

    $btnScan.Enabled = $true
    $btnScan.Text = "Scan"
    $form.Cursor = [System.Windows.Forms.Cursors]::Default
}

$btnScan.Add_Click($scanAction)
$btnRefresh.Add_Click($scanAction)

$btnSelectAll.Add_Click({
        foreach ($row in $dataGrid.Rows) {
            if (!$row.Cells["Select"].ReadOnly) { $row.Cells["Select"].Value = $true }
        }
    })

$btnDeselectAll.Add_Click({
        foreach ($row in $dataGrid.Rows) { $row.Cells["Select"].Value = $false }
    })

$btnLogoff.Add_Click({
        $selected = @()
        foreach ($row in $dataGrid.Rows) {
            if ($row.Cells["Select"].Value -eq $true) {
                $selected += [PSCustomObject]@{
                    Host     = $row.Cells["Host"].Value
                    Username = $row.Cells["Username"].Value
                }
            }
        }

        if ($selected.Count -eq 0) {
            [System.Windows.Forms.MessageBox]::Show(
                "No sessions selected.`nPlease check the boxes for sessions you want to log off.",
                "Nothing Selected", "OK", "Information"
            )
            return
        }

        $sessionList = ($selected | ForEach-Object {
                "  $([char]0x2022) $($_.Username) on $($_.Host)"
            }) -join "`n"
        $confirmMsg = "Are you sure you want to log off the following $($selected.Count) session(s)?`n`n$sessionList"
        $confirm = [System.Windows.Forms.MessageBox]::Show($confirmMsg, "Confirm Logoff", "YesNo", "Warning")

        if ($confirm -ne "Yes") { return }

        $form.Cursor = [System.Windows.Forms.Cursors]::WaitCursor
        $statusLabel.Text = "Logging off $($selected.Count) session(s)..."
        $form.Refresh()

        $logoffSuccess = 0
        $logoffFail = 0

        foreach ($s in $selected) {
            $statusLabel.Text = "Logging off $($s.Username) on $($s.Host)..."
            $form.Refresh()

            $success = Invoke-RemoteLogoff -ComputerName $s.Host -Username $s.Username
            if ($success) { $logoffSuccess++ } else { $logoffFail++ }
        }

        $form.Cursor = [System.Windows.Forms.Cursors]::Default
        $statusLabel.Text = "Logoff complete. Success: $logoffSuccess | Failed: $logoffFail"

        [System.Windows.Forms.MessageBox]::Show(
            "Logoff Results:`n`n  Success: $logoffSuccess`n  Failed: $logoffFail`n`nClick Refresh to update the session list.",
            "Logoff Complete", "OK", "Information"
        )
    })

$form.CancelButton = $btnClose
[void]$form.ShowDialog()
$form.Dispose()
