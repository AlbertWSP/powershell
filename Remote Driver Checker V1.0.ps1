Add-Type -AssemblyName System.Windows.Forms
Add-Type -AssemblyName System.Drawing

# ══════════════════════════════════════════════════════════════════════════════
#  Remote Driver Checker v1.0
#  Check installed driver versions on remote PCs via WinRM
#  Features:
#    - Scan multiple remote PCs for installed drivers
#    - Filter by keyword (e.g. Intel, Arc, NPU, BIOS, Realtek)
#    - DataGridView with sortable columns
#    - Export results to CSV
#    - Dark theme GUI matching Setup File Deployment tool
# ══════════════════════════════════════════════════════════════════════════════

$script:allResults = @()
$script:isScanning = $false
$script:cancelScan = $false

# ── Helper: Append colored text to RichTextBox ────────────────────────────────
function Write-Log {
    param(
        [System.Windows.Forms.RichTextBox]$LogBox,
        [string]$Message,
        [System.Drawing.Color]$Color = [System.Drawing.Color]::White
    )
    $LogBox.SelectionStart = $LogBox.TextLength
    $LogBox.SelectionLength = 0
    $LogBox.SelectionColor = $Color
    $LogBox.AppendText("$Message`r`n")
    $LogBox.ScrollToCaret()
    [System.Windows.Forms.Application]::DoEvents()
}

# ══════════════════════════════════════════════════════════════════════════════
#  GUI SETUP
# ══════════════════════════════════════════════════════════════════════════════

$form = New-Object System.Windows.Forms.Form
$form.Text = "Remote Driver Checker v1.0"
$form.Size = New-Object System.Drawing.Size(920, 780)
$form.StartPosition = "CenterScreen"
$form.FormBorderStyle = "Sizable"
$form.MinimumSize = New-Object System.Drawing.Size(800, 600)
$form.BackColor = [System.Drawing.Color]::FromArgb(30, 30, 30)
$form.ForeColor = [System.Drawing.Color]::White

$yPos = 10

# ── Host Names ─────────────────────────────────────────────────────────────────

$labelHost = New-Object System.Windows.Forms.Label
$labelHost.Text = "Target Host Names (one per line):"
$labelHost.Location = New-Object System.Drawing.Point(10, $yPos)
$labelHost.Size = New-Object System.Drawing.Size(300, 20)
$form.Controls.Add($labelHost)
$yPos += 22

$textBoxHosts = New-Object System.Windows.Forms.TextBox
$textBoxHosts.Location = New-Object System.Drawing.Point(10, $yPos)
$textBoxHosts.Size = New-Object System.Drawing.Size(300, 80)
$textBoxHosts.Multiline = $true
$textBoxHosts.ScrollBars = "Vertical"
$textBoxHosts.Font = New-Object System.Drawing.Font("Consolas", 10)
$form.Controls.Add($textBoxHosts)

# ── Filter Keywords ────────────────────────────────────────────────────────────

$labelFilter = New-Object System.Windows.Forms.Label
$labelFilter.Text = "Filter Keywords (comma-separated, e.g. Intel,Arc,NPU):"
$labelFilter.Location = New-Object System.Drawing.Point(320, 10)
$labelFilter.Size = New-Object System.Drawing.Size(400, 20)
$form.Controls.Add($labelFilter)

$textBoxFilter = New-Object System.Windows.Forms.TextBox
$textBoxFilter.Location = New-Object System.Drawing.Point(320, 32)
$textBoxFilter.Size = New-Object System.Drawing.Size(400, 24)
$textBoxFilter.Font = New-Object System.Drawing.Font("Consolas", 10)
$textBoxFilter.Text = "Intel,Dell,BIOS,Graphics,NPU,Serial IO,Camera,Realtek,Audio"
$form.Controls.Add($textBoxFilter)

# ── Quick filter buttons ──────────────────────────────────────────────────────

$quickFilterY = 62

$btnFilterAll = New-Object System.Windows.Forms.Button
$btnFilterAll.Text = "All Drivers"
$btnFilterAll.Location = New-Object System.Drawing.Point(320, $quickFilterY)
$btnFilterAll.Size = New-Object System.Drawing.Size(85, 25)
$btnFilterAll.FlatStyle = "Flat"
$btnFilterAll.ForeColor = [System.Drawing.Color]::White
$btnFilterAll.Add_Click({ $textBoxFilter.Text = "" })
$form.Controls.Add($btnFilterAll)

$btnFilterIntel = New-Object System.Windows.Forms.Button
$btnFilterIntel.Text = "Intel"
$btnFilterIntel.Location = New-Object System.Drawing.Point(410, $quickFilterY)
$btnFilterIntel.Size = New-Object System.Drawing.Size(60, 25)
$btnFilterIntel.FlatStyle = "Flat"
$btnFilterIntel.ForeColor = [System.Drawing.Color]::White
$btnFilterIntel.Add_Click({ $textBoxFilter.Text = "Intel" })
$form.Controls.Add($btnFilterIntel)

$btnFilterDell = New-Object System.Windows.Forms.Button
$btnFilterDell.Text = "Dell/BIOS"
$btnFilterDell.Location = New-Object System.Drawing.Point(475, $quickFilterY)
$btnFilterDell.Size = New-Object System.Drawing.Size(75, 25)
$btnFilterDell.FlatStyle = "Flat"
$btnFilterDell.ForeColor = [System.Drawing.Color]::White
$btnFilterDell.Add_Click({ $textBoxFilter.Text = "Dell,BIOS,Firmware" })
$form.Controls.Add($btnFilterDell)

$btnFilterNetwork = New-Object System.Windows.Forms.Button
$btnFilterNetwork.Text = "Network"
$btnFilterNetwork.Location = New-Object System.Drawing.Point(555, $quickFilterY)
$btnFilterNetwork.Size = New-Object System.Drawing.Size(70, 25)
$btnFilterNetwork.FlatStyle = "Flat"
$btnFilterNetwork.ForeColor = [System.Drawing.Color]::White
$btnFilterNetwork.Add_Click({ $textBoxFilter.Text = "Network,Ethernet,Wi-Fi,Wireless,Bluetooth" })
$form.Controls.Add($btnFilterNetwork)

$btnFilterDisplay = New-Object System.Windows.Forms.Button
$btnFilterDisplay.Text = "Display"
$btnFilterDisplay.Location = New-Object System.Drawing.Point(630, $quickFilterY)
$btnFilterDisplay.Size = New-Object System.Drawing.Size(65, 25)
$btnFilterDisplay.FlatStyle = "Flat"
$btnFilterDisplay.ForeColor = [System.Drawing.Color]::White
$btnFilterDisplay.Add_Click({ $textBoxFilter.Text = "Graphics,Display,Arc,GPU,Video" })
$form.Controls.Add($btnFilterDisplay)

# ── Scan & Action Buttons ──────────────────────────────────────────────────────

$btnScan = New-Object System.Windows.Forms.Button
$btnScan.Text = [char]0x1F50D + "  Scan Drivers"
$btnScan.Location = New-Object System.Drawing.Point(730, 30)
$btnScan.Size = New-Object System.Drawing.Size(150, 34)
$btnScan.Font = New-Object System.Drawing.Font("Segoe UI", 9.5, [System.Drawing.FontStyle]::Bold)
$btnScan.BackColor = [System.Drawing.Color]::FromArgb(0, 120, 212)
$btnScan.ForeColor = [System.Drawing.Color]::White
$btnScan.FlatStyle = "Flat"
$form.Controls.Add($btnScan)

$btnExport = New-Object System.Windows.Forms.Button
$btnExport.Text = "Export CSV"
$btnExport.Location = New-Object System.Drawing.Point(730, 68)
$btnExport.Size = New-Object System.Drawing.Size(150, 28)
$btnExport.FlatStyle = "Flat"
$btnExport.BackColor = [System.Drawing.Color]::FromArgb(50, 50, 50)
$btnExport.ForeColor = [System.Drawing.Color]::White
$btnExport.Enabled = $false
$form.Controls.Add($btnExport)

$yPos = 110

# ── Progress Bar ───────────────────────────────────────────────────────────────

$progressBar = New-Object System.Windows.Forms.ProgressBar
$progressBar.Location = New-Object System.Drawing.Point(10, $yPos)
$progressBar.Size = New-Object System.Drawing.Size(780, 18)
$progressBar.Minimum = 0
$progressBar.Maximum = 100
$progressBar.Style = "Continuous"
$progressBar.Anchor = [System.Windows.Forms.AnchorStyles]::Top -bor [System.Windows.Forms.AnchorStyles]::Left -bor [System.Windows.Forms.AnchorStyles]::Right
$form.Controls.Add($progressBar)

$labelPercent = New-Object System.Windows.Forms.Label
$labelPercent.Text = ""
$labelPercent.Location = New-Object System.Drawing.Point(800, $yPos)
$labelPercent.Size = New-Object System.Drawing.Size(80, 18)
$labelPercent.ForeColor = [System.Drawing.Color]::FromArgb(0, 200, 0)
$labelPercent.Font = New-Object System.Drawing.Font("Consolas", 8.5, [System.Drawing.FontStyle]::Bold)
$labelPercent.Anchor = [System.Windows.Forms.AnchorStyles]::Top -bor [System.Windows.Forms.AnchorStyles]::Right
$form.Controls.Add($labelPercent)
$yPos += 22

$labelStatus = New-Object System.Windows.Forms.Label
$labelStatus.Text = "Ready. Enter host names and click Scan Drivers."
$labelStatus.Location = New-Object System.Drawing.Point(10, $yPos)
$labelStatus.Size = New-Object System.Drawing.Size(870, 18)
$labelStatus.ForeColor = [System.Drawing.Color]::FromArgb(180, 180, 180)
$labelStatus.Font = New-Object System.Drawing.Font("Segoe UI", 8.5)
$labelStatus.Anchor = [System.Windows.Forms.AnchorStyles]::Top -bor [System.Windows.Forms.AnchorStyles]::Left -bor [System.Windows.Forms.AnchorStyles]::Right
$form.Controls.Add($labelStatus)
$yPos += 22

# ── DataGridView ──────────────────────────────────────────────────────────────

$dataGrid = New-Object System.Windows.Forms.DataGridView
$dataGrid.Location = New-Object System.Drawing.Point(10, $yPos)
$dataGrid.Size = New-Object System.Drawing.Size(880, 400)
$dataGrid.Anchor = [System.Windows.Forms.AnchorStyles]::Top -bor [System.Windows.Forms.AnchorStyles]::Bottom -bor [System.Windows.Forms.AnchorStyles]::Left -bor [System.Windows.Forms.AnchorStyles]::Right
$dataGrid.BackgroundColor = [System.Drawing.Color]::FromArgb(25, 25, 25)
$dataGrid.ForeColor = [System.Drawing.Color]::White
$dataGrid.GridColor = [System.Drawing.Color]::FromArgb(60, 60, 60)
$dataGrid.BorderStyle = "None"
$dataGrid.CellBorderStyle = "SingleHorizontal"
$dataGrid.ColumnHeadersBorderStyle = "Single"
$dataGrid.EnableHeadersVisualStyles = $false
$dataGrid.DefaultCellStyle.BackColor = [System.Drawing.Color]::FromArgb(30, 30, 30)
$dataGrid.DefaultCellStyle.ForeColor = [System.Drawing.Color]::White
$dataGrid.DefaultCellStyle.SelectionBackColor = [System.Drawing.Color]::FromArgb(0, 100, 180)
$dataGrid.DefaultCellStyle.SelectionForeColor = [System.Drawing.Color]::White
$dataGrid.DefaultCellStyle.Font = New-Object System.Drawing.Font("Consolas", 9)
$dataGrid.ColumnHeadersDefaultCellStyle.BackColor = [System.Drawing.Color]::FromArgb(50, 50, 50)
$dataGrid.ColumnHeadersDefaultCellStyle.ForeColor = [System.Drawing.Color]::White
$dataGrid.ColumnHeadersDefaultCellStyle.Font = New-Object System.Drawing.Font("Segoe UI", 9, [System.Drawing.FontStyle]::Bold)
$dataGrid.RowHeadersVisible = $false
$dataGrid.AllowUserToAddRows = $false
$dataGrid.ReadOnly = $true
$dataGrid.SelectionMode = "FullRowSelect"
$dataGrid.MultiSelect = $true
$dataGrid.AutoSizeColumnsMode = "Fill"
$dataGrid.AllowUserToResizeColumns = $true
$dataGrid.ClipboardCopyMode = "EnableAlwaysIncludeHeaderText"
$form.Controls.Add($dataGrid)

# Define columns
$colHost = New-Object System.Windows.Forms.DataGridViewTextBoxColumn
$colHost.Name = "Host"
$colHost.HeaderText = "Host"
$colHost.Width = 90
$colHost.FillWeight = 10
$dataGrid.Columns.Add($colHost) | Out-Null

$colDevice = New-Object System.Windows.Forms.DataGridViewTextBoxColumn
$colDevice.Name = "DeviceName"
$colDevice.HeaderText = "Device Name"
$colDevice.Width = 280
$colDevice.FillWeight = 35
$dataGrid.Columns.Add($colDevice) | Out-Null

$colVersion = New-Object System.Windows.Forms.DataGridViewTextBoxColumn
$colVersion.Name = "DriverVersion"
$colVersion.HeaderText = "Driver Version"
$colVersion.Width = 130
$colVersion.FillWeight = 15
$dataGrid.Columns.Add($colVersion) | Out-Null

$colMfg = New-Object System.Windows.Forms.DataGridViewTextBoxColumn
$colMfg.Name = "Manufacturer"
$colMfg.HeaderText = "Manufacturer"
$colMfg.Width = 130
$colMfg.FillWeight = 15
$dataGrid.Columns.Add($colMfg) | Out-Null

$colDate = New-Object System.Windows.Forms.DataGridViewTextBoxColumn
$colDate.Name = "DriverDate"
$colDate.HeaderText = "Driver Date"
$colDate.Width = 100
$colDate.FillWeight = 12
$dataGrid.Columns.Add($colDate) | Out-Null

$colClass = New-Object System.Windows.Forms.DataGridViewTextBoxColumn
$colClass.Name = "DeviceClass"
$colClass.HeaderText = "Class"
$colClass.Width = 100
$colClass.FillWeight = 13
$dataGrid.Columns.Add($colClass) | Out-Null

# ── Summary Label (bottom) ────────────────────────────────────────────────────

$labelSummary = New-Object System.Windows.Forms.Label
$labelSummary.Text = ""
$labelSummary.Location = New-Object System.Drawing.Point(10, 558)
$labelSummary.Size = New-Object System.Drawing.Size(600, 20)
$labelSummary.ForeColor = [System.Drawing.Color]::FromArgb(180, 180, 180)
$labelSummary.Font = New-Object System.Drawing.Font("Segoe UI", 8.5)
$labelSummary.Anchor = [System.Windows.Forms.AnchorStyles]::Bottom -bor [System.Windows.Forms.AnchorStyles]::Left
$form.Controls.Add($labelSummary)

# ── Log Panel (bottom) ────────────────────────────────────────────────────────

$richLog = New-Object System.Windows.Forms.RichTextBox
$richLog.Location = New-Object System.Drawing.Point(10, 580)
$richLog.Size = New-Object System.Drawing.Size(880, 120)
$richLog.Font = New-Object System.Drawing.Font("Consolas", 8.5)
$richLog.BackColor = [System.Drawing.Color]::FromArgb(20, 20, 20)
$richLog.ForeColor = [System.Drawing.Color]::White
$richLog.ReadOnly = $true
$richLog.WordWrap = $false
$richLog.ScrollBars = [System.Windows.Forms.RichTextBoxScrollBars]::Both
$richLog.Anchor = [System.Windows.Forms.AnchorStyles]::Bottom -bor [System.Windows.Forms.AnchorStyles]::Left -bor [System.Windows.Forms.AnchorStyles]::Right
$form.Controls.Add($richLog)

# ══════════════════════════════════════════════════════════════════════════════
#  SCAN BUTTON HANDLER
# ══════════════════════════════════════════════════════════════════════════════

$btnScan.Add_Click({

        $hostInput = $textBoxHosts.Text
        $hostNames = $hostInput -split "`r`n" | ForEach-Object { $_.Trim() } | Where-Object { $_ -match '\S' }

        if ($hostNames.Count -eq 0) {
            [System.Windows.Forms.MessageBox]::Show("No host names entered.", "Validation",
                [System.Windows.Forms.MessageBoxButtons]::OK, [System.Windows.Forms.MessageBoxIcon]::Warning)
            return
        }

        # Parse filter keywords
        $filterText = $textBoxFilter.Text.Trim()
        $keywords = @()
        if ($filterText) {
            $keywords = $filterText -split "," | ForEach-Object { $_.Trim() } | Where-Object { $_ -ne "" }
        }

        # Lock UI
        $script:isScanning = $true
        $script:cancelScan = $false
        $script:allResults = @()
        $btnScan.Enabled = $false
        $btnScan.Text = "Scanning..."
        $btnExport.Enabled = $false
        $textBoxHosts.Enabled = $false
        $textBoxFilter.Enabled = $false
        $dataGrid.Rows.Clear()
        $richLog.Clear()
        $progressBar.Value = 0
        $progressBar.Maximum = $hostNames.Count
        $labelPercent.Text = "0%"
        $labelSummary.Text = ""

        Write-Log $richLog "=================================================" ([System.Drawing.Color]::Cyan)
        Write-Log $richLog "  Remote Driver Checker v1.0" ([System.Drawing.Color]::Cyan)
        Write-Log $richLog "=================================================" ([System.Drawing.Color]::Cyan)
        Write-Log $richLog "  Hosts    : $($hostNames.Count)" ([System.Drawing.Color]::Cyan)
        if ($keywords.Count -gt 0) {
            Write-Log $richLog "  Filter   : $($keywords -join ', ')" ([System.Drawing.Color]::Cyan)
        }
        else {
            Write-Log $richLog "  Filter   : (none - showing all drivers)" ([System.Drawing.Color]::Cyan)
        }
        Write-Log $richLog "=================================================" ([System.Drawing.Color]::Cyan)
        Write-Log $richLog "" ([System.Drawing.Color]::White)

        $hostIndex = 0
        $totalDrivers = 0
        $hostsOK = 0
        $hostsFail = 0

        foreach ($name in $hostNames) {

            $hostIndex++
            $pct = [math]::Round(($hostIndex / $hostNames.Count) * 100)
            $labelStatus.Text = "[$hostIndex/$($hostNames.Count)] Scanning $name..."
            $labelStatus.ForeColor = [System.Drawing.Color]::FromArgb(180, 180, 180)
            [System.Windows.Forms.Application]::DoEvents()

            Write-Log $richLog "--- $name ---" ([System.Drawing.Color]::Cyan)

            # Ping check
            if (!(Test-Connection -ComputerName $name -Count 1 -Quiet)) {
                Write-Log $richLog "  [!] OFFLINE or unreachable." ([System.Drawing.Color]::Red)
                $hostsFail++
                $progressBar.Value = $hostIndex
                $labelPercent.Text = "$pct%"
                [System.Windows.Forms.Application]::DoEvents()
                continue
            }

            Write-Log $richLog "  [OK] Online." ([System.Drawing.Color]::FromArgb(0, 200, 0))

            # WinRM check
            $labelStatus.Text = "[$hostIndex/$($hostNames.Count)] Checking WinRM on $name..."
            [System.Windows.Forms.Application]::DoEvents()

            $winrmTest = Test-WSMan -ComputerName $name -ErrorAction SilentlyContinue
            if (-not $winrmTest) {
                Write-Log $richLog "  [!] WinRM unavailable. Cannot scan drivers." ([System.Drawing.Color]::Red)
                $hostsFail++
                $progressBar.Value = $hostIndex
                $labelPercent.Text = "$pct%"
                [System.Windows.Forms.Application]::DoEvents()
                continue
            }

            # Query drivers remotely
            $labelStatus.Text = "[$hostIndex/$($hostNames.Count)] Querying drivers on $name..."
            [System.Windows.Forms.Application]::DoEvents()

            try {
                $drivers = Invoke-Command -ComputerName $name -ScriptBlock {
                    Get-CimInstance Win32_PnPSignedDriver |
                    Where-Object { $_.DeviceName -ne $null -and $_.DeviceName -ne "" } |
                    Select-Object DeviceName, DriverVersion, Manufacturer, DriverDate, DeviceClass, InfName |
                    Sort-Object DeviceClass, DeviceName
                } -ErrorAction Stop

                if (-not $drivers -or $drivers.Count -eq 0) {
                    Write-Log $richLog "  [!] No driver data returned." ([System.Drawing.Color]::Yellow)
                    $hostsFail++
                    $progressBar.Value = $hostIndex
                    $labelPercent.Text = "$pct%"
                    [System.Windows.Forms.Application]::DoEvents()
                    continue
                }

                # Apply keyword filter
                $filtered = $drivers
                if ($keywords.Count -gt 0) {
                    $filtered = $drivers | Where-Object {
                        $dev = $_
                        $match = $false
                        foreach ($kw in $keywords) {
                            if ($dev.DeviceName -like "*$kw*" -or
                                $dev.Manufacturer -like "*$kw*" -or
                                $dev.DeviceClass -like "*$kw*") {
                                $match = $true
                                break
                            }
                        }
                        $match
                    }
                }

                $count = 0
                if ($filtered) {
                    $filteredArr = @($filtered)
                    $count = $filteredArr.Count

                    foreach ($drv in $filteredArr) {
                        # Format driver date
                        $dateStr = ""
                        if ($drv.DriverDate) {
                            try { $dateStr = ([datetime]$drv.DriverDate).ToString("yyyy-MM-dd") }
                            catch { $dateStr = $drv.DriverDate.ToString() }
                        }

                        $rowIndex = $dataGrid.Rows.Add(
                            $name,
                            $drv.DeviceName,
                            $drv.DriverVersion,
                            $drv.Manufacturer,
                            $dateStr,
                            $drv.DeviceClass
                        )

                        # Store for export
                        $script:allResults += [PSCustomObject]@{
                            Host          = $name
                            DeviceName    = $drv.DeviceName
                            DriverVersion = $drv.DriverVersion
                            Manufacturer  = $drv.Manufacturer
                            DriverDate    = $dateStr
                            DeviceClass   = $drv.DeviceClass
                            InfName       = $drv.InfName
                        }
                    }
                }

                $totalDrivers += $count
                $hostsOK++
                $filterNote = if ($keywords.Count -gt 0) { " (filtered)" } else { "" }
                Write-Log $richLog "  [OK] $count driver(s) found${filterNote}." ([System.Drawing.Color]::FromArgb(0, 200, 0))
            }
            catch {
                Write-Log $richLog "  [X] ERROR: $($_.Exception.Message)" ([System.Drawing.Color]::Red)
                $hostsFail++
            }

            $progressBar.Value = $hostIndex
            $labelPercent.Text = "$pct%"
            [System.Windows.Forms.Application]::DoEvents()
        }

        # ── Scan complete ──────────────────────────────────────────────────────
        $progressBar.Value = $progressBar.Maximum
        $labelPercent.Text = "100%"

        Write-Log $richLog "" ([System.Drawing.Color]::White)
        Write-Log $richLog "=================================================" ([System.Drawing.Color]::Yellow)
        Write-Log $richLog "  SCAN COMPLETE" ([System.Drawing.Color]::Yellow)
        Write-Log $richLog "  Hosts OK: $hostsOK | Failed: $hostsFail | Drivers: $totalDrivers" ([System.Drawing.Color]::Yellow)
        Write-Log $richLog "=================================================" ([System.Drawing.Color]::Yellow)

        $labelSummary.Text = "$totalDrivers driver(s) found across $hostsOK host(s). $hostsFail host(s) failed."
        $labelSummary.ForeColor = if ($hostsFail -gt 0) { [System.Drawing.Color]::Orange } else { [System.Drawing.Color]::FromArgb(0, 200, 0) }

        if ($totalDrivers -eq 0 -and $hostsOK -gt 0 -and $keywords.Count -gt 0) {
            $labelStatus.Text = "No drivers matched the filter. Try different keywords or click 'All Drivers'."
            $labelStatus.ForeColor = [System.Drawing.Color]::Orange
        }
        elseif ($totalDrivers -gt 0) {
            $labelStatus.Text = "[OK] Scan complete. $totalDrivers driver(s) listed."
            $labelStatus.ForeColor = [System.Drawing.Color]::FromArgb(0, 200, 0)
        }
        else {
            $labelStatus.Text = "[!] Scan complete. No data retrieved."
            $labelStatus.ForeColor = [System.Drawing.Color]::Orange
        }

        # Restore UI
        $script:isScanning = $false
        $btnScan.Enabled = $true
        $btnScan.Text = [char]0x1F50D + "  Scan Drivers"
        $btnExport.Enabled = ($script:allResults.Count -gt 0)
        $textBoxHosts.Enabled = $true
        $textBoxFilter.Enabled = $true
    })

# ══════════════════════════════════════════════════════════════════════════════
#  EXPORT CSV HANDLER
# ══════════════════════════════════════════════════════════════════════════════

$btnExport.Add_Click({
        if ($script:allResults.Count -eq 0) {
            [System.Windows.Forms.MessageBox]::Show("No results to export.", "Export",
                [System.Windows.Forms.MessageBoxButtons]::OK,
                [System.Windows.Forms.MessageBoxIcon]::Information)
            return
        }
        $saveDialog = New-Object System.Windows.Forms.SaveFileDialog
        $saveDialog.Filter = "CSV files (*.csv)|*.csv"
        $saveDialog.FileName = "DriverReport_$(Get-Date -Format 'yyyyMMdd_HHmmss').csv"
        $saveDialog.InitialDirectory = "C:\temp"
        if ($saveDialog.ShowDialog() -eq [System.Windows.Forms.DialogResult]::OK) {
            try {
                $script:allResults | Export-Csv -Path $saveDialog.FileName -NoTypeInformation -Encoding UTF8
                Write-Log $richLog "CSV exported: $($saveDialog.FileName)" ([System.Drawing.Color]::FromArgb(0, 200, 0))
                [System.Windows.Forms.MessageBox]::Show("Exported $($script:allResults.Count) row(s) to:`n$($saveDialog.FileName)",
                    "Export Successful",
                    [System.Windows.Forms.MessageBoxButtons]::OK,
                    [System.Windows.Forms.MessageBoxIcon]::Information)
            }
            catch {
                [System.Windows.Forms.MessageBox]::Show("Export failed:`n$($_.Exception.Message)", "Export Error",
                    [System.Windows.Forms.MessageBoxButtons]::OK,
                    [System.Windows.Forms.MessageBoxIcon]::Error)
            }
        }
    })

# ── Ctrl+C copy from DataGridView ─────────────────────────────────────────────
$dataGrid.Add_KeyDown({
        if ($_.Control -and $_.KeyCode -eq 'C') {
            if ($dataGrid.SelectedRows.Count -gt 0) {
                $clipData = $dataGrid.GetClipboardContent()
                if ($clipData) {
                    [System.Windows.Forms.Clipboard]::SetDataObject($clipData)
                }
            }
        }
    })

# ── Show Form ──────────────────────────────────────────────────────────────────
[void]$form.ShowDialog()
$form.Dispose()