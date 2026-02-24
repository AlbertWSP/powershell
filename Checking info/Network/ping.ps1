Add-Type -AssemblyName System.Windows.Forms

# Create main form
$form = New-Object System.Windows.Forms.Form
$form.Text = "Ping Multiple Hosts"
$form.Size = New-Object System.Drawing.Size(600, 400)

# Create results textbox
$textBoxResults = New-Object System.Windows.Forms.TextBox
$textBoxResults.Location = New-Object System.Drawing.Point(10, 40)
$textBoxResults.Size = New-Object System.Drawing.Size(565, 280)
$textBoxResults.Multiline = $true
$textBoxResults.ScrollBars = "Vertical"
$textBoxResults.ReadOnly = $true
$form.Controls.Add($textBoxResults)

# Create Load button
$loadButton = New-Object System.Windows.Forms.Button
$loadButton.Location = New-Object System.Drawing.Point(10, 10)
$loadButton.Size = New-Object System.Drawing.Size(100, 23)
$loadButton.Text = "Load Hosts File"
$loadButton.Add_Click({
    $openFileDialog = New-Object System.Windows.Forms.OpenFileDialog
    $openFileDialog.Filter = "Text files (*.txt)|*.txt|All files (*.*)|*.*"
    $openFileDialog.InitialDirectory = [Environment]::GetFolderPath("Desktop")
    
    if ($openFileDialog.ShowDialog() -eq 'OK') {
        $script:hostsList = Get-Content -Path $openFileDialog.FileName
        $textBoxResults.Text = "Loaded ${($script:hostsList.Count)} hosts from $($openFileDialog.FileName)`r`n"
    }
})
$form.Controls.Add($loadButton)

# Create Start button
$startButton = New-Object System.Windows.Forms.Button
$startButton.Location = New-Object System.Drawing.Point(120, 10)
$startButton.Size = New-Object System.Drawing.Size(100, 23)
$startButton.Text = "Start Ping"
$startButton.Add_Click({
    if (-not $script:hostsList) {
        [System.Windows.Forms.MessageBox]::Show("Please load a hosts file first", "Error")
        return
    }
    
    $textBoxResults.Clear()
    $results = @()
    
    foreach ($hostname in $script:hostsList) {
        if ([string]::IsNullOrWhiteSpace($hostname)) { continue }
        
        $result = Test-Connection -ComputerName $hostname -Count 1 -ErrorAction SilentlyContinue
        if ($result) {
            $status = "Success - ${hostname} responded in $($result.ResponseTime)ms"
            $textBoxResults.AppendText("$status`r`n")
        } else {
            $status = "Failed - ${hostname} did not respond"
            $textBoxResults.AppendText("$status`r`n")
        }
        $results += $status
    }
    
    # Save results to log file
    $timestamp = Get-Date -Format "yyyyMMdd_HHmmss"
    $logPath = Join-Path $PSScriptRoot "ping_results_${timestamp}.log"
    $results | Out-File -FilePath $logPath
    $textBoxResults.AppendText("`r`nResults saved to: $logPath")
})
$form.Controls.Add($startButton)

# Show the form
$form.ShowDialog()
