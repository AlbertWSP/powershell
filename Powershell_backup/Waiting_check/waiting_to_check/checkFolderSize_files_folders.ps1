# Add required assembly for Windows Forms
Add-Type -AssemblyName System.Windows.Forms

# Set up report paths
$PowershellPath = "C:\temp\Powershell"
$ReportPath = Join-Path $PowershellPath "Report"
if((Test-Path $ReportPath) -eq $False) {
    New-Item -Path $ReportPath -ItemType Directory
}

# Function to convert bytes to MB or GB
function Format-FileSize {
    param (
        [long]$Size
    )
    
    if ($Size -ge 1GB) {
        return "{0:N2} GB" -f ($Size / 1GB)
    }
    else {
        return "{0:N2} MB" -f ($Size / 1MB)
    }
}

# Function to get color based on size
function Get-SizeColor {
    param (
        [long]$Size
    )
    
    if ($Size -ge 1GB) {
        return "Red"
    }
    elseif ($Size -ge 500MB) {
        return "Yellow"
    }
    else {
        return "Green"
    }
}

# Function to get folder statistics
function Get-FolderStats ($folderPath) {
    try {
        $folderSize = (Get-ChildItem -Path $folderPath -Recurse | Measure-Object -Property Length -Sum).Sum
        $fileCount = (Get-ChildItem -Path $folderPath -Recurse -File | Measure-Object).Count
        $folderCount = (Get-ChildItem -Path $folderPath -Recurse -Directory | Measure-Object).Count
        return @{
            Size = $folderSize
            FileCount = $fileCount
            FolderCount = $folderCount
        }
    }
    catch {
        return @{
            Size = 0
            FileCount = 0
            FolderCount = 0
        }
    }
}

# Function to show file selection dialog
function Show-FileSelectionDialog {
    $form = New-Object System.Windows.Forms.Form
    $form.Text = "Select Paths File"
    $form.Size = New-Object System.Drawing.Size(500, 150)
    $form.StartPosition = "CenterScreen"

    $label = New-Object System.Windows.Forms.Label
    $label.Text = "Select the text file containing folder paths:"
    $label.AutoSize = $true
    $label.Location = New-Object System.Drawing.Point(10, 20)
    $form.Controls.Add($label)

    $textbox = New-Object System.Windows.Forms.TextBox
    $textbox.Size = New-Object System.Drawing.Size(360, 20)
    $textbox.Location = New-Object System.Drawing.Point(10, 50)
    $form.Controls.Add($textbox)

    $browseButton = New-Object System.Windows.Forms.Button
    $browseButton.Text = "Browse"
    $browseButton.Location = New-Object System.Drawing.Point(380, 48)
    $browseButton.Add_Click({
        $openFileDialog = New-Object System.Windows.Forms.OpenFileDialog
        $openFileDialog.Filter = "Text files (*.txt)|*.txt"
        $openFileDialog.Title = "Select Paths File"
        if ($openFileDialog.ShowDialog() -eq [System.Windows.Forms.DialogResult]::OK) {
            $textbox.Text = $openFileDialog.FileName
        }
    })
    $form.Controls.Add($browseButton)

    $button = New-Object System.Windows.Forms.Button
    $button.Text = "Process Folders"
    $button.Location = New-Object System.Drawing.Point(10, 80)
    $button.Add_Click({
        $pathsFile = $textbox.Text
        if (Test-Path $pathsFile) {
            Process-Folders $pathsFile
            $form.Close()
        }
        else {
            [System.Windows.Forms.MessageBox]::Show("Please select a valid file path.", "Error", [System.Windows.Forms.MessageBoxButtons]::OK, [System.Windows.Forms.MessageBoxIcon]::Error)
        }
    })
    $form.Controls.Add($button)

    $form.ShowDialog()
}

# Function to process folders from file
function Process-Folders ($pathsFile) {
    # Read the list of folders from the text file
    $folders = Get-Content -Path $pathsFile
    
    Write-Host "`nStarting Folder Size Analysis" -ForegroundColor Cyan
    Write-Host "==========================" -ForegroundColor Cyan
    Write-Host ""
    
    # Initialize array to store results
    $results = @()
    
    foreach ($folder in $folders) {
        if (Test-Path $folder) {
            Write-Host "Processing folder: $folder" -ForegroundColor Yellow
            
            # Get folder statistics
            $stats = Get-FolderStats $folder
            
            # Create result object
            $result = [PSCustomObject]@{
                FolderPath = $folder
                FolderSize = Format-FileSize -Size $stats.Size
                FileCount = $stats.FileCount
                FolderCount = $stats.FolderCount
            }
            $results += $result
            
            # Display information
            Write-Host "Size: " -NoNewline
            Write-Host $result.FolderSize -ForegroundColor (Get-SizeColor -Size $stats.Size)
            Write-Host "Files: " -NoNewline
            Write-Host $result.FileCount -ForegroundColor Magenta
            Write-Host "Subfolders: " -NoNewline
            Write-Host $result.FolderCount -ForegroundColor Cyan
            Write-Host "-------------------"
        }
        else {
            Write-Host "Folder not found: $folder" -ForegroundColor Red
        }
    }
    
    # Export results
    $today = Get-Date -Format "yyyyMMddHHmm"
    $reportName = Join-Path $ReportPath "FolderSize_Report_$today.csv"
    
    # Export to CSV with UTF8 encoding
    $results | Export-Csv -Path $reportName -NoTypeInformation -Force -Encoding UTF8
    
    Write-Host "`nReport generated successfully:" -ForegroundColor Green
    Write-Host "Report location: $reportName" -ForegroundColor Green
}

# Show the file selection dialog
Show-FileSelectionDialog