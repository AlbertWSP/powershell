# Add required assembly for Windows Forms
Add-Type -AssemblyName System.Windows.Forms

# Set up report paths
$PowershellPath = "C:\temp\Powershell"
$ReportPath = Join-Path $PowershellPath "Report"

# Function to extract folder path from ACL report filename
function Get-FolderPathFromReportName {
    param (
        [string]$ReportName
    )
    
    # Remove prefix "ACL_Report_"
    $withoutPrefix = $ReportName -replace '^ACL_Report_', ''
    
    # Remove depth info and timestamp
    $withoutSuffix = $withoutPrefix -replace '_Depth\d+_\d+\.csv$', ''
    $withoutSuffix = $withoutSuffix -replace '_\d+\.csv$', ''
    
    # Replace underscores back to path separators
    $folderPath = $withoutSuffix -replace '_', '\'
    
    return $folderPath
}

# Function to select a file containing folder paths
function Show-FileSelectionDialog {
    $openFileDialog = New-Object System.Windows.Forms.OpenFileDialog
    $openFileDialog.Filter = "Text files (*.txt)|*.txt"
    $openFileDialog.Title = "Select Text File Containing Folder Paths"
    
    if ($openFileDialog.ShowDialog() -eq [System.Windows.Forms.DialogResult]::OK) {
        return $openFileDialog.FileName
    }
    else {
        return $null
    }
}

# Main function
function Check-Missing-ACLReports {
    # Select the file containing folder paths
    Write-Host "Please select the text file containing the list of folders to check..." -ForegroundColor Yellow
    $pathsFile = Show-FileSelectionDialog
    
    if ($pathsFile -eq $null) {
        Write-Host "Operation cancelled." -ForegroundColor Yellow
        return
    }
    
    # Check if the file exists
    if (-not (Test-Path $pathsFile)) {
        Write-Host "File not found: $pathsFile" -ForegroundColor Red
        return
    }
    
    # Check if report directory exists
    if (-not (Test-Path $ReportPath)) {
        Write-Host "Report directory not found: $ReportPath" -ForegroundColor Red
        Write-Host "All folders need to be scanned." -ForegroundColor Cyan
        
        # Display folders that need to be scanned
        $folderPaths = Get-Content -Path $pathsFile
        Write-Host "`nFolders that need ACL reports:" -ForegroundColor Green
        foreach ($folderPath in $folderPaths) {
            if (Test-Path $folderPath) {
                Write-Host "  - $folderPath" -ForegroundColor White
            }
            else {
                Write-Host "  - $folderPath (Path not found)" -ForegroundColor Red
            }
        }
        return
    }
    
    # Get ACL report files
    $aclReportFiles = Get-ChildItem -Path $ReportPath -Filter "ACL_Report_*.csv" -ErrorAction SilentlyContinue | Select-Object -ExpandProperty Name
    
    if ($aclReportFiles -eq $null -or $aclReportFiles.Count -eq 0) {
        Write-Host "No ACL reports found in $ReportPath" -ForegroundColor Red
        Write-Host "All folders need to be scanned." -ForegroundColor Cyan
        
        # Display folders that need to be scanned
        $folderPaths = Get-Content -Path $pathsFile
        Write-Host "`nFolders that need ACL reports:" -ForegroundColor Green
        foreach ($folderPath in $folderPaths) {
            if (Test-Path $folderPath) {
                Write-Host "  - $folderPath" -ForegroundColor White
            }
            else {
                Write-Host "  - $folderPath (Path not found)" -ForegroundColor Red
            }
        }
        return
    }
    
    # Extract folder paths from report names
    $scannedFolders = @()
    foreach ($reportFile in $aclReportFiles) {
        $folderPath = Get-FolderPathFromReportName -ReportName $reportFile
        if (-not ($scannedFolders -contains $folderPath)) {
            $scannedFolders += $folderPath
        }
    }
    
    # Read folder paths that should be scanned
    $requiredFolders = Get-Content -Path $pathsFile
    
    # Find missing scans
    $missingScans = @()
    $existingScans = @()
    $invalidPaths = @()
    
    foreach ($folderPath in $requiredFolders) {
        if (-not (Test-Path $folderPath)) {
            $invalidPaths += $folderPath
        }
        elseif ($scannedFolders -contains $folderPath) {
            # Find the most recent report for this folder
            $folderReports = $aclReportFiles | Where-Object { (Get-FolderPathFromReportName -ReportName $_) -eq $folderPath }
            $latestReport = $folderReports | Sort-Object -Descending | Select-Object -First 1
            
            $existingScans += [PSCustomObject]@{
                FolderPath = $folderPath
                ReportFile = $latestReport
                ReportDate = (Get-ChildItem -Path (Join-Path $ReportPath $latestReport)).LastWriteTime
            }
        }
        else {
            $missingScans += $folderPath
        }
    }
    
    # Display results
    Write-Host "`n==============================================" -ForegroundColor Cyan
    Write-Host "ACL Report Status Check" -ForegroundColor Cyan
    Write-Host "==============================================`n" -ForegroundColor Cyan
    
    Write-Host "Total folders to check: $($requiredFolders.Count)" -ForegroundColor White
    Write-Host "Folders with reports: $($existingScans.Count)" -ForegroundColor Green
    Write-Host "Folders without reports: $($missingScans.Count)" -ForegroundColor Yellow
    Write-Host "Invalid paths: $($invalidPaths.Count)`n" -ForegroundColor Red
    
    if ($existingScans.Count -gt 0) {
        Write-Host "Folders with existing ACL reports:" -ForegroundColor Green
        foreach ($scan in $existingScans) {
            Write-Host "  - $($scan.FolderPath)" -ForegroundColor White
            Write-Host "    Last report: $($scan.ReportFile) ($($scan.ReportDate))" -ForegroundColor Gray
        }
        Write-Host ""
    }
    
    if ($missingScans.Count -gt 0) {
        Write-Host "Folders that need ACL reports:" -ForegroundColor Yellow
        foreach ($folderPath in $missingScans) {
            Write-Host "  - $folderPath" -ForegroundColor White
        }
        Write-Host ""
    }
    
    if ($invalidPaths.Count -gt 0) {
        Write-Host "Invalid folder paths:" -ForegroundColor Red
        foreach ($folderPath in $invalidPaths) {
            Write-Host "  - $folderPath" -ForegroundColor White
        }
        Write-Host ""
    }
    
    # Ask if user wants to save a list of missing scans
    if ($missingScans.Count -gt 0) {
        $result = [System.Windows.Forms.MessageBox]::Show(
            "Do you want to save the list of folders that need ACL reports to a file?", 
            "Save Missing Scans", 
            [System.Windows.Forms.MessageBoxButtons]::YesNo, 
            [System.Windows.Forms.MessageBoxIcon]::Question)
            
        if ($result -eq [System.Windows.Forms.DialogResult]::Yes) {
            $saveFileDialog = New-Object System.Windows.Forms.SaveFileDialog
            $saveFileDialog.Filter = "Text files (*.txt)|*.txt"
            $saveFileDialog.Title = "Save List of Folders Needing ACL Reports"
            $saveFileDialog.FileName = "MissingACLReports.txt"
            
            if ($saveFileDialog.ShowDialog() -eq [System.Windows.Forms.DialogResult]::OK) {
                $missingScans | Out-File -FilePath $saveFileDialog.FileName -Force
                Write-Host "Missing folders list saved to: $($saveFileDialog.FileName)" -ForegroundColor Green
            }
        }
    }
    
    # Ask if user wants to scan missing folders now
    if ($missingScans.Count -gt 0) {
        $result = [System.Windows.Forms.MessageBox]::Show(
            "Do you want to scan the missing $($missingScans.Count) folders now?", 
            "Scan Missing Folders", 
            [System.Windows.Forms.MessageBoxButtons]::YesNo, 
            [System.Windows.Forms.MessageBoxIcon]::Question)
            
        if ($result -eq [System.Windows.Forms.DialogResult]::Yes) {
            # Save missing scans to a temporary file
            $tempFile = [System.IO.Path]::GetTempFileName()
            $missingScans | Out-File -FilePath $tempFile -Force
            
            # Ask for scan depth
            $depthForm = New-Object System.Windows.Forms.Form
            $depthForm.Text = "Select Scan Depth"
            $depthForm.Size = New-Object System.Drawing.Size(300, 150)
            $depthForm.StartPosition = "CenterScreen"
            
            $depthLabel = New-Object System.Windows.Forms.Label
            $depthLabel.Text = "Select scan depth (0 = root only, 1-10 = levels of subfolders):"
            $depthLabel.AutoSize = $true
            $depthLabel.Location = New-Object System.Drawing.Point(10, 10)
            $depthForm.Controls.Add($depthLabel)
            
            $depthNumeric = New-Object System.Windows.Forms.NumericUpDown
            $depthNumeric.Location = New-Object System.Drawing.Point(10, 40)
            $depthNumeric.Size = New-Object System.Drawing.Size(100, 20)
            $depthNumeric.Minimum = 0
            $depthNumeric.Maximum = 10
            $depthNumeric.Value = 2
            $depthForm.Controls.Add($depthNumeric)
            
            $okButton = New-Object System.Windows.Forms.Button
            $okButton.Text = "OK"
            $okButton.Location = New-Object System.Drawing.Point(120, 40)
            $okButton.Add_Click({
                $depthForm.DialogResult = [System.Windows.Forms.DialogResult]::OK
                $depthForm.Close()
            })
            $depthForm.Controls.Add($okButton)
            
            $depthResult = $depthForm.ShowDialog()
            
            if ($depthResult -eq [System.Windows.Forms.DialogResult]::OK) {
                $depth = $depthNumeric.Value
                
                # Run ACL_Report.ps1 with the temporary file and selected depth
                Write-Host "Launching ACL scan with depth $depth..." -ForegroundColor Green
                
                # Get the full path to ACL_Report.ps1
                $scriptPath = Join-Path (Get-Location) "ACL_Report.ps1"
                
                if (Test-Path $scriptPath) {
                    # Create a modified version of the script that accepts command line parameters
                    $tempScriptPath = [System.IO.Path]::GetTempFileName() + ".ps1"
                    $scriptContent = Get-Content -Path $scriptPath -Raw
                    
                    # Add parameter handling at the beginning
                    $parameterCode = @"
param(
    [string]`$pathsFile,
    [int]`$depth = 2
)

# If parameters are provided, use them instead of the GUI
if (`$pathsFile -and (Test-Path `$pathsFile)) {
    Process-Folders `$pathsFile `$depth
    exit
}

"@
                    
                    $modifiedScript = $parameterCode + $scriptContent
                    $modifiedScript | Out-File -FilePath $tempScriptPath -Force
                    
                    # Run the modified script with parameters
                    Write-Host "Running ACL scan for missing folders..." -ForegroundColor Cyan
                    PowerShell -ExecutionPolicy Bypass -File $tempScriptPath -pathsFile $tempFile -depth $depth
                    
                    # Clean up
                    Remove-Item -Path $tempScriptPath -Force
                }
                else {
                    Write-Host "ACL_Report.ps1 not found at: $scriptPath" -ForegroundColor Red
                    Write-Host "Please run the ACL_Report.ps1 script manually with this input file: $tempFile" -ForegroundColor Yellow
                }
            }
            
            # Clean up the temporary file
            Remove-Item -Path $tempFile -Force -ErrorAction SilentlyContinue
        }
    }
}

# Run the main function
Check-Missing-ACLReports
