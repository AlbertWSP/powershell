# 1. Load Windows Forms to enable the File Picker UI
Add-Type -AssemblyName System.Windows.Forms

# 2. Open File Dialog to select the .ps1 file
$FileBrowser = New-Object System.Windows.Forms.OpenFileDialog -Property @{
    InitialDirectory = [Environment]::GetFolderPath('Desktop')
    Filter = "PowerShell Scripts (*.ps1)|*.ps1"
    Title = "Select the PowerShell script to run"
}

if ($FileBrowser.ShowDialog() -ne "OK") {
    Write-Host "No file selected. Exiting." -ForegroundColor Red
    exit
}

$SelectedScript = $FileBrowser.FileName

# 3. Prompt for Account and Password using a secure UI
# This creates a $Cred object containing the username and encrypted password
$Cred = Get-Credential -Message "Enter credentials for the account to run the script"

# Replace Step 4 in your code with this:
try {
    Start-Process -FilePath "powershell.exe" `
                  -ArgumentList "-ExecutionPolicy Bypass -File `"$SelectedScript`"" `
                  -Credential $Cred `
                  -ErrorAction Stop
    Write-Host "Success: Process started as $($Cred.UserName)" -ForegroundColor Green
}
catch {
    [System.Windows.Forms.MessageBox]::Show("Failed to start process: $($_.Exception.Message)", "Error")
}
