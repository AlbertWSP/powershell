Add-Type -AssemblyName System.Windows.Forms

$FileBrowser = New-Object System.Windows.Forms.OpenFileDialog -Property @{
    InitialDirectory = Split-Path (Get-Item $PSCommandPath).FullName
    Filter = "PowerShell Scripts (*.ps1)|*.ps1"
    Title = "Select script to run"
}

if ($FileBrowser.ShowDialog() -ne "OK") { exit }

$Cred = Get-Credential -Message "Enter AD admin credentials"
if (-not $Cred) { exit }

try {
    Start-Process -FilePath "powershell.exe" `
                  -ArgumentList "-ExecutionPolicy Bypass -File `"$($FileBrowser.FileName)`"" `
                  -Credential $Cred `
                  -ErrorAction Stop
    [System.Windows.Forms.MessageBox]::Show("Script launched as $($Cred.UserName)", "Success")
} catch {
    [System.Windows.Forms.MessageBox]::Show("Error: $($_.Exception.Message)", "Launch Failed")
}