# Add Windows Forms assembly
Add-Type -AssemblyName System.Windows.Forms

# Helper Functions
function GetSamAccountName {
    param ([string]$Email)
    $user = Get-ADUser -Filter {Mail -eq $Email} -Properties SamAccountName
    return $user.SamAccountName
}

function ExportGroupMembers {
    param ([string]$groupName, [string]$suffix = "")
    $timestamp = Get-Date -Format "yyyyMMdd_HHmmss"
    $csvPath = "C:\temp\${groupName}_members${suffix}_${timestamp}.csv"
    $members = Get-ADGroupMember -Identity $groupName | Get-ADUser -Properties DisplayName, Mail
    $members | Select-Object Name, SamAccountName, DisplayName, Mail | Export-Csv -Path $csvPath -NoTypeInformation
    return @{
        Path = $csvPath
        Members = $members
    }
}

function ShowFileDialog {
    $openFileDialog = New-Object System.Windows.Forms.OpenFileDialog
    $openFileDialog.Filter = "Text files (*.txt)|*.txt|All files (*.*)|*.*"
    $openFileDialog.Title = "Select file containing members to remove"
    $openFileDialog.InitialDirectory = [Environment]::GetFolderPath('Desktop')
    
    if ($openFileDialog.ShowDialog() -eq [System.Windows.Forms.DialogResult]::OK) {
        return $openFileDialog.FileName
    }
    return $null
}

function RemoveGroupMembers {
    param (
        [string]$groupName,
        [string]$inputFile
    )
    
    if (-not (Test-Path $inputFile)) {
        Write-Host "Error: Input file not found" -ForegroundColor Red
        return
    }
    
    $members = Get-Content $inputFile | Where-Object { -not [string]::IsNullOrWhiteSpace($_) } | ForEach-Object {
        $member = $_.Trim()
        if ($member -match "^[\w\.-]+@[\w\.-]+\.\w+$") {
            GetSamAccountName -Email $member
        } else {
            $member
        }
    } | Where-Object { $_ }
    
    if ($members.Count -eq 0) { 
        Write-Host "No valid members found in input file" -ForegroundColor Yellow
        return 
    }
    
    try {
        # Show members to be removed
        Write-Host "`nMembers to be removed:" -ForegroundColor Cyan
        $members | ForEach-Object { Write-Host $_ -ForegroundColor Yellow }
        
        $confirm = Read-Host "`nDo you want to proceed? (Y/N)"
        if ($confirm -ne 'Y') {
            Write-Host "Operation cancelled" -ForegroundColor Yellow
            return
        }
        
        Remove-ADGroupMember -Identity $groupName -Members $members -Confirm:$false
        
        $result = ExportGroupMembers -groupName $groupName -suffix "_after_removal"
        Write-Host "`nMembers removed successfully. Updated members exported to: $($result.Path)" -ForegroundColor Green
        Write-Host "`nRemaining Members:" -ForegroundColor Cyan
        $result.Members | Sort-Object Name | ForEach-Object { Write-Host $_.Name -ForegroundColor Yellow }
    }
    catch {
        Write-Host "Error: $_" -ForegroundColor Red
    }
}

# Create export directory if needed
if (-not (Test-Path "C:\temp")) { New-Item -ItemType Directory -Path "C:\temp" }

# Main script
do {
    Clear-Host
    Write-Host "AD Group Member Removal Tool" -ForegroundColor Magenta
    Write-Host "--------------------------" -ForegroundColor Magenta
    Write-Host "1. Remove members using text file" -ForegroundColor Cyan
    Write-Host "2. Exit" -ForegroundColor Red
    
    switch (Read-Host "`nSelect option (1-2)") {
        "1" {
            $groupName = Read-Host "`nEnter AD group name"
            try {
                $null = Get-ADGroup -Identity $groupName
                
                # Show current members
                $initial = ExportGroupMembers -groupName $groupName
                Write-Host "`nCurrent members exported to: $($initial.Path)"
                Write-Host "`nCurrent Members:" -ForegroundColor Cyan
                $initial.Members | Sort-Object Name | ForEach-Object { Write-Host $_.Name -ForegroundColor Yellow }
                
                # Get input file path using GUI
                Write-Host "`nPlease select the text file containing members to remove..." -ForegroundColor Yellow
                $inputFile = ShowFileDialog
                if ($inputFile) {
                    RemoveGroupMembers -groupName $groupName -inputFile $inputFile
                } else {
                    Write-Host "`nNo file selected. Operation cancelled." -ForegroundColor Yellow
                }
                
                Read-Host "`nPress Enter to continue"
            }
            catch {
                Write-Host "Error: Group not found" -ForegroundColor Red
                Read-Host "Press Enter to continue"
            }
        }
        "2" { exit }
        default { 
            Write-Host "Invalid option" -ForegroundColor Red
            Start-Sleep -Seconds 2
        }
    }
} while ($true)