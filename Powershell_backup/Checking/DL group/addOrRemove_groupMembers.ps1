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

function ManageGroupMembers {
    param (
        [string]$groupName,
        [string]$operation,
        [string]$memberInput
    )
    
    if ([string]::IsNullOrWhiteSpace($memberInput)) { return }
    
    $members = $memberInput -split "," | ForEach-Object {
        $member = $_.Trim()
        if ($member -match "^[\w\.-]+@[\w\.-]+\.\w+$") {
            GetSamAccountName -Email $member
        } else {
            $member
        }
    } | Where-Object { $_ }
    
    if ($members.Count -eq 0) { return }
    
    try {
        if ($operation -eq "add") {
            Add-ADGroupMember -Identity $groupName -Members $members
        } else {
            Remove-ADGroupMember -Identity $groupName -Members $members -Confirm:$false
        }
        
        $result = ExportGroupMembers -groupName $groupName -suffix "_updated"
        Write-Host "`nOperation successful. Updated members exported to: $($result.Path)" -ForegroundColor Green
        Write-Host "`nCurrent Members:" -ForegroundColor Cyan
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
    Write-Host "AD Group Management" -ForegroundColor Magenta
    Write-Host "-------------------" -ForegroundColor Magenta
    Write-Host "1. Add members to group" -ForegroundColor Cyan
    Write-Host "2. Remove members from group" -ForegroundColor Cyan
    Write-Host "3. Exit" -ForegroundColor Red
    
    switch (Read-Host "`nSelect option (1-3)") {
        "1" { $operation = "add" }
        "2" { $operation = "remove" }
        "3" { exit }
        default { 
            Write-Host "Invalid option" -ForegroundColor Red
            continue 
        }
    }
    
    # Get group name
    $groupName = Read-Host "`nEnter AD group name"
    try {
        $null = Get-ADGroup -Identity $groupName
        
        # Show current members
        $initial = ExportGroupMembers -groupName $groupName
        Write-Host "`nCurrent members exported to: $($initial.Path)"
        Write-Host "`nCurrent Members:" -ForegroundColor Cyan
        $initial.Members | Sort-Object Name | ForEach-Object { Write-Host $_.Name -ForegroundColor Yellow }
        
        # Get members to add/remove
        $memberInput = Read-Host "`nEnter email addresses or SAM account names (comma-separated)"
        ManageGroupMembers -groupName $groupName -operation $operation -memberInput $memberInput
        
        Read-Host "`nPress Enter to continue"
    }
    catch {
        Write-Host "Error: Group not found" -ForegroundColor Red
        Read-Host "Press Enter to continue"
    }
} while ($true)