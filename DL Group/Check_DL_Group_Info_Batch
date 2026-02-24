Import-Module ActiveDirectory

# === Values to configure ===
$TextFilePath = "C:\Temp\GroupEmails.txt"      # Path to text file with group emails (one per line)

# Read group emails from text file
$GroupEmails = Get-Content -Path $TextFilePath | Where-Object { $_ -match '\S' }

if (-not $GroupEmails) {
    Write-Host "No group emails found in $TextFilePath"
    exit
}

$results = @()

foreach ($GroupEmail in $GroupEmails) {
    $GroupEmail = $GroupEmail.Trim()
    
    Write-Host "Checking: $GroupEmail"
    
    $g = Get-ADGroup -Filter "Mail -eq '$GroupEmail'" -Properties ManagedBy, Description, GroupScope, GroupCategory, Mail
    
    if ($g) {
        $owner = if ($g.ManagedBy) {
            try {
                $o = Get-ADObject -Identity $g.ManagedBy -Properties DisplayName, Name
                if ($o.DisplayName) { $o.DisplayName } else { $o.Name }
            } catch { "N/A (ManagedBy not found)" }
        } else { "N/A" }
        
        $results += [pscustomobject]@{
            GroupName         = $g.Name
            Email             = $g.Mail
            SamAccountName    = $g.SamAccountName
            Description       = $g.Description
            GroupCategory     = $g.GroupCategory
            GroupScope        = $g.GroupScope
            Owner             = $owner
            ManagedByDN       = $g.ManagedBy
            DistinguishedName = $g.DistinguishedName
        }
    } else {
        Write-Host "  - Group with email $GroupEmail not found"
        $results += [pscustomobject]@{
            GroupName         = "N/A"
            Email             = $GroupEmail
            SamAccountName    = "N/A"
            Description       = "N/A"
            GroupCategory     = "N/A"
            GroupScope        = "N/A"
            Owner             = "N/A"
            ManagedByDN       = "N/A"
            DistinguishedName = "N/A"
        }
    }
}

# Display results
Write-Host "`n====== Group Information ======"
$results | Format-Table -AutoSize

# Export to CSV
$Timestamp = Get-Date -Format "yyyyMMdd_HHmmss"
$ExportPath = "C:\Temp\GroupInfo_${Timestamp}.csv"
$results | Export-Csv -Path $ExportPath -NoTypeInformation -Encoding UTF8
Write-Host "`nResults exported to: $ExportPath"