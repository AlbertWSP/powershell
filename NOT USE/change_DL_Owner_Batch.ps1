Import-Module ActiveDirectory

# Single owner for all groups
$NewOwnerSam = "HKGL737296"

# Read group emails from text file (one email per line)
$GroupEmails = Get-Content -Path "C:\Temp\GroupEmails.txt"

$results = @()

foreach ($GroupEmail in $GroupEmails) {
    $GroupEmail = $GroupEmail.Trim()  # Remove whitespace
    
    if ([string]::IsNullOrWhiteSpace($GroupEmail)) {
        continue  # Skip empty lines
    }
    
    # Get the group
    $g = Get-ADGroup -Filter "Mail -eq '$GroupEmail'" -Properties ManagedBy, Description, GroupScope, GroupCategory, Mail
    
    if ($g) {
        # Change the owner
        try {
            $newOwner = Get-ADUser -Identity $NewOwnerSam
            Set-ADGroup -Identity $g.DistinguishedName -ManagedBy $newOwner.DistinguishedName
            
            # Verify the change
            $gUpdated = Get-ADGroup -Identity $g.DistinguishedName -Properties ManagedBy, Description, GroupScope, GroupCategory, Mail
            
            $owner = if ($gUpdated.ManagedBy) {
                try {
                    $o = Get-ADObject -Identity $gUpdated.ManagedBy -Properties DisplayName, Name
                    if ($o.DisplayName) { $o.DisplayName } else { $o.Name }
                } catch { "N/A (ManagedBy not found)" }
            } else { "N/A" }
            
            $results += [pscustomobject]@{
                GroupName         = $gUpdated.Name
                Email             = $gUpdated.Mail
                SamAccountName    = $gUpdated.SamAccountName
                Description       = $gUpdated.Description
                GroupCategory     = $gUpdated.GroupCategory
                GroupScope        = $gUpdated.GroupScope
                NewOwner          = $owner
                NewOwnerSam       = $NewOwnerSam
                Status            = "Changed Successfully"
                ManagedByDN       = $gUpdated.ManagedBy
                DistinguishedName = $gUpdated.DistinguishedName
            }
        } catch {
            $results += [pscustomobject]@{
                GroupName         = $g.Name
                Email             = $g.Mail
                SamAccountName    = $g.SamAccountName
                Description       = $g.Description
                GroupCategory     = $g.GroupCategory
                GroupScope        = $g.GroupScope
                NewOwner          = "N/A"
                NewOwnerSam       = $NewOwnerSam
                Status            = "Error: $($_.Exception.Message)"
                ManagedByDN       = $g.ManagedBy
                DistinguishedName = $g.DistinguishedName
            }
        }
    } else {
        $results += [pscustomobject]@{
            GroupName         = "N/A"
            Email             = $GroupEmail
            SamAccountName    = "N/A"
            Description       = "N/A"
            GroupCategory     = "N/A"
            GroupScope        = "N/A"
            NewOwner          = "N/A"
            NewOwnerSam       = $NewOwnerSam
            Status            = "Group not found"
            ManagedByDN       = "N/A"
            DistinguishedName = "N/A"
        }
    }
}

# Display results
$results | Format-Table -AutoSize

# Export to CSV
$Timestamp = Get-Date -Format "yyyyMMdd_HHmmss"
$ExportPath = "C:\Temp\GroupOwnerChange_${Timestamp}.csv"
$results | Export-Csv -Path $ExportPath -NoTypeInformation -Encoding UTF8
Write-Host "`nResults exported to: $ExportPath"