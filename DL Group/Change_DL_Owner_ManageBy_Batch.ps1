Import-Module ActiveDirectory

# === Values to configure ===
$TextFilePath = "C:\Temp\GroupEmails.txt"      # Path to text file with group emails (one per line)
$NewOwnerSam  = "HKRK738628"                   # New Owner's SamAccountName (User)

# Read group emails from text file
$GroupEmails = Get-Content -Path $TextFilePath | Where-Object { $_ -match '\S' }

if (-not $GroupEmails) {
    Write-Host "No group emails found in $TextFilePath"
    exit
}

$results = @()

foreach ($GroupEmail in $GroupEmails) {
    $GroupEmail = $GroupEmail.Trim()
    
    Write-Host "Processing: $GroupEmail"
    
    # 1) Get the Group and New Owner
    $group = Get-ADGroup -Filter "Mail -eq '$GroupEmail'" -Properties Mail
    
    if (-not $group) {
        Write-Host "  - Group with email $GroupEmail not found"
        $results += [pscustomobject]@{
            GroupEmail      = $GroupEmail
            Group           = "N/A"
            Status          = "Not found"
            Owner_or_ManagedBy        = "N/A"
        }
        continue
    }
    
    try {
        $newOwner = Get-ADUser -Identity $NewOwnerSam -Properties SID, DisplayName
        
        # 2) Set ManagedBy (Owner)
        Set-ADGroup -Identity $group.DistinguishedName -Replace @{ managedBy = $newOwner.DistinguishedName }
        
        # 3) Add ACL to the Group: Allow Owner to write to member attribute
        $aclPath = "AD:\$($group.DistinguishedName)"
        $acl     = Get-Acl $aclPath
        
        # GUID for member attribute
        $guidMemberAttr = [Guid]'bf9679c0-0de6-11d0-a285-00aa003049e2'
        
        # Create Access Rule using SID
        $sid     = [System.Security.Principal.SecurityIdentifier]$newOwner.SID
        
        $rights  = [System.DirectoryServices.ActiveDirectoryRights]::WriteProperty
        $type    = [System.Security.AccessControl.AccessControlType]::Allow
        $inherit = [System.DirectoryServices.ActiveDirectorySecurityInheritance]::None
        
        $rule = New-Object System.DirectoryServices.ActiveDirectoryAccessRule `
            ($sid, $rights, $type, $guidMemberAttr, $inherit)
        
        # Avoid adding duplicate ACE
        $exists = $acl.Access | Where-Object {
            $_.IdentityReference -eq $sid -and
            $_.AccessControlType -eq "Allow" -and
            ($_.ActiveDirectoryRights -band $rights) -and
            $_.ObjectType -eq $guidMemberAttr
        }
        
        if (-not $exists) {
            $null = $acl.AddAccessRule($rule)
            Set-Acl -Path $aclPath -AclObject $acl
        }
        
        # 4) Verification
        $g2 = Get-ADGroup -Identity $group.DistinguishedName -Properties ManagedBy
        $ownerDisplay = if ($g2.ManagedBy) {
            $o = Get-ADObject -Identity $g2.ManagedBy -Properties DisplayName, Name
            if ($o.DisplayName) { $o.DisplayName } else { $o.Name }
        } else { "N/A" }
        
        $acl2 = Get-Acl $aclPath
        $hasWriteMember = $acl2.Access | Where-Object {
            $_.IdentityReference -eq $sid -and
            $_.AccessControlType -eq "Allow" -and
            ($_.ActiveDirectoryRights -band $rights) -and
            $_.ObjectType -eq $guidMemberAttr
        } | Measure-Object | Select-Object -ExpandProperty Count
        
        Write-Host "  - Successfully updated"
        
        $results += [pscustomobject]@{
            GroupEmail      = $GroupEmail
            Group           = $group.Name
            Status          = "Success"
            New_Owner_or_ManagedBy       = $ownerDisplay
        }
    }
    catch {
        Write-Host "  - Error: $($_.Exception.Message)"
        $results += [pscustomobject]@{
            GroupEmail      = $GroupEmail
            Group           = $group.Name
            Status          = "Error: $($_.Exception.Message)"
            New_Owner_or_ManagedBy       = "N/A"
        }
    }
}

# Display results
Write-Host "`n====== Results ======"
$results | Format-Table -AutoSize

# Export to CSV
$Timestamp = Get-Date -Format "yyyyMMdd_HHmmss"
$ExportPath = "C:\Temp\AD_DLGroupOwnerChange_${Timestamp}.csv"
$results | Export-Csv -Path $ExportPath -NoTypeInformation -Encoding UTF8
Write-Host "`nResults exported to: $ExportPath"