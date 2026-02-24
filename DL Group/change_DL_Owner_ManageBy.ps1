Import-Module ActiveDirectory

# === Values to configure ===
$GroupEmail  = "WSPJ2153EM@WSPGroup.com.hk"      # Group Email
$NewOwnerSam = "HKGL737296"           # New Owner's SamAccountName (User)

# 1) Get the Group and New Owner
$group    = Get-ADGroup -Filter "Mail -eq '$GroupEmail'" -Properties Mail
$newOwner = Get-ADUser -Identity $NewOwnerSam -Properties SID, DisplayName

if (-not $group) {
    Write-Host "Group with email $GroupEmail not found"
    exit
}

# 2) Set ManagedBy (Owner)
Set-ADGroup -Identity $group.DistinguishedName -Replace @{ managedBy = $newOwner.DistinguishedName }

# 3) Add ACL to the Group: Allow Owner to write to member attribute (equivalent to checking "Manager can update membership list")
$aclPath = "AD:\$($group.DistinguishedName)"
$acl     = Get-Acl $aclPath

# GUID for member attribute (common approach: add WriteProperty permission to member attribute)
$guidMemberAttr = [Guid]'bf9679c0-0de6-11d0-a285-00aa003049e2'

# Create Access Rule using SID (more stable than using NTAccount)
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

# 4) Verification Output
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

[pscustomobject]@{
    Group           = $group.Name
    GroupEmail      = $group.Mail
    ManagedBy       = $ownerDisplay
    ManagedByDN     = $g2.ManagedBy
    CanUpdateMember = ($hasWriteMember -gt 0)
}