Import-Module ActiveDirectory

# === 你要改的值 ===
$GroupName   = "HK - Project - 2539506A - BS (HKSTP DCS)"      # Group Name / SamAccountName / DN 都可以
$NewOwnerSam = "HKGL737296"           # 新 Owner 的 SamAccountName（User）

# 1) 取新 Owner
$newOwner = Get-ADUser -Identity $NewOwnerSam -Properties SID, DisplayName

# 2) 設定 ManagedBy（Owner）
Set-ADGroup -Identity $GroupName -Replace @{ managedBy = $newOwner.DistinguishedName }

# 3) 對該 Group 加 ACL：允許 Owner 寫入 member 屬性（等同勾 "Manager can update membership list"）
$group   = Get-ADGroup -Identity $GroupName
$aclPath = "AD:\$($group.DistinguishedName)"
$acl     = Get-Acl $aclPath

# member attribute 的 GUID（常用做法：對 member 屬性加 WriteProperty）
$guidMemberAttr = [Guid]'bf9679c0-0de6-11d0-a285-00aa003049e2'

# 用 SID 建立 Access Rule（比 NTAccount 更穩）
$sid     = [System.Security.Principal.SecurityIdentifier]$newOwner.SID

$rights  = [System.DirectoryServices.ActiveDirectoryRights]::WriteProperty
$type    = [System.Security.AccessControl.AccessControlType]::Allow
$inherit = [System.DirectoryServices.ActiveDirectorySecurityInheritance]::None

$rule = New-Object System.DirectoryServices.ActiveDirectoryAccessRule `
    ($sid, $rights, $type, $guidMemberAttr, $inherit)

# 避免重覆加同一條 ACE
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

# 4) 驗證輸出
$g2 = Get-ADGroup -Identity $GroupName -Properties ManagedBy
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
    ManagedBy       = $ownerDisplay
    ManagedByDN     = $g2.ManagedBy
    CanUpdateMember = ($hasWriteMember -gt 0)
}
