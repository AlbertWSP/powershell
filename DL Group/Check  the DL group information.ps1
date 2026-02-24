Import-Module ActiveDirectory

$GroupEmail = "HK-HongKong-AllEM-BS1Staff@WSPGroup.com"   # 改成你的群組 email

$g = Get-ADGroup -Filter "Mail -eq '$GroupEmail'" -Properties ManagedBy, Description, GroupScope, GroupCategory, Mail

if ($g) {
    $owner = if ($g.ManagedBy) {
        try {
            $o = Get-ADObject -Identity $g.ManagedBy -Properties DisplayName, Name
            if ($o.DisplayName) { $o.DisplayName } else { $o.Name }
        } catch { "N/A (ManagedBy not found)" }
    } else { "N/A" }

    [pscustomobject]@{
        GroupName         = $g.Name
        SamAccountName    = $g.SamAccountName
        Email             = $g.Mail
        Description       = $g.Description
        GroupCategory     = $g.GroupCategory
        GroupScope        = $g.GroupScope
        Owner             = $owner
        ManagedByDN       = $g.ManagedBy
        DistinguishedName = $g.DistinguishedName
    }
} else {
    Write-Host "Group with email $GroupEmail not found"
}