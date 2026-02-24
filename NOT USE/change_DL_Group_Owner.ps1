Import-Module ActiveDirectory

$GroupEmail = "HK-HongKong-AllEM-BS1Staff@WSPGroup.com" # 改成你的群組 email

$NewOwnerSam = "HKGL737296"   # 新 owner 的 SamAccountName

$group = Get-ADGroup -Filter "Mail -eq '$GroupEmail'" -Properties Mail

if ($group) {
    $newOwner = Get-ADUser -Identity $NewOwnerSam
    Set-ADGroup -Identity $group.DistinguishedName -ManagedBy $newOwner.DistinguishedName
    Write-Host "Owner changed successfully for group: $($group.Name)"
} else {
    Write-Host "Group with email $GroupEmail not found"
}