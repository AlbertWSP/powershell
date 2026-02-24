Import-Module ActiveDirectory


$SearchBaseOU = "OU=Security,OU=Groups,OU=HK,OU=WSPObjects,DC=corp,DC=pbwan,DC=net"


$ExportPath = "C:\Temp\LocationSecurityGroups.csv"


Get-ADGroup -Filter * -SearchBase $SearchBaseOU -SearchScope Subtree -Properties Name, Description, GroupCategory, DistinguishedName, ManagedBy |
    Select-Object Name, Description, GroupCategory, DistinguishedName, @{
        Name = "Owner"
        Expression = {
            if ($_.ManagedBy) {
                (Get-ADUser $_.ManagedBy -Properties DisplayName).DisplayName
            } else {
                "N/A"
            }
        }
    } |
    Export-Csv -Path $ExportPath -NoTypeInformation -Encoding UTF8

Write-Host "C:\Temp\LocationSecurityGroups.csv"