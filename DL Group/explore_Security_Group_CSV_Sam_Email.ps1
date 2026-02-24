Import-Module ActiveDirectory


$SearchBaseOU = "OU=Messaging,OU=Groups,OU=HK,OU=WSPObjects,DC=corp,DC=pbwan,DC=net"

# Extract OU name for filename
$OUName = ($SearchBaseOU -split ",")[0] -replace "OU="
$Timestamp = Get-Date -Format "yyyyMMdd_HHmmss"
$ExportPath = "C:\Temp\LocationSecurityGroups_${OUName}_${Timestamp}.csv"


Get-ADGroup -Filter * -SearchBase $SearchBaseOU -SearchScope Subtree -Properties Name, Description, GroupCategory, DistinguishedName, ManagedBy, Mail, SAMAccountName |
    Select-Object Name, Description, GroupCategory, DistinguishedName, Mail, SAMAccountName, @{
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

Write-Host $ExportPath