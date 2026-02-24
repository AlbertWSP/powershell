### Get member list of an AD group ###
import-module ActiveDirectory
"Please enter the AD Group name:"
$groupName = Read-Host
$ADGroup = get-adgroup -Filter {Name -eq $groupName}
if ($ADGroup -eq $NULL) {
    Write-Host "Group not found."
    exit
    }

$members = Get-ADGroupMember -Identity $ADGroup.DistinguishedName
$results = @()
 
foreach ( $member in $members) {
    $user =Get-ADUser -Identity $member.DistinguishedName
    #write-host $user.Name "-" $user.UserPrincipalName
    $results += New-Object PSObject -Property @{
        Name = $user.Name
        Email = $user.UserPrincipalName
        }
}
$fileName = "c:\temp\$groupName Members.csv"
$results | Export-Csv -Path $fileName -NoTypeInformation
