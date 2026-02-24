
<# 
.SYNOPSIS
Batch set AD group ManagedBy and enable "Manager can update membership list" permission.

.DESCRIPTION
1) Set managedBy on group
2) Add an ACE on the group DACL allowing the manager to WriteProperty on the "member" attribute 
   (equivalent to ticking "Manager can update membership list" in ADUC).  [1](https://activedirectoryfaq.com/2021/03/manager-can-update-membership-list/)[3](https://acloudguy.com/2020/07/18/active-directory-set-manager-can-update-membership-list-with-powershell/)

.PARAMETER CsvPath
CSV file path. Columns supported:
- GroupIdentity  (required) : Name / SamAccountName / DistinguishedName
- Owner          (required) : sAMAccountName / UPN / DistinguishedName (User/Group/Contact)

.PARAMETER SearchBaseOU
OU DN to search groups from.

.PARAMETER Owner
Owner value for OU mode (single owner for all groups in OU). Accept sAMAccountName/UPN/DN.

.PARAMETER Subtree
Search groups in OU subtree.

.PARAMETER ReportPath
Output report CSV path. If not supplied, script creates one in C:\Temp.

.EXAMPLE
# CSV mode:
.\Set-GroupOwnerMembershipBatch.ps1 -CsvPath C:\Temp\grp.csv -WhatIf

.EXAMPLE
# OU mode (same owner for all groups in OU):
.\Set-GroupOwnerMembershipBatch.ps1 -SearchBaseOU "OU=Security,OU=Groups,OU=HK,OU=WSPObjects,DC=corp,DC=pbwan,DC=net" `
  -Owner "HKRK738628" -Subtree -Verbose
#>

[CmdletBinding(SupportsShouldProcess=$true)]
param(
    [Parameter(ParameterSetName="CSV", Mandatory=$true)]
    [string]$CsvPath,

    [Parameter(ParameterSetName="OU", Mandatory=$true)]
    [string]$SearchBaseOU,

    [Parameter(ParameterSetName="OU", Mandatory=$true)]
    [string]$Owner,

    [Parameter(ParameterSetName="OU")]
    [switch]$Subtree,

    [string]$ReportPath
)

Import-Module ActiveDirectory

# ---- Helpers ----
function Resolve-ADPrincipal {
    param(
        [Parameter(Mandatory=$true)]
        [string]$Identity
    )

    # If it's a DN, try direct bind first
    if ($Identity -match '^(CN|OU|DC)=' ) {
        try { return Get-ADObject -Identity $Identity -Properties objectSid, displayName, name, sAMAccountName -ErrorAction Stop }
        catch { }
    }

    # Try sAMAccountName / UPN
    $escaped = $Identity.Replace('\','\\').Replace('(','\28').Replace(')','\29')  # light escape for LDAP filter safety
    $filter = "(|(sAMAccountName=$escaped)(userPrincipalName=$escaped)(name=$escaped))"

    $obj = Get-ADObject -LDAPFilter $filter -Properties objectSid, displayName, name, sAMAccountName |
           Select-Object -First 1

    return $obj
}

function Ensure-ReportPath {
    param([string]$Path)

    if ($Path) { return $Path }

    $outDir = "C:\Temp"
    if (-not (Test-Path $outDir)) { New-Item $outDir -ItemType Directory -Force | Out-Null }

    $ts = Get-Date -Format "yyyyMMdd_HHmmss"
    return Join-Path $outDir "Batch_SetGroupOwnerMembership_$ts.csv"
}

$ReportPath = Ensure-ReportPath -Path $ReportPath

# member attribute GUID used to scope WriteProperty to the Member attribute in the group's ACL [3](https://acloudguy.com/2020/07/18/active-directory-set-manager-can-update-membership-list-with-powershell/)[4](https://serverfault.com/questions/118625/how-to-set-or-clear-manager-can-update-membership-list-using-powershell)[5](https://community.spiceworks.com/t/powershell-active-directory-group-manager-can-update-membership-list/940569)
$guidMemberAttr = [Guid]'bf9679c0-0de6-11d0-a285-00aa003049e2'
$rights  = [System.DirectoryServices.ActiveDirectoryRights]::WriteProperty
$type    = [System.Security.AccessControl.AccessControlType]::Allow
$inherit = [System.DirectoryServices.ActiveDirectorySecurityInheritance]::None

# Build workload
$workItems = @()

if ($PSCmdlet.ParameterSetName -eq "CSV") {
    if (-not (Test-Path $CsvPath)) { throw "CSV not found: $CsvPath" }
    $csv = Import-Csv -Path $CsvPath

    foreach ($row in $csv) {
        if (-not $row.GroupIdentity -or -not $row.Owner) { 
            Write-Warning "Skip row missing GroupIdentity/Owner: $($row | ConvertTo-Json -Compress)"
            continue
        }
        $workItems += [pscustomobject]@{ GroupIdentity = $row.GroupIdentity; Owner = $row.Owner }
    }
}
else {
    $scope = if ($Subtree) { "Subtree" } else { "OneLevel" }
    $groups = Get-ADGroup -Filter * -SearchBase $SearchBaseOU -SearchScope $scope
    foreach ($g in $groups) {
        $workItems += [pscustomobject]@{ GroupIdentity = $g.DistinguishedName; Owner = $Owner }
    }
}

# Process
$results = New-Object System.Collections.Generic.List[object]

foreach ($item in $workItems) {

    $status = "OK"
    $err = $null
    $managedBySet = $false
    $aceAdded = $false

    try {
        $group = Get-ADGroup -Identity $item.GroupIdentity -Properties managedBy -ErrorAction Stop

        $ownerObj = Resolve-ADPrincipal -Identity $item.Owner
        if (-not $ownerObj) { throw "Owner not found in AD: $($item.Owner)" }
        if (-not $ownerObj.objectSid) { throw "Owner object has no objectSid: $($ownerObj.DistinguishedName)" }

        $ownerSid = New-Object System.Security.Principal.SecurityIdentifier($ownerObj.objectSid, 0)

        # 1) Set ManagedBy
        if ($PSCmdlet.ShouldProcess($group.Name, "Set ManagedBy to $($ownerObj.DistinguishedName)")) {
            Set-ADGroup -Identity $group.DistinguishedName -Replace @{ managedBy = $ownerObj.DistinguishedName } -ErrorAction Stop
            $managedBySet = $true
        }

        # 2) Add ACE for "WriteProperty on member attribute"
        $aclPath = "AD:\$($group.DistinguishedName)"
        $acl = Get-Acl $aclPath

        $exists = $acl.Access | Where-Object {
            $_.IdentityReference -eq $ownerSid -and
            $_.AccessControlType -eq "Allow" -and
            ($_.ActiveDirectoryRights -band $rights) -and
            $_.ObjectType -eq $guidMemberAttr
        }

        if (-not $exists) {
            $rule = New-Object System.DirectoryServices.ActiveDirectoryAccessRule `
                ($ownerSid, $rights, $type, $guidMemberAttr, $inherit)

            if ($PSCmdlet.ShouldProcess($group.Name, "Add ACL ACE: Allow WriteProperty on member for $($item.Owner)")) {
                $null = $acl.AddAccessRule($rule)
                Set-Acl -Path $aclPath -AclObject $acl -ErrorAction Stop
                $aceAdded = $true
            }
        }
        else {
            $aceAdded = $false # already exists
        }
    }
    catch {
        $status = "FAILED"
        $err = $_.Exception.Message
    }

    $results.Add([pscustomobject]@{
        GroupIdentity = $item.GroupIdentity
        OwnerInput    = $item.Owner
        ManagedBySet  = $managedBySet
        AceAdded      = $aceAdded
        Status        = $status
        Error         = $err
        Timestamp     = (Get-Date)
    }) | Out-Null
}

# Export report
$results | Export-Csv -Path $ReportPath -NoTypeInformation -Encoding UTF8
Write-Host "Done. Report: $ReportPath"
