# ============================================
# AD Group Nested "Member Of" Export Tool (Simplified)
# ============================================

Import-Module ActiveDirectory

# ==============================
# Function: Get Nested Membership
# ==============================
function Get-NestedGroupMembership {
    param (
        [string]$GroupDN
    )

    $visited = @()
    $queue = @($GroupDN)
    $results = @()

    while ($queue.Count -gt 0) {

        $current = $queue[0]

        if ($queue.Count -gt 1) {
            $queue = $queue[1..($queue.Count - 1)]
        }
        else {
            $queue = @()
        }

        if ($visited -contains $current) { continue }
        $visited += $current

        try {
            $group = Get-ADGroup -Identity $current -Properties MemberOf -ErrorAction Stop

            foreach ($parentDN in $group.MemberOf) {

                if ($visited -notcontains $parentDN) {

                    try {
                        $parentGroup = Get-ADGroup -Identity $parentDN -Properties GroupScope, GroupCategory

                        $results += $parentGroup
                        $queue += $parentDN

                    }
                    catch {
                        Write-Warning "Failed to resolve parent group: $parentDN"
                    }
                }
            }

        }
        catch {
            Write-Warning "Failed to query group: $current"
        }
    }

    return $results
}

# ==============================
# Input
# ==============================
$groupName = Read-Host "Please enter the AD Group name (Name/SamAccountName/DN)"

# ==============================
# Resolve Group
# ==============================
try {
    $ADGroup = Get-ADGroup -Identity $groupName -ErrorAction Stop
}
catch {
    try {
        $ADGroup = Get-ADGroup -Filter "Name -eq '$groupName'" -ErrorAction Stop
    }
    catch {
        Write-Host "Group not found: $groupName" -ForegroundColor Red
        exit
    }
}

# ==============================
# Get Nested Membership
# ==============================
Write-Host "Processing nested membership..." -ForegroundColor Yellow

$memberOfGroups = Get-NestedGroupMembership -GroupDN $ADGroup.DistinguishedName

# ==============================
# Handle Empty Result
# ==============================
if (-not $memberOfGroups) {
    Write-Host "No group memberships found." -ForegroundColor Yellow
}

# ==============================
# Build Output
# ==============================
$results = $memberOfGroups | Sort-Object Name -Unique | ForEach-Object {
    [pscustomobject]@{
        SourceGroupName    = $ADGroup.Name
        MemberOfGroupName  = $_.Name
        MemberOfSamAccount = $_.SamAccountName
        GroupScope         = $_.GroupScope
        GroupCategory      = $_.GroupCategory
        DistinguishedName  = $_.DistinguishedName
    }
}

# ==============================
# Export
# ==============================
$folder = "C:\Temp"
if (!(Test-Path $folder)) {
    New-Item -Path $folder -ItemType Directory | Out-Null
}

$safeName = ($ADGroup.Name -replace '[\\/:*?"<>|]', '_')
$fileName = Join-Path $folder "$safeName-memberOf.csv"

try {
    $results | Export-Csv -Path $fileName -NoTypeInformation -Encoding UTF8
    Write-Host "Exported: $fileName" -ForegroundColor Green
}
catch {
    Write-Host "Failed to export CSV" -ForegroundColor Red
}

# ==============================
# Optional View
# ==============================
$results | Out-GridView -Title "Nested Membership: $($ADGroup.Name)"