# Read the network paths and groups from text files
$networkPaths = Get-Content -Path "C:\temp\DL_Group\paths.txt"
$groups = Get-Content -Path "C:\temp\DL_Group\groups.txt"

# Ensure the number of paths matches the number of groups
if ($networkPaths.Count -ne $groups.Count) {
    Write-Error "The number of paths does not match the number of groups."
    exit
}

# Loop through each network path and add the corresponding group
for ($i = 0; $i -lt $networkPaths.Count; $i++) {
    $path = $networkPaths[$i]
    $group = $groups[$i]
    
    # Add the group to the network folder
    $acl = Get-Acl -Path $path
    $accessRule = New-Object System.Security.AccessControl.FileSystemAccessRule($group, "FullControl", "ContainerInherit,ObjectInherit", "None", "Allow")
    $acl.SetAccessRule($accessRule)
    Set-Acl -Path $path -AclObject $acl
    
    Write-Output "Added $group to $path"
}