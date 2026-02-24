# Prompt the user to enter the file path
$filePath = Read-Host "Please enter the file path"

# Check if the file exists
if (Test-Path $filePath) {
    # Get the access control list (ACL) for the file
    $acl = Get-Acl $filePath

    # Display the groups that have access to the file
    Write-Host "Groups with access to the file:"
    foreach ($access in $acl.Access) {
        if ($access.IdentityReference -match "S-1-5-32-") {
            Write-Host $access.IdentityReference
        }
    }
} else {
    Write-Host "The file path you entered does not exist."
}
