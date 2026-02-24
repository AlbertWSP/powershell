# Script to remove a group from a specified path

# Function to write logs
function Write-Log {
    param(
        [string]$Message,
        [string]$Type = "INFO"
    )
    $timestamp = Get-Date -Format "yyyy-MM-dd HH:mm:ss"
    Write-Host "[$timestamp] $Type : $Message"
}

# Function to check groups in path
function Get-GroupsInPath {
    param(
        [Parameter(Mandatory=$true)]
        [string]$Path
    )

    try {
        # Check if path exists
        if (-not (Test-Path $Path)) {
            Write-Log "Path '$Path' does not exist." -Type "ERROR"
            return $null
        }

        # Get current ACL
        $acl = Get-Acl $Path
        $accessRules = $acl.Access

        # Get all groups with access
        $groups = $accessRules | Select-Object IdentityReference, FileSystemRights, AccessControlType | 
                 Where-Object { $_.IdentityReference -like "*\*" } |  # Only get domain groups
                 Sort-Object IdentityReference

        if ($groups) {
            Write-Log "Groups found in path '$Path':" -Type "INFO"
            $index = 1
            foreach ($group in $groups) {
                Write-Log "$index. Group: $($group.IdentityReference), Rights: $($group.FileSystemRights), Type: $($group.AccessControlType)" -Type "INFO"
                $index++
            }
            return $groups
        } else {
            Write-Log "No groups found in path '$Path'" -Type "WARNING"
            return $null
        }
    }
    catch {
        Write-Log "Error occurred while checking groups: $_" -Type "ERROR"
        return $null
    }
}

# Function to remove group from path
function Remove-GroupFromPath {
    param(
        [Parameter(Mandatory=$true)]
        [string]$Path,
        
        [Parameter(Mandatory=$true)]
        [string]$GroupName
    )

    try {
        # Check if path exists
        if (-not (Test-Path $Path)) {
            Write-Log "Path '$Path' does not exist." -Type "ERROR"
            return $false
        }

        # Get current ACL
        $acl = Get-Acl $Path
        $accessRules = $acl.Access

        # Find and remove the specified group's access rule
        $groupRule = $accessRules | Where-Object { $_.IdentityReference -eq $GroupName }
        
        if ($groupRule) {
            $acl.RemoveAccessRule($groupRule) | Out-Null
            try {
                Set-Acl -Path $Path -AclObject $acl -ErrorAction Stop
                Write-Log "Successfully removed group '$GroupName' from path '$Path'" -Type "SUCCESS"
                return $true
            }
            catch {
                Write-Log "Failed to remove group '$GroupName' from path '$Path'. Error: $_" -Type "ERROR"
                return $false
            }
        } else {
            Write-Log "Group '$GroupName' not found in path '$Path'" -Type "WARNING"
            return $false
        }
    }
    catch {
        Write-Log "Error occurred: $_" -Type "ERROR"
        return $false
    }
}

# Function to process user choice
function Process-UserChoice {
    param(
        [string]$UserChoice,
        [array]$ExistingGroups,
        [string]$Path
    )

    if ($UserChoice -eq 'all') {
        foreach ($group in $ExistingGroups) {
            Write-Log "Processing group: $($group.IdentityReference)" -Type "INFO"
            $result = Remove-GroupFromPath -Path $Path -GroupName $group.IdentityReference

            if ($result) {
                Write-Log "Operation completed successfully for path: $Path, group: $($group.IdentityReference)" -Type "SUCCESS"
            } else {
                Write-Log "Operation failed for path: $Path, group: $($group.IdentityReference)" -Type "ERROR"
            }
        }
    }
    elseif ($UserChoice -ne 'none') {
        # Split the input by commas and process each choice
        $choices = $UserChoice -split ',' | ForEach-Object { $_.Trim() }
        
        foreach ($choice in $choices) {
            # Try to parse as number first
            if ([int]::TryParse($choice, [ref]$null)) {
                $index = [int]$choice
                if ($index -gt 0 -and $index -le $ExistingGroups.Count) {
                    $selectedGroup = $ExistingGroups[$index - 1]
                    Write-Log "Processing group: $($selectedGroup.IdentityReference)" -Type "INFO"
                    $result = Remove-GroupFromPath -Path $Path -GroupName $selectedGroup.IdentityReference

                    if ($result) {
                        Write-Log "Operation completed successfully for path: $Path, group: $($selectedGroup.IdentityReference)" -Type "SUCCESS"
                    } else {
                        Write-Log "Operation failed for path: $Path, group: $($selectedGroup.IdentityReference)" -Type "ERROR"
                    }
                }
            }
            else {
                # Try to find group by name
                $selectedGroup = $ExistingGroups | Where-Object { $_.IdentityReference -like "*$choice*" }
                if ($selectedGroup) {
                    Write-Log "Processing group: $($selectedGroup.IdentityReference)" -Type "INFO"
                    $result = Remove-GroupFromPath -Path $Path -GroupName $selectedGroup.IdentityReference

                    if ($result) {
                        Write-Log "Operation completed successfully for path: $Path, group: $($selectedGroup.IdentityReference)" -Type "SUCCESS"
                    } else {
                        Write-Log "Operation failed for path: $Path, group: $($selectedGroup.IdentityReference)" -Type "ERROR"
                    }
                }
                else {
                    Write-Log "No group found matching '$choice'" -Type "WARNING"
                }
            }
        }
    }
    else {
        Write-Log "Skipping groups for path: $Path" -Type "INFO"
    }
}

# Main script execution
try {
    # Define the directory containing input files
    $inputDir = "C:\temp\Powershell"
    $pathsFile = Join-Path $inputDir "paths.txt"
    $groupsFile = Join-Path $inputDir "groups.txt"

    # Check if input files exist
    if (-not (Test-Path $pathsFile)) {
        Write-Log "paths.txt file not found at $pathsFile. Please create it with the paths to process." -Type "ERROR"
        exit 1
    }
    if (-not (Test-Path $groupsFile)) {
        Write-Log "groups.txt file not found at $groupsFile. Please create it with the groups to remove." -Type "ERROR"
        exit 1
    }

    # Read paths and groups from files
    $paths = Get-Content $pathsFile | Where-Object { $_ -ne "" }  # Skip empty lines
    $groups = Get-Content $groupsFile | Where-Object { $_ -ne "" }  # Skip empty lines

    if ($paths.Count -eq 0) {
        Write-Log "No paths found in $pathsFile" -Type "ERROR"
        exit 1
    }
    if ($groups.Count -eq 0) {
        Write-Log "No groups found in $groupsFile" -Type "ERROR"
        exit 1
    }

    # Process each path
    foreach ($path in $paths) {
        Write-Log "`nChecking groups in path: $path" -Type "INFO"
        $existingGroups = Get-GroupsInPath -Path $path
        
        if ($existingGroups) {
            Write-Log "`nWhich groups would you like to remove? (Enter numbers or group names separated by commas, or 'all' for all groups, or 'none' to skip)" -Type "INFO"
            $userChoice = Read-Host "Your choice"
            
            Process-UserChoice -UserChoice $userChoice -ExistingGroups $existingGroups -Path $path
        }
    }

    Write-Log "All operations completed" -Type "SUCCESS"
}
catch {
    Write-Log "An unexpected error occurred: $_" -Type "ERROR"
    exit 1
}
