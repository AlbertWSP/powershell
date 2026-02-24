
# Define input and output file paths
$InputPath  = "C:\Temp\groups_.txt"   # <-- Change this to your TXT file path
$OutputPath = "C:\Temp\groups.csv"   # <-- Change this to your desired CSV path

# Regex pattern: two or more spaces separate columns
$pattern = '^(?<GroupEmail>\S+)\s{2,}(?<Group>.+?)\s*$'

# Read and parse file
$rows = @()
Get-Content -Path $InputPath | ForEach-Object {
    $line = $_.Trim()
    if ([string]::IsNullOrWhiteSpace($line)) { return }

    # Skip header line if it contains "GroupEmail" and "Group"
    if ($line -match 'groupemail' -and $line -match '(^|\s)group(\s|$)') { return }

    $m = [regex]::Match($line, $pattern)
    if ($m.Success) {
        $rows += [PSCustomObject]@{
            GroupEmail = $m.Groups['GroupEmail'].Value
            Group      = $m.Groups['Group'].Value
        }
    }
}

# Export to CSV
$rows | Export-Csv -Path $OutputPath -NoTypeInformation -Encoding UTF8
Write-Host "CSV file created at: $OutputPath"
