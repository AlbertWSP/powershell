# Ensure the Active Directory module for PowerShell is installed
Import-Module ActiveDirectory

# Path to the text file containing the list of computer names
$computerListPath = "C:\temp\ComputerList.txt"

# Import the list of computer names from the text file
$computerNames = Get-Content -Path $computerListPath

# Initialize an array to store the results
$results = @()

# Function to convert Windows build number to a readable format
function ConvertWindowsBuild {
    param(
        [string] $OperatingSystem,
        [string] $OperatingSystemVersion
    )
    $WinBuilds = @{
        '10.0 (22621)' = "Windows 11 22H2"
        '10.0 (19045)' = "Windows 10 22H2"
        '10.0 (22000)' = "Windows 11 21H2"
        '10.0 (19044)' = "Windows 10 21H2"
        '10.0 (19043)' = "Windows 10 21H1"
        '10.0 (19042)' = "Windows 10 20H2"
        '10.0 (18362)' = "Windows 10 1903"
        '10.0 (17763)' = "Windows 10 1809"
        '10.0 (17134)' = "Windows 10 1803"
        '10.0 (16299)' = "Windows 10 1709"
        '10.0 (15063)' = "Windows 10 1703"
        '10.0 (14393)' = "Windows 10 1607"
        '10.0 (26100)' = "Windows 11 24H2"
        '10.0 (26200)' = "Windows 11 25H2"
        '10.0 (22631)' = "Windows 11 23H2"
    }
    $WinBuild = $WinBuilds[$OperatingSystemVersion]
    if ($WinBuild) {
        $WinBuild
    } else {
        'Unknown'
    }
}

# Loop through each computer name and get details from AD
foreach ($computerName in $computerNames) {
    $computer = Get-ADComputer -Identity $computerName -Properties Name, OperatingSystem, OperatingSystemVersion, IPv4Address, LastLogonDate
    if ($computer) {
        $results += [PSCustomObject]@{
            Name = $computer.Name
            IPv4Address = $computer.IPv4Address
            OperatingSystem = $computer.OperatingSystem
            Build = ConvertWindowsBuild -OperatingSystem $computer.OperatingSystem -OperatingSystemVersion $computer.OperatingSystemVersion
            LastLogonDate = $computer.LastLogonDate
        }
    }
}

# Generate a timestamp for the filename
$timestamp = Get-Date -Format "yyyyMMdd_HHmmss"

# Path to the CSV file where the results will be exported
$csvPath = "C:\temp\Output_win11_$timestamp.csv"

# Export the results to a CSV file
$results | Export-Csv -Path $csvPath -NoTypeInformation -Encoding UTF8

Write-Host "Export completed. Check the CSV file at $csvPath"