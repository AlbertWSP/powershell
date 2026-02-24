# Method 1: Using the call operator (&)
Write-Host "Method 1: Using the call operator"
& ".\ChildScript.ps1" -Name "John"

# Method 2: Using Invoke-Expression
Write-Host "`nMethod 2: Using Invoke-Expression"
$command = ".\ChildScript.ps1 -Name 'Jane'"
Invoke-Expression $command

# Method 3: Using Start-Process
Write-Host "`nMethod 3: Using Start-Process"
Start-Process powershell.exe -ArgumentList "-File `".\ChildScript.ps1`" -Name `"Bob`""

# Method 4: Using dot sourcing
Write-Host "`nMethod 4: Using dot sourcing"
. ".\ChildScript.ps1" -Name "Alice"