function Get-FolderSize {
    param (
        [Parameter(Mandatory=$true)]
        [string]$FolderPath
    )
    
    $size = (Get-ChildItem -Path $FolderPath -Recurse -Force | Measure-Object -Property Length -Sum).Sum
    
    if ($size -ge 1GB) {
        return "{0:N2} GB" -f ($size / 1GB)
    } elseif ($size -ge 1MB) {
        return "{0:N2} MB" -f ($size / 1MB)
    } elseif ($size -ge 1KB) {
        return "{0:N2} KB" -f ($size / 1KB)
    } else {
        return "$size Bytes"
    }
}

$folderPath = "\\10.35.33.93\c$\Users\Ella.Lee\Documents"
$size = Get-FolderSize -FolderPath $folderPath
Write-Host "Folder size: $size"