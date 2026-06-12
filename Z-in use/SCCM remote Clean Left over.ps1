$RemoteComputer = "2HM62T3"

Invoke-Command -ComputerName $RemoteComputer -ScriptBlock {
    function Remove-SCCM {
        # Stop SCCM services
        Get-Service -Name CcmExec -ErrorAction SilentlyContinue | Stop-Service -Force -Verbose
        Get-Service -Name ccmsetup -ErrorAction SilentlyContinue | Stop-Service -Force -Verbose

        # Remove folders
        $paths = @("$env:WinDir\CCM", "$env:WinDir\CCMSetup", "$env:WinDir\CCMCache")
        foreach ($path in $paths) {
            if (Test-Path $path) {
                takeown /F $path /R /A /D Y | Out-Null
                Remove-Item -Path $path -Force -Recurse -Confirm:$false -Verbose -ErrorAction SilentlyContinue
            }
        }

        # Remove smscfg.ini
        Remove-Item -Path "$env:WinDir\smscfg.ini" -Force -Confirm:$false -Verbose -ErrorAction SilentlyContinue

        # Remove SCCM certificates
        Remove-Item -Path 'HKLM:\Software\Microsoft\SystemCertificates\SMS\Certificates\*' -Force -Confirm:$false -Verbose -ErrorAction SilentlyContinue

        # Remove registry keys
        $registryPaths = @(
            'HKLM:\SOFTWARE\Microsoft\CCM',
            'HKLM:\SOFTWARE\Wow6432Node\Microsoft\CCM',
            'HKLM:\SOFTWARE\Microsoft\SMS',
            'HKLM:\SOFTWARE\Wow6432Node\Microsoft\SMS',
            'HKLM:\Software\Microsoft\CCMSetup',
            'HKLM:\Software\Wow6432Node\Microsoft\CCMSetup',
            'HKLM:\SYSTEM\CurrentControlSet\Services\CcmExec',
            'HKLM:\SYSTEM\CurrentControlSet\Services\ccmsetup'
        )
        foreach ($regPath in $registryPaths) {
            if (Test-Path $regPath) {
                Remove-Item -Path $regPath -Force -Recurse -Confirm:$false -Verbose -ErrorAction SilentlyContinue
            }
        }

        # Remove SCCM namespaces from WMI
        $wmiNamespaces = @(
            @{ Namespace = "root"; Name = "CCM" },
            @{ Namespace = "root"; Name = "CCMVDI" },
            @{ Namespace = "root"; Name = "SmsDm" },
            @{ Namespace = "root\cimv2"; Name = "sms" }
        )

        foreach ($ns in $wmiNamespaces) {
            try {
                Get-CimInstance -Namespace $ns.Namespace -Query "SELECT * FROM __Namespace WHERE Name='$($ns.Name)'" -ErrorAction SilentlyContinue |
                Remove-CimInstance -Verbose -Confirm:$false -ErrorAction SilentlyContinue
            }
            catch {
                Write-Verbose "WMI namespace $($ns.Name) not found or could not be removed."
            }
        }

        Write-Host "[$env:COMPUTERNAME] All traces of SCCM have been removed"
    }

    # Execute function
    Remove-SCCM
}
