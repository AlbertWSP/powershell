function TakeScreenshot {
    Param(
        [string] $ObjectID,
        [int] $Left,
        [int] $Top,
        [int] $Width,
        [int] $Height
    )

    Add-Type -AssemblyName System.Windows.Forms
    Add-Type -AssemblyName System.Drawing

    try {
        $bitmap  = New-Object System.Drawing.Bitmap $Width, $Height
        $graphic = [System.Drawing.Graphics]::FromImage($bitmap)
        $graphic.CopyFromScreen($Left, $Top, 0, 0, $bitmap.Size)

        $FileName = Join-Path $PSScriptRoot "${ObjectID}_EA12.png"
        $bitmap.Save($FileName, "PNG")
    } catch {
        Write-Error "Failed to take or save screenshot. More Info: $_"
    }
}

function SetFullScreen {
    Add-Type -AssemblyName System.Windows.Forms
    [System.Windows.Forms.SendKeys]::SendWait("{F11}")
}

function GetActiveWindowScreenBounds {
    Add-Type @"
        using System;
        using System.Runtime.InteropServices;
        public class UserWindow {
            [DllImport("user32.dll")]
            public static extern IntPtr GetForegroundWindow();

            [DllImport("user32.dll")]
            public static extern bool GetWindowRect(IntPtr hWnd, out RECT lpRect);
        }

        public struct RECT {
            public int Left;
            public int Top;
            public int Right;
            public int Bottom;
        }
"@

    $activeHandle = [UserWindow]::GetForegroundWindow()
    $rect = New-Object RECT
    [UserWindow]::GetWindowRect($activeHandle, [ref]$rect) | Out-Null

    return $rect
}

# Path to the file containing the list of email addresses
$emailListPath = Join-Path $PSScriptRoot "email_list.txt"

# Read the email addresses from the file
$emailAddresses = Get-Content -Path $emailListPath

# Set the active window to full screen
SetFullScreen

foreach ($email in $emailAddresses) {
    try {
        # Get the screen bounds of the active window
        $rect = GetActiveWindowScreenBounds
        $captureLeft = $rect.Left
        $captureTop = 1100
        $captureWidth = 500
        $captureHeight = 500

        # Call the TakeScreenshot function with the email as the parameter and capture area
        TakeScreenshot -ObjectID $email -Left $captureLeft -Top $captureTop -Width $captureWidth -Height $captureHeight

        # Retrieve the AD object
        $obj = Get-ADObject -Filter {Mail -eq $email} -Properties objectClass, samaccountname
        if (-not $obj) {
            throw "Object not found for email: $email"
        }

        if ($obj.objectClass -eq "user") {
            write-host "###### Basic Info ###### " -ForegroundColor Cyan
            get-aduser -identity $obj.samaccountname -Properties * | Select-Object DisplayName, 
            @{Label="LastName";Expression={$_.sn}}, @{Label="FirstName";Expression={$_.GivenName}}, sAMAccountName, 
            @{Label="UPN Logon";Expression={$_.userPrincipalName}}, @{Label="FullName";Expression={$_.Name}}, Mail, Description, 
            @{Label="ExpireDate";Expression={[datetime]::FromFileTime($_.accountExpires)}}, Enabled, 
            @{Label="EmployeeType";Expression={$_.EmployeeType}},extensionAttribute12| format-list
            # Wait for 1 second before taking another screenshot
            Start-Sleep -Seconds 1
            TakeScreenshot -ObjectID $email -Left $captureLeft -Top $captureTop -Width $captureWidth -Height $captureHeight
        } elseif ($obj.objectClass -eq "group") {
            $obj | Select-Object samaccountname, mail, displayname, description, extensionAttribute9, distinguishedName, grouptype, managedby | format-list
            write-host "`n####### Group member(s) of ", $obj.samaccountname ," #######`n" -ForegroundColor Cyan
            get-adgroupmember -identity $obj.samaccountname | format-wide -column 3
            write-host "`n####### End #######`n" -ForegroundColor Cyan
        } else {
            write-host "`n####### No Record!!! #######`n" -ForegroundColor Red
        }
    } catch {
        Write-Error "An error occurred: $_"
    }
}