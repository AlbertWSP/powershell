Add-Type -AssemblyName System.Windows.Forms

# Create form
$form = New-Object System.Windows.Forms.Form
$form.Text = "Windows 11 Remote Installer"
$form.Size = New-Object System.Drawing.Size(400, 200)
$form.StartPosition = "CenterScreen"

# Create hostname label and textbox
$labelHostName = New-Object System.Windows.Forms.Label
$labelHostName.Location = New-Object System.Drawing.Point(10, 20)
$labelHostName.Size = New-Object System.Drawing.Size(100, 20)
$labelHostName.Text = "Hostname:"
$form.Controls.Add($labelHostName)

$textBoxHostName = New-Object System.Windows.Forms.TextBox
$textBoxHostName.Location = New-Object System.Drawing.Point(120, 20)
$textBoxHostName.Size = New-Object System.Drawing.Size(200, 20)
$form.Controls.Add($textBoxHostName)

# Create install button
$buttonInstall = New-Object System.Windows.Forms.Button
$buttonInstall.Location = New-Object System.Drawing.Point(120, 60)
$buttonInstall.Size = New-Object System.Drawing.Size(150, 30)
$buttonInstall.Text = "Install Windows 11"
$buttonInstall.Add_Click({
    $hostname = $textBoxHostName.Text
    if ([string]::IsNullOrWhiteSpace($hostname)) {
        [System.Windows.Forms.MessageBox]::Show("請輸入主機名稱", "錯誤")
        return
    }

    try {
        $credential = Get-Credential -Message "輸入遠端連線的憑證"
        $session = New-PSSession -ComputerName $hostname -Credential $credential -ErrorAction Stop

        Invoke-Command -Session $session -ScriptBlock {
            # 建立安裝目錄
            $FolderPath = "C:\WindowsUpgrade"
            if (!(Test-Path -Path $FolderPath)) {
                New-Item -Path $FolderPath -ItemType Directory -Force
            }

            # 下載 Windows 11 安裝助理
            $DownloadUrl = "https://go.microsoft.com/fwlink/?linkid=2171764"
            $InstallerPath = "$FolderPath\Windows11InstallationAssistant.exe"
            
            Write-Host "下載 Windows 11 安裝助理中..."
            Invoke-WebRequest -Uri $DownloadUrl -OutFile $InstallerPath -ErrorAction Stop

            # 開始安裝
            Write-Host "開始安裝 Windows 11..."
            Start-Process -FilePath $InstallerPath `
                         -ArgumentList "/auto upgrade /quiet /noreboot /Compat IgnoreWarning" `
                         -Verb RunAs -Wait -NoNewWindow
        }

        Remove-PSSession $session
        [System.Windows.Forms.MessageBox]::Show("安裝過程已成功完成！", "成功")
    }
    catch {
        [System.Windows.Forms.MessageBox]::Show("錯誤: $_", "安裝失敗")
    }
})
$form.Controls.Add($buttonInstall)

# Show form
$form.ShowDialog()
