Add-Type -AssemblyName System.Windows.Forms
Add-Type -AssemblyName System.Drawing

# 創建表單
$Form = New-Object System.Windows.Forms.Form
$Form.Text = "PowerShell GUI"
$Form.Size = New-Object System.Drawing.Size(400, 300)
$Form.StartPosition = "CenterScreen"

# 按鈕 1
$Button1 = New-Object System.Windows.Forms.Button
$Button1.Text = "執行腳本 1"
$Button1.Size = New-Object System.Drawing.Size(120, 30)
$Button1.Location = New-Object System.Drawing.Point(20, 20)
$Button1.Add_Click({
    Start-Process -FilePath "powershell.exe" -ArgumentList "-File `"C:\temp\Powershell\DL_group\check the access right in the path.ps1`""
})
$Form.Controls.Add($Button1)

# 按鈕 2
$Button2 = New-Object System.Windows.Forms.Button
$Button2.Text = "執行腳本 2"
$Button2.Size = New-Object System.Drawing.Size(120, 30)
$Button2.Location = New-Object System.Drawing.Point(20, 60)
$Button2.Add_Click({
    Start-Process -FilePath "powershell.exe" -ArgumentList "-File `".\Script2.ps1`""
})
$Form.Controls.Add($Button2)

# 按鈕 3
$Button3 = New-Object System.Windows.Forms.Button
$Button3.Text = "執行腳本 3"
$Button3.Size = New-Object System.Drawing.Size(120, 30)
$Button3.Location = New-Object System.Drawing.Point(20, 100)
$Button3.Add_Click({
    Start-Process -FilePath "powershell.exe" -ArgumentList "-File `".\Script3.ps1`""
})
$Form.Controls.Add($Button3)

# 顯示表單
[void] $Form.ShowDialog()
