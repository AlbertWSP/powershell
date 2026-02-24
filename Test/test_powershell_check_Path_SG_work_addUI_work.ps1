Add-Type -AssemblyName System.Windows.Forms
Add-Type -AssemblyName System.Drawing

# 1. 建立主視窗
$form = New-Object System.Windows.Forms.Form
$form.Text = "ACL Permissions Viewer"
$form.Size = New-Object System.Drawing.Size(850, 600)
$form.StartPosition = "CenterScreen"

# 2. 路徑輸入標籤
$label = New-Object System.Windows.Forms.Label
$label.Location = New-Object System.Drawing.Point(15, 15)
$label.Size = New-Object System.Drawing.Size(200, 20)
$label.Text = "Folder Path:"
$form.Controls.Add($label)

# 3. 路徑輸入框
$textBox = New-Object System.Windows.Forms.TextBox
$textBox.Location = New-Object System.Drawing.Point(15, 40)
$textBox.Size = New-Object System.Drawing.Size(680, 25)
$textBox.Text = "\\dcasi300dat01\DCASI300\DAT08\SSAE\Proposals\ICT\2025\WSP-SSAE-O-250126_AAT MHS"
$form.Controls.Add($textBox)

# 4. 瀏覽按鈕
$btnBrowse = New-Object System.Windows.Forms.Button
$btnBrowse.Location = New-Object System.Drawing.Point(710, 38)
$btnBrowse.Size = New-Object System.Drawing.Size(100, 28)
$btnBrowse.Text = "Browse..."
$btnBrowse.Add_Click({
    $browser = New-Object System.Windows.Forms.FolderBrowserDialog
    if($browser.ShowDialog() -eq "OK") { $textBox.Text = $browser.SelectedPath }
})
$form.Controls.Add($btnBrowse)

# 5. 結果顯示區域 (DataGridView)
$dataGridView = New-Object System.Windows.Forms.DataGridView
$dataGridView.Location = New-Object System.Drawing.Point(15, 130)
$dataGridView.Size = New-Object System.Drawing.Size(800, 350)
$dataGridView.ReadOnly = $true
$dataGridView.AutoSizeColumnsMode = "AllCells"
$dataGridView.BackgroundColor = "White"
$dataGridView.RowHeadersVisible = $false
$form.Controls.Add($dataGridView)

# 6. 擁有者資訊標籤
$lblOwner = New-Object System.Windows.Forms.Label
$lblOwner.Location = New-Object System.Drawing.Point(15, 500)
$lblOwner.Size = New-Object System.Drawing.Size(800, 25)
$lblOwner.ForeColor = "DarkBlue"
$form.Controls.Add($lblOwner)

# 7. 執行按鈕 (點擊後直接更新 UI)
$btnCheck = New-Object System.Windows.Forms.Button
$btnCheck.Location = New-Object System.Drawing.Point(15, 80)
$btnCheck.Size = New-Object System.Drawing.Size(150, 35)
$btnCheck.Text = "Check Permissions"
$btnCheck.BackColor = "LightGreen"

$btnCheck.Add_Click({
    $path = $textBox.Text
    if (Test-Path $path) {
        try {
            $acl = Get-Acl -Path $path
            $lblOwner.Text = "Owner: $($acl.Owner)"
            
            # 將資料轉換為 DataTable 以便在 GridView 顯示
            $results = $acl.Access | Select-Object IdentityReference, FileSystemRights, AccessControlType
            
            # 轉換資料格式
            $dt = New-Object System.Data.DataTable
            $dt.Columns.Add("IdentityReference")
            $dt.Columns.Add("FileSystemRights")
            $dt.Columns.Add("AccessControlType")
            
            foreach ($item in $results) {
                $row = $dt.NewRow()
                $row.IdentityReference = $item.IdentityReference.ToString()
                $row.FileSystemRights = $item.FileSystemRights.ToString()
                $row.AccessControlType = $item.AccessControlType.ToString()
                $dt.Rows.Add($row)
            }
            $dataGridView.DataSource = $dt
        } catch {
            [System.Windows.Forms.MessageBox]::Show("Access Denied or Error: $($_.Exception.Message)")
        }
    } else {
        [System.Windows.Forms.MessageBox]::Show("Invalid Path!")
    }
})
$form.Controls.Add($btnCheck)

# 顯示視窗
$form.ShowDialog()
