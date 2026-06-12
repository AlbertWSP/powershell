<#
.SYNOPSIS
    AD Security Group Manager (GUI) - Version 2.7

.DESCRIPTION
    WinForms GUI 工具，用作：
      1) 檢視 Folder ACL 權限 (ACL Viewer)
      2) 查閱 AD Group 成員 (Audit Members)
      3) 批量新增/移除 AD Group 成員 (Bulk Updates)
      4) 檢查指定 User 對 Folder 的有效權限 (Access Checker)

    重要說明（UI 版面修正）：
    - 本工具使用 WinForms + 手動座標 (Location) 方式建立 UI。
    - 在某些 DPI/縮放/Tab 切換的 Layout pass 下，TextBox 可能會被 WinForms 自動重新計算尺寸，
      造成「TextBox 向右無限伸展，右邊按鈕被擠走/消失」。
    - 因此本版本加入：
        (a) 對關鍵 TextBox 設定 AutoSize = $false
        (b) 在 TabPage.Enter 事件觸發後再重新計算 TextBox.Width
      以確保按鈕永遠可見，版面穩定。

.PREREQUISITES
    - PowerShell 5.1+
    - RSAT ActiveDirectory module
    - 如需 Excel export：本機需安裝 Microsoft Excel（COM）

.NOTES
    Author : Albert Ng
    Updated: 2026-03-18

#>

# =================================================================================
# 1. Self-Elevation (Request Administrator Privileges)
# =================================================================================
# 目的：
#   - 部分 AD 操作/ACL 讀取可能需要 elevated 權限（視環境而定）。
#   - 若非 Admin，會自動重新以 RunAs 啟動同一份腳本。

if (-not ([Security.Principal.WindowsPrincipal][Security.Principal.WindowsIdentity]::GetCurrent()).IsInRole(
        [Security.Principal.WindowsBuiltInRole]::Administrator
    )) {

    $newProcess = New-Object System.Diagnostics.ProcessStartInfo 'PowerShell'
    $newProcess.Arguments = "-NoProfile -ExecutionPolicy Bypass -File `"$PSCommandPath`""
    $newProcess.Verb = 'runas'

    try {
        [System.Diagnostics.Process]::Start($newProcess) | Out-Null
    }
    catch {
        # 使用者取消 UAC，不再繼續
    }

    exit
}

# =================================================================================
# 2. Dependencies / Module / Membership Check
# =================================================================================
# 目的：
#   - 載入 WinForms 所需 assembly
#   - 檢查 RSAT AD module 是否存在
#   - 依公司支援群組 (Support Group) 判斷是否具備操作權限（如無權限，仍可繼續但改動操作可能失敗）

Add-Type -AssemblyName System.Windows.Forms
Add-Type -AssemblyName System.Drawing

if (-not (Get-Module -ListAvailable ActiveDirectory)) {
    [System.Windows.Forms.MessageBox]::Show("Active Directory module is required. Please install RSAT.")
    exit
}

Import-Module ActiveDirectory

# =================================================================================
# 2.1 Helper Functions
# =================================================================================

function Set-TextboxWidthToLeaveRoom {
    <#
    .SYNOPSIS
        計算 TextBox 寬度，確保右邊按鈕永遠可見。

    .DESCRIPTION
        在 WinForms 使用「固定座標 + Anchor」混合時，
        可能出現 TextBox 往右伸展，覆蓋右邊按鈕。
        此 helper 會用「右邊控件最左 X」去計算 TextBox.Width。

        注意：
        - 只適合「TextBox 在左，按鈕在右」同一行的佈局
        - 要配合 TextBox.AutoSize = $false，否則 WinForms 仍可能覆蓋寬度

    .PARAMETER TextBox
        需要被計算寬度的 TextBox 控件

    .PARAMETER RightControls
        TextBox 右邊的控件（通常是 Button）陣列

    .PARAMETER Gap
        TextBox 與右邊控件之間的間距

    .EXAMPLE
        Set-TextboxWidthToLeaveRoom -TextBox $txtGroupSearch -RightControls @($btnFetch,$btnExport,$btnClear)
    #>
    param(
        [Parameter(Mandatory)][System.Windows.Forms.Control]$TextBox,
        [Parameter(Mandatory)][System.Windows.Forms.Control[]]$RightControls,
        [int]$RightPadding = 20,
        [int]$Gap = 10
    )

    $minRightX = ($RightControls | Measure-Object -Property Left -Minimum).Minimum
    $newWidth = $minRightX - $TextBox.Left - $Gap
    if ($newWidth -lt 120) { $newWidth = 120 }
    $TextBox.Width = $newWidth
}

function Resolve-ADGroupByNameOrMail {
    <#
    .SYNOPSIS
        以 Group Name 或 mail 屬性解析 AD Group。
    #>
    param([string]$SearchValue)

    $val = $SearchValue.Trim()
    Get-ADGroup -Filter { Name -eq $val -or mail -eq $val } -ErrorAction Stop
}

function Resolve-ADUserByMailOrSam {
    <#
    .SYNOPSIS
        以 User mail 或 SamAccountName 解析 AD User。
    #>
    param([string]$SearchValue)

    $val = $SearchValue.Trim()
    Get-ADUser -Filter { mail -eq $val -or SamAccountName -eq $val } -ErrorAction Stop
}

function Update-Status {
    <#
    .SYNOPSIS
        更新底部 StatusStrip 的訊息並立即刷新 UI。
    #>
    param([string]$Message)

    $statusLabel.Text = $Message
    [System.Windows.Forms.Application]::DoEvents()
}

# =================================================================================
# 2.2 Export Helpers (Excel COM)
# =================================================================================
# 目的：
#   - 將 DataTable 安全地輸出到 Excel（.xlsx）
#   - 包括 COM cleanup，避免 Excel.exe 背景殘留

function Export-DataTableToExcel {
    param(
        [Parameter(Mandatory = $true)]
        [System.Data.DataTable]$DataTable,

        [Parameter(Mandatory = $true)]
        [string]$FilePath,

        [string]$WorksheetName = "ACL Viewer",
        [string]$FolderPath = "",
        [string]$Owner = ""
    )

    $excel = $null
    $workbook = $null
    $worksheet = $null
    $saveSucceeded = $false

    try {
        $excel = New-Object -ComObject Excel.Application
        $excel.Visible = $false
        $excel.DisplayAlerts = $false

        $workbook = $excel.Workbooks.Add()
        $worksheet = $workbook.Worksheets.Item(1)

        if ($WorksheetName.Length -gt 31) {
            $WorksheetName = $WorksheetName.Substring(0, 31)
        }
        $worksheet.Name = $WorksheetName

        # 匯出標頭資訊
        $worksheet.Cells.Item(1, 1) = "ACL Viewer Export"
        $worksheet.Cells.Item(2, 1) = "Folder Path"
        $worksheet.Cells.Item(2, 2) = $FolderPath
        $worksheet.Cells.Item(3, 1) = "Owner"
        $worksheet.Cells.Item(3, 2) = $Owner
        $worksheet.Cells.Item(4, 1) = "Export Time"
        $worksheet.Cells.Item(4, 2) = (Get-Date).ToString("yyyy-MM-dd HH:mm:ss")

        $worksheet.Range("A1:B1").Merge() | Out-Null
        $worksheet.Range("A1").Font.Bold = $true
        $worksheet.Range("A1").Font.Size = 14

        # 表格 header
        $startRow = 6
        for ($c = 0; $c -lt $DataTable.Columns.Count; $c++) {
            $worksheet.Cells.Item($startRow, $c + 1) = $DataTable.Columns[$c].ColumnName
        }

        $headerRange = $worksheet.Range(
            $worksheet.Cells.Item($startRow, 1),
            $worksheet.Cells.Item($startRow, $DataTable.Columns.Count)
        )
        $headerRange.Font.Bold = $true
        $headerRange.Interior.ColorIndex = 15

        # 表格資料
        for ($r = 0; $r -lt $DataTable.Rows.Count; $r++) {
            for ($c = 0; $c -lt $DataTable.Columns.Count; $c++) {
                $value = $DataTable.Rows[$r][$c]
                if ($null -eq $value) { $value = "" }
                $worksheet.Cells.Item($startRow + 1 + $r, $c + 1) = $value.ToString()
            }
        }

        # 邊框 + AutoFit
        $lastRow = $startRow + $DataTable.Rows.Count
        $lastCol = $DataTable.Columns.Count
        $usedRange = $worksheet.Range(
            $worksheet.Cells.Item($startRow, 1),
            $worksheet.Cells.Item($lastRow, $lastCol)
        )
        $usedRange.Borders.LineStyle = 1
        $worksheet.Columns.AutoFit() | Out-Null

        # SaveAs: 51 = xlOpenXMLWorkbook (.xlsx)
        $workbook.SaveAs($FilePath, 51)
        $saveSucceeded = $true
    }
    catch {
        throw "Excel SaveAs failed: $($_.Exception.Message)"
    }
    finally {
        # COM cleanup（避免 Excel.exe 殘留）
        if ($worksheet) {
            try { [void][System.Runtime.InteropServices.Marshal]::FinalReleaseComObject($worksheet) } catch {}
            $worksheet = $null
        }
        if ($workbook) {
            try { $workbook.Close($false) } catch {}
            try { [void][System.Runtime.InteropServices.Marshal]::FinalReleaseComObject($workbook) } catch {}
            $workbook = $null
        }
        if ($excel) {
            try { $excel.Quit() } catch {}
            try { [void][System.Runtime.InteropServices.Marshal]::FinalReleaseComObject($excel) } catch {}
            $excel = $null
        }

        [GC]::Collect()
        [GC]::WaitForPendingFinalizers()
        [GC]::Collect()
        [GC]::WaitForPendingFinalizers()
    }

    if (-not $saveSucceeded -or -not (Test-Path $FilePath)) {
        throw "Excel export did not complete successfully. File was not created."
    }
}

# =================================================================================
# 2.3 Support Group Membership Check
# =================================================================================

$supportGroupName = "GRP-RBC-R-ASI-WSP-ClientSideSupport"
$isSupportMember = $false

try {
    $currentUserSam = [Security.Principal.WindowsIdentity]::GetCurrent().Name.Split('\')[-1]
    $groupMembers = Get-ADGroupMember -Identity $supportGroupName -Recursive |
        Select-Object -ExpandProperty SamAccountName

    if ($groupMembers -contains $currentUserSam) {
        $isSupportMember = $true
    }
}
catch {
    [System.Windows.Forms.MessageBox]::Show(
        "Warning: Could not verify membership in '$supportGroupName'.`n`nError: $($_.Exception.Message)",
        "Membership Check Warning",
        "OK",
        "Warning"
    )
}

if (-not $isSupportMember) {
    $msgText = @"
PERMISSION WARNING:

You are currently NOT a member of '$supportGroupName'.

You may 'Search' groups, but 'Modify' operations will likely fail due to restricted access.

Do you want to continue?
"@

    if ([System.Windows.Forms.MessageBox]::Show($msgText, "Access Check", "YesNo", "Warning") -eq "No") {
        exit
    }
}

# =================================================================================
# 3. UI Interface Code
# =================================================================================
# 目的：
#   - 建立主視窗 + TabControl + 四個功能 Tab
#   - 設定 Anchors / AutoSize / 事件（Click / KeyDown / Enter）

$form = New-Object System.Windows.Forms.Form
$form.Text = "AD Security Group Manager - Version 2.5"
$form.Size = New-Object System.Drawing.Size(850, 1000)
$form.StartPosition = "CenterScreen"

# Status Bar
$statusStrip = New-Object System.Windows.Forms.StatusStrip
$statusLabel = New-Object System.Windows.Forms.ToolStripStatusLabel
$statusLabel.Text = "Ready"
$statusStrip.Items.Add($statusLabel) | Out-Null
$form.Controls.Add($statusStrip)

# Main Container
$tabControl = New-Object System.Windows.Forms.TabControl
$tabControl.Size = New-Object System.Drawing.Size(830, 880)
$tabControl.Anchor = "Top, Bottom, Left, Right"
$tabControl.Location = New-Object System.Drawing.Point(5, 5)
$form.Controls.Add($tabControl)

# --------------------------
# TAB: ACL VIEWER
# --------------------------
$tabACL = New-Object System.Windows.Forms.TabPage
$tabACL.Text = "ACL Viewer"
$tabControl.TabPages.Add($tabACL)

$lblPath = New-Object System.Windows.Forms.Label
$lblPath.Location = "20, 10"
$lblPath.Size = "100, 20"
$lblPath.Text = "Folder Path:"
$tabACL.Controls.Add($lblPath)

$txtPath = New-Object System.Windows.Forms.TextBox
$txtPath.Location = "20, 35"
$txtPath.Size = "790, 25"
$txtPath.Anchor = "Top, Left, Right"
$txtPath.Add_KeyDown({
        param($s, $e)
        if ($e.Control -and $e.KeyCode -eq [System.Windows.Forms.Keys]::A) {
            $s.SelectAll()
            $e.SuppressKeyPress = $true
            $e.Handled = $true
        }
    })
$tabACL.Controls.Add($txtPath)

# (略) ... 你其餘 ACL Viewer controls/logic 照原本放落去

# --------------------------
# TAB: AUDIT MEMBERS（關鍵：防止 TextBox 拉長遮按鈕）
# --------------------------
$tabView = New-Object System.Windows.Forms.TabPage
$tabView.Text = "Audit Members"
$tabControl.TabPages.Add($tabView)

# (略) ... hint panel 照原本

$txtGroupSearch = New-Object System.Windows.Forms.TextBox
$txtGroupSearch.Location = "20, 170"
$txtGroupSearch.Width = 350
$txtGroupSearch.Anchor = "Top, Left"
$txtGroupSearch.AutoSize = $false  # ✅ 關閉 AutoSize，避免 layout pass 覆蓋 Width
$tabView.Controls.Add($txtGroupSearch)

$btnFetch = New-Object System.Windows.Forms.Button
$btnFetch.Text = "Fetch"
$btnFetch.Location = "380, 168"
$btnFetch.Width = 80
$btnFetch.Anchor = "Top, Right"
$tabView.Controls.Add($btnFetch)

$btnExport = New-Object System.Windows.Forms.Button
$btnExport.Text = "Export CSV"
$btnExport.Location = "470, 168"
$btnExport.Width = 100
$btnExport.Enabled = $false
$btnExport.Anchor = "Top, Right"
$tabView.Controls.Add($btnExport)

$btnClearTab1 = New-Object System.Windows.Forms.Button
$btnClearTab1.Text = "Clear All"
$btnClearTab1.Location = "580, 168"
$btnClearTab1.Width = 100
$btnClearTab1.Anchor = "Top, Right"
$tabView.Controls.Add($btnClearTab1)

# ✅ 第一次 render 後先計一次（初始）
Set-TextboxWidthToLeaveRoom -TextBox $txtGroupSearch -RightControls @($btnFetch, $btnExport, $btnClearTab1)

# ✅ Tab 切換進入時再計一次（WinForms 會重新 layout）
$tabView.Add_Enter({
    Set-TextboxWidthToLeaveRoom -TextBox $txtGroupSearch -RightControls @($btnFetch, $btnExport, $btnClearTab1)
})

# --------------------------
# TAB: BULK UPDATES（關鍵：Target group TextBox + Upload button）
# --------------------------
$tabUpdate = New-Object System.Windows.Forms.TabPage
$tabUpdate.Text = "Bulk Updates"
$tabControl.TabPages.Add($tabUpdate)

$txtTargetGroup = New-Object System.Windows.Forms.TextBox
$txtTargetGroup.Location = "20, 195"
$txtTargetGroup.Width = 400
$txtTargetGroup.Anchor = "Top, Left"
$txtTargetGroup.AutoSize = $false  # ✅ 同上，避免 layout 覆蓋 Width
$tabUpdate.Controls.Add($txtTargetGroup)

$btnImportFile = New-Object System.Windows.Forms.Button
$btnImportFile.Text = "Upload CSV/TXT"
$btnImportFile.Location = "430, 193"
$btnImportFile.Width = 130
$btnImportFile.Anchor = "Top, Right"
$tabUpdate.Controls.Add($btnImportFile)

Set-TextboxWidthToLeaveRoom -TextBox $txtTargetGroup -RightControls @($btnImportFile)

$tabUpdate.Add_Enter({
    Set-TextboxWidthToLeaveRoom -TextBox $txtTargetGroup -RightControls @($btnImportFile)
})

# ---------------------------------------------------------------------------------
# 其他 Tab（Access Checker / ACL Export / Logic Handling）照原本放落去即可
# ---------------------------------------------------------------------------------

# 建議：首次顯示後做一次全域對齊（防 DPI / 字體延遲）
$form.Add_Shown({
    try {
        Set-TextboxWidthToLeaveRoom -TextBox $txtGroupSearch -RightControls @($btnFetch, $btnExport, $btnClearTab1)
        Set-TextboxWidthToLeaveRoom -TextBox $txtTargetGroup -RightControls @($btnImportFile)
    } catch {}
})

$form.ShowDialog()