# 1. 定義你想查詢的子群組名稱
$TargetGroupName = "GRP-FSU-HKKWN200-SSAE-PM"

# 2. 取得該群組的物件
$GroupObj = Get-ADGroup -Identity $TargetGroupName

# 3. 使用 LDAP 過濾器進行遞迴搜尋 (找出所有上層父群組)
$ParentGroups = Get-ADGroup -LDAPFilter "(member:1.2.840.113556.1.4.1941:=$($GroupObj.DistinguishedName))"

# 4. 輸出結果
$ParentGroups | Select-Object Name, DistinguishedName | Format-Table -AutoSize
