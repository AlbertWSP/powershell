🏷️ 命名与备注一致性框架

1. 命名规范 (Filename)格式： Verb-Noun.ps1规则： 文件名必须与脚本内主函数的名称（或 .SYNOPSIS 中的核心动作）完全匹配。示例： Get-SystemHealth.ps1 或 New-UserAccount.ps1。


2. 内部备注框架 (Internal Template)将以下代码块置于脚本最顶部。这是 PowerShell 官方标准的 Comment-Based Help 格式。powershell<#
.SYNOPSIS
    [动作-对象] - 简短描述脚本功能 (例如: 获取系统健康状态报告)

.DESCRIPTION
    详细描述脚本具体执行了什么逻辑，解决了什么问题。

.PARAMETER <ParameterName>
    描述输入参数的用途及其对脚本的影响。

.EXAMPLE
    .\Verb-Noun.ps1 -ParameterName "Value"
    展示一个实际运行的例子。

.NOTES
    文件名：Verb-Noun.ps1 (必须与此处一致)
    创建日期：YYYY-MM-DD
    作者：Your Name
#>

param (
    # 这里定义参数
)
Use code with caution.

🛠️ 分类命名建议 (Naming Groups)

类别文件名前缀 (Approved Verbs)内部备注备注示例查询/读取Get- (例: Get-DiskUsage.ps1).SYNOPSIS 获取磁盘占用情况配置/修改Set- (例: Set-AppConfig.ps1).SYNOPSIS 修改应用程序配置新建/创建New- (例: New-BackupFile.ps1).SYNOPSIS 创建新的备份文件移除/删除Remove- (例: Remove-OldLogs.ps1).SYNOPSIS 删除过期日志文件执行/启动Invoke- (例: Invoke-Update.ps1).SYNOPSIS 触发系统更新程序💡 


为什么这样做？
自解释性： 看到文件名就知道用法，不需要打开文件。
系统集成： 这种备注格式支持 Get-Help .\YourScript.ps1 命令，能像系统自带命令一样显示说明文档。
标准统一： 方便以后将多个 .ps1 脚本快速整合进一个 Module (.psm1) 中。