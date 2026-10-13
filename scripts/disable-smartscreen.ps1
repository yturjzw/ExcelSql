<#
.SYNOPSIS
    一次性关闭本机 SmartScreen 对应用的「检查应用和文件」，并放行 ExcelSqlConsole 目录。
.DESCRIPTION
    自签名证书无法获得 SmartScreen 的信誉，双击自签名的 ExcelSqlConsole.exe 会反复弹
    「Windows 无法验证发布者」提示。本脚本（需管理员）：
      1. 关闭 SmartScreen「检查应用和文件」（注册表 SmartScreenEnabled=Off + EnableSmartScreen=0）；
      2. 把 dist\ExcelSqlConsole 加入 Defender 排除项（额外保险，服务若未运行会自动跳过）；
      3. 重启资源管理器使设置立即生效。
    仅适合个人自用；若以后要分发给别人，需购买商业代码签名证书。
.EXAMPLE
    右键本文件 -> 使用 PowerShell 运行（或右键管理员运行），UAC 弹窗点「是」。
#>
$ErrorActionPreference = "Continue"

# 1) 自提权
$isAdmin = ([Security.Principal.WindowsPrincipal][Security.Principal.WindowsIdentity]::GetCurrent()).IsInRole([Security.Principal.WindowsBuiltInRole]::Administrator)
if (-not $isAdmin) {
    try {
        Start-Process powershell -Verb RunAs -ArgumentList "-ExecutionPolicy Bypass -File `"$PSCommandPath`""
        Write-Host "已请求管理员权限，请在弹出的 UAC 窗口点「是」。"
    } catch {
        Write-Host "无法提权：$($_.Exception.Message)"
        Read-Host "按回车退出"
    }
    exit
}

Write-Host "=== 关闭 SmartScreen 应用检查 ==="

# 2) 关闭 SmartScreen「检查应用和文件」
reg add "HKLM\SOFTWARE\Microsoft\Windows\CurrentVersion\Explorer" /v SmartScreenEnabled /t REG_SZ /d Off /f | Out-Null
reg add "HKLM\SOFTWARE\Policies\Microsoft\Windows\System" /v EnableSmartScreen /t REG_DWORD /d 0 /f | Out-Null
Write-Host "  已写入 SmartScreenEnabled=Off / EnableSmartScreen=0"

# 3) 放行本程序目录（Defender 排除，容错）
$exeDir = Join-Path (Split-Path $PSScriptRoot) "dist\ExcelSqlConsole"
try {
    Add-MpPreference -ExclusionPath $exeDir -ErrorAction Stop
    Write-Host "  已加入 Defender 排除：$exeDir"
} catch {
    Write-Host "  Defender 排除添加失败（服务可能未运行，可忽略）：$($_.Exception.Message.Split("`n")[0])"
}

# 4) 重启资源管理器立即生效
Write-Host "正在重启资源管理器（桌面会短暂刷新）..."
Stop-Process -Name explorer -Force -ErrorAction SilentlyContinue

Write-Host ""
Write-Host "完成。之后双击 ExcelSqlConsole.exe 不再弹 SmartScreen 提示。"
Read-Host "按回车退出"