<#
.SYNOPSIS
    给打包产物 ExcelSqlConsole.exe 打上本机代码签名，并把证书加入当前用户的
    "受信任的发布者" 与 "受信任的根证书颁发机构" 存储。
.DESCRIPTION
    PyInstaller 的产物无 Authenticode 签名，Windows SmartScreen 每次启动都会
    弹 "Microsoft Defender SmartScreen 阻止了无法识别的应用启动"。
    本脚本用自签名代码签名证书（首次自动生成，5 年有效，之后复用）给 exe 签名
    （带 RFC3161 时间戳），并尝试信任证书 —— 之后本机双击不再被 SmartScreen 拦截，
    发布者显示 "ExcelSql Console"。

    注意：把签名证书加入 "受信任的根证书颁发机构" 时，Windows 会弹安全确认框，
    需要你点一次 "是"（这是系统对根证书的强制保护，脚本无法替代）。若脚本运行后
    最终状态仍为 UnknownError，请双击脚本导出的 .cer 手动安装到"受信任的根"。

.PARAMETER ExePath
    要签名的 exe 路径。默认 scripts\..\dist\ExcelSqlConsole\ExcelSqlConsole.exe。
.EXAMPLE
    powershell -ExecutionPolicy Bypass -File scripts\sign.ps1
#>
[CmdletBinding()]
param(
    [string]$ExePath = ""
)

$ErrorActionPreference = "Stop"

$scriptDir = Split-Path -Parent $MyInvocation.MyCommand.Path
if (-not $ExePath) {
    $ExePath = Join-Path $scriptDir "..\dist\ExcelSqlConsole\ExcelSqlConsole.exe"
}

function Add-CertToStore([string]$StoreName, [byte[]]$Der) {
    Add-Type -TypeDefinition @'
using System;
using System.Runtime.InteropServices;
public static class CertStoreSigner2 {
    [DllImport("crypt32.dll", CharSet=CharSet.Unicode, SetLastError=true)]
    public static extern IntPtr CertOpenSystemStoreW(IntPtr hProv, string szSubsystemProtocol);
    [DllImport("crypt32.dll", SetLastError=true)]
    public static extern bool CertAddEncodedCertificateToStore(IntPtr hStore, int encEncodingType, byte[] pbEncoded, int cbEncoded, int dwAddDisposition, out IntPtr ppvContext);
    [DllImport("crypt32.dll", SetLastError=true)]
    public static extern bool CertCloseStore(IntPtr hStore, uint dwFlags);
}
'@
    $h = [CertStoreSigner2]::CertOpenSystemStoreW([IntPtr]::Zero, $StoreName)
    if ($h -eq [IntPtr]::Zero) { throw "打开证书存储 $StoreName 失败: $([Runtime.InteropServices.Marshal]::GetLastWin32Error())" }
    try {
        $ctx = [IntPtr]::Zero
        # 底层 API 写根存储（避免 .NET X509Store.Add 的 CryptUI 向导弹窗）
        $ok = [CertStoreSigner2]::CertAddEncodedCertificateToStore($h, 1, $Der, $Der.Length, 3, [ref]$ctx)
        if (-not $ok) { throw "添加证书到 $StoreName 失败: $([Runtime.InteropServices.Marshal]::GetLastWin32Error())" }
    } finally {
        [CertStoreSigner2]::CertCloseStore($h, 0) | Out-Null
    }
}

$exe = (Resolve-Path $ExePath).Path
Write-Host "== 代码签名: $exe =="

# 1) 复用或生成自签名代码签名证书（当前用户存储，免管理员）
$cert = Get-ChildItem Cert:\CurrentUser\My -CodeSigningCert -ErrorAction SilentlyContinue |
    Where-Object { $_.Subject -match "ExcelSql Console" } | Select-Object -First 1
if (-not $cert) {
    Write-Host "生成自签名代码签名证书（5 年有效）..."
    $cert = New-SelfSignedCertificate -Type CodeSigningCert `
        -Subject "CN=ExcelSql Console, O=Local Dev, C=CN" `
        -FriendlyName "ExcelSql Console Code Signing (local)" `
        -CertStoreLocation "Cert:\CurrentUser\My" `
        -NotAfter (Get-Date).AddYears(5) `
        -KeyAlgorithm RSA -KeyLength 2048
}
Write-Host "证书: $($cert.Thumbprint) 有效期至 $($cert.NotAfter)"

# 2) 签名（优先带时间戳；时间戳服务不可达时降级为无时间戳签名）
$sig = $null
foreach ($ts in @("http://timestamp.digicert.com", "http://timestamp.sectigo.com")) {
    try {
        $sig = Set-AuthenticodeSignature -FilePath $exe -Certificate $cert -HashAlgorithm SHA256 -TimestampServer $ts -ErrorAction Stop
        Write-Host "已用时间戳 $ts 写入签名（Set 返回 $($sig.Status)）"
        if ($sig.Status -eq "Valid") { break }
    } catch { Write-Host "时间戳服务不可达: $ts" }
}
if (-not $sig -or $sig.Status -ne "Valid") {
    $sig = Set-AuthenticodeSignature -FilePath $exe -Certificate $cert -HashAlgorithm SHA256
    Write-Host "无时间戳签名（Set 返回 $($sig.Status)）"
}

# 3) 信任证书：加入 受信任的根 + 受信任的发布者（当前用户）。
#    已存在则跳过（幂等，避免对根存储的重复 REPLACE 触发确认）；每次重建后签名在，证书不变。
$cer = Join-Path $scriptDir "..\ExcelSqlConsole_trust.cer"
Export-Certificate -Cert $cert -FilePath $cer -Type CERT -Force | Out-Null
$der = [System.IO.File]::ReadAllBytes($cer)

foreach ($storeName in @("Root", "TrustedPublisher")) {
    $exists = Get-ChildItem "Cert:\CurrentUser\$storeName" -ErrorAction SilentlyContinue |
        Where-Object { $_.Thumbprint -eq $cert.Thumbprint }
    if ($exists) {
        Write-Host "证书已在 $storeName，跳过"
        continue
    }
    try {
        Add-CertToStore $storeName $der
        Write-Host "证书已加入 $storeName"
    } catch {
        Write-Warning "加入 $storeName 失败（$($_.Exception.Message)）。可能因根证书安全确认弹出未点击；可手动双击 $cer 安装到""受信任的根证书颁发机构""。"
    }
}

# 4) 最终校验
$v = Get-AuthenticodeSignature $exe
Write-Host "== 最终状态: $($v.Status)  发布者: $($v.SignerCertificate.Subject)  带时间戳: $($null -ne $v.TimeStamperCertificate) =="
if ($v.Status -ne "Valid") {
    Write-Warning "签名暂未生效（根证书尚未受信任）。请双击 $cer → 安装证书 → 选择""受信任的根证书颁发机构"" → 确定，即可恢复。"
} else {
    Write-Host "签名有效，SmartScreen 不再拦截（本机）。"
}