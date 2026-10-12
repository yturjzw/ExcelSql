<#
.SYNOPSIS
    给打包产物 ExcelSqlConsole.exe 打上本机代码签名，并把证书加入当前用户的
    "受信任的发布者" 与 "受信任的根证书颁发机构" 存储。
.DESCRIPTION
    PyInstaller 的产物无 Authenticode 签名，Windows SmartScreen 每次启动都会
    弹 "Microsoft Defender SmartScreen 阻止了无法识别的应用启动"。
    本脚本用自签名代码签名证书（首次自动生成，5 年有效，之后复用）给 exe 签名
    （带 RFC3161 时间戳），并信任证书 —— 之后本机双击不再被 SmartScreen 拦截，
    发布者显示 "ExcelSql Console"。

    注意：信任与签名都是"本机"行为；把产物拷到别的电脑仍会提示未知发布者，
    需在那台机器上执行本脚本（或安装证书 / 买商业签名证书）。

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

# 注意：param 默认值表达式里 $PSScriptRoot 尚未初始化（PowerShell 5.1 怪癖），
# 故在此处计算脚本目录。
$scriptDir = Split-Path -Parent $MyInvocation.MyCommand.Path
if (-not $ExePath) {
    $ExePath = Join-Path $scriptDir "..\dist\ExcelSqlConsole\ExcelSqlConsole.exe"
}

function Add-CertToStore([string]$StoreName, [byte[]]$Der) {
    Add-Type -TypeDefinition @'
using System;
using System.Runtime.InteropServices;
public static class CertStoreSigner {
    [DllImport("crypt32.dll", CharSet=CharSet.Unicode, SetLastError=true)]
    public static extern IntPtr CertOpenSystemStoreW(IntPtr hProv, string szSubsystemProtocol);
    [DllImport("crypt32.dll", SetLastError=true)]
    public static extern bool CertAddEncodedCertificateToStore(IntPtr hStore, int encEncodingType, byte[] pbEncoded, int cbEncoded, int dwAddDisposition, out IntPtr ppvContext);
    [DllImport("crypt32.dll", SetLastError=true)]
    public static extern bool CertCloseStore(IntPtr hStore, uint dwFlags);
}
'@
    $h = [CertStoreSigner]::CertOpenSystemStoreW([IntPtr]::Zero, $StoreName)
    if ($h -eq [IntPtr]::Zero) { throw "打开证书存储 $StoreName 失败: $([Runtime.InteropServices.Marshal]::GetLastWin32Error())" }
    try {
        $ctx = [IntPtr]::Zero
        # 用 Crypt32 底层 API 写入，绕开 .NET X509Store.Add 对根存储的 UI 确认向导
        $ok = [CertStoreSigner]::CertAddEncodedCertificateToStore($h, 1, $Der, $Der.Length, 3, [ref]$ctx)
        if (-not $ok) { throw "添加证书到 $StoreName 失败: $([Runtime.InteropServices.Marshal]::GetLastWin32Error())" }
    } finally {
        [CertStoreSigner]::CertCloseStore($h, 0) | Out-Null
    }
    Write-Host "  证书已加入 $StoreName"
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
        if ($sig.Status -eq "Valid") { Write-Host "签名成功（时间戳: $ts）"; break }
    } catch { Write-Host "时间戳服务不可达: $ts" }
}
if (-not $sig -or $sig.Status -ne "Valid") {
    $sig = Set-AuthenticodeSignature -FilePath $exe -Certificate $cert -HashAlgorithm SHA256
    if ($sig.Status -eq "Valid") { Write-Host "签名成功（无时间戳）" }
}

# 3) 信任证书：加入 受信任的根 + 受信任的发布者（当前用户）
if ($sig.Status -eq "Valid") {
    $cer = Join-Path $env:TEMP "excel_sql_cert.cer"
    Export-Certificate -Cert $cert -FilePath $cer -Type CERT -Force | Out-Null
    $der = [System.IO.File]::ReadAllBytes($cer)
    Add-CertToStore "Root" $der
    Add-CertToStore "TrustedPublisher" $der
    Remove-Item $cer -Force -ErrorAction SilentlyContinue
}

# 4) 最终校验
$v = Get-AuthenticodeSignature $exe
Write-Host "== 最终状态: $($v.Status)  发布者: $($v.SignerCertificate.Subject)  带时间戳: $($null -ne $v.TimeStamperCertificate) =="
if ($v.Status -ne "Valid") { throw "签名未生效: $($v.StatusMessage)" }