# SPDX-License-Identifier: GPL-3.0-or-later
# 可选：为发布 exe 添加自签名（完整性校验用，不能消除 SmartScreen 提示）
# 用法: powershell -ExecutionPolicy Bypass -File scripts/self-sign.ps1 [-Path dist\Md2Docx_Release]
#        [-TimestampServer <url>]（默认 digicert，离线环境可换内网 RFC3161 服务器）
param(
    [string]$Path = "dist\Md2Docx_Release",
    [string]$Subject = "CN=Md2Docx Open Source",
    [string]$TimestampServer = "http://timestamp.digicert.com"
)

$ErrorActionPreference = "Stop"

# 定位 signtool：优先 PATH，其次在 Windows SDK 目录里按版本号数值降序取最新
# （目录名字符串排序在版本位数变化时会选错，如 10.0.999 < 10.0.1000）
$signTool = Get-Command signtool.exe -ErrorAction SilentlyContinue
if (-not $signTool) {
    $sdkRoots = @(
        "${env:ProgramFiles(x86)}\Windows Kits\10\bin",
        "$env:ProgramFiles\Windows Kits\10\bin"
    )
    foreach ($root in $sdkRoots) {
        if (Test-Path $root) {
            $signTool = Get-ChildItem $root -Recurse -Filter signtool.exe | Sort-Object {
                # 取路径中的四段版本号转数值，倒序即"最新在前"
                if ($_.FullName -match '(\d+)\.(\d+)\.(\d+)\.(\d+)') {
                    [int64]$Matches[1] * 1e9 + [int64]$Matches[2] * 1e6 + [int64]$Matches[3] * 1e3 + [int64]$Matches[4]
                } else { 0 }
            } -Descending | Select-Object -First 1
            if ($signTool) { break }
        }
    }
}
if (-not $signTool) {
    Write-Error "未找到 signtool.exe，请安装 Windows SDK（https://developer.microsoft.com/windows/downloads/windows-sdk/）"
    exit 1
}
$signToolPath = if ($signTool.FullName) { $signTool.FullName } else { "$signTool" }
Write-Host "signtool: $signToolPath"

# 创建（或复用）代码签名证书
$cert = Get-ChildItem Cert:\CurrentUser\My -CodeSigningCert |
    Where-Object Subject -eq $Subject | Select-Object -First 1
if (-not $cert) {
    $cert = New-SelfSignedCertificate -Type CodeSigningCert -Subject $Subject `
        -CertStoreLocation Cert:\CurrentUser\My -NotAfter (Get-Date).AddYears(3)
    Write-Host "已创建自签名证书: $($cert.Thumbprint)"
}

# 导出证书供用户按需导入"受信任的发布者"。
# 注意：-Path 指向发布目录时证书会随目录一起打进 zip（打包前运行即可携带）
$cerPath = Join-Path $Path "md2docx-publisher.cer"
Export-Certificate -Cert $cert -FilePath $cerPath | Out-Null
Write-Host "证书已导出: $cerPath（用户可导入以信任发布者）"

# 签名全部 exe（带时间戳；离线环境可经 -TimestampServer 关闭或指向内网服务器）
Get-ChildItem $Path -Filter *.exe | ForEach-Object {
    Write-Host "签名: $($_.Name)"
    & $signToolPath sign /fd SHA256 /tr $TimestampServer /td SHA256 `
        /sha1 $cert.Thumbprint $_.FullName
    if ($LASTEXITCODE -ne 0) { Write-Error "签名失败: $($_.Name)" }
}
Write-Host "完成。注意：自签名不会消除 SmartScreen 提示（见 docs/SIGNING.md）"
