$ErrorActionPreference = 'Stop'
Set-Location $PSScriptRoot

$root = Split-Path $PSScriptRoot -Parent
$appExe = Join-Path $root '小说校对助手.exe'
$readme = Join-Path $root '使用说明.txt'
$out = Join-Path $root '小说校对助手_安装程序_v1.1.2.exe'
$payload = Join-Path $env:TEMP 'novelproofreader_setup_payload'

if (-not (Test-Path $appExe)) { throw "找不到主程序：$appExe" }
if (-not (Test-Path $readme)) { throw "找不到使用说明：$readme" }

if (Test-Path $payload) { Remove-Item $payload -Recurse -Force }
New-Item -ItemType Directory -Path $payload | Out-Null

Copy-Item $appExe (Join-Path $payload '小说校对助手.exe') -Force
Copy-Item $readme (Join-Path $payload '使用说明.txt') -Force

$configTemplate = @"
# 小说校对助手 DeepSeek 在线校对配置
# 填好 ApiKey 后，把 Enabled 改成 true，重新打开程序即可默认勾选在线校对。
# 也可以在程序里手动勾选【DeepSeek 在线校对】。
# 注意：ApiKey 以明文保存在本文件中，请勿把本文件发给别人。

[DeepSeek]
Enabled=false
ApiKey=
ApiBase=https://api.deepseek.com
Model=deepseek-chat
TimeoutSeconds=120
ChunkSize=6000
"@
$cleanConfig = Join-Path $payload '小说校对助手.ini'
[System.IO.File]::WriteAllText($cleanConfig, $configTemplate, (New-Object System.Text.UTF8Encoding($true)))

$csc = Join-Path $env:WINDIR 'Microsoft.NET\Framework64\v4.0.30319\csc.exe'
if (-not (Test-Path $csc)) {
    $csc = Join-Path $env:WINDIR 'Microsoft.NET\Framework\v4.0.30319\csc.exe'
}
if (-not (Test-Path $csc)) { throw '未找到 .NET Framework 4.x 的 csc.exe。' }

$arguments = @(
    '/nologo',
    '/codepage:65001',
    '/target:winexe',
    '/platform:anycpu',
    '/optimize+',
    "/out:$out",
    "/win32icon:$PSScriptRoot\icon.ico",
    "/win32manifest:$PSScriptRoot\app.manifest",
    '/r:System.dll',
    '/r:System.Core.dll',
    '/r:System.Drawing.dll',
    '/r:System.Windows.Forms.dll',
    '/r:Microsoft.CSharp.dll',
    "/resource:$(Join-Path $payload '小说校对助手.exe'),Payload.App",
    "/resource:$(Join-Path $payload '使用说明.txt'),Payload.Readme",
    "/resource:$cleanConfig,Payload.Config",
    'Installer.cs',
    'InstallerAssemblyInfo.cs'
)

& $csc $arguments
if ($LASTEXITCODE -ne 0) { throw "安装程序编译失败，退出码 $LASTEXITCODE" }

# Safety check: the real API key in the workspace config must never enter the installer.
$key = ''
$appIni = Join-Path $root '小说校对助手.ini'
if (Test-Path $appIni) {
    foreach ($line in [System.IO.File]::ReadAllLines($appIni, [System.Text.Encoding]::UTF8)) {
        if ($line -match '^[ \t]*ApiKey[ \t]*=[ \t]*(.+?)[ \t]*$') {
            $key = $Matches[1].Trim()
            break
        }
    }
}

$cleanText = [System.IO.File]::ReadAllText($cleanConfig, [System.Text.Encoding]::UTF8)
if ($cleanText -match '(?m)^[ \t]*ApiKey[ \t]*=[ \t]*\S+') {
    Remove-Item $out -Force -ErrorAction SilentlyContinue
    throw '安装包内的模板配置包含 ApiKey，已中止。'
}

if ($key.Length -gt 0) {
    $filesToCheck = @($out, $appExe, $cleanConfig)
    foreach ($file in $filesToCheck) {
        $bytes = [System.IO.File]::ReadAllBytes($file)
        $utf8 = [System.Text.Encoding]::UTF8.GetString($bytes)
        $unicode = [System.Text.Encoding]::Unicode.GetString($bytes)
        if ($utf8.Contains($key) -or $unicode.Contains($key)) {
            Remove-Item $out -Force -ErrorAction SilentlyContinue
            throw "检测到安装包或主程序中包含工作区配置里的 API Key，已中止：$file"
        }
    }
    Write-Host '安全校验：安装包中未包含工作区配置的 API Key。'
} else {
    Write-Host '安全校验：工作区配置中没有填写 API Key。'
}

Remove-Item $payload -Recurse -Force -ErrorAction SilentlyContinue
Write-Host "安装程序已生成：$out"
