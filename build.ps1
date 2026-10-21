$ErrorActionPreference = 'Stop'
Set-Location $PSScriptRoot

$csc = Join-Path $env:WINDIR 'Microsoft.NET\Framework64\v4.0.30319\csc.exe'
if (-not (Test-Path $csc)) {
    $csc = Join-Path $env:WINDIR 'Microsoft.NET\Framework\v4.0.30319\csc.exe'
}
if (-not (Test-Path $csc)) {
    throw '未找到 .NET Framework 4.x 的 csc.exe。'
}

$arguments = @(
    '/nologo',
    '/codepage:65001',
    '/target:winexe',
    '/platform:anycpu',
    '/optimize+',
    '/out:小说校对助手.exe',
    '/win32icon:icon.ico',
    '/win32manifest:app.manifest',
    '/r:System.dll',
    '/r:System.Core.dll',
    '/r:System.Xml.dll',
    '/r:System.Xml.Linq.dll',
    '/r:System.Drawing.dll',
    '/r:System.Windows.Forms.dll',
    '/r:System.IO.Compression.dll',
    '/r:System.IO.Compression.FileSystem.dll',
    '/r:Microsoft.CSharp.dll',
    '/r:System.Web.Extensions.dll',
    'Program.cs',
    'MainForm.cs',
    'Scanner.cs',
    'DocxTool.cs',
    'WordCom.cs',
    'WordApplyHelper.cs',
    'DeepSeekConfig.cs',
    'DeepSeekProofreader.cs',
    'AssemblyInfo.cs'
)

& $csc $arguments
if ($LASTEXITCODE -ne 0) {
    throw "编译失败，退出码 $LASTEXITCODE"
}

Copy-Item -Path '小说校对助手.exe' -Destination (Join-Path (Split-Path $PSScriptRoot -Parent) '小说校对助手.exe') -Force
Write-Host '编译完成：小说校对助手.exe'
