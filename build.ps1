# 一键编译：生成单文件 exe（自带 .NET Framework 4.8，无需安装任何东西）
# 用法：powershell -ExecutionPolicy Bypass -File build.ps1

$ErrorActionPreference = 'Stop'
$root = Split-Path -Parent $MyInvocation.MyCommand.Path
$csc  = 'C:\Windows\Microsoft.NET\Framework64\v4.0.30319\csc.exe'

if (-not (Test-Path $csc)) {
    Write-Host "找不到 C# 编译器：$csc" -ForegroundColor Red
    Write-Host "请确认系统是 Windows 10/11 且 .NET Framework 4.x 未损坏。" -ForegroundColor Red
    exit 1
}

$dist = Join-Path $root 'dist'
if (-not (Test-Path $dist)) { New-Item -ItemType Directory -Path $dist | Out-Null }
$exe = Join-Path $dist 'ReportTool.exe'

$sources = @(
    (Join-Path $root 'core\Model.cs'),
    (Join-Path $root 'core\XlsxReader.cs'),
    (Join-Path $root 'core\Lexer.cs'),
    (Join-Path $root 'core\Evaluator.cs'),
    (Join-Path $root 'core\Calculator.cs'),
    (Join-Path $root 'core\XlsxWriter.cs'),
    (Join-Path $root 'app\Pipeline.cs'),
    (Join-Path $root 'app\Summary.cs'),
    (Join-Path $root 'app\Chart.cs'),
    (Join-Path $root 'app\Inbox.cs'),
    (Join-Path $root 'app\MainForm.cs'),
    (Join-Path $root 'app\Program.cs')
)

$refs = @(
    '/r:System.dll',
    '/r:System.Core.dll',
    '/r:System.Drawing.dll',
    '/r:System.Windows.Forms.dll',
    '/r:System.IO.Compression.dll',
    '/r:System.IO.Compression.FileSystem.dll',
    '/r:System.Xml.dll'
)

$args = @('/nologo', '/target:winexe', '/optimize+', '/langversion:5', "/out:$exe") + $refs + $sources

Write-Host "正在编译…" -ForegroundColor Cyan
& $csc $args
if ($LASTEXITCODE -ne 0) { Write-Host "编译失败。" -ForegroundColor Red; exit $LASTEXITCODE }

$size = [math]::Round((Get-Item $exe).Length / 1KB, 0)
Write-Host "编译成功：$exe ($size KB)" -ForegroundColor Green
