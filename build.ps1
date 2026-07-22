param(
    [string]$Python = "python",
    [switch]$Force
)

$ErrorActionPreference = "Stop"
$ProjectDir = Split-Path -Parent $MyInvocation.MyCommand.Path
$VenvPython = Join-Path $ProjectDir ".venv\Scripts\python.exe"
$BuildDir = Join-Path $ProjectDir "build"
$IconPath = Join-Path $BuildDir "app.ico"
$VersionPath = Join-Path $ProjectDir "version_info.txt"
$ExeName = "ExcelMergeTool-Windows-x64.exe"
$ExePath = Join-Path $ProjectDir "dist\$ExeName"
$ChecksumPath = Join-Path $ProjectDir "dist\SHA256SUMS.txt"

Set-Location $ProjectDir

if ((Test-Path -LiteralPath $ExePath) -and -not $Force) {
    $Answer = Read-Host "构建会覆盖 $ExePath。输入 Y 继续"
    if ($Answer -notin @("Y", "y")) {
        Write-Host "已取消构建。"
        exit 1
    }
}

if (-not (Test-Path -LiteralPath $VenvPython)) {
    & $Python -m venv (Join-Path $ProjectDir ".venv")
    if ($LASTEXITCODE -ne 0) {
        throw "创建项目虚拟环境失败。"
    }
}

& $VenvPython -m pip install --disable-pip-version-check -r requirements-build.txt
if ($LASTEXITCODE -ne 0) {
    throw "安装构建依赖失败。"
}

& $VenvPython -m unittest discover -s tests -v
if ($LASTEXITCODE -ne 0) {
    throw "自动化测试未通过，已停止构建。"
}

New-Item -ItemType Directory -Path $BuildDir -Force | Out-Null
Add-Type -AssemblyName System.Drawing
$Bitmap = [System.Drawing.Bitmap]::new(256, 256)
$Graphics = [System.Drawing.Graphics]::FromImage($Bitmap)
$Graphics.SmoothingMode = [System.Drawing.Drawing2D.SmoothingMode]::AntiAlias
$Background = [System.Drawing.SolidBrush]::new([System.Drawing.Color]::FromArgb(31, 78, 120))
$WhitePen = [System.Drawing.Pen]::new([System.Drawing.Color]::White, 14)
$ArrowPen = [System.Drawing.Pen]::new([System.Drawing.Color]::FromArgb(169, 208, 142), 18)
$ArrowPen.EndCap = [System.Drawing.Drawing2D.LineCap]::ArrowAnchor
try {
    $Graphics.FillRectangle($Background, 0, 0, 256, 256)
    $Graphics.DrawRectangle($WhitePen, 42, 42, 76, 172)
    $Graphics.DrawLine($WhitePen, 42, 98, 118, 98)
    $Graphics.DrawLine($WhitePen, 42, 156, 118, 156)
    $Graphics.DrawRectangle($WhitePen, 138, 42, 76, 172)
    $Graphics.DrawLine($WhitePen, 138, 98, 214, 98)
    $Graphics.DrawLine($WhitePen, 138, 156, 214, 156)
    $Graphics.DrawLine($ArrowPen, 91, 128, 180, 128)
    $Handle = $Bitmap.GetHicon()
    $Icon = [System.Drawing.Icon]::FromHandle($Handle)
    $Stream = [System.IO.File]::Create($IconPath)
    try {
        $Icon.Save($Stream)
    }
    finally {
        $Stream.Dispose()
        $Icon.Dispose()
    }
}
finally {
    $ArrowPen.Dispose()
    $WhitePen.Dispose()
    $Background.Dispose()
    $Graphics.Dispose()
    $Bitmap.Dispose()
}

$PyInstallerArgs = @(
    "--noconfirm",
    "--clean",
    "--onefile",
    "--windowed",
    "--name", "ExcelMergeTool-Windows-x64",
    "--icon", $IconPath,
    "--version-file", $VersionPath,
    "--add-data", "$IconPath;."
)
& $VenvPython -m PyInstaller @PyInstallerArgs excel_merge_tool.pyw
if ($LASTEXITCODE -ne 0) {
    throw "PyInstaller 构建失败。"
}

$Hash = (Get-FileHash -LiteralPath $ExePath -Algorithm SHA256).Hash.ToLowerInvariant()
"$Hash  $ExeName" | Set-Content -LiteralPath $ChecksumPath -Encoding ascii

Write-Host "构建完成：$ExePath"
Write-Host "校验文件：$ChecksumPath"

