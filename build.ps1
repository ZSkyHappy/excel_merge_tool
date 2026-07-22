param(
    [string]$Python = "python",
    [switch]$Force
)

$ErrorActionPreference = "Stop"
$ProjectDir = Split-Path -Parent $MyInvocation.MyCommand.Path
$VenvPython = Join-Path $ProjectDir ".venv\Scripts\python.exe"
$ExePath = Join-Path $ProjectDir "dist\Excel批量合并工具.exe"

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
}

& $VenvPython -m pip install --disable-pip-version-check -r requirements-build.txt
& $VenvPython -m unittest discover -s tests -v
& $VenvPython -m PyInstaller `
    --noconfirm `
    --clean `
    --onefile `
    --windowed `
    --name "Excel批量合并工具" `
    excel_merge_tool.pyw

Write-Host "构建完成：$ExePath"

