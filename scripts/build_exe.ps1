$ErrorActionPreference = "Stop"

$ProjectRoot = Split-Path -Parent $PSScriptRoot
Set-Location $ProjectRoot

$python = "python"
if ($env:PYTHON) {
    $python = $env:PYTHON
}

$appName = [Text.Encoding]::UTF8.GetString([Convert]::FromBase64String("5rW35Y2X5ZSu55S157uT566X6Ieq5Yqo5YyW5bel5YW3"))

& $python -m pip install -r requirements.txt
& $python -m PyInstaller `
    --noconfirm `
    --clean `
    --onefile `
    --windowed `
    --name $appName `
    --paths "$ProjectRoot\src" `
    "$ProjectRoot\scripts\run_gui.py"

Write-Host "Build complete: $ProjectRoot\dist\$appName.exe"
