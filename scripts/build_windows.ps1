param(
    [switch]$OneFile = $true
)

$ErrorActionPreference = "Stop"

$projectRoot = Split-Path -Parent $PSScriptRoot
Set-Location $projectRoot

python -m pip install --upgrade pip
python -m pip install -e . pyinstaller

$pyInstallerArgs = @(
    "--noconfirm",
    "--clean",
    "--name", "compressPPTX",
    "--specpath", "build",
    "--paths", "src",
    "compress_pptx.py"
)

if ($OneFile) {
    $pyInstallerArgs += "--onefile"
}

python -m PyInstaller @pyInstallerArgs

Write-Host ""
Write-Host "Build complete."
Write-Host "Executable: dist\\compressPPTX.exe"
Write-Host "If audio/video compression is needed, provide ffmpeg.exe next to the executable or in PATH."
