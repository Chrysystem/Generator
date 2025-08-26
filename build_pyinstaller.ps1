# Requires: PowerShell on Windows
# Usage: Right-click -> Run with PowerShell (or run from a PowerShell prompt)

$ErrorActionPreference = "Stop"

Write-Host "== Build Generator (PyInstaller) ==" -ForegroundColor Cyan

# 1) Create and activate venv
$venvPath = "outlook_env"
if (-Not (Test-Path $venvPath)) {
  Write-Host "Creating virtual environment..." -ForegroundColor Yellow
  python -m venv $venvPath
}

$venvActivate = Join-Path $venvPath "Scripts\Activate.ps1"
. $venvActivate

# 2) Upgrade pip and install deps
Write-Host "Installing dependencies..." -ForegroundColor Yellow
python -m pip install --upgrade pip
if (Test-Path requirements.txt) {
  pip install -r requirements.txt
}
# Ensure pyinstaller present
pip install pyinstaller

# 3) Validate required resources
$required = @(
  "excel_to_word_app.py",
  "logo-Toyota-Solo.ico",
  "LogoTMH.png",
  "Datas"
)

$missing = @()
foreach ($item in $required) {
  if (-Not (Test-Path $item)) { $missing += $item }
}
if ($missing.Count -gt 0) {
  Write-Host "Missing required files/folders: $($missing -join ', ')" -ForegroundColor Red
  exit 1
}

# 4) Clean previous build artifacts
Write-Host "Cleaning previous build artifacts..." -ForegroundColor Yellow
if (Test-Path "build") { try { Remove-Item -Recurse -Force "build" -ErrorAction Stop } catch {} }
if (Test-Path "dist") { try { Remove-Item -Recurse -Force "dist" -ErrorAction Stop } catch {} }

# 5) Build using .spec if present, otherwise fallback CLI
$specFile = "excel_to_word_app.spec"
if (Test-Path $specFile) {
  Write-Host "Building with spec: $specFile" -ForegroundColor Yellow
  pyinstaller --clean --noconfirm $specFile
} else {
  Write-Host "Spec not found. Building with CLI options..." -ForegroundColor Yellow
  pyinstaller --noconfirm --clean --onefile --windowed `
    --name Generator_Formation `
    --icon logo-Toyota-Solo.ico `
    --add-data "Datas;Datas" `
    --add-data "LogoTMH.png;." `
    --add-data "logo-Toyota-Solo.ico;." `
    --hidden-import win32com `
    --hidden-import win32com.client `
    --hidden-import pythoncom `
    excel_to_word_app.py
}

if ($LASTEXITCODE -ne 0) {
  Write-Host "Build failed." -ForegroundColor Red
  exit 1
}

Write-Host "Build succeeded!" -ForegroundColor Green
Write-Host "Output folder: dist" -ForegroundColor Green
Get-ChildItem dist -Recurse | ForEach-Object { $_.FullName }
