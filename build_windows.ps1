# Requires: PowerShell on Windows
# Usage: Right-click -> Run with PowerShell (or run from a PowerShell prompt)

$ErrorActionPreference = "Stop"

Write-Host "== Build Generator (Windows) ==" -ForegroundColor Cyan

# 1) Create and activate venv
if (-Not (Test-Path .venv)) {
  Write-Host "Creating virtual environment..." -ForegroundColor Yellow
  python -m venv .venv
}

$venvActivate = Join-Path ".venv" "Scripts\Activate.ps1"
. $venvActivate

# 2) Upgrade pip and install deps
Write-Host "Installing dependencies..." -ForegroundColor Yellow
python -m pip install --upgrade pip
if (Test-Path requirements.txt) {
  pip install -r requirements.txt
}

# Ensure pyinstaller is present
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

# 4) Run PyInstaller with existing spec
Write-Host "Building executable with PyInstaller..." -ForegroundColor Yellow
pyinstaller --clean --noconfirm excel_to_word_app.spec

if ($LASTEXITCODE -ne 0) {
  Write-Host "Build failed." -ForegroundColor Red
  exit 1
}

Write-Host "Build succeeded!" -ForegroundColor Green
Write-Host "Output folder: dist/Generator_Formation" -ForegroundColor Green
Write-Host "Executable: dist/Generator_Formation/Generator_Formation.exe" -ForegroundColor Green
