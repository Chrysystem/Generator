@echo off
setlocal enabledelayedexpansion

echo == Build Generator (Windows) ==

REM 1) Create venv if missing
if not exist .venv (
  echo Creating virtual environment...
  python -m venv .venv
)

REM 2) Activate venv
call .venv\Scripts\activate.bat

REM 3) Install deps
python -m pip install --upgrade pip
if exist requirements.txt (
  pip install -r requirements.txt
)
pip install pyinstaller

REM 4) Validate required resources
set MISSING=
for %%F in (excel_to_word_app.py logo-Toyota-Solo.ico LogoTMH.png) do (
  if not exist "%%F" (
    set MISSING=!MISSING! %%F
  )
)
if not exist Datas (
  set MISSING=!MISSING! Datas
)
if not "!MISSING!"=="" (
  echo Missing required files/folders: !MISSING!
  exit /b 1
)

REM 5) Build
pyinstaller --clean --noconfirm excel_to_word_app.spec
if errorlevel 1 (
  echo Build failed.
  exit /b 1
)

echo Build succeeded!
echo Output folder: dist\Generator_Formation
echo Executable: dist\Generator_Formation\Generator_Formation.exe
endlocal
