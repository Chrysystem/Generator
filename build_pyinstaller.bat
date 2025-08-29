@echo off
setlocal enabledelayedexpansion

echo == Build Generator (PyInstaller) ==

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

REM 5) Clean previous artifacts
if exist build rmdir /s /q build
if exist dist rmdir /s /q dist

REM 6) Build using spec if available, else fallback CLI
if exist excel_to_word_app.spec (
  echo Building with spec...
  pyinstaller --clean --noconfirm excel_to_word_app.spec
) else (
  echo Spec not found. Building with CLI options...
  pyinstaller --noconfirm --clean --onefile --windowed ^
    --name Generator_Formation ^
    --icon logo-Toyota-Solo.ico ^
    --add-data "Datas;Datas" ^
    --add-data "LogoTMH.png;." ^
    --add-data "logo-Toyota-Solo.ico;." ^
    --hidden-import win32com ^
    --hidden-import win32com.client ^
    --hidden-import pythoncom ^
    excel_to_word_app.py
)

if errorlevel 1 (
  echo Build failed.
  exit /b 1
)

echo Build succeeded!
echo Output folder: dist
for /r dist %%i in (*) do echo %%i

endlocal

