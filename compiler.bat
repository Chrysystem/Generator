@echo off
echo ========================================
echo    Compilation Generator Formation
echo ========================================
echo.

echo Installation des dependances...
python -m pip install -r requirements.txt

echo.
echo Lancement de la compilation...
python build_exe.py

echo.
echo Appuyez sur une touche pour fermer...
pause > nul
