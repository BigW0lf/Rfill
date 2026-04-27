@echo off
echo ============================================================
echo  Rfill — Build EXE
echo ============================================================

echo [1/2] Generation de l'icone...
python make_ico.py
if errorlevel 1 (
    echo ERREUR : make_ico.py a echoue.
    pause
    exit /b 1
)

echo [2/2] Compilation avec PyInstaller...
pyinstaller rfill.spec --clean --noconfirm
if errorlevel 1 (
    echo ERREUR : PyInstaller a echoue.
    pause
    exit /b 1
)

echo.
echo ============================================================
echo  Build termine : dist\Rfill\Rfill.exe
echo ============================================================
pause
