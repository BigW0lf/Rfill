@echo off
setlocal
set ROOT=%~dp0
set PYTHON=%ROOT%build_env\Scripts\python.exe

echo ============================================================
echo  Rfill — Build complet (EXE + Installer)
echo ============================================================

echo [1/3] Generation de l'icone...
"%PYTHON%" "%ROOT%build_tools\make_ico.py"
if errorlevel 1 ( echo ERREUR make_ico.py & pause & exit /b 1 )

echo [2/3] Compilation PyInstaller...
"%PYTHON%" -m PyInstaller "%ROOT%build_tools\rfill.spec" --clean --noconfirm
if errorlevel 1 ( echo ERREUR PyInstaller & pause & exit /b 1 )

echo [3/3] Creation de l'installer Inno Setup...
set ISCC=%LOCALAPPDATA%\Programs\Inno Setup 6\ISCC.exe
if not exist "%ISCC%" set ISCC=C:\Program Files (x86)\Inno Setup 6\ISCC.exe
if not exist "%ISCC%" set ISCC=C:\Program Files\Inno Setup 6\ISCC.exe
if not exist "%ISCC%" ( echo ERREUR : ISCC.exe introuvable & pause & exit /b 1 )
"%ISCC%" "%ROOT%build_tools\rfill_installer.iss"
if errorlevel 1 ( echo ERREUR Inno Setup & pause & exit /b 1 )

echo.
echo ============================================================
echo  Build termine !
echo  - Executable : dist\Rfill.exe
echo  - Installer  : installer\Rfill_Setup_1.0.0.exe
echo ============================================================
pause
