@echo off
setlocal

set SPHINXOPTS=
set SPHINXBUILD=sphinx-build
set SOURCEDIR=source
set BUILDDIR=build

if "%1" == "" goto help
if "%1" == "help" goto help
if "%1" == "clean" goto clean
if "%1" == "html" goto html
if "%1" == "open" goto open

:help
echo Utilisation : make.bat [html ^| clean ^| open]
echo   html   - Genere la documentation HTML dans build/html/
echo   clean  - Supprime le dossier build/
echo   open   - Ouvre la doc dans le navigateur
goto end

:clean
echo Nettoyage de build/...
rmdir /s /q %BUILDDIR% 2>NUL
echo Fait.
goto end

:html
echo Generation de la documentation HTML...
%SPHINXBUILD% -b html -W --keep-going %SPHINXOPTS% %SOURCEDIR% %BUILDDIR%/html
if errorlevel 1 (
    echo.
    echo ERREUR : sphinx-build a echoue. Voir les messages ci-dessus.
    pause
    exit /b 1
)
echo.
echo Documentation generee dans : %BUILDDIR%\html\index.html
goto end

:open
start "" "%~dp0build\html\index.html"
goto end

:end
endlocal
