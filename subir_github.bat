REM Hook automatico activo: valida titulos y bundles antes de cada commit
@echo off
cd /d "%~dp0"

git diff --quiet
if %errorlevel%==0 (
    echo No hay cambios para subir.
    pause
    exit /b
)

set /p MSG=Escribe el mensaje del commit: 

git add .
git commit -m "%MSG%"
if errorlevel 1 (
    echo Commit cancelado o con error.
    pause
    exit /b
)

git pull --rebase origin main
if errorlevel 1 (
    echo Error en pull --rebase.
    pause
    exit /b
)

git push origin main

pause