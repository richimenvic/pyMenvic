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
git pull --rebase origin main
git push origin main

pause