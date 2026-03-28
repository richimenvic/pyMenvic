@echo off
cd /d "%~dp0"

set /p MSG=Escribe el mensaje del commit: 

git add .
git commit -m "%MSG%"
git pull --rebase origin main
git push origin main

pause