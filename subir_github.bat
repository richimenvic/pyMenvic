@echo off
cd /d "%~dp0"

git add .
git commit -m "update"
git pull --rebase origin main
git push origin main

pause