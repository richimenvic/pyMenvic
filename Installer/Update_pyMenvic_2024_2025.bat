@echo off
title pyMenvic Updater for Revit 2024-2025
setlocal EnableExtensions

set "EXT_PATH=%AppData%\pyRevit\Extensions\pyMenvic.extension"

echo ==========================================
echo Updating pyMenvic for Revit 2024-2025...
echo ==========================================
echo.

where git >nul 2>nul
if errorlevel 1 (
    echo ERROR: Git no esta instalado o no esta en PATH.
    pause
    exit /b 1
)

if not exist "%EXT_PATH%" (
    echo ERROR: No existe la carpeta:
    echo %EXT_PATH%
    echo.
    echo Ejecuta primero el instalador.
    pause
    exit /b 1
)

git -C "%EXT_PATH%" pull
if errorlevel 1 (
    echo ERROR: No se pudo actualizar pyMenvic.
    echo Verifica que la extension haya sido instalada desde GitHub.
    pause
    exit /b 1
)

echo.
echo Actualizacion completada.
echo Cierra y vuelve a abrir Revit 2024 o 2025.
pause
