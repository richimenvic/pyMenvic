@echo off
title pyMenvic Installer for Revit 2024-2025
setlocal EnableExtensions

set "EXT_DIR=%AppData%\pyRevit\Extensions"
set "EXT_PATH=%EXT_DIR%\pyMenvic.extension"
set "CONFIG_DIR=%AppData%\pyRevit"
set "CONFIG_FILE=%CONFIG_DIR%\pyRevit_config.ini"
set "REPO_URL=https://github.com/richimenvic/pyMenvic.git"

echo ==========================================
echo Installing pyMenvic for Revit 2024-2025...
echo ==========================================
echo.

where pyrevit >nul 2>nul
if errorlevel 1 (
    echo ERROR: pyRevit no esta instalado o no esta en PATH.
    pause
    exit /b 1
)

where git >nul 2>nul
if errorlevel 1 (
    echo ERROR: Git no esta instalado o no esta en PATH.
    pause
    exit /b 1
)

if not exist "%EXT_DIR%" mkdir "%EXT_DIR%"
if not exist "%CONFIG_DIR%" mkdir "%CONFIG_DIR%"

echo Vinculando pyRevit con los Revit instalados...
pyrevit attach master default --installed
if errorlevel 1 (
    echo ERROR: No se pudo vincular pyRevit.
    echo Prueba primero con: pyrevit clones
    pause
    exit /b 1
)

if exist "%EXT_PATH%" (
    echo La extension ya existe. Actualizando...
    git -C "%EXT_PATH%" pull
    if errorlevel 1 (
        echo ERROR: No se pudo actualizar pyMenvic.
        pause
        exit /b 1
    )
) else (
    echo Clonando pyMenvic desde GitHub...
    git clone "%REPO_URL%" "%EXT_PATH%"
    if errorlevel 1 (
        echo ERROR: No se pudo clonar pyMenvic.
        pause
        exit /b 1
    )
)

if not exist "%CONFIG_FILE%" (
    type nul > "%CONFIG_FILE%"
)

findstr /R /C:"^\[pyMenvic\.extension\]$" "%CONFIG_FILE%" >nul
if errorlevel 1 (
    echo.>> "%CONFIG_FILE%"
    echo [pyMenvic.extension]>> "%CONFIG_FILE%"
    echo disabled = false>> "%CONFIG_FILE%"
    echo private_repo = false>> "%CONFIG_FILE%"
    echo username = "">> "%CONFIG_FILE%"
    echo password = "">> "%CONFIG_FILE%"
)

echo.
echo Instalacion completada correctamente.
echo Cierra y vuelve a abrir Revit 2024 o 2025.
pause
