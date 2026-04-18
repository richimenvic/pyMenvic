@echo off
title pyMenvic Installer (ZIP) for Revit 2024-2025
setlocal EnableExtensions

set "EXT_DIR=%AppData%\pyRevit\Extensions"
set "EXT_PATH=%EXT_DIR%\pyMenvic.extension"
set "CONFIG_DIR=%AppData%\pyRevit"
set "CONFIG_FILE=%CONFIG_DIR%\pyRevit_config.ini"
set "ZIP_URL=https://github.com/richimenvic/pyMenvic/archive/refs/heads/main.zip"
set "TEMP_ZIP=%TEMP%\pyMenvic_main.zip"
set "TEMP_UNZIP=%TEMP%\pyMenvic_main_unzip"

echo ==========================================
echo Installing pyMenvic from ZIP for Revit 2024-2025...
echo ==========================================
echo.

where pyrevit >nul 2>nul
if errorlevel 1 (
    echo ERROR: pyRevit no esta instalado o no esta en PATH.
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

echo Descargando ZIP desde GitHub...
powershell -NoProfile -ExecutionPolicy Bypass -Command "Invoke-WebRequest -Uri '%ZIP_URL%' -OutFile '%TEMP_ZIP%'"
if errorlevel 1 (
    echo ERROR: No se pudo descargar el ZIP.
    pause
    exit /b 1
)

if exist "%TEMP_UNZIP%" rmdir /s /q "%TEMP_UNZIP%"
mkdir "%TEMP_UNZIP%"

echo Descomprimiendo archivos...
powershell -NoProfile -ExecutionPolicy Bypass -Command "Expand-Archive -LiteralPath '%TEMP_ZIP%' -DestinationPath '%TEMP_UNZIP%' -Force"
if errorlevel 1 (
    echo ERROR: No se pudo descomprimir el ZIP.
    pause
    exit /b 1
)

if exist "%EXT_PATH%" (
    echo Eliminando instalacion anterior...
    rmdir /s /q "%EXT_PATH%"
)

echo Copiando pyMenvic.extension...
xcopy "%TEMP_UNZIP%\pyMenvic-main" "%EXT_PATH%" /E /I /Y >nul
if errorlevel 1 (
    echo ERROR: No se pudo copiar la extension.
    pause
    exit /b 1
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

del /q "%TEMP_ZIP%" >nul 2>nul
if exist "%TEMP_UNZIP%" rmdir /s /q "%TEMP_UNZIP%"

echo.
echo Instalacion completada correctamente.
echo Cierra y vuelve a abrir Revit 2024 o 2025.
pause
