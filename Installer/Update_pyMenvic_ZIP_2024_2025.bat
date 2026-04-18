@echo off
title pyMenvic Updater (ZIP) for Revit 2024-2025
setlocal EnableExtensions

set "EXT_DIR=%AppData%\pyRevit\Extensions"
set "EXT_PATH=%EXT_DIR%\pyMenvic.extension"
set "ZIP_URL=https://github.com/richimenvic/pyMenvic/archive/refs/heads/main.zip"
set "TEMP_ZIP=%TEMP%\pyMenvic_main.zip"
set "TEMP_UNZIP=%TEMP%\pyMenvic_main_unzip"

echo ==========================================
echo Updating pyMenvic from ZIP for Revit 2024-2025...
echo ==========================================
echo.

if not exist "%EXT_PATH%" (
    echo ERROR: No existe la carpeta:
    echo %EXT_PATH%
    echo.
    echo Ejecuta primero el instalador.
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

echo Reemplazando archivos...
rmdir /s /q "%EXT_PATH%"
xcopy "%TEMP_UNZIP%\pyMenvic-main" "%EXT_PATH%" /E /I /Y >nul
if errorlevel 1 (
    echo ERROR: No se pudo actualizar la extension.
    pause
    exit /b 1
)

del /q "%TEMP_ZIP%" >nul 2>nul
if exist "%TEMP_UNZIP%" rmdir /s /q "%TEMP_UNZIP%"

echo.
echo Actualizacion completada.
echo Cierra y vuelve a abrir Revit 2024 o 2025.
pause
