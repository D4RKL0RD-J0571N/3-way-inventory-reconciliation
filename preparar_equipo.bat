@echo off
title Configurar ReconcileAI en este equipo
cd /d "%~dp0"

echo ======================================================================
echo   Configurador Rapido de ReconcileAI para Nuevos Equipos
echo ======================================================================
echo.

where python >nul 2>nul
if %ERRORLEVEL% NEQ 0 (
    echo [ERROR] Python no esta instalado o no se encuentra en el PATH.
    echo Por favor instala Python 3.10+ marcando la casilla "Add Python to PATH".
    pause
    exit /b 1
)

echo 1. Creando entorno virtual aislado (.venv)...
python -m venv .venv

echo 2. Actualizando pip e instalando dependencias requeridas...
".venv\Scripts\python.exe" -m pip install --upgrade pip
".venv\Scripts\python.exe" -m pip install -r requirements.txt

echo.
echo ======================================================================
echo   CONFIGURACION COMPLETADA EXITOSAMENTE!
echo ======================================================================
echo Ahora puedes iniciar la aplicacion con doble clic en:
echo   'iniciar_asistente.bat'
echo ======================================================================
pause
