@echo off
title ReconcileAI - Copiloto Universal 3-Way
cd /d "%~dp0"

echo ======================================================================
echo   ReconcileAI - Copiloto Universal de Conciliacion 3-Way
echo   Auditoria Inteligente de Inventarios y Recepcion de Mercaderia
echo ======================================================================
echo.

:: 1. Verificar si existe ejecutable compilado en dist/ReconcileAI
if exist "dist\ReconcileAI\ReconcileAI.exe" (
    echo Iniciando version compilada independiente...
    start "" "dist\ReconcileAI\ReconcileAI.exe"
    exit /b 0
)

:: 2. Verificar si existe entorno virtual local .venv
if exist ".venv\Scripts\python.exe" (
    echo Iniciando mediante entorno virtual local (.venv)...
    ".venv\Scripts\python.exe" run_standalone.py
    goto FINISH
)

:: 3. Verificar si python esta en el PATH
where python >nul 2>nul
if %ERRORLEVEL% EQU 0 (
    echo Iniciando mediante Python del sistema...
    python run_standalone.py
    goto FINISH
)

:: 4. Si no se detecta Python ni ejecutable
echo [AVISO] No se detecto Python instalado en esta computadora.
echo.
echo Opciones para ejecutar en este equipo:
echo  1. Si tienes Python: ejecuta 'preparar_equipo.bat' para instalar dependencias.
echo  2. Si es una PC de bodega/oficina sin permisos: compila el paquete portable
echo     con 'python tools/build_exe.py' y copia la carpeta 'dist/ReconcileAI'.
echo.
pause
exit /b 1

:FINISH
pause
