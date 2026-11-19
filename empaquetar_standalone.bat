@echo off
title Empaquetar Version Standalone - ReconcileAI
cd /d "%~dp0"

echo ======================================================================
echo   EMPAQUETADOR STANDALONE - RECONCILEAI
echo   Genera ejecutable portable unico (sin requerir .NET runtime)
echo ======================================================================
echo.

echo [1/2] Compilando Acr.Desktop Standalone (Windows x64 Single-File comprimido)...
dotnet publish "src\dotnet\Acr.Desktop\Acr.Desktop.csproj" -c Release -r win-x64 --self-contained true -p:PublishSingleFile=true -p:IncludeNativeLibrariesForSelfExtract=true -p:EnableCompressionInSingleFile=true -o "publish\desktop-standalone"

if %ERRORLEVEL% NEQ 0 (
    echo [ERROR] Fallo la publicacion de Acr.Desktop.
    pause
    exit /b 1
)

echo.
echo [2/2] Empaquetando Web Edition (React + Vite)...
if exist "web\package.json" (
    cd web
    call npm run build
    cd ..
    powershell -NoProfile -Command "Compress-Archive -Path 'web\dist\*' -DestinationPath 'publish\web-edition.zip' -Force"
)

echo.
echo ======================================================================
echo   EMPAQUETADO COMPLETADO EXITOSAMENTE
echo ======================================================================
echo.
echo Ubicaciones de los paquetes listos para copiar y usar en cualquier PC:
echo.
echo  * Desktop App Portable:
echo    publish\desktop-standalone\Acr.Desktop.exe
echo    (Un solo archivo .exe de ~62MB, 100%% autocontenido, doble clic y listo)
echo.
echo  * Web Edition (Bundle Estatico):
echo    publish\web-edition.zip
echo.
pause
