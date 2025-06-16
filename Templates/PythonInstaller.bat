@echo off
setlocal

REM URL del instalador
set "PYTHON_URL=https://www.python.org/ftp/python/3.13.5/python-3.13.5-amd64.exe"

REM Ruta temporal para guardar el instalador
set "INSTALLER_PATH=%TEMP%\python-installer.exe"

REM Descargar el instalador usando PowerShell
echo Descargando Python...
powershell -Command "Invoke-WebRequest -Uri '%PYTHON_URL%' -OutFile '%INSTALLER_PATH%'"

IF NOT EXIST "%INSTALLER_PATH%" (
    echo ERROR: No se pudo descargar el instalador.
    exit /b 1
)

REM Ejecutar el instalador en modo silencioso
echo Instalando Python en modo silencioso...
"%INSTALLER_PATH%" /quiet InstallAllUsers=1 PrependPath=1 Include_test=0

REM Verificar instalación
echo Verificando instalacion de Python...
python --version

REM Fin
endlocal
pause
