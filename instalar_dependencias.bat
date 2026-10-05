@echo off
cd /d "%~dp0"
echo Instalando dependencias necesarias para Kardex...
py -m pip install --upgrade pip
py -m pip install -r requirements_app.txt
if errorlevel 1 (
    echo.
    echo No se pudo instalar con py. Probando con python...
    python -m pip install --upgrade pip
    python -m pip install -r requirements_app.txt
)
echo.
echo Proceso terminado. Si no hubo errores, abra el sistema con abrir_kardex.bat o desde VS Code.
pause
