@echo off
cd /d "%~dp0"
echo Verificando dependencias...
py -c "import openpyxl, docx" 2>NUL
if errorlevel 1 (
    echo Faltan dependencias. Se instalaran automaticamente.
    call instalar_dependencias.bat
)
echo Iniciando Kardex...
py app_kardex.py
if errorlevel 1 (
    echo.
    echo No se pudo iniciar con py. Probando con python...
    python app_kardex.py
)
pause
