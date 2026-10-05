Set-Location -Path $PSScriptRoot
Write-Host "Instalando dependencias necesarias para Kardex..."
python -m pip install --upgrade pip
python -m pip install -r requirements_app.txt
Write-Host "Proceso terminado."
Read-Host "Presione Enter para cerrar"
