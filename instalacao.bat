@echo off
echo Instalando bibliotecas pandas, openpyxl e requests...
pip install pandas openpyxl requests

if %errorlevel% neq 0 (
    echo Falha ao instalar bibliotecas.
    pause
    exit /b 1
)

echo Depedenncias instaladas com sucesso.