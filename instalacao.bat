@echo off
@echo Instalando dependencias json, pandas, pywin32, openpyxl e pillow...
pip install json pandas pywin32 openpyxl pillow

if %errorlevel% neq 0 (
    echo Falha ao instalar bibliotecas.
    pause
    exit /b 1
)

echo Depedenncias instaladas com sucesso.