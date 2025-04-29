@echo off
REM 1. Definir o diretório de instalação do Python
set "TARGET_DIR=%~dp0Python"
set "PYTHON_EXE=%TARGET_DIR%\python.exe"

REM 2. Verificar se o Python está instalado
if not exist "%PYTHON_EXE%" (
    echo ❌ Python não foi encontrado em "%TARGET_DIR%".
    pause
    exit /b
)

REM 3. Executar o script main.py com o Python instalado localmente
echo [INFO] Executando o script main.py...
"%PYTHON_EXE%" "%~dp0main.py"

echo.
echo ✅ Execução concluída.
pause
