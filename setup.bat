@echo off
setlocal enabledelayedexpansion

set "TARGET_DIR=%~dp0Python"
if not exist "%TARGET_DIR%" (
    mkdir "%TARGET_DIR%"
)
echo Diretório de instalação: %TARGET_DIR%

set "PYTHON_URL=https://www.python.org/ftp/python/3.10.11/python-3.10.11-amd64.exe"
set INSTALLER=python-installer.exe

echo [2/5] A descarregar o instalador...
powershell -Command "Invoke-WebRequest -Uri '%PYTHON_URL%' -OutFile '%INSTALLER%'"

echo [3/5] A instalar Python localmente na pasta atual...
start /wait "" "%INSTALLER%" InstallAllUsers=0 PrependPath=0 Include_tcltk=1 Include_pip=1 TargetDir="%TARGET_DIR%" ForceInstall=1

echo [4/5] A verificar instalação...
set "PYTHON_EXE=%TARGET_DIR%\python.exe"
if not exist "%PYTHON_EXE%" (
    echo ❌ Python não foi instalado corretamente em %TARGET_DIR%.
    pause
    exit /b
)

echo ✅ Python encontrado em: %PYTHON_EXE%

echo [5/5] A instalar bibliotecas opcionais...
"%PYTHON_EXE%" -m ensurepip
"%PYTHON_EXE%" -m pip install --upgrade pip
"%PYTHON_EXE%" -m pip install pandas openpyxl matplotlib numpy

echo ✅ Instalação finalizada com sucesso!
"%PYTHON_EXE%" --version
pause
