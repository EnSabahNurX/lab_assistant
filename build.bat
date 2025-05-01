@echo off
REM Caminho local para o Python
set PYTHON_DIR=%~dp0Python
set PYINSTALLER_EXE=%PYTHON_DIR%\Scripts\pyinstaller.exe

REM Vai para a pasta onde está este .bat
cd /d "%~dp0"

REM Compila o main.py para um executável sem terminal
%PYINSTALLER_EXE% --noconsole --onefile CPK.py

REM Verifica se o .exe foi criado e copia para a pasta atual
if exist dist\CPK.exe (
    copy /Y dist\CPK.exe .
    echo ✅ Executável CPK.exe copiado para a pasta atual.
) else (
    echo ❌ Erro: O executável não foi criado.
    pause
    exit /b
)

REM Limpeza de ficheiros temporários
rmdir /s /q dist
rmdir /s /q build
del /q CPK.spec

echo 🧹 Limpeza concluída.
pause
