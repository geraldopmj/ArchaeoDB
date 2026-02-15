@echo off
cd /d "%~dp0"

if not exist ".venv\Scripts\pythonw.exe" (
  echo ERRO: venv nao encontrado. Rode setup.bat primeiro.
  exit /b 1
)

if not exist "main.py" (
  echo ERRO: main.py nao encontrado.
  exit /b 1
)

start "" ".venv\Scripts\pythonw.exe" "main.py"
