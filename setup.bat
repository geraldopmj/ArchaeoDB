@echo off
setlocal enabledelayedexpansion

cd /d "%~dp0"

if not exist requirements.txt (
  echo ERRO: requirements.txt nao encontrado na raiz do projeto.
  exit /b 1
)

if not exist .venv (
  echo Criando venv em .venv ...
  py -3 -m venv .venv
  if errorlevel 1 (
    echo ERRO: falha ao criar venv. Verifique se o Python foi instalado e o launcher "py" existe.
    exit /b 1
  )
) else (
  echo venv .venv ja existe.
)

echo Atualizando pip ...
".venv\Scripts\python.exe" -m pip install --upgrade pip

echo Instalando dependencias ...
".venv\Scripts\python.exe" -m pip install -r requirements.txt
if errorlevel 1 (
  echo ERRO: falha ao instalar requirements.
  exit /b 1
)

echo OK: ambiente pronto.
pause
