@echo off
setlocal

REM Definir o nome da pasta do ambiente virtual
set "venv_dir=venv"

REM Verificar se o ambiente virtual já existe
if exist "%venv_dir%\Scripts\activate" (
    echo Ambiente virtual ja existe.
) else (
    echo Criando ambiente virtual...
    python -m venv %venv_dir%
)

echo Ativando ambiente virtual...
call %venv_dir%\Scripts\activate

echo Instalando dependencias...
pip install -r requirements.txt

echo Verificando a presenca do conversor.py...
if exist conversor.py (
    echo "conversor.py encontrado."
) else (
    echo "conversor.py não encontrado."
    pause
    exit /b 1
)

echo Executando o conversor.py...
python conversor.py

echo Configuracao e execucao concluidas.
pause

endlocal
