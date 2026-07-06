@echo off
title Hapvida - Robo de Disparo de E-mail
color 1F
cls

echo.
echo  =====================================================
echo   HAPVIDA - ROBO DE DISPARO DE E-MAIL
echo  =====================================================
echo.

cd /d "%~dp0"

:: Verificar Python
python --version >nul 2>&1
if errorlevel 1 (
    echo  [ERRO] Python nao encontrado!
    echo  Instale em: https://www.python.org/downloads/
    echo  Marque "Add Python to PATH" durante a instalacao!
    pause & exit /b
)

echo  Instalando dependencias...
python -m pip install pandas openpyxl --quiet --disable-pip-version-check
echo  Dependencias OK!
echo.

echo  Escolha uma opcao:
echo   [1] Enviar e-mail de TESTE (recomendado antes do envio real)
echo   [2] Iniciar envio para todos os clientes pendentes
echo.
set /p opcao="Digite 1 ou 2 e pressione Enter: "

if "%opcao%"=="1" (
    set /p testeemail="Digite o e-mail de teste (ex: voce@gmail.com): "
    python robo_disparo_email.py --teste %testeemail%
) else (
    python robo_disparo_email.py
)

echo.
echo  =====================================================
echo   Processo encerrado.
echo   Confira o arquivo: log_envios.csv
echo  =====================================================
echo.
pause
