@echo off
title BeefPassion OTE - Calculadora
color 1F
echo.
echo  =====================================================
echo   BeefPassion . Calculadora OTE  ^|  Stoic Capital
echo  =====================================================
echo.

cd /d "%~dp0"

echo  Verificando dependencias...
pip install -r requirements.txt --quiet

echo.
echo  Iniciando servidor...
echo  Acesse: http://localhost:5000
echo  Para encerrar: feche esta janela ou pressione Ctrl+C
echo.

start "" http://localhost:5000
python app.py

pause
