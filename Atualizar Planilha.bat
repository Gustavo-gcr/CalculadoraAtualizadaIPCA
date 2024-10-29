@echo off


echo Atualizando bibliotecas Python...
pip install --upgrade pandas requests openpyxl streamlit

set /p mes="Digite o mes (em numero, por exemplo 09 para Setembro): "
set /p ano="Digite o ano (por exemplo 2024): "

python atualizar.py %mes% %ano%

pause
