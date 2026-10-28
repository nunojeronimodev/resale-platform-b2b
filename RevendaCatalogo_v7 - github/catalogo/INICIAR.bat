@echo off
cd /d "%~dp0"
title Catalogo
color 1F
cls

echo.
echo  ================================
echo    Catalogo Comercial
echo  ================================
echo.

if not exist ".venv\Scripts\activate.bat" (
    echo  A criar ambiente virtual...
    python -m venv .venv
)

call .venv\Scripts\activate.bat

echo  A verificar dependencias...
pip install -r requirements.txt -q --disable-pip-version-check

echo  A iniciar servidor em http://localhost:5050
echo.

start "" http://localhost:5050
python app.py

pause
