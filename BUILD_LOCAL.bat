@echo off
cd /d "%~dp0"
python -m pip install --upgrade pip
pip install -r requirements.txt
pip install pyinstaller
pyinstaller --noconfirm --clean --windowed --name "Apurador Fiscal" --icon assets\app_icon.ico --add-data "assets;assets" main.py
echo.
echo Programa compilado em: dist\Apurador Fiscal\
pause
