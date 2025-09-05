@echo off
set PLAYWRIGHT_BROWSERS_PATH=0
python -m pip install --upgrade pip
pip install -r requirements.txt
playwright install chromium
pip install pyinstaller
pyinstaller --onefile --name nkp_autofill --clean main.py
echo Done. See .\dist\nkp_autofill.exe
