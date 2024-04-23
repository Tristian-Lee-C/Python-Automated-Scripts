@echo off
cd /d "%APPDATA%\..\Local\Programs\Python\Python310"
python -m pip install --upgrade pip
python -m ensurepip --upgrade
python -m pip install openpyxl
python -m pip install --upgrade pywin32
python Scripts/pywin32_postinstall.py -install
python -m pip install tk
python -m pip install mttkinter
python -m pip install pyinstaller
pause