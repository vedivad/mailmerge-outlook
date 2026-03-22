@echo off
pip install -r requirements.txt
pyinstaller --onefile --windowed --name MailMerge main.py
echo Build complete. See dist\MailMerge.exe
