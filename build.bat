@echo off
pip install -r requirements.txt
python -m PyInstaller --onedir --windowed --icon=favicon.ico --name MailMerge --add-data "data/templates;data/templates" --add-data "translations;translations" --add-data "data/contacts.csv;data" --add-data "favicon.ico;." main.py
echo Build complete.

"C:\Program Files (x86)\Inno Setup 6\ISCC.exe" installer.iss
echo Installer created.
