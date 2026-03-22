[Setup]
AppName=MailMerge
AppVersion=1.0
DefaultDirName={localappdata}\MailMerge
DefaultGroupName=MailMerge
OutputBaseFilename=MailMerge-Setup
Compression=lzma2
SolidCompression=yes
PrivilegesRequired=lowest
SetupIconFile=favicon.ico
UninstallDisplayIcon={app}\MailMerge.exe

[Files]
; PyInstaller --add-data already bundles templates, contacts.csv, and favicon.ico into dist\MailMerge\
Source: "dist\MailMerge\*"; DestDir: "{app}"; Flags: ignoreversion recursesubdirs createallsubdirs

[Icons]
Name: "{group}\MailMerge"; Filename: "{app}\MailMerge.exe"
Name: "{userdesktop}\MailMerge"; Filename: "{app}\MailMerge.exe"
