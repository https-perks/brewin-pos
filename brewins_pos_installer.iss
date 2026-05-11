[Setup]
AppName=BrewIns POS
AppVersion=2.0.0
DefaultDirName={localappdata}\BrewInsPOS
OutputBaseFilename=BrewIns_POS_Setup
Compression=lzma
SolidCompression=yes
PrivilegesRequired=lowest
ArchitecturesInstallIn64BitMode=x64os
SetupIconFile=website\icon.ico

[Files]
; Main app EXE
Source: "website\release\BrewInsPOS_2.0.0.exe"; DestDir: "{app}"; DestName: "BrewInsPOS.exe"; Flags: ignoreversion
; Updater EXE
Source: "website\release\updater.exe"; DestDir: "{app}"; Flags: ignoreversion
; Backend PY files
Source: "website\backend\*"; DestDir: "{app}\backend"; Flags: ignoreversion recursesubdirs createallsubdirs
; Database file
Source: "website\backend\pos.db"; DestDir: "{app}\backend"; Flags: ignoreversion
; SQL schema for resets/rebuild
Source: "website\backend\models.sql"; DestDir: "{app}\backend"; Flags: ignoreversion
; Templates (HTML)
Source: "website\templates\*"; DestDir: "{app}\templates"; Flags: ignoreversion recursesubdirs createallsubdirs
; Static assets (JS, CSS)
Source: "website\static\*"; DestDir: "{app}\static"; Flags: ignoreversion recursesubdirs createallsubdirs
; Backup/restore utilities
Source: "website\catalog_backup\*"; DestDir: "{app}\catalog_backup"; Flags: ignoreversion recursesubdirs createallsubdirs
; Excel file
Source: "website\SchoolCafe_POS.xlsx"; DestDir: "{app}"; Flags: ignoreversion
; Icon
Source: "website\icon.ico"; DestDir: "{app}"; Flags: ignoreversion

[Tasks]
Name: "desktopicon"; Description: "Create a desktop icon"; Flags: unchecked
Name: "startmenuicon"; Description: "Pin to start"; Flags: unchecked

[Icons]
Name: "{group}\BrewIns POS"; Filename: "{app}\BrewInsPOS.exe"; IconFilename: "{app}\icon.ico"; Tasks: startmenuicon
Name: "{userdesktop}\BrewIns POS"; Filename: "{app}\BrewInsPOS.exe"; IconFilename: "{app}\icon.ico"; Tasks: desktopicon

[Run]
Filename: "{app}\BrewInsPOS.exe"; Description: "Launch BrewIns POS"; Flags: nowait postinstall runasoriginaluser
