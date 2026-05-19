[Setup]
AppName=BrewIns POS
AppVersion=2.1.0
DefaultDirName={localappdata}\BrewInsPOS
DefaultGroupName=BrewIns POS
OutputBaseFilename=BrewIns_POS_Setup_2.1.0
Compression=lzma
SolidCompression=yes
PrivilegesRequired=lowest
ArchitecturesInstallIn64BitMode=x64os
SetupIconFile=website\icon.ico
UninstallDisplayIcon={app}\icon.ico

[Files]
; Main app EXE
Source: website\release\BrewInsPOS_2.1.0.exe; DestDir: {app}; DestName: BrewInsPOS.exe; Flags: ignoreversion

; Updater EXE
Source: "website\release\updater.exe"; DestDir: "{app}"; Flags: ignoreversion

; Backend files
Source: "website\backend\*"; DestDir: "{app}\backend"; Flags: ignoreversion recursesubdirs createallsubdirs; Excludes: "pos.db,__pycache__\*,*.pyc"

; Existing database - only install if one does not already exist
Source: "website\backend\pos.db"; DestDir: "{app}\backend"; Flags: onlyifdoesntexist

; Templates - IMPORTANT: pulls admin/login.html, subnavs, reports, etc.
Source: "website\templates\*"; DestDir: "{app}\templates"; Flags: ignoreversion recursesubdirs createallsubdirs

; Static assets - CSS/JS/images
Source: "website\static\*"; DestDir: "{app}\static"; Flags: ignoreversion recursesubdirs createallsubdirs; Excludes: "__pycache__\*,*.pyc"

; Catalog backup files
Source: "website\catalog_backup\*"; DestDir: "{app}\catalog_backup"; Flags: ignoreversion recursesubdirs createallsubdirs

; Root support files
Source: "website\SchoolCafe_POS.xlsx"; DestDir: "{app}"; Flags: ignoreversion skipifsourcedoesntexist
Source: "website\version.txt"; DestDir: "{app}"; Flags: ignoreversion skipifsourcedoesntexist
Source: "website\icon.ico"; DestDir: "{app}"; Flags: ignoreversion

[Tasks]
Name: "desktopicon"; Description: "Create a desktop icon"; Flags: unchecked
Name: "startmenuicon"; Description: "Create a Start Menu shortcut"; Flags: checkedonce

[Icons]
Name: "{group}\BrewIns POS"; Filename: "{app}\BrewInsPOS.exe"; WorkingDir: "{app}"; IconFilename: "{app}\icon.ico"; Tasks: startmenuicon
Name: "{userdesktop}\BrewIns POS"; Filename: "{app}\BrewInsPOS.exe"; WorkingDir: "{app}"; IconFilename: "{app}\icon.ico"; Tasks: desktopicon

[Run]
Filename: "{app}\BrewInsPOS.exe"; Description: "Launch BrewIns POS"; WorkingDir: "{app}"; Flags: nowait postinstall runasoriginaluser
