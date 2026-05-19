@echo off
setlocal

set /p VERSION=<version.txt

echo Preparing GitHub release %VERSION%...

if not exist release\BrewInsPOS_%VERSION%.exe (
    echo ERROR: Build the app first using build_pos_exe.bat
    pause
    exit /b 1
)

set "INNO_FILE=..\brewins_pos_installer.iss"

if not exist "%INNO_FILE%" (
    echo ERROR: Inno Setup script not found at %INNO_FILE%
    pause
    exit /b 1
)

echo Updating Inno Setup script...

powershell -NoProfile -ExecutionPolicy Bypass -Command ^
  "$p = '%INNO_FILE%';" ^
  "$v = '%VERSION%';" ^
  "$lines = Get-Content $p;" ^
  "$lines = $lines | ForEach-Object {" ^
  "  if ($_ -match '^AppVersion=') { 'AppVersion=' + $v }" ^
  "  elseif ($_ -match '^OutputBaseFilename=') { 'OutputBaseFilename=BrewIns_POS_Setup_' + $v }" ^
  "  elseif ($_ -match 'Source: ""website\\release\\BrewInsPOS_.*\.exe""') { 'Source: ""website\release\BrewInsPOS_' + $v + '.exe""; DestDir: ""{app}""; DestName: ""BrewInsPOS.exe""; Flags: ignoreversion' }" ^
  "  else { $_ }" ^
  "};" ^
  "Set-Content -Path $p -Value $lines -Encoding UTF8"

if errorlevel 1 (
    echo ERROR: Failed to update Inno Setup script.
    pause
    exit /b 1
)

echo Building installer with Inno Setup...

set "ISCC=C:\Program Files (x86)\Inno Setup 6\ISCC.exe"

if not exist "%ISCC%" (
    set "ISCC=C:\Program Files\Inno Setup 6\ISCC.exe"
)

if not exist "%ISCC%" (
    echo ERROR: Could not find ISCC.exe.
    pause
    exit /b 1
)

"%ISCC%" "%INNO_FILE%"

if errorlevel 1 (
    echo ERROR: Inno Setup build failed.
    pause
    exit /b 1
)

set "INSTALLER=..\Output\BrewIns_POS_Setup_%VERSION%.exe"

if exist "%INSTALLER%" (
    copy /Y "%INSTALLER%" "release\BrewIns_POS_Setup_%VERSION%.exe" >nul
) else (
    echo WARNING: Installer was not found at:
    echo %INSTALLER%
)

echo Updating version.json...

(
echo {
echo   "latest_version": "%VERSION%",
echo   "download_url": "https://github.com/https-perks/brewin-pos/releases/download/v%VERSION%/BrewInsPOS_%VERSION%.exe"
echo }
)> version.json

copy /Y version.json release\version_%VERSION%.json >nul

echo --------------------------------------------------------
echo GitHub Release Prep Complete!
echo.
echo Upload BOTH files to GitHub Release v%VERSION%:
echo   release\BrewInsPOS_%VERSION%.exe
echo   release\BrewIns_POS_Setup_%VERSION%.exe
echo.
echo version.json points to:
echo   BrewInsPOS_%VERSION%.exe
echo.
echo Commit:
echo   git add version.json ../brewins_pos_installer.iss
echo   git commit -m "Release v%VERSION%"
echo   git push
echo --------------------------------------------------------

pause