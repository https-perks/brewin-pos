@echo off
setlocal

REM Read version
set /p VERSION=<version.txt

echo Preparing GitHub release %VERSION%...

REM Make sure release EXE exists
if not exist release\BrewInsPOS_%VERSION%.exe (
    echo ERROR: Build the app first using build_pos_exe.bat
    pause
    exit /b
)

REM --------------------------------------------------------
REM Update Inno Setup .iss file
REM --------------------------------------------------------

set "INNO_FILE=..\brewins_pos_installer.iss"

if exist "%INNO_FILE%" (
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
) else (
    echo ERROR: Inno Setup script not found at %INNO_FILE%
    pause
    exit /b 1
)

REM --------------------------------------------------------
REM Build installer with Inno Setup
REM --------------------------------------------------------

echo Building installer with Inno Setup...

set "ISCC=C:\Program Files (x86)\Inno Setup 6\ISCC.exe"

if not exist "%ISCC%" (
    set "ISCC=C:\Program Files\Inno Setup 6\ISCC.exe"
)

if not exist "%ISCC%" (
    echo ERROR: Could not find ISCC.exe.
    echo Check your Inno Setup install path.
    pause
    exit /b 1
)

if not exist "%INNO_FILE%" (
    echo ERROR: Inno Setup script not found:
    echo %INNO_FILE%
    pause
    exit /b 1
)

"%ISCC%" "%INNO_FILE%"

if errorlevel 1 (
    echo ERROR: Inno Setup build failed.
    pause
    exit /b 1
)

echo Installer build complete.


REM --------------------------------------------------------
REM Update version.json automatically
REM --------------------------------------------------------

echo Updating version.json...

(
echo {
echo   "latest_version": "%VERSION%",
echo   "download_url": "https://github.com/https-perks/brewin-pos/releases/download/v%VERSION%/BrewInsPOS_%VERSION%.exe"
echo }
)> version.json

REM Copy version.json to release folder for convenience
copy /y version.json release\version_%VERSION%.json >nul

echo --------------------------------------------------------
echo GitHub Release Prep Complete!
echo
echo Updated:
echo   version.json
echo   %INNO_FILE%
echo
echo Upload the following file to your GitHub Release:
echo   release\BrewInsPOS_%VERSION%.exe
echo
echo Then commit and push your updated version.json and .iss:
echo   git add version.json ../BrewIns_POS_Setup.iss
echo   git commit -m "Release v%VERSION%"
echo   git push
echo --------------------------------------------------------

pause