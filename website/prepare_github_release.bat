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

REM Update version.json automatically
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
echo Upload the following file to your GitHub Release:
echo   release\BrewInsPOS_%VERSION%.exe
echo
echo Then commit and push your updated version.json:
echo   git add version.json
echo   git commit -m "Release v%VERSION%"
echo   git push
echo --------------------------------------------------------

pause
