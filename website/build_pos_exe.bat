@echo off
setlocal

REM --------------------------------------------------------
REM Activate virtual environment
REM --------------------------------------------------------
call venv\Scripts\activate.bat
if errorlevel 1 (
    echo ERROR: Could not activate venv.
    pause
    exit /b
)

REM --------------------------------------------------------
REM Read version number
REM --------------------------------------------------------
set /p VERSION=<version.txt

echo.
echo ================================
echo   Building BrewIns POS v%VERSION%
echo ================================
echo.


REM --------------------------------------------------------
REM Clean old folders
REM --------------------------------------------------------
rmdir /s /q build 2>nul
rmdir /s /q dist 2>nul


REM --------------------------------------------------------
REM Run PyInstaller
REM --------------------------------------------------------
pyinstaller ^
  --windowed ^
  --onefile ^
  --name BrewInsPOS ^
  --icon=icon.ico ^
  --collect-all backend ^
  --add-data "backend;backend" ^
  --add-data "backend/pos.db;backend" ^
  --hidden-import backend ^
  --hidden-import backend.db_ops ^
  --hidden-import backend.build_database ^
  --hidden-import backend.backup_restore ^
  --add-data "templates;templates" ^
  --add-data "static;static" ^
  --add-data "catalog_backup;catalog_backup" ^
  --add-data "SchoolCafe_POS.xlsx;." ^
  --add-data "icon.ico;." ^
  app_window.py

if errorlevel 1 (
    echo ERROR: PyInstaller failed!
    pause
    exit /b
)


REM --------------------------------------------------------
REM Rename exe
REM --------------------------------------------------------
rename dist\BrewInsPOS.exe BrewInsPOS_%VERSION%.exe


REM --------------------------------------------------------
REM Move exe
REM --------------------------------------------------------
if not exist release mkdir release
move /Y dist\BrewInsPOS_%VERSION%.exe release\ >nul


REM --------------------------------------------------------
REM Build updater
REM --------------------------------------------------------
pyinstaller --onefile --windowed --name updater updater.py
copy /Y dist\updater.exe release\ >nul

echo.
echo Build completed: release\BrewInsPOS_%VERSION%.exe
pause
