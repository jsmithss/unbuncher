@echo off
REM Build script for Excel Unmerger executable
REM Run this to create the standalone executable

echo ============================================
echo Building Excel Unmerger Executable
echo ============================================
echo.

REM Check if Python is available
python --version >nul 2>&1
if errorlevel 1 (
    echo ERROR: Python is not installed or not in PATH
    echo.
    echo Please install Python from: https://www.python.org/downloads/
    echo Make sure to check "Add Python to PATH" during installation
    pause
    exit /b 1
)

echo Found Python version:
python --version

REM Check if required packages are installed
echo.
echo Checking required packages...
python -c "import openpyxl" >nul 2>&1
if errorlevel 1 (
    echo Installing openpyxl...
    python -m pip install openpyxl
    if errorlevel 1 (
        echo ERROR: Failed to install openpyxl
        pause
        exit /b 1
    )
) else (
    echo ✓ openpyxl already installed
)

python -c "import PyInstaller" >nul 2>&1
if errorlevel 1 (
    echo Installing PyInstaller...
    python -m pip install pyinstaller
    if errorlevel 1 (
        echo ERROR: Failed to install PyInstaller
        pause
        exit /b 1
    )
) else (
    echo ✓ PyInstaller already installed
)

REM Clean previous builds
echo Cleaning previous builds...
if exist "dist" rmdir /s /q "dist"
if exist "build" rmdir /s /q "build"
if exist "*.spec" del /q "*.spec"

REM Build the executable
echo Building executable...
python -m PyInstaller --onefile ^
    --name="Excel_Unmerger" ^
    --icon=unbunch.ico ^
    --console ^
    --distpath="./dist" ^
    --workpath="./build" ^
    --specpath="." ^
    unmerge_excel.py

if errorlevel 1 (
    echo ERROR: Build failed
    pause
    exit /b 1
)

REM Copy additional files to dist folder
echo Copying additional files...
copy "README.txt" "dist\" >nul 2>&1
copy "*.bat" "dist\" >nul 2>&1

echo.
echo ============================================
echo Build completed successfully!
echo ============================================
echo.
echo The executable is located in the 'dist' folder:
echo   Excel_Unmerger.exe
echo.
echo You can distribute the entire 'dist' folder to your team.
echo.
pause
