@echo off
:: KAAVAL Build Script — Creates single-file EXE
:: Usage: Double-click this file or run in terminal
cd /d "%~dp0"

echo ============================================
echo   KAAVAL - Building Production EXE...
echo ============================================
echo.

:: Activate virtual environment if it exists
if exist ".venv\Scripts\activate.bat" (
    call .venv\Scripts\activate.bat
)

:: Install PyInstaller if needed
pip install pyinstaller --quiet

:: Clean old build artifacts
if exist "dist" rd /s /q "dist"
if exist "build" rd /s /q "build"

:: Build using spec file
pyinstaller kaaval.spec --clean --noconfirm

echo.
echo ============================================
if exist "dist\KAAVAL.exe" (
    echo   BUILD SUCCESSFUL!
    echo   EXE Location: dist\KAAVAL.exe
    echo.
    echo   Creating desktop shortcut...
    powershell -Command "$WshShell = New-Object -ComObject WScript.Shell; $Shortcut = $WshShell.CreateShortcut([Environment]::GetFolderPath('Desktop') + '\KAAVAL.lnk'); $Shortcut.TargetPath = '%~dp0dist\KAAVAL.exe'; $Shortcut.WorkingDirectory = '%~dp0'; $Shortcut.IconLocation = '%~dp0kaaval\kaaval.ico'; $Shortcut.Description = 'KAAVAL - Tamil Nadu Police EOW'; $Shortcut.Save()"
    echo   Desktop shortcut created!
) else (
    echo   BUILD FAILED - Check errors above
)
echo ============================================
pause
