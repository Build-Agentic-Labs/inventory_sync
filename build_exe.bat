@echo off
echo ============================================
echo Building Inventory Sync with Ghostscript
echo ============================================
echo.

REM Install PyInstaller if needed
echo Ensuring PyInstaller is installed...
pip install pyinstaller

echo.
echo Building executable...
echo This will include Ghostscript automatically if found.
echo.

REM Delete old build if exists
if exist "dist\InventorySync.exe" del "dist\InventorySync.exe"

python -m PyInstaller inventory_sync.spec --clean

echo.
echo ============================================
if exist "dist\InventorySync.exe" (
    echo BUILD SUCCESSFUL!
    echo.
    echo Your exe is at: dist\InventorySync.exe
    echo.
    echo The exe includes:
    echo   - Inventory Sync application
    echo   - Ghostscript ^(if found on this system^)
    echo   - All required Python libraries
    echo.
    echo Users can just run InventorySync.exe - no additional installation needed!
) else (
    echo BUILD FAILED - Check errors above
)
echo ============================================
pause
