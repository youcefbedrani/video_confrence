@echo off
echo.
echo  ============================================
echo    Nextcloud Enterprise Installer Builder
echo    Windows 10/11 - Docker Engine (No Login)
echo  ============================================
echo.

python --version >nul 2>&1
if errorlevel 1 (
    echo  ERROR: Python not found.
    echo  Download: https://www.python.org/downloads/
    pause & exit /b 1
)

echo  [1/2] Installing Dependencies...
python -m pip install pyinstaller cryptography Pillow --quiet --upgrade

echo  [2/3] Generating App Icon...
python -c "from PIL import Image; Image.open('../logo.png').save('logo.ico', format='ICO', sizes=[(256, 256)])"

echo  [3/3] Building .exe...
python -m PyInstaller ^
    --onefile ^
    --windowed ^
    --uac-admin ^
    --clean ^
    --name "NextcloudInstaller" ^
    --icon "logo.ico" ^
    --add-data "../logo.png;." ^
    --collect-all cryptography ^
    --collect-all PIL ^
    installer.py

echo.
if exist "dist\NextcloudInstaller.exe" (
    echo  ============================================
    echo   SUCCESS!
    echo   dist\NextcloudInstaller.exe
    echo.
    echo   Send this ONE file to your client.
    echo   They double-click it.
    echo   No Docker account. No login. No signup.
    echo  ============================================
) else (
    echo  BUILD FAILED - check errors above.
)
echo.
pause
