@echo off
echo.
echo  ============================================
echo    Nextcloud Uninstaller
echo  ============================================
echo.
echo  This will completely remove Nextcloud.
echo  All data (files, users, settings) will be deleted.
echo.
set /p confirm="  Type YES to confirm: "
if /i not "%confirm%"=="YES" (
    echo  Cancelled.
    pause & exit /b 0
)

echo.
echo  Stopping Nextcloud containers...
cd /d "C:\Nextcloud-LAN"
docker compose down -v

echo  Removing auto-start task...
schtasks /delete /tn "NextcloudAutoStart" /f >nul 2>&1

echo  Removing firewall rule...
netsh advfirewall firewall delete rule name="Nextcloud" >nul 2>&1

echo  Removing files...
cd /
rmdir /s /q "C:\Nextcloud-LAN" >nul 2>&1

echo.
echo  ============================================
echo   Nextcloud removed successfully.
echo  ============================================
pause
