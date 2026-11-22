@echo off
chcp 65001 >nul
title Library Hotspot
net session >nul 2>&1
if %errorlevel% neq 0 (
  echo Requesting administrator privileges...
  powershell -NoProfile -Command "Start-Process -FilePath '%~f0' -Verb RunAs"
  exit /b
)

echo.
echo == Starting Wi-Fi hotspot for the library ==
echo SSID: Maktaba
echo Password: 00000000
echo.

netsh wlan set hostednetwork mode=allow ssid=Maktaba key=00000000
netsh wlan start hostednetwork
echo.
netsh wlan show hostednetwork | findstr /C:"SSID" /C:"Status" /C:"Mode"
echo.

echo If Status = Running, now run:
echo   "%~dp0تشغيل المكتبة.bat" --lan-ip 192.168.137.1
echo Then students join Wi-Fi "Maktaba" and open http://192.168.137.1/
echo.
echo If it fails, use Windows Settings:  ms-settings:mobile-devices/hotspot
echo.
pause
