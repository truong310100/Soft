@echo off
CHCP 1258 >nul 2>&1
CHCP 65001 >nul 2>&1
echo ===========================
echo   SETUP PRINT SERVER VHU
echo ===========================
echo.
set /p username="Username: "
set /p password="Password: "
set printer=VHU_Printsv
set domain=vanhien
set pass=123456
cmdkey /add:172.17.34.18 /user:%domain%\%username% /pass:%password%%pass%
cmdkey /add:172.17.34.15 /user:%domain%\%username% /pass:%password%%pass%
ping -n 3 127.0.0.1 >nul
cd C:\
net use Z: \\172.17.34.15\home-vhu$\%username%
start "" explorer \\172.17.34.18\%printer%
control printers
echo.
echo Nhớ set Authentication trong VHU_Printsv
pause