@echo off
CHCP 1258 >nul 2>&1
CHCP 65001 >nul 2>&1
Title Support User

::===========================================================================
@REM fsutil dirty query %systemdrive%  >nul 2>&1 || (
@REM echo.
@REM echo =====================================================
@REM echo   BẠN CẦN CHẠY VỚI QUYỀN "QUẢN TRỊ VIÊN" ĐỂ TIẾP TỤC
@REM echo      YOU NEED "RUN AS ADMINISTRATOR" TO CONTINUE
@REM echo =====================================================
@REM echo.
@REM pause & exit
@REM )
::===========================================================================

mode con: cols=84 lines=35
chcp 65001 >nul
color f0

>nul 2>&1 "%SYSTEMROOT%\system32\cacls.exe" "%SYSTEMROOT%\system32\config\system"
if '%errorlevel%' NEQ '0' (
    echo  Chay CMD Với Quyền Quản Trị - Run as Administrator...
    goto goUAC 
) else (
 goto goADMIN )

:goUAC
    echo Set UAC = CreateObject^("Shell.Application"^) > "%temp%\getadmin.vbs"
    set params = %*:"=""
    echo UAC.ShellExecute "cmd.exe", "/c %~s0 %params%", "", "runas", 1 >> "%temp%\getadmin.vbs"
    "%temp%\getadmin.vbs"
    del "%temp%\getadmin.vbs"
    exit /B

:goADMIN
    pushd "%CD%"
    CD /D "%~dp0"
:main
cls
::===========================================================================
echo.
echo   _______________________________________________________________________________
echo  ^| I. Fix Error                         ^| II. Setup                              ^|
echo  ^|  1. On/Off Uac        = Nhấn phím A  ^|  1. Active Windows      = Nhấn phím F  ^|
echo  ^|  2. Error Wifi        = Nhấn phím B  ^|  2. Set Password Device = Nhấn phím G  ^|
echo  ^|  3. Error 0x000007c   = Nhấn phím C  ^|  3. Install Azure       = Nhấn phím H  ^|
echo  ^|  4. Error 0x0000011b  = Nhấn phím D  ^|  4. Install Font HH     = Nhấn phím I  ^|
echo  ^|  5. Reset Onedrive    = Nhấn phím E  ^|  5. Install Timezone    = Nhấn phím J  ^|
echo  ^|______________________________________^|  6. Key VNC             = Nhấn phím K  ^|
echo  ^| III. ###                             ^|  7. Merge Regedit       = Nhấn phím L  ^|
echo  ^|  1. Reset Password    = Nhấn phím O  ^|  8. Install Apps Fast   = Nhấn phím M  ^|
echo  ^|  2. Open Aio Tools    = Nhấn phím P  ^|  9. Rename This PC      = Nhấn phím N  ^|
echo  ^|  3. Optimize Windowns = Nhấn phím Q  ^|  10.Backup Driver       = Nhấn phím S  ^|
echo  ^|  4. Googledrive       = Nhấn phím R  ^|           Thoát = Nhấn phím T          ^|
echo  ^|______________________________________^|________________________________________^|
echo.

choice /N /C ABCDEFGHIJKLMNOPQRST /M "* Nhập Lựa Chọn Của Bạn : "

if ERRORLEVEL  20 goto  :Exit T
if ERRORLEVEL  19 goto  :BackupDriver S
if ERRORLEVEL  18 goto  :GoogleDrive R
if ERRORLEVEL  17 goto  :OptimizeWindows Q
if ERRORLEVEL  16 goto  :AioTools P
if ERRORLEVEL  15 goto  :ResetPassword O
if ERRORLEVEL  14 goto  :RenameThisPC N
if ERRORLEVEL  13 goto  :InstallAppsFast M
if ERRORLEVEL  12 goto  :MergeRegedit L
if ERRORLEVEL  11 goto  :KeyVNC K
if ERRORLEVEL  10 goto  :Timezone J
if ERRORLEVEL  9 goto  :FontHH I
if ERRORLEVEL  8 goto  :Azure H
if ERRORLEVEL  7 goto  :SetPasswordDevice G
if ERRORLEVEL  6 goto  :ActiveWindows F
if ERRORLEVEL  5 goto  :ResetOnedrive E
if ERRORLEVEL  4 goto  :Error0x0000011b D
if ERRORLEVEL  3 goto  :Error0x000007c C
if ERRORLEVEL  2 goto  :ErrorWifi B
if ERRORLEVEL  1 goto :UAC A

:UAC
start UAC.cmd
goto :main

:ErrorWifi
netsh winsock reset
ipconfig /release
netsh int ip reset
ipconfig /renew
ipconfig /flushdns
shutdown /r /t 10
timeout 5
goto :main

:Error0x000007c
reg add "HKEY_LOCAL_MACHINE\SYSTEM\CurrentControlSet\Policies\Microsoft\FeatureManagement\Overrides" /v "713073804" /t REG_DWORD /d 0 /f
echo Đã hoàn tất...!
pause & goto :main

:Error0x0000011b
reg add HKEY_LOCAL_MACHINE\System\CurrentControlSet\Control\Print /v RpcAuthnLevelPrivacyEnabled /t REG_dword /d 0 /f
echo Đã hoàn tất...!
pause & goto :main

:ResetOnedrive
set Onedrive1=%localappdata%\Microsoft\OneDrive\onedrive.exe
set Onedrive2=C:\Program Files\Microsoft OneDrive\onedrive.exe
set Onedrive3=C:\Program Files (x86)\Microsoft OneDrive\onedrive.exe
start "OneDrive" "%Onedrive1%" /reset
start "OneDrive" "%Onedrive2%" /reset
start "OneDrive" "%Onedrive3%" /reset
echo Đã hoàn tất...!
pause & goto :main

:ActiveWindows
cd ActivateAIOToolsv3.1.3\BIN\ActWin10Digital\ActWin10All\
start temp.cmd
cd ..\..\..\..
goto :main

:SetPasswordDevice
net user admin HH@@2016
echo Đã hoàn tất...!
pause & goto :main

:Azure
start "" "ms-settings:workplace"
goto :main

:FontHH
start FontHH
goto :main

:Timezone
TZUTIL /S "SE Asia Standard Time"
echo Đã hoàn tất...!
pause & goto :main

:KeyVNC
echo Z55CY-MJYV9-GFGVJ-SDGB6-L2PDA|clip
echo Đã hoàn tất...!
echo Key VNC đã được copy vào Clipboard
pause & goto :main

:MergeRegedit
for %%f in (*.reg) do regedit /s "%%f"
echo Đã hoàn tất...!
pause & goto :main

:InstallAppsFast
for %%a in (*.exe) do start "" "%%~fa"
start AppFast\InstallAppsFast.cmd
goto :main

:RenameThisPC
set /p Rename="Nhập tên thiết bị: "
echo.
wmic computersystem where name="%computername%" call rename name="%Rename%"
echo Tên Thiết bị đổi thành %Rename%
pause & goto :main

:ResetPassword
start ResetPassword.txt
goto :main

:AioTools
start ActivateAIOToolsv3.1.3\ActivateAIOToolsv3.1.3bySavio.cmd
goto :main

:OptimizeWindows
start OptimizeWindows.vbs
goto :main

:GoogleDrive
start https://drive.google.com/drive/folders/1byWx3tlr2ql2mqJnHdAiZkAGoemQMNdp?usp=sharing
goto :main

:BackupDriver
start BackupDriver.cmd
goto :main

:Exit
exit