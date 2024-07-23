@echo off
CHCP 1258 >nul 2>&1
CHCP 65001 >nul 2>&1
Title Install Apps

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

mode con: cols=123 lines=35
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
@echo.

:begin
::===========================================================================
cd ..
cd ..
robocopy Soft C:\Soft /E
cd Soft
cd Apps
robocopy Shortcut C:\Users\Public\Desktop /E

net user admin HH@@2016
wmic UserAccount where Name='admin' set PasswordExpires=False

start "" "ms-settings:workplace"
start FontHH
TZUTIL /S "SE Asia Standard Time"
echo Z55CY-MJYV9-GFGVJ-SDGB6-L2PDA|clip
@REM start ActivateAIOToolsv3.1.3\BIN\ActWin10Digital\ActWin10All\temp.cmd
start ActivateAIOToolsv3.1.3\BIN\ActWinOfficeOnline\ActWinOfficeOnline.cmd
for %%f in (*.reg) do regedit /s "%%f"

echo ===========================================================================
echo.
set /p Rename="Nhập tên thiết bị: "
echo.
wmic computersystem where name="%computername%" call rename name="%Rename%"
echo Computer name changed to %Rename%
pause & exit