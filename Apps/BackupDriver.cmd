@echo off
CHCP 1258 >nul 2>&1
CHCP 65001 >nul 2>&1
@REM mode con: cols=50 lines=20
color f0
Title SETUP TOOLS
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

echo.
::===========================================================================
cd C:\
mkdir Soft
cd Soft
setlocal enabledelayedexpansion
for /f "tokens=2 delims==" %%a in ('wmic csproduct get name /value ^| find "Name"') do (
    set "computerName=%%a"
)
set "computerName=!computerName:~1,-1!"
set "computerName=!computerName: =_!"

for /f "tokens=1-5 delims=/ " %%a in ('date /t') do (
    set day=%%a
    set month=%%b
    set year=%%c)
set formattedDate=!day!.!month!.!year!

mkdir "BackupDriver_!computerName!_!formattedDate!"
dism /online /export-driver /destination:C:\Soft\BackupDriver_!computerName!_!formattedDate!
start C:\Soft\BackupDriver_!computerName!_!formattedDate!
echo.
echo ===========================================================================
echo.
echo Toàn bộ Driver đã được sao lưu vào ổ đĩa C:\Soft\BackupDriver_!computerName!_!formattedDate!
echo *Lưu ý: Phải copy Soft vào ổ C rồi mới chạy tools này
pause & exit