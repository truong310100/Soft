@echo off
CHCP 1258 >nul 2>&1
CHCP 65001 >nul 2>&1
Title CHUYỂN ĐỔI PHIÊN BẢN WINDOWS

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

mode con: cols=70 lines=10
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
::===========================================================================
echo.
echo   _____________________________________________________________
echo  ^|  1. Chuyển đến Microsoft Windows Pro         = Nhấn phím 1  ^|
echo  ^|  2. Chuyển đến Microsoft Windows Enterprise  = Nhấn phím 2  ^|
echo  ^|_____________________________________________________________^|
echo.
choice /N /C 12 /M "* Nhập Lựa Chọn Của Bạn: "

if ERRORLEVEL 2 goto :Enterprise
if ERRORLEVEL 1 goto :Pro

:Pro
echo Chuyển đến Microsoft Windows Pro
sc config LicenseManager start= auto & net start LicenseManager
sc config wuauserv start= auto & net start wuauserv
changepk.exe /productkey VK7JG-NPHTM-C97JM-9MPGT-3V66T
pause & exit

:Enterprise
echo Chuyển đến Microsoft Windows Enterprise
sc config LicenseManager start= auto & net start LicenseManager
sc config wuauserv start= auto & net start wuauserv
changepk.exe /productkey BNJ87-JYH2B-DGXWM-JJVMR-Q9MPF
pause & exit
