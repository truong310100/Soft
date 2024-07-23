@echo off
CHCP 1258 >nul 2>&1
CHCP 65001 >nul 2>&1
mode con: cols=50 lines=20
color 3F
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
for /F "tokens=1,2 delims=#" %%a in ('"prompt #$H#$E# & echo on & for %%b in (1) do rem"') do (set "DEL=%%a")
for /f "tokens=2 delims==" %%a in ('wmic os get caption /value') do set "os_name=%%a"
set "text = Thiết bị đang sử dụng phiên bản:"
if /I "%os_name%" == "Microsoft Windows 10 Pro" (
  call :ColorText 2F " Phien Ban Windows  =  %os_name%"
) else if /I "%os_name%" == "Microsoft Windows 11 Pro" (
  call :ColorText 2F " Phien Ban Windows  =  %os_name%"
) else (
  call :ColorText 4F " Phien Ban Windows  =  %os_name%"
)

goto :Beginoffile
:ColorText
echo off
<nul set /p ".=%DEL%" > "%~2"
findstr /v /a:%1 /R "^$" "%~2" nul
del "%~2" > nul 2>&1
goto :eof  
:Beginoffile
::===========================================================================
echo.
echo   ________________________________________
echo  ^|  1. Chuyển đổi Windows  = Nhấn phím 1  ^|
echo  ^|  2. Cài đặt phần mềm    = Nhấn phím 2  ^|
echo  ^|  3. Hỗ Trợ Người Dùng   = Nhấn phím 3  ^|
echo  ^|  4. Chữ ký Outlook      = Nhấn phím 4  ^|
echo  ^|  5. Kiểm tra thiết bị   = Nhấn phím 5  ^|
echo  ^|________________________________________^|
echo  ^|        Thoát     = Nhấn phím 6         ^|
echo  ^|        Contact   = Nhấn phím 7         ^|
echo  ^|                                        ^|
echo  ^|__Phiên bản v23.07.24_10:21_byTruongNL__^|
echo.

choice /N /C 1234567 /M "* Nhập Lựa Chọn Của Bạn: "

if ERRORLEVEL 7 goto :Contact 7 
if ERRORLEVEL 6 goto :Exit 6 
if ERRORLEVEL 5 goto :CheckingDevice 5
if ERRORLEVEL 4 goto :Signature 4
if ERRORLEVEL 3 goto :SupportUser 3
if ERRORLEVEL 2 goto :Install 2
if ERRORLEVEL 1 goto :ConvertWin 1

:ConvertWin
start Apps\ConvertWin.cmd
exit

:Install
echo CÀI ĐẶT CẤU HÌNH VÀ PHẦN MỀM
start Apps\InstallAppsFast.cmd
start Apps\Config.cmd
exit

:SupportUser
echo Support User
start Apps\5.SupportUser.cmd
goto :main

:Signature
echo Signature HHH
cd Apps\signatureHHH
start run.exe
cd ..\..
goto :main

:CheckingDevice
echo Checking Device
cd Apps\CheckingDevice
start run.exe
cd ..\..
goto :main

:Exit
exit

:Contact
start Apps\Contact.vbs
goto :main