@echo off
CHCP 1258 >nul 2>&1
CHCP 65001 >nul 2>&1


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
    echo  Chay CMD Voi Quyen Quan tri - Run as Administrator...
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
echo Đang cài đặt, chương trình hoàn thành trong vài giây...
echo.
start C:\Soft\Apps\UniKeyNT.exe
start C:\Soft\Apps\OfficeSetup.exe
@REM start C:\Soft\Apps\MSTeamsSetup_c_l_.exe
@REM start C:\Soft\Apps\MSTeams-x64.msix
start C:\Soft\Apps\AnyDesk.exe --install "C:\Program Files (x86)\AnyDesk" --start-with-win --create-desktop-icon
start C:\Soft\Apps\7z2301-x64.exe /S /D="C:\Program Files\7-Zip"
start C:\Soft\Apps\FoxitReader10.exe /VERYSILENT /SUPPRESSMSGBOXES /SP-
start C:\Soft\Apps\UltraViewer_setup_6.6_vi.exe /VERYSILENT /SUPPRESSMSGBOXES
start C:\Soft\Apps\SetupVNCv4.5.3.exe /VERYSILENT /SUPPRESSMSGBOXES
echo Z55CY-MJYV9-GFGVJ-SDGB6-L2PDA|clip
echo.
echo Key VNC = Z55CY-MJYV9-GFGVJ-SDGB6-L2PDA
echo Đã xong, nhấn phím bất kỳ để thoát...
pause >nul & exit