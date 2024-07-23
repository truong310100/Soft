@echo off
CHCP 1258 >nul 2>&1
CHCP 65001 >nul 2>&1
Title Fast Copy

set /p A="Đường dẫn đầu: "
dir /a-d /s /b "%A%" | find /c "\"
for %%i in ("%A%") do set "last_folder=%%~nxi"

set /p B="Đường dẫn cuối: "
if not exist "%A%" (
  echo Path does not exist.
  pause
)

robocopy "%A%" "%B%\%last_folder%" /E
echo Đã xong...Nhấn phím bất kỳ để thoát
pause >nul