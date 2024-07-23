@echo off
CHCP 1258 >nul 2>&1
CHCP 65001 >nul 2>&1
title TÌM SHORTCUT TRONG FOLDER

@echo off
echo.
set /p path="Nhập đường dẫn cần tìm Shortcut: "
dir /s /b "%path%\*.lnk"
pause
