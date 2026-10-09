@echo off
echo ======================================================
echo    BAT DAU DONG GOI UNG DUNG SO-HOA-VBCC THANH EXE
echo ======================================================
echo.

echo [Buoc 1/3] Cai dat PyInstaller qua uv...
call uv add --dev pyinstaller
if %errorlevel% neq 0 (
    echo [!] Loi khi cai dat PyInstaller qua uv.
    pause
    exit /b %errorlevel%
)
echo.

echo [Buoc 2/3] Dang dong goi bang PyInstaller...
call uv run pyinstaller --noconfirm --onedir --console --name "so-hoa-vbcc" --collect-all pillow_heif --hidden-import openpyxl --hidden-import thefuzz main.py
if %errorlevel% neq 0 (
    echo [!] Qua trinh dong goi gap loi.
    pause
    exit /b %errorlevel%
)
echo.

echo [Buoc 3/3] Dong bo thu muc data va file .env...
if not exist "dist\so-hoa-vbcc\data" (
    xcopy /E /I /Y "data" "dist\so-hoa-vbcc\data" > nul
) else (
    xcopy /E /Y "data" "dist\so-hoa-vbcc\data" > nul
)

if exist ".env" (
    copy /Y ".env" "dist\so-hoa-vbcc\.env" > nul
)

echo ======================================================
echo    DONG GOI THANH CONG!
echo ======================================================
echo  Thu muc ung dung: dist\so-hoa-vbcc\
echo  File thuc thi:    dist\so-hoa-vbcc\so-hoa-vbcc.exe
echo ======================================================
pause
