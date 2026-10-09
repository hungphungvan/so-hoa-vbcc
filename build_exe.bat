@echo off
chcp 65001 > nul
echo ======================================================
echo    BẮT ĐẦU ĐÓNG GÓI ỨNG DỤNG SỐ HÓA VBCC THÀNH FILE EXE
echo ======================================================
echo.

echo [Bước 1/3] Kiểm tra và cài đặt PyInstaller...
uv add --dev pyinstaller
if %errorlevel% neq 0 (
    echo [!] Lỗi khi cài đặt PyInstaller qua uv.
    pause
    exit /b %errorlevel%
)
echo.

echo [Bước 2/3] Đang đóng gói ứng dụng (PyInstaller --onedir)...
uv run pyinstaller --noconfirm --onedir --console ^
    --name "so-hoa-vbcc" ^
    --collect-all pillow_heif ^
    --hidden-import openpyxl ^
    --hidden-import thefuzz ^
    main.py

if %errorlevel% neq 0 (
    echo [!] Quá trình đóng gói gặp lỗi.
    pause
    exit /b %errorlevel%
)
echo.

echo [Bước 3/3] Đồng bộ thư mục data và file .env...
if not exist "dist\so-hoa-vbcc\data" (
    xcopy /E /I /Y "data" "dist\so-hoa-vbcc\data" > nul
) else (
    xcopy /E /Y "data" "dist\so-hoa-vbcc\data" > nul
)

if exist ".env" (
    copy /Y ".env" "dist\so-hoa-vbcc\.env" > nul
)

echo ======================================================
echo    ĐÓNG GÓI THÀNH CÔNG!
echo ======================================================
echo  Thư mục ứng dụng: dist\so-hoa-vbcc\
echo  File thực thi:    dist\so-hoa-vbcc\so-hoa-vbcc.exe
echo.
echo  Ghi chú: Khi mang sang máy khác, bạn chỉ cần nén 
echo  toàn bộ thư mục 'dist\so-hoa-vbcc' thành file .zip.
echo ======================================================
pause
