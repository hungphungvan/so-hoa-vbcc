#!/bin/bash

# Di chuyển Terminal vào đúng thư mục chứa file .command này
cd "$(dirname "$0")"

# Dọn dẹp màn hình Terminal cho gọn gàng
clear

echo "Đang khởi động Hệ Thống Số Hóa VBCC..."

# Ưu tiên chạy bằng môi trường ảo .venv (nếu có)
if [ -f ".venv/bin/python" ]; then
    .venv/bin/python main.py
else
    # Dự phòng chạy bằng lệnh uv nếu không gọi được python trực tiếp
    uv run main.py
fi
