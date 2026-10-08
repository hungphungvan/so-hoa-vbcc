import base64
import io
import json
import platform
import re
import subprocess
import tkinter as tk
from pathlib import Path
from tkinter import filedialog

from PIL import Image
from pillow_heif import register_heif_opener

# Đăng ký bộ đọc HEIC cho hệ thống (chỉ cần gọi 1 lần khi import utils)
register_heif_opener()

def natural_sort_key(file_path: str | Path):
    """
    Tách tên file thành các phần chữ và số để sắp xếp tự nhiên (Natural Sort).
    Đảm bảo thứ tự 1, 2, ..., 9, 10 cũng như 01, 02, ..., 10.
    """
    name = Path(file_path).name
    tokens = [int(text) if text.isdigit() else text.lower() for text in re.split(r'(\d+)', name)]
    return (tokens, name)


def encode_image(image_path: str) -> str:
    """Đọc ảnh (tự động convert HEIC sang JPEG trong RAM) và chuyển sang base64."""
    path = Path(image_path)

    # Nếu là file HEIC/HEIF từ iPhone
    if path.suffix.lower() in [".heic", ".heif"]:
        img = Image.open(path)
        if img.mode != "RGB":
            img = img.convert("RGB")

        buffer = io.BytesIO()
        img.save(buffer, format="JPEG")
        return base64.b64encode(buffer.getvalue()).decode("utf-8")

    # Nếu là ảnh bình thường thì đọc trực tiếp
    else:
        with open(path, "rb") as image_file:
            return base64.b64encode(image_file.read()).decode("utf-8")

def read_text_file(file_path: Path) -> str:
    """Đọc nội dung text từ file."""
    with open(file_path, "r", encoding="utf-8") as file:
        return file.read().strip()

def select_folder(title: str = "Chọn thư mục") -> str:
    """Mở hộp thoại chọn thư mục (Hỗ trợ Native macOS, ép focus)."""
    if platform.system() == "Darwin":
        script = f"""
        tell application (path to frontmost application as text)
            activate
            set f to choose folder with prompt "{title}"
        end tell
        return POSIX path of f
        """
        result = subprocess.run(['osascript', '-e', script], capture_output=True, text=True, check=False)
        return result.stdout.strip() if result.returncode == 0 else ""
    else:
        root = tk.Tk()
        root.withdraw()
        root.attributes('-topmost', True)
        folder = filedialog.askdirectory(title=title)
        root.destroy()
        return folder

def select_file(title: str = "Chọn file", filetypes: list | None = None) -> str:
    """Mở hộp thoại chọn file cụ thể (Hỗ trợ Native macOS, ép focus)."""
    if platform.system() == "Darwin":
        script = f"""
        tell application (path to frontmost application as text)
            activate
            set f to choose file with prompt "{title}"
        end tell
        return POSIX path of f
        """
        result = subprocess.run(['osascript', '-e', script], capture_output=True, text=True, check=False)
        return result.stdout.strip() if result.returncode == 0 else ""
    else:
        if filetypes is None:
            filetypes = [("Excel files", "*.xlsx *.xls")]
        root = tk.Tk()
        root.withdraw()
        root.attributes('-topmost', True)
        file_path = filedialog.askopenfilename(title=title, filetypes=filetypes)
        root.destroy()
        return file_path

def save_file(title: str = "Lưu file", default_name: str = "", filetypes: list | None = None) -> str:
    """Mở hộp thoại Lưu file - Save As (Hỗ trợ Native macOS, ép focus)."""
    if platform.system() == "Darwin":
        script = f"""
        tell application (path to frontmost application as text)
            activate
            set f to choose file name with prompt "{title}" default name "{default_name}"
        end tell
        return POSIX path of f
        """
        result = subprocess.run(['osascript', '-e', script], capture_output=True, text=True, check=False)
        return result.stdout.strip() if result.returncode == 0 else ""
    else:
        if filetypes is None:
            filetypes = [("Excel files", "*.xlsx"), ("All files", "*.*")]
        root = tk.Tk()
        root.withdraw()
        root.attributes('-topmost', True)
        file_path = filedialog.asksaveasfilename(
            title=title,
            initialfile=default_name,
            defaultextension=".xlsx",
            filetypes=filetypes
        )
        root.destroy()
        return file_path

def open_file_in_os(file_path: str | Path) -> bool:
    """Mở file bằng ứng dụng mặc định của hệ thống (hỗ trợ macOS, Windows, Linux)."""
    target = Path(file_path).resolve()
    if not target.exists():
        print(f"[-] File không tồn tại: {target}")
        return False

    try:
        sys_name = platform.system()
        if sys_name == "Darwin":
            subprocess.run(["open", str(target)], check=False)
        elif sys_name == "Windows":
            import os
            os.startfile(str(target))
        else:
            subprocess.run(["xdg-open", str(target)], check=False)
        return True
    except Exception as e:  # noqa: BLE001
        print(f"[-] Không thể mở file tự động: {e}")
        return False

def load_history(history_file: str) -> dict:
    """Đọc và kiểm tra tính hợp lệ của file lịch sử chuẩn hóa (JSON)"""
    path = Path(history_file)
    if not path.exists():
        return {}
    try:
        with open(path, 'r', encoding='utf-8') as f:
            data = json.load(f)
            for v in data.values():
                if not isinstance(v, dict):
                    print("[!] Cấu trúc lịch sử cũ không tương thích. Hệ thống sẽ học lại từ đầu...")
                    return {}
            return data
    except Exception:  # noqa: BLE001
        return {}

def save_history(data: dict, history_file: str):
    """Ghi đè dữ liệu mới vào file lịch sử JSON"""
    with open(history_file, 'w', encoding='utf-8') as f:
        json.dump(data, f, ensure_ascii=False, indent=4)

def is_empty(val) -> bool:
    """Kiểm tra xem một ô có trống không (chấp nhận STT không phải là số)."""
    return val is None or str(val).strip() == ""

def trim_ghost_rows(ws) -> int:
    """
    Tìm và chặt bỏ toàn bộ các dòng rác ở cuối bảng tính siêu tốc.
    Xử lý triệt để bẫy ô rác đơn độc lơ lửng ở đáy trang tính (như kẹt ô ở dòng 1.048.576).
    """
    if not hasattr(ws, '_cells'):
        return 0

    # Gom nhóm và đếm số ô có dữ liệu thật trên từng dòng
    row_counts = {}
    for (r, _), cell in ws._cells.items():
        if cell.value is not None and str(cell.value).strip() != "":
            row_counts[r] = row_counts.get(r, 0) + 1

    if not row_counts:
        return 0

    # Lấy danh sách các dòng có dữ liệu đã sắp xếp tăng dần
    sorted_rows = sorted(row_counts.keys())

    # Quét ngược từ đáy lên để gọt sạch các ô rác cô lập / lơ lửng ở đáy
    while sorted_rows:
        last_row = sorted_rows[-1]

        # Nếu chỉ còn 1 dòng duy nhất thì giữ lại
        if len(sorted_rows) == 1:
            break

        prev_row = sorted_rows[-2]
        gap = last_row - prev_row

        # BẮT BẪY Ô RÁC ĐÁY:
        # Nếu dòng đáy có ít ô (<= 2 ô) VÀ cách dòng dữ liệu bên trên từ 2 dòng trống trở lên (gap > 2)
        # -> Đây 100% là ô rác vô tình bị rơi xuống đáy trang tính
        if row_counts[last_row] <= 2 and gap > 2:
            sorted_rows.pop()
        else:
            break

    real_max = sorted_rows[-1]
    deleted_count = 0
    max_row = ws.max_row

    if max_row > real_max:
        deleted_count = max_row - real_max
        ws.delete_rows(real_max + 1, deleted_count)

    return deleted_count

def get_header_keywords(config_file, default_keywords) -> list:
    """Đọc danh sách từ khóa nhận diện tiêu đề từ file cấu hình riêng của từng module."""
    path = Path(config_file)
    if path.exists():
        try:
            with open(path, 'r', encoding='utf-8') as f:
                file_kws = [line.strip().lower() for line in f if line.strip()]
                if file_kws:
                    return file_kws
        except (OSError, UnicodeDecodeError) as e:
            print(f"[!] Lỗi đọc file {config_file}: {e}. Sẽ dùng từ khóa mặc định.")

    return default_keywords

def find_data_start_row(
    ws,
    max_scan: int = 15,
    config_file: str = "data/tu_khoa_nhan_dien.txt",
    default_keywords: list | None = None,
    default_start_row: int | None = 3
) -> int | None:
    """
    Quét các dòng đầu tiên để tự động tìm dòng tiêu đề của bảng dữ liệu.
    Sử dụng cơ chế đếm từ khóa (cố định yêu cầu >= 3 từ khóa) để chống nhận diện nhầm.
    """
    if default_keywords is None:
        default_keywords = ['stt', 'số thứ tự', 'họ và tên', 'tt', 'sbd']

    keywords = get_header_keywords(config_file, default_keywords)

    for row_idx in range(1, min(max_scan, ws.max_row) + 1):
        # Mở rộng quét lên 30 cột để bắt hết form ngang dài
        row_vals = [str(ws.cell(row=row_idx, column=c).value).strip().lower()
                    for c in range(1, 31) if ws.cell(row=row_idx, column=c).value]

        # Đếm số lượng từ khóa xuất hiện trên cùng một dòng
        matched_count = sum(1 for kw in keywords if kw in row_vals)

        # CHỐT CỨNG: Phải đạt ít nhất 3 từ khóa mới xác định là dòng tiêu đề
        if matched_count >= 3:
            return row_idx + 1

    return default_start_row
