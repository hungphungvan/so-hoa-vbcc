import base64
import io
import json
import platform
import subprocess
import tkinter as tk
from pathlib import Path
from tkinter import filedialog

from PIL import Image
from pillow_heif import register_heif_opener

# Đăng ký bộ đọc HEIC cho hệ thống (chỉ cần gọi 1 lần khi import utils)
register_heif_opener()

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
    """Tìm và chặt bỏ toàn bộ các dòng rác (trống hoàn toàn) ở cuối bảng tính."""
    max_row = ws.max_row
    max_col = ws.max_column
    real_max = 1

    for r in range(max_row, 0, -1):
        has_data = False
        for c in range(1, max_col + 1):
            if not is_empty(ws.cell(row=r, column=c).value):
                has_data = True
                break

        if has_data:
            real_max = r
            break

    deleted_count = 0
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

def find_data_start_row(ws, max_scan=15, config_file="data/tu_khoa_nhan_dien.txt", default_keywords=None) -> int:
    """Quét các dòng đầu tiên để tự động tìm dòng tiêu đề của bảng dữ liệu."""
    if default_keywords is None:
        default_keywords = ['stt', 'số thứ tự', 'họ và tên', 'tt', 'sbd']

    keywords = get_header_keywords(config_file, default_keywords)

    for row_idx in range(1, min(max_scan, ws.max_row) + 1):
        row_vals = [str(ws.cell(row=row_idx, column=c).value).strip().lower()
                    for c in range(1, 20) if ws.cell(row=row_idx, column=c).value]

        for val in row_vals:
            if any(kw == val for kw in keywords):
                return row_idx + 1

    return 2
