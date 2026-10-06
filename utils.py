import base64
import tkinter as tk
from tkinter import filedialog
from pathlib import Path

def encode_image(image_path: str) -> str:
    """Đọc và mã hóa file ảnh sang base64 string."""
    with open(image_path, "rb") as image_file:
        return base64.b64encode(image_file.read()).decode("utf-8")

def read_text_file(file_path: Path) -> str:
    """Đọc nội dung text từ file."""
    with open(file_path, "r", encoding="utf-8") as file:
        return file.read().strip()

def select_folder(title: str = "Chọn thư mục") -> str:
    """Mở hộp thoại chọn thư mục bằng Tkinter."""
    root = tk.Tk()
    root.withdraw()
    root.attributes('-topmost', True)
    folder = filedialog.askdirectory(title=title)
    root.destroy()
    return folder

def select_file(title: str = "Chọn file", filetypes: list = None) -> str:
    """Mở hộp thoại chọn file cụ thể bằng Tkinter."""
    if filetypes is None:
        filetypes = [("Excel files", "*.xlsx *.xls")]

    root = tk.Tk()
    root.withdraw()
    root.attributes('-topmost', True) # Ép cửa sổ luôn nổi lên trên cùng
    file_path = filedialog.askopenfilename(title=title, filetypes=filetypes)
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
            # Kiểm tra tương thích phiên bản
            for k, v in data.items():
                if not isinstance(v, dict):
                    print("[!] Cấu trúc lịch sử cũ không tương thích. Hệ thống sẽ học lại từ đầu...")
                    return {}
            return data
    except Exception:
        return {}

def save_history(data: dict, history_file: str):
    """Ghi đè dữ liệu mới vào file lịch sử JSON"""
    with open(history_file, 'w', encoding='utf-8') as f:
        json.dump(data, f, ensure_ascii=False, indent=4)

def save_file(title: str = "Lưu file", default_name: str = "", filetypes: list = None) -> str:
    """Mở hộp thoại Lưu file (Save As) bằng Tkinter."""
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
