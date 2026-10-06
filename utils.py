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
    return folder

def select_file(title: str = "Chọn file", filetypes: list = None) -> str:
    """Mở hộp thoại chọn file cụ thể bằng Tkinter."""
    if filetypes is None:
        filetypes = [("Excel files", "*.xlsx *.xls")]

    root = tk.Tk()
    root.withdraw()
    root.attributes('-topmost', True) # Ép cửa sổ luôn nổi lên trên cùng
    file_path = filedialog.askopenfilename(title=title, filetypes=filetypes)

    # Hủy root sau khi chọn xong để tránh treo trên macOS
    root.destroy()

    return file_path
