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

def is_empty(val) -> bool:
    """Kiểm tra xem một ô có trống không (chấp nhận STT không phải là số)."""
    if val is None:
        return True
    if str(val).strip() == "":
        return True
    return False

def trim_ghost_rows(ws) -> int:
    """Tìm và chặt bỏ toàn bộ các dòng rác (trống hoàn toàn) ở cuối bảng tính."""
    max_row = ws.max_row
    max_col = ws.max_column
    real_max = 1

    # Quét từ dưới lên trên để tìm dòng thực sự có chứa dữ liệu
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

def find_data_start_row(ws, max_scan=15) -> int:
    """
    Quét các dòng đầu tiên để tự động tìm dòng tiêu đề của bảng dữ liệu.
    Dữ liệu sẽ bắt đầu ở dòng ngay bên dưới dòng tiêu đề đó.
    """
    keywords = ['stt', 'số thứ tự', 'họ và tên', 'họ tên', 'số hiệu bằng']

    # Quét 15 dòng đầu tiên
    for row_idx in range(1, min(max_scan, ws.max_row) + 1):
        # Lấy giá trị của 20 cột đầu tiên đưa về chữ thường để so sánh
        row_vals = [str(ws.cell(row=row_idx, column=c).value).strip().lower()
                    for c in range(1, 20) if ws.cell(row=row_idx, column=c).value]

        # Nếu phát hiện thấy từ khóa của dòng tiêu đề
        for val in row_vals:
            if any(kw == val for kw in keywords):
                return row_idx + 1  # Trả về dòng ngay dưới dòng tiêu đề

    return 2 # Trả về 2 nếu không tìm thấy (giữ nguyên logic cũ làm fallback)
