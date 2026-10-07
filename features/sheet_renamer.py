from pathlib import Path

import openpyxl

from utils import select_folder


def run_sheet_renamer():
    print("\n" + "="*60)
    print(" BẮT ĐẦU ĐỔI TÊN SHEET THÀNH 'Data'")
    print("="*60)

    print("[+] Đang mở cửa sổ chọn thư mục... (Kiểm tra taskbar nếu không thấy)")
    folder_path = select_folder("Chọn thư mục chứa các file Excel cần đổi tên sheet")

    if not folder_path:
        print("[-] Đã hủy chọn thư mục. Kết thúc thao tác.")
        return

    target_dir = Path(folder_path)
    # Lọc ra các file excel, bỏ qua các file tạm đang mở (bắt đầu bằng ~)
    excel_files = [f for f in target_dir.iterdir() if f.is_file() and f.suffix in ['.xlsx', '.xls'] and not f.name.startswith('~')]

    if not excel_files:
        print(f"[-] Không tìm thấy file Excel nào trong thư mục {target_dir.name}")
        return

    print(f"\n[+] Đã tìm thấy {len(excel_files)} file. Bắt đầu xử lý...\n")

    count = 0
    for file_path in excel_files:
        try:
            wb = openpyxl.load_workbook(file_path)
            # Lấy sheet đầu tiên đang active
            sheet = wb.active

            # FIX 1: Linter bắt buộc phải kiểm tra None trước khi gọi thuộc tính
            if sheet is None:
                print(f"  [!] {file_path.name}: File không có sheet nào hợp lệ.")
                continue

            if sheet.title != "Data":
                old_name = sheet.title
                sheet.title = "Data"
                wb.save(file_path)
                print(f"  -> {file_path.name}: '{old_name}' => 'Data'")
                count += 1
            else:
                print(f"  -> {file_path.name}: Đã là 'Data', bỏ qua.")

        # FIX 2: Cắm cờ báo cho linter biết mình cố tình bắt lỗi để không sập vòng lặp
        except Exception as e:  # noqa: BLE001
            print(f"  [!] LỖI khi xử lý {file_path.name}: {e}")

    print("\n" + "="*60)
    print(f"[=] HOÀN TẤT! Đã đổi tên sheet thành công cho {count}/{len(excel_files)} file.")
