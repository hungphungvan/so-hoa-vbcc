from pathlib import Path

import openpyxl
from openpyxl.styles import PatternFill

from utils import find_data_start_row, is_empty, select_folder, trim_ghost_rows


def run_data_cleaner():
    print("\n" + "="*60)
    print(" BẮT ĐẦU CHUẨN HOÁ DỮ LIỆU ĐỂ ĐƯA LÊN HỆ THỐNG")
    print("="*60)

    print("[+] Đang mở cửa sổ chọn thư mục... (Kiểm tra taskbar nếu không thấy)")
    folder_path = select_folder("Chọn thư mục chứa các file Excel cần xử lý")

    if not folder_path:
        print("[-] Đã hủy chọn thư mục. Kết thúc thao tác.")
        return

    target_dir = Path(folder_path)
    excel_files = [f for f in target_dir.iterdir() if f.is_file() and f.suffix in ['.xlsx', '.xls'] and not f.name.startswith('~')]

    if not excel_files:
        print(f"[-] Không tìm thấy file Excel nào trong thư mục {target_dir.name}")
        return

    print(f"\n[+] Đã tìm thấy {len(excel_files)} file. Bắt đầu xử lý...\n")

    # Cấu hình màu sắc
    yellow_fill = PatternFill(start_color="FFFFFF00", end_color="FFFFFF00", fill_type="solid") # Vàng cho ô thiếu
    red_fill = PatternFill(start_color="FFFF9999", end_color="FFFF9999", fill_type="solid")    # Đỏ nhạt cho dòng rác

    # 18 cột bắt buộc (Đã bỏ Cột A-STT và Cột E-Họ tên)
    MANDATORY_COLS = {
        "Số hiệu bằng (B)": 2,
        "Tên văn bằng (D)": 4,
        "Ngày sinh (I)": 9,
        "Giới tính (J)": 10,
        "Dân tộc (K)": 11,
        "Quốc tịch (L)": 12,
        "Nơi sinh (M)": 13,
        "Tên trường (N)": 14,
        "Niên khoá (O)": 15,
        "Hội đồng thi (Q)": 17,
        "Địa danh cấp bằng (R)": 18,
        "Ngày cấp bằng (S)": 19,
        "Tên cơ quan (T)": 20,
        "Chức danh (U)": 21,
        "Người ký (V)": 22,
        "Trạng thái số hóa (Y)": 25,
        "Mã đơn vị (AA)": 27,
        "Tên đơn vị (AB)": 28
    }

    files_with_issues = []
    count_processed = 0

    for file_path in excel_files:
        print(f"\n--- Đang xử lý: {file_path.name} ---")
        try:
            wb = openpyxl.load_workbook(file_path)
            sheet = wb.active

            if sheet is None:
                print("  [!] File không có sheet nào hợp lệ.")
                continue

            is_modified = False

            # 1. Chuẩn hóa tên sheet thành Data
            if sheet.title != "Data":
                sheet.title = "Data"
                is_modified = True
                print("  -> Đã đổi tên sheet thành 'Data'.")

            # 2. Cắt tỉa dòng ma (trắng tinh ở cuối file)
            ghost_deleted = trim_ghost_rows(sheet)
            if ghost_deleted > 0:
                is_modified = True
                print(f"  -> Đã gọt sạch {ghost_deleted} 'dòng ma' ở đuôi file.")

            # 3. Định vị dòng dữ liệu
            start_row = find_data_start_row(sheet)

            # 4. Rà soát dữ liệu & Tô màu
            has_issue = False
            red_rows = 0
            yellow_cells = 0

            for r in range(start_row, sheet.max_row + 1):
                # Bỏ qua nếu dòng này hoàn toàn trống (chỉ soi từ cột 1 đến 30)
                if all(is_empty(sheet.cell(row=r, column=c).value) for c in range(1, 30)):
                    continue

                # KIỂM ĐỊNH 1: Dòng rác (Thiếu Họ Tên ở Cột 5) -> Tô ĐỎ NHẠT cả dòng
                if is_empty(sheet.cell(row=r, column=5).value):
                    for c in range(1, 30): # Tô đến cột AC cho gọn mắt
                        sheet.cell(row=r, column=c).fill = red_fill
                    has_issue = True
                    is_modified = True
                    red_rows += 1

                # KIỂM ĐỊNH 2: Dòng hợp lệ (Có Họ tên) -> Soi 18 cột bắt buộc -> Tô VÀNG ô thiếu
                else:
                    for col_idx in MANDATORY_COLS.values():
                        cell = sheet.cell(row=r, column=col_idx)
                        if is_empty(cell.value):
                            cell.fill = yellow_fill
                            has_issue = True
                            is_modified = True
                            yellow_cells += 1

            if has_issue:
                files_with_issues.append(file_path.name)
                print(f"  [!] PHÁT HIỆN LỖI: {red_rows} dòng rác (Tô đỏ) | {yellow_cells} ô thiếu (Tô vàng).")
            else:
                print("  -> OK: Dữ liệu hoàn hảo, không có lỗi.")

            if is_modified:
                wb.save(file_path)
            count_processed += 1

        except Exception as e:  # noqa: BLE001
            print(f"  [!] LỖI khi quét file: {e}")

    print("\n" + "="*60)
    print(f"[=] HOÀN TẤT CHUẨN HOÁ ({count_processed}/{len(excel_files)} file)!")

    # 5. Xuất Bảng phong thần
    if files_with_issues:
        print("\n" + "!"*60)
        print(" [BẢNG PHONG THẦN - CÁC FILE CẦN BẠN MỞ RA SỬA TAY]")
        print(" -> Dòng bôi ĐỎ NHẠT : Kéo chọn và ấn Delete (do nhập rác).")
        print(" -> Ô bôi VÀNG      : Bổ sung dữ liệu còn thiếu.")
        print(" -> Cột STT (A)     : Kéo chuột đánh lại sau khi xóa dòng đỏ.")
        for f in files_with_issues:
            print(f"  - {f}")
        print("!"*60)
    else:
        print("\n[+] Tuyệt vời! Toàn bộ file đều hoàn hảo 100%, sẵn sàng up hệ thống.")
