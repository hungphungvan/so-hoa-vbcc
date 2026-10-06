import os
from pathlib import Path
import openpyxl
from utils import select_folder, is_empty, trim_ghost_rows, find_data_start_row

def parse_selection(selection_str: str) -> list:
    """Chuyển chuỗi nhập (vd: '5, 7, 10-12') thành danh sách các dòng cần xóa, xếp ngược từ lớn đến bé."""
    if selection_str.lower() == 'all':
        return 'all'

    selected_rows = set()
    parts = selection_str.split(',')
    for part in parts:
        part = part.strip()
        if not part:
            continue
        if '-' in part:
            try:
                start, end = map(int, part.split('-'))
                selected_rows.update(range(start, end + 1))
            except ValueError:
                pass
        else:
            try:
                selected_rows.add(int(part))
            except ValueError:
                pass
    return sorted(list(selected_rows), reverse=True)

def run_data_cleaner():
    print("\n" + "="*60)
    print(" BẮT ĐẦU DỌN DẸP DỮ LIỆU THỪA & CHUẨN HÓA STT")
    print("="*60)

    print("[+] Đang mở cửa sổ chọn thư mục... (Kiểm tra taskbar nếu không thấy)")
    folder_path = select_folder("Chọn thư mục chứa các file Excel cần dọn dẹp")

    if not folder_path:
        print("[-] Đã hủy chọn thư mục. Kết thúc thao tác.")
        return

    target_dir = Path(folder_path)
    excel_files = [f for f in target_dir.iterdir() if f.is_file() and f.suffix in ['.xlsx', '.xls'] and not f.name.startswith('~')]

    if not excel_files:
        print(f"[-] Không tìm thấy file Excel nào trong thư mục {target_dir.name}")
        return

    print(f"\n[+] Đã tìm thấy {len(excel_files)} file. Bắt đầu xử lý...\n")

    columns_to_check = {
        "STT": 1,                 # Cột A
        "Số hiệu bằng": 2,        # Cột B
        "Họ và tên": 5,           # Cột E
        "Ngày sinh": 9,           # Cột I
        "Tên trường": 14,         # Cột N
        "Hội đồng thi": 17,       # Cột Q
        "Mã đơn vị": 27,          # Cột AA
        "Tên đơn vị": 28          # Cột AB
    }

    for file_path in excel_files:
        print(f"\n--- Đang xử lý: {file_path.name} ---")
        try:
            wb = openpyxl.load_workbook(file_path)
            ws = wb.active
            is_modified = False

            # ĐỊNH VỊ DÒNG BẮT ĐẦU DỮ LIỆU (tái sử dụng từ utils)
            start_row = find_data_start_row(ws)
            if start_row != 2:
                print(f"  -> Bảng dữ liệu tự động xác định bắt đầu từ dòng {start_row}.")

            # Bước 1: Tự động "cắt tỉa" dòng rác ở cuối file (tái sử dụng từ utils)
            ghost_deleted = trim_ghost_rows(ws)
            if ghost_deleted > 0:
                print(f"  -> Đã tự động cắt bỏ {ghost_deleted} 'dòng ma' trống ở cuối bảng.")
                is_modified = True

            # Bước 2: Quét lỗi thiếu cột (sử dụng start_row)
            invalid_rows = []
            for row_idx in range(start_row, ws.max_row + 1):
                if all(is_empty(ws.cell(row=row_idx, column=c).value) for c in range(1, 30)):
                    continue

                missing_fields = []
                for field_name, col_idx in columns_to_check.items():
                    val = ws.cell(row=row_idx, column=col_idx).value
                    if is_empty(val):
                        missing_fields.append(field_name)

                if missing_fields:
                    stt = ws.cell(row=row_idx, column=1).value
                    name = ws.cell(row=row_idx, column=5).value
                    identifier = f"STT: {stt if not is_empty(stt) else '[Trống]'} | Tên: {name if not is_empty(name) else '[Trống]'}"

                    invalid_rows.append({
                        "row": row_idx,
                        "identifier": identifier,
                        "missing": missing_fields
                    })

            if invalid_rows:
                print(f"  [!] Phát hiện {len(invalid_rows)} dòng bị thiếu dữ liệu:")
                for item in invalid_rows:
                    missing_str = ", ".join(item["missing"])
                    print(f"      + Dòng {item['row']:<3} | {item['identifier']:<40} (Thiếu: {missing_str})")

                while True:
                    print("\n  Lựa chọn xử lý:")
                    print("  - Nhập các dòng muốn xóa (VD: 5, 7, 10-12)")
                    print("  - Nhập 'all' để xóa TẤT CẢ các dòng lỗi trên")
                    print("  - Nhập '0' hoặc để trống để BỎ QUA không xóa")

                    choice = input("  -> Lựa chọn của bạn: ").strip()

                    if choice in ['0', '']:
                        print("  -> Bỏ qua thao tác xóa dòng lỗi.")
                        break

                    rows_to_delete = []
                    if choice.lower() == 'all':
                        rows_to_delete = sorted([item['row'] for item in invalid_rows], reverse=True)
                    else:
                        parsed = parse_selection(choice)
                        valid_flagged = [item['row'] for item in invalid_rows]
                        rows_to_delete = sorted([r for r in parsed if r in valid_flagged], reverse=True)

                    if not rows_to_delete:
                        print("  [!] Cú pháp không hợp lệ hoặc dòng đã chọn không nằm trong danh sách. Vui lòng thử lại.")
                        continue

                    for r in rows_to_delete:
                        ws.delete_rows(r)

                    print(f"  -> Đã xóa {len(rows_to_delete)} dòng lỗi!")
                    is_modified = True
                    break
            else:
                if not ghost_deleted:
                    print("  -> File sạch sẽ, không có dòng lỗi thiếu cột!")

            # Bước 3: Đánh lại STT (Cột 1) - cũng sử dụng start_row
            stt_counter = 1
            stt_updated = False
            for row_idx in range(start_row, ws.max_row + 1):
                if not all(is_empty(ws.cell(row=row_idx, column=c).value) for c in range(1, 30)):
                    current_stt = ws.cell(row=row_idx, column=1).value
                    if current_stt != stt_counter:
                        ws.cell(row=row_idx, column=1).value = stt_counter
                        stt_updated = True
                    stt_counter += 1

            if stt_updated:
                print(f"  -> Đã tự động đánh lại STT liên tục từ 1 đến {stt_counter - 1}.")
                is_modified = True

            if is_modified:
                wb.save(file_path)
                print("  -> Đã lưu lại file thành công.")

        except Exception as e:
            print(f"  [!] LỖI khi quét file {file_path.name}: {e}")

    print("\n" + "="*60)
    print("[=] HOÀN TẤT DỌN DẸP & CHUẨN HÓA STT!")
