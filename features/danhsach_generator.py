from datetime import UTC, datetime
from pathlib import Path

import openpyxl

from utils import find_data_start_row, select_folder


def generate_danhsach(template_danhsach_path: str):
    print("\n" + "="*60)
    print(" BẮT ĐẦU TỔNG HỢP DANH SÁCH TÀI LIỆU SỐ HÓA")
    print("="*60)

    print("[+] Đang mở cửa sổ chọn thư mục... (Kiểm tra thanh taskbar nếu không thấy)")

    output_dir = select_folder("Chọn thư mục chứa các file Excel của trường")

    if not output_dir:
        print("[-] Bạn đã hủy chọn thư mục. Kết thúc chương trình.")
        return

    print(f"[+] Thư mục đang xử lý: {output_dir}")

    nam_chung = input("\n[?] Nhập Năm cấp bằng (VD: 2008): ").strip()

    tinh_trang = "Nguyên vẹn"
    ngay_lap_so = datetime.now(tz=UTC).astimezone().strftime("%d/%m/%Y")

    # BƯỚC MỞ TEMPLATE: Ép buộc phải có sheet 'Data'
    try:
        wb = openpyxl.load_workbook(template_danhsach_path)
        if 'Data' not in wb.sheetnames:
            print("[-] LỖI: File template không có sheet 'Data'. Vui lòng kiểm tra lại file mẫu!")
            return

        ws = wb['Data']

    except Exception as e:  # noqa: BLE001
        print(f"[-] LỖI: Không thể mở file {template_danhsach_path}. Chi tiết: {e}")
        return

    # Dòng 1 là tên bảng, dòng 2 là tiêu đề cột -> Bắt đầu ghi từ dòng 3
    row_idx = 3

    target_path = Path(output_dir)
    excel_files = sorted([f for f in target_path.iterdir()
                   if f.is_file() and f.suffix in ['.xlsx', '.xls']
                   and not f.name.startswith('~')
                   and not f.name.startswith('DS_Tai_Lieu_So_Hoa')])

    if not excel_files:
        print(f"[-] Không tìm thấy file Excel trường nào trong thư mục {output_dir}")
        return

    print(f"\n[+] Đã tìm thấy {len(excel_files)} file. Bắt đầu xử lý tự động...\n")

    stt_sogoc_counter = 1
    default_kws = ['stt', 'số hiệu bằng', 'họ và tên', 'hội đồng thi', 'nơi sinh', 'giới tính', 'dân tộc']

    for file_path in excel_files:
        file_name = file_path.name

        ten_truong_da_hoc = ""
        ma_don_vi = ""
        ten_don_vi = ""

        try:
            wb_school = openpyxl.load_workbook(file_path, data_only=True)
            target_sheet = None
            start_row = None

            for sheet_name in wb_school.sheetnames:
                ws_temp = wb_school[sheet_name]
                found_row = find_data_start_row(
                    ws_temp,
                    config_file="data/tu_khoa_cleaner.txt",
                    default_keywords=default_kws,
                    default_start_row=None
                )
                if found_row is not None:
                    target_sheet = ws_temp
                    start_row = found_row
                    break

            if target_sheet is None or start_row is None:
                print(f"  [!] BỎ QUA {file_name}: Không tìm thấy sheet dữ liệu chuẩn.")
                wb_school.close()
                continue

            COL_TEN_TRUONG = 14
            COL_MA_DV = 27
            COL_TEN_DV = 28

            for r in range(start_row, target_sheet.max_row + 1):
                val_stt = target_sheet.cell(row=r, column=1).value

                if isinstance(val_stt, (int, float)) or (isinstance(val_stt, str) and val_stt.strip().isdigit()):

                    cur_truong = str(target_sheet.cell(row=r, column=COL_TEN_TRUONG).value or "").strip()
                    cur_ma = str(target_sheet.cell(row=r, column=COL_MA_DV).value or "").strip()
                    cur_ten_dv = str(target_sheet.cell(row=r, column=COL_TEN_DV).value or "").strip()

                    if not ten_truong_da_hoc and cur_truong and cur_truong.lower() != "none" and not cur_truong.isdigit():
                        ten_truong_da_hoc = cur_truong

                    if not ma_don_vi and cur_ma and cur_ma.lower() != "none":
                        ma_don_vi = cur_ma

                    if not ten_don_vi and cur_ten_dv and cur_ten_dv.lower() != "none" and not cur_ten_dv.isdigit():
                        ten_don_vi = cur_ten_dv

                    if ten_truong_da_hoc and ten_don_vi:
                        break

            wb_school.close()

        except Exception as e:  # noqa: BLE001
            print(f"  [!] Lỗi đọc dữ liệu từ file {file_name}: {e}")
            continue

        if not ten_truong_da_hoc or ten_truong_da_hoc.lower() == "none" or ten_truong_da_hoc.isdigit():
            ten_truong_da_hoc = file_name.rsplit('_', 1)[0].strip()

        if not ten_don_vi or ten_don_vi.lower() == "none" or ten_don_vi.isdigit():
            ten_don_vi = ten_truong_da_hoc

        if ma_don_vi.lower() == "none":
            ma_don_vi = ""

        stt_so_str = f"{stt_sogoc_counter:02d}"
        ma_so_goc = f"{nam_chung}_VP_{stt_so_str}"
        ten_tai_lieu = f"Sổ gốc văn bằng năm {nam_chung} - {ten_truong_da_hoc}"

        print("-" * 50)
        print(f"File: {file_name}")
        print(f"Trường: {ten_truong_da_hoc} | Đơn vị: {ten_don_vi} (Mã: {ma_don_vi})")
        print(f"-> Gán mã sổ: {ma_so_goc}")

        ws.cell(row=row_idx, column=1, value=stt_sogoc_counter)
        ws.cell(row=row_idx, column=2, value=ma_don_vi)
        ws.cell(row=row_idx, column=3, value=ten_don_vi)
        ws.cell(row=row_idx, column=4, value=ma_so_goc)
        ws.cell(row=row_idx, column=5, value=ten_tai_lieu)
        ws.cell(row=row_idx, column=6, value=nam_chung)
        ws.cell(row=row_idx, column=7, value=nam_chung)
        ws.cell(row=row_idx, column=8, value=ngay_lap_so)
        ws.cell(row=row_idx, column=9, value="")
        ws.cell(row=row_idx, column=10, value=tinh_trang)
        ws.cell(row=row_idx, column=11, value="")

        row_idx += 1
        stt_sogoc_counter += 1

    output_danhsach = target_path / f"DS_Tai_Lieu_So_Hoa_{nam_chung}.xlsx"
    wb.save(output_danhsach)

    print("\n" + "="*60)
    print(f"[=] HOÀN TẤT! Đã tổng hợp xong {stt_sogoc_counter - 1} trường.")
    print(f"[=] File danh sách được lưu tại: {output_danhsach}")
