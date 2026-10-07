from datetime import UTC, datetime
from pathlib import Path

import openpyxl

# Import trực tiếp các hàm dùng chung từ utils.py của bạn
from utils import load_history, select_folder


def generate_danhsach(template_danhsach_path: str):
    print("\n" + "="*60)
    print(" BẮT ĐẦU TỔNG HỢP DANH SÁCH TÀI LIỆU SỐ HÓA")
    print("="*60)

    print("[+] Đang mở cửa sổ chọn thư mục... (Kiểm tra thanh taskbar nếu không thấy)")

    # Sử dụng hàm chọn thư mục chuyên nghiệp của bạn
    output_dir = select_folder("Chọn thư mục chứa các file Excel của trường")

    if not output_dir:
        print("[-] Bạn đã hủy chọn thư mục. Kết thúc chương trình.")
        return

    print(f"[+] Thư mục đang xử lý: {output_dir}")

    # 1. Chỉ hỏi Năm đúng 1 lần
    nam_chung = input("\n[?] Nhập Năm cấp bằng (VD: 2008): ").strip()

    # Cố định Tình trạng tài liệu và ngày lập
    tinh_trang = "Nguyên vẹn"
    ngay_lap_so = datetime.now(tz=UTC).astimezone().strftime("%d/%m/%Y")

    # 2. Đọc bộ nhớ lịch sử bằng hàm load_history từ utils
    history_file = "data/lich_su_chuan_hoa.json"
    history_data = load_history(history_file)

    # Xây dựng từ điển để tra cứu Đơn vị dựa trên Tên trường
    school_info_db = {}
    # FIX: Dùng values() thay vì items() nếu không cần key cũ
    for raw_name, info in history_data.items():
        if isinstance(info, dict):
            # Key là tên trường, Value là toàn bộ info
            ten_truong = info.get("ten_truong", raw_name)
            school_info_db[ten_truong] = info

    # 3. Mở file template Danh sách
    try:
        wb = openpyxl.load_workbook(template_danhsach_path)
        # Bắt lỗi an toàn nếu template không có sheet Data
        if 'Data' in wb.sheetnames:
            ws = wb['Data']
        else:
            ws = wb.active
            if ws is not None:
                ws.title = 'Data'

        # Kiểm tra None để linter không phàn nàn
        if ws is None:
            print("[-] LỖI: Template Excel không có bảng tính nào hợp lệ.")
            return

    except Exception as e:  # noqa: BLE001 (FIX: Chặn cảnh báo blind exception)
        print(f"[-] LỖI: Không thể mở file {template_danhsach_path}. Chi tiết: {e}")
        return

    start_row = 2
    row_idx = start_row

    # FIX: Thay thế os.listdir bằng pathlib cho đồng bộ và hiện đại
    target_path = Path(output_dir)
    excel_files = [f.name for f in target_path.iterdir()
                   if f.is_file() and f.suffix in ['.xlsx', '.xls']
                   and not f.name.startswith('~')
                   and not f.name.startswith('DS_Tai_Lieu_So_Hoa')]

    if not excel_files:
        print(f"[-] Không tìm thấy file Excel trường nào trong thư mục {output_dir}")
        return

    print(f"\n[+] Đã tìm thấy {len(excel_files)} file. Bắt đầu duyệt từng trường...\n")

    # 4. Duyệt từng file, lấy thông tin và hỏi STT
    for file_name in excel_files:
        # Cắt lấy Tên trường từ tên file
        base_name = file_name.rsplit('_', 1)[0]
        ten_truong_da_hoc = base_name.strip()

        # Mặc định lấy tên trường làm đơn vị
        ma_don_vi = ""
        ten_don_vi = ten_truong_da_hoc

        # Nếu có trong lịch sử thì kéo thông tin chuẩn ra
        if ten_truong_da_hoc in school_info_db:
            ma_don_vi = school_info_db[ten_truong_da_hoc].get("ma_don_vi", "")
            ten_don_vi = school_info_db[ten_truong_da_hoc].get("ten_don_vi", ten_truong_da_hoc)

        print("-" * 50)
        print(f"File: {file_name}")
        print(f"Trường trên bằng: {ten_truong_da_hoc}")
        print(f"Đơn vị quản lý : {ten_don_vi} (Mã: {ma_don_vi})")

        # Tạm dừng để hỏi STT
        stt_so = input("--> Nhập STT sổ gốc (VD: 01, 15...): ").strip()

        # Tính toán công thức
        ma_so_goc = f"{nam_chung}_VP_{stt_so}"
        ten_tai_lieu = f"Sổ gốc văn bằng năm {nam_chung} - {ten_truong_da_hoc}"

        # FIX: Tránh lỗi MergedCell bằng cách truyền tham số value=
        ws.cell(row=row_idx, column=1, value=row_idx - 1)  # STT (tự tăng từ 1)
        ws.cell(row=row_idx, column=2, value=ma_don_vi)    # Mã đơn vị
        ws.cell(row=row_idx, column=3, value=ten_don_vi)   # Tên đơn vị
        ws.cell(row=row_idx, column=4, value=ma_so_goc)    # Mã sổ gốc
        ws.cell(row=row_idx, column=5, value=ten_tai_lieu) # Tên tài liệu
        ws.cell(row=row_idx, column=6, value=nam_chung)    # Từ năm
        ws.cell(row=row_idx, column=7, value=nam_chung)    # Đến năm
        ws.cell(row=row_idx, column=8, value=ngay_lap_so)  # Ngày lập (hôm nay)
        ws.cell(row=row_idx, column=9, value="")           # Số trang (bỏ trống)
        ws.cell(row=row_idx, column=10, value=tinh_trang)  # Nguyên vẹn
        ws.cell(row=row_idx, column=11, value="")          # Thông tư (bỏ trống)

        row_idx += 1

    # 5. Lưu file tổng hợp
    output_danhsach = target_path / f"DS_Tai_Lieu_So_Hoa_{nam_chung}.xlsx"
    wb.save(output_danhsach)

    print("\n" + "="*60)
    print(f"[=] HOÀN TẤT! Đã tổng hợp xong {len(excel_files)} trường.")
    print(f"[=] File danh sách được lưu tại: {output_danhsach}")
