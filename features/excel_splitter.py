import json
import re
import warnings
from pathlib import Path

import openpyxl
import pandas as pd
from openpyxl.utils.cell import column_index_from_string
from thefuzz import process

warnings.filterwarnings("ignore", category=UserWarning, module="openpyxl")

def load_mapping_rules(mapping_file: str) -> dict:
    mapping = {}
    with open(mapping_file, 'r', encoding='utf-8') as f:
        for line in f:
            if '#' in line:
                line = line.split('#')[0]
            line = line.strip()
            if not line:
                continue
            if '=' in line:
                key, val = line.split('=', 1)
                key = key.strip().upper()
                val = val.strip()
                if val:
                    if (val.startswith('"') and val.endswith('"')) or (val.startswith("'") and val.endswith("'")):
                        mapping[key] = ('const', val[1:-1])
                    else:
                        mapping[key] = ('col', val.upper())
    return mapping

def load_school_catalog(csv_file: str) -> dict:
    catalog = {}
    with open(csv_file, 'r', encoding='utf-8') as f:
        for line in f:
            line = line.strip()
            if not line or ';' not in line or line.startswith('ma_truong'):
                continue
            ma_truong, ten_truong = line.split(';', 1)
            catalog[ten_truong.strip()] = ma_truong.strip()
    return catalog

def run_excel_splitter(input_file: str, template_file: str, mapping_file: str, output_dir: str):
    print(f"\n[+] Đang nạp quy tắc ánh xạ từ: {Path(mapping_file).name}")
    mapping_rules = load_mapping_rules(mapping_file)

    catalog_file = Path("data/danh_muc_truong.csv")
    standard_schools = {}
    if catalog_file.exists():
        standard_schools = load_school_catalog(str(catalog_file))

    history_file = Path("data/lich_su_chuan_hoa.json")
    approved_mapping = {}
    if history_file.exists():
        try:
            with open(history_file, 'r', encoding='utf-8') as f:
                approved_mapping = json.load(f)
                # FIX 1: Chỉ dùng values() vì không cần khóa k
                for v in approved_mapping.values():
                    if not isinstance(v, dict):
                        print("[!] Cấu trúc lịch sử cũ không còn tương thích. Hệ thống sẽ học lại từ đầu...")
                        approved_mapping = {}
                        break
        # FIX 2: Thêm noqa để bỏ qua cảnh báo vơ vét lỗi
        except Exception:  # noqa: BLE001
            approved_mapping = {}

    try:
        df = pd.read_excel(input_file, sheet_name=0, header=None)
    except Exception as e:  # noqa: BLE001
        print(f"[-] LỖI đọc file input: {e}")
        return

    SCHOOL_TEMPLATE_COL = "N"
    if SCHOOL_TEMPLATE_COL not in mapping_rules:
        print(f"[-] LỖI: Bắt buộc cấu hình ánh xạ cho cột {SCHOOL_TEMPLATE_COL} để tách file.")
        return

    # FIX 3: Thêm _ trước map_type vì biến này không được xài đến
    _map_type, map_val = mapping_rules[SCHOOL_TEMPLATE_COL]
    school_col_idx = column_index_from_string(map_val) - 1

    DATA_START_ROW = 3
    df_data = df.iloc[DATA_START_ROW:].copy()
    df_data = df_data.dropna(subset=[school_col_idx])

    raw_schools = df_data[school_col_idx].unique()

    # ---------------------------------------------------------
    # LUỒNG XỬ LÝ: ĐƠN VỊ QUẢN LÝ & TÊN TRƯỜNG ĐÃ HỌC
    # ---------------------------------------------------------
    list_standard_names = list(standard_schools.keys())

    for raw_name in raw_schools:
        raw_name_str = str(raw_name).strip()

        # 1. Nếu tên chuẩn 100% -> Trường độc lập, Đơn vị = Tên trường
        if raw_name_str in standard_schools:
            if raw_name_str not in approved_mapping:
                approved_mapping[raw_name_str] = {
                    "ten_truong": raw_name_str,
                    "ten_don_vi": raw_name_str,
                    "ma_don_vi": standard_schools[raw_name_str]
                }
            continue

        # 2. Đã từng được xác nhận trong lịch sử
        if raw_name_str in approved_mapping:
            continue

        # 3. Tên lạ (Sai chính tả hoặc Trường giải thể) -> Hỏi người dùng 2 bước
        # FIX 4: Xóa chữ 'f' vô nghĩa ở chuỗi có dấu '='
        print("\n" + "="*60)
        print(f"[?] PHÁT HIỆN TRƯỜNG CHƯA CHUẨN: '{raw_name_str}'")

        # --- BƯỚC 3.1: Xác định ĐƠN VỊ QUẢN LÝ ---
        unit_name = ""
        unit_code = ""
        if list_standard_names:
            matches = process.extract(raw_name_str, list_standard_names, limit=3)
            print(f"\n[BƯỚC 1] Đơn vị nào đang quản lý hồ sơ của '{raw_name_str}'?")
            print("Gợi ý đơn vị:")
            for idx, (m_name, score) in enumerate(matches, start=1):
                print(f"  {idx}. {m_name} (Mã: {standard_schools[m_name]}) - Giống: {score}%")
            print("  0. Tự tìm kiếm thủ công trong danh mục")

            while True:
                choice = input(f"Chọn (0-{len(matches)}): ").strip()
                if choice.isdigit() and 0 <= int(choice) <= len(matches):
                    choice = int(choice)
                    break
                print("[-] Nhập sai, thử lại!")

            if choice == 0:
                while True:
                    keyword = input("Nhập từ khóa tìm kiếm Đơn vị (VD: 'Chuyen', 'Tam Duong'): ").strip()
                    search_results = process.extract(keyword, list_standard_names, limit=5)
                    for i, (s_name, s_score) in enumerate(search_results, start=1):
                        print(f"  {i}. {s_name} (Mã: {standard_schools[s_name]})")
                    print("  0. Tìm từ khóa khác")

                    s_choice = input(f"Chọn đơn vị đúng (1-{len(search_results)}) hoặc 0: ").strip()
                    if s_choice.isdigit() and 1 <= int(s_choice) <= len(search_results):
                        unit_name = search_results[int(s_choice) - 1][0]
                        unit_code = standard_schools[unit_name]
                        break
            else:
                unit_name = matches[choice - 1][0]
                unit_code = standard_schools[unit_name]

        # --- BƯỚC 3.2: Xác định TÊN TRƯỜNG ĐÃ HỌC ---
        # FIX 5: Xóa chữ 'f' vô nghĩa ở chuỗi này
        print("\n[BƯỚC 2] 'Tên trường đã học' ghi trên phôi bằng của học sinh sẽ là gì?")
        print(f"  1. Giữ nguyên gốc: '{raw_name_str}' (Dành cho trường giải thể/sát nhập)")
        print(f"  2. Lấy tên Đơn vị: '{unit_name}' (Dành cho lỗi sai chính tả)")
        print("  3. Tự gõ tên trường chuẩn xác...")

        final_school_name = ""
        while True:
            t_choice = input("Lựa chọn (1-3): ").strip()
            if t_choice == '1':
                final_school_name = raw_name_str
                break
            elif t_choice == '2':
                final_school_name = unit_name
                break
            elif t_choice == '3':
                final_school_name = input("Nhập tên trường chuẩn xác: ").strip()
                if final_school_name:
                    break
            print("[-] Vui lòng chọn 1, 2 hoặc 3.")

        # Lưu lại lịch sử
        approved_mapping[raw_name_str] = {
            "ten_truong": final_school_name,
            "ten_don_vi": unit_name,
            "ma_don_vi": unit_code
        }
        with open(history_file, 'w', encoding='utf-8') as f:
            json.dump(approved_mapping, f, ensure_ascii=False, indent=4)

        print(f"  -> ĐÃ LƯU: Tên trên bằng [{final_school_name}] | Đơn vị quản lý [{unit_name}]")

    print("\n[+] Bắt đầu tách và kết xuất Excel...\n")

    # ---------------------------------------------------------
    # TIẾN HÀNH TÁCH FILE
    # ---------------------------------------------------------
    for raw_school in raw_schools:
        raw_name_str = str(raw_school).strip()
        school_info = approved_mapping[raw_name_str]

        final_school_name = school_info["ten_truong"]
        unit_name = school_info["ten_don_vi"]
        unit_code = school_info["ma_don_vi"]

        df_school = df_data[df_data[school_col_idx] == raw_school]

        wb = openpyxl.load_workbook(template_file)
        ws = wb['Data']

        ma_dv_col, ten_dv_col = None, None
        for cell in ws[2]:
            header_val = str(cell.value).strip() if cell.value else ""
            if header_val == "Mã đơn vị":
                ma_dv_col = cell.column
            elif header_val == "Tên đơn vị":
                ten_dv_col = cell.column

        start_row = 3
        stt_counter = 1

        for _, row in df_school.iterrows():
            # FIX 6: Truyền value vào tham số để tránh lỗi gán cho MergedCell
            ws.cell(row=start_row, column=1, value=stt_counter)

            for temp_letter, (m_type, m_val) in mapping_rules.items():
                if temp_letter == "A":
                    continue
                col_idx = column_index_from_string(temp_letter)
                if m_type == 'const':
                    ws.cell(row=start_row, column=col_idx, value=m_val)
                elif m_type == 'col':
                    demo_idx = column_index_from_string(m_val) - 1
                    if demo_idx < len(row):
                        # Cột N (Tên trường) giờ sẽ dùng final_school_name (Tên lịch sử hoặc Tên đã sửa chính tả)
                        if demo_idx == school_col_idx:
                            val = final_school_name
                        else:
                            val = row.iloc[demo_idx]

                        if pd.isna(val):
                            val = ""
                        else:
                            if temp_letter in ["I", "S"]:
                                try:
                                    val = pd.to_datetime(val, dayfirst=True).strftime("%d/%m/%Y")
                                except Exception:  # noqa: BLE001
                                    val = str(val).strip()
                            elif isinstance(val, pd.Timestamp):
                                val = val.strftime("%d/%m/%Y")
                            elif isinstance(val, float) and val.is_integer():
                                val = str(int(val))
                            else:
                                val = str(val).strip()

                        ws.cell(row=start_row, column=col_idx, value=val)

            # Ghi tự động Đơn vị quản lý vào các cột tương ứng nếu có
            if ma_dv_col:
                ws.cell(row=start_row, column=ma_dv_col, value=unit_code)
            if ten_dv_col:
                ws.cell(row=start_row, column=ten_dv_col, value=unit_name)

            start_row += 1
            stt_counter += 1

        # Tên file xuất ra chỉ lấy Tên trường đã học và Số lượng học sinh
        safe_school_name = re.sub(r'[<>:"/\\|?*]', '_', final_school_name).strip()
        num_students = len(df_school)

        file_name = f"{safe_school_name}_{num_students}.xlsx"

        output_path = Path(output_dir) / file_name
        wb.save(output_path)
        print(f"  -> Đã tạo: {file_name}")

    print(f"\n[=] HOÀN TẤT! Toàn bộ file đã được lưu tại: {output_dir}\n")
