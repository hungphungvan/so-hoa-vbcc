import pandas as pd
import openpyxl
from openpyxl.utils.cell import column_index_from_string
from pathlib import Path
import warnings
import re

warnings.filterwarnings("ignore", category=UserWarning, module="openpyxl")

def load_mapping_rules(mapping_file: str) -> dict:
    """Đọc quy tắc theo cấu trúc: CỘT_TEMPLATE = CỘT_DEMO | "GIÁ TRỊ TĨNH" # Ghi chú"""
    mapping = {}
    with open(mapping_file, 'r', encoding='utf-8') as f:
        for line in f:
            # Vứt bỏ phần text phía sau dấu # (phần gợi nhớ)
            if '#' in line:
                line = line.split('#')[0]

            line = line.strip()
            if not line:
                continue

            if '=' in line:
                key, val = line.split('=', 1)
                key = key.strip().upper() # Ký hiệu cột Template (A, B, C...)
                val = val.strip()

                if val:
                    # Nếu bọc trong ngoặc kép -> Giá trị tĩnh
                    if (val.startswith('"') and val.endswith('"')) or (val.startswith("'") and val.endswith("'")):
                        mapping[key] = ('const', val[1:-1])
                    else:
                        # Ký hiệu cột từ file Demo
                        mapping[key] = ('col', val.upper())
    return mapping

def run_excel_splitter(input_file: str, template_file: str, mapping_file: str, output_dir: str):
    print(f"\n[+] Đang nạp quy tắc ánh xạ từ: {Path(mapping_file).name}")
    mapping_rules = load_mapping_rules(mapping_file)

    print(f"[+] Đang phân tích dữ liệu từ file Demo: {Path(input_file).name}")
    try:
        df = pd.read_excel(input_file, sheet_name=0, header=None)
    except Exception as e:
        print(f"[-] LỖI đọc file input: {e}")
        return

    # 1. Mỏ neo tách file: Cột N ("Tên trường đã học" trong Template)
    SCHOOL_TEMPLATE_COL = "N"
    if SCHOOL_TEMPLATE_COL not in mapping_rules:
        print(f"[-] LỖI: Bắt buộc phải cấu hình ánh xạ cho cột {SCHOOL_TEMPLATE_COL} (Tên trường đã học) để tách file.")
        return

    map_type, map_val = mapping_rules[SCHOOL_TEMPLATE_COL]
    if map_type != 'col':
        print(f"[-] LỖI: Cột {SCHOOL_TEMPLATE_COL} phải trỏ đến một ký tự cột của file Demo.")
        return

    school_col_idx = column_index_from_string(map_val) - 1

    # 2. Định vị dữ liệu bắt đầu từ dòng số 4 của file Demo
    DATA_START_ROW = 3
    df_data = df.iloc[DATA_START_ROW:].copy()
    df_data = df_data.dropna(subset=[school_col_idx])

    schools = df_data[school_col_idx].unique()
    print(f"[+] Tìm thấy {len(schools)} trường. Bắt đầu tách và kết xuất...\n")

    # 3. Tiến hành tách file
    for school in schools:
        df_school = df_data[df_data[school_col_idx] == school]

        wb = openpyxl.load_workbook(template_file)
        ws = wb['Data']
        start_row = 3
        stt_counter = 1

        for _, row in df_school.iterrows():
            # Tự động đánh STT tại cột A
            ws.cell(row=start_row, column=1).value = stt_counter

            for temp_letter, (map_type, map_val) in mapping_rules.items():
                if temp_letter == "A": # Bỏ qua cột A
                    continue

                col_idx = column_index_from_string(temp_letter)

                if map_type == 'const':
                    ws.cell(row=start_row, column=col_idx).value = map_val

                elif map_type == 'col':
                    demo_idx = column_index_from_string(map_val) - 1

                    if demo_idx < len(row):
                        val = row.iloc[demo_idx]

                        if pd.isna(val):
                            val = ""
                        else:
                            # Cột I (Ngày sinh) và S (Ngày cấp) trong Template
                            if temp_letter in ["I", "S"]:
                                try:
                                    parsed_date = pd.to_datetime(val, dayfirst=True)
                                    val = parsed_date.strftime("%d/%m/%Y")
                                except Exception:
                                    val = str(val).strip()
                            elif isinstance(val, pd.Timestamp):
                                val = val.strftime("%d/%m/%Y")
                            elif isinstance(val, float) and val.is_integer():
                                val = str(int(val))
                            else:
                                val = str(val).strip()

                        ws.cell(row=start_row, column=col_idx).value = val

            start_row += 1
            stt_counter += 1

        # Lưu file
        safe_school_name = re.sub(r'[<>:"/\\|?*]', '_', str(school)).strip()
        num_students = len(df_school)
        file_name = f"{safe_school_name}_{num_students}.xlsx"
        output_path = Path(output_dir) / file_name

        wb.save(output_path)
        print(f"  -> Đã tạo: {file_name}")

    print(f"\n[=] HOÀN TẤT! Toàn bộ file đã được lưu tại: {output_dir}\n")
