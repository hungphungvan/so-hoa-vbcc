import re
import pandas as pd
import openpyxl
from openpyxl.utils.cell import column_index_from_string
from pathlib import Path

# Thêm 2 dòng này để tắt các cảnh báo vô hại của openpyxl
import warnings
warnings.filterwarnings("ignore", category=UserWarning, module="openpyxl")

def load_mapping_rules(mapping_file: str) -> dict:
    """Đọc quy tắc ánh xạ, phân biệt giữa cột (A,B,C) và giá trị tĩnh ("Text")"""
    mapping = {}
    with open(mapping_file, 'r', encoding='utf-8') as f:
        for line in f:
            line = line.strip()
            if not line or line.startswith('#'):
                continue
            if '=' in line:
                key, val = line.split('=', 1)
                val = val.strip()
                if val:
                    # Nếu có ngoặc kép ở đầu và cuối -> Là giá trị cố định (constant)
                    if (val.startswith('"') and val.endswith('"')) or (val.startswith("'") and val.endswith("'")):
                        mapping[key.strip()] = ('const', val[1:-1]) # Xóa ngoặc kép và lưu
                    else:
                        # Ngược lại -> Là ký tự cột (column)
                        mapping[key.strip()] = ('col', val.upper())
    return mapping

def run_excel_splitter(input_file: str, template_file: str, mapping_file: str, output_dir: str):
    print(f"\n[+] Đang nạp quy tắc ánh xạ từ: {Path(mapping_file).name}")
    mapping_rules = load_mapping_rules(mapping_file)

    print(f"[+] Đang phân tích dữ liệu từ file Demo: {Path(input_file).name}")
    try:
        # Đọc toàn bộ dữ liệu, header=None để tính tọa độ cột tuyệt đối (A=0, B=1...)
        df = pd.read_excel(input_file, sheet_name=0, header=None)
    except Exception as e:
        print(f"[-] LỖI đọc file input: {e}")
        return

    # 1. Xác định cột chứa Tên Trường để làm mỏ neo tách file
    school_col_name = "Tên trường đã học"
    if school_col_name not in mapping_rules:
        print(f"[-] LỖI: File cấu hình txt thiếu dòng '{school_col_name} = ...'")
        return

    map_type, map_val = mapping_rules[school_col_name]
    if map_type != 'col':
        print(f"[-] LỖI: Cột '{school_col_name}' phải là ký hiệu cột (ví dụ G), không thể là giá trị tĩnh.")
        return

    school_col_idx = column_index_from_string(map_val) - 1

    # 2. Định vị dữ liệu
    # Dữ liệu học sinh bắt đầu từ dòng số 4 trong Excel (tương ứng index 3 của pandas)
    DATA_START_ROW = 3
    df_data = df.iloc[DATA_START_ROW:].copy()

    # Lọc bỏ các dòng trống (dòng không có tên trường)
    df_data = df_data.dropna(subset=[school_col_idx])

    schools = df_data[school_col_idx].unique()
    print(f"[+] Tìm thấy {len(schools)} trường. Bắt đầu tách và kết xuất...\n")

    # 3. Quét Header của file Template (nằm ở dòng 2)
    wb_temp = openpyxl.load_workbook(template_file)
    ws_temp = wb_temp['Data']
    template_headers = {str(cell.value).strip(): i for i, cell in enumerate(ws_temp[2], start=1) if cell.value}
    wb_temp.close()

    # 4. Tiến hành tách file
    for school in schools:
        df_school = df_data[df_data[school_col_idx] == school]

        # Mở một bản copy của Template
        wb = openpyxl.load_workbook(template_file)
        ws = wb['Data']
        start_row = 3 # Ghi dữ liệu từ dòng số 3
        stt_counter = 1 # Khởi tạo biến đếm số thứ tự bắt đầu từ 1 cho mỗi trường

        for _, row in df_school.iterrows():
            # 1. Tự động ghi Số thứ tự (STT) nếu cột này có tồn tại trong file Template
            if "STT" in template_headers:
                stt_col_idx = template_headers["STT"]
                ws.cell(row=start_row, column=stt_col_idx).value = stt_counter

            # 2. Ghi các dữ liệu khác theo quy tắc ánh xạ
            for temp_col, (map_type, map_val) in mapping_rules.items():
                if temp_col == "STT":
                    continue

                if temp_col in template_headers:
                    col_idx = template_headers[temp_col]

                    # NẾU LÀ GIÁ TRỊ CỐ ĐỊNH -> Điền thẳng vào ô
                    if map_type == 'const':
                        ws.cell(row=start_row, column=col_idx).value = map_val

                    # NẾU LÀ TỌA ĐỘ CỘT -> Lấy từ dòng dữ liệu của file Demo
                    elif map_type == 'col':
                        demo_idx = column_index_from_string(map_val) - 1

                        if demo_idx < len(row):
                            val = row.iloc[demo_idx]

                            # Chuẩn hóa dữ liệu chống lỗi
                            if pd.isna(val):
                                val = ""
                            else:
                                if temp_col in ["Ngày, tháng, năm sinh", "Ngày tháng năm cấp bằng"]:
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
            stt_counter += 1 # Tăng số thứ tự lên 1 cho học sinh tiếp theo

        # Lưu file
        safe_school_name = re.sub(r'[<>:"/\\|?*]', '_', str(school)).strip()
        num_students = len(df_school)

        # Đặt tên file theo quy tắc: Tên_trường_Số_học_sinh.xlsx
        file_name = f"{safe_school_name}_{num_students}.xlsx"
        output_path = Path(output_dir) / file_name

        # Lưu file
        wb.save(output_path)
        print(f"  -> Đã tạo: {output_path.name} (Chứa {num_students} học sinh)")

    print(f"\n[=] HOÀN TẤT! Toàn bộ file đã được lưu tại: {output_dir}\n")
