import json
import re
import time
from pathlib import Path

import pandas as pd

from config import MODEL_NAME, api_client

# Nhúng thêm hàm save_file từ utils
from utils import encode_image, read_text_file, save_file


def extract_table_from_image(image_path: str, prompt_text: str, max_retries: int = 3) -> str:
    """Gửi ảnh qua API có kèm cơ chế tự động thử lại khi lỗi mạng."""
    base64_image = encode_image(image_path)

    for attempt in range(max_retries):
        try:
            response = api_client.chat.completions.create(
                model=MODEL_NAME,
                messages=[
                    {
                        "role": "user",
                        "content": [
                            {"type": "text", "text": prompt_text},
                            {
                                "type": "image_url",
                                "image_url": {"url": f"data:image/jpeg;base64,{base64_image}"},
                            },
                        ],
                    }
                ],
                temperature=0.1,
            )

            # FIX 1: Ép kiểu an toàn. Nếu API trả về None, lấy chuỗi rỗng.
            content = response.choices[0].message.content
            return content if content is not None else ""

        # FIX 2: Thêm comment noqa để báo linter bỏ qua quy tắc bắt lỗi chung (BLE001) tại đây
        except Exception as e:  # noqa: BLE001
            print(f"\n      [!] Lỗi API (Lần thử {attempt + 1}/{max_retries}): {e}")
            if attempt == max_retries - 1:
                return ""
            print("      -> Đợi 5 giây rồi thử lại...")
            time.sleep(5)

def run_ocr_feature(folder_path_str: str):
    """Quét ảnh trong thư mục trường, tìm prompt theo 3 cấp, lưu thẳng ra Excel qua hộp thoại Save As."""
    data_dir = Path(folder_path_str)

    valid_extensions = {".jpg", ".jpeg", ".png", ".heic", ".heif"}
    image_files = sorted([f for f in data_dir.iterdir() if f.suffix.lower() in valid_extensions])

    if not image_files:
        print(f"[-] Thư mục '{data_dir.name}' trống hoặc không chứa file ảnh hợp lệ.")
        return

    print(f"\n[+] Tìm thấy {len(image_files)} ảnh trong thư mục '{data_dir.name}' (Năm học: {data_dir.parent.name}).")

    # CẤU TRÚC TÌM KIẾM PROMPT
    local_prompt = data_dir / "prompt.txt"
    year_prompt = data_dir.parent / "prompt.txt"
    default_prompt = Path(__file__).parent.parent / "data" / "default_prompt.txt"

    if local_prompt.is_file():
        print(f"[+] Đang dùng prompt ưu tiên của TRƯỜNG tại: {local_prompt}")
        prompt_text = read_text_file(local_prompt)
    elif year_prompt.is_file():
        print(f"[+] Đang dùng prompt chung của NĂM HỌC tại: {year_prompt}")
        prompt_text = read_text_file(year_prompt)
    elif default_prompt.is_file():
        print(f"[+] Đang dùng prompt MẶC ĐỊNH của hệ thống tại: {default_prompt}")
        prompt_text = read_text_file(default_prompt)
    else:
        print("[-] LỖI: Không tìm thấy file prompt nào trong hệ thống.")
        return

    print("--- Bắt đầu OCR ---")
    all_students_data = []

    for img_path in image_files:
        print(f"  -> Xử lý: {img_path.name}...", end=" ")

        json_str = ""

        try:
            result_text = extract_table_from_image(str(img_path), prompt_text)

            if not result_text:
                print("BỎ QUA DO LỖI MẠNG")
                continue

            match = re.search(r'\[.*\]', result_text, re.DOTALL)

            if match:
                json_str = match.group(0)
                page_data = json.loads(json_str)

                if isinstance(page_data, list):
                    all_students_data.extend(page_data)
                    print(f"OK ({len(page_data)} học sinh)")
                else:
                    print("LỖI (Dữ liệu không phải là list JSON)")
            else:
                print("LỖI (AI không trả về mảng JSON nào)")
                print(f"\n--- DỮ LIỆU THÔ AI TRẢ VỀ ---\n{result_text}\n-----------------------------\n")

        except json.JSONDecodeError as e:
            print(f"LỖI Parse JSON ({e})")
            print(f"\n--- CHUỖI JSON BỊ LỖI ---\n{json_str}\n-------------------------\n")
        except Exception as e:  # noqa: BLE001
            print(f"LỖI HỆ THỐNG ({e})")

    if all_students_data:
        truong_name = data_dir.name

        print("\n[+] Đang mở cửa sổ lưu file... (Kiểm tra taskbar nếu không thấy)")

        # Mở hộp thoại Save As, gợi ý sẵn tên file
        default_filename = f"truong_{truong_name}_ketqua.xlsx"
        save_path_str = save_file(title="Lưu file kết quả OCR", default_name=default_filename)

        if not save_path_str:
            print("[-] Bạn đã hủy lưu file. Dữ liệu chưa được xuất ra!")
            return

        excel_output = Path(save_path_str)
        # Đổi đuôi file excel thành .json để lưu file backup chung chỗ
        json_output = excel_output.with_suffix('.json')

        # 1. Lưu file backup JSON
        with open(json_output, "w", encoding="utf-8") as f:
            json.dump(all_students_data, f, ensure_ascii=False, indent=4)

        # 2. Lưu trực tiếp ra Excel bằng Pandas
        df = pd.DataFrame(all_students_data)
        df.to_excel(excel_output, index=False)

        print(f"\n[=] HOÀN TẤT: Đã trích xuất {len(all_students_data)} học sinh.")
        print(f"  -> File Excel : {excel_output}")
        print(f"  -> File JSON  : {json_output}\n")
    else:
        print("\n[-] Không có dữ liệu hợp lệ để lưu.\n")
