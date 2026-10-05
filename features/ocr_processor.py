import json
import csv
import re
import time
from pathlib import Path
from config import api_client, MODEL_NAME
from utils import encode_image, read_text_file

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
            return response.choices[0].message.content

        except Exception as e:
            print(f"\n      [!] Lỗi API (Lần thử {attempt + 1}/{max_retries}): {e}")
            if attempt == max_retries - 1:
                return ""  # Hết lượt thử, trả về chuỗi rỗng để không làm chết chương trình
            print("      -> Đợi 5 giây rồi thử lại...")
            time.sleep(5)

def run_ocr_feature(folder_path_str: str):
    """Quét ảnh trong thư mục trường, tìm prompt theo 3 cấp, lưu CSV."""
    data_dir = Path(folder_path_str)

    valid_extensions = {".jpg", ".jpeg", ".png"}
    image_files = sorted([f for f in data_dir.iterdir() if f.suffix.lower() in valid_extensions])

    if not image_files:
        print(f"[-] Thư mục '{data_dir.name}' trống hoặc không chứa file ảnh hợp lệ.")
        return

    print(f"\n[+] Tìm thấy {len(image_files)} ảnh trong thư mục '{data_dir.name}' (Năm học: {data_dir.parent.name}).")

    # CẤU TRÚC TÌM KIẾM PROMPT MỚI (3 Cấp bậc)
    local_prompt = data_dir / "prompt.txt"                        # Cấp 1: Thư mục Trường (VD: data/2007/01/prompt.txt)
    year_prompt = data_dir.parent / "prompt.txt"                  # Cấp 2: Thư mục Năm (VD: data/2007/prompt.txt)
    default_prompt = Path(__file__).parent / "default_prompt.txt" # Cấp 3: Mặc định hệ thống

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
        try:
            result_text = extract_table_from_image(str(img_path), prompt_text)

            # Nếu hết 3 lần thử mà vẫn lỗi, bỏ qua ảnh này và chạy tiếp ảnh sau
            if not result_text:
                print("BỎ QUA DO LỖI MẠNG")
                continue

            # Sử dụng Regex để trích xuất toàn bộ nội dung nằm giữa [ và ]
            match = re.search(r'\[.*\]', result_text, re.DOTALL)

            if match:
                json_str = match.group(0)
                page_data = json.loads(json_str)

                if isinstance(page_data, list):
                    all_students_data.extend(page_data)
                    print(f"OK ({len(page_data)} học sinh)")
                else:
                    print(f"LỖI (Dữ liệu không phải là list JSON)")
            else:
                print("LỖI (AI không trả về mảng JSON nào)")
                print(f"\n--- DỮ LIỆU THÔ AI TRẢ VỀ ---\n{result_text}\n-----------------------------\n")

        except json.JSONDecodeError as e:
            print(f"LỖI Parse JSON ({e})")
            print(f"\n--- CHUỖI JSON BỊ LỖI ---\n{json_str}\n-------------------------\n")
        except Exception as e:
            print(f"LỖI HỆ THỐNG ({e})")

    if all_students_data:
        truong_name = data_dir.name

        # Định nghĩa đường dẫn cho cả 2 file lưu ở thư mục Năm học
        csv_output = data_dir.parent / f"truong_{truong_name}_ketqua.csv"
        json_output = data_dir.parent / f"truong_{truong_name}_ketqua.json"

        # 1. LƯU FILE JSON
        with open(json_output, "w", encoding="utf-8") as f:
            json.dump(all_students_data, f, ensure_ascii=False, indent=4)

        # 2. TỰ ĐỘNG LẤY CỘT VÀ LƯU FILE CSV
        fieldnames = []
        for row in all_students_data:
            for key in row.keys():
                if key not in fieldnames:
                    fieldnames.append(key)

        with open(csv_output, "w", encoding="utf-8-sig", newline="") as f:
            writer = csv.DictWriter(f, fieldnames=fieldnames)
            writer.writeheader()
            for row in all_students_data:
                clean_row = {key: row.get(key, "") for key in fieldnames}
                writer.writerow(clean_row)

        print(f"\n[=] HOÀN TẤT: Lưu {len(all_students_data)} học sinh.")
        print(f"  -> Đã tạo: {csv_output}")
        print(f"  -> Đã tạo: {json_output}\n")
    else:
        print("\n[-] Không có dữ liệu hợp lệ để lưu.\n")
