from pathlib import Path
from utils import select_folder, select_file
from features.ocr_processor import run_ocr_feature
from features.excel_splitter import run_excel_splitter
from features.danhsach_generator import generate_danhsach
from features.sheet_renamer import run_sheet_renamer

def show_menu():
    print("="*40)
    print("HỆ THỐNG SỐ HÓA VĂN BẰNG CHỨNG CHỈ")
    print("="*40)
    print("1. Chạy OCR trích xuất ảnh sổ sang Excel")
    print("2. Tách trường từ file Excel")
    print("3. Tổng hợp Danh sách tài liệu số hóa")
    print("4. Đồng loạt đổi tên Sheet thành 'Data'")
    print("0. Thoát")
    print("="*40)

def main():
    while True:
        show_menu()
        choice = input("Nhập lựa chọn của bạn (0-4): ").strip()

        if choice == "1":
            print("\nVui lòng chọn thư mục chứa ảnh...")
            folder = select_folder("Chọn thư mục ảnh trường (vd: 01, 02)")
            if folder:
                run_ocr_feature(folder)
            else:
                print("[-] Đã hủy chọn thư mục.\n")

        elif choice == "2":
            template_file = Path("data/template.xlsx")
            mapping_file = Path("data/quy_tac_anh_xa.txt")

            if not template_file.exists():
                print(f"[-] LỖI: Không tìm thấy file mẫu tại '{template_file}'. Vui lòng bổ sung!")
                continue
            if not mapping_file.exists():
                print(f"[-] LỖI: Không tìm thấy file quy tắc tại '{mapping_file}'. Vui lòng bổ sung!")
                continue

            print("\n[Bước 1/2] Vui lòng chọn File dữ liệu CẦN TÁCH...")
            input_file = select_file("Chọn File Excel Dữ Liệu Gốc")
            if not input_file:
                print("[-] Đã hủy thao tác.")
                continue

            print("\n[Bước 2/2] Vui lòng chọn THƯ MỤC LƯU CÁC FILE KẾT QUẢ...")
            output_dir = select_folder("Chọn thư mục xuất kết quả")
            if not output_dir:
                print("[-] Đã hủy thao tác.")
                continue

            run_excel_splitter(str(input_file), str(template_file), str(mapping_file), str(output_dir))

        elif choice == "3":
            template_danhsach_path = Path("data/template_danhsach.xlsx")

            if not template_danhsach_path.exists():
                print(f"[-] LỖI: Không tìm thấy file mẫu tại '{template_danhsach_path}'. Vui lòng bổ sung!")
                continue

            generate_danhsach(str(template_danhsach_path))

        elif choice == "4":
            run_sheet_renamer()

        elif choice == "0":
            print("Đang thoát chương trình...")
            break

        else:
            print("\n[-] Lựa chọn không hợp lệ. Vui lòng thử lại!\n")

if __name__ == "__main__":
    main()
