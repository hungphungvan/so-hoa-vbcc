from utils import select_folder
from features.ocr_processor import run_ocr_feature

def show_menu():
    print("="*40)
    print("HỆ THỐNG SỐ HÓA VĂN BẰNG CHỨNG CHỈ")
    print("="*40)
    print("1. Chạy OCR trích xuất ảnh sổ sang CSV")
    print("2. [Tính năng tương lai] Gộp nhiều CSV thành Excel")
    print("3. [Tính năng tương lai] Chuẩn hóa tên & ngày sinh")
    print("0. Thoát")
    print("="*40)

def main():
    while True:
        show_menu()
        choice = input("Nhập lựa chọn của bạn (0-3): ").strip()

        if choice == "1":
            print("\nVui lòng chọn thư mục chứa ảnh...")
            folder = select_folder("Chọn thư mục ảnh trường (vd: 01, 02)")
            if folder:
                run_ocr_feature(folder)
            else:
                print("[-] Đã hủy chọn thư mục.\n")

        elif choice == "0":
            print("Đang thoát chương trình...")
            break

        else:
            print("\n[-] Tính năng chưa khả dụng hoặc lựa chọn không hợp lệ. Vui lòng thử lại!\n")

if __name__ == "__main__":
    main()
