import sys
from src.converters.excel_to_pdf import excel_to_pdf

def main():
    if len(sys.argv) != 3:
        print("Usage: python main_excel_to_pdf.py <input_excel_path> <output_pdf_path>")
        sys.exit(1)

    input_excel = sys.argv[1]
    output_pdf = sys.argv[2]

    try:
        excel_to_pdf(input_excel, output_pdf)
        print(f"Chuyển đổi thành công từ {input_excel} sang {output_pdf}")
    except Exception as e:
        print(f"Lỗi: {e}")

if __name__ == "__main__":
    main()
