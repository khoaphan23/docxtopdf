import pandas as pd
from matplotlib.backends.backend_pdf import PdfPages
import matplotlib.pyplot as plt

def excel_to_pdf(excel_path, pdf_path):
    # Đọc file Excel vào DataFrame
    df = pd.read_excel(excel_path)

    # Tạo một file PDF để ghi bảng vào
    with PdfPages(pdf_path) as pdf:
        # Tạo figure
        fig, ax = plt.subplots(figsize=(8, 6))
        ax.axis('tight')
        ax.axis('off')
        
        # Tạo bảng trong figure
        table = ax.table(cellText=df.values, colLabels=df.columns, loc='center')

        # Lưu bảng vào PDF
        pdf.savefig(fig, bbox_inches='tight')
        plt.close()

    print(f"Đã chuyển đổi {excel_path} sang {pdf_path}")
