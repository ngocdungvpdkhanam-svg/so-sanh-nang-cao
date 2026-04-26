import streamlit as st
import pandas as pd
from openpyxl import load_workbook
from openpyxl.styles import PatternFill
import io

st.title("🚀 Công cụ So sánh Excel Online")

col1, col2 = st.columns(2)
with col1:
    file_goc = st.file_uploader("Tải File Gốc (xlsx)", type="xlsx")
with col2:
    file_ss = st.file_uploader("Tải File So Sánh (xlsx)", type="xlsx")

col_a = st.text_input("Tên cột File Gốc cần so sánh (ví dụ: Ma_Hang)", "ID")
col_b = st.text_input("Tên cột File So Sánh cần so sánh", "ID")

if st.button("Bắt đầu so sánh và bôi đỏ"):
    if file_goc and file_ss:
        # Đọc dữ liệu
        df_goc = pd.read_excel(file_goc)
        df_ss = pd.read_excel(file_ss)
        
        # Lấy danh sách trùng khớp
        trung_khop = df_ss[col_b].unique()
        
        # Xử lý file Excel để giữ nguyên định dạng và bôi màu
        file_goc.seek(0)
        wb = load_workbook(file_goc)
        ws = wb.active
        
        red_fill = PatternFill(start_color="FFFF0000", end_color="FFFF0000", fill_type="solid")
        
        # Tìm vị trí cột A trong File Gốc
        header = [cell.value for cell in ws[1]]
        try:
            col_idx = header.index(col_a) + 1
            
            for row in range(2, ws.max_row + 1):
                cell_value = ws.cell(row=row, column=col_idx).value
                if cell_value in trung_khop:
                    for cell in ws[row]:
                        cell.fill = red_fill
            
            # Xuất file
            output = io.BytesIO()
            wb.save(output)
            st.success("Đã xử lý xong!")
            st.download_button("Tải file kết quả tại đây", data=output.getvalue(), file_name="ket_qua_so_sanh.xlsx")
            
        except ValueError:
            st.error(f"Không tìm thấy cột '{col_a}' trong file gốc.")
    else:
        st.warning("Vui lòng tải lên cả 2 file.")
