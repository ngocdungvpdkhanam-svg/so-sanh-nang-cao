import streamlit as st
import pandas as pd
from openpyxl import load_workbook
from openpyxl.styles import PatternFill
import io

st.set_page_config(page_title="Excel Fast Compare", layout="wide")

st.title("🚀 Công cụ So sánh Excel Tốc độ cao")

# Sidebar tải file
st.sidebar.header("1. Tải lên dữ liệu")
file_goc = st.sidebar.file_uploader("Tải File Gốc", type="xlsx")
file_ss = st.sidebar.file_uploader("Tải File So Sánh", type="xlsx")

if file_goc and file_ss:
    # Đọc nhanh header để người dùng chọn tên cột
    df_goc_preview = pd.read_excel(file_goc, nrows=0)
    df_ss_preview = pd.read_excel(file_ss, nrows=0)
    
    st.header("2. Thiết lập điều kiện")
    num_pairs = st.number_input("Số lượng cặp cột so sánh:", min_value=1, max_value=10, value=1)
    
    col_mappings = []
    cols = st.columns(num_pairs)
    for i in range(num_pairs):
        with cols[i]:
            st.subheader(f"Cặp {i+1}")
            g_col = st.selectbox(f"Cột File Gốc {i+1}", df_goc_preview.columns, key=f"g_{i}")
            s_col = st.selectbox(f"Cột File So Sánh {i+1}", df_ss_preview.columns, key=f"s_{i}")
            col_mappings.append({'goc': g_col, 'ss': s_col})
            
    logic_type = st.radio("Chọn phép toán logic:", ("VÀ (AND)", "HOẶC (OR)"))

    if st.button("Bắt đầu xử lý nhanh"):
        try:
            with st.spinner('Đang xử lý thần tốc...'):
                # Đọc toàn bộ dữ liệu vào Pandas (Chỉ đọc các cột cần thiết để tiết kiệm RAM)
                cols_goc = [m['goc'] for m in col_mappings]
                cols_ss = [m['ss'] for m in col_mappings]
                
                df_goc = pd.read_excel(file_goc, usecols=cols_goc)
                df_ss = pd.read_excel(file_ss, usecols=cols_ss)

                # Chuẩn hóa dữ liệu: ép kiểu về string và xóa khoảng trắng để so sánh chính xác
                for col in cols_goc: df_goc[col] = df_goc[col].astype(str).str.strip()
                for col in cols_ss: df_ss[col] = df_ss[col].astype(str).str.strip()

                # TÍNH TOÁN LOGIC BẰNG PANDAS (Đây là phần giúp tăng tốc)
                masks = []
                for mapping in col_mappings:
                    # Kiểm tra xem giá trị trong cột file gốc có nằm trong tập hợp giá trị của cột file so sánh không
                    set_ss = set(df_ss[mapping['ss']].unique())
                    masks.append(df_goc[mapping['goc']].isin(set_ss))

                # Kết hợp các mask theo logic AND/OR
                final_mask = masks[0]
                for m in masks[1:]:
                    if "AND" in logic_type:
                        final_mask &= m
                    else:
                        final_mask |= m

                # Lấy danh sách index của các hàng cần bôi đỏ (Pandas index bắt đầu từ 0, Excel từ 2)
                rows_to_color = df_goc.index[final_mask].tolist()
                rows_to_color = [r + 2 for r in rows_to_color] 

                # BƯỚC TÔ MÀU: Chỉ thực hiện trên những hàng đã biết chắc là trùng
                file_goc.seek(0)
                wb = load_workbook(file_goc)
                ws = wb.active
                red_fill = PatternFill(start_color="FFFF0000", end_color="FFFF0000", fill_type="solid")
                
                for row_idx in rows_to_color:
                    for cell in ws[row_idx]:
                        cell.fill = red_fill

                # Xuất file
                output = io.BytesIO()
                wb.save(output)
                
                st.success(f"Xử lý xong! Đã bôi đỏ {len(rows_to_color)} hàng.")
                st.download_button("📥 Tải File Kết Quả", output.getvalue(), "ket_qua_nhanh.xlsx")

        except Exception as e:
            st.error(f"Lỗi: {e}")
else:
    st.info("Vui lòng tải file để bắt đầu.")
