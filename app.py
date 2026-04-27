import streamlit as st
import pandas as pd
from openpyxl import load_workbook
from openpyxl.styles import PatternFill
from openpyxl.utils import column_index_from_string
import io

st.set_page_config(page_title="Công cụ So sánh Excel Đa điều kiện", layout="wide")

st.title("🚀 Công cụ So sánh & Bôi đỏ Excel Online")
st.markdown("""
Công cụ này giúp bạn so sánh nhiều cặp cột giữa 2 file Excel. 
Nếu thỏa mãn điều kiện (VÀ/HOẶC), hàng ở file gốc sẽ được **bôi đỏ**.
""")

# --- BƯỚC 1: TẢI FILE ---
st.sidebar.header("1. Tải lên dữ liệu")
file_goc = st.sidebar.file_uploader("Tải File Gốc (xlsx)", type="xlsx")
file_ss = st.sidebar.file_uploader("Tải File So Sánh (xlsx)", type="xlsx")

if file_goc and file_ss:
    # --- BƯỚC 2: CẤU HÌNH ---
    st.header("2. Thiết lập điều kiện so sánh")
    
    num_pairs = st.number_input("Số lượng cặp cột muốn so sánh:", min_value=1, max_value=10, value=1)
    
    col_mappings = []
    cols = st.columns(num_pairs)
    
    for i in range(num_pairs):
        with cols[i]:
            st.subheader(f"Cặp {i+1}")
            g_col = st.text_input(f"Cột File Gốc {i+1} (VD: A)", key=f"g_{i}").upper()
            s_col = st.text_input(f"Cột File So Sánh {i+1} (VD: B)", key=f"s_{i}").upper()
            col_mappings.append({'goc': g_col, 'ss': s_col})
            
    logic_type = st.radio("Chọn phép toán logic:", ("VÀ (AND) - Phải khớp tất cả các cặp", "HOẶC (OR) - Chỉ cần khớp 1 trong các cặp"))

    # --- BƯỚC 3: XỬ LÝ ---
    if st.button("Bắt đầu xử lý"):
        try:
            with st.spinner('Đang xử lý dữ liệu...'):
                # Đọc file so sánh (dùng data_only=True để lấy giá trị sau công thức)
                wb_ss = load_workbook(file_ss, data_only=True)
                ws_ss = wb_ss.active
                
                # Tạo tập hợp dữ liệu tra cứu cho từng cột ở file so sánh
                data_ss_sets = []
                for mapping in col_mappings:
                    idx = column_index_from_string(mapping['ss'])
                    # Lấy dữ liệu cột, bỏ qua None, chuyển về string và strip khoảng trắng
                    values = {str(ws_ss.cell(row=r, column=idx).value).strip() 
                              for r in range(1, ws_ss.max_row + 1) 
                              if ws_ss.cell(row=r, column=idx).value is not None}
                    data_ss_sets.append(values)

                # Mở file gốc để tô màu
                wb_goc = load_workbook(file_goc)
                ws_goc = wb_goc.active
                red_fill = PatternFill(start_color="FFFF0000", end_color="FFFF0000", fill_type="solid")
                
                count = 0
                # Duyệt qua từng hàng ở file gốc (từ hàng 2 để tránh header)
                for row in range(2, ws_goc.max_row + 1):
                    match_results = []
                    for i in range(num_pairs):
                        g_idx = column_index_from_string(col_mappings[i]['goc'])
                        val_goc = str(ws_goc.cell(row=row, column=g_idx).value).strip()
                        match_results.append(val_goc in data_ss_sets[i])
                    
                    # Kiểm tra logic
                    should_color = False
                    if "AND" in logic_type:
                        if all(match_results): should_color = True
                    else:
                        if any(match_results): should_color = True
                    
                    if should_color:
                        for cell in ws_goc[row]:
                            cell.fill = red_fill
                        count += 1

                # Xuất file kết quả ra bộ nhớ đệm (Buffer)
                output = io.BytesIO()
                wb_goc.save(output)
                processed_data = output.getvalue()
                
                st.success(f"Đã xử lý xong! Tìm thấy {count} hàng trùng khớp.")
                st.download_button(
                    label="📥 Tải File Kết Quả",
                    data=processed_data,
                    file_name="ket_qua_so_sanh.xlsx",
                    mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
                )
        except Exception as e:
            st.error(f"Có lỗi xảy ra: {e}. Vui lòng kiểm tra lại tên cột (A, B, C...)")
else:
    st.info("Vui lòng tải lên cả 2 file Excel ở thanh bên trái để bắt đầu.")
