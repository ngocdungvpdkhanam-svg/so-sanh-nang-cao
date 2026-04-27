import streamlit as st
import pandas as pd
from openpyxl import load_workbook
from openpyxl.styles import PatternFill
from openpyxl.utils import column_index_from_string
import io

st.set_page_config(page_title="Excel Compare by Column Letter", layout="wide")

st.title("🚀 So sánh Excel theo Tên Cột (A, B, C...)")
st.markdown("Nhập trực tiếp các cột cần so sánh theo ký hiệu Excel (ví dụ: cột A, cột B...).")

# --- BƯỚC 1: TẢI FILE ---
st.sidebar.header("1. Tải lên dữ liệu")
file_goc = st.sidebar.file_uploader("Tải File Gốc", type="xlsx")
file_ss = st.sidebar.file_uploader("Tải File So Sánh", type="xlsx")

if file_goc and file_ss:
    # --- BƯỚC 2: CẤU HÌNH ---
    st.header("2. Thiết lập điều kiện so sánh")
    num_pairs = st.number_input("Số lượng cặp cột muốn so sánh:", min_value=1, max_value=10, value=1)
    
    col_mappings = []
    cols = st.columns(num_pairs)
    
    for i in range(num_pairs):
        with cols[i]:
            st.subheader(f"Cặp {i+1}")
            g_letter = st.text_input(f"Cột File GỐC {i+1} (Ví dụ: A)", key=f"gl_{i}").upper().strip()
            s_letter = st.text_input(f"Cột File SO SÁNH {i+1} (Ví dụ: B)", key=f"sl_{i}").upper().strip()
            if g_letter and s_letter:
                col_mappings.append({'g_letter': g_letter, 's_letter': s_letter})
            
    logic_type = st.radio("Chọn phép toán logic:", ("VÀ (AND) - Khớp tất cả các cặp", "HOẶC (OR) - Chỉ cần khớp 1 cặp"))

    # --- BƯỚC 3: XỬ LÝ ---
    if st.button("Bắt đầu xử lý nhanh"):
        if len(col_mappings) < num_pairs:
            st.error("Vui lòng nhập đầy đủ tên cột cho các cặp đã chọn.")
        else:
            try:
                with st.spinner('Đang tính toán...'):
                    # 1. Đọc file so sánh bằng Pandas (Không lấy header để dùng index)
                    # Chúng ta đọc toàn bộ file, sau đó sẽ lọc theo index của cột
                    df_ss = pd.read_excel(file_ss, header=None)
                    
                    # 2. Đọc file gốc bằng Pandas (Không lấy header)
                    df_goc = pd.read_excel(file_goc, header=None)

                    # 3. Tạo các mặt nạ (masks) để so sánh
                    masks = []
                    for mapping in col_mappings:
                        # Chuyển chữ cái cột sang index (0-based) cho Pandas
                        idx_g = column_index_from_string(mapping['g_letter']) - 1
                        idx_s = column_index_from_string(mapping['s_letter']) - 1
                        
                        # Lấy danh sách giá trị duy nhất ở cột file so sánh (ép kiểu string, xóa khoảng trắng)
                        val_ss = set(df_ss[idx_s].astype(str).str.strip().unique())
                        
                        # Tạo mask cho file gốc
                        mask = df_goc[idx_g].astype(str).str.strip().isin(val_ss)
                        masks.append(mask)

                    # 4. Kết hợp Logic
                    final_mask = masks[0]
                    for m in masks[1:]:
                        if "AND" in logic_type:
                            final_mask &= m
                        else:
                            final_mask |= m

                    # Lấy danh sách hàng cần bôi đỏ (Pandas index + 1 để ra dòng Excel)
                    # Vì df đọc header=None nên hàng 1 trong Excel là index 0
                    rows_to_color = df_goc.index[final_mask].tolist()
                    rows_to_color = [r + 1 for r in rows_to_color]

                    # 5. Dùng Openpyxl để tô màu (Giữ nguyên định dạng file gốc)
                    file_goc.seek(0)
                    wb = load_workbook(file_goc)
                    ws = wb.active
                    red_fill = PatternFill(start_color="FFFF0000", end_color="FFFF0000", fill_type="solid")
                    
                    for r_idx in rows_to_color:
                        for cell in ws[r_idx]:
                            cell.fill = red_fill

                    # 6. Xuất file
                    output = io.BytesIO()
                    wb.save(output)
                    
                    st.success(f"Hoàn thành! Đã bôi đỏ {len(rows_to_color)} hàng.")
                    st.download_button("📥 Tải File Kết Quả", output.getvalue(), "ket_qua_so_sanh.xlsx")
            
            except Exception as e:
                st.error(f"Lỗi: Có thể bạn nhập tên cột không tồn tại hoặc file không đúng định dạng. Chi tiết: {e}")
else:
    st.info("Vui lòng tải 2 file Excel để bắt đầu.")
