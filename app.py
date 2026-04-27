import streamlit as st
import pandas as pd
from openpyxl import load_workbook
from openpyxl.styles import PatternFill
from openpyxl.utils import column_index_from_string
import io

st.set_page_config(page_title="Turbo Excel Compare", layout="wide")

st.title("⚡ Công cụ So sánh Excel - Phiên bản Turbo")
st.markdown("Đã tối ưu hóa thuật toán để xử lý hàng chục nghìn dòng trong vài giây.")

# --- BƯỚC 1: TẢI FILE ---
st.sidebar.header("1. Tải lên dữ liệu")
file_goc = st.sidebar.file_uploader("Tải File Gốc", type="xlsx")
file_ss = st.sidebar.file_uploader("Tải File So Sánh", type="xlsx")

if file_goc and file_ss:
    # --- BƯỚC 2: CẤU HÌNH ---
    st.header("2. Thiết lập điều kiện")
    num_pairs = st.number_input("Số lượng cặp cột muốn so sánh:", min_value=1, max_value=10, value=1)
    
    col_mappings = []
    ui_cols = st.columns(num_pairs)
    
    for i in range(num_pairs):
        with ui_cols[i]:
            st.subheader(f"Cặp {i+1}")
            g_letter = st.text_input(f"Cột Gốc (VD: A)", key=f"gl_{i}").upper().strip()
            s_letter = st.text_input(f"Cột So Sánh (VD: B)", key=f"sl_{i}").upper().strip()
            if g_letter and s_letter:
                col_mappings.append({
                    'g_idx': column_index_from_string(g_letter) - 1,
                    's_idx': column_index_from_string(s_letter) - 1
                })
            
    logic_type = st.radio("Chọn phép toán logic:", ("VÀ (AND)", "HOẶC (OR)"), horizontal=True)

    if st.button("🚀 XỬ LÝ SIÊU TỐC"):
        if len(col_mappings) < num_pairs:
            st.error("Vui lòng nhập đủ tên cột.")
        else:
            try:
                with st.spinner('Đang tính toán thần tốc...'):
                    # Dùng engine 'calamine' để đọc file với tốc độ tối đa
                    # header=None để tính theo cột A, B, C...
                    df_ss = pd.read_excel(file_ss, header=None, engine='calamine')
                    df_goc = pd.read_excel(file_goc, header=None, engine='calamine')

                    # 1. Tối ưu tra cứu: Chuyển các cột so sánh thành set (Hash Set)
                    # Việc kiểm tra 1 giá trị trong Set nhanh gấp hàng nghìn lần trong List
                    ss_lookup_sets = []
                    for m in col_mappings:
                        # Lấy cột tương ứng, xóa bỏ giá trị trống, chuyển thành string và lưu vào set
                        s_set = set(df_ss[m['s_idx']].dropna().astype(str).str.strip().unique())
                        ss_lookup_sets.append(s_set)

                    # 2. Vector hóa việc so sánh (Vô cùng nhanh)
                    masks = []
                    for i, m in enumerate(col_mappings):
                        # Kiểm tra toàn bộ cột của file gốc xem có nằm trong set tra cứu không
                        mask = df_goc[m['g_idx']].astype(str).str.strip().isin(ss_lookup_sets[i])
                        masks.append(mask)

                    # 3. Tính toán logic mask cuối cùng
                    final_mask = masks[0]
                    for m in masks[1:]:
                        if "AND" in logic_type:
                            final_mask &= m
                        else:
                            final_mask |= m

                    # 4. Xác định các hàng cần tô màu (Lấy index)
                    rows_to_color = df_goc.index[final_mask].tolist()
                    rows_to_color = [r + 1 for r in rows_to_color] # Excel bắt đầu từ 1

                    # 5. Bước tô màu (Dùng Openpyxl tối ưu)
                    file_goc.seek(0)
                    wb = load_workbook(file_goc)
                    ws = wb.active
                    red_fill = PatternFill(start_color="FFFF0000", end_color="FFFF0000", fill_type="solid")
                    
                    # Tối ưu: Chỉ duyệt qua các hàng có trong danh sách trùng
                    # Và chỉ tô màu đến cột cuối cùng có dữ liệu (không tô hết 16.000 cột của Excel)
                    max_col = ws.max_column
                    for r_idx in rows_to_color:
                        for c_idx in range(1, max_col + 1):
                            ws.cell(row=r_idx, column=c_idx).fill = red_fill

                    # 6. Xuất file
                    output = io.BytesIO()
                    wb.save(output)
                    
                    st.success(f"Xong! Đã xử lý {len(df_goc)} dòng. Tìm thấy {len(rows_to_color)} hàng trùng khớp.")
                    st.download_button("📥 Tải Kết Quả", output.getvalue(), "ket_qua_turbo.xlsx")
            
            except Exception as e:
                st.error(f"Lỗi kỹ thuật: {e}")
else:
    st.info("Chờ tải file...")
