import streamlit as st
import pandas as pd
from openpyxl import load_workbook
from openpyxl.styles import PatternFill
from openpyxl.utils import column_index_from_string
import io

st.set_page_config(page_title="Chuyên gia So sánh Excel", layout="wide")

st.title("🛡️ Công cụ So sánh Excel (Chính xác Tuyệt đối)")

# --- HÀM CHUẨN HÓA DỮ LIỆU (QUAN TRỌNG NHẤT) ---
def clean_data(value):
    if value is None or pd.isna(value):
        return ""
    # Chuyển về string, viết thường, xóa khoảng trắng 2 đầu
    s = str(value).strip().lower()
    # Nếu là số dạng 100.0 thì chuyển về 100
    if s.endswith('.0'):
        s = s[:-2]
    return s

# --- BƯỚC 1: TẢI FILE ---
st.sidebar.header("1. Tải lên dữ liệu")
file_goc = st.sidebar.file_uploader("Tải File Gốc", type="xlsx")
file_ss = st.sidebar.file_uploader("Tải File So Sánh", type="xlsx")

if file_goc and file_ss:
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

    if st.button("🚀 BẮT ĐẦU KIỂM TRA"):
        try:
            with st.spinner('Đang phân tích dữ liệu chuyên sâu...'):
                # Đọc file bằng engine mặc định để đảm bảo tương thích tốt nhất
                df_goc = pd.read_excel(file_goc, header=None)
                df_ss = pd.read_excel(file_ss, header=None)

                # 1. Chuẩn hóa và tạo tập hợp tra cứu
                lookup_sets = []
                for m in col_mappings:
                    # Áp dụng hàm clean_data cho toàn bộ cột file so sánh
                    s_set = set(df_ss[m['s_idx']].apply(clean_data).unique())
                    # Loại bỏ giá trị rỗng khỏi set so sánh để tránh bôi đỏ các hàng trống
                    if "" in s_set: s_set.remove("")
                    lookup_sets.append(s_set)

                # 2. So sánh từng hàng của file gốc
                results_indices = []
                for idx, row in df_goc.iterrows():
                    row_matches = []
                    for i, m in enumerate(col_mappings):
                        val_goc = clean_data(row[m['g_idx']])
                        # Kiểm tra xem giá trị gốc có trong tập hợp so sánh không
                        row_matches.append(val_goc != "" and val_goc in lookup_sets[i])
                    
                    # Áp dụng logic AND/OR
                    if "AND" in logic_type:
                        if all(row_matches) and len(row_matches) > 0:
                            results_indices.append(idx + 1)
                    else:
                        if any(row_matches):
                            results_indices.append(idx + 1)

                # 3. Tô màu bằng Openpyxl
                file_goc.seek(0)
                wb = load_workbook(file_goc)
                ws = wb.active
                red_fill = PatternFill(start_color="FFFF0000", end_color="FFFF0000", fill_type="solid")
                
                max_col = ws.max_column
                for r_idx in results_indices:
                    for c_idx in range(1, max_col + 1):
                        ws.cell(row=r_idx, column=c_idx).fill = red_fill

                # 4. Xuất file
                output = io.BytesIO()
                wb.save(output)
                
                st.success(f"Đã kiểm tra xong! Tìm thấy {len(results_indices)} hàng khớp hoàn toàn.")
                st.download_button("📥 Tải Kết Quả", output.getvalue(), "ket_qua_chinh_xac.xlsx")

        except Exception as e:
            st.error(f"Lỗi: {e}")
