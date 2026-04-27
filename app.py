import pandas as pd
from openpyxl import load_workbook
from openpyxl.styles import PatternFill
from google.colab import files
from openpyxl.utils import column_index_from_string

# 1. Tải file lên
print("--- BƯỚC 1: TẢI FILE GỐC ---")
uploaded_goc = files.upload()
file_goc_name = list(uploaded_goc.keys())[0]

print("\n--- BƯỚC 2: TẢI FILE SO SÁNH ---")
uploaded_ss = files.upload()
file_ss_name = list(uploaded_ss.keys())[0]

# 2. Thiết lập cấu hình so sánh
print("\n--- BƯỚC 3: CẤU HÌNH SO SÁNH ĐA ĐIỀU KIỆN ---")
num_conditions = int(input("Bạn muốn so sánh bao nhiêu cặp cột? (VD: 2): "))

conditions = []
for i in range(num_conditions):
    print(f"\nCặp thứ {i+1}:")
    c_goc = input(f"  - Tên cột File Gốc (VD: A): ").upper()
    c_ss = input(f"  - Tên cột File So Sánh tương ứng (VD: B): ").upper()
    conditions.append({'goc': c_goc, 'ss': c_ss})

logic_type = input("\nChọn phép toán logic (Gõ 'AND' để tất cả phải khớp, 'OR' để chỉ cần 1 cái khớp): ").upper()

# 3. Xử lý dữ liệu
print("\nĐang xử lý dữ liệu...")

# Đọc file so sánh
wb_ss = load_workbook(file_ss_name, data_only=True)
ws_ss = wb_ss.active

# Chuyển đổi dữ liệu file so sánh thành một danh sách để tra cứu cho nhanh
data_ss_sets = []
for cond in conditions:
    col_idx = column_index_from_string(cond['ss'])
    # Lấy toàn bộ giá trị của cột đó trong file so sánh, đưa vào set để tìm kiếm cực nhanh
    values = {str(ws_ss.cell(row=r, column=col_idx).value).strip() for r in range(1, ws_ss.max_row + 1) if ws_ss.cell(row=r, column=col_idx).value is not None}
    data_ss_sets.append(values)

# Mở file gốc để tô màu
wb_goc = load_workbook(file_goc_name)
ws_goc = wb_goc.active
red_fill = PatternFill(start_color="FFFF0000", end_color="FFFF0000", fill_type="solid")

count = 0
# Duyệt từng hàng của file gốc
for row in range(2, ws_goc.max_row + 1):
    results = []
    
    # Kiểm tra từng điều kiện
    for i in range(num_conditions):
        col_goc_idx = column_index_from_string(conditions[i]['goc'])
        val_goc = str(ws_goc.cell(row=row, column=col_goc_idx).value).strip()
        
        # Kiểm tra xem giá trị này có nằm trong cột tương ứng của file SS không
        is_match = val_goc in data_ss_sets[i]
        results.append(is_match)
    
    # Áp dụng logic AND hoặc OR
    should_highlight = False
    if logic_type == "AND":
        if all(results): # Tất cả đều True
            should_highlight = True
    else: # Mặc định là OR
        if any(results): # Chỉ cần 1 cái True
            should_highlight = True
            
    if should_highlight:
        for cell in ws_goc[row]:
            cell.fill = red_fill
        count += 1

# 4. Xuất kết quả
output_name = "ket_qua_so_sanh_nang_cao.xlsx"
wb_goc.save(output_name)
print(f"\n--- HOÀN THÀNH ---")
print(f"Đã tìm thấy và bôi đỏ {count} hàng thỏa mãn điều kiện {logic_type}.")
files.download(output_name)
