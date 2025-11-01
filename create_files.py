import pandas as pd
import numpy as np

print("Bắt đầu tạo file Excel phức tạp...")

# --- 1. Định nghĩa Cấu trúc Cột (Columns) ---
# Mục tiêu: ~30 cột, 3 cấp độ (Category, Group, Metric)

# Cấp 1: 2 Categories chính
col_level_0_names = ['Financials', 'Operations']

# Cấp 2: Các nhóm con
fin_groups = ['Product A', 'Product B', 'Product C'] # 3 nhóm
ops_groups = ['Region North', 'Region South', 'Region East', 'Region West'] # 4 nhóm

# Cấp 3: Các chỉ số
fin_metrics = ['Revenue', 'COGS', 'Profit', 'Margin %'] # 4 chỉ số
ops_metrics = ['Transactions', 'New Users', 'Avg. Ticket Size', 'Support Tickets'] # 4 chỉ số

# Tạo danh sách các tuple cho MultiIndex Cột
col_tuples = []

# Financials: 3 groups * 4 metrics = 12 cột
for group in fin_groups:
    for metric in fin_metrics:
        col_tuples.append((col_level_0_names[0], group, metric))
        
# Operations: 4 groups * 4 metrics = 16 cột
for group in ops_groups:
    for metric in ops_metrics:
        col_tuples.append((col_level_0_names[1], group, metric))

# Tạo MultiIndex cho cột
# Tổng số cột = 12 + 16 = 28 cột (gần 30)
column_index = pd.MultiIndex.from_tuples(col_tuples, names=['Category', 'Group', 'Metric'])


# --- 2. Định nghĩa Cấu trúc Dòng (Rows) ---
# Mục tiêu: 50 dòng, 2 cấp độ (Department, Team)

row_level_0_names = ['Sales', 'Marketing', 'Engineering', 'Data', 'Support'] # 5 departments
row_level_1_names = [f'Team_{chr(65+i)}' for i in range(10)] # 10 teams (A-J)

# Tạo MultiIndex cho dòng
# Tổng số dòng = 5 * 10 = 50 dòng
row_index = pd.MultiIndex.from_product([row_level_0_names, row_level_1_names], names=['Department', 'Team'])


# --- 3. Tạo Dữ liệu Giả (Dummy Data) ---
num_rows = 50
num_cols = 28

# Bắt đầu với dữ liệu số nguyên ngẫu nhiên
data = np.random.randint(100, 5000, size=(num_rows, num_cols))

# Tạo DataFrame
df = pd.DataFrame(data, index=row_index, columns=column_index)

# --- 4. Tăng độ phức tạp (Mixed Types & NaNs) ---
# Thao tác này rất quan trọng để testing code đọc dữ liệu
print("Đang thêm độ phức tạp: mixed types và NaNs...")

# Chuyển đổi một số cột sang float (ví dụ: Revenue, Avg. Ticket Size)
# Sử dụng .loc với tuple slicing để chọn cột đa cấp

# *** ĐÂY LÀ CÁC DÒNG ĐÃ SỬA ***
# Lỗi cũ: ('Financials', :, 'Revenue')
# Sửa đúng: ('Financials', slice(None), 'Revenue')
df.loc[:, ('Financials', slice(None), 'Revenue')] = np.random.rand(num_rows, len(fin_groups)) * 10000 + 5000

# Lỗi cũ: ('Operations', :, 'Avg. Ticket Size')
# Sửa đúng: ('Operations', slice(None), 'Avg. Ticket Size')
df.loc[:, ('Operations', slice(None), 'Avg. Ticket Size')] = np.random.rand(num_rows, len(ops_groups)) * 100 + 50

# Lỗi cũ: ('Financials', :, 'Margin %')
# Sửa đúng: ('Financials', slice(None), 'Margin %')
df.loc[:, ('Financials', slice(None), 'Margin %')] = np.random.rand(num_rows, len(fin_groups)) * 0.5 + 0.1

# Thêm một số giá trị rỗng (NaN) một cách có chủ đích
df.iloc[3:8, 2] = np.nan      # Thêm NaN vào cột ('Financials', 'Product A', 'Profit')
df.iloc[10:15, 7] = np.nan   # Thêm NaN vào cột ('Financials', 'Product B', 'Margin %')
df.iloc[25, 15:20] = np.nan  # Thêm NaN vào 1 dòng ở nhiều cột Operations

# --- 5. Xuất ra file Excel ---
file_name = 'complex_hierarchical_data_test.xlsx'
print(f"Đang ghi dữ liệu ra file '{file_name}'...")

# 'index=True' là bắt buộc để lưu cấu trúc multi-index của dòng
# 'header=True' là bắt buộc để lưu cấu trúc multi-index của cột
df.to_excel(file_name, sheet_name='ComplexReport', index=True, header=True)

print("\n-------------------------------------------------")
print(f"Đã tạo file '{file_name}' thành công.")
print(f"Kích thước: {df.shape[0]} dòng x {df.shape[1]} cột")
print("Xem trước 5 dòng đầu của dữ liệu (phần Financials):")
print(df['Financials'].head())
print("-------------------------------------------------")