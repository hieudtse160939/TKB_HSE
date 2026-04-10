import pandas as pd
import openpyxl
import streamlit as st
import io
import re

# =====================================================================
# 1. BỘ TỪ ĐIỂN TỰ ĐỘNG CỦA TRƯỜNG HOA SEN
# =====================================================================
TU_DIEN = {
    ("Vân", "L"): "Cô Vân (Lý)",
    ("Vân", "CĐ Lý"): "Cô Vân (Lý)",
    ("Vân", "Đ"): "Cô Vân (Địa)",
    ("Vân", "A"): "Cô Thảo Vân (Anh)",
    ("Vân", "IELTS/A"): "Cô Thảo Vân (Anh)",
    ("Vân", "TNHN"): "Cô Vân (Lý)",
    
    ("Nhung", "V"): "Cô T.Nhung (Văn)",
    ("Nhung", "CĐ Văn"): "Cô T.Nhung (Văn)",
    ("Nhung", "A"): "Cô Nhung (Anh)",
    ("Nhung", "AVTH"): "Cô Nhung (Anh)",
    ("Nhung", "TNHN"): "Cô Nhung (TNHN)",
    
    ("Tâm", "V"): "Thầy Tâm (Văn)",
    ("Tâm", "CĐ Văn"): "Thầy Tâm (Văn)",
    ("Tâm", "AVTH"): "Cô Tâm (Anh)",
    
    ("Ngọc", "L"): "Cô Ngọc (Lý)",
    ("Ngọc", "CĐ Lý"): "Cô Ngọc (Lý)",
    ("Ngọc", "KTPL"): "Thầy Ngọc (KTPL)",
    ("Ngọc", "CĐ KTPL"): "Thầy Ngọc (KTPL)",
    ("Ngọc", "CN"): "Cô Ngọc (Công Nghệ)",

    ("Phương", "V"): "Cô Phương (Văn)",
    ("Phương", "Su"): "Cô Phương (Sử)",
    ("Phương", "CĐ Sử"): "Cô Phương (Sử)",
    
    ("Anh", "KTPL"): "Cô Lan Anh (GDCD/KTPL)",
    ("Anh", "CĐ KTPL"): "Cô Lan Anh (GDCD/KTPL)",
    ("Anh", "GDCD"): "Cô Lan Anh (GDCD/KTPL)",
    
    ("Nghĩa", "T"): "Thầy Nghĩa (Toán)",
    ("Nghĩa", "CĐ Toán"): "Thầy Nghĩa (Toán)",
    ("nghĩa", "CĐ Toán"): "Thầy Nghĩa (Toán)",
    
    ("Bình", "V"): "Thầy/Cô Bình (Văn)",
    ("Bình", "CĐ Văn"): "Thầy/Cô Bình (Văn)",
    
    ("Bảo", "Su"): "Thầy Bảo (Sử/GDĐP)",
    ("Bảo", "SĐ"): "Thầy Bảo (Sử/GDĐP)",
    
    ("Chi", "H"): "Cô Chi (Hóa)",
    ("Chi", "TNHN"): "Cô Chi (Hóa)",
    
    ("Diệp", "V"): "Cô Diệp (Văn)",
    ("Diệp", "CĐ Văn"): "Cô Diệp (Văn)",
    ("Diệp", "Su"): "Cô Diệp (Sử)",
    ("Diệp", "GDĐP"): "Cô Diệp (Sử)",
    
    ("Xuân", "GDCD"): "Cô Xuân (GDCD/KTPL)",
    ("Xuân", "KTPL"): "Cô Xuân (GDCD/KTPL)",
    ("Xuân", "CĐ KTPL"): "Cô Xuân (GDCD/KTPL)",
    
    ("Vinh", "GDCD"): "Cô Vinh (GDCD/KTPL)",
    ("Vinh", "KTPL"): "Cô Vinh (GDCD/KTPL)",
    ("Vinh", "CĐ KTPL"): "Cô Vinh (GDCD/KTPL)",
}

def get_standard_name(ten_ky_hieu, mon_hoc):
    if mon_hoc == 'Chủ nhiệm':
        return f"{ten_ky_hieu} (GVCN)"
    if (ten_ky_hieu, mon_hoc) in TU_DIEN:
        return TU_DIEN[(ten_ky_hieu, mon_hoc)]
    return ten_ky_hieu

def phan_loai_khoi(ten_lop):
    """Hàm tự động phân loại Khối THCS và THPT"""
    ten_lop = str(ten_lop).strip()
    if ten_lop.startswith(('10', '11', '12')):
        return 'THPT'
    elif ten_lop.startswith(('6', '7', '8', '9')):
        return 'THCS'
    else:
        return 'Khác'

# =====================================================================
# 2. HÀM XỬ LÝ CHÍNH TRÊN MEMORY
# =====================================================================
def process_tkb_data(uploaded_file):
    wb = openpyxl.load_workbook(uploaded_file)
    
    for sheet_name in wb.sheetnames:
        ws = wb[sheet_name]
        merged_ranges = list(ws.merged_cells.ranges)
        
        for merged_range in merged_ranges:
            min_col, min_row, max_col, max_row = merged_range.bounds
            top_left_cell_value = ws.cell(row=min_row, column=min_col).value
            ws.unmerge_cells(str(merged_range))
            for row in range(min_row, max_row + 1):
                for col in range(min_col, max_col + 1):
                    ws.cell(row=row, column=col).value = top_left_cell_value
                    
        max_row_current = ws.max_row
        if max_row_current > 66:
            ws.delete_rows(67, max_row_current - 66)
            
    virtual_workbook = io.BytesIO()
    wb.save(virtual_workbook)
    virtual_workbook.seek(0)
    
    df = pd.read_excel(virtual_workbook) 
    
    class_row_idx = 3
    gvcn_row_idx = 4
    classes = {} 
    gvcn = {}    
    
    class_row = df.iloc[class_row_idx]
    gvcn_row = df.iloc[gvcn_row_idx]
    
    for col_idx, val in enumerate(class_row):
        # Regex CHỈ bắt tên lớp bắt đầu bằng số từ 6 đến 12 (6, 7, 8, 9, 10, 11, 12)
        if pd.notna(val) and isinstance(val, str) and re.match(r'^(1[0-2]|[6-9])', val.strip()): 
            class_name = val.strip()
            classes[col_idx] = class_name
            gv_val = str(gvcn_row.iloc[col_idx]).strip()
            if '-' in gv_val:
                gvcn[class_name] = gv_val.split('-')[0].strip()
            else:
                gvcn[class_name] = gv_val

    records = []
    for row_idx in range(5, len(df)):
        row = df.iloc[row_idx]
        for col_idx, class_name in classes.items():
            cell_val = str(row.iloc[col_idx]).strip()
            if cell_val == 'nan' or cell_val == '' or cell_val in ['CHÀO CỜ', 'SINH HOẠT ĐẦU GIỜ', 'THỂ DỤC THỂ THAO']:
                continue
            
            khoi = phan_loai_khoi(class_name)

            if cell_val.lower() == 'chủ nhiệm':
                ten_goc = gvcn.get(class_name, 'Unknown')
                records.append({
                    'Khối': khoi,
                    'Giáo viên': get_standard_name(ten_goc, 'Chủ nhiệm'),
                    'Lớp': class_name, 
                    'Môn': 'Chủ nhiệm'
                })
            elif '-' in cell_val:
                parts = cell_val.split('-')
                if len(parts) >= 2:
                    mon_hoc = "-".join(parts[:-1]).strip() 
                    ten_goc = parts[-1].strip()        
                    records.append({
                        'Khối': khoi,
                        'Giáo viên': get_standard_name(ten_goc, mon_hoc),
                        'Lớp': class_name, 
                        'Môn': mon_hoc
                    })
            else:
                records.append({
                    'Khối': khoi,
                    'Giáo viên': 'Chung (Không ghi tên)',
                    'Lớp': class_name, 
                    'Môn': cell_val
                })

    df_records = pd.DataFrame(records)
    if df_records.empty:
        return None
        
    df_summary = df_records.groupby(['Khối', 'Giáo viên', 'Lớp', 'Môn']).size().reset_index(name='Số tiết')
    df_summary = df_summary.sort_values(by=['Khối', 'Giáo viên', 'Lớp', 'Môn'])
    
    return df_summary

# =====================================================================
# 3. GIAO DIỆN STREAMLIT VỚI FILTER VÀ EXPORT
# =====================================================================
st.set_page_config(page_title="Thống kê TKB", page_icon="📊", layout="wide")

st.title("📊 Công cụ Xử lý & Thống kê TKB")

# --- KHỞI TẠO SESSION STATE ---
if 'df_data' not in st.session_state:
    st.session_state.df_data = None

# --- KHU VỰC UPLOAD ---
uploaded_file = st.file_uploader("Kéo thả hoặc chọn file TKB (.xlsx) vào đây", type=["xlsx"])

if uploaded_file is not None:
    if st.button("Bắt đầu phân tích dữ liệu", type="primary"):
        with st.spinner('Đang xử lý dữ liệu...'):
            df_ket_qua = process_tkb_data(uploaded_file)
            
            if df_ket_qua is not None:
                st.session_state.df_data = df_ket_qua
                st.success("Đã phân tích xong! Bạn có thể dùng bộ lọc bên dưới.")
            else:
                st.error("Không tìm thấy dữ liệu hợp lệ trong file Excel.")

# --- KHU VỰC BỘ LỌC & HIỂN THỊ ---
if st.session_state.df_data is not None:
    df = st.session_state.df_data
    
    st.divider()
    st.subheader("🔍 Bộ lọc dữ liệu")
    
    col1, col2, col3 = st.columns(3)
    
    with col1:
        danh_sach_khoi = sorted(df['Khối'].unique().tolist())
        khoi_chon = st.multiselect("📚 Chọn Khối (Để trống để xem tất cả):", danh_sach_khoi)

    with col2:
        danh_sach_gv = sorted(df['Giáo viên'].unique().tolist())
        gv_chon = st.multiselect("👩‍🏫 Chọn Giáo viên:", danh_sach_gv)
        
    with col3:
        danh_sach_lop = sorted(df['Lớp'].unique().tolist())
        lop_chon = st.multiselect("🏫 Chọn Lớp:", danh_sach_lop)

    # --- ÁP DỤNG BỘ LỌC ---
    df_filtered = df.copy()
    
    if len(khoi_chon) > 0:
        df_filtered = df_filtered[df_filtered['Khối'].isin(khoi_chon)]

    if len(gv_chon) > 0:
        df_filtered = df_filtered[df_filtered['Giáo viên'].isin(gv_chon)]
        
    if len(lop_chon) > 0:
        df_filtered = df_filtered[df_filtered['Lớp'].isin(lop_chon)]

    # --- HIỂN THỊ KẾT QUẢ ĐÃ LỌC ---
    st.markdown(f"**Hiển thị {len(df_filtered)} kết quả:**")
    st.dataframe(df_filtered, use_container_width=True)
    
    # Tính tổng số tiết theo bộ lọc
    tong_so_tiet = df_filtered['Số tiết'].sum()
    st.metric(label="🎯 Tổng số tiết (theo bộ lọc hiện tại)", value=f"{tong_so_tiet} tiết")
    
    st.divider()
    
    # --- XUẤT FILE DỮ LIỆU ĐÃ LỌC ---
    csv = df_filtered.to_csv(index=False, encoding='utf-8-sig').encode('utf-8-sig')
    st.download_button(
        label="📥 Tải file kết quả (Excel/CSV)",
        data=csv,
        file_name="ThongKe_TKB_DaLoc.csv",
        mime="text/csv",
    )
