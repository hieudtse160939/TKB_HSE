import pandas as pd
import openpyxl
import streamlit as st
import io
import re

# =====================================================================
# 1. BỘ TỪ ĐIỂN ĐẦY ĐỦ (Đã bao gồm Chuyên đề CĐ)
# =====================================================================
TU_DIEN = {
    # --- TỔ VĂN ---
    ("Việt", "V"): "Thầy Lê Công Quốc Việt (Văn)",
    ("Việt", "CĐ Văn"): "Thầy Lê Công Quốc Việt (Văn)",
    ("Nhung", "V"): "Cô Trần Thị Tuyết Nhung (Văn)",
    ("Nhung", "CĐ Văn"): "Cô Trần Thị Tuyết Nhung (Văn)",
    ("Tâm", "V"): "Thầy Ngô Đình Tâm (Văn)",
    ("Tâm", "CĐ Văn"): "Thầy Ngô Đình Tâm (Văn)",
    ("Bình", "V"): "Cô Lê Thị Bình (Văn)",
    ("Bình", "CĐ Văn"): "Cô Lê Thị Bình (Văn)",
    ("Hạnh", "V"): "Cô Lê Ngọc Bích Hạnh (Văn)",
    ("Hạnh", "CĐ Văn"): "Cô Lê Ngọc Bích Hạnh (Văn)",
    ("Nhàn", "V"): "Cô Lưu Thị Nhàn (Văn)",
    ("Nhàn", "CĐ Văn"): "Cô Lưu Thị Nhàn (Văn)",
    ("Thoa", "V"): "Cô Nguyễn Thị Thoa (Văn)",
    ("Thoa", "CĐ Văn"): "Cô Nguyễn Thị Thoa (Văn)",
    ("Oanh", "V"): "Cô Võ Thị Oanh (Văn)",
    ("Oanh", "CĐ Văn"): "Cô Võ Thị Oanh (Văn)",
    ("Như", "V"): "Cô Võ Nguyễn Huỳnh Như (Văn)",
    ("Như", "CĐ Văn"): "Cô Võ Nguyễn Huỳnh Như (Văn)",
    ("Diệp", "V"): "Cô Phan Thị Ngọc Diệp (Văn)",
    ("Diệp", "CĐ Văn"): "Cô Phan Thị Ngọc Diệp (Văn)",
    ("Nam", "V"): "Thầy Nguyễn Văn Nam (Văn)",

    # --- TỔ TOÁN ---
    ("Nghĩa", "T"): "Thầy Nghĩa (Toán)",
    ("Nghĩa", "CĐ Toán"): "Thầy Nghĩa (Toán)",
    ("Như", "T"): "Cô Nguyễn Trương Quỳnh Như (Toán)",
    ("Như", "CĐ Toán"): "Cô Nguyễn Trương Quỳnh Như (Toán)",
    ("Nhi", "T"): "Cô Lê Thị Yến Nhi (Toán)",
    ("Nhi", "CĐ Toán"): "Cô Lê Thị Yến Nhi (Toán)",
    ("Sang", "T"): "Thầy Huỳnh Phước Sang (Toán)",
    ("Vũ", "T"): "Thầy Hồ Thế Vũ (Toán)",
    ("Thọ", "T"): "Thầy Phan Minh Thọ (Toán)",
    ("Thuận", "T"): "Thầy Lư Quang Thuận (Toán)",
    ("Ánh", "T"): "Cô Lê Thị Ngọc Ánh (Toán)",

    # --- TỔ ANH VĂN ---
    ("Vân", "A"): "Cô Nguyễn Thảo Vân (Anh)",
    ("Lài", "A"): "Cô Nguyễn Thị Lài (Anh)",
    ("Nhi", "A"): "Cô Trần Trúc Nhi (Anh)",
    ("Nhung", "A"): "Cô Nguyễn Thị Phương Nhung (Anh)",
    ("Biền", "A"): "Cô Đinh Thị Biền (Anh)",
    ("Giang", "A"): "Cô Bùi Thị Giang (Anh)",

    # --- CÁC MÔN KHÁC ---
    ("Bảo", "Su"): "Thầy Vương Quốc Bảo (Sử)",
    ("Phương", "Su"): "Cô Tống Thị Phương (Sử)",
    ("Vy", "Đ"): "Cô Phan Thị Thảo Vy (Địa)",
    ("Ngân", "L"): "Cô Hà Thị Kim Ngân (Lý)",
    ("Ngân", "CĐ Lý"): "Cô Hà Thị Kim Ngân (Lý)",
    ("Chi", "H"): "Cô Trần Thị Quí Chi (Hóa)",
    ("Xuân", "KTPL"): "Cô Xuân (KTPL)",
    ("Vinh", "KTPL"): "Cô Vinh (KTPL)",
    ("Anh", "KTPL"): "Cô Lan Anh (KTPL)",
}

# Danh sách từ khóa để nhận diện Môn học
DS_MON = ["V", "T", "L", "H", "A", "SU", "Đ", "KTPL", "GDCD", "CĐ", "SI", "CN", "KHTN", "SĐ", "GDĐP"]

def get_standard_info(p1, p2):
    """Xử lý thông minh để tách Môn và Tên giáo viên"""
    p1, p2 = p1.strip(), p2.strip()
    
    # 1. Kiểm tra trong từ điển (thử cả 2 chiều)
    if (p1, p2) in TU_DIEN:
        return TU_DIEN[(p1, p2)], p2
    if (p2, p1) in TU_DIEN:
        return TU_DIEN[(p2, p1)], p1

    # 2. Nếu không có trong từ điển, tự đoán dựa trên danh sách môn
    # Nếu phần nào chứa chữ "CĐ" hoặc nằm trong DS_MON thì đó là Môn
    p1_is_mon = any(m in p1.upper() for m in DS_MON)
    if p1_is_mon:
        return f"{p2} ({p1})", p1
    else:
        return f"{p1} ({p2})", p2

def phan_loai_khoi(ten_lop):
    ten_lop = str(ten_lop).strip()
    if ten_lop.startswith(('10', '11', '12')): return 'THPT'
    if ten_lop.startswith(('6', '7', '8', '9')): return 'THCS'
    return 'Khác'

# =====================================================================
# 2. XỬ LÝ FILE
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
        if ws.max_row > 66:
            ws.delete_rows(67, ws.max_row - 66)

    virtual_workbook = io.BytesIO()
    wb.save(virtual_workbook)
    virtual_workbook.seek(0)
    df = pd.read_excel(virtual_workbook)

    # Lấy thông tin Lớp và GVCN
    class_row_idx, gvcn_row_idx = 3, 4
    classes, gvcn = {}, {}
    class_row = df.iloc[class_row_idx]
    gvcn_row = df.iloc[gvcn_row_idx]

    for col_idx, val in enumerate(class_row):
        if pd.notna(val) and isinstance(val, str) and re.match(r'^(1[0-2]|[6-9])', val.strip()):
            class_name = val.strip()
            classes[col_idx] = class_name
            gv_val = str(gvcn_row.iloc[col_idx]).strip()
            gvcn[class_name] = gv_val.split('-')[0].strip() if '-' in gv_val else gv_val

    records = []
    for row_idx in range(5, len(df)):
        row = df.iloc[row_idx]
        for col_idx, class_name in classes.items():
            cell_val = str(row.iloc[col_idx]).strip()
            if cell_val in ['nan', '', 'CHÀO CỜ', 'SINH HOẠT ĐẦU GIỜ', 'THỂ DỤC THỂ THAO']:
                continue
            
            khoi = phan_loai_khoi(class_name)

            if cell_val.lower() == 'chủ nhiệm':
                ten_gv = gvcn.get(class_name, 'Unknown')
                records.append({'Khối': khoi, 'Giáo viên': f"{ten_gv} (GVCN)", 'Lớp': class_name, 'Môn': 'Chủ nhiệm'})
            elif '-' in cell_val:
                parts = cell_val.split('-')
                p1, p2 = parts[0], parts[1]
                ten_hien_thi, mon_hoc = get_standard_info(p1, p2)
                records.append({'Khối': khoi, 'Giáo viên': ten_hien_thi, 'Lớp': class_name, 'Môn': mon_hoc})
            else:
                records.append({'Khối': khoi, 'Giáo viên': f"Chung ({cell_val})", 'Lớp': class_name, 'Môn': cell_val})

    df_res = pd.DataFrame(records)
    if df_res.empty: return None
    # Thống kê số tiết
    df_summary = df_res.groupby(['Khối', 'Giáo viên', 'Lớp', 'Môn']).size().reset_index(name='Số tiết')
    return df_summary

# =====================================================================
# 3. GIAO DIỆN STREAMLIT
# =====================================================================
st.set_page_config(page_title="Hoa Sen TKB", layout="wide")
st.title("📊 Thống kê Tiết dạy Hoa Sen (Full Chuyên đề)")

if 'df_data' not in st.session_state:
    st.session_state.df_data = None

uploaded_file = st.file_uploader("Tải file TKB (.xlsx)", type=["xlsx"])

if uploaded_file:
    if st.button("Phân tích dữ liệu", type="primary"):
        res = process_tkb_data(uploaded_file)
        if res is not None:
            st.session_state.df_data = res
            st.success("Đã xử lý xong dữ liệu!")

if st.session_state.df_data is not None:
    df = st.session_state.df_data
    
    # Khu vực Bộ lọc
    st.divider()
    col_f1, col_f2, col_f3 = st.columns(3)
    k_filter = col_f1.multiselect("Lọc Khối:", sorted(df['Khối'].unique()))
    g_filter = col_f2.multiselect("Lọc Giáo viên:", sorted(df['Giáo viên'].unique()))
    l_filter = col_f3.multiselect("Lọc Lớp:", sorted(df['Lớp'].unique()))

    # Áp dụng lọc
    df_filtered = df.copy()
    if k_filter: df_filtered = df_filtered[df_filtered['Khối'].isin(k_filter)]
    if g_filter: df_filtered = df_filtered[df_filtered['Giáo viên'].isin(g_filter)]
    if l_filter: df_filtered = df_filtered[df_filtered['Lớp'].isin(l_filter)]

    # Hiển thị
    c_left, c_right = st.columns([1, 2])
    
    with c_left:
        st.subheader("📈 Tổng số tiết/GV")
        tong_hop = df_filtered.groupby('Giáo viên')['Số tiết'].sum().reset_index().sort_values('Số tiết', ascending=False)
        st.dataframe(tong_hop, hide_index=True, use_container_width=True)
        st.metric("Tổng tiết hiển thị", f"{tong_hop['Số tiết'].sum()} tiết")

    with c_right:
        st.subheader("📋 Chi tiết tiết dạy")
        st.dataframe(df_filtered, hide_index=True, use_container_width=True)

    # Download
    csv = df_filtered.to_csv(index=False, encoding='utf-8-sig').encode('utf-8-sig')
    st.download_button("📥 Tải kết quả CSV", data=csv, file_name="ThongKe_TKB_HoaSen.csv")
