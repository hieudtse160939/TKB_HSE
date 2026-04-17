import pandas as pd
import openpyxl
import streamlit as st
import io
import re

# =====================================================================
# 1. BỘ TỪ ĐIỂN TỰ ĐỘNG (ĐÃ CẬP NHẬT)
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
    ("Nhi", "T"): "Cô Lê Thị Yến Nhi (Toán)",
    ("Sang", "T"): "Thầy Huỳnh Phước Sang (Toán)",
    ("Sanh", "T"): "Thầy Đặng Văn Sanh (Toán)",
    ("Vũ", "T"): "Thầy Hồ Thế Vũ (Toán)",
    # --- TỔ ANH ---
    ("Vân", "A"): "Cô Nguyễn Thảo Vân (Anh)",
    ("Lài", "A"): "Cô Nguyễn Thị Lài (Anh)",
    # --- CÁC TRƯỜNG HỢP ĐẢO NGƯỢC ---
    ("V", "Việt"): "Thầy Lê Công Quốc Việt (Văn)",
    ("V", "Nhung"): "Cô Trần Thị Tuyết Nhung (Văn)",
    ("T", "Nghĩa"): "Thầy Nghĩa (Toán)",
    ("L", "Vân"): "Cô Nguyễn Thị Vân (Lý)",
}

def get_standard_name(p1, p2):
    """p1 và p2 là 2 phần tách ra từ dấu '-' """
    # Thử trường hợp 1: (Tên, Môn)
    if (p1, p2) in TU_DIEN:
        return TU_DIEN[(p1, p2)]
    # Thử trường hợp 2: (Môn, Tên)
    if (p2, p1) in TU_DIEN:
        return TU_DIEN[(p2, p1)]
    # Nếu không có trong từ điển, trả về dạng Tên (Môn) dựa trên logic thông thường
    # Giả định phần ngắn hơn là môn
    if len(p1) < len(p2):
        return f"{p2} ({p1})"
    return f"{p1} ({p2})"

def phan_loai_khoi(ten_lop):
    ten_lop = str(ten_lop).strip()
    if ten_lop.startswith(('10', '11', '12')): return 'THPT'
    elif ten_lop.startswith(('6', '7', '8', '9')): return 'THCS'
    return 'Khác'

# =====================================================================
# 2. XỬ LÝ DỮ LIỆU
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
                ten_goc = gvcn.get(class_name, 'Unknown')
                records.append({'Khối': khoi, 'Giáo viên': f"{ten_goc} (GVCN)", 'Lớp': class_name, 'Môn': 'Chủ nhiệm'})
            elif '-' in cell_val:
                parts = [p.strip() for p in cell_val.split('-')]
                records.append({
                    'Khối': khoi,
                    'Giáo viên': get_standard_name(parts[0], parts[1]),
                    'Lớp': class_name,
                    'Môn': parts[0] if len(parts[0]) < len(parts[1]) else parts[1]
                })
            else:
                records.append({'Khối': khoi, 'Giáo viên': 'Chung (Không ghi tên)', 'Lớp': class_name, 'Môn': cell_val})

    df_res = pd.DataFrame(records)
    if df_res.empty: return None
    df_summary = df_res.groupby(['Khối', 'Giáo viên', 'Lớp', 'Môn']).size().reset_index(name='Số tiết')
    return df_summary

# =====================================================================
# 3. GIAO DIỆN
# =====================================================================
st.set_page_config(page_title="Thống kê TKB", layout="wide")
st.title("📊 Công cụ Xử lý & Thống kê TKB")

if 'df_data' not in st.session_state:
    st.session_state.df_data = None

uploaded_file = st.file_uploader("Chọn file TKB (.xlsx)", type=["xlsx"])

if uploaded_file:
    if st.button("Bắt đầu phân tích dữ liệu", type="primary"):
        with st.spinner('Đang xử lý...'):
            res = process_tkb_data(uploaded_file)
            if res is not None:
                st.session_state.df_data = res
                st.success("Phân tích thành công!")

if st.session_state.df_data is not None:
    df = st.session_state.df_data
    st.divider()
    
    # Bộ lọc
    c1, c2, c3 = st.columns(3)
    khoi_chon = c1.multiselect("📚 Khối:", sorted(df['Khối'].unique()))
    gv_chon = c2.multiselect("👩‍🏫 Giáo viên:", sorted(df['Giáo viên'].unique()))
    lop_chon = c3.multiselect("🏫 Lớp:", sorted(df['Lớp'].unique()))

    # Áp dụng lọc
    df_f = df.copy()
    if khoi_chon: df_f = df_f[df_f['Khối'].isin(khoi_chon)]
    if gv_chon: df_f = df_f[df_f['Giáo viên'].isin(gv_chon)]
    if lop_chon: df_f = df_f[df_f['Lớp'].isin(lop_chon)]

    # HIỂN THỊ KẾT QUẢ
    col_summary, col_details = st.columns([1, 2])

    with col_summary:
        st.subheader("📈 Tổng tiết theo GV")
        gv_sum = df_f.groupby('Giáo viên')['Số tiết'].sum().reset_index().sort_values('Số tiết', ascending=False)
        st.dataframe(gv_sum, hide_index=True, use_container_width=True)
        st.metric("Tổng số tiết", f"{gv_sum['Số tiết'].sum()} tiết")

    with col_details:
        st.subheader("📋 Chi tiết")
        st.dataframe(df_f, hide_index=True, use_container_width=True)

    # Export
    csv = df_f.to_csv(index=False, encoding='utf-8-sig').encode('utf-8-sig')
    st.download_button("📥 Tải CSV kết quả", data=csv, file_name="ThongKe_TKB.csv", mime="text/csv")
