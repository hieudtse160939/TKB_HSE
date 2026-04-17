import pandas as pd
import openpyxl
import streamlit as st
import io
import re

# =====================================================================
# 1. CẤU HÌNH & TỪ ĐIỂN ƯU TIÊN (Dành cho các trường hợp đặc biệt)
# =====================================================================
TU_DIEN_UU_TIEN = {
    ("Vân", "L"): "Cô Vân (Lý)",
    ("Vân", "CĐ Lý"): "Cô Vân (Lý)",
    ("Vân", "Đ"): "Cô Vân (Địa)",
    ("Vân", "A"): "Cô Thảo Vân (Anh)",
    ("Nhung", "V"): "Cô T.Nhung (Văn)",
    ("Nhung", "CĐ Văn"): "Cô T.Nhung (Văn)",
    ("Nghĩa", "T"): "Thầy Nghĩa (Toán)",
    ("nghĩa", "CĐ Toán"): "Thầy Nghĩa (Toán)",
}

# =====================================================================
# 2. HÀM HỖ TRỢ TRA CỨU TÊN GIÁO VIÊN
# =====================================================================
def get_standard_name(ten_tat, mon_hoc, teacher_dict):
    """
    Quy trình tra cứu:
    1. Kiểm tra từ điển ưu tiên.
    2. Tra cứu trong file DSGVBM dựa trên Tên và Môn.
    3. Nếu không thấy, trả về định dạng mặc định.
    """
    # 1. Ưu tiên từ điển viết tay
    if (ten_tat, mon_hoc) in TU_DIEN_UU_TIEN:
        return TU_DIEN_UU_TIEN[(ten_tat, mon_hoc)]
    
    # 2. Chuẩn hóa tên môn để khớp với file DSGVBM
    # Ví dụ: "CĐ Văn" -> "Văn", "T" -> "Toán"
    mon_sach = mon_hoc.replace("CĐ ", "").strip()
    if mon_sach == "V": mon_sach = "Văn"
    if mon_sach == "T": mon_sach = "Toán"
    if mon_sach == "L": mon_sach = "Lý"
    if mon_sach == "H": mon_sach = "Hóa"
    if mon_sach == "A": mon_sach = "Anh"

    # 3. Tra cứu trong dictionary được tạo từ file DSGVBM
    # Khớp theo Tên (chữ cuối) và Môn (có chứa từ khóa môn học)
    key = (ten_tat.lower(), mon_sach.lower())
    if key in teacher_dict:
        return teacher_dict[key]
    
    # 4. Mặc định nếu không tìm thấy
    return f"{ten_tat} ({mon_hoc})"

def phan_loai_khoi(ten_lop):
    ten_lop = str(ten_lop).strip()
    if ten_lop.startswith(('10', '11', '12')): return 'THPT'
    if ten_lop.startswith(('6', '7', '8', '9')): return 'THCS'
    return 'Khác'

# =====================================================================
# 3. XỬ LÝ FILE TKB
# =====================================================================
def process_tkb_data(uploaded_tkb, teacher_dict):
    wb = openpyxl.load_workbook(uploaded_tkb)
    
    # Unmerge tất cả các cell
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
                    
        # Giới hạn dòng xử lý (thường TKB không quá 66 dòng)
        if ws.max_row > 66:
            ws.delete_rows(67, ws.max_row - 66)
            
    virtual_workbook = io.BytesIO()
    wb.save(virtual_workbook)
    virtual_workbook.seek(0)
    df = pd.read_excel(virtual_workbook) 
    
    # Xác định vị trí lớp và GVCN (Dòng 4 và 5 trong Excel tương ứng idx 3, 4)
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
    # Bắt đầu duyệt từ dòng tiết học (idx 5)
    for row_idx in range(5, len(df)):
        row = df.iloc[row_idx]
        for col_idx, class_name in classes.items():
            cell_val = str(row.iloc[col_idx]).strip()
            if cell_val in ['nan', '', 'CHÀO CỜ', 'SINH HOẠT ĐẦU GIỜ', 'THỂ DỤC THỂ THAO']:
                continue
            
            khoi = phan_loai_khoi(class_name)

            # Trường hợp tiết Chủ nhiệm
            if cell_val.lower() == 'chủ nhiệm':
                ten_goc = gvcn.get(class_name, 'Unknown')
                records.append({
                    'Khối': khoi,
                    'Giáo viên': f"{ten_goc} (GVCN)",
                    'Lớp': class_name, 'Môn': 'Chủ nhiệm'
                })
            # Trường hợp tiết có gạch nối (Môn - TênGV) hoặc (CĐ Môn - TênGV)
            elif '-' in cell_val:
                parts = cell_val.split('-')
                mon_hoc = "-".join(parts[:-1]).strip() 
                ten_tat = parts[-1].strip()        
                records.append({
                    'Khối': khoi,
                    'Giáo viên': get_standard_name(ten_tat, mon_hoc, teacher_dict),
                    'Lớp': class_name, 'Môn': mon_hoc
                })
            else:
                records.append({
                    'Khối': khoi, 'Giáo viên': 'Chung (Không tên)',
                    'Lớp': class_name, 'Môn': cell_val
                })

    return pd.DataFrame(records)

# =====================================================================
# 4. GIAO DIỆN STREAMLIT
# =====================================================================
st.set_page_config(page_title="Hoa Sen TKB Tool", layout="wide")
st.title("📊 Công cụ Thống kê Tiết dạy Hoa Sen")

# --- BƯỚC 1: TẢI DANH SÁCH GIÁO VIÊN ---
st.sidebar.header("1. Cài đặt Danh bạ")
file_dsgv = st.sidebar.file_uploader("Tải file DSGVBM (.xlsx hoặc .csv)", type=["xlsx", "csv"])

teacher_dict = {}
if file_dsgv:
    try:
        if file_dsgv.name.endswith('csv'):
            df_ds = pd.read_csv(file_dsgv)
        else:
            df_ds = pd.read_excel(file_dsgv)
        
        # Tạo từ điển tra cứu nhanh
        for _, r in df_ds.iterrows():
            ho_ten = str(r['Họ tên GV']).strip()
            mon = str(r['GVBM']).strip()
            ten_rieng = ho_ten.split()[-1].lower()
            # Key: (tên, môn_viết_tắt) -> Value: Họ tên đầy đủ (Môn)
            teacher_dict[(ten_rieng, mon.lower())] = f"{ho_ten} ({mon})"
        st.sidebar.success(f"Đã nạp {len(teacher_dict)} giáo viên!")
    except Exception as e:
        st.sidebar.error(f"Lỗi file danh sách: {e}")

# --- BƯỚC 2: TẢI TKB VÀ XỬ LÝ ---
st.subheader("2. Phân tích Thời khóa biểu")
file_tkb = st.file_uploader("Tải file TKB (.xlsx)", type=["xlsx"])

if file_tkb and teacher_dict:
    if st.button("Bắt đầu phân tích", type="primary"):
        df_raw = process_tkb_data(file_tkb, teacher_dict)
        
        if not df_raw.empty:
            # Lưu vào session state
            st.session_state.data = df_raw.groupby(['Khối', 'Giáo viên', 'Lớp', 'Môn']).size().reset_index(name='Số tiết')
            st.success("Phân tích hoàn tất!")

# --- BƯỚC 3: HIỂN THỊ KẾT QUẢ ---
if 'data' in st.session_state:
    df = st.session_state.data
    
    # Bộ lọc
    st.divider()
    c1, c2 = st.columns(2)
    with c1:
        gv_filter = st.multiselect("Chọn Giáo viên:", sorted(df['Giáo viên'].unique()))
    with c2:
        khoi_filter = st.multiselect("Chọn Khối:", sorted(df['Khối'].unique()))
    
    df_f = df.copy()
    if gv_filter: df_f = df_f[df_f['Giáo viên'].isin(gv_filter)]
    if khoi_filter: df_f = df_f[df_f['Khối'].isin(khoi_filter)]

    # Hiển thị bảng tổng hợp và chi tiết
    col_left, col_right = st.columns([1, 2])
    
    with col_left:
        st.subheader("📈 Tổng số tiết/GV")
        summary = df_f.groupby('Giáo viên')['Số tiết'].sum().reset_index().sort_values('Số tiết', ascending=False)
        st.dataframe(summary, hide_index=True, use_container_width=True)
        st.metric("Tổng cộng", f"{summary['Số tiết'].sum()} tiết")

    with col_right:
        st.subheader("📋 Chi tiết phân bổ")
        st.dataframe(df_f, hide_index=True, use_container_width=True)

    # Xuất file
    csv = df_f.to_csv(index=False, encoding='utf-8-sig').encode('utf-8-sig')
    st.download_button("📥 Tải kết quả (CSV)", data=csv, file_name="ThongKe_TKB.csv", mime="text/csv")
else:
    if not teacher_dict:
        st.info("Vui lòng tải file Danh sách GV (DSGVBM) ở thanh bên trái trước.")
