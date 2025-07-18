import streamlit as st
import pandas as pd
import sys
import os
from pathlib import Path

# Định nghĩa base path (hỗ trợ khi frozen vs chạy script)
def get_base_path():
    if getattr(sys, 'frozen', False):
        return sys._MEIPASS
    return os.path.dirname(os.path.abspath(__file__))

BASE_PATH = get_base_path()
UPLOAD_DIR = os.path.join(BASE_PATH, 'uploaded_files')
os.makedirs(UPLOAD_DIR, exist_ok=True)

@st.cache_data
def load_data(path, sheet_name, header_row=1):
    """
    Đọc file Excel tại path và sheet, bỏ qua header_row đầu,
    trả về DataFrame với tên cột được chuẩn hóa.
    """
    df = pd.read_excel(path, sheet_name=sheet_name, header=header_row)
    df.columns = [str(col).strip() for col in df.columns]
    return df

# Lấy danh sách file đã upload và file mẫu
def get_files():
    files = {}
    default_path = os.path.join(BASE_PATH, 'TestSearch.xlsx')
    if os.path.exists(default_path):
        files['TestSearch.xlsx'] = default_path
    for f in Path(UPLOAD_DIR).iterdir():
        if f.suffix.lower() in ['.xlsx', '.xls']:
            files[f.name] = str(f)
    return files

# Hàm main
def main():
    # Cấu hình layout rộng
    st.set_page_config(layout='wide')
    st.title('🔍 Ứng dụng Tìm Kiếm Nhanh Trong Sheets')
    st.markdown('---')

    # 1) Upload file mới
    st.header('1. Upload File')
    uploaded = st.file_uploader(
        label='Chọn file Excel (xlsx hoặc xls)',
        type=['xlsx', 'xls'],
        accept_multiple_files=True,
        key='upload'
    )
    if uploaded:
        saved_any = False
        for up in uploaded:
            existing = get_files()
            if up.name in existing:
                st.warning(f"File '{up.name}' đã tồn tại, bỏ qua.")
            else:
                dst = os.path.join(UPLOAD_DIR, up.name)
                with open(dst, 'wb') as out:
                    out.write(up.getbuffer())
                st.success(f"Đã lưu file '{up.name}'.")
                saved_any = True
        if saved_any:
            st.cache_data.clear()
            return

    # 2) Chọn file để tìm kiếm
    st.header('2. Chọn File')
    files = get_files()
    if not files:
        st.error('Chưa có file nào. Vui lòng upload file.')
        return
    selected_file = st.selectbox('Chọn file', list(files.keys()), key='file_sel')

    # Lấy danh sách sheet trong file
    try:
        xl = pd.ExcelFile(files[selected_file])
        sheets = xl.sheet_names
    except Exception as e:
        st.error(f"Không đọc được file: {e}")
        return

    # 3) Chọn chế độ tìm kiếm
    st.header('3. Chế độ Tìm kiếm')
    mode = st.radio(
        label='Chọn chế độ',
        options=['Một sheet', 'Tất cả sheets'],
        key='mode'
    )
    if mode == 'Một sheet':
        sheet = st.selectbox('Chọn sheet', sheets, key='sheet_sel')
    else:
        sheet = None

    # 4) Nhập điều kiện tìm kiếm và hiển thị
    if mode == 'Tất cả sheets':
        st.header('4. Tìm kiếm chung (All Sheets)')
        with st.form(key='search_all_form'):
            query = st.text_input('Nhập từ khóa chung', key='query')
            submit_all = st.form_submit_button('Tìm')
        if submit_all and query:
            results = []
            for sh in sheets:
                df = load_data(files[selected_file], sh)
                mask = df.astype(str).apply(lambda col: col.str.contains(query, case=False, na=False))
                matched = df[mask.any(axis=1)]
                if not matched.empty:
                    matched.insert(0, 'Sheet', sh)
                    results.append(matched)
            st.markdown('---')
            st.header('Kết quả Chung')
            if results:
                total = sum(len(df_) for df_ in results)
                st.success(f'Tìm thấy {total} kết quả.')
                for df_res in results:
                    st.dataframe(df_res, use_container_width=True)
            else:
                st.error('Không tìm thấy kết quả phù hợp.')
    else:
        # 4a) Tìm trong 1 sheet: filters in sidebar, results in main
        df0 = load_data(files[selected_file], sheet)
        st.sidebar.header(f'Lọc theo cột (Sheet: {sheet})')
        with st.sidebar.form(key='filter_form'):
            filters = {}
            for col in df0.columns:
                filters[col] = st.text_input(label=col, key=f'filter_{col}')
            submit = st.form_submit_button('Tìm (Enter)')
        if submit:
            filtered = df0.copy()
            for c, v in filters.items():
                if v:
                    filtered = filtered[filtered[c].astype(str).str.contains(v, case=False, na=False)]
            st.markdown('---')
            st.header(f'Kết quả Sheet: {sheet}')
            if not filtered.empty:
                filtered.insert(0, 'Sheet', sheet)
                st.dataframe(filtered, use_container_width=True)
            else:
                st.error('Không tìm thấy kết quả phù hợp.')

if __name__ == '__main__':
    main()
