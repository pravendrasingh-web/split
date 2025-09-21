import streamlit as st
import openpyxl
from openpyxl import Workbook
import os
import tempfile
import zipfile
import pandas as pd
from datetime import datetime

# ===== PAGE CONFIG =====
st.set_page_config(
    page_title="ExcelSplit Pro",
    page_icon="📊",
    layout="centered",
    initial_sidebar_state="collapsed"
)

# ===== CUSTOM CSS =====
st.markdown("""
    <style>
    body {
        font-family: "Inter", sans-serif;
        background: #f9fafe;
    }
    .main-title {
        font-size: 2.8rem;
        font-weight: 800;
        background: linear-gradient(90deg, #4e54c8, #8f94fb);
        -webkit-background-clip: text;
        -webkit-text-fill-color: transparent;
        text-align: center;
        margin: 0.8rem 0 0.5rem;
        letter-spacing: -0.5px;
        text-shadow: 0 2px 8px rgba(78,84,200,0.2);
    }
    .subtitle {
        text-align: center;
        color: #555;
        font-size: 1.2rem;
        margin-bottom: 1rem;
        font-weight: 500;
    }
    .made-by {
        text-align: center;
        color: #777;
        font-size: 0.95rem;
        margin-top: -0.3rem;
        font-style: italic;
    }
    .card {
        background: rgba(255, 255, 255, 0.85);
        border-radius: 16px;
        padding: 1.5rem;
        box-shadow: 0 8px 24px rgba(0,0,0,0.06);
        border: 1px solid rgba(230,230,250,0.6);
        backdrop-filter: blur(12px);
        margin: 1.2rem 0;
        transition: all 0.25s ease;
    }
    .card:hover {
        transform: translateY(-4px);
        box-shadow: 0 10px 28px rgba(0,0,0,0.08);
    }
    .stButton>button {
        width: 100%;
        height: 3.3rem;
        border-radius: 14px !important;
        font-weight: 600 !important;
        font-size: 1.1rem !important;
        border: none;
        cursor: pointer;
        transition: all 0.25s ease;
    }
    .stButton>button[kind="primary"] {
        background: linear-gradient(90deg, #4e54c8, #8f94fb);
        color: white;
    }
    .stButton>button[kind="primary"]:hover {
        background: linear-gradient(90deg, #3c40b5, #737be0);
        transform: translateY(-2px);
        box-shadow: 0 6px 16px rgba(78,84,200,0.35);
    }
    .metric-box {
        background: #f4f6ff;
        border: 1px solid #e0e3ff;
        border-radius: 14px;
        padding: 1rem;
        text-align: center;
        margin: 0.8rem 0;
    }
    .metric-value {
        font-size: 2.2rem;
        font-weight: 800;
        color: #4e54c8;
        margin: 0;
    }
    .metric-label {
        color: #666;
        font-size: 0.95rem;
        font-weight: 500;
    }
    .dataframe {
        border: 1px solid #eee;
        border-radius: 12px;
        overflow-y: auto;
        max-height: 350px;
        margin-top: 0.8rem;
    }
    .footer {
        text-align: center;
        padding: 1.5rem 0;
        color: #777;
        font-size: 0.95rem;
        border-top: 1px solid #eaeaea;
        margin-top: 2rem;
        background: linear-gradient(90deg, #fafbff, #f6f8ff);
    }
    .footer a {
        color: #4e54c8;
        text-decoration: none;
    }
    .footer a:hover {
        text-decoration: underline;
    }
    </style>
""", unsafe_allow_html=True)

# ===== HELPER: Copy cell with style =====
def copy_cell(source_cell, target_cell):
    target_cell.value = source_cell.value
    if source_cell.has_style:
        target_cell.font = source_cell.font.copy()
        target_cell.border = source_cell.border.copy()
        target_cell.fill = source_cell.fill.copy()
        target_cell.number_format = source_cell.number_format
        target_cell.protection = source_cell.protection.copy()
        target_cell.alignment = source_cell.alignment.copy()

# ===== FUNCTION: Split Excel by column =====
def split_excel_by_column(ws, headers, column_name, output_dir):
    if column_name not in headers:
        return None, f"❌ Column '{column_name}' not found!"

    col_idx = headers.index(column_name) + 1
    groups = {}
    for row in ws.iter_rows(min_row=2):
        value = row[col_idx - 1].value
        key = str(value) if value is not None else "Unknown"
        key = "".join(c if c not in '<>:"/\\|?*' else "_" for c in key)
        if key not in groups:
            groups[key] = []
        groups[key].append(row)

    saved_files = []
    for value, rows in groups.items():
        new_wb = Workbook()
        new_ws = new_wb.active
        new_ws.title = ws.title
        for i, cell in enumerate(ws[1], start=1):
            new_cell = new_ws.cell(row=1, column=i)
            copy_cell(cell, new_cell)
            if cell.column_letter in ws.column_dimensions:
                new_ws.column_dimensions[new_cell.column_letter].width = ws.column_dimensions[cell.column_letter].width
        for r_idx, row in enumerate(rows, start=2):
            for c_idx, cell in enumerate(row, start=1):
                new_cell = new_ws.cell(row=r_idx, column=c_idx)
                copy_cell(cell, new_cell)
        for merged_range in ws.merged_cells.ranges:
            new_ws.merge_cells(str(merged_range))
        file_name = f"{value}.xlsx"
        save_path = os.path.join(output_dir, file_name)
        new_wb.save(save_path)
        saved_files.append(save_path)
    return saved_files, None

# ===== FUNCTION: Create ZIP =====
def create_zip(files, zip_name):
    with zipfile.ZipFile(zip_name, 'w') as zf:
        for f in files:
            zf.write(f, os.path.basename(f))
    return zip_name

# ===== UI =====
st.markdown('<h1 class="main-title">📊 ExcelSplit Pro</h1>', unsafe_allow_html=True)
st.markdown('<p class="subtitle">Split Excel files intelligently — preserve formatting, preview results, download instantly</p>', unsafe_allow_html=True)
st.markdown('<p class="made-by">Made with ❤️ by Pravendra Singh Rawat</p>', unsafe_allow_html=True)

with st.container():
    st.markdown('<div class="card">', unsafe_allow_html=True)
    st.subheader("📤 Upload Excel File")
    st.info("Drag & drop your `.xlsx` file below. All styles, colors, and formats will be preserved!")
    uploaded_file = st.file_uploader("", type=["xlsx"], label_visibility="collapsed", key="uploader")
    if uploaded_file:
        st.success(f"✅ File loaded: `{uploaded_file.name}`")
    st.markdown('</div>', unsafe_allow_html=True)

if not uploaded_file:
    st.stop()

with tempfile.NamedTemporaryFile(delete=False, suffix=".xlsx") as tmp:
    tmp.write(uploaded_file.getvalue())
    tmp_path = tmp.name

try:
    wb = openpyxl.load_workbook(tmp_path)
    ws = wb.active
    headers = [cell.value for cell in next(ws.iter_rows(min_row=1, max_row=1))]
    col1, col2 = st.columns([3, 1])
    with col1:
        st.markdown('<div class="card">', unsafe_allow_html=True)
        st.subheader("⚙️ Configure Split")
        selected_column = st.selectbox("Select column to split by:", headers)
        col_idx = headers.index(selected_column) + 1
        unique_values = set()
        for row in ws.iter_rows(min_row=2):
            value = row[col_idx - 1].value
            key = str(value) if value is not None else "Unknown"
            key = "".join(c if c not in '<>:"/\\|?*' else "_" for c in key)
            unique_values.add(key)
        group_count = len(unique_values)
        st.markdown('<div class="metric-box">', unsafe_allow_html=True)
        st.markdown(f'<span class="metric-value">{group_count}</span>', unsafe_allow_html=True)
        st.markdown('<span class="metric-label">Files to be created</span>', unsafe_allow_html=True)
        st.markdown('</div>', unsafe_allow_html=True)
        st.markdown('<div class="metric-box">', unsafe_allow_html=True)
        st.markdown(f'<span class="metric-label">Based on column</span>', unsafe_allow_html=True)
        st.code(selected_column)
        st.markdown('</div>', unsafe_allow_html=True)
        st.markdown('</div>', unsafe_allow_html=True)
    with col2:
        st.markdown('<div class="card">', unsafe_allow_html=True)
        st.markdown('<div class="metric-box">', unsafe_allow_html=True)
        st.markdown(f'<span class="metric-value">{group_count}</span>', unsafe_allow_html=True)
        st.markdown('<span class="metric-label">Files to be created</span>', unsafe_allow_html=True)
        st.markdown('</div>', unsafe_allow_html=True)
        st.markdown('<div class="metric-box">', unsafe_allow_html=True)
        st.markdown('<span class="metric-label">Based on column</span>', unsafe_allow_html=True)
        st.code(selected_column)
        st.markdown('</div>', unsafe_allow_html=True)
        st.markdown('</div>', unsafe_allow_html=True)
    with st.expander("🔍 Data Preview (First 5 Rows)"):
        df = pd.read_excel(uploaded_file)
        st.dataframe(df.head(), use_container_width=True)
    st.markdown('<div class="card">', unsafe_allow_html=True)
    if st.button("🚀 Generate Split Files", key="split_btn"):
        with st.spinner("Processing... Preserving styles, widths, and formatting"):
            with tempfile.TemporaryDirectory() as tmpdir:
                saved_files, error = split_excel_by_column(ws, headers, selected_column, tmpdir)
                if error:
                    st.error(error)
                elif saved_files:
                    zip_path = os.path.join(tempfile.gettempdir(), "ExcelSplit_Pro_Output.zip")
                    create_zip(saved_files, zip_path)
                    st.success(f"🎉 Success! Created **{len(saved_files)}** beautifully formatted Excel files.")
                    with open(zip_path, "rb") as f:
                        st.download_button(
                            label="📥 Download All Files (ZIP)",
                            data=f,
                            file_name=f"ExcelSplit_Output_{datetime.now().strftime('%Y%m%d_%H%M')}.zip",
                            mime="application/zip",
                            key="download_btn"
                        )
                    with st.expander("📁 Sample Output Files", expanded=True):
                        for fname in saved_files[:5]:
                            st.code("📄 " + os.path.basename(fname))
                        if len(saved_files) > 5:
                            st.caption(f"... and {len(saved_files) - 5} more files")
    st.markdown('</div>', unsafe_allow_html=True)
except Exception as e:
    st.error(f"⚠️ Error: {str(e)}")
    st.info("Tip: Make sure your file is a valid .xlsx and not corrupted.")
finally:
    if 'tmp_path' in locals():
        os.unlink(tmp_path)

st.markdown("""
    <div class="footer">
        <p>ExcelSplit Pro v1.0 • Made with Python, Streamlit & ❤️ by <strong>Pravendra Singh Rawat</strong></p>
        <p>Preserves all formatting • No data leaves your computer • Secure local processing</p>
    </div>
""", unsafe_allow_html=True)
