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
    layout="wide",
    initial_sidebar_state="collapsed"
)

# ===== CUSTOM CSS =====
st.markdown("""
    <style>
    /* ===== BASE ===== */
    .main-title {
        font-size: 2.8rem;
        font-weight: 800;
        background: linear-gradient(90deg, #4e54c8, #8f94fb);
        -webkit-background-clip: text;
        -webkit-text-fill-color: transparent;
        text-align: center;
        margin: 0.2rem 0 0.5rem;
        letter-spacing: -0.5px;
    }
    .subtitle {
        text-align: center;
        color: #666;
        font-size: 1.1rem;
        margin-bottom: 2rem;
        font-weight: 500;
    }
    .made-by {
        text-align: center;
        color: #888;
        font-size: 0.9rem;
        margin-top: -0.5rem;
        font-style: italic;
    }

    /* ===== CARDS ===== */
    .upload-card, .preview-card, .action-card {
        background: white;
        border-radius: 16px;
        padding: 1.5rem;
        box-shadow: 0 4px 20px rgba(0,0,0,0.08);
        border: 1px solid #eee;
        transition: all 0.3s ease;
        margin-bottom: 1.5rem;
    }
    .upload-card:hover, .preview-card:hover, .action-card:hover {
        transform: translateY(-2px);
        box-shadow: 0 6px 25px rgba(0,0,0,0.12);
    }

    /* ===== BUTTONS ===== */
    .stButton>button {
        width: 100%;
        height: 3.2rem;
        border-radius: 12px;
        font-weight: 600;
        font-size: 1.1rem;
        transition: all 0.2s ease;
        border: none;
        margin-top: 0.5rem;
    }
    .primary-btn {
        background: linear-gradient(90deg, #4e54c8, #8f94fb) !important;
        color: white !important;
    }
    .primary-btn:hover {
        background: linear-gradient(90deg, #3a40b0, #7a83e0) !important;
        transform: translateY(-1px);
        box-shadow: 0 4px 15px rgba(78, 84, 200, 0.4);
    }
    .download-btn {
        background: linear-gradient(90deg, #28a745, #3fd16d) !important;
        color: white !important;
    }
    .download-btn:hover {
        background: linear-gradient(90deg, #218838, #36b65d) !important;
        box-shadow: 0 4px 15px rgba(40, 167, 69, 0.4);
    }

    /* ===== METRICS ===== */
    .metric-card {
        background: #f8f9ff;
        padding: 1rem;
        border-radius: 12px;
        text-align: center;
        border: 1px solid #eef2ff;
    }
    .metric-value {
        font-size: 2rem;
        font-weight: 700;
        color: #4e54c8;
        margin: 0.5rem 0;
    }
    .metric-label {
        color: #666;
        font-size: 0.9rem;
        font-weight: 500;
    }

    /* ===== PROGRESS ===== */
    .stProgress > div > div > div > div {
        background-color: #4e54c8;
    }

    /* ===== FOOTER ===== */
    .footer {
        text-align: center;
        padding: 2rem 0 1rem;
        color: #888;
        font-size: 0.9rem;
        border-top: 1px solid #eee;
        margin-top: 3rem;
    }
    .footer a {
        color: #4e54c8;
        text-decoration: none;
    }
    .footer a:hover {
        text-decoration: underline;
    }

    /* ===== EXPANDER ===== */
    .streamlit-expanderHeader {
        background-color: #f8f9ff !important;
        border-radius: 8px !important;
        font-weight: 600;
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

        # Copy headers
        for i, cell in enumerate(ws[1], start=1):
            new_cell = new_ws.cell(row=1, column=i)
            copy_cell(cell, new_cell)
            if cell.column_letter in ws.column_dimensions:
                new_ws.column_dimensions[new_cell.column_letter].width = ws.column_dimensions[cell.column_letter].width

        # Copy data rows
        for r_idx, row in enumerate(rows, start=2):
            for c_idx, cell in enumerate(row, start=1):
                new_cell = new_ws.cell(row=r_idx, column=c_idx)
                copy_cell(cell, new_cell)

        # Copy merged cells
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

# ===== APP HEADER =====
st.markdown('<h1 class="main-title">📊 ExcelSplit Pro</h1>', unsafe_allow_html=True)
st.markdown('<p class="subtitle">Split Excel files intelligently — preserve formatting, preview results, download instantly</p>', unsafe_allow_html=True)
st.markdown('<p class="made-by">Made with ❤️ by Pravedra Singh Rawat</p>', unsafe_allow_html=True)

# ===== INIT STATE =====
if "file_processed" not in st.session_state:
    st.session_state.file_processed = False
    st.session_state.group_count = 0
    st.session_state.selected_column = ""

# ===== STEP 1: UPLOAD =====
col1, col2 = st.columns([2, 1])

with col1:
    st.markdown('<div class="upload-card">', unsafe_allow_html=True)
    st.subheader("📤 Upload Excel File")
    st.info("Drag & drop your `.xlsx` file below. All styles, colors, and formats will be preserved!")
    uploaded_file = st.file_uploader("", type=["xlsx"], label_visibility="collapsed", key="uploader")
    st.markdown('</div>', unsafe_allow_html=True)

if not uploaded_file:
    st.stop()

# Load workbook
with tempfile.NamedTemporaryFile(delete=False, suffix=".xlsx") as tmp:
    tmp.write(uploaded_file.getvalue())
    tmp_path = tmp.name

try:
    wb = openpyxl.load_workbook(tmp_path)
    ws = wb.active
    headers = [cell.value for cell in next(ws.iter_rows(min_row=1, max_row=1))]

    # ===== STEP 2: COLUMN SELECTION + PREVIEW STATS =====
    col1, col2 = st.columns([2, 1])

    with col1:
        st.markdown('<div class="preview-card">', unsafe_allow_html=True)
        st.subheader("⚙️ Configure Split")
        selected_column = st.selectbox(
            "Select column to split by:",
            headers,
            help="Each unique value becomes a separate file"
        )

        # Preview how many files will be created
        col_idx = headers.index(selected_column) + 1
        unique_values = set()
        for row in ws.iter_rows(min_row=2):
            value = row[col_idx - 1].value
            key = str(value) if value is not None else "Unknown"
            key = "".join(c if c not in '<>:"/\\|?*' else "_" for c in key)
            unique_values.add(key)

        st.session_state.group_count = len(unique_values)
        st.session_state.selected_column = selected_column

        st.markdown('</div>', unsafe_allow_html=True)

    with col2:
        st.markdown('<div class="metric-card">', unsafe_allow_html=True)
        st.markdown('<div class="metric-label">Files to be created</div>', unsafe_allow_html=True)
        st.markdown(f'<div class="metric-value">{st.session_state.group_count}</div>', unsafe_allow_html=True)
        st.markdown('<div class="metric-label">Based on column</div>', unsafe_allow_html=True)
        st.code(st.session_state.selected_column)
        st.markdown('</div>', unsafe_allow_html=True)

    # ===== STEP 3: DATA PREVIEW =====
    with st.expander("🔍 Data Preview (First 5 Rows)", expanded=False):
        df = pd.read_excel(uploaded_file)
        st.dataframe(df.head(), use_container_width=True)

    # ===== STEP 4: SPLIT BUTTON =====
    st.markdown('<div class="action-card">', unsafe_allow_html=True)
    if st.button("🚀 Generate Split Files", key="split_btn", help="Click to start splitting your Excel file"):
        with st.spinner("Processing... Preserving styles, widths, and formatting"):
            progress_bar = st.progress(0)
            status_text = st.empty()

            with tempfile.TemporaryDirectory() as tmpdir:
                saved_files, error = split_excel_by_column(ws, headers, selected_column, tmpdir)

                if error:
                    st.error(error)
                elif saved_files:
                    progress_bar.progress(50)
                    status_text.text("📦 Packaging files into ZIP...")

                    zip_path = os.path.join(tempfile.gettempdir(), "ExcelSplit_Pro_Output.zip")
                    create_zip(saved_files, zip_path)

                    progress_bar.progress(100)
                    status_text.text("✅ Done! Ready to download.")

                    st.success(f"🎉 Success! Created **{len(saved_files)}** beautifully formatted Excel files.")

                    with open(zip_path, "rb") as f:
                        st.download_button(
                            label="📥 Download All Files (ZIP)",
                            data=f,
                            file_name=f"ExcelSplit_Output_{datetime.now().strftime('%Y%m%d_%H%M')}.zip",
                            mime="application/zip",
                            key="download_btn"
                        )

                    # Show sample filenames
                    with st.expander("📁 Sample Output Files", expanded=True):
                        for fname in saved_files[:5]:
                            st.code("📄 " + os.path.basename(fname))
                        if len(saved_files) > 5:
                            st.caption(f"... and {len(saved_files) - 5} more files")

                    st.session_state.file_processed = True

            progress_bar.empty()
            status_text.empty()

    st.markdown('</div>', unsafe_allow_html=True)

except Exception as e:
    st.error(f"⚠️ Error: {str(e)}")
    st.info("Tip: Make sure your file is a valid .xlsx and not corrupted.")

finally:
    if 'tmp_path' in locals():
        os.unlink(tmp_path)

# ===== FOOTER =====
st.markdown("""
    <div class="footer">
        <p>ExcelSplit Pro v1.0 • Made with Python, Streamlit & ❤️ by <strong>Pravedra Singh Rawat</strong></p>
        <p>Preserves all formatting • No data leaves your computer • 100% client-side processing</p>
    </div>
""", unsafe_allow_html=True)
