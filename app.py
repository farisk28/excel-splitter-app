import streamlit as st
import pandas as pd
import openpyxl
from openpyxl.drawing.image import Image
from openpyxl.styles import Border, Side, Font, PatternFill, Alignment
import io
import zipfile
import xml.etree.ElementTree as ET

# =====================================================================
# 1. FUNGSI FORMAT NOMINAL & PEMBERSIH TEKS
# =====================================================================
def format_transaction_amount(val):
    if val is None:
        return ""
    val_str = str(val).strip()
    if val_str == "#VALUE!" or not val_str:
        return ""
    try:
        clean_str = val_str.replace(',', '').replace('.', '')
        if val_str.endswith('.0'):
            val_str = val_str[:-2]
            clean_str = val_str.replace(',', '').replace('.', '')
        num = float(clean_str)
        if 0 < num < 1000:
            num = num * 1000
        return f"{int(num):,}"
    except ValueError:
        return val_str

# =====================================================================
# 2. FUNGSI STYLING TABEL & BORDER EXCEL
# =====================================================================
def apply_excel_styling(ws, headers, num_rows):
    thin_border = Border(
        left=Side(style='thin', color='D9D9D9'),
        right=Side(style='thin', color='D9D9D9'),
        top=Side(style='thin', color='D9D9D9'),
        bottom=Side(style='thin', color='D9D9D9')
    )
    header_fill = PatternFill(start_color="F2F2F2", end_color="F2F2F2", fill_type="solid")
    header_font = Font(name="Calibri", size=11, bold=True)
    
    # Format baris Header
    for col_idx in range(1, len(headers) + 1):
        cell = ws.cell(row=1, column=col_idx)
        cell.font = header_font
        cell.fill = header_fill
        cell.border = thin_border
        cell.alignment = Alignment(vertical="center", horizontal="center")

    # Format seluruh baris Data
    for row_idx in range(2, num_rows + 2):
        for col_idx in range(1, len(headers) + 1):
            cell = ws.cell(row=row_idx, column=col_idx)
            cell.border = thin_border
            cell.alignment = Alignment(vertical="center")

# =====================================================================
# 3. FUNGSI EKSTRAKSI GAMBAR PER SHEET
# =====================================================================
def extract_sheet_images(file_bytes, sheet_name):
    row_images_map = {}
    try:
        wb = openpyxl.load_workbook(io.BytesIO(file_bytes), data_only=True)
        if sheet_name in wb.sheetnames:
            ws = wb[sheet_name]
            if hasattr(ws, '_images'):
                for img in ws._images:
                    img_row = img.anchor._from.row + 1
                    img_col = img.anchor._from.col + 1
                    img_bytes = io.BytesIO(img._data())
                    if img_row not in row_images_map:
                        row_images_map[img_row] = []
                    row_images_map[img_row].append((img_col, img_bytes))
    except Exception:
        pass
    return row_images_map

# =====================================================================
# 4. FUNGSI MEMBACA MULTI-SHEET DARI SELURUH FILE
# =====================================================================
def process_multisheet_excel(uploaded_files):
    """
    Membaca seluruh sheet dari setiap file yang diunggah.
    Struktur kembalian:
    sheets_dict = {
        'sheet_name': {
            'headers': [...],
            'rows': [{'data': {...}, 'images': [...]}, ...]
        }
    }
    """
    sheets_dict = {}

    for file in uploaded_files:
        file_bytes = file.getvalue()
        wb = openpyxl.load_workbook(io.BytesIO(file_bytes), data_only=True)

        for s_name in wb.sheetnames:
            ws = wb[s_name]
            current_headers = [str(cell.value).strip() for cell in ws[1] if cell.value is not None]
            
            if not current_headers:
                continue

            if s_name not in sheets_dict:
                sheets_dict[s_name] = {
                    'headers': current_headers,
                    'rows': []
                }

            headers = sheets_dict[s_name]['headers']
            row_images_map = extract_sheet_images(file_bytes, s_name)

            for row_idx in range(2, ws.max_row + 1):
                row_vals = []
                row_dict = {}
                for c in range(1, len(headers) + 1):
                    col_name = headers[c - 1]
                    raw_val = ws.cell(row=row_idx, column=c).value
                    
                    if col_name in ['transaction_amount', 'amount']:
                        val = format_transaction_amount(raw_val)
                    else:
                        val = "" if raw_val is None or str(raw_val).strip() == "#VALUE!" else str(raw_val).strip()
                    
                    row_vals.append(val)
                    row_dict[col_name] = val
                
                row_images = row_images_map.get(row_idx, [])
                if any(v != "" for v in row_vals) or len(row_images) > 0:
                    sheets_dict[s_name]['rows'].append({
                        'data': row_dict,
                        'images': row_images
                    })

    return sheets_dict

# =====================================================================
# 5. FUNGSI MEMBUAT WORKBOOK OUTPUT MULTI-SHEET DENGAN BORDER & GAMBAR
# =====================================================================
def create_multisheet_workbook(sheets_dict, target_val, selected_column):
    wb = openpyxl.Workbook()
    wb.remove(wb.active) # Hapus sheet default

    for s_name, s_content in sheets_dict.items():
        headers = s_content['headers']
        rows = s_content['rows']

        # Filter baris di sheet ini yang memiliki nilai sesuai kolom pilihan
        if selected_column in headers:
            filtered_rows = [r for r in rows if r['data'].get(selected_column, "") == target_val]
        else:
            filtered_rows = rows # Jika sheet tidak memiliki kolom tersebut, sertakan semua baris

        clean_sheet_name = str(s_name)[:30].replace("/", "_").replace("\\", "_").replace("?", "_")
        ws = wb.create_sheet(title=clean_sheet_name)
        ws.views.sheetView[0].showGridLines = True

        # Tulis Header
        for col_idx, h_text in enumerate(headers, start=1):
            ws.cell(row=1, column=col_idx, value=h_text)

        # Tulis Data & Tempelkan Gambar
        for row_offset, item in enumerate(filtered_rows, start=2):
            row_dict = item['data']
            for col_idx, h_text in enumerate(headers, start=1):
                val = row_dict.get(h_text, "")
                ws.cell(row=row_offset, column=col_idx, value=val)

            if item['images']:
                for col_idx, img_bytes in item['images']:
                    img_bytes.seek(0)
                    new_img = Image(img_bytes)
                    new_img.width = 100
                    new_img.height = 100
                    col_letter = openpyxl.utils.get_column_letter(col_idx)
                    ws.add_image(new_img, f"{col_letter}{row_offset}")
                    ws.row_dimensions[row_offset].height = 80

        # Terapkan Border
        apply_excel_styling(ws, headers, len(filtered_rows))

    buffer = io.BytesIO()
    wb.save(buffer)
    buffer.seek(0)
    return buffer

# =====================================================================
# 6. INTERFACE STREAMLIT
# =====================================================================
st.set_page_config(page_title="Multi-Sheet Excel Splitter", layout="wide")
st.title("📊 Multi-Sheet Excel Splitter & Filter")
st.write("Unggah file Excel multi-sheet, pilih kolom filter bersama, dan pisahkan seluruh sheet sekaligus ke dalam masing-masing file!")

uploaded_files = st.file_uploader(
    "Pilih satu atau beberapa file Excel (.xlsx)", 
    type=["xlsx"], 
    accept_multiple_files=True
)

if uploaded_files:
    with st.spinner("Membaca seluruh sheet dan data tabel..."):
        sheets_dict = process_multisheet_excel(uploaded_files)
    
    if sheets_dict:
        st.success(f"Berhasil membaca **{len(sheets_dict)}** sheet: {', '.join([f'**{name}** ({len(info['rows'])} baris)' for name, info in sheets_dict.items()])}")

        # Tampilkan Pratinjau untuk Masing-Masing Sheet
        st.markdown("### 👀 Pratinjau Data Tiap Sheet")
        tabs = st.tabs(list(sheets_dict.keys()))
        for idx, (s_name, s_info) in enumerate(sheets_dict.items()):
            with tabs[idx]:
                df_preview = pd.DataFrame([r['data'] for r in s_info['rows']])
                st.dataframe(df_preview.head(10))

        st.markdown("---")

        # Cari Kolom Filter yang Tersedia di Seluruh Sheet
        all_header_sets = [set(info['headers']) for info in sheets_dict.values()]
        common_columns = list(set.intersection(*all_header_sets)) if all_header_sets else []
        
        # Jika tidak ada kolom yang sama persis, gabungkan semua nama kolom
        if not common_columns:
            common_columns = list(set.union(*all_header_sets))

        st.subheader("⚙️ Atur Filter Pemisahan Data")
        selected_column = st.selectbox(
            "Pilih kolom dasar pemisahan (kolom ini akan memfilter seluruh sheet):",
            options=common_columns
        )

        if selected_column:
            # Ambil semua nilai unik dari kolom terpilih di semua sheet
            all_unique_values = set()
            for s_info in sheets_dict.values():
                if selected_column in s_info['headers']:
                    for r in s_info['rows']:
                        v = r['data'].get(selected_column, "").strip()
                        if v:
                            all_unique_values.add(v)
            
            unique_list = sorted(list(all_unique_values))
            st.info(f"Ditemukan **{len(unique_list)}** kategori unik pada kolom **'{selected_column}'**.")

            st.markdown("---")

            # Tombol Unduh ZIP Berisi File Multi-Sheet
            st.subheader("📥 Unduh Hasil Filter")
            st.write("Setiap file Excel hasil unduhan akan berisi **seluruh sheet** yang sudah terfilter secara serentak.")

            if st.button("🚀 Proses & Buat File ZIP Multi-Sheet"):
                with st.spinner("Menyusun file Excel multi-sheet..."):
                    zip_buffer = io.BytesIO()
                    with zipfile.ZipFile(zip_buffer, "w", zipfile.ZIP_DEFLATED) as zip_file:
                        for val in unique_list:
                            excel_buf = create_multisheet_workbook(sheets_dict, val, selected_column)
                            clean_filename = str(val).replace("/", "_").replace("\\", "_").replace("?", "_")
                            zip_file.writestr(f"{clean_filename}.xlsx", excel_buf.getvalue())

                    zip_buffer.seek(0)
                    st.download_button(
                        label=f"📥 Download File ZIP ({len(unique_list)} File Excel Multi-Sheet)",
                        data=zip_buffer,
                        file_name=f"Hasil_Filter_MultiSheet_{selected_column}.zip",
                        mime="application/zip"
                    )
