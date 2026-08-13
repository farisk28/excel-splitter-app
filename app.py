import streamlit as st
import pandas as pd
import openpyxl
from openpyxl.drawing.image import Image
import io
import zipfile

# =====================================================================
# FUNGSI MEMBACA DATA TEKS DAN GAMBAR MENGGUNAKAN OPENPYXL
# =====================================================================
def process_excel_with_images(uploaded_files, selected_column):
    """
    Membaca beberapa file Excel, menggabungkan data teks beserta gambar,
    lalu mengelompokkan (memisahkan) data berdasarkan kolom terpilih.
    """
    all_rows_data = [] # Menampung data teks (dict)
    all_images_data = [] # Menampung objek gambar dan barisnya
    
    headers = []

    for file_idx, file in enumerate(uploaded_files):
        # Memuat workbook menggunakan openpyxl
        wb = openpyxl.load_workbook(file, data_only=True)
        ws = wb.active

        # Read headers from row 1
        current_headers = [str(cell.value) if cell.value is not None else '' for cell in ws[1]]
        if not headers:
            headers = current_headers

        # Peta gambar untuk sheet ini: {row_index: list_of_images}
        row_images_map = {}
        if hasattr(ws, '_images'):
            for img in ws._images:
                # Ambil koordinat baris (1-based index)
                img_row = img.anchor._from.row + 1
                
                # Extract image bytes
                img_bytes = io.BytesIO(img._data())
                
                if img_row not in row_images_map:
                    row_images_map[img_row] = []
                row_images_map[img_row].append((img.anchor._from.col + 1, img_bytes))

        # Membaca isi data dari baris ke-2 ke atas
        for row_idx in range(2, ws.max_row + 1):
            row_vals = [ws.cell(row=row_idx, column=c).value for c in range(1, len(headers) + 1)]
            
            # Cek apakah baris tidak kosong
            if any(v is not None for v in row_vals):
                row_dict = {headers[c]: row_vals[c] for c in range(len(headers))}
                row_images = row_images_map.get(row_idx, [])
                
                all_rows_data.append({
                    'data': row_dict,
                    'images': row_images
                })

    # Mengelompokkan data berdasarkan nilai unik kolom filter
    grouped_data = {}
    for item in all_rows_data:
        val = str(item['data'].get(selected_column, 'Uncategorized'))
        if val not in grouped_data:
            grouped_data[val] = []
        grouped_data[val].append(item)

    return headers, grouped_data

# =====================================================================
# FUNGSI MEMBUAT WORKBOOK BARU LENGKAP DENGAN GAMBAR
# =====================================================================
def create_excel_workbook(headers, items):
    wb = openpyxl.Workbook()
    ws = wb.active
    ws.title = "Report"

    # Tulis Header
    for col_idx, h_text in enumerate(headers, start=1):
        ws.cell(row=1, column=col_idx, value=h_text)

    # Tulis Baris Data & Tempelkan Gambar
    for row_offset, item in enumerate(items, start=2):
        row_dict = item['data']
        for col_idx, h_text in enumerate(headers, start=1):
            val = row_dict.get(h_text, "")
            ws.cell(row=row_offset, column=col_idx, value=str(val) if val is not None else "")

        # Tempelkan Gambar ke Sel yang Sesuai
        for col_idx, img_bytes in item['images']:
            img_bytes.seek(0)
            new_img = Image(img_bytes)
            # Atur ukuran gambar agar rapi di sel
            new_img.width = 100
            new_img.height = 100
            
            col_letter = openpyxl.utils.get_column_letter(col_idx)
            ws.add_image(new_img, f"{col_letter}{row_offset}")
            ws.row_dimensions[row_offset].height = 80 # Tinggikan baris sel

    buffer = io.BytesIO()
    wb.save(buffer)
    buffer.seek(0)
    return buffer

# =====================================================================
# INTERFACE STREAMLIT
# =====================================================================
st.set_page_config(page_title="Excel Image Splitter App", layout="wide")
st.title("📊 Excel Splitter & Filter (Dengan Dukungan Gambar)")
st.write("Unggah file Excel, pilih kolom filter. Gambar di dalam sel akan tetap dipertahankan pada file output!")

uploaded_files = st.file_uploader(
    "Pilih satu atau beberapa file Excel (.xlsx)", 
    type=["xlsx"], 
    accept_multiple_files=True
)

if uploaded_files:
    # Membaca header dari file pertama untuk pilihan dropdown
    wb_temp = openpyxl.load_workbook(uploaded_files[0], data_only=True)
    first_sheet = wb_temp.active
    available_columns = [str(cell.value) for cell in first_sheet[1] if cell.value is not None]

    selected_column = st.selectbox("Pilih kolom yang ingin dijadikan dasar pemisahan:", options=available_columns)

    download_format = st.radio(
        "Pilih format hasil unduhan:",
        options=["1 File ZIP (Banyak File Excel Terpisah)", "1 File Excel (Multi-Sheet)"]
    )

    if st.button("🚀 Proses & Pisahkan File"):
        with st.spinner("Memproses data dan memindahkan gambar... Mohon tunggu."):
            headers, grouped_data = process_excel_with_images(uploaded_files, selected_column)
            
            st.success(f"Berhasil memisahkan data menjadi **{len(grouped_data)}** kategori!")

            # --- OPSI 1: FILE ZIP ---
            if download_format == "1 File ZIP (Banyak File Excel Terpisah)":
                zip_buffer = io.BytesIO()
                with zipfile.ZipFile(zip_buffer, "w", zipfile.ZIP_DEFLATED) as zip_file:
                    for val, items in grouped_data.items():
                        excel_buf = create_excel_workbook(headers, items)
                        clean_filename = str(val).replace("/", "_").replace("\\", "_").replace("?", "_")
                        zip_file.writestr(f"{clean_filename}.xlsx", excel_buf.getvalue())

                zip_buffer.seek(0)
                st.download_button(
                    label="📥 Download File ZIP",
                    data=zip_buffer,
                    file_name=f"Hasil_Filter_{selected_column}.zip",
                    mime="application/zip"
                )

            # --- OPSI 2: MULTI-SHEET ---
            else:
                wb_multi = openpyxl.Workbook()
                wb_multi.remove(wb_multi.active) # Hapus sheet default

                for val, items in grouped_data.items():
                    clean_sheet_name = str(val)[:30].replace("/", "_").replace("\\", "_").replace("?", "_")
                    ws = wb_multi.create_sheet(title=clean_sheet_name)

                    for col_idx, h_text in enumerate(headers, start=1):
                        ws.cell(row=1, column=col_idx, value=h_text)

                    for row_offset, item in enumerate(items, start=2):
                        row_dict = item['data']
                        for col_idx, h_text in enumerate(headers, start=1):
                            v = row_dict.get(h_text, "")
                            ws.cell(row=row_offset, column=col_idx, value=str(v) if v is not None else "")

                        for col_idx, img_bytes in item['images']:
                            img_bytes.seek(0)
                            new_img = Image(img_bytes)
                            new_img.width = 100
                            new_img.height = 100
                            col_letter = openpyxl.utils.get_column_letter(col_idx)
                            ws.add_image(new_img, f"{col_letter}{row_offset}")
                            ws.row_dimensions[row_offset].height = 80

                multi_buffer = io.BytesIO()
                wb_multi.save(multi_buffer)
                multi_buffer.seek(0)

                st.download_button(
                    label="📥 Download File Excel Multi-Sheet",
                    data=multi_buffer,
                    file_name=f"Hasil_Filter_{selected_column}.xlsx",
                    mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
                )
