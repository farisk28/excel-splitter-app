import streamlit as st
import pandas as pd
import openpyxl
from openpyxl.drawing.image import Image
import io
import zipfile

# =====================================================================
# FUNGSI PEMBERSIH TEKS SEL (MENGHILANGKAN .0 DENGAN RAPI)
# =====================================================================
def clean_cell_value(val):
    if val is None:
        return ""
    val_str = str(val).strip()
    
    # Membersihkan teks error #VALUE!
    if val_str == "#VALUE!":
        return ""
    
    # Menghilangkan akhiran .0 dari pembacaan float (misal: 15.0 menjadi 15)
    if val_str.endswith(".0"):
        val_str = val_str[:-2]
        
    return val_str

# =====================================================================
# FUNGSI MEMBACA DATA TEKS DAN GAMBAR MENGGUNAKAN OPENPYXL
# =====================================================================
def process_excel_files(uploaded_files):
    """
    Membaca seluruh file Excel yang diunggah, menggabungkan data teks
    beserta objek gambar jika ada.
    """
    all_rows_data = [] 
    headers = []

    for file in uploaded_files:
        wb = openpyxl.load_workbook(file, data_only=True)
        ws = wb.active

        # Read headers from row 1
        current_headers = [clean_cell_value(cell.value) for cell in ws[1] if cell.value is not None]
        if not headers:
            headers = current_headers

        # Peta gambar untuk sheet ini: {row_index: list_of_images}
        row_images_map = {}
        if hasattr(ws, '_images'):
            for img in ws._images:
                img_row = img.anchor._from.row + 1
                img_bytes = io.BytesIO(img._data())
                
                if img_row not in row_images_map:
                    row_images_map[img_row] = []
                row_images_map[img_row].append((img.anchor._from.col + 1, img_bytes))

        # Membaca isi data dari baris ke-2 ke atas
        for row_idx in range(2, ws.max_row + 1):
            row_vals = []
            for c in range(1, len(headers) + 1):
                val = clean_cell_value(ws.cell(row=row_idx, column=c).value)
                row_vals.append(val)
            
            # Cek apakah baris memiliki isi (tidak kosong seluruhnya)
            if any(v != "" for v in row_vals):
                row_dict = {headers[c]: row_vals[c] for c in range(len(headers))}
                row_images = row_images_map.get(row_idx, [])
                
                all_rows_data.append({
                    'data': row_dict,
                    'images': row_images
                })

    return headers, all_rows_data

# =====================================================================
# FUNGSI MEMBUAT WORKBOOK BARU
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
            ws.cell(row=row_offset, column=col_idx, value=val)

        # Tempelkan Gambar jika ada
        if item['images']:
            for col_idx, img_bytes in item['images']:
                img_bytes.seek(0)
                new_img = Image(img_bytes)
                new_img.width = 100
                new_img.height = 100
                
                col_letter = openpyxl.utils.get_column_letter(col_idx)
                ws.add_image(new_img, f"{col_letter}{row_offset}")
                ws.row_dimensions[row_offset].height = 80

    buffer = io.BytesIO()
    wb.save(buffer)
    buffer.seek(0)
    return buffer

# =====================================================================
# INTERFACE STREAMLIT
# =====================================================================
st.set_page_config(page_title="Excel Splitter & Filter App", layout="wide")
st.title("📊 Aplikasi Penggabung & Pemisah Berkas Excel")
st.write("Unggah beberapa file Excel, lihat pratinjau data gabungan, pilih kolom filter, dan unduh hasilnya dalam format Excel Multi-Sheet atau File ZIP.")

uploaded_files = st.file_uploader(
    "Pilih satu atau beberapa file Excel (.xlsx)", 
    type=["xlsx"], 
    accept_multiple_files=True
)

if uploaded_files:
    # Membaca seluruh data file untuk pratinjau dan pemrosesan
    headers, all_rows_data = process_excel_files(uploaded_files)
    
    if all_rows_data:
        # Konversi ke DataFrame Pandas khusus untuk Pratinjau di Web
        preview_df = pd.DataFrame([item['data'] for item in all_rows_data])
        
        st.success(f"Berhasil membaca dan menggabungkan **{len(uploaded_files)}** file! Total data: **{len(preview_df)}** baris.")
        
        # 1. FITUR PRATINJAU DATA GABUNGAN
        with st.expander("👀 Lihat Pratinjau Data Gabungan (Seluruh File)", expanded=True):
            st.dataframe(preview_df.head(15))

        st.markdown("---")

        # 2. PILIHAN KOLOM FILTER
        st.subheader("⚙️ Atur Filter Pemisahan Data")
        selected_column = st.selectbox(
            "Pilih kolom yang ingin dijadikan dasar pemisahan:",
            options=headers
        )

        if selected_column:
            # Mengelompokkan data berdasarkan kolom terpilih
            grouped_data = {}
            for item in all_rows_data:
                val = item['data'].get(selected_column, "Uncategorized").strip()
                if not val:
                    val = "Uncategorized"
                if val not in grouped_data:
                    grouped_data[val] = []
                grouped_data[val].append(item)

            st.info(f"Ditemukan **{len(grouped_data)}** kategori unik pada kolom **'{selected_column}'**.")

            st.markdown("---")

            # 3. OPSI UNDUHAN
            st.subheader("📥 Unduh Hasil Filter")
            download_format = st.radio(
                "Pilih format hasil unduhan:",
                options=["1 File ZIP (Banyak File Excel Terpisah)", "1 File Excel (Multi-Sheet)"]
            )

            # --- OPSI A: FILE ZIP ---
            if download_format == "1 File ZIP (Banyak File Excel Terpisah)":
                zip_buffer = io.BytesIO()
                with zipfile.ZipFile(zip_buffer, "w", zipfile.ZIP_DEFLATED) as zip_file:
                    for val, items in grouped_data.items():
                        excel_buf = create_excel_workbook(headers, items)
                        clean_filename = str(val).replace("/", "_").replace("\\", "_").replace("?", "_")
                        zip_file.writestr(f"{clean_filename}.xlsx", excel_buf.getvalue())

                zip_buffer.seek(0)
                st.download_button(
                    label="📥 Download File ZIP (Kumpulan File Excel)",
                    data=zip_buffer,
                    file_name=f"Hasil_Filter_{selected_column}.zip",
                    mime="application/zip"
                )

            # --- OPSI B: MULTI-SHEET ---
            else:
                wb_multi = openpyxl.Workbook()
                wb_multi.remove(wb_multi.active)

                for val, items in grouped_data.items():
                    clean_sheet_name = str(val)[:30].replace("/", "_").replace("\\", "_").replace("?", "_")
                    ws = wb_multi.create_sheet(title=clean_sheet_name)

                    for col_idx, h_text in enumerate(headers, start=1):
                        ws.cell(row=1, column=col_idx, value=h_text)

                    for row_offset, item in enumerate(items, start=2):
                        row_dict = item['data']
                        for col_idx, h_text in enumerate(headers, start=1):
                            v = row_dict.get(h_text, "")
                            ws.cell(row=row_offset, column=col_idx, value=v)

                        if item['images']:
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
