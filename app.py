import streamlit as st
import pandas as pd
import openpyxl
from openpyxl.drawing.image import Image
from openpyxl.styles import Border, Side, Font, PatternFill, Alignment
import io
import zipfile
import xml.etree.ElementTree as ET

# =====================================================================
# FUNGSI EKSTRAKSI GAMBAR IN-CELL (EXCEL 365 RICH DATA) & FLOATING
# =====================================================================
def extract_row_images_map(file_bytes):
    """
    Ekstrak peta gambar per baris {row_idx: [(col_idx, img_bytes)]}
    dari file Excel, mendukung gambar In-Cell (RichData) dan Floating.
    """
    row_images_map = {}
    
    # 1. Coba ekstraksi In-Cell Images dari arsip Zip (XML RichData)
    try:
        with zipfile.ZipFile(io.BytesIO(file_bytes), 'r') as z:
            namelist = z.namelist()
            
            # Cek apakah ada folder richData dan media
            if 'xl/richData/richValueRel.xml' in namelist and 'xl/richData/_rels/richValueRel.xml.rels' in namelist:
                # Read rels
                rels_xml = z.read("xl/richData/_rels/richValueRel.xml.rels")
                root_rels = ET.fromstring(rels_xml)
                rel_id_to_target = {}
                for elem in root_rels:
                    rel_id_to_target[elem.attrib['Id']] = elem.attrib['Target'].replace('../media/', '')
                
                # Read richValueRel.xml
                rich_rel_xml = z.read("xl/richData/richValueRel.xml")
                root_rich_rel = ET.fromstring(rich_rel_xml)
                index_to_media = []
                for elem in root_rich_rel:
                    r_id = elem.attrib['{http://schemas.openxmlformats.org/officeDocument/2006/relationships}id']
                    if r_id in rel_id_to_target:
                        index_to_media.append(rel_id_to_target[r_id])

                # Read rdrichvalue.xml
                rv_xml = z.read("xl/richData/rdrichvalue.xml")
                root_rv = ET.fromstring(rv_xml)
                rv_index_to_media = []
                for rv in root_rv:
                    val_idx = int(rv[0].text)
                    if val_idx < len(index_to_media):
                        rv_index_to_media.append(index_to_media[val_idx])

                # Read sheet1.xml untuk memetakan baris
                sheet_xml = z.read("xl/worksheets/sheet1.xml")
                root_sheet = ET.fromstring(sheet_xml)
                
                ns_main = '{http://schemas.openxmlformats.org/spreadsheetml/2006/main}'
                sheet_data = root_sheet.find(f'{ns_main}sheetData')
                
                if sheet_data is not None:
                    for row_elem in sheet_data.findall(f'{ns_main}row'):
                        r_idx = int(row_elem.attrib['r'])
                        for c_elem in row_elem.findall(f'{ns_main}c'):
                            vm_attr = c_elem.attrib.get('vm')
                            if vm_attr is not None:
                                vm_idx = int(vm_attr) - 1
                                if 0 <= vm_idx < len(rv_index_to_media):
                                    media_name = rv_index_to_media[vm_idx]
                                    media_path = f"xl/media/{media_name}"
                                    if media_path in namelist:
                                        img_data = z.read(media_path)
                                        # Kolom Evidence standar berada di kolom 9 (Kolom I)
                                        col_idx = 9
                                        if r_idx not in row_images_map:
                                            row_images_map[r_idx] = []
                                        row_images_map[r_idx].append((col_idx, io.BytesIO(img_data)))
    except Exception:
        pass

    # 2. Coba ekstraksi Gambar Floating biasa menggunakan OpenPyXL
    try:
        wb = openpyxl.load_workbook(io.BytesIO(file_bytes), data_only=True)
        ws = wb.active
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
# FUNGSI FORMAT NOMINAL TRANSACTION AMOUNT
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
# FUNGSI MENAMBAHKAN BORDER & STYLING EXCEL
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
    
    for col_idx in range(1, len(headers) + 1):
        cell = ws.cell(row=1, column=col_idx)
        cell.font = header_font
        cell.fill = header_fill
        cell.border = thin_border
        cell.alignment = Alignment(vertical="center", horizontal="center")

    for row_idx in range(2, num_rows + 2):
        for col_idx in range(1, len(headers) + 1):
            cell = ws.cell(row=row_idx, column=col_idx)
            cell.border = thin_border
            cell.alignment = Alignment(vertical="center")

# =====================================================================
# FUNGSI MEMBACA DATA TEKS & GAMBAR LENGKAP
# =====================================================================
def process_excel_files(uploaded_files):
    all_rows_data = [] 
    headers = []

    for file in uploaded_files:
        file_bytes = file.getvalue()
        row_images_map = extract_row_images_map(file_bytes)

        wb = openpyxl.load_workbook(io.BytesIO(file_bytes), data_only=True)
        ws = wb.active

        current_headers = [str(cell.value).strip() for cell in ws[1] if cell.value is not None]
        if not headers:
            headers = current_headers

        for row_idx in range(2, ws.max_row + 1):
            row_vals = []
            row_dict = {}
            for c in range(1, len(headers) + 1):
                col_name = headers[c - 1]
                raw_val = ws.cell(row=row_idx, column=c).value
                
                if col_name == 'transaction_amount':
                    val = format_transaction_amount(raw_val)
                else:
                    val = "" if raw_val is None or str(raw_val).strip() == "#VALUE!" else str(raw_val).strip()
                
                row_vals.append(val)
                row_dict[col_name] = val
            
            row_images = row_images_map.get(row_idx, [])
            if any(v != "" for v in row_vals) or len(row_images) > 0:
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
    ws.views.sheetView[0].showGridLines = True

    for col_idx, h_text in enumerate(headers, start=1):
        ws.cell(row=1, column=col_idx, value=h_text)

    for row_offset, item in enumerate(items, start=2):
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

    apply_excel_styling(ws, headers, len(items))

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
    headers, all_rows_data = process_excel_files(uploaded_files)
    
    if all_rows_data:
        preview_df = pd.DataFrame([item['data'] for item in all_rows_data])
        
        st.success(f"Berhasil membaca dan menggabungkan **{len(uploaded_files)}** file! Total data: **{len(preview_df)}** baris.")
        
        with st.expander("👀 Lihat Pratinjau Data Gabungan (Seluruh File)", expanded=True):
            st.dataframe(preview_df.head(15))

        st.markdown("---")

        st.subheader("⚙️ Atur Filter Pemisahan Data")
        selected_column = st.selectbox(
            "Pilih kolom yang ingin dijadikan dasar pemisahan:",
            options=headers
        )

        if selected_column:
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
                    ws.views.sheetView[0].showGridLines = True

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

                    apply_excel_styling(ws, headers, len(items))

                multi_buffer = io.BytesIO()
                wb_multi.save(multi_buffer)
                multi_buffer.seek(0)

                st.download_button(
                    label="📥 Download File Excel Multi-Sheet",
                    data=multi_buffer,
                    file_name=f"Hasil_Filter_{selected_column}.xlsx",
                    mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
                )
