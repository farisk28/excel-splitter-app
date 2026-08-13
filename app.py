import streamlit as st
import pandas as pd
import io
import zipfile
import re

# =====================================================================
# FUNGSI PEMBERSIH & STANDARISASI FORMAT TRANSACTION AMOUNT
# =====================================================================
def format_transaction_amount(val):
    if pd.isna(val):
        return val
    
    val_str = str(val).strip()
    
    # 1. Hapus karakter non-angka (seperti koma atau titik)
    clean_digits = re.sub(r'[^\d]', '', val_str)
    
    if not clean_digits:
        return val_str
    
    num = int(clean_digits)
    
    # 2. Koreksi nilai jika terbaca di bawah 1000 (misal: 5 menjadi 5000)
    if 0 < num < 1000:
        num = num * 1000
        
    # 3. Format ulang menjadi string dengan pemisah ribuan koma (contoh: 5,000)
    return f"{num:,}"

# =====================================================================
# 1. KONFIGURASI HALAMAN STREAMLIT
# =====================================================================
st.set_page_config(page_title="Excel Splitter & Filter App", layout="wide")
st.title("📊 Aplikasi Penggabung & Pemisah Berkas Excel")
st.write("Unggah beberapa file Excel, pilih kolom filter, dan unduh hasilnya dalam format Excel Multi-Sheet atau File ZIP.")

# =====================================================================
# 2. KOMPONEN UNGGAH BERKAS
# =====================================================================
uploaded_files = st.file_uploader(
    "Pilih satu atau beberapa file Excel (.xlsx)", 
    type=["xlsx"], 
    accept_multiple_files=True
)

if uploaded_files:
    df_list = []
    for file in uploaded_files:
        try:
            # Membaca seluruh data sebagai string
            temp_df = pd.read_excel(file, dtype=str)
            
            # Terapkan fungsi pembersihan format angka jika kolom transaction_amount ditemukan
            if 'transaction_amount' in temp_df.columns:
                temp_df['transaction_amount'] = temp_df['transaction_amount'].apply(format_transaction_amount)
                
            df_list.append(temp_df)
        except Exception as e:
            st.error(f"Gagal membaca file {file.name}: {e}")

    if df_list:
        combined_df = pd.concat(df_list, ignore_index=True)
        st.success(f"Berhasil menggabungkan {len(uploaded_files)} file! Total baris data: **{len(combined_df)}**.")
        
        with st.expander("👀 Lihat Pratinjau Data Gabungan"):
            st.dataframe(combined_df.head(10))

        st.markdown("---")

        # =====================================================================
        # 3. PILIHAN KOLOM FILTER
        # =====================================================================
        st.subheader("⚙️ Atur Filter Pemisahan Data")
        available_columns = list(combined_df.columns)
        selected_column = st.selectbox(
            "Pilih kolom yang ingin dijadikan dasar pemisahan:",
            options=available_columns
        )

        if selected_column:
            unique_values = combined_df[selected_column].dropna().unique()
            st.info(f"Ditemukan **{len(unique_values)}** kategori unik pada kolom **'{selected_column}'**.")

            st.markdown("---")

            # =====================================================================
            # 4. OPSI UNDUHAN (EXCEL MULTI-SHEET ATAU ZIP)
            # =====================================================================
            st.subheader("📥 Unduh Hasil Filter")
            download_format = st.radio(
                "Pilih format hasil unduhan:",
                options=["1 File Excel (Multi-Sheet)", "1 File ZIP (Banyak File Excel Terpisah)"]
            )

            # --- OPSI A: 1 FILE EXCEL MULTI-SHEET ---
            if download_format == "1 File Excel (Multi-Sheet)":
                output_excel = io.BytesIO()
                with pd.ExcelWriter(output_excel, engine='openpyxl') as writer:
                    for val in unique_values:
                        filtered_df = combined_df[combined_df[selected_column] == val]
                        clean_sheet_name = str(val)[:30].replace("/", "_").replace("\\", "_").replace("?", "_")
                        filtered_df.to_excel(writer, sheet_name=clean_sheet_name, index=False)
                
                output_excel.seek(0)
                st.download_button(
                    label="📥 Download File Excel Multi-Sheet",
                    data=output_excel,
                    file_name=f"Hasil_Filter_{selected_column}.xlsx",
                    mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
                )

            # --- OPSI B: 1 FILE ZIP ---
            else:
                zip_buffer = io.BytesIO()
                with zipfile.ZipFile(zip_buffer, "w", zipfile.ZIP_DEFLATED) as zip_file:
                    for val in unique_values:
                        filtered_df = combined_df[combined_df[selected_column] == val]
                        excel_buffer = io.BytesIO()
                        with pd.ExcelWriter(excel_buffer, engine='openpyxl') as writer:
                            filtered_df.to_excel(writer, index=False)
                        
                        excel_buffer.seek(0)
                        clean_filename = str(val).replace("/", "_").replace("\\", "_")
                        zip_file.writestr(f"{clean_filename}.xlsx", excel_buffer.getvalue())
                
                zip_buffer.seek(0)
                st.download_button(
                    label="📥 Download File ZIP (Kumpulan File Excel)",
                    data=zip_buffer,
                    file_name=f"Hasil_Filter_{selected_column}.zip",
                    mime="application/zip"
                )
