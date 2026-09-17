import streamlit as st
import os
import io
import pandas as pd
from datetime import datetime
from googleapiclient.discovery import build
from googleapiclient.http import MediaIoBaseUpload, MediaIoBaseDownload
from google.oauth2.credentials import Credentials

# Konfigurasi Page Streamlit
st.set_page_config(page_title="Database BPS | Dashboard", page_icon="🗄️", layout="wide")

# Custom CSS untuk mempercantik UI dan memperbesar font navigasi
st.markdown("""
<style>
    /* CSS untuk memperbesar Menu Navigasi */
    [data-testid="stSidebarNavItems"] li {
        font-size: 1.25rem !important; 
        margin-bottom: 5px;
    }
    [data-testid="stSidebarNavItems"] a span {
        font-size: 1.15rem !important;
        font-weight: 500;
    }
    
    .main-header {
        font-size: 2.5rem;
        color: #1E88E5;
        font-weight: 700;
        margin-bottom: 0px;
    }
    .sub-header {
        font-size: 1.2rem;
        color: #555555;
        margin-bottom: 2rem;
    }
    .file-card {
        background-color: #f8f9fa;
        padding: 15px;
        border-radius: 10px;
        border: 1px solid #e9ecef;
        margin-bottom: 10px;
    }
    .stButton button {
        border-radius: 20px;
    }
</style>
""", unsafe_allow_html=True)

# ==================== KONEKSI GOOGLE DRIVE ====================

SCOPES = ['https://www.googleapis.com/auth/drive']

@st.cache_resource
def get_drive_service():
    """Fungsi untuk inisialisasi koneksi ke Google Drive & Sheets API."""
    try:
        if "gcp_oauth" not in st.secrets:
             return None, None, "Kredensial OAuth belum diisi di secrets."
             
        oauth_dict = st.secrets["gcp_oauth"]
        
        if not oauth_dict.get("refresh_token"):
             return None, None, "Refresh token tidak ditemukan."
             
        creds = Credentials(
            token=None,
            refresh_token=oauth_dict["refresh_token"],
            token_uri="https://oauth2.googleapis.com/token",
            client_id=oauth_dict["client_id"],
            client_secret=oauth_dict["client_secret"],
            scopes=SCOPES
        )
        drive_svc = build('drive', 'v3', credentials=creds)
        sheets_svc = build('sheets', 'v4', credentials=creds)
        return drive_svc, sheets_svc, "Berhasil"
    except Exception as e:
        return None, None, str(e)

# Inisialisasi Service
service, sheets_service, status_msg = get_drive_service()

# Ambil Folder ID dari secrets
try:
    FOLDER_ID = st.secrets.get("drive_folder_id", "")
except Exception:
    FOLDER_ID = ""

def upload_file_to_drive(file_obj, filename, mimetype):
    if not service or not FOLDER_ID:
        return False, "Google Drive belum terkonfigurasi."
    try:
        file_metadata = {'name': filename, 'parents': [FOLDER_ID]}
        media = MediaIoBaseUpload(file_obj, mimetype=mimetype, resumable=True)
        uploaded_file = service.files().create(
            body=file_metadata, media_body=media, fields='id, name'
        ).execute()
        return True, uploaded_file.get('name')
    except Exception as e:
        return False, str(e)

@st.cache_data(ttl=60) # Cache hasil pencarian selama 60 detik agar lebih cepat
def search_files_in_drive(query=""):
    if not service or not FOLDER_ID:
        return []
    try:
        q = f"'{FOLDER_ID}' in parents and trashed = false"
        if query:
            q += f" and name contains '{query}'"
        results = service.files().list(
            q=q, pageSize=100, fields="nextPageToken, files(id, name, mimeType, size, createdTime, webContentLink)"
        ).execute()
        return results.get('files', [])
    except Exception as e:
        st.error(f"Error mengambil data: {e}")
        return []

def format_date(date_str):
    if not date_str: return "-"
    dt = datetime.strptime(date_str[:19], "%Y-%m-%dT%H:%M:%S")
    return dt.strftime("%d %b %Y, %H:%M")

def format_size(size_bytes):
    if not size_bytes: return "Unknown"
    size = int(size_bytes)
    if size < 1024: return f"{size} B"
    elif size < 1024 * 1024: return f"{size/1024:.1f} KB"
    else: return f"{size/(1024*1024):.1f} MB"

def display_file_card(file):
    """Fungsi pembantu untuk merender tampilan kartu file."""
    file_name = file.get('name')
    file_id = file.get('id')
    web_link = file.get('webContentLink') # Link unduh langsung dari Google
    
    with st.container():
        st.markdown(f'<div class="file-card">', unsafe_allow_html=True)
        col1, col2 = st.columns([3, 1])
        with col1:
            st.markdown(f"**📄 {file_name}**")
            st.caption(f"📅 {format_date(file.get('createdTime'))} &nbsp; | &nbsp; 💾 {format_size(file.get('size'))}")
        with col2:
            if web_link:
                # Tombol unduh langsung (Sangat cepat tanpa fetch ke Python)
                st.link_button("⬇️ Unduh Sekarang", web_link, type="primary", use_container_width=True)
            else:
                st.button("File Tidak Tersedia", disabled=True, use_container_width=True)
        st.markdown('</div>', unsafe_allow_html=True)


# ==================== FUNGSI HALAMAN ====================

def page_beranda():
    st.markdown('<p class="main-header">Selamat Datang di Database BPS 🗄️</p>', unsafe_allow_html=True)
    st.markdown('<p class="sub-header">Pusat pengelolaan dan penyimpanan data terpadu Anda.</p>', unsafe_allow_html=True)
    
    col1, col2, col3 = st.columns(3)
    
    files = search_files_in_drive("")
    total_files = len(files)
    total_size = sum([int(f.get('size', 0)) for f in files])
    
    with col1:
        st.info(f"### 📂 {total_files}\n**Total File Tersimpan**")
    with col2:
        st.success(f"### 💾 {format_size(total_size)}\n**Total Kapasitas Terpakai**")
    with col3:
        st.warning(f"### ⏱️ {datetime.now().strftime('%H:%M')}\n**Waktu Akses Terakhir**")

    st.markdown("---")
    st.markdown("### 📌 Fitur Utama")
    st.markdown("""
    - **📂 Semua Data**: Lihat seluruh daftar file yang ada di direktori penyimpanan Anda.
    - **🔍 Cari Data**: Temukan file spesifik dengan cepat menggunakan kata kunci.
    - **📤 Upload Data**: Simpan file baru ke dalam sistem Cloud secara instan.
    """)

def page_semua_data():
    st.markdown('<p class="main-header">Semua Data</p>', unsafe_allow_html=True)
    st.markdown("Lihat dan kelola semua file yang ada di penyimpanan Drive Anda.")
    st.markdown("---")
    
    with st.spinner("Memuat data dari Drive..."):
        files = search_files_in_drive("")
        
        if not files:
            st.info("Penyimpanan kosong. Belum ada file yang diunggah.")
        else:
            sort_by = st.selectbox("Urutkan berdasarkan:", ["Terbaru", "Terlama", "Nama (A-Z)"])
            
            if sort_by == "Terbaru":
                files = sorted(files, key=lambda x: x.get('createdTime', ''), reverse=True)
            elif sort_by == "Terlama":
                files = sorted(files, key=lambda x: x.get('createdTime', ''))
            elif sort_by == "Nama (A-Z)":
                files = sorted(files, key=lambda x: x.get('name', '').lower())
            
            st.write(f"Menampilkan **{len(files)}** file.")
            
            for file in files:
                display_file_card(file)
            
            if st.button("🔄 Segarkan Data"):
                search_files_in_drive.clear()
                st.rerun()

def page_cari_data():
    st.markdown('<p class="main-header">Cari Data</p>', unsafe_allow_html=True)
    st.markdown("Masukkan kata kunci untuk menemukan file tertentu.")
    
    search_query = st.text_input("🔍 Ketik nama file:", placeholder="Misal: laporan_keuangan.pdf")
    
    if search_query:
        with st.spinner(f"Mencari '{search_query}'..."):
            files = search_files_in_drive(search_query)
            
            st.markdown("---")
            if not files:
                st.warning(f"Tidak ditemukan file dengan nama yang mengandung '{search_query}'.")
            else:
                st.success(f"Ditemukan {len(files)} file yang cocok.")
                for file in files:
                    display_file_card(file)
    else:
        st.info("👆 Masukkan kata kunci pencarian di atas.")

def page_upload_data():
    st.markdown('<p class="main-header">Upload Data Baru</p>', unsafe_allow_html=True)
    st.markdown("Pilih file dari komputer Anda untuk disimpan dengan aman ke Cloud.")
    st.markdown("---")
    
    with st.container():
        # Gunakan session state key agar uploader bisa direset/dikosongkan
        if 'uploader_key' not in st.session_state:
            st.session_state.uploader_key = 1
            
        uploaded_file = st.file_uploader(
            "Seret dan lepas file di sini, atau klik untuk memilih file.", 
            help="Batas ukuran file mengikuti konfigurasi Streamlit Anda.",
            key=str(st.session_state.uploader_key)
        )
        
        if uploaded_file is not None:
            col1, col2 = st.columns([1, 4])
            with col1:
                if st.button("☁️ Mulai Upload", type="primary", use_container_width=True):
                    if service and FOLDER_ID:
                        with st.spinner("Sedang mengunggah ke Drive... mohon tunggu."):
                            success, msg = upload_file_to_drive(uploaded_file, uploaded_file.name, uploaded_file.type)
                            
                            if success:
                                st.success(f"✅ Berhasil! File '{msg}' telah tersimpan.")
                                # Hapus cache pencarian agar file baru langsung muncul
                                search_files_in_drive.clear() 
                                # Reset uploader
                                st.session_state.uploader_key += 1
                                st.rerun()
                            else:
                                st.error(f"❌ Gagal mengunggah file: {msg}")
                    else:
                        st.error("Koneksi Google Drive belum siap.")

# ==================== LOGIKA DATABASE GOOGLE SHEETS ====================
DB_FILE_NAME = "Database_Universal_BPS"

@st.cache_resource
def get_or_create_database():
    if not service or not sheets_service or not FOLDER_ID:
        return None
    
    try:
        # Cari file database
        q = f"'{FOLDER_ID}' in parents and name = '{DB_FILE_NAME}' and trashed = false"
        results = service.files().list(q=q, fields="files(id, name)").execute()
        files = results.get('files', [])
        
        if files:
            return files[0]['id']
        else:
            # Buat file baru
            file_metadata = {
                'name': DB_FILE_NAME,
                'mimeType': 'application/vnd.google-apps.spreadsheet',
                'parents': [FOLDER_ID]
            }
            file = service.files().create(body=file_metadata, fields='id').execute()
            spreadsheet_id = file.get('id')
            
            # Tambahkan header
            headers = [['Tahun', 'Bulan', 'Kategori', 'Indikator', 'Satuan', 'Nilai', 'Keterangan']]
            body = {'values': headers}
            sheets_service.spreadsheets().values().update(
                spreadsheetId=spreadsheet_id, range='Sheet1!A1:G1',
                valueInputOption='USER_ENTERED', body=body
            ).execute()
            return spreadsheet_id
    except Exception as e:
        st.error(f"Error Database: {e}")
        return None

def page_isi_data():
    st.markdown('<p class="main-header">Isi Data Form</p>', unsafe_allow_html=True)
    st.markdown("Masukkan data statistik baru ke dalam database.")
    st.markdown("---")
    
    db_id = get_or_create_database()
    if not db_id:
        st.error("Database belum siap atau koneksi terputus.")
        return
        
    st.info("Silakan isi tabel di bawah ini. Anda bisa menambah baris dengan mengklik area kosong di bawah tabel.")
    
    # Template DataFrame kosong
    if "form_data" not in st.session_state:
        st.session_state.form_data = pd.DataFrame(
            columns=['Tahun', 'Bulan', 'Kategori', 'Indikator', 'Satuan', 'Nilai', 'Keterangan']
        )
        
    edited_df = st.data_editor(
        st.session_state.form_data, 
        num_rows="dynamic",
        use_container_width=True,
        hide_index=True,
        column_config={
            "Tahun": st.column_config.NumberColumn("Tahun", min_value=1900, max_value=2100, step=1, format="%d"),
            "Bulan": st.column_config.NumberColumn("Bulan (1-12)", min_value=1, max_value=12, step=1),
            "Nilai": st.column_config.NumberColumn("Nilai / Angka")
        }
    )
    
    # Filter baris yang kosong
    valid_df = edited_df.dropna(how='all')
    is_empty = valid_df.empty
    
    if st.button("💾 Simpan ke Database", type="primary", disabled=is_empty):
        with st.spinner("Mengecek dan menyimpan data..."):
            try:
                # Ambil data yang sudah ada di database untuk cek duplikat
                result = sheets_service.spreadsheets().values().get(
                    spreadsheetId=db_id, range="Sheet1!A:G"
                ).execute()
                existing_rows = result.get('values', [])
                
                # Konversi input ke list of strings
                new_rows = valid_df.fillna("").astype(str).values.tolist()
                
                # Saring data yang benar-benar baru (belum ada di database)
                rows_to_insert = []
                for row in new_rows:
                    # Samakan panjang elemen jika perlu
                    while len(row) < 7: row.append("")
                    if row not in existing_rows:
                        rows_to_insert.append(row)
                        
                if not rows_to_insert:
                    st.warning("⚠️ Semua baris yang Anda masukkan sudah ada di database (Duplikat). Tidak ada data baru yang dikirim.")
                else:
                    body = {'values': rows_to_insert}
                    
                    sheets_service.spreadsheets().values().append(
                        spreadsheetId=db_id,
                        range="Sheet1!A:G",
                        valueInputOption="USER_ENTERED",
                        insertDataOption="INSERT_ROWS",
                        body=body
                    ).execute()
                    
                    st.success(f"✅ {len(rows_to_insert)} baris data baru berhasil disimpan ke database!")
                    # Kosongkan form setelah simpan
                    st.session_state.form_data = pd.DataFrame(
                        columns=['Tahun', 'Bulan', 'Kategori', 'Indikator', 'Satuan', 'Nilai', 'Keterangan']
                    )
                    st.rerun()
            except Exception as e:
                st.error(f"Gagal menyimpan data: {e}")

def page_ekstrak_data():
    st.markdown('<p class="main-header">Ekstrak & Unduh Data</p>', unsafe_allow_html=True)
    st.markdown("Tarik data dari database dan ekspor menjadi file Excel.")
    st.markdown("---")
    
    db_id = get_or_create_database()
    if not db_id:
        st.error("Database belum siap.")
        return
        
    with st.spinner("Mengambil data dari database..."):
        try:
            result = sheets_service.spreadsheets().values().get(
                spreadsheetId=db_id, range="Sheet1!A:G"
            ).execute()
            values = result.get('values', [])
            
            if not values or len(values) == 1:
                st.info("Database masih kosong. Belum ada data yang bisa diekstrak.")
                return
                
            # Jadikan baris pertama sebagai header
            headers = values[0]
            data = values[1:]
            
            # Sesuaikan jumlah kolom data dengan header jika ada yang terpotong
            clean_data = []
            for row in data:
                row_copy = list(row)
                while len(row_copy) < len(headers):
                    row_copy.append("")
                clean_data.append(row_copy)
                
            df = pd.DataFrame(clean_data, columns=headers)
            
            # --- FITUR FILTER ---
            st.markdown("### 🔍 Filter Data")
            col1, col2, col3, col4 = st.columns(4)
            
            with col1:
                kategori_unik = ["Semua"] + list(df['Kategori'].unique()) if 'Kategori' in df.columns else ["Semua"]
                pilih_kategori = st.selectbox("Kategori", kategori_unik)
                
            with col2:
                indikator_unik = ["Semua"] + list(df['Indikator'].unique()) if 'Indikator' in df.columns else ["Semua"]
                pilih_indikator = st.selectbox("Indikator", indikator_unik)
                    
            with col3:
                tahun_unik = ["Semua"] + list(df['Tahun'].unique()) if 'Tahun' in df.columns else ["Semua"]
                pilih_tahun = st.selectbox("Tahun", tahun_unik)
                
            with col4:
                bulan_unik = ["Semua"] + list(df['Bulan'].unique()) if 'Bulan' in df.columns else ["Semua"]
                pilih_bulan = st.selectbox("Bulan", bulan_unik)
                
            search_kw = st.text_input("🔍 Cari kata kunci (opsional):", placeholder="Contoh: padi")
                    
            # Terapkan filter
            df_filtered = df.copy()
            if pilih_kategori != "Semua":
                df_filtered = df_filtered[df_filtered['Kategori'] == pilih_kategori]
            if pilih_indikator != "Semua":
                df_filtered = df_filtered[df_filtered['Indikator'] == pilih_indikator]
            if pilih_tahun != "Semua":
                df_filtered = df_filtered[df_filtered['Tahun'] == pilih_tahun]
            if pilih_bulan != "Semua":
                df_filtered = df_filtered[df_filtered['Bulan'] == pilih_bulan]
                
            if search_kw:
                # Cari kata kunci di seluruh kolom
                mask = df_filtered.apply(lambda row: row.astype(str).str.contains(search_kw, case=False).any(), axis=1)
                df_filtered = df_filtered[mask]
                
            st.markdown("### 📊 Pratinjau Data")
            st.dataframe(df_filtered, use_container_width=True)
            st.caption(f"Menampilkan {len(df_filtered)} baris data hasil filter (dari total {len(df)} baris).")
            
            # --- EKSPOR KE EXCEL ---
            if not df_filtered.empty:
                output = io.BytesIO()
                with pd.ExcelWriter(output, engine='openpyxl') as writer:
                    df_filtered.to_excel(writer, index=False, sheet_name='Data_BPS')
                excel_data = output.getvalue()
                
                st.download_button(
                    label="⬇️ Unduh sebagai Excel (.xlsx)",
                    data=excel_data,
                    file_name=f"Ekstrak_Data_{datetime.now().strftime('%Y%m%d')}.xlsx",
                    mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                    type="primary"
                )
                
        except Exception as e:
            st.error(f"Gagal memuat data: {e}")


# ==================== KONFIGURASI NAVIGASI ====================

pg_beranda = st.Page(page_beranda, title="Beranda", icon="🏠", default=True)
pg_semua = st.Page(page_semua_data, title="Semua Data", icon="📂")
pg_cari = st.Page(page_cari_data, title="Cari Data", icon="🔍")
pg_upload = st.Page(page_upload_data, title="Upload File", icon="📤")
pg_isi_data = st.Page(page_isi_data, title="Isi Data Form", icon="📝")
pg_ekstrak_data = st.Page(page_ekstrak_data, title="Tarik/Ekstrak Data", icon="📊")

pg = st.navigation([pg_beranda, pg_semua, pg_cari, pg_upload, pg_isi_data, pg_ekstrak_data])

# ==================== UI SIDEBAR (STATUS DI BAWAH MENU) ====================

with st.sidebar:
    # Spacer untuk mendorong konten ke bawah
    st.markdown('<div style="height: 25vh;"></div>', unsafe_allow_html=True)
    
    st.markdown("---")
    st.markdown("### Status Sistem")
    if not service:
        st.error("🔴 Terputus")
    elif not FOLDER_ID:
        st.warning("🟡 Menunggu Folder ID")
    else:
        st.success("🟢 Terhubung ke Database")
        
    st.markdown("---")
    st.caption("© 2026 Database BPS Dashboard")

# Jalankan navigasi utama
pg.run()
