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
    """Fungsi untuk inisialisasi koneksi ke Google Drive API."""
    try:
        if "gcp_oauth" not in st.secrets:
             return None, "Kredensial OAuth belum diisi di secrets."
             
        oauth_dict = st.secrets["gcp_oauth"]
        
        if not oauth_dict.get("refresh_token"):
             return None, "Refresh token tidak ditemukan."
             
        creds = Credentials(
            token=None,
            refresh_token=oauth_dict["refresh_token"],
            token_uri="https://oauth2.googleapis.com/token",
            client_id=oauth_dict["client_id"],
            client_secret=oauth_dict["client_secret"],
            scopes=SCOPES
        )
        service = build('drive', 'v3', credentials=creds)
        return service, "Berhasil"
    except Exception as e:
        return None, str(e)

# Inisialisasi Service
service, status_msg = get_drive_service()

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



# ==================== KONFIGURASI NAVIGASI ====================

pg_beranda = st.Page(page_beranda, title="Beranda", icon="🏠", default=True)
pg_semua = st.Page(page_semua_data, title="Semua Data", icon="📂")
pg_cari = st.Page(page_cari_data, title="Cari Data", icon="🔍")
pg_upload = st.Page(page_upload_data, title="Upload Data", icon="📤")

pg = st.navigation([pg_beranda, pg_semua, pg_cari, pg_upload])

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
