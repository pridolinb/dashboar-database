import streamlit as st
import os
import io
import pandas as pd
from datetime import datetime
from googleapiclient.discovery import build
from googleapiclient.http import MediaIoBaseUpload, MediaIoBaseDownload
from google.oauth2.credentials import Credentials
import openpyxl

# Konfigurasi Page Streamlit
st.set_page_config(page_title="SANDEQ | Dashboard PDRB BPS", page_icon="🗄️", layout="wide")

# Logo dan Nama Sistem di Kiri Atas (Tepat di atas menu navigasi)
st.logo("sandeq_logo.svg", icon_image="sandeq_icon.svg", size="large")

# Custom CSS untuk mempercantik UI dan memperbesar font navigasi
st.markdown("""
<style>
    /* Styling Logo Sandeq di bagian atas sidebar */
    [data-testid="stSidebarHeader"] {
        padding-top: 1.5rem !important;
        padding-bottom: 0.5rem !important;
        height: auto !important;
    }
    [data-testid="stSidebarLogo"] {
        height: auto !important;
        max-height: none !important;
        width: 100% !important;
    }
    [data-testid="stSidebarLogo"] img {
        height: auto !important;
        max-height: 110px !important;
        width: 100% !important;
        object-fit: contain !important;
        object-position: left center !important;
    }

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
    """Fungsi untuk inisialisasi koneksi ke Google Drive."""
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

TRIWULAN_LIST = ["Triwulan 1", "Triwulan 2", "Triwulan 3", "Triwulan 4"]
TRIWULAN_CODE = {
    "Triwulan 1": "TW1",
    "Triwulan 2": "TW2",
    "Triwulan 3": "TW3",
    "Triwulan 4": "TW4"
}

@st.cache_data(ttl=300)
def get_or_create_subfolder(parent_id, folder_name):
    """Mencari atau membuat subfolder di dalam parent_id jika belum ada."""
    if not service or not parent_id:
        return None
    try:
        q = f"'{parent_id}' in parents and name = '{folder_name}' and mimeType = 'application/vnd.google-apps.folder' and trashed = false"
        res = service.files().list(q=q, fields="files(id, name)").execute()
        files = res.get('files', [])
        if files:
            return files[0]['id']
            
        file_metadata = {
            'name': folder_name,
            'mimeType': 'application/vnd.google-apps.folder',
            'parents': [parent_id]
        }
        created = service.files().create(body=file_metadata, fields='id').execute()
        return created.get('id')
    except Exception as e:
        return None

def get_triwulan_folder_id(year, triwulan):
    """Mendapatkan ID subfolder Triwulan di dalam folder Tahun (Opsi A)."""
    if not FOLDER_ID:
        return None
    year_folder_id = get_or_create_subfolder(FOLDER_ID, str(year))
    if not year_folder_id:
        return None
    tw_folder_id = get_or_create_subfolder(year_folder_id, triwulan)
    return tw_folder_id

def format_sandeq_filename(year, triwulan, filename):
    """Memformat nama file otomatis agar diawali {Tahun}_{TW}_{NamaFile}."""
    tw_code = TRIWULAN_CODE.get(triwulan, "TW1")
    prefix = f"{year}_{tw_code}_"
    clean_name = filename.strip()
    if clean_name.startswith(prefix):
        return clean_name
    return f"{prefix}{clean_name}"

@st.cache_data(ttl=60)
def get_available_years():
    """Mengambil daftar tahun yang tersedia di Drive atau default beberapa tahun terkini."""
    curr = datetime.now().year
    years = {str(curr), str(curr - 1), str(curr - 2), str(curr + 1)}
    if service and FOLDER_ID:
        try:
            q = f"'{FOLDER_ID}' in parents and mimeType = 'application/vnd.google-apps.folder' and trashed = false"
            res = service.files().list(q=q, fields="files(id, name)").execute()
            for f in res.get('files', []):
                if f['name'].isdigit():
                    years.add(f['name'])
        except Exception:
            pass
    return sorted(list(years), reverse=True)

def get_search_target_folders(year="Semua", triwulan="Semua"):
    """Mengambil daftar folder ID target berdasarkan filter Tahun dan Triwulan."""
    if not service or not FOLDER_ID:
        return []
        
    if year != "Semua" and triwulan != "Semua":
        fid = get_triwulan_folder_id(year, triwulan)
        return [fid] if fid else []
        
    target_ids = []
    
    # Ambil folder-folder tahun di bawah FOLDER_ID
    try:
        q = f"'{FOLDER_ID}' in parents and mimeType = 'application/vnd.google-apps.folder' and trashed = false"
        res = service.files().list(q=q, fields="files(id, name)").execute()
        year_folders = res.get('files', [])
    except Exception:
        year_folders = []
        
    if year != "Semua":
        matched_years = [f for f in year_folders if f['name'] == str(year)]
    else:
        matched_years = year_folders
        # Sertakan FOLDER_ID langsung agar file lama yang tersimpan di root tetap terbaca saat pilih 'Semua'
        target_ids.append(FOLDER_ID)
        
    for yf in matched_years:
        try:
            q_tw = f"'{yf['id']}' in parents and mimeType = 'application/vnd.google-apps.folder' and trashed = false"
            res_tw = service.files().list(q=q_tw, fields="files(id, name)").execute()
            tw_folders = res_tw.get('files', [])
            for twf in tw_folders:
                if triwulan == "Semua" or twf['name'] == triwulan:
                    target_ids.append(twf['id'])
        except Exception:
            pass
            
    return target_ids

def upload_file_to_drive(file_obj, filename, mimetype, target_folder_id=None, app_properties=None):
    if not service or not FOLDER_ID:
        return False, "Google Drive belum terkonfigurasi."
    try:
        dest_folder = target_folder_id if target_folder_id else FOLDER_ID
        file_metadata = {'name': filename, 'parents': [dest_folder]}
        if app_properties:
            file_metadata['appProperties'] = app_properties
            
        media = MediaIoBaseUpload(file_obj, mimetype=mimetype, resumable=True)
        uploaded_file = service.files().create(
            body=file_metadata, media_body=media, fields='id, name'
        ).execute()
        return True, uploaded_file.get('name')
    except Exception as e:
        return False, str(e)

@st.cache_data(ttl=60) # Cache hasil pencarian selama 60 detik agar lebih cepat
def search_files_in_drive(query="", mime_type=None, app_properties=None, year="Semua", triwulan="Semua"):
    if not service or not FOLDER_ID:
        return []
    try:
        target_folders = get_search_target_folders(year, triwulan)
        if not target_folders:
            return []
            
        if len(target_folders) == 1:
            parents_q = f"'{target_folders[0]}' in parents"
        else:
            parents_q = "(" + " or ".join([f"'{fid}' in parents" for fid in target_folders]) + ")"
            
        q = f"{parents_q} and trashed = false and mimeType != 'application/vnd.google-apps.folder'"
        if query:
            q += f" and name contains '{query}'"
        if mime_type:
            q += f" and mimeType = '{mime_type}'"
        if app_properties:
            for k, v in app_properties.items():
                q += f" and appProperties has {{ key='{k}' and value='{v}' }}"
                
        results = service.files().list(
            q=q, pageSize=100, fields="nextPageToken, files(id, name, mimeType, size, createdTime, webContentLink)"
        ).execute()
        return results.get('files', [])
    except Exception as e:
        st.error(f"Error mengambil data: {e}")
        return []

def download_file_from_drive(file_id):
    if not service: return None
    try:
        request = service.files().get_media(fileId=file_id)
        fh = io.BytesIO()
        downloader = MediaIoBaseDownload(fh, request)
        done = False
        while done is False:
            status, done = downloader.next_chunk()
        fh.seek(0)
        return fh
    except Exception as e:
        st.error(f"Error mengunduh: {e}")
        return None

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
    st.markdown('<p class="main-header">Selamat Datang di SANDEQ 🗄️</p>', unsafe_allow_html=True)
    st.markdown('<p class="sub-header">Sistem Administrasi Data Elektronik Quartal PDRB - BPS Provinsi Sulawesi Barat</p>', unsafe_allow_html=True)
    
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
    - **📂 Semua Data**: Lihat seluruh file berdasarkan filter Tahun dan Triwulan.
    - **📤 Upload File**: Unggah file dari perangkat Anda ke folder Tahun > Triwulan yang dipilih.
    - **📝 Isi Data Form**: Buat tabel Excel fleksibel langsung dari dashboard dan simpan per Triwulan.
    - **📊 Tarik/Ekstrak Data**: Buka, saring, dan unduh data Excel berdasarkan periode Triwulan.
    """)

def page_semua_data():
    st.markdown('<p class="main-header">Semua Data</p>', unsafe_allow_html=True)
    st.markdown("Lihat, cari, dan kelola semua file yang ada di penyimpanan Drive Anda berdasarkan Tahun dan Triwulan.")
    st.markdown("---")
    
    # Filter pencarian, Tahun, Triwulan, dan Pengurutan
    col_search, col_yr, col_tw, col_sort = st.columns([2.5, 1, 1, 1])
    with col_search:
        search_query = st.text_input("🔍 Cari File:", placeholder="Ketik nama file yang ingin dicari...")
    with col_yr:
        filter_year = st.selectbox("Tahun:", ["Semua"] + get_available_years())
    with col_tw:
        filter_tw = st.selectbox("Triwulan:", ["Semua"] + TRIWULAN_LIST)
    with col_sort:
        sort_by = st.selectbox("Urutkan:", ["Terbaru", "Terlama", "Nama (A-Z)"])
    
    with st.spinner("Memuat data dari Drive..."):
        files = search_files_in_drive(search_query, year=filter_year, triwulan=filter_tw)
        
        if not files:
            if search_query:
                st.warning(f"Tidak ditemukan file dengan nama yang mengandung '{search_query}'.")
            else:
                st.info("Tidak ada file pada periode/folder yang dipilih.")
        else:
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

def page_upload_data():
    st.markdown('<p class="main-header">Upload File</p>', unsafe_allow_html=True)
    st.markdown("Pilih satu atau beberapa file dari komputer Anda untuk disimpan ke folder Tahun & Triwulan terkait.")
    st.markdown("---")
    
    # Pilih Tahun dan Triwulan Target
    col_yr, col_tw = st.columns(2)
    with col_yr:
        avail_years = get_available_years()
        curr_year = str(datetime.now().year)
        yr_idx = avail_years.index(curr_year) if curr_year in avail_years else 0
        upload_year = st.selectbox("📅 Pilih Tahun:", avail_years, index=yr_idx)
    with col_tw:
        upload_tw = st.selectbox("📊 Pilih Triwulan:", TRIWULAN_LIST)
        
    st.markdown("---")
    
    with st.container():
        # Gunakan session state key agar uploader bisa direset/dikosongkan
        if 'uploader_key' not in st.session_state:
            st.session_state.uploader_key = 1
            
        uploaded_files = st.file_uploader(
            "Seret dan lepas file di sini, atau klik untuk memilih file (bisa pilih banyak file sekaligus).", 
            accept_multiple_files=True,
            help="Batas ukuran file mengikuti konfigurasi Streamlit Anda.",
            key=str(st.session_state.uploader_key)
        )
        
        if uploaded_files:
            tw_code = TRIWULAN_CODE.get(upload_tw, "TW1")
            st.info(f"📁 Terpilih **{len(uploaded_files)}** file. Nama file akan otomatis diawali format: `{upload_year}_{tw_code}_` dan disimpan ke folder **{upload_year} > {upload_tw}**.")
            col1, col2 = st.columns([1, 4])
            with col1:
                if st.button("☁️ Mulai Upload Semua", type="primary", use_container_width=True):
                    if service and FOLDER_ID:
                        total_files = len(uploaded_files)
                        progress_bar = st.progress(0)
                        status_text = st.empty()
                        
                        berhasil = 0
                        gagal = []
                        
                        # Dapatkan folder ID triwulan tujuan (otomatis buat jika belum ada)
                        target_folder = get_triwulan_folder_id(upload_year, upload_tw)
                        if not target_folder:
                            st.error(f"Gagal menyiapkan folder Google Drive untuk {upload_year} > {upload_tw}.")
                            return
                        
                        for idx, file in enumerate(uploaded_files):
                            final_name = format_sandeq_filename(upload_year, upload_tw, file.name)
                            status_text.text(f"Mengunggah ({idx + 1}/{total_files}): {final_name}...")
                            success, msg = upload_file_to_drive(file, final_name, file.type, target_folder_id=target_folder)
                            if success:
                                berhasil += 1
                            else:
                                gagal.append((final_name, msg))
                            progress_bar.progress((idx + 1) / total_files)
                            
                        # Hapus cache pencarian agar data terbaru langsung muncul
                        search_files_in_drive.clear()
                        
                        if berhasil > 0:
                            st.success(f"✅ Berhasil mengunggah **{berhasil}** file ke folder **{upload_year} > {upload_tw}**!")
                        if gagal:
                            for nama_file, err in gagal:
                                st.error(f"❌ Gagal mengunggah '{nama_file}': {err}")
                                
                        if berhasil > 0 and not gagal:
                            # Reset uploader otomatis jika semua sukses
                            st.session_state.uploader_key += 1
                            st.rerun()
                    else:
                        st.error("Koneksi Google Drive belum siap.")

def page_isi_data():
    st.markdown('<p class="main-header">Isi Data Form (Kanvas Excel)</p>', unsafe_allow_html=True)
    st.markdown("Ketik nama **Judul Kolom (Header)** di **baris paling pertama (No. 1)**, lalu isi datanya di baris-baris bawahnya.")
    st.markdown("---")
    
    if not service or not FOLDER_ID:
        st.error("Koneksi Google Drive belum siap atau terputus.")
        return
        
    # Kunci unik agar data_editor bisa di-reset sepenuhnya
    if "canvas_key" not in st.session_state:
        st.session_state.canvas_key = 1
        
    # Setup kanvas kosong (misal 10 kolom A s.d. J)
    if "canvas_data" not in st.session_state:
        # Buat dataframe kosong dengan nama kolom A, B, C...
        cols = ['A', 'B', 'C', 'D', 'E', 'F', 'G', 'H', 'I', 'J']
        empty_data = [["" for _ in cols] for _ in range(15)]
        st.session_state.canvas_data = pd.DataFrame(empty_data, columns=cols)
        
    st.info("💡 **Tips:** Baris No.1 otomatis akan dijadikan *Header* (Judul Kolom) saat file disimpan. Kolom yang sama sekali kosong tidak akan ikut disimpan.")
    
    # Render data editor dengan key yang dinamis
    edited_df = st.data_editor(
        st.session_state.canvas_data, 
        num_rows="dynamic",
        use_container_width=True,
        key=f"canvas_editor_{st.session_state.canvas_key}"
    )
    
    st.markdown("---")
    st.markdown("### 📅 Pengaturan Periode Data")
    col_yr, col_tw = st.columns(2)
    with col_yr:
        avail_years = get_available_years()
        curr_year = str(datetime.now().year)
        yr_idx = avail_years.index(curr_year) if curr_year in avail_years else 0
        form_year = st.selectbox("Tahun Data:", avail_years, index=yr_idx, key="form_year")
    with col_tw:
        form_tw = st.selectbox("Triwulan Data:", TRIWULAN_LIST, key="form_tw")
    
    st.markdown("---")
    
    col_name, col_btn = st.columns([3, 1])
    with col_name:
        tw_code = TRIWULAN_CODE.get(form_tw, "TW1")
        file_name_input = st.text_input("Nama File Excel Anda:", placeholder="Contoh: Produksi_Padi", help=f"Nama file otomatis diawali '{form_year}_{tw_code}_'. Tidak perlu menuliskan .xlsx.")
    with col_btn:
        st.write("") # Spacer agar sejajar dengan input
        st.write("")
        simpan_btn = st.button("💾 Simpan", type="primary", use_container_width=True)
        
    if simpan_btn:
        if file_name_input.strip() == "":
            st.warning("Silakan masukkan nama file terlebih dahulu.")
            return
            
        with st.spinner("Memproses tabel dan menyimpan ke Drive..."):
            try:
                # 1. Bersihkan dataframe dari kolom & baris yang kosong semua
                import numpy as np
                cleaned_df = edited_df.replace("", np.nan)
                
                # Buang baris yang isinya kosong semua, dan kolom yang isinya kosong semua
                cleaned_df = cleaned_df.dropna(how='all', axis=0).dropna(how='all', axis=1)
                
                if cleaned_df.empty:
                    st.warning("Tabel masih benar-benar kosong.")
                    return
                
                # 2. Jadikan baris pertama yang tidak kosong sebagai HEADER
                header_row = cleaned_df.iloc[0]
                data_rows = cleaned_df.iloc[1:]
                data_rows.columns = header_row
                
                # 3. Ubah ke file Excel (di memory)
                output = io.BytesIO()
                with pd.ExcelWriter(output, engine='openpyxl') as writer:
                    data_rows.to_excel(writer, index=False, sheet_name='Sheet1')
                output.seek(0)
                
                # 4. Upload ke Google Drive ke subfolder Tahun > Triwulan
                base_name = file_name_input.strip()
                if not base_name.endswith('.xlsx'):
                    base_name += ".xlsx"
                final_filename = format_sandeq_filename(form_year, form_tw, base_name)
                mime = "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
                
                target_folder = get_triwulan_folder_id(form_year, form_tw)
                if not target_folder:
                    st.error(f"Gagal menyiapkan folder Google Drive untuk {form_year} > {form_tw}.")
                    return
                
                # Metadata ini akan membedakan file ini dari file Excel upload-an biasa
                app_props = {
                    "source": "isi_data_form",
                    "year": str(form_year),
                    "triwulan": form_tw
                }
                
                success, msg = upload_file_to_drive(output, final_filename, mime, target_folder_id=target_folder, app_properties=app_props)
                
                if success:
                    st.success(f"✅ Berhasil! File '{msg}' telah disimpan ke folder {form_year} > {form_tw} di Google Drive.")
                    search_files_in_drive.clear()
                    
                    # Hapus data sebelumnya
                    del st.session_state.canvas_data
                    # Naikkan key agar komponen UI refresh sepenuhnya
                    st.session_state.canvas_key += 1
                    
                    st.rerun()
                else:
                    st.error(f"Gagal mengunggah file: {msg}")
            except Exception as e:
                st.error(f"Terjadi kesalahan saat memproses tabel: {str(e)}")

def page_ekstrak_data():
    st.markdown('<p class="main-header">Tarik & Filter Excel</p>', unsafe_allow_html=True)
    st.markdown("Lihat isi file Excel yang Anda buat melalui menu *Isi Data Form*, dan saring isinya berdasarkan periode Triwulan.")
    st.markdown("---")
    
    if not service:
        st.error("Koneksi Google Drive belum siap.")
        return
        
    col_yr, col_tw = st.columns(2)
    with col_yr:
        filter_year = st.selectbox("Filter Tahun:", ["Semua"] + get_available_years(), key="ekstrak_year")
    with col_tw:
        filter_tw = st.selectbox("Filter Triwulan:", ["Semua"] + TRIWULAN_LIST, key="ekstrak_tw")
        
    with st.spinner("Mencari daftar file formulir Excel di Drive..."):
        # Cari hanya file berformat Excel (.xlsx) yang dibuat melalui Isi Data Form
        excel_mime = "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
        files = search_files_in_drive("", mime_type=excel_mime, app_properties={"source": "isi_data_form"}, year=filter_year, triwulan=filter_tw)
        
    if not files:
        st.info("Belum ada file Excel formulir yang cocok dengan filter periode Tahun & Triwulan yang dipilih.")
        return
        
    # Buat pilihan file
    file_options = {f"{f.get('name')}  (Dibuat: {format_date(f.get('createdTime'))})": f.get('id') for f in files}
    pilihan_file = st.selectbox("Pilih File Excel yang ingin dibuka:", list(file_options.keys()))
    file_id = file_options[pilihan_file]
    
    st.markdown("---")
    
    if file_id:
        with st.spinner(f"Membuka {pilihan_file}..."):
            try:
                # Unduh file Excel dari Drive ke memori
                file_stream = download_file_from_drive(file_id)
                if not file_stream:
                    st.error("Gagal mengunduh file.")
                    return
                    
                # Baca menggunakan pandas
                df = pd.read_excel(file_stream, engine='openpyxl')
                
                if df.empty:
                    st.warning("File Excel ini kosong.")
                    return
                    
                st.markdown("### 🔍 Filter Data")
                search_kw = st.text_input("Ketik kata kunci untuk mencari di kolom apa saja:", placeholder="Contoh: Budi, Padi, 2026...")
                
                df_filtered = df.copy()
                if search_kw:
                    # Terapkan filter text mengandung kata kunci di baris manapun
                    mask = df_filtered.apply(lambda row: row.astype(str).str.contains(search_kw, case=False).any(), axis=1)
                    df_filtered = df_filtered[mask]
                    
                st.markdown("### 📊 Pratinjau Tabel")
                st.dataframe(df_filtered, use_container_width=True)
                st.caption(f"Menampilkan {len(df_filtered)} baris dari total {len(df)} baris data.")
                
                # Tombol unduh hasil saringan
                if not df_filtered.empty:
                    output = io.BytesIO()
                    with pd.ExcelWriter(output, engine='openpyxl') as writer:
                        df_filtered.to_excel(writer, index=False, sheet_name='Filtered_Data')
                    output.seek(0)
                    
                    # Bersihkan nama file asli dan pastikan berekstensi .xlsx
                    nama_asli = pilihan_file.split('  (')[0].strip()
                    if nama_asli.endswith('.xlsx'):
                        nama_unduhan = f"Filter_{nama_asli}"
                    else:
                        nama_unduhan = f"Filter_{nama_asli}.xlsx"
                        
                    col1, col2 = st.columns([1, 4])
                    with col1:
                        st.download_button(
                            label="⬇️ Unduh Data",
                            data=output,
                            file_name=nama_unduhan,
                            mime=excel_mime,
                            type="primary",
                            use_container_width=True
                        )
            except Exception as e:
                st.error(f"Gagal membaca atau memproses file Excel: {e}")

# ==================== KONFIGURASI NAVIGASI ====================

pg_beranda = st.Page(page_beranda, title="Beranda", icon="🏠", default=True)
pg_semua = st.Page(page_semua_data, title="Semua Data", icon="📂")
pg_upload = st.Page(page_upload_data, title="Upload File", icon="📤")
pg_isi_data = st.Page(page_isi_data, title="Isi Data Form", icon="📝")
pg_ekstrak_data = st.Page(page_ekstrak_data, title="Tarik/Ekstrak Data", icon="📊")

pg = st.navigation([pg_beranda, pg_semua, pg_upload, pg_isi_data, pg_ekstrak_data])

# ==================== UI SIDEBAR ====================

with st.sidebar:
    # Spacer untuk mendorong status ke bawah menu navigasi
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
    st.caption("© 2026 SANDEQ - BPS")

# Jalankan navigasi utama
pg.run()
