import streamlit as st
import os
import io
import json
import hashlib
import re
import time
import threading
import pandas as pd
from datetime import datetime
from googleapiclient.discovery import build
from googleapiclient.http import MediaIoBaseUpload, MediaIoBaseDownload
from google.oauth2.credentials import Credentials
import openpyxl

# Konfigurasi Page Streamlit
st.set_page_config(page_title="SANDEQ | Dashboard PDRB BPS", page_icon="🗄️", layout="wide")

# Pastikan file aset SVG selalu tersedia agar st.logo tidak pernah crash di Streamlit Cloud
def _ensure_svg_assets():
    logo_path = "sandeq_logo.svg"
    icon_path = "sandeq_icon.svg"
    try:
        if not os.path.exists(logo_path):
            with open(logo_path, "w", encoding="utf-8") as f:
                f.write(
                    '<svg xmlns="http://www.w3.org/2000/svg" viewBox="0 0 300 98" width="300" height="98">\n'
                    '  <text x="2" y="36" font-family="-apple-system, BlinkMacSystemFont, \'Segoe UI\', Roboto, Helvetica, Arial, sans-serif" font-size="38" font-weight="900" fill="#1E88E5" letter-spacing="2">SANDEQ</text>\n'
                    '  <text x="2" y="60" font-family="-apple-system, BlinkMacSystemFont, \'Segoe UI\', Roboto, Helvetica, Arial, sans-serif" font-size="13.5" font-weight="600" fill="#334155">Sistem Administrasi Data Elektronik</text>\n'
                    '  <text x="2" y="79" font-family="-apple-system, BlinkMacSystemFont, \'Segoe UI\', Roboto, Helvetica, Arial, sans-serif" font-size="13.5" font-weight="700" fill="#1E88E5">Quartal PDRB</text>\n'
                    '  <line x1="2" y1="94" x2="298" y2="94" stroke="#CBD5E1" stroke-width="2" />\n'
                    '</svg>\n'
                )
        if not os.path.exists(icon_path):
            with open(icon_path, "w", encoding="utf-8") as f:
                f.write(
                    '<svg xmlns="http://www.w3.org/2000/svg" viewBox="0 0 32 32" width="32" height="32">\n'
                    '  <rect width="32" height="32" rx="6" fill="#1E88E5"/>\n'
                    '  <text x="16" y="23" font-family="-apple-system, BlinkMacSystemFont, \'Segoe UI\', Roboto, Helvetica, Arial, sans-serif" font-size="20" font-weight="900" fill="#FFFFFF" text-anchor="middle">S</text>\n'
                    '</svg>\n'
                )
        st.logo(logo_path, icon_image=icon_path, size="large")
    except Exception:
        pass

_ensure_svg_assets()

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

# ==================== KONEKSI GOOGLE DRIVE (THREAD-SAFE & AUTO-RECONNECT) ====================

SCOPES = ['https://www.googleapis.com/auth/drive']
_drive_lock = threading.RLock()
_thread_local = threading.local()

def _build_fresh_drive_service():
    """Membuat koneksi Google Drive baru yang segar (mencegah SSL timeout & thread crash)."""
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
        raw_srv = build('drive', 'v3', credentials=creds, cache_discovery=False)
        return raw_srv, "Berhasil"
    except Exception as e:
        return None, str(e)

def get_drive_service(force_reconnect=False):
    """Mengambil instance service Google Drive yang terisolasi per-thread dan otomatis menyambung ulang."""
    with _drive_lock:
        if force_reconnect or getattr(_thread_local, "raw_service", None) is None:
            srv, msg = _build_fresh_drive_service()
            _thread_local.raw_service = srv
            _thread_local.status_msg = msg
        return _thread_local.raw_service, getattr(_thread_local, "status_msg", "Berhasil")

def reset_all_caches_and_connections():
    """Membersihkan seluruh cache aplikasi dan mereset koneksi Google Drive."""
    try:
        st.cache_data.clear()
    except Exception:
        pass
    try:
        st.cache_resource.clear()
    except Exception:
        pass
    with _drive_lock:
        _thread_local.raw_service = None

def auto_recover_system(error_msg="", force_immediate_rerun=False):
    """
    Sistem pemulihan otomatis (Self-Healing):
    Saat mendeteksi error/koneksi putus, otomatis bersihkan cache, sambung ulang Drive, dan rerun aplikasi.
    Dilengkapi pembatas (cooldown) agar tidak terjebak infinite loop jika jaringan Google sedang gangguan.
    """
    reset_all_caches_and_connections()
    get_drive_service(force_reconnect=True)
    
    now_ts = time.time()
    last_ts = st.session_state.get("_last_auto_recover_ts", 0.0)
    recover_count = st.session_state.get("_auto_recover_count", 0)
    
    # Reset penghitung jika sudah lewat 20 detik sejak pemulihan terakhir
    if now_ts - last_ts > 20.0:
        recover_count = 0
        
    if force_immediate_rerun or recover_count < 2:
        st.session_state["_auto_recover_count"] = recover_count + 1
        st.session_state["_last_auto_recover_ts"] = now_ts
        st.rerun()
    else:
        st.warning("⚠️ Sistem mendeteksi gangguan koneksi sementara ke server Google Drive.")
        if error_msg:
            st.caption(f"Detail teknis: `{error_msg}`")
        if st.button("🔄 Reboot & Pulihkan Aplikasi Sekarang", type="primary", width="stretch"):
            st.session_state["_auto_recover_count"] = 0
            st.session_state["_last_auto_recover_ts"] = 0.0
            reset_all_caches_and_connections()
            st.rerun()

class _ResilientRequestProxy:
    """Membungkus setiap perintah API Google Drive dengan Auto-Reconnect & Auto-Retry (3x)."""
    def __init__(self, method_name, args, kwargs):
        self.method_name = method_name
        self.args = args
        self.kwargs = kwargs

    def execute(self, *exec_args, **exec_kwargs):
        last_err = None
        for attempt in range(3):
            try:
                with _drive_lock:
                    raw_srv, err_msg = get_drive_service(force_reconnect=(attempt > 0))
                    if not raw_srv:
                        raise RuntimeError(err_msg or "Koneksi Google Drive belum siap.")
                    # Jika ada stream upload (media_body), kembalikan posisi baca ke awal saat retry
                    media = self.kwargs.get("media_body")
                    if media is not None and hasattr(media, "_fd") and hasattr(media._fd, "seek"):
                        try:
                            media._fd.seek(0)
                        except Exception:
                            pass
                    files_res = raw_srv.files()
                    api_method = getattr(files_res, self.method_name)
                    req = api_method(*self.args, **self.kwargs)
                    return req.execute(*exec_args, **exec_kwargs)
            except Exception as e:
                last_err = e
                with _drive_lock:
                    _thread_local.raw_service = None
                if attempt < 2:
                    time.sleep(0.4 * (attempt + 1))
        raise last_err

class _ResilientFilesProxy:
    def list(self, *args, **kwargs):
        return _ResilientRequestProxy("list", args, kwargs)
    def create(self, *args, **kwargs):
        return _ResilientRequestProxy("create", args, kwargs)
    def update(self, *args, **kwargs):
        return _ResilientRequestProxy("update", args, kwargs)
    def delete(self, *args, **kwargs):
        return _ResilientRequestProxy("delete", args, kwargs)
    def get_media(self, *args, **kwargs):
        with _drive_lock:
            raw_srv, err_msg = get_drive_service(force_reconnect=False)
            if not raw_srv:
                raise RuntimeError(err_msg or "Koneksi Google Drive belum siap.")
            return raw_srv.files().get_media(*args, **kwargs)

class _ResilientDriveService:
    """Proxy Service Google Drive yang 100% Thread-Safe dan tahan putus koneksi."""
    def files(self):
        return _ResilientFilesProxy()
    def is_connected(self):
        raw_srv, _ = get_drive_service(force_reconnect=False)
        if raw_srv is None:
            raw_srv, _ = get_drive_service(force_reconnect=True)
        return raw_srv is not None
    def __bool__(self):
        return self.is_connected()

# Inisialisasi Service dengan Proxy Auto-Reconnect
_raw_init, status_msg = get_drive_service()
service = _ResilientDriveService() if _raw_init is not None else None
if service is None:
    # Coba sambung ulang sekali lagi secara otomatis
    _raw_init, status_msg = get_drive_service(force_reconnect=True)
    if _raw_init is not None:
        service = _ResilientDriveService()

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

@st.cache_data(ttl=600, max_entries=64, show_spinner=False)
def find_subfolder(parent_id, folder_name):
    """Mencari subfolder di dalam parent_id tanpa membuatnya jika belum ada."""
    if not service or not parent_id:
        return None
    try:
        q = f"'{parent_id}' in parents and name = '{folder_name}' and mimeType = 'application/vnd.google-apps.folder' and trashed = false"
        res = service.files().list(q=q, fields="files(id, name)").execute()
        files = res.get('files', [])
        if files:
            return files[0]['id']
        return None
    except Exception:
        return None

@st.cache_data(ttl=600, max_entries=64, show_spinner=False)
def get_or_create_subfolder(parent_id, folder_name):
    """Mencari atau membuat subfolder di dalam parent_id jika belum ada."""
    if not service or not parent_id:
        return None
    try:
        found = find_subfolder(parent_id, folder_name)
        if found:
            return found
            
        file_metadata = {
            'name': folder_name,
            'mimeType': 'application/vnd.google-apps.folder',
            'parents': [parent_id]
        }
        created = service.files().create(body=file_metadata, fields='id').execute()
        # Bersihkan cache struktur folder saat ada folder baru yang dibuat
        find_subfolder.clear()
        get_search_target_folders.clear()
        get_available_years.clear()
        return created.get('id')
    except Exception as e:
        return None

def get_triwulan_folder_id(year, triwulan):
    """Mendapatkan ID subfolder Triwulan di dalam folder Tahun."""
    if not FOLDER_ID:
        return None
    year_folder_id = get_or_create_subfolder(FOLDER_ID, str(year))
    if not year_folder_id:
        return None
    tw_folder_id = get_or_create_subfolder(year_folder_id, triwulan)
    return tw_folder_id

def get_target_division_folder_id(year, triwulan, division):
    """Mendapatkan ID subfolder Divisi di dalam Tahun > Triwulan."""
    tw_folder_id = get_triwulan_folder_id(year, triwulan)
    if not tw_folder_id:
        return None
    div_folder_id = get_or_create_subfolder(tw_folder_id, str(division))
    return div_folder_id

def format_sandeq_filename(year, triwulan, filename):
    """Memformat nama file otomatis agar diawali {Tahun}_{TW}_{NamaFile}."""
    tw_code = TRIWULAN_CODE.get(triwulan, "TW1")
    prefix = f"{year}_{tw_code}_"
    clean_name = filename.strip()
    if clean_name.startswith(prefix):
        return clean_name
    return f"{prefix}{clean_name}"

@st.cache_data(ttl=600, max_entries=4, show_spinner=False)
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

# ==================== MANAJEMEN PENGGUNA & KEAMANAN ====================

def hash_password(password: str, salt: str = "sandeq_bps_sulbar_2026") -> str:
    """Mengacak password dengan salt menggunakan SHA-256."""
    return hashlib.sha256(f"{salt}{password}".encode("utf-8")).hexdigest()

def verify_password(plain_password: str, stored_hash: str, salt: str = "sandeq_bps_sulbar_2026") -> bool:
    """Memverifikasi kecocokan password dengan hash tersimpan."""
    if not stored_hash:
        return False
    if hash_password(plain_password, salt) == stored_hash:
        return True
    if plain_password == stored_hash:
        return True
    return False

@st.cache_data(ttl=300, max_entries=4, show_spinner=False)
def load_users_from_drive():
    """Mengambil database akun pengguna dari file sandeq_users.json di Google Drive."""
    DEFAULT_SEED = {
        "neraca": {
            "name": "Tim Neraca",
            "division": "Neraca",
            "role": "admin",
            "password_hash": hash_password("admin123")
        },
        "produksi": {
            "name": "Tim Produksi",
            "division": "Produksi",
            "role": "user",
            "password_hash": hash_password("user123")
        },
        "distribusi": {
            "name": "Tim Distribusi",
            "division": "Distribusi",
            "role": "user",
            "password_hash": hash_password("user123")
        },
        "ipds": {
            "name": "Tim IPDS",
            "division": "IPDS",
            "role": "user",
            "password_hash": hash_password("user123")
        },
        "sosial": {
            "name": "Tim Sosial",
            "division": "Sosial",
            "role": "user",
            "password_hash": hash_password("user123")
        }
    }
    
    if not service or not FOLDER_ID:
        return DEFAULT_SEED
        
    try:
        q = f"'{FOLDER_ID}' in parents and name = 'sandeq_users.json' and trashed = false"
        res = service.files().list(q=q, fields="files(id, name)").execute()
        files = res.get('files', [])
        
        if files:
            file_id = files[0]['id']
            last_dl_err = None
            for attempt in range(3):
                try:
                    with _drive_lock:
                        raw_srv, _ = get_drive_service(force_reconnect=(attempt > 0))
                        if not raw_srv:
                            raise RuntimeError("Koneksi Drive belum siap.")
                        request = raw_srv.files().get_media(fileId=file_id)
                        fh = io.BytesIO()
                        downloader = MediaIoBaseDownload(fh, request)
                        done = False
                        while not done:
                            status, done = downloader.next_chunk()
                    fh.seek(0)
                    users_dict = json.loads(fh.read().decode('utf-8'))
                    return users_dict
                except Exception as dl_e:
                    last_dl_err = dl_e
                    with _drive_lock:
                        _thread_local.raw_service = None
                    if attempt < 2:
                        time.sleep(0.4 * (attempt + 1))
            raise last_dl_err
        else:
            seed_users = DEFAULT_SEED.copy()
            if "users" in st.secrets:
                for uname, udata in st.secrets["users"].items():
                    if uname not in seed_users and isinstance(udata, dict):
                        u_div = udata.get("division", "Neraca")
                        seed_users[uname] = {
                            "name": udata.get("name", uname),
                            "division": u_div,
                            "role": "admin" if u_div.lower() == "neraca" else udata.get("role", "user"),
                            "password_hash": hash_password(str(udata.get("password", "user123")))
                        }
            json_bytes = json.dumps(seed_users, indent=2).encode('utf-8')
            media = MediaIoBaseUpload(io.BytesIO(json_bytes), mimetype='application/json', resumable=True)
            file_metadata = {
                'name': 'sandeq_users.json',
                'parents': [FOLDER_ID],
                'mimeType': 'application/json'
            }
            service.files().create(body=file_metadata, media_body=media, fields='id').execute()
            return seed_users
    except Exception as e:
        return DEFAULT_SEED

def save_users_to_drive(users_dict):
    """Menyimpan database pengguna ke sandeq_users.json di Google Drive."""
    if not service or not FOLDER_ID:
        return False, "Google Drive belum terhubung."
    try:
        json_bytes = json.dumps(users_dict, indent=2).encode('utf-8')
        media = MediaIoBaseUpload(io.BytesIO(json_bytes), mimetype='application/json', resumable=True)
        
        q = f"'{FOLDER_ID}' in parents and name = 'sandeq_users.json' and trashed = false"
        res = service.files().list(q=q, fields="files(id, name)").execute()
        files = res.get('files', [])
        
        if files:
            file_id = files[0]['id']
            service.files().update(fileId=file_id, media_body=media).execute()
        else:
            file_metadata = {
                'name': 'sandeq_users.json',
                'parents': [FOLDER_ID],
                'mimeType': 'application/json'
            }
            service.files().create(body=file_metadata, media_body=media, fields='id').execute()
            
        load_users_from_drive.clear()
        return True, "Berhasil"
    except Exception as e:
        return False, str(e)

def get_registered_divisions():
    """Mengambil daftar semua divisi yang terdaftar di database users."""
    DEFAULT_DIVS = ["Neraca", "Produksi", "Distribusi", "IPDS", "Sosial"]
    try:
        users_dict = load_users_from_drive()
        divs = set(DEFAULT_DIVS)
        for u in users_dict.values():
            if isinstance(u, dict) and "division" in u and u["division"]:
                divs.add(u["division"])
        return sorted(list(divs))
    except Exception:
        return DEFAULT_DIVS

@st.cache_data(ttl=600, max_entries=32, show_spinner=False)
def get_search_target_folders(year="Semua", triwulan="Semua", division="Semua"):
    """
    Mengambil daftar folder ID target berdasarkan filter Tahun, Triwulan, dan Divisi.
    Dioptimalkan dengan query batch per-level (memangkas 16-20 panggilan API menjadi maksimal 3 panggilan cepat)
    dan di-cache selama 10 menit.
    """
    if not service or not FOLDER_ID:
        return []
        
    # Kasus spesifik: Tahun, Triwulan, dan Divisi tertentu
    if year != "Semua" and triwulan != "Semua" and division != "Semua":
        yf_id = find_subfolder(FOLDER_ID, str(year))
        if not yf_id:
            return []
        tw_id = find_subfolder(yf_id, triwulan)
        if not tw_id:
            return []
        div_id = find_subfolder(tw_id, division)
        return [div_id] if div_id else []
        
    target_ids = []
    
    # 1. Batch Level Tahun
    try:
        q = f"'{FOLDER_ID}' in parents and mimeType = 'application/vnd.google-apps.folder' and trashed = false"
        res = service.files().list(q=q, pageSize=100, fields="files(id, name)").execute()
        year_folders = res.get('files', [])
    except Exception:
        year_folders = []
        
    if year != "Semua":
        matched_years = [f for f in year_folders if f['name'] == str(year)]
    else:
        matched_years = year_folders
        # Jika mencari 'Semua Divisi' dan 'Semua Tahun', sertakan FOLDER_ID root untuk file lama
        if division == "Semua":
            target_ids.append(FOLDER_ID)
            
    if not matched_years:
        return target_ids
        
    # 2. Batch Level Triwulan (Sekaligus untuk semua tahun yang cocok dalam 1 panggilan API)
    try:
        yr_parents_q = "(" + " or ".join([f"'{yf['id']}' in parents" for yf in matched_years]) + ")"
        q_tw = f"{yr_parents_q} and mimeType = 'application/vnd.google-apps.folder' and trashed = false"
        if triwulan != "Semua":
            q_tw += f" and name = '{triwulan}'"
        res_tw = service.files().list(q=q_tw, pageSize=200, fields="files(id, name, parents)").execute()
        tw_folders = res_tw.get('files', [])
    except Exception:
        tw_folders = []
        
    if not tw_folders:
        return target_ids
        
    # 3. Batch Level Divisi (Sekaligus untuk semua triwulan yang cocok dalam 1 panggilan API)
    try:
        tw_parents_q = "(" + " or ".join([f"'{twf['id']}' in parents" for twf in tw_folders]) + ")"
        q_div = f"{tw_parents_q} and mimeType = 'application/vnd.google-apps.folder' and trashed = false"
        res_div = service.files().list(q=q_div, pageSize=500, fields="files(id, name, parents)").execute()
        div_folders = res_div.get('files', [])
        
        tw_with_divs = set()
        for df in div_folders:
            for p_id in df.get('parents', []):
                tw_with_divs.add(p_id)
            if division == "Semua" or df['name'] == division:
                target_ids.append(df['id'])
                
        # Jika ada folder triwulan yang belum punya subfolder divisi sama sekali
        if division == "Semua":
            for twf in tw_folders:
                if twf['id'] not in tw_with_divs:
                    target_ids.append(twf['id'])
    except Exception:
        if division == "Semua":
            for twf in tw_folders:
                target_ids.append(twf['id'])
                
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

def delete_file_from_drive(file_id):
    """Menghapus file secara permanen dari Google Drive."""
    if not service:
        return False, "Google Drive belum terkonfigurasi."
    try:
        service.files().delete(fileId=file_id).execute()
        return True, "File berhasil dihapus."
    except Exception as e:
        return False, str(e)

@st.cache_data(ttl=300, max_entries=32, show_spinner=False)
def search_files_in_drive(query="", mime_type=None, app_properties=None, year="Semua", triwulan="Semua", division="Semua"):
    if not service or not FOLDER_ID:
        return []
    try:
        target_folders = get_search_target_folders(year, triwulan, division)
        if not target_folders:
            return []
            
        if len(target_folders) == 1:
            parents_q = f"'{target_folders[0]}' in parents"
        else:
            parents_q = "(" + " or ".join([f"'{fid}' in parents" for fid in target_folders]) + ")"
            
        q = f"{parents_q} and trashed = false and mimeType != 'application/vnd.google-apps.folder' and name != 'sandeq_users.json'"
        if query:
            q += f" and name contains '{query}'"
        if mime_type:
            q += f" and mimeType = '{mime_type}'"
        if app_properties:
            for k, v in app_properties.items():
                q += f" and appProperties has {{ key='{k}' and value='{v}' }}"
                
        results = service.files().list(
            q=q, pageSize=500, fields="nextPageToken, files(id, name, mimeType, size, createdTime, webContentLink, webViewLink, appProperties, parents)"
        ).execute()
        return results.get('files', [])
    except Exception as e:
        if type(e).__name__ in ("RerunException", "StopException"):
            raise e
        auto_recover_system(f"Gagal mengambil data Drive: {e}")
        return []

def download_file_from_drive(file_id):
    if not service:
        return None
    last_err = None
    for attempt in range(3):
        try:
            with _drive_lock:
                raw_srv, _ = get_drive_service(force_reconnect=(attempt > 0))
                if not raw_srv:
                    raise RuntimeError("Koneksi Drive belum siap.")
                request = raw_srv.files().get_media(fileId=file_id)
                fh = io.BytesIO()
                downloader = MediaIoBaseDownload(fh, request)
                done = False
                while done is False:
                    status, done = downloader.next_chunk()
            fh.seek(0)
            return fh
        except Exception as e:
            last_err = e
            with _drive_lock:
                _thread_local.raw_service = None
            if attempt < 2:
                time.sleep(0.4 * (attempt + 1))
    auto_recover_system(f"Gagal mengunduh file: {last_err}")
    return None

@st.cache_data(ttl=300, max_entries=10, show_spinner=False)
def load_excel_dataframe_from_drive(file_id):
    """Mengunduh dan membaca file Excel ke DataFrame dengan cache hemat memori (max 10 file di RAM)."""
    file_stream = download_file_from_drive(file_id)
    if not file_stream:
        return None
    return pd.read_excel(file_stream, engine='openpyxl')

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
    props = file.get('appProperties') or {}
    
    user_info = st.session_state.get('user_info', {})
    user_div = user_info.get('division', '')
    is_admin = (user_info.get('role') == 'admin' or user_div == 'Neraca')
    user_uname = st.session_state.get('username', '')
    
    file_div = props.get('division', '')
    uploader_name = props.get('uploader_name', props.get('uploader', ''))
    
    # Cek hak akses hapus file: Admin atau pemilik/divisi yang mengunggah
    can_delete = is_admin or (file_div == user_div) or (props.get('uploader') == user_uname)
    
    with st.container():
        st.markdown('<div class="file-card">', unsafe_allow_html=True)
        col1, col2, col3 = st.columns([3, 1, 0.8])
        with col1:
            st.markdown(f"**📄 {file_name}**")
            
            # Badge info pengunggah & divisi
            if file_div:
                badge_info = f"🏢 Divisi: **{file_div}**"
                if uploader_name:
                    badge_info += f" &nbsp;|&nbsp; 👤 Oleh: **{uploader_name}**"
            else:
                badge_info = "📁 *File Database Awal (Khusus Admin)*"
                
            st.caption(f"📅 {format_date(file.get('createdTime'))} &nbsp;|&nbsp; 💾 {format_size(file.get('size'))} &nbsp;|&nbsp; {badge_info}")
            
        with col2:
            link = web_link or file.get('webViewLink')
            if link:
                # Tombol unduh langsung Google Drive
                st.link_button("⬇️ Unduh", link, type="primary", width="stretch")
            else:
                st.button("File Tidak Tersedia", disabled=True, width="stretch")
                
        with col3:
            if can_delete:
                with st.popover("🗑️", width="stretch", help="Hapus file ini"):
                    st.write(f"Hapus permanen file **{file_name}**?")
                    if st.button("Ya, Hapus", key=f"del_{file_id}", type="primary", width="stretch"):
                        ok, msg = delete_file_from_drive(file_id)
                        if ok:
                            st.success("File berhasil dihapus.")
                            search_files_in_drive.clear()
                            st.rerun()
                        else:
                            st.error(f"Gagal: {msg}")
            else:
                st.write("")
        st.markdown('</div>', unsafe_allow_html=True)


# ==================== OTENTIKASI & KEAMANAN ====================

def login_screen():
    st.markdown("""
    <div style="text-align: center; margin-top: 30px; margin-bottom: 20px;">
        <h1 style="color: #1E88E5; font-size: 2.6rem; margin-bottom: 5px;">SANDEQ 🗄️</h1>
        <p style="color: #555; font-size: 1.15rem; margin-top: 0;">Sistem Administrasi Data Elektronik Quartal PDRB<br><b>BPS Provinsi Sulawesi Barat</b></p>
    </div>
    """, unsafe_allow_html=True)
    
    col_l, col_m, col_r = st.columns([1, 1.4, 1])
    with col_m:
        with st.container(border=True):
            tab_login, tab_register = st.tabs(["🔐 Masuk", "📝 Buat Akun Baru"])
            
            with tab_login:
                st.subheader("Masuk ke Dashboard")
                st.caption("Masukkan username dan password akun divisi Anda.")
                with st.form("form_masuk", clear_on_submit=False):
                    username_input = st.text_input("Username:", placeholder="Contoh: neraca atau produksi")
                    password_input = st.text_input("Password:", type="password", placeholder="••••••••")
                    submitted_login = st.form_submit_button("Masuk ke Dashboard 🚀", type="primary", width="stretch")
                    
                    if submitted_login:
                        user_clean = username_input.strip().lower()
                        users_dict = load_users_from_drive()
                        
                        if user_clean in users_dict:
                            user_data = users_dict[user_clean]
                            stored_hash = user_data.get("password_hash") or user_data.get("password", "")
                            
                            if verify_password(password_input, stored_hash):
                                u_div = user_data.get("division", "Umum")
                                u_role = "admin" if u_div.lower() == "neraca" else user_data.get("role", "user")
                                
                                st.session_state["authenticated"] = True
                                st.session_state["username"] = user_clean
                                st.session_state["user_info"] = {
                                    "name": user_data.get("name", user_clean),
                                    "division": u_div,
                                    "role": u_role
                                }
                                search_files_in_drive.clear()
                                st.success(f"Selamat datang, {user_data.get('name', user_clean)}!")
                                st.rerun()
                            else:
                                st.error("Password yang Anda masukkan salah.")
                        else:
                            st.error("Username tidak ditemukan dalam sistem.")
                            
            with tab_register:
                st.subheader("Daftar Akun Baru")
                st.caption("Akun yang baru dibuat akan langsung aktif dan tersimpan di database Google Drive.")
                with st.form("form_daftar", clear_on_submit=False):
                    reg_name = st.text_input("Nama Lengkap:", placeholder="Contoh: Budi Santoso")
                    reg_username = st.text_input("Username:", placeholder="Contoh: budi_santoso (tanpa spasi)").strip().lower()
                    reg_division = st.selectbox(
                        "Pilih Divisi:", 
                        ["Neraca", "Produksi", "Distribusi", "IPDS", "Sosial"],
                        help="Akun dengan divisi 'Neraca' otomatis memiliki hak akses Administrator (dapat melihat dan mengelola seluruh data sistem)."
                    )   
                    reg_password = st.text_input("Kata Sandi (Password):", type="password", placeholder="Minimal 6 karakter")
                    reg_password2 = st.text_input("Konfirmasi Kata Sandi:", type="password", placeholder="Ulangi kata sandi")
                    submitted_register = st.form_submit_button("Daftar Akun", type="primary", width="stretch")
                    
                    if submitted_register:
                        if not reg_name.strip():
                            st.error("Mohon isi Nama Lengkap")
                        elif len(reg_username) < 3 or not re.match(r"^[a-zA-Z0-9_]+$", reg_username):
                            st.error("Username harus minimal 3 karakter dan hanya boleh berisi huruf, angka, atau garis bawah (_).")
                        elif len(reg_password) < 6:
                            st.error("Kata sandi harus minimal 6 karakter.")
                        elif reg_password != reg_password2:
                            st.error("Konfirmasi kata sandi tidak cocok.")
                        else:
                            users_dict = load_users_from_drive()
                            if reg_username in users_dict:
                                st.error(f"Username '{reg_username}' sudah digunakan. Silakan pilih username lain.")
                            else:
                                reg_role = "admin" if reg_division == "Neraca" else "user"
                                new_user_obj = {
                                    "name": reg_name.strip(),
                                    "division": reg_division,
                                    "role": reg_role,
                                    "password_hash": hash_password(reg_password),
                                    "created_at": datetime.now().strftime("%Y-%m-%d %H:%M:%S")
                                }
                                users_dict[reg_username] = new_user_obj
                                ok, err = save_users_to_drive(users_dict)
                                if ok:
                                    st.session_state["authenticated"] = True
                                    st.session_state["username"] = reg_username
                                    st.session_state["user_info"] = {
                                        "name": new_user_obj["name"],
                                        "division": new_user_obj["division"],
                                        "role": new_user_obj["role"]
                                    }
                                    search_files_in_drive.clear()
                                    st.success(f"Akun '{reg_username}' berhasil dibuat! Mengalihkan ke dashboard...")
                                    st.rerun()
                                else:
                                    st.error(f"Gagal menyimpan akun ke Google Drive: {err}")

# Pengecekan status login (Gatekeeper) dengan pelindung Self-Healing
if not st.session_state.get("authenticated", False):
    st.markdown("""
    <style>
        [data-testid="stSidebar"] {
            display: none !important;
        }
    </style>
    """, unsafe_allow_html=True)
    try:
        login_screen()
    except Exception as e:
        if type(e).__name__ in ("RerunException", "StopException"):
            raise e
        auto_recover_system(str(e))
    st.stop()


# ==================== FUNGSI HALAMAN ====================

def page_beranda():
    user_info = st.session_state.get('user_info', {})
    user_div = user_info.get('division', 'Neraca')
    is_admin = (user_info.get('role') == 'admin' or user_div == 'Neraca')
    
    st.markdown('<p class="main-header">Selamat Datang di SANDEQ 🗄️</p>', unsafe_allow_html=True)
    st.markdown('<p class="sub-header">Sistem Administrasi Data Elektronik Quartal PDRB - BPS Provinsi Sulawesi Barat</p>', unsafe_allow_html=True)
    
    col1, col2, col3 = st.columns(3)
    
    if is_admin:
        files = search_files_in_drive("", division="Semua")
        stat_title = "Total File (Semua Divisi)"
        cap_title = "Total Kapasitas (Semua Divisi)"
    else:
        raw_files = search_files_in_drive("", division=user_div)
        files = [f for f in raw_files if (f.get('appProperties') or {}).get('division') == user_div]
        stat_title = f"Total File (Divisi {user_div})"
        cap_title = f"Kapasitas Terpakai ({user_div})"
        
    total_files = len(files)
    total_size = sum([int(f.get('size', 0)) for f in files])
    
    with col1:
        st.info(f"### 📂 {total_files}\n**{stat_title}**")
    with col2:
        st.success(f"### 💾 {format_size(total_size)}\n**{cap_title}**")
    with col3:
        st.warning(f"### ⏱️ {datetime.now().strftime('%H:%M')}\n**Waktu Akses Terakhir**")

    st.markdown("---")
    st.markdown("### 📌 Fitur Utama")
    st.markdown("""
    - **📂 Semua Data**: Lihat seluruh file berdasarkan filter Tahun dan Triwulan.
    - **📤 Upload File**: Unggah file dari perangkat Anda ke folder Tahun > Triwulan > Divisi terkait.
    - **📝 Isi Data Form**: Buat tabel Excel fleksibel langsung dari dashboard dan simpan per Triwulan.
    - **📊 Tarik/Ekstrak Data**: Buka, saring, dan unduh data Excel berdasarkan periode Triwulan.
    """)

def page_semua_data():
    st.markdown('<p class="main-header">Semua Data</p>', unsafe_allow_html=True)
    st.markdown("Lihat, cari, dan kelola semua file yang ada di penyimpanan Drive Anda berdasarkan Tahun dan Triwulan.")
    st.markdown("---")
    
    user_info = st.session_state.get('user_info', {})
    user_div = user_info.get('division', 'Neraca')
    is_admin = (user_info.get('role') == 'admin' or user_div == 'Neraca')
    
    # Filter pencarian, Tahun, Triwulan, Divisi (jika admin), dan Pengurutan
    if is_admin:
        col_search, col_yr, col_tw, col_div, col_sort = st.columns([2.2, 1, 1, 1, 1])
        with col_search:
            search_query = st.text_input("🔍 Cari File:", placeholder="Ketik nama file yang ingin dicari...")
        with col_yr:
            filter_year = st.selectbox("Tahun:", ["Semua"] + get_available_years())
        with col_tw:
            filter_tw = st.selectbox("Triwulan:", ["Semua"] + TRIWULAN_LIST)
        with col_div:
            registered_divs = get_registered_divisions()
            filter_div = st.selectbox("Divisi:", ["Semua"] + registered_divs)
        with col_sort:
            sort_by = st.selectbox("Urutkan:", ["Terbaru", "Terlama", "Nama (A-Z)"])
        
        with st.spinner("Memuat data dari Drive..."):
            # Gunakan query="" agar berbagi cache penuh dengan halaman Beranda (0 panggilan API saat cari nama file)
            files = search_files_in_drive("", year=filter_year, triwulan=filter_tw, division=filter_div)
            if filter_div != "Semua":
                files = [f for f in files if (f.get('appProperties') or {}).get('division') == filter_div]
    else:
        col_search, col_yr, col_tw, col_sort = st.columns([2.5, 1, 1, 1])
        with col_search:
            search_query = st.text_input("🔍 Cari File:", placeholder="Ketik nama file yang ingin dicari...")
        with col_yr:
            filter_year = st.selectbox("Tahun:", ["Semua"] + get_available_years())
        with col_tw:
            filter_tw = st.selectbox("Triwulan:", ["Semua"] + TRIWULAN_LIST)
        with col_sort:
            sort_by = st.selectbox("Urutkan:", ["Terbaru", "Terlama", "Nama (A-Z)"])
            
        with st.spinner(f"Memuat data Divisi {user_div} dari Drive..."):
            raw_files = search_files_in_drive("", year=filter_year, triwulan=filter_tw, division=user_div)
            files = [f for f in raw_files if (f.get('appProperties') or {}).get('division') == user_div]
            
    # Filter nama file secara instan di memori (tanpa memanggil ulang API Google Drive)
    if search_query:
        q_clean = search_query.strip().lower()
        files = [f for f in files if q_clean in f.get('name', '').lower()]
            
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
        
        # Pagination (15 file per halaman agar render browser & RAM server selalu ringan)
        ITEMS_PER_PAGE = 15
        total_files = len(files)
        total_pages = max(1, (total_files + ITEMS_PER_PAGE - 1) // ITEMS_PER_PAGE)
        
        if total_pages > 1:
            col_info, col_page = st.columns([3, 1])
            with col_page:
                current_page = st.number_input(
                    f"Halaman (dari {total_pages}):",
                    min_value=1,
                    max_value=total_pages,
                    value=1,
                    step=1,
                    key="semua_data_page"
                )
            start_idx = (current_page - 1) * ITEMS_PER_PAGE
            end_idx = min(start_idx + ITEMS_PER_PAGE, total_files)
            with col_info:
                st.write(f"Menampilkan **{start_idx + 1}–{end_idx}** dari total **{total_files}** file (Halaman {current_page}/{total_pages}).")
            paged_files = files[start_idx:end_idx]
        else:
            st.write(f"Menampilkan **{total_files}** file.")
            paged_files = files
        
        for file in paged_files:
            display_file_card(file)
        
        if st.button("🔄 Segarkan Data"):
            search_files_in_drive.clear()
            get_search_target_folders.clear()
            st.rerun()

def page_upload_data():
    st.markdown('<p class="main-header">Upload File</p>', unsafe_allow_html=True)
    st.markdown("Pilih satu atau beberapa file dari komputer Anda untuk disimpan ke folder Tahun, Triwulan, dan Divisi terkait.")
    st.markdown("---")
    
    user_info = st.session_state.get('user_info', {})
    user_div = user_info.get('division', 'Neraca')
    is_admin = (user_info.get('role') == 'admin' or user_div == 'Neraca')
    user_uname = st.session_state.get('username', '')
    user_realname = user_info.get('name', user_uname)
    
    # Pilih Tahun, Triwulan, dan Divisi Target
    if is_admin:
        col_yr, col_tw, col_div = st.columns(3)
        with col_yr:
            avail_years = get_available_years()
            curr_year = str(datetime.now().year)
            yr_idx = avail_years.index(curr_year) if curr_year in avail_years else 0
            upload_year = st.selectbox("📅 Pilih Tahun:", avail_years, index=yr_idx)
        with col_tw:
            upload_tw = st.selectbox("📊 Pilih Triwulan:", TRIWULAN_LIST)
        with col_div:
            registered_divs = get_registered_divisions()
            div_idx = registered_divs.index(user_div) if user_div in registered_divs else 0
            upload_div = st.selectbox("🏢 Divisi Target:", registered_divs, index=div_idx)
    else:
        col_yr, col_tw, col_div = st.columns(3)
        with col_yr:
            avail_years = get_available_years()
            curr_year = str(datetime.now().year)
            yr_idx = avail_years.index(curr_year) if curr_year in avail_years else 0
            upload_year = st.selectbox("📅 Pilih Tahun:", avail_years, index=yr_idx)
        with col_tw:
            upload_tw = st.selectbox("📊 Pilih Triwulan:", TRIWULAN_LIST)
        with col_div:
            upload_div = user_div
            st.text_input("🏢 Divisi Pengunggah:", value=user_div, disabled=True)
        
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
            st.info(f"📁 Terpilih **{len(uploaded_files)}** file. Nama file akan otomatis diawali format: `{upload_year}_{tw_code}_` dan disimpan ke folder **{upload_year} > {upload_tw} > {upload_div}**.")
            col1, col2 = st.columns([1, 4])
            with col1:
                if st.button("☁️ Mulai Upload Semua", type="primary", width="stretch"):
                    if service and FOLDER_ID:
                        total_files = len(uploaded_files)
                        progress_bar = st.progress(0)
                        status_text = st.empty()
                        
                        berhasil = 0
                        gagal = []
                        
                        # Dapatkan folder ID divisi tujuan (otomatis buat jika belum ada)
                        target_folder = get_target_division_folder_id(upload_year, upload_tw, upload_div)
                        if not target_folder:
                            st.error(f"Gagal menyiapkan folder Google Drive untuk {upload_year} > {upload_tw} > {upload_div}.")
                            return
                        
                        app_props = {
                            "division": str(upload_div),
                            "uploader": str(user_uname),
                            "uploader_name": str(user_realname),
                            "year": str(upload_year),
                            "triwulan": str(upload_tw)
                        }
                        
                        for idx, file in enumerate(uploaded_files):
                            final_name = format_sandeq_filename(upload_year, upload_tw, file.name)
                            status_text.text(f"Mengunggah ({idx + 1}/{total_files}): {final_name}...")
                            success, msg = upload_file_to_drive(file, final_name, file.type, target_folder_id=target_folder, app_properties=app_props)
                            if success:
                                berhasil += 1
                            else:
                                gagal.append((final_name, msg))
                            progress_bar.progress((idx + 1) / total_files)
                            
                        # Hapus cache pencarian agar data terbaru langsung muncul
                        search_files_in_drive.clear()
                        
                        if berhasil > 0:
                            st.success(f"✅ Berhasil mengunggah **{berhasil}** file ke folder **{upload_year} > {upload_tw} > {upload_div}**!")
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
        
    user_info = st.session_state.get('user_info', {})
    user_div = user_info.get('division', 'Neraca')
    is_admin = (user_info.get('role') == 'admin' or user_div == 'Neraca')
    user_uname = st.session_state.get('username', '')
    user_realname = user_info.get('name', user_uname)
    
    # Kunci unik agar data_editor bisa di-reset sepenuhnya
    if "canvas_key" not in st.session_state:
        st.session_state.canvas_key = 1
        
    # Setup kanvas kosong (misal 10 kolom A s.d. J)
    if "canvas_data" not in st.session_state:
        cols = ['A', 'B', 'C', 'D', 'E', 'F', 'G', 'H', 'I', 'J']
        empty_data = [["" for _ in cols] for _ in range(15)]
        st.session_state.canvas_data = pd.DataFrame(empty_data, columns=cols)
        
    st.info("💡 **Tips:** Baris No.1 otomatis akan dijadikan *Header* (Judul Kolom) saat file disimpan. Kolom yang sama sekali kosong tidak akan ikut disimpan.")
    
    # Render data editor dengan key yang dinamis
    edited_df = st.data_editor(
        st.session_state.canvas_data, 
        num_rows="dynamic",
        width="stretch",
        key=f"canvas_editor_{st.session_state.canvas_key}"
    )
    
    st.markdown("---")
    st.markdown("### 📅 Pengaturan Periode & Divisi Data")
    if is_admin:
        col_yr, col_tw, col_div = st.columns(3)
        with col_yr:
            avail_years = get_available_years()
            curr_year = str(datetime.now().year)
            yr_idx = avail_years.index(curr_year) if curr_year in avail_years else 0
            form_year = st.selectbox("Tahun Data:", avail_years, index=yr_idx, key="form_year")
        with col_tw:
            form_tw = st.selectbox("Triwulan Data:", TRIWULAN_LIST, key="form_tw")
        with col_div:
            registered_divs = get_registered_divisions()
            div_idx = registered_divs.index(user_div) if user_div in registered_divs else 0
            form_div = st.selectbox("🏢 Divisi Pemilik Data:", registered_divs, index=div_idx, key="form_div")
    else:
        col_yr, col_tw, col_div = st.columns(3)
        with col_yr:
            avail_years = get_available_years()
            curr_year = str(datetime.now().year)
            yr_idx = avail_years.index(curr_year) if curr_year in avail_years else 0
            form_year = st.selectbox("Tahun Data:", avail_years, index=yr_idx, key="form_year")
        with col_tw:
            form_tw = st.selectbox("Triwulan Data:", TRIWULAN_LIST, key="form_tw")
        with col_div:
            form_div = user_div
            st.text_input("🏢 Divisi Pemilik Data:", value=user_div, disabled=True, key="form_div_disp")
    
    st.markdown("---")
    
    col_name, col_btn = st.columns([3, 1])
    with col_name:
        tw_code = TRIWULAN_CODE.get(form_tw, "TW1")
        file_name_input = st.text_input("Nama File Excel Anda:", placeholder="Contoh: Produksi_Padi", help=f"Nama file otomatis diawali '{form_year}_{tw_code}_'. Tidak perlu menuliskan .xlsx.")
    with col_btn:
        st.write("") # Spacer agar sejajar dengan input
        st.write("")
        simpan_btn = st.button("💾 Simpan", type="primary", width="stretch")
        
    if simpan_btn:
        if file_name_input.strip() == "":
            st.warning("Silakan masukkan nama file terlebih dahulu.")
            return
            
        with st.spinner("Memproses tabel dan menyimpan ke Drive..."):
            try:
                import numpy as np
                cleaned_df = edited_df.replace("", np.nan)
                cleaned_df = cleaned_df.dropna(how='all', axis=0).dropna(how='all', axis=1)
                
                if cleaned_df.empty:
                    st.warning("Tabel masih benar-benar kosong.")
                    return
                
                header_row = cleaned_df.iloc[0]
                data_rows = cleaned_df.iloc[1:]
                data_rows.columns = header_row
                
                output = io.BytesIO()
                with pd.ExcelWriter(output, engine='openpyxl') as writer:
                    data_rows.to_excel(writer, index=False, sheet_name='Sheet1')
                output.seek(0)
                
                base_name = file_name_input.strip()
                if not base_name.endswith('.xlsx'):
                    base_name += ".xlsx"
                final_filename = format_sandeq_filename(form_year, form_tw, base_name)
                mime = "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
                
                target_folder = get_target_division_folder_id(form_year, form_tw, form_div)
                if not target_folder:
                    st.error(f"Gagal menyiapkan folder Google Drive untuk {form_year} > {form_tw} > {form_div}.")
                    return
                
                app_props = {
                    "source": "isi_data_form",
                    "year": str(form_year),
                    "triwulan": str(form_tw),
                    "division": str(form_div),
                    "uploader": str(user_uname),
                    "uploader_name": str(user_realname)
                }
                
                success, msg = upload_file_to_drive(output, final_filename, mime, target_folder_id=target_folder, app_properties=app_props)
                
                if success:
                    st.success(f"✅ Berhasil! File '{msg}' telah disimpan ke folder {form_year} > {form_tw} > {form_div} di Google Drive.")
                    search_files_in_drive.clear()
                    
                    del st.session_state.canvas_data
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
        
    user_info = st.session_state.get('user_info', {})
    user_div = user_info.get('division', 'Neraca')
    is_admin = (user_info.get('role') == 'admin' or user_div == 'Neraca')
    
    excel_mime = "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
    
    if is_admin:
        col_yr, col_tw, col_div = st.columns(3)
        with col_yr:
            filter_year = st.selectbox("Filter Tahun:", ["Semua"] + get_available_years(), key="ekstrak_year")
        with col_tw:
            filter_tw = st.selectbox("Filter Triwulan:", ["Semua"] + TRIWULAN_LIST, key="ekstrak_tw")
        with col_div:
            registered_divs = get_registered_divisions()
            filter_div = st.selectbox("Filter Divisi:", ["Semua"] + registered_divs, key="ekstrak_div")
            
        with st.spinner("Mencari daftar file formulir Excel di Drive..."):
            files = search_files_in_drive("", mime_type=excel_mime, app_properties={"source": "isi_data_form"}, year=filter_year, triwulan=filter_tw, division=filter_div)
            if filter_div != "Semua":
                files = [f for f in files if (f.get('appProperties') or {}).get('division') == filter_div]
    else:
        col_yr, col_tw = st.columns(2)
        with col_yr:
            filter_year = st.selectbox("Filter Tahun:", ["Semua"] + get_available_years(), key="ekstrak_year")
        with col_tw:
            filter_tw = st.selectbox("Filter Triwulan:", ["Semua"] + TRIWULAN_LIST, key="ekstrak_tw")
            
        with st.spinner(f"Mencari daftar file formulir Excel Divisi {user_div} di Drive..."):
            raw_files = search_files_in_drive("", mime_type=excel_mime, app_properties={"source": "isi_data_form"}, year=filter_year, triwulan=filter_tw, division=user_div)
            files = [f for f in raw_files if (f.get('appProperties') or {}).get('division') == user_div]
        
    if not files:
        st.info("Belum ada file Excel formulir yang cocok dengan filter periode yang dipilih.")
        return
        
    # Buat pilihan file dengan label informatif
    file_options = {}
    for f in files:
        props = f.get('appProperties') or {}
        div_lbl = f" [{props.get('division')}]" if props.get('division') else ""
        label = f"{f.get('name')}{div_lbl}  (Dibuat: {format_date(f.get('createdTime'))})"
        file_options[label] = f.get('id')
        
    pilihan_file = st.selectbox("Pilih File Excel yang ingin dibuka:", list(file_options.keys()))
    file_id = file_options[pilihan_file]
    
    st.markdown("---")
    
    if file_id:
        with st.spinner(f"Membuka {pilihan_file}..."):
            try:
                df = load_excel_dataframe_from_drive(file_id)
                if df is None:
                    st.error("Gagal mengunduh file.")
                    return
                
                if df.empty:
                    st.warning("File Excel ini kosong.")
                    return
                    
                st.markdown("### 🔍 Filter Data")
                search_kw = st.text_input("Ketik kata kunci untuk mencari di kolom apa saja:", placeholder="Contoh: Budi, Padi, 2026...")
                
                df_filtered = df.copy()
                if search_kw:
                    mask = df_filtered.apply(lambda row: row.astype(str).str.contains(search_kw, case=False).any(), axis=1)
                    df_filtered = df_filtered[mask]
                    
                st.markdown("### 📊 Pratinjau Tabel")
                st.dataframe(df_filtered, width="stretch")
                st.caption(f"Menampilkan {len(df_filtered)} baris dari total {len(df)} baris data.")
                
                if not df_filtered.empty:
                    output = io.BytesIO()
                    with pd.ExcelWriter(output, engine='openpyxl') as writer:
                        df_filtered.to_excel(writer, index=False, sheet_name='Filtered_Data')
                    output.seek(0)
                    
                    nama_asli = pilihan_file.split('  (')[0].split(' [')[0].strip()
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
                            width="stretch"
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
    st.markdown('<div style="height: 10vh;"></div>', unsafe_allow_html=True)
    st.markdown("---")
    
    # Informasi Profil Pengguna
    user_info = st.session_state.get('user_info', {})
    u_name = user_info.get('name', st.session_state.get('username', 'Pengguna'))
    u_div = user_info.get('division', '-')
    is_admin = (user_info.get('role') == 'admin' or u_div == 'Neraca')
    
    st.markdown(f"""
    <div style="background: #eef2f6; padding: 12px; border-radius: 8px; border-left: 4px solid #1E88E5; margin-bottom: 10px;">
        <div style="font-size: 0.8rem; color: #666; font-weight: 600;">LOGIN SEBAGAI:</div>
        <div style="font-size: 1.05rem; font-weight: 700; color: #1a202c;">{u_name}</div>
        <div style="font-size: 0.85rem; color: #4a5568; margin-top: 3px;">🏢 Divisi: <b>{u_div}</b> &nbsp;</div>
    </div>
    """, unsafe_allow_html=True)
    
    if st.button("🚪 Keluar (Logout)", width="stretch"):
        st.session_state.clear()
        reset_all_caches_and_connections()
        st.rerun()
        
    st.markdown("---")
    st.markdown("### Status Sistem")
    if not service:
        # Coba pulihkan otomatis jika terdeteksi terputus
        auto_recover_system("Koneksi Drive terputus di Sidebar")
        st.error("🔴 Terputus")
    elif not FOLDER_ID:
        st.warning("🟡 Menunggu Folder ID")
    else:
        st.success("🟢 Terhubung ke Database")
        
    if st.button("🔄 Pulihkan / Bersihkan Cache", width="stretch", help="Klik jika aplikasi terasa lambat atau data belum terbarui"):
        st.session_state["_auto_recover_count"] = 0
        st.session_state["_last_auto_recover_ts"] = 0.0
        reset_all_caches_and_connections()
        st.rerun()
        
    st.markdown("---")
    st.caption("© 2026 SANDEQ - BPS")

# Jalankan navigasi utama dengan pelindung Self-Healing (Auto-Rerun & Auto-Clear Cache saat error)
try:
    pg.run()
    # Jika render halaman berhasil tanpa kendala, reset penghitung auto-recovery
    st.session_state["_auto_recover_count"] = 0
except Exception as e:
    if type(e).__name__ in ("RerunException", "StopException"):
        raise e
    auto_recover_system(str(e))

