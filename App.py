import datetime
import pandas as pd
import streamlit as st
import cloudinary
import cloudinary.uploader
import json
import os
import urllib.parse
import gspread
from google.oauth2.service_account import Credentials
from fpdf import FPDF
import requests
from io import BytesIO
from PIL import Image
import base64
import tempfile # TAMBAHAN PENTING UNTUK FIX PDF

# Mengimpor modul PPTX
try:
    from pptx import Presentation
    from pptx.util import Inches, Pt
    from pptx.dml.color import RGBColor
    HAS_PPTX = True
except ImportError:
    HAS_PPTX = False

# -------------------------------------------------------------------------
# SETUP HALAMAN & FUNGSI LOGO
# -------------------------------------------------------------------------
st.set_page_config(page_title="Report SPS - Okta Pradika", page_icon="⚡", layout="wide")

def get_base64_of_bin_file(bin_file):
    try:
        with open(bin_file, 'rb') as f:
            data = f.read()
        return base64.b64encode(data).decode()
    except:
        return None

# Fungsi untuk memunculkan logo presisi di tengah atas halaman
def render_header_logo():
    bg_base64 = get_base64_of_bin_file("logo.png")
    if bg_base64:
        st.markdown(
            f"""
            <div style="display: flex; justify-content: center; align-items: center; margin-bottom: 20px;">
                <img src="data:image/png;base64,{bg_base64}" style="max-height: 120px; width: auto; object-fit: contain;">
            </div>
            """, 
            unsafe_allow_html=True
        )

# -------------------------------------------------------------------------
# CUSTOM CSS (ADAPTIF UNTUK LIGHT MODE & DARK MODE)
# -------------------------------------------------------------------------
st.markdown("""
    <style>
        h1, h2, h3 { color: var(--primary-color) !important; font-family: 'Segoe UI', sans-serif; }
        
        div[data-testid="stExpander"] details {
            border: 1px solid var(--primary-color); 
            border-radius: 10px; 
            background-color: var(--secondary-background-color);
            margin-bottom: 10px; 
            box-shadow: 0 4px 6px rgba(0,0,0,0.1); 
            transition: all 0.3s ease;
        }
        div[data-testid="stExpander"] details:hover { border-color: var(--primary-color); box-shadow: 0 6px 12px rgba(0,0,0,0.15); }
        div[data-testid="stExpander"] summary { font-size: 16px !important; font-weight: 600 !important; color: var(--text-color) !important; padding: 10px; }
        
        .stButton>button { 
            border-radius: 8px; font-weight: bold; transition: all 0.3s; 
            border: 1px solid var(--primary-color); 
            background-color: var(--secondary-background-color) !important; 
            color: var(--text-color) !important;
        }
        .stButton>button:hover { 
            transform: translateY(-2px); 
            background-color: var(--primary-color) !important; 
            color: white !important; 
        }
        
        .footer-okta { 
            text-align: center; padding: 25px; margin-top: 50px; 
            color: var(--text-color); font-size: 15px; 
            border-top: 1px solid var(--primary-color); 
            background-color: var(--secondary-background-color); 
            border-radius: 10px; opacity: 0.8;
        }
        .footer-okta span { color: var(--primary-color); font-weight: 800; letter-spacing: 1px; font-size: 16px; }
        
        .login-box { 
            border: 2px solid var(--primary-color); 
            padding: 40px 30px; 
            border-radius: 16px; 
            background-color: var(--secondary-background-color);
            text-align: center; 
            height: 100%; 
            box-shadow: 0 10px 20px rgba(0,0,0,0.05);
            transition: all 0.3s ease;
        }
        .login-box:hover {
            transform: translateY(-5px);
            box-shadow: 0 15px 30px rgba(0,0,0,0.1);
        }
        .login-title { color: var(--text-color); font-size: 24px; font-weight: 700; margin-bottom: 15px; }
        .login-desc { color: var(--text-color); font-size: 15px; margin-bottom: 25px; line-height: 1.6; opacity: 0.9; }
    </style>
""", unsafe_allow_html=True)

# -------------------------------------------------------------------------
# SISTEM LOGIN & ROLE AKSES (VIEWER / ADMIN)
# -------------------------------------------------------------------------
if 'role' not in st.session_state:
    st.session_state['role'] = None

if st.session_state['role'] is None:
    st.markdown("<br><br>", unsafe_allow_html=True)
    render_header_logo() 
    
    st.markdown("<h1 style='text-align: center; font-size: 40px;'>Portal PM Dashboard</h1>", unsafe_allow_html=True)
    st.markdown("<p style='text-align: center; font-size: 18px; margin-bottom: 50px; letter-spacing: 1px; opacity: 0.8;'>Sistem Pelaporan Terpadu Preventive Maintenance Site Telekomunikasi</p>", unsafe_allow_html=True)
    
    col_v, col_space, col_a = st.columns([4, 1, 4])
    
    with col_v:
        st.markdown("""
            <div class='login-box'>
                <div style='font-size: 55px; margin-bottom: 15px;'>👁️</div>
                <div class='login-title'>Mode Viewer</div>
                <div class='login-desc'>Akses publik untuk memantau hasil laporan, melihat dokumentasi foto, dan mengunduh rekapitulasi data (PDF/Excel/PPTX) tanpa hak modifikasi.</div>
            </div>
        """, unsafe_allow_html=True)
        st.write("")
        if st.button("Masuk sebagai Viewer", type="secondary", use_container_width=True):
            st.session_state['role'] = 'Viewer'
            st.rerun()
            
    with col_a:
        st.markdown("""
            <div class='login-box'>
                <div style='font-size: 55px; margin-bottom: 15px;'>🔐</div>
                <div class='login-title'>Mode Admin</div>
                <div class='login-desc'>Akses khusus operasional untuk input form laporan baru, revisi data lapangan, penambahan dokumentasi, dan analisa grafik Power.</div>
            </div>
        """, unsafe_allow_html=True)
        admin_pass = st.text_input("Kata Sandi Admin:", type="password", placeholder="Masukkan Sandi...", key="pwd_login")
        if st.button("Login Admin", type="primary", use_container_width=True):
            if admin_pass == "KUT2027":
                st.session_state['role'] = 'Admin'
                st.rerun()
            else:
                st.error("❌ Kata sandi salah! Akses ditolak.")
                
    st.markdown("<br><br>", unsafe_allow_html=True)
    st.markdown("""
        <div style='text-align: center; font-size: 13px; opacity: 0.7;'>
            Secured System Application & Database Management<br>
            <b>Created By Okta Pradika © 2026</b>
        </div>
    """, unsafe_allow_html=True)
    st.stop()

# -------------------------------------------------------------------------
# 1. KONFIGURASI CLOUDINARY
# -------------------------------------------------------------------------
cloudinary.config( 
  cloud_name = "fxm61tjv", 
  api_key = "624877324969231", 
  api_secret = "LIFO6pfEg9fOM3nbsY8FBbVTpSI",
  secure = True
)

def upload_image(file_obj, folder_name="solar_bts_healthcheck"):
    if file_obj is not None:
        try:
            response = cloudinary.uploader.upload(file_obj.getvalue(), folder=folder_name)
            return response.get('secure_url')
        except Exception: return None
    return None

def upload_multiple_images(file_objs, folder_name="solar_bts_healthcheck"):
    urls = []
    if file_objs:
        for file in file_objs:
            url = upload_image(file, folder_name)
            if url: urls.append(url)
    return urls

def upload_datalog(file_obj, folder_name="prev_datalog"):
    if file_obj is not None:
        try:
            response = cloudinary.uploader.upload(file_obj.getvalue(), folder=folder_name, resource_type="raw", public_id=file_obj.name)
            return response.get('secure_url')
        except Exception as e: 
            st.error(f"Gagal upload datalog {file_obj.name}: {e}")
            return None
    return None

def tampilkan_grid_foto(url_data, caption=""):
    if not url_data: return
    urls = []
    if isinstance(url_data, str) and url_data.startswith("http"): urls = [url_data]
    elif isinstance(url_data, list): urls = [u for u in url_data if isinstance(u, str) and u.startswith("http")]
        
    if urls:
        if caption: st.markdown(f"*{caption}*")
        img_html = '<div style="display: grid; grid-template-columns: repeat(auto-fit, minmax(280px, 1fr)); gap: 12px; margin-bottom: 15px;">'
        for u in urls:
            img_html += f'<div style="background-color: var(--secondary-background-color); border: 1px solid var(--primary-color); border-radius: 8px; padding: 6px; text-align: center;"><a href="{u}" target="_blank"><img src="{u}" style="max-width: 100%; height: auto; max-height: 400px; object-fit: contain; border-radius: 6px;"></a></div>'
        img_html += '</div>'
        st.markdown(img_html, unsafe_allow_html=True)

# -------------------------------------------------------------------------
# 2. GENERATOR PDF & PPTX
# -------------------------------------------------------------------------
def clean_text(text):
    if not text: return "-"
    return str(text).encode('latin-1', 'ignore').decode('latin-1')

def build_pdf(r):
    pdf = FPDF()
    pdf.add_page()
    pdf.set_auto_page_break(auto=True, margin=15)
    
    pdf.set_font("helvetica", "B", 14)
    pdf.cell(0, 10, clean_text("BERITA ACARA PREVENTIVE MAINTENANCE"), ln=True, align="C")
    pdf.set_font("helvetica", "B", 12)
    pdf.cell(0, 6, clean_text("SITE TELEKOMUNIKASI"), ln=True, align="C")
    pdf.ln(5)
    
    pdf.set_font("helvetica", "", 10)
    intro = f"Berdasarkan hasil inspeksi dan pengerjaan lapangan pada tanggal {r.get('timestamp', '-')}, dengan ini diterangkan bahwa tim teknisi telah melaksanakan kegiatan Preventive Maintenance (Pemeliharaan Berkala) pada:"
    pdf.multi_cell(0, 6, clean_text(intro))
    pdf.ln(2)
    
    pdf.set_font("helvetica", "B", 10)
    pdf.cell(40, 6, clean_text("Nama Site / ID"), 0, 0)
    pdf.cell(0, 6, clean_text(f": {r.get('site_name', '-')}"), 0, 1)
    pdf.cell(40, 6, clean_text("Kategori Site"), 0, 0)
    pdf.cell(0, 6, clean_text(f": {r.get('kategori', 'SPS')}"), 0, 1)
    pdf.cell(40, 6, clean_text("Regional / NOP"), 0, 0)
    pdf.cell(0, 6, clean_text(f": {r.get('nop', '-')}"), 0, 1)
    pdf.cell(40, 6, clean_text("Pelaksana (Tim)"), 0, 0)
    pdf.cell(0, 6, clean_text(f": {r.get('teknisi', '-')}"), 0, 1)
    pdf.cell(40, 6, clean_text("Status Akhir"), 0, 0)
    pdf.cell(0, 6, clean_text(f": {r.get('status', '-')}"), 0, 1)
    pdf.ln(5)
    
    pdf.set_font("helvetica", "B", 11)
    pdf.cell(0, 8, clean_text("1. RINCIAN KEGIATAN & TINDAKAN (ACTION)"), ln=True)
    pdf.set_font("helvetica", "", 10)
    action_text = r.get('action', '-')
    if not action_text.strip(): action_text = "Tidak ada tindakan/catatan khusus yang dilaporkan oleh tim."
    pdf.multi_cell(0, 6, clean_text("Adapun rincian tindakan dan pekerjaan yang telah diselesaikan oleh tim di lokasi adalah sebagai berikut:\n" + action_text))
    pdf.ln(2)
    
    pdf.set_font("helvetica", "B", 10)
    pdf.cell(0, 6, clean_text("Penggunaan / Penggantian Sparepart:"), ln=True)
    pdf.set_font("helvetica", "", 10)
    pdf.multi_cell(0, 6, clean_text(r.get('sparepart', '-')))
    pdf.ln(5)
    
    pdf.set_font("helvetica", "B", 11)
    pdf.cell(0, 8, clean_text("2. HASIL PENGECEKAN PARAMETER SITE"), ln=True)
    pdf.set_font("helvetica", "", 10)
    
    pdf.cell(0, 6, clean_text(f"[v] Fisik & Lingkungan : {r.get('site_cond', '-')} | Fisik Tower: {r.get('tower_cond', '-')}"), ln=True)
    pdf.cell(0, 6, clean_text(f"[v] Modul Surya (SPS)  : Kondisi Shading {r.get('shading_status', '-')}"), ln=True)
    pdf.cell(0, 6, clean_text(f"[v] Jaringan Grid PLN  : {r.get('pln_status', '-')}"), ln=True)
    pdf.cell(0, 6, clean_text(f"[v] Rectifier & Beban  : {r.get('rect_brand', '-')} | Out: {r.get('rect_out_v', '-')}V | Load BTS: {r.get('total_load', '-')}A"), ln=True)
    pdf.cell(0, 6, clean_text(f"[v] Genset & Level BBM : {r.get('genset_status', '-')} | Ketersediaan BBM: {r.get('fuel_pct', '-')}%"), ln=True)
    pdf.cell(0, 6, clean_text(f"[v] Sistem Grounding   : Terukur {r.get('earth_ohm', '-')} Ohm"), ln=True)
    pdf.ln(5)
    
    pdf.multi_cell(0, 6, clean_text("Demikian Berita Acara ini dibuat sebenar-benarnya sesuai dengan kondisi aktual di lapangan untuk dapat dipergunakan sebagaimana mestinya."))
    pdf.ln(10)
    
    pdf.cell(90, 6, clean_text("Mengetahui / Menyetujui,"), 0, 0, "C")
    pdf.cell(90, 6, clean_text("Dibuat Oleh,"), 0, 1, "C")
    pdf.ln(20)
    pdf.set_font("helvetica", "B", 10)
    pdf.cell(90, 6, clean_text("(..........................................)"), 0, 0, "C")
    pdf.cell(90, 6, clean_text(f"( {r.get('teknisi', 'Tim Teknisi')} )"), 0, 1, "C")
    pdf.set_font("helvetica", "", 10)
    pdf.cell(90, 6, clean_text("Koordinator / PIC Area"), 0, 0, "C")
    pdf.cell(90, 6, clean_text("Pelaksana Lapangan"), 0, 1, "C")
    
    pdf.add_page()
    pdf.set_font("helvetica", "B", 12)
    pdf.cell(0, 10, clean_text("LAMPIRAN DOKUMENTASI FOTO (FULL)"), ln=True, align="C")
    pdf.ln(5)
    
    def draw_photo_grid(url_list, title):
        urls = [u for u in url_list if isinstance(u, str) and u.startswith("http")]
        if not urls: return
        pdf.set_font("helvetica", "B", 11)
        pdf.cell(0, 8, clean_text(title), ln=True)
        max_img_w = 160 
        for u in urls:
            # PENTING: Memaksa Cloudinary convert ke JPG agar format stabil & ukuran ideal
            opt_url = u
            if "upload/v" in opt_url:
                opt_url = opt_url.replace("upload/v", "upload/c_limit,w_800,f_jpg/v")

            try:
                response = requests.get(opt_url, timeout=12)
                if response.status_code == 200:
                    img = Image.open(BytesIO(response.content))
                    
                    # PASTIKAN Format RGB agar tidak crash di PDF jika gambar transparent/PNG
                    if img.mode in ('RGBA', 'P', 'LA'):
                        img = img.convert('RGB')
                        
                    # FIX TERPENTING: Simpan sebagai file fisik Temp. 
                    # FPDF versi lama HANYA BISA membaca file path (string), BUKAN object memori!
                    fd, temp_path = tempfile.mkstemp(suffix=".jpg")
                    os.close(fd)
                    img.save(temp_path, format="JPEG", quality=85)
                    
                    w_orig, h_orig = img.size
                    calc_h = (max_img_w / w_orig) * h_orig
                    img_w_adj = max_img_w
                    if calc_h > 240: 
                        calc_h = 240
                        img_w_adj = (calc_h / h_orig) * w_orig
                    
                    if pdf.get_y() + calc_h > 275: 
                        pdf.add_page()
                        
                    # Inject gambar dari file fisik sementara
                    pdf.image(temp_path, x=(210 - img_w_adj)/2, y=pdf.get_y(), w=img_w_adj)
                    pdf.set_y(pdf.get_y() + calc_h + 10)
                    
                    # Langsung hapus file dari server setelah dimasukkan ke PDF
                    if os.path.exists(temp_path):
                        os.remove(temp_path)
                else:
                    pdf.set_font("helvetica", "I", 10)
                    pdf.cell(0, 10, clean_text(f"[Gagal muat gambar (HTTP {response.status_code})]"), ln=True)
            except Exception as e:
                pdf.set_font("helvetica", "I", 10)
                pdf.cell(0, 10, clean_text(f"[Error koneksi/format sistem: {str(e)[:40]}]"), ln=True)
        pdf.ln(5)
        
    fisik_urls = (r.get('url_sites') or []) + (r.get('url_shadings') or []) + (r.get('extras_fisik') or [])
    draw_photo_grid(fisik_urls, "A. Dokumentasi Fisik & Lingkungan Site")
    
    panel_urls = []
    for p in r.get('panel_data', []):
        if isinstance(p.get('URL_Before'), list): panel_urls.extend(p.get('URL_Before'))
        elif p.get('URL_Before'): panel_urls.append(p.get('URL_Before'))
        panel_urls.extend(p.get('URLs_Before') or [])
        if isinstance(p.get('URL_After'), list): panel_urls.extend(p.get('URL_After'))
        elif p.get('URL_After'): panel_urls.append(p.get('URL_After'))
        panel_urls.extend(p.get('URLs_After') or [])
    panel_urls.extend(r.get('extras_panel') or [])
    draw_photo_grid(panel_urls, "B. Dokumentasi Modul Surya (SPS)")
    
    elek_urls = (r.get('url_rects') or []) + (r.get('url_gensets') or []) + (r.get('extras_elektrikal') or [])
    draw_photo_grid(elek_urls, "C. Dokumentasi Rectifier & Mesin Genset")
    
    bat_urls = []
    for b in r.get('battery_data', []): bat_urls.extend(b.get('URL_Fotos') or ( [b.get('URL_Foto')] if b.get('URL_Foto') else [] ))
    bat_urls.extend(r.get('url_grds') or [])
    bat_urls.extend(r.get('extras_baterai') or [])
    draw_photo_grid(bat_urls, "D. Dokumentasi Bank Baterai & Grounding")

    try: return bytes(pdf.output())
    except Exception:
        out = pdf.output(dest='S')
        return out.encode('latin-1', 'ignore') if isinstance(out, str) else out

# FUNGSI EXPORT PPTX
def build_pptx(db_list):
    prs = Presentation()
    
    # Title Slide
    slide_layout = prs.slide_layouts[0]
    slide = prs.slides.add_slide(slide_layout)
    title = slide.shapes.title
    subtitle = slide.placeholders[1]
    title.text = "Laporan Lengkap Preventive Maintenance"
    subtitle.text = f"Total Site Terinspeksi: {len(db_list)}\nGenerated on: {datetime.date.today()}"
    
    # Isi Slide Tiap Site menggunakan Tabel yang Teratur
    for r in db_list:
        slide_layout = prs.slide_layouts[5]
        slide = prs.slides.add_slide(slide_layout)
        
        title = slide.shapes.title
        title.text = f"Site: {r.get('site_name', '-')} | Status: {r.get('status', '-')}"
        
        rows = 6
        cols = 2
        left = Inches(0.5)
        top = Inches(1.5)
        width = Inches(4.5)
        height = Inches(3.0)
        
        table_shape = slide.shapes.add_table(rows, cols, left, top, width, height)
        table = table_shape.table
        
        table_data = [
            ("Tanggal", r.get('timestamp', '-')),
            ("Kategori / NOP", f"{r.get('kategori', 'SPS')} / {r.get('nop', '-')}"),
            ("Teknisi", r.get('teknisi', '-')),
            ("Tegangan / Load", f"{r.get('rect_out_v', '-')} V / {r.get('total_load', '-')} A"),
            ("Action Lapangan", r.get('action', '-')),
            ("Sparepart Diganti", r.get('sparepart', '-'))
        ]
        
        for row_idx, (k, v) in enumerate(table_data):
            table.cell(row_idx, 0).text = k
            table.cell(row_idx, 1).text = str(v)
            for cell in [table.cell(row_idx, 0), table.cell(row_idx, 1)]:
                for paragraph in cell.text_frame.paragraphs:
                    paragraph.font.size = Pt(13)
        
        site_urls = r.get('url_sites', [])
        if site_urls and len(site_urls) > 0:
            try:
                img_url = site_urls[0].replace("upload/v", "upload/c_limit,w_500,q_80,f_jpg/v")
                resp = requests.get(img_url, timeout=5)
                if resp.status_code == 200:
                    image_stream = BytesIO(resp.content)
                    slide.shapes.add_picture(image_stream, Inches(5.3), Inches(1.5), width=Inches(4.2))
            except:
                pass
                
    out = BytesIO()
    prs.save(out)
    return out.getvalue()

# -------------------------------------------------------------------------
# 3. SETUP DATABASE (SINKRONISASI KE GOOGLE SHEETS)
# -------------------------------------------------------------------------
SHEET_ID = "1HvgVicTWwO4RMQI6ZR3Mu3IgGicwjcLZl9mDN1auvJU"
SHEET_NAME = "Report Preventive"

def connect_gsheets():
    try:
        creds_json = st.secrets["gcp_json"]
        creds_dict = json.loads(creds_json)
        scopes = ["https://www.googleapis.com/auth/spreadsheets", "https://www.googleapis.com/auth/drive"]
        creds = Credentials.from_service_account_info(creds_dict, scopes=scopes)
        client = gspread.authorize(creds)
        return client.open_by_key(SHEET_ID).worksheet(SHEET_NAME)
    except Exception as e:
        st.error(f"⚠️ Koneksi Spreadsheet Gagal! Detail Error: {e}")
        return None

@st.cache_data(ttl=600) 
def fetch_data_from_gsheets():
    sheet = connect_gsheets()
    if sheet:
        data = sheet.get_all_values()
        if len(data) > 1: 
            db = []
            for row in data[1:]:
                if len(row) >= 6: 
                    try: db.append(json.loads(row[5]))
                    except: pass
            return db
    return []

if 'laporan_db' not in st.session_state:
    st.session_state['laporan_db'] = fetch_data_from_gsheets()

# -------------------------------------------------------------------------
# 4. NAVIGASI UTAMA
# -------------------------------------------------------------------------
st.sidebar.markdown("<h2 style='text-align: center; color: var(--primary-color);'>⚡ NAVIGASI</h2>", unsafe_allow_html=True)

if st.session_state['role'] == 'Admin':
    st.sidebar.markdown("<div style='text-align: center; background-color: var(--secondary-background-color); padding: 10px; border-radius: 8px; border: 1px solid var(--primary-color);'>Status: <b>🟢 ADMIN</b></div>", unsafe_allow_html=True)
    menu_options = ["📝 Form Preventive Check", "📊 Hasil Laporan & Dashboard", "📈 Monitoring Improvement"]
else:
    st.sidebar.markdown("<div style='text-align: center; background-color: var(--secondary-background-color); padding: 10px; border-radius: 8px; border: 1px solid gray;'>Status: <b>👁️ VIEWER</b></div>", unsafe_allow_html=True)
    menu_options = ["📊 Hasil Laporan & Dashboard", "📈 Monitoring Improvement"]

st.sidebar.write("")
menu = st.sidebar.radio("Pilih Operasional:", menu_options)
st.sidebar.markdown("---")

if st.sidebar.button("🚪 Keluar Akun (Log Out)", use_container_width=True):
    st.session_state['role'] = None
    st.rerun()

st.sidebar.markdown("---")
st.sidebar.info(f"📂 **Database Taut:**\n\n`Report Preventive`\n\n🔗 [Buka Spreadsheet Target](https://docs.google.com/spreadsheets/d/{SHEET_ID}/edit)")
st.sidebar.markdown("---")
st.sidebar.markdown("<div style='text-align: center; font-size: 13px; opacity: 0.7;'>System & Database Architecture<br><b style='color:var(--primary-color);'>Created By Okta Pradika</b></div>", unsafe_allow_html=True)

# =========================================================================
# MENU 1: FORM PENGECEKAN LENGKAP (HANYA ADMIN)
# =========================================================================
if menu == "📝 Form Preventive Check" and st.session_state['role'] == 'Admin':
    
    render_header_logo()
    st.markdown("<h1 style='text-align: center;'>⚡ Form Preventive Maintenance</h1>", unsafe_allow_html=True)
    st.info("💡 Data dan Lampiran (Foto & Datalog) Anda akan dienkripsi dan dikirim langsung ke Google Sheets & Cloudinary.")
    
    tabs = st.tabs(["📌 1. Info Site", "☀️ 2. SPS Panel", "🔌 3. PLN & Recti", "⛽ 4. Genset & BBM", "🔋 5. Baterai & Gnd", "📤 6. Upload & Submit"])

    with tabs[0]:
        st.markdown("#### Tentukan Kategori & Identitas Site")
        site_category = st.radio("Kategori Site (Wajib Pilih):", ["SPS", "Site Reguler"], horizontal=True)
        st.markdown("<hr style='margin:10px 0;'>", unsafe_allow_html=True)
        
        c1, c2 = st.columns(2)
        with c1:
            site_name = st.text_input("Nama / ID Site", placeholder="Contoh: BTS-PKY-001")
            nop_area = st.selectbox("NOP Area", ["Palangkaraya", "Pangkalan Bun", "Tarakan", "Pontianak", "Lainnya"])
            check_date = st.date_input("Tanggal Pengecekan", value=datetime.date.today())
        with c2:
            technician_name = st.text_input("Nama Pelaksana / Teknisi")
            weather = st.selectbox("Kondisi Cuaca", ["Cerah", "Berawan", "Hujan", "Kabut Asap menutup cahaya matahari", "Lembab tidak ada panas", "Badai"])
        st.divider()
        sc1, sc2 = st.columns(2)
        with sc1:
            site_condition = st.selectbox("Kondisi Halaman", ["Bersih & Aman", "Banyak Rumput", "Genangan Air"])
            tower_condition = st.selectbox("Fisik Tower", ["Aman", "Berkarat", "Rusak"])
        with sc2:
            site_photos = st.file_uploader("📸 Foto View Site (Bisa >1)", accept_multiple_files=True, key="spics")

    with tabs[1]:
        shading_status = st.selectbox("Status Shading?", ["Aman", "Sedikit Shading", "Kritis"])
        shading_photos = st.file_uploader("📸 Foto Shading (Bisa >1)", accept_multiple_files=True, key="shd")
        st.divider()
        
        konfigurasi_panel = st.radio("Metode Pengukuran SPS:", ["Individu (Satuan)", "Seri / String (Grup)"])
        panel_data = []

        if konfigurasi_panel == "Individu (Satuan)":
            num_panels = st.number_input("Jumlah Panel", min_value=1, value=24)
            for i in range(int(num_panels)):
                with st.expander(f"⚙️ Panel #{i+1}", expanded=(i==0)):
                    c1, c2, c3 = st.columns([1, 1, 1])
                    with c1:
                        voc = st.number_input(f"Voc [V]", value=21.5, key=f"v_{i}")
                        isc = st.number_input(f"Isc [A]", value=5.2, key=f"i_{i}")
                        p_cond = st.selectbox(f"Fisik", ["Baik", "Kotor", "Retak", "Delaminasi"], key=f"c_{i}")
                    with c2: pb = st.file_uploader(f"📸 KONDISI PANEL", key=f"pb_{i}")
                    with c3: pa = st.file_uploader(f"✨ PENGUKURAN", key=f"pa_{i}")
                    panel_data.append({"tipe": "Individu", "id": f"Panel #{i+1}", "voc": voc, "isc": isc, "kondisi": p_cond, "foto_before_obj": pb, "foto_after_obj": pa})
        else: 
            num_strings = st.number_input("Jumlah String", min_value=1, value=4)
            for i in range(int(num_strings)):
                with st.expander(f"🔌 String Seri #{i+1}", expanded=(i==0)):
                    c1, c2 = st.columns([1, 1.5])
                    with c1:
                        qty = st.number_input(f"Isi Panel per Seri", value=6, key=f"sq_{i}")
                        voc = st.number_input(f"Total Voc [V]", value=129.0, key=f"sv_{i}")
                        isc = st.number_input(f"Isc [A]", value=5.2, key=f"si_{i}")
                        p_cond = st.selectbox(f"Kondisi", ["Baik", "Kotor", "Retak"], key=f"sc_{i}")
                    with c2:
                        pb = st.file_uploader(f"📸 KONDISI PANEL (Bisa >1)", accept_multiple_files=True, key=f"spb_{i}")
                        pa = st.file_uploader(f"✨ PENGUKURAN (Bisa >1)", accept_multiple_files=True, key=f"spa_{i}")
                    panel_data.append({"tipe": "Seri", "id": f"String #{i+1}", "qty": qty, "voc": voc, "isc": isc, "kondisi": p_cond, "foto_before_objs": pb, "foto_after_objs": pa})

    with tabs[2]:
        c1, c2 = st.columns(2)
        with c1:
            pln_status = st.selectbox("Status PLN", ["Normal", "Padam", "Power Perusahaan", "Tidak Ada PLN"])
            jb_enclosure = st.selectbox("Box DC & Seal", ["Bersih", "Bocor", "Sarang Serangga"])
            jb_spd = st.selectbox("Arrester SPD", ["Normal", "Rusak"])
        with c2:
            rect_brand = st.text_input("Merek Rectifier", placeholder="Huawei / ZTE")
            rect_out_v = st.number_input("Tegangan Output Rectifier [V]", value=53.5)
            total_load_a = st.number_input("Total Beban BTS [A]", value=18.2)
            rect_alarm = st.selectbox("Alarm Rectifier", ["No Alarm", "Ada Alarm"])
            rect_photos = st.file_uploader("📸 Foto Rectifier (Bisa >1)", accept_multiple_files=True, key="rect_pics")

    with tabs[3]:
        g1, g2 = st.columns(2)
        with g1:
            genset_status = st.selectbox("Ketersediaan Genset", ["Ada (Beroperasi)", "Tidak Ada"])
            genset_brand = st.text_input("Merek Genset", placeholder="Perkins")
            hour_meter = st.number_input("Hour Meter", value=1250.5, step=0.1)
        with g2:
            tank_capacity = st.selectbox("Kapasitas Tangki [L]", [200, 300, 500, 1000])
            current_fuel_pct = st.slider("Level BBM [%]", 0, 100, 75)
            est_fuel = (current_fuel_pct / 100.0) * tank_capacity
            st.info(f"📌 **Volume BBM:** {est_fuel:.1f} Liter.")
            genset_photos = st.file_uploader("📸 Foto Genset & BBM (Bisa >1)", accept_multiple_files=True, key="gen_pics")

    with tabs[4]:
        b1, b2 = st.columns(2)
        with b1:
            num_bat = st.number_input("Jumlah Baterai", min_value=1, value=4)
            bat_data = []
            for j in range(int(num_bat)):
                with st.expander(f"🔋 Baterai #{j+1}", expanded=(j==0)):
                    bc1, bc2 = st.columns(2)
                    with bc1:
                        bv = st.number_input(f"Voltase [V]", value=12.2, key=f"bv_{j}")
                        bt = st.number_input(f"Load Charging [A]", value=28.0, key=f"bt_{j}")
                    with bc2:
                        bc = st.selectbox(f"Kondisi", ["Normal", "Bengkak", "Rusak tidak bisa charging", "Over Charging", "Korosi"], key=f"bc_{j}")
                        bp = st.file_uploader(f"Foto Bat #{j+1}", accept_multiple_files=True, key=f"bp_{j}")
                    bat_data.append({"id": f"Baterai #{j+1}", "voltase": bv, "suhu": bt, "kondisi": bc, "foto_objs": bp})
        with b2:
            earth_resistance = st.number_input("Tahanan Grounding [Ohm]", value=2.1)
            grd_cable = st.selectbox("Kabel Grounding", ["Kuat", "Kendor", "Putus"])
            grd_photos = st.file_uploader("📸 Foto Grounding (Bisa >1)", accept_multiple_files=True, key="grd")

    with tabs[5]:
        st.markdown("📂 **Upload Datalog (Bisa .csv, .xlsx, .zip, dll):**")
        datalog_files = st.file_uploader("Semua file tersimpan utuh", accept_multiple_files=True, key="datalog_all")
        st.divider()
        action_taken = st.text_area("🔧 Rincian Pekerjaan & Action di Lapangan:")
        sparepart_needed = st.text_input("📦 Penggantian Sparepart:")
        final_status = st.radio("Status Akhir Site:", ["Normal", "Minor Issue", "Major/Critical"])

        if st.button("🚀 UPLOAD & SINKRONKAN DATA", type="primary", use_container_width=True):
            if not site_name:
                st.error("⚠️ Mohon isi Nama Site!")
            else:
                with st.spinner("⏳ Sedang memproses dan mengamankan data ke Cloud..."):
                    url_sites = upload_multiple_images(site_photos, "prev_view")
                    url_shadings = upload_multiple_images(shading_photos, "prev_shade")
                    url_rects = upload_multiple_images(rect_photos, "prev_rect")
                    url_gensets = upload_multiple_images(genset_photos, "prev_genset")
                    url_grds = upload_multiple_images(grd_photos, "prev_grd")
                    
                    datalog_urls = []
                    if datalog_files:
                        for df in datalog_files:
                            du = upload_datalog(df, "prev_datalog")
                            if du: datalog_urls.append({"name": df.name, "url": du})

                    p_res = []
                    for p in panel_data:
                        if p.get("tipe") == "Individu":
                            p_res.append({"Tipe": "Individu", "Panel": p["id"], "Voc": p["voc"], "Isc": p["isc"], "Kondisi": p["kondisi"], "URL_Before": upload_image(p["foto_before_obj"]), "URL_After": upload_image(p["foto_after_obj"])})
                        else:
                            p_res.append({"Tipe": "Seri", "Panel": p["id"], "Qty": p["qty"], "Voc": p["voc"], "Isc": p["isc"], "Kondisi": p["kondisi"], "URLs_Before": upload_multiple_images(p["foto_before_objs"]), "URLs_After": upload_multiple_images(p["foto_after_objs"])})
                    
                    b_res = []
                    for b in bat_data: b_res.append({"Baterai": b["id"], "Voltase": b["voltase"], "Suhu": b["suhu"], "Kondisi": b["kondisi"], "URL_Fotos": upload_multiple_images(b["foto_objs"])})

                    # --- SIMPAN VARIABEL KATEGORI ---
                    report_dict = {
                        "timestamp": datetime.datetime.now().strftime("%Y-%m-%d %H:%M:%S"),
                        "site_name": site_name, "kategori": site_category, "nop": nop_area, "teknisi": technician_name,
                        "status": final_status, "action": action_taken, "sparepart": sparepart_needed,
                        "site_cond": site_condition, "tower_cond": tower_condition, "url_sites": url_sites,
                        "shading_status": shading_status, "url_shadings": url_shadings,
                        "panel_data": p_res,
                        "pln_status": pln_status, "rect_brand": rect_brand, "rect_out_v": rect_out_v, 
                        "total_load": total_load_a, "rect_alarm": rect_alarm, "url_rects": url_rects,
                        "genset_status": genset_status, "genset_brand": genset_brand, "hour_meter": hour_meter,
                        "tank_capacity": tank_capacity, "fuel_pct": current_fuel_pct, "fuel_liter": est_fuel,
                        "url_gensets": url_gensets,
                        "battery_data": b_res, "earth_ohm": earth_resistance, "url_grds": url_grds,
                        "datalog_files": datalog_urls,
                        "extras_fisik": [], "extras_panel": [], "extras_baterai": [], "extras_elektrikal": []
                    }
                    
                    sheet = connect_gsheets()
                    if sheet:
                        row_data = [
                            report_dict.get('timestamp', ''), report_dict.get('site_name', ''),
                            report_dict.get('nop', ''), report_dict.get('teknisi', ''),
                            report_dict.get('status', ''), json.dumps(report_dict)
                        ]
                        sheet.append_row(row_data)
                        st.cache_data.clear()
                        st.session_state['laporan_db'].append(report_dict)
                        st.success("✅ BERHASIL! Data telah diamankan di Google Sheets.")
                    else:
                        st.error("⚠️ Gagal menyambung ke Spreadsheet, mohon periksa Setting Secret API Anda.")

# =========================================================================
# MENU 2: HASIL LAPORAN (DITARIK DARI SPREADSHEET)
# =========================================================================
elif menu == "📊 Hasil Laporan & Dashboard":
    
    render_header_logo()
    st.markdown("<h1 style='text-align: center;'>📊 Dashboard Analytics & Report</h1>", unsafe_allow_html=True)
    
    db = st.session_state['laporan_db']

    if not db:
        st.warning("⚠ Belum ada data di Spreadsheet / Koneksi Sedang Proses.")
    else:
        total_sites = len(db)
        col_m1, col_m2, col_m3 = st.columns(3)
        with col_m1: st.metric("Total Site Ter-Inspeksi", total_sites)
        with col_m2: st.metric("Database Terhubung", "Google Sheets", "Online")
        with col_m3: st.metric("Cloud Storage", "Cloudinary", "Active")
        
        st.markdown("---")
        
        # ---------------------------------------------------------------------
        # FILTER KATEGORI SITE 
        # ---------------------------------------------------------------------
        st.markdown("#### 🗂️ Filter Kategori Site", unsafe_allow_html=True)
        filter_kat = st.radio("Pilih kategori laporan yang ingin ditampilkan:", ["Semua", "SPS", "Site Reguler"], horizontal=True)
        
        filtered_indices = []
        for idx, r in enumerate(db):
            kat = r.get('kategori', 'SPS')
            if filter_kat == "Semua" or kat == filter_kat:
                filtered_indices.append(idx)
                
        if len(filtered_indices) == 0:
            st.info(f"Tidak ada data Laporan untuk kategori: {filter_kat}")
            st.stop()

        # ---------------------------------------------------------------------
        # TABEL EXCEL FLAT & POWERPOINT
        # ---------------------------------------------------------------------
        st.markdown("### 📥 Ekspor Laporan Rekapitulasi (Excel & PPTX)")
        
        summary_list = []
        export_db = [db[idx] for idx in filtered_indices]
        
        for r in export_db:
            panel_issues = [p.get('Panel', 'Panel') for p in r.get('panel_data', []) if p.get('kondisi') and p.get('kondisi') != "Baik"]
            bat_issues = [b.get('Baterai', 'Baterai') for b in r.get('battery_data', []) if b.get('Kondisi') and b.get('Kondisi') != "Normal"]

            summary_list.append({
                "Timestamp": r.get('timestamp', '-'),
                "Nama Site": r.get('site_name', '-'),
                "Kategori": r.get('kategori', 'SPS'),
                "NOP": r.get('nop', '-'),
                "Teknisi": r.get('teknisi', '-'),
                "Status Akhir": r.get('status', '-'),
                "Kondisi Site": r.get('site_cond', '-'),
                "Fisik Tower": r.get('tower_cond', '-'),
                "Shading Panel": r.get('shading_status', '-'),
                "SPS Rusak (Jml)": len(panel_issues),
                "SPS Rusak (Detail)": ", ".join(panel_issues) if panel_issues else "Aman (Baik)",
                "Load Beban (A)": r.get('total_load', '-'),
                "PLN": r.get('pln_status', '-'),
                "Rectifier (Merek)": r.get('rect_brand', '-'),
                "Voltase Recti (V)": r.get('rect_out_v', '-'),
                "Genset": r.get('genset_status', '-'),
                "Level BBM (%)": r.get('fuel_pct', '-'),
                "Baterai Rusak (Jml)": len(bat_issues),
                "Baterai Rusak (Detail)": ", ".join(bat_issues) if bat_issues else "Aman (Normal)",
                "Grounding (Ohm)": r.get('earth_ohm', '-'),
                "Action / Pekerjaan": r.get('action', '-'),
                "Sparepart Diganti": r.get('sparepart', '-')
            })
            
        df_export = pd.DataFrame(summary_list)
        
        col_dl_ex, col_dl_ppt = st.columns(2)
        with col_dl_ex:
            try:
                excel_buffer = BytesIO()
                with pd.ExcelWriter(excel_buffer, engine='xlsxwriter') as writer:
                    df_export.to_excel(writer, index=False, sheet_name='Database PM')
                    workbook = writer.book
                    worksheet = writer.sheets['Database PM']
                    
                    header_format = workbook.add_format({'bold': True, 'font_color': 'white', 'bg_color': '#112240', 'border': 1, 'align': 'center', 'valign': 'vcenter'})
                    for col_num, value in enumerate(df_export.columns.values):
                        worksheet.write(0, col_num, value, header_format)

                    cell_center = workbook.add_format({'align': 'center', 'valign': 'vcenter', 'border': 1})
                    cell_left = workbook.add_format({'align': 'left', 'valign': 'vcenter', 'border': 1, 'text_wrap': True})

                    col_formats = [
                        (20, cell_center), (18, cell_center), (15, cell_center), (15, cell_center), (20, cell_center), (15, cell_center), 
                        (18, cell_center), (15, cell_center), (15, cell_center), (15, cell_center), (25, cell_left),   
                        (15, cell_center), (15, cell_center), (18, cell_center), (15, cell_center), (15, cell_center), (15, cell_center), 
                        (18, cell_center), (25, cell_left), (15, cell_center), (40, cell_left), (30, cell_left)        
                    ]
                    for i, (w, fmt) in enumerate(col_formats):
                        worksheet.set_column(i, i, w, fmt)

                    worksheet.freeze_panes(1, 0)
                    worksheet.autofilter(0, 0, len(df_export), len(df_export.columns) - 1)
                    
                file_data = excel_buffer.getvalue()
                st.download_button(label=f"📊 Download Excel ({filter_kat})", data=file_data, file_name=f"Database_PM_Master_{datetime.date.today()}.xlsx", mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet", type="primary", use_container_width=True)
            except Exception:
                csv_data = df_export.to_csv(index=False, sep=",").encode('utf-8')
                st.download_button(label="📥 Download Laporan Darurat (Raw CSV)", data=csv_data, file_name=f"Database_PM_{datetime.date.today()}.csv", mime="text/csv", type="secondary")

        with col_dl_ppt:
            if HAS_PPTX:
                with st.spinner("Menyiapkan PPTX..."):
                    pptx_data = build_pptx(export_db)
                    st.download_button(label=f"📽 Download Presentasi Report ({filter_kat})", data=pptx_data, file_name=f"Report_PM_KUT_{datetime.date.today()}.pptx", mime="application/vnd.openxmlformats-officedocument.presentationml.presentation", type="primary", use_container_width=True)
            else:
                st.error("Library `python-pptx` belum ter-install di server.")

        st.divider()

        # Daftar list data per-site
        for i in reversed(filtered_indices):
            r = db[i]
            site_id = r.get('site_name', 'Unknown')
            kategori_site = r.get('kategori', 'SPS')
            stat = r.get('status', '')
            icon = "🟢" if stat == "Normal" else "🟡" if stat == "Minor Issue" else "🔴"
            
            wa_text = f"""*BERITA ACARA PREVENTIVE MAINTENANCE* ⚡\n📍 *Site:* {site_id} ({kategori_site} - {r.get('nop', '-')})\n📅 *Tanggal:* {r.get('timestamp', '-')}\n👷 *Pelaksana:* {r.get('teknisi', '-')}\n📊 *Status:* {r.get('status', '-')}\n\n*RINCIAN TINDAKAN:*\n{r.get('action', '-')}\n\n*POWER & LOAD:*\n- PLN: {r.get('pln_status', '-')}\n- Rectifier: {r.get('rect_brand', '-')} ({r.get('rect_out_v', '-')}V)\n- Load BTS: {r.get('total_load', '-')} A\n\n*SPAREPART:*\n{r.get('sparepart', '-')}"""
            wa_url = f"https://wa.me/?text={urllib.parse.quote(wa_text)}"

            with st.expander(f"{icon}  |  {site_id}  |  {kategori_site}  |  {r.get('timestamp', '')}  |  Status: {stat}"):
                
                if st.checkbox("📄 Buat Berita Acara (PDF Resmi)", key=f"prep_pdf_{i}"):
                    with st.spinner("⏳ Rendering Dokumen PDF Resolusi Tinggi..."):
                        try:
                            pdf_bytes = build_pdf(r)
                            st.download_button(label="📥 Download PDF Berita Acara", data=pdf_bytes, file_name=f"Berita_Acara_{site_id}.pdf", mime="application/pdf", key=f"dl_pdf_{i}", type="primary")
                        except Exception as e:
                            st.error(f"Terjadi kesalahan saat menyusun PDF: {e}")
                
                st.markdown("<hr style='border: 1px solid var(--primary-color); margin: 15px 0;'>", unsafe_allow_html=True)
                c_btn1, c_btn2 = st.columns(2)
                with c_btn1: st.link_button("📱 Share Rangkuman ke WhatsApp", wa_url, use_container_width=True)
                with c_btn2: st.link_button("📈 Buka Database Spreadsheet Target", f"https://docs.google.com/spreadsheets/d/{SHEET_ID}/edit", use_container_width=True)
                
                st.markdown(f"**👨‍🔧 Pelaksana (Teknisi):** {r.get('teknisi', '-')} &nbsp;&nbsp;|&nbsp;&nbsp; **⚡ Beban Load:** {r.get('total_load', '-')} A")
                st.markdown(f"**🔧 Action Dikerjakan:** {r.get('action', '-')}")
                if r.get('sparepart'): st.warning(f"**📦 Sparepart Diganti:** {r['sparepart']}")
                
                st.markdown("<br>", unsafe_allow_html=True)
                
                ltab1, ltab2, ltab3, ltab4, ltab5, ltab6 = st.tabs(["🏗️ 1. Fisik", "☀️ 2. Panel SPS", "🔌 3. Recti & Genset", "🔋 4. Baterai & Gnd", "📂 5. Datalog", "📈 6. Analisa Power BBU"])
                
                with ltab1:
                    st.write(f"- Kondisi Site: {r.get('site_cond', '-')} | Tower: {r.get('tower_cond', '-')} | Shading: {r.get('shading_status', '-')}")
                    tampilkan_grid_foto(r.get('url_sites'), "📸 View Site")
                    tampilkan_grid_foto(r.get('url_shadings'), "📸 Shading")
                    tampilkan_grid_foto(r.get('extras_fisik'), "📸 Tambahan")

                with ltab2:
                    for p in r.get('panel_data', []):
                        st.markdown(f"**{p.get('Panel', '-')}** | Voc: {p.get('Voc','-')}V | Isc: {p.get('Isc','-')}A")
                        cb, ca = st.columns(2)
                        with cb: tampilkan_grid_foto(p.get('URL_Before') or p.get('URLs_Before'), "📸 KONDISI PANEL")
                        with ca: tampilkan_grid_foto(p.get('URL_After') or p.get('URLs_After'), "✨ PENGUKURAN")
                        st.divider()
                    tampilkan_grid_foto(r.get('extras_panel'), "📸 Tambahan Panel")

                with ltab3:
                    c_p1, c_p2 = st.columns(2)
                    with c_p1:
                        st.write(f"**PLN:** {r.get('pln_status', '-')} | **Rectifier:** {r.get('rect_brand', '-')} ({r.get('rect_out_v', '-')}V)")
                        tampilkan_grid_foto(r.get('url_rects'), "📸 Rectifier")
                    with c_p2:
                        st.write(f"**Genset:** {r.get('genset_status', '-')} | **BBM:** {r.get('fuel_pct', '-')}%")
                        tampilkan_grid_foto(r.get('url_gensets'), "📸 Genset")
                    st.divider()
                    tampilkan_grid_foto(r.get('extras_elektrikal'), "📸 Tambahan Mesin/Elektrikal")

                with ltab4:
                    st.markdown("### 🔋 Bank Baterai")
                    for b in r.get('battery_data', []):
                        with st.container(border=True):
                            bc1, bc2 = st.columns([2, 1])
                            with bc1:
                                st.markdown(f"**{b.get('Baterai', 'Baterai')}**")
                                st.write(f"🔹 **Voltase:** {b.get('Voltase', '-') } V  |  🌡 **Load Charging:** {b.get('Suhu', '-') } A")
                                st.write(f"🔍 **Kondisi Fisik:** {b.get('Kondisi', '-')}")
                            bat_urls = b.get('URL_Fotos') or b.get('URL_Foto')
                            if bat_urls:
                                st.markdown("")
                                tampilkan_grid_foto(bat_urls, "📸 Dokumentasi Baterai")
                        st.markdown("")
                    st.markdown("---")
                    st.markdown("### 🌍 Sistem Grounding")
                    st.write(f"⚡ **Tahanan Grounding:** {r.get('earth_ohm', '-')} Ohm")
                    tampilkan_grid_foto(r.get('url_grds'), "📸 Foto Grounding")
                    tampilkan_grid_foto(r.get('extras_baterai'), "📸 Tambahan Baterai & Grounding")

                with ltab5:
                    st.markdown("📂 **Datalog Universal yang Tersimpan:**")
                    dls = r.get('datalog_files', [])
                    if dls:
                        for dl in dls: st.markdown(f"- 🔗 [{dl['name']}]({dl['url']})")
                    else: st.info("Tidak ada file datalog diunggah pada site ini.")

                with ltab6:
                    st.markdown("### ⚡ Grafik Analisa Power & Voltage")
                    st.info(f"Sistem sedang melacak file data histori **`{site_id}.xlsx`** di repositori...")
                    
                    excel_filename = f"{site_id}.xlsx"
                    if os.path.exists(excel_filename):
                        try:
                            df_power = pd.read_excel(excel_filename, engine='openpyxl')
                            if 'Begin Time' in df_power.columns and 'MinVoltageOfBBU(V)' in df_power.columns:
                                df_power['Waktu'] = pd.to_datetime(df_power['Begin Time']).dt.strftime('%H:%M')
                                df_power.set_index('Waktu', inplace=True)
                                
                                chart_cols = ['MaxVoltageOfBBU(V)', 'MinVoltageOfBBU(V)', 'AvgVoltageOfBBU(V)']
                                chart_data = df_power[chart_cols].copy()
                                chart_data['Batas Hold (55V)'] = 55.0
                                
                                st.line_chart(chart_data, color=["#64FFDA", "#FF5252", "#FFC107", "#FFFFFF"])
                                
                                min_voltage = df_power['MinVoltageOfBBU(V)'].min()
                                if min_voltage < 55.0:
                                    st.error(f"🚨 **ANALISA DROP VOLTAGE:** Ditemukan tegangan Drop di bawah batas 55V! Tegangan terendah terekam di angka **{min_voltage} V**.")
                                else:
                                    st.success(f"✅ **ANALISA STABIL:** Kondisi Power aman. Tegangan berhasil di-hold (tidak jatuh di bawah batas 55V). Tegangan terendah terekam: {min_voltage} V.")
                            else:
                                st.warning(f"File '{excel_filename}' ditemukan, tapi struktur kolomnya tidak sesuai format U2000/U2020. Pastikan ada kolom 'Begin Time' dan 'MinVoltageOfBBU(V)'.")
                        except Exception as e:
                            st.error(f"Gagal membaca file {excel_filename}. Pesan Error: {e}")
                    else:
                        st.warning(f"File log power **`{excel_filename}`** belum di-upload ke repositori/folder lokal server. Silakan upload file excel dengan nama site tersebut untuk melihat grafik otomatis.")


                st.markdown("<br>", unsafe_allow_html=True)
                
                # =========================================================
                # BLOK EDIT DATA (KHUSUS ADMIN)
                # =========================================================
                if st.session_state['role'] == 'Admin':
                    with st.container(border=True):
                        st.markdown("<h4 style='color: var(--primary-color);'>🛠 REVISI DATA & TAMBAH LAMPIRAN</h4>", unsafe_allow_html=True)
                        c_edit1, c_edit2 = st.columns(2)
                        with c_edit1:
                            new_site = st.text_input("Edit Nama Site / ID", r.get('site_name',''), key=f"esite_{i}")
                            kat_options = ["SPS", "Site Reguler"]
                            curr_kat = r.get('kategori', 'SPS')
                            if curr_kat not in kat_options: curr_kat = "SPS"
                            new_kat = st.selectbox("Edit Kategori Site", kat_options, index=kat_options.index(curr_kat), key=f"ekat_{i}")
                            
                            nop_options = ["Palangkaraya", "Pangkalan Bun", "Tarakan", "Pontianak", "Lainnya"]
                            curr_nop = r.get('nop', 'Palangkaraya')
                            if curr_nop not in nop_options: curr_nop = "Lainnya"
                            new_nop = st.selectbox("Edit NOP Area", nop_options, index=nop_options.index(curr_nop), key=f"enop_{i}")
                            new_tek = st.text_input("Edit Teknisi", r.get('teknisi',''), key=f"et_{i}")
                        with c_edit2:
                            status_options = ["Normal", "Minor Issue", "Major/Critical"]
                            curr_stat = r.get('status', 'Normal')
                            if curr_stat not in status_options: curr_stat = "Normal"
                            new_status = st.selectbox("Edit Status Akhir", status_options, index=status_options.index(curr_stat), key=f"estat_{i}")
                            new_sp = st.text_input("Edit Sparepart", r.get('sparepart',''), key=f"es_{i}")
                        new_act = st.text_area("Edit Action / Tindakan", r.get('action',''), key=f"ea_{i}")
                        st.markdown("**Tambah Lampiran Susulan (Otomatis Masuk Cloud):**")
                        up_f = st.file_uploader("📸 Fisik / Shading", accept_multiple_files=True, key=f"uf_{i}")
                        up_p = st.file_uploader("📸 Panel Surya", accept_multiple_files=True, key=f"up_{i}")
                        up_b = st.file_uploader("📸 Baterai / Grounding", accept_multiple_files=True, key=f"ub_{i}")
                        up_e = st.file_uploader("📸 Elektrikal / Mesin", accept_multiple_files=True, key=f"ue_{i}")
                        up_dl = st.file_uploader("📂 Datalog (Zip, Csv, xlsx)", accept_multiple_files=True, key=f"udl_{i}")

                        if st.button("💾 Simpan Perubahan ke Server", key=f"btn_{i}"):
                            with st.spinner("Mengirim Revisi ke Database Utama..."):
                                r['site_name'], r['kategori'], r['nop'], r['status'], r['teknisi'], r['action'], r['sparepart'] = new_site, new_kat, new_nop, new_status, new_tek, new_act, new_sp
                                if up_f: r['extras_fisik'] = r.get('extras_fisik', []) + upload_multiple_images(up_f)
                                if up_p: r['extras_panel'] = r.get('extras_panel', []) + upload_multiple_images(up_p)
                                if up_b: r['extras_baterai'] = r.get('extras_baterai', []) + upload_multiple_images(up_b)
                                if up_e: r['extras_elektrikal'] = r.get('extras_elektrikal', []) + upload_multiple_images(up_e)
                                if up_dl:
                                    new_datalog = []
                                    for df in up_dl:
                                        du = upload_datalog(df, "prev_datalog")
                                        if du: new_datalog.append({"name": df.name, "url": du})
                                    r['datalog_files'] = r.get('datalog_files', []) + new_datalog
                                
                                sheet = connect_gsheets()
                                if sheet:
                                    row_num = i + 2 
                                    sheet.update_cell(row_num, 2, new_site)
                                    sheet.update_cell(row_num, 3, new_nop)
                                    sheet.update_cell(row_num, 4, new_tek)
                                    sheet.update_cell(row_num, 5, new_status)
                                    sheet.update_cell(row_num, 6, json.dumps(r))
                                    st.cache_data.clear() 
                            st.success("✅ REVISI BERHASIL! Data & Foto tersimpan permanen.")
                            st.rerun()

# =========================================================================
# MENU 3: EXCEL LIVE-EDITOR (MONITORING IMPROVEMENT)
# =========================================================================
elif menu == "📈 Monitoring Improvement":
    render_header_logo()
    st.markdown("<h1 style='text-align: center;'>📈 Master Tracker Improvement</h1>", unsafe_allow_html=True)
    
    file_master = "Monitoring_Availability_Improvement_Visit_SPS_NOP_PLK.xlsx"
    
    if not os.path.exists(file_master):
        st.error(f"❌ File '{file_master}' tidak ditemukan di sistem/server. Pastikan Anda telah meletakkan file tersebut satu folder dengan aplikasi.")
    else:
        tab_dashboard, tab_tim, tab_editor = st.tabs(["📊 Kurva S", "👥 Progress Tim (Done/Berjalan)", "📝 Live Editor Data"])
        
        # --- TAB 1: VISUALISASI DASHBOARD KURVA S ---
        with tab_dashboard:
            st.markdown("### 📈 S-Curve Target vs Aktual")
            st.info("💡 Data grafik ini dibaca secara Real-Time dari **Sheet 'Kurva S'** di file Excel Master Anda.")
            
            try:
                df_kurva = pd.read_excel(file_master, sheet_name='Kurva S', header=None)
                total_site = df_kurva.iloc[6, 13] if pd.notna(df_kurva.iloc[6, 13]) else 0
                minggu_berj = df_kurva.iloc[7, 13] if pd.notna(df_kurva.iloc[7, 13]) else 0
                site_plan = df_kurva.iloc[8, 13] if pd.notna(df_kurva.iloc[8, 13]) else 0
                
                m1, m2, m3 = st.columns(3)
                m1.metric(label="Total Site Target", value=f"{int(total_site)} Site")
                m2.metric(label="Site dengan Plan", value=f"{int(site_plan)} Site")
                m3.metric(label="Minggu Berjalan", value=f"Week {int(minggu_berj)}")
                
                st.markdown("---")
                
                df_chart = df_kurva.iloc[6:58, :11].copy()
                df_chart.columns = ['Week', 'Mulai', 'Label', 'Plan_Minggu', 'Plan_Kumulatif', 'Actual_Minggu', 'Actual_Kumulatif', 'Pct_Plan', 'Pct_Actual', 'Gap', 'Achievement']
                df_chart = df_chart.dropna(subset=['Week', 'Plan_Kumulatif'])
                
                df_chart['Week'] = "W" + df_chart['Week'].astype(str)
                df_chart.set_index('Week', inplace=True)
                
                col_chart, col_table = st.columns([2, 1])
                with col_chart:
                    st.markdown("**Perbandingan Kumulatif Plan vs Actual**")
                    st.line_chart(df_chart[['Plan_Kumulatif', 'Actual_Kumulatif']], color=["#FF5252", "#64FFDA"])
                    
                with col_table:
                    st.markdown("**Tabel Gap Mingguan**")
                    st.dataframe(df_chart[['Plan_Kumulatif', 'Actual_Kumulatif', 'Gap']], use_container_width=True)
            except Exception as e:
                st.error(f"Gagal memvisualisasikan Grafik S-Curve: {e}")

        # --- TAB 2: PROGRESS TIM (DONE VS MASIH BERJALAN) ---
        with tab_tim:
            st.markdown("### 👥 Dashboard Eksekusi Lapangan")
            st.info("💡 Sistem akan membaca Sheet **Tracker Improvement** untuk menghitung berapa site yang sudah Selesai (Done) dan masih dalam pengerjaan (On Progress).")
            
            try:
                # Membaca Sheet Tracker
                df_track = pd.read_excel(file_master, sheet_name='Tracker Improvement', header=None)
                headers_track = df_track.iloc[3].fillna("").astype(str).tolist()
                df_t = df_track.iloc[4:].copy()
                df_t.columns = headers_track
                
                # Coba cari otomatis kolom yang mengandung kata status / progress / keterangan
                status_cols = [c for c in headers_track if 'status' in c.lower() or 'progress' in c.lower() or 'keterangan' in c.lower() or 'aktual' in c.lower()]
                default_idx = headers_track.index(status_cols[0]) if status_cols else 0
                
                st.markdown("**Pilih kolom pada Excel yang berisi status pengerjaan tim:**")
                sel_col = st.selectbox("Pilih Kolom Status:", headers_track, index=default_idx)
                
                if sel_col:
                    # Ambil data kolom tersebut, rapikan teksnya (kapital & buang spasi ujung)
                    df_t[sel_col] = df_t[sel_col].astype(str).str.strip().str.upper()
                    # Buang cell yang kosong (NaN)
                    df_valid = df_t[~df_t[sel_col].isin(['NAN', 'NAT', 'NONE', ''])]
                    
                    # Hitung kemunculan masing-masing status
                    counts = df_valid[sel_col].value_counts().reset_index()
                    counts.columns = ['Status', 'Jumlah']
                    
                    if not counts.empty:
                        # Keyword cerdas untuk mendeteksi mana yang Selesai dan Berjalan
                        done_kws = ['DONE', 'SELESAI', 'CLOSE', 'OK', 'COMPLETED']
                        prog_kws = ['PROGRESS', 'BERJALAN', 'ON GOING', 'OPEN', 'PENDING', 'ON PROGRESS']
                        
                        # Menjumlahkan angka berdasarkan kecocokan keyword
                        jml_done = counts[counts['Status'].apply(lambda x: any(k in x for k in done_kws))]['Jumlah'].sum()
                        jml_prog = counts[counts['Status'].apply(lambda x: any(k in x for k in prog_kws))]['Jumlah'].sum()
                        jml_lain = counts['Jumlah'].sum() - (jml_done + jml_prog)
                        
                        st.markdown("#### 🎯 Hasil Kalkulasi Progress")
                        c_t1, c_t2, c_t3 = st.columns(3)
                        c_t1.metric("✅ TOTAL SELESAI (Done)", f"{jml_done} Site")
                        c_t2.metric("⏳ SEDANG BERJALAN (On Progress)", f"{jml_prog} Site")
                        c_t3.metric("📝 Lainnya / Belum Mulai", f"{jml_lain} Site")
                        
                        st.markdown("#### 📊 Grafik Breakdown Status")
                        st.bar_chart(counts.set_index('Status'), color="#64FFDA")
                        
                        with st.expander("Tampilkan Rekap Tabel Mentah Status", expanded=False):
                            st.dataframe(counts, use_container_width=True)
                    else:
                        st.warning(f"Data pada kolom '{sel_col}' kosong atau seluruhnya terbaca sebagai cell kosong.")
            except Exception as e:
                st.error(f"Gagal memproses data tracker tim: {e}")

        # --- TAB 3: EDITOR EXCEL LANGSUNG & VIEWER DASHBOARD ---
        with tab_editor:
            st.markdown("### 📝 Tabel Master Tracker & Viewer")
            st.write("Klik ganda (*double click*) pada sel tabel untuk merubah isi data secara instan.")
            
            sheet_options = ["Tracker Improvement", "Jadwal Visit SPS", "Dashboard", "Kurva S"]
            sheet_choice = st.selectbox("Pilih Sheet Excel yang Ingin Diedit / Dilihat:", sheet_options)
            
            try:
                df_raw = pd.read_excel(file_master, sheet_name=sheet_choice, header=None)
                
                if sheet_choice == "Dashboard":
                    st.info("💡 Karena sheet Dashboard bawaan Excel biasanya berisi desain sel yang di-merge dan banyak bagan, tampilannya di web mungkin terlihat sebagai sekumpulan teks kasar. Gunakan **Tab 👥 Progress Tim** di atas untuk tampilan Dashboard yang jauh lebih rapi.")
                    st.dataframe(df_raw, use_container_width=True, height=500)
                
                else:
                    headers = df_raw.iloc[3].fillna("").astype(str).tolist()
                    df_data = df_raw.iloc[4:].copy()
                    df_data.columns = headers
                    df_data = df_data.reset_index(drop=True)
                    
                    edited_df = st.data_editor(
                        df_data,
                        use_container_width=True,
                        num_rows="dynamic",
                        key=f"editor_{sheet_choice}"
                    )
                    
                    if st.button("💾 Simpan Perubahan & Download Excel Utuh", type="primary"):
                        state_key = f"editor_{sheet_choice}"
                        changes = st.session_state[state_key]
                        
                        if changes.get("edited_rows") or changes.get("added_rows"):
                            with st.spinner("Menyuntikkan data baru ke Master Excel (Menjaga Format)..."):
                                import openpyxl
                                wb = openpyxl.load_workbook(file_master)
                                ws = wb[sheet_choice]
                                
                                for row_idx_str, col_changes in changes.get("edited_rows", {}).items():
                                    row_idx = int(row_idx_str)
                                    excel_row = row_idx + 5 
                                    
                                    for col_name, new_val in col_changes.items():
                                        if col_name in headers:
                                            col_idx = headers.index(col_name) + 1
                                            ws.cell(row=excel_row, column=col_idx).value = new_val
                                            
                                for added_row in changes.get("added_rows", []):
                                    excel_row = ws.max_row + 1
                                    for col_name, new_val in added_row.items():
                                        if col_name in headers:
                                            col_idx = headers.index(col_name) + 1
                                            ws.cell(row=excel_row, column=col_idx).value = new_val
                                
                                out_buffer = BytesIO()
                                wb.save(out_buffer)
                                
                                st.success("✅ Berhasil! Perubahan telah disuntikkan ke file asli.")
                                st.download_button(
                                    label=f"📥 Download File Master Terupdate",
                                    data=out_buffer.getvalue(),
                                    file_name=f"Updated_{file_master}",
                                    mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                                    type="primary"
                                )
                        else:
                            st.warning("Belum ada sel yang diubah. Klik ganda pada tabel untuk mengubah isinya.")
                            
            except Exception as e:
                st.error(f"Terjadi kendala saat membaca sheet {sheet_choice}: {e}")

# -------------------------------------------------------------------------
# FOOTER HAK CIPTA OKTA PRADIKA
# -------------------------------------------------------------------------
st.markdown("""
    <div class="footer-okta">
        🚀 System Application & Database Management<br>
        <span>Create Data By Okta Pradika</span> © 2026
    </div>
""", unsafe_allow_html=True)
