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
import tempfile 

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
# CUSTOM CSS
# -------------------------------------------------------------------------
st.markdown("""
    <style>
        h1, h2, h3 { color: var(--primary-color) !important; font-family: 'Segoe UI', sans-serif; }
        div[data-testid="stExpander"] details {
            border: 1px solid var(--primary-color); border-radius: 10px; background-color: var(--secondary-background-color); margin-bottom: 10px; box-shadow: 0 4px 6px rgba(0,0,0,0.1); transition: all 0.3s ease;
        }
        div[data-testid="stExpander"] details:hover { border-color: var(--primary-color); box-shadow: 0 6px 12px rgba(0,0,0,0.15); }
        div[data-testid="stExpander"] summary { font-size: 16px !important; font-weight: 600 !important; color: var(--text-color) !important; padding: 10px; }
        .stButton>button { border-radius: 8px; font-weight: bold; transition: all 0.3s; border: 1px solid var(--primary-color); background-color: var(--secondary-background-color) !important; color: var(--text-color) !important; }
        .stButton>button:hover { transform: translateY(-2px); background-color: var(--primary-color) !important; color: white !important; }
        .footer-okta { text-align: center; padding: 25px; margin-top: 50px; color: var(--text-color); font-size: 15px; border-top: 1px solid var(--primary-color); background-color: var(--secondary-background-color); border-radius: 10px; opacity: 0.8; }
        .footer-okta span { color: var(--primary-color); font-weight: 800; letter-spacing: 1px; font-size: 16px; }
        .login-box { border: 2px solid var(--primary-color); padding: 40px 30px; border-radius: 16px; background-color: var(--secondary-background-color); text-align: center; height: 100%; box-shadow: 0 10px 20px rgba(0,0,0,0.05); transition: all 0.3s ease; }
        .login-box:hover { transform: translateY(-5px); box-shadow: 0 15px 30px rgba(0,0,0,0.1); }
        .login-title { color: var(--text-color); font-size: 24px; font-weight: 700; margin-bottom: 15px; }
        .login-desc { color: var(--text-color); font-size: 15px; margin-bottom: 25px; line-height: 1.6; opacity: 0.9; }
    </style>
""", unsafe_allow_html=True)

# -------------------------------------------------------------------------
# SISTEM LOGIN & ROLE AKSES 
# -------------------------------------------------------------------------
if 'role' not in st.session_state: st.session_state['role'] = None

if st.session_state['role'] is None:
    st.markdown("<br><br>", unsafe_allow_html=True)
    render_header_logo() 
    st.markdown("<h1 style='text-align: center; font-size: 40px;'>Portal PM Dashboard</h1>", unsafe_allow_html=True)
    st.markdown("<p style='text-align: center; font-size: 18px; margin-bottom: 50px; letter-spacing: 1px; opacity: 0.8;'>Sistem Pelaporan Terpadu Preventive Maintenance Site Telekomunikasi</p>", unsafe_allow_html=True)
    col_v, col_space, col_a = st.columns([4, 1, 4])
    
    with col_v:
        st.markdown("""
            <div class='login-box'>
                <div style='font-size: 55px; margin-bottom: 15px;'>👷‍♂️</div>
                <div class='login-title'>Mode Tim Lapangan</div>
                <div class='login-desc'>Akses publik langsung untuk tim lapangan. Bisa mengisi Form Preventive Check (SiUPDATE) dan memantau dashboard tanpa hak akses modifikasi data.</div>
            </div>
        """, unsafe_allow_html=True)
        st.write("")
        if st.button("Masuk sebagai Tim Lapangan", type="secondary", use_container_width=True):
            st.session_state['role'] = 'Lapangan'
            st.rerun()
            
    with col_a:
        st.markdown("""
            <div class='login-box'>
                <div style='font-size: 55px; margin-bottom: 15px;'>🔐</div>
                <div class='login-title'>Mode Admin</div>
                <div class='login-desc'>Akses khusus operasional untuk revisi data lapangan, perbaikan dashboard, penambahan dokumentasi, dan analisa grafik Power.</div>
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
    st.markdown("<div style='text-align: center; font-size: 13px; opacity: 0.7;'>Secured System Application & Database Management<br><b>Created By Okta Pradika © 2026</b></div>", unsafe_allow_html=True)
    st.stop()

# -------------------------------------------------------------------------
# KONFIGURASI CLOUDINARY
# -------------------------------------------------------------------------
cloudinary.config( cloud_name = "fxm61tjv", api_key = "624877324969231", api_secret = "LIFO6pfEg9fOM3nbsY8FBbVTpSI", secure = True )

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
        except Exception as e: return None
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
# TANDA TANGAN & PDF/PPTX GENERATOR (Bawaan Asli Palangkaraya)
# -------------------------------------------------------------------------
TTD_OKTA_B64 = "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"

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
    
    # Deteksi apakah ini format Pangkalanbun atau Palangkaraya di PDF
    if r.get('kategori') == 'Pangkalan Bun':
        pdf.set_font("helvetica", "", 10)
        pdf.multi_cell(0, 6, clean_text(f"SITE VISIT REPORT: {r.get('site_name', '-')} (Pangkalan Bun)"))
        pdf.cell(0, 6, clean_text(f"TE Name: {r.get('teknisi', '-')} | Ticket#: {r.get('pb_ticket', '-')}"), ln=True)
        pdf.cell(0, 6, clean_text(f"Catuan Power: {r.get('pb_power', '-')}"), ln=True)
        pdf.ln(5)
        pdf.multi_cell(0, 6, clean_text(f"Temuan / Catatan:\n{r.get('action', '-')}"))
    else:
        # Format Asli Palangkaraya
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

    pdf.set_font("helvetica", "", 10)
    pdf.cell(90, 6, "", 0, 0) 
    pdf.cell(90, 6, clean_text("Data dibuat oleh,"), 0, 1, "C")
    try:
        fd, ttd_path = tempfile.mkstemp(suffix=".png")
        os.close(fd)
        with open(ttd_path, "wb") as f:
            f.write(base64.b64decode(TTD_OKTA_B64))
        pdf.image(ttd_path, x=135, y=pdf.get_y() + 2, w=30)
        os.remove(ttd_path)
    except Exception: pass
        
    pdf.ln(22)
    pdf.set_font("helvetica", "B", 10)
    pdf.cell(90, 6, "", 0, 0) 
    pdf.cell(90, 6, clean_text("(          Okta Pradika          )"), 0, 1, "C")
    pdf.set_font("helvetica", "", 10)
    pdf.cell(90, 6, "", 0, 0) 
    pdf.cell(90, 6, clean_text("Koordinator KUT Palangkaraya"), 0, 1, "C")
    
    # ------------------ LAMPIRAN FOTO ------------------
    pdf.add_page()
    pdf.set_font("helvetica", "B", 12)
    pdf.cell(0, 10, clean_text("LAMPIRAN DOKUMENTASI FOTO"), ln=True, align="C")
    pdf.ln(5)
    
    def draw_photo_grid(url_list, title):
        urls = [u for u in url_list if isinstance(u, str) and u.startswith("http")]
        if not urls: return
        pdf.set_font("helvetica", "B", 11)
        pdf.cell(0, 8, clean_text(title), ln=True)
        max_img_w = 160 
        for u in urls:
            opt_url = u
            if "upload/v" in opt_url: opt_url = opt_url.replace("upload/v", "upload/c_limit,w_800,f_jpg/v")
            try:
                response = requests.get(opt_url, timeout=12)
                if response.status_code == 200:
                    img = Image.open(BytesIO(response.content))
                    if img.mode in ('RGBA', 'P', 'LA'): img = img.convert('RGB')
                    fd, temp_path = tempfile.mkstemp(suffix=".jpg")
                    os.close(fd)
                    img.save(temp_path, format="JPEG", quality=85)
                    w_orig, h_orig = img.size
                    calc_h = (max_img_w / w_orig) * h_orig
                    img_w_adj = max_img_w
                    if calc_h > 240: 
                        calc_h = 240
                        img_w_adj = (calc_h / h_orig) * w_orig
                    if pdf.get_y() + calc_h > 275: pdf.add_page()
                    pdf.image(temp_path, x=(210 - img_w_adj)/2, y=pdf.get_y(), w=img_w_adj)
                    pdf.set_y(pdf.get_y() + calc_h + 10)
                    if os.path.exists(temp_path): os.remove(temp_path)
            except Exception: pass
        pdf.ln(5)
        
    if r.get('kategori') == 'Pangkalan Bun':
        draw_photo_grid(r.get('pb_fotos_info', []), "A. Informasi Site")
        draw_photo_grid(r.get('pb_fotos_recti', []), "B. Data Rectifier")
        draw_photo_grid(r.get('pb_fotos_kelistrikan', []), "C. Data Kelistrikan & Genset")
    else:
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

def build_pptx(db_list):
    prs = Presentation()
    slide_layout = prs.slide_layouts[0]
    slide = prs.slides.add_slide(slide_layout)
    title = slide.shapes.title
    subtitle = slide.placeholders[1]
    title.text = "Laporan Lengkap Preventive Maintenance"
    subtitle.text = f"Total Site Terinspeksi: {len(db_list)}\nGenerated on: {datetime.date.today()}"
    
    for r in db_list:
        slide_layout = prs.slide_layouts[5]
        slide = prs.slides.add_slide(slide_layout)
        title = slide.shapes.title
        title.text = f"Site: {r.get('site_name', '-')} | Status: {r.get('status', '-')}"
        
        table_shape = slide.shapes.add_table(6, 2, Inches(0.5), Inches(1.5), Inches(4.5), Inches(3.0))
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
                for paragraph in cell.text_frame.paragraphs: paragraph.font.size = Pt(13)
        
        site_urls = r.get('url_sites', [])
        if site_urls and len(site_urls) > 0:
            try:
                img_url = site_urls[0].replace("upload/v", "upload/c_limit,w_500,q_80,f_jpg/v")
                resp = requests.get(img_url, timeout=5)
                if resp.status_code == 200:
                    image_stream = BytesIO(resp.content)
                    slide.shapes.add_picture(image_stream, Inches(5.3), Inches(1.5), width=Inches(4.2))
            except: pass
                
    out = BytesIO()
    prs.save(out)
    return out.getvalue()

# -------------------------------------------------------------------------
# SETUP DATABASE (SINKRONISASI KE GOOGLE SHEETS)
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

if 'laporan_db' not in st.session_state: st.session_state['laporan_db'] = fetch_data_from_gsheets()

# -------------------------------------------------------------------------
# FUNGSI PEMBANTU UNTUK INPUT + UPLOAD (Spesifik untuk Pangkalanbun)
# -------------------------------------------------------------------------
def input_with_photo(label, input_type="text", key="", options=None, value=""):
    c1, c2 = st.columns([2.5, 1])
    with c1:
        if input_type == "text": 
            v = st.text_input(label, value=value, key=f"t_{key}")
        elif input_type == "select": 
            v = st.selectbox(label, options=options, key=f"s_{key}")
        elif input_type == "number":
            v = st.number_input(label, value=0, key=f"n_{key}")
    with c2:
        f = st.file_uploader(f"📸 {label}", key=f"f_{key}")
    return v, f

# -------------------------------------------------------------------------
# NAVIGASI UTAMA
# -------------------------------------------------------------------------
st.sidebar.markdown("<h2 style='text-align: center; color: var(--primary-color);'>⚡ NAVIGASI</h2>", unsafe_allow_html=True)

if st.session_state['role'] == 'Admin':
    st.sidebar.markdown("<div style='text-align: center; background-color: var(--secondary-background-color); padding: 10px; border-radius: 8px; border: 1px solid var(--primary-color);'>Status: <b>🟢 ADMIN</b></div>", unsafe_allow_html=True)
    menu_options = ["📝 Form Preventive Check", "📊 Hasil Laporan & Dashboard", "📈 Monitoring Improvement"]
else:
    st.sidebar.markdown("<div style='text-align: center; background-color: var(--secondary-background-color); padding: 10px; border-radius: 8px; border: 1px solid gray;'>Status: <b>👷‍♂️ TIM LAPANGAN</b></div>", unsafe_allow_html=True)
    menu_options = ["📸 Form SiUPDATE", "📊 Hasil Laporan & Dashboard", "📈 Monitoring Improvement"]

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
# MENU 1: FORM PENGISIAN (ADMIN: PALANGKARAYA/PANGKALANBUN | LAPANGAN: SIUPDATE)
# =========================================================================
if menu in ["📝 Form Preventive Check", "📸 Form SiUPDATE"]:
    render_header_logo()
    
    if menu == "📸 Form SiUPDATE":
        st.markdown("<h1 style='text-align: center;'>📸 Form SiUPDATE (Pelaporan Lapangan)</h1>", unsafe_allow_html=True)
        standar_form = "SiUPDATE"
    else:
        st.markdown("<h1 style='text-align: center;'>⚡ Form Preventive Maintenance</h1>", unsafe_allow_html=True)
        # --- TOGGLE PILIHAN FORM BAGI ADMIN ---
        st.info("Pilih format pengisian data berdasarkan Standar NOP / Regional:")
        standar_form = st.radio("Format Form Laporan:", ["NOP Palangkaraya (Format Asli)", "NOP Pangkalan Bun (Site Visit Report)"], horizontal=True)

    # -------------------------------------------------------------------------
    # BLOK A: FORM NOP PALANGKARAYA & SIUPDATE (KODE LAMA UTUH)
    # -------------------------------------------------------------------------
    if standar_form in ["NOP Palangkaraya (Format Asli)", "SiUPDATE"]:
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
            st.markdown("### 📝 Kelengkapan Data SPS & SCC")
            c_pnl1, c_pnl2, c_pnl3 = st.columns(3)
            with c_pnl1: status_panel = st.selectbox("Status Panel", ["Normal", "Kotor", "Retak / Pecah", "Terbakar", "Kritis / Kosong"])
            with c_pnl2: jml_panel_rusak = st.number_input("Jumlah Panel Rusak", min_value=0, value=0)
            with c_pnl3: merk_panel = st.text_input("Type Merk Panel", placeholder="Contoh: Canadian Solar 550Wp")

            c_scc1, c_scc2, c_scc3 = st.columns(3)
            with c_scc1: status_scc = st.selectbox("Status SCC", ["Normal", "Alarm", "Rusak", "Tidak Ada"])
            with c_scc2: jml_scc_nok = st.number_input("Jumlah SCC NOK", min_value=0, value=0)
            with c_scc3: merk_scc = st.text_input("Type/Merk SCC", placeholder="Contoh: Huawei / Shoto")
            st.divider()

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
            st.markdown("### 🔌 Pengecekan Jaringan PLN & PSB")
            c_psb1, c_psb2 = st.columns(2)
            with c_psb1: possibility_psb = st.selectbox("Possibility PSB (PLN)", ["Bisa Dilakukan", "Sulit / Terlalu Jauh", "Sudah Tersambung", "Tidak Memungkinkan"])
            with c_psb2: jarak_pln = st.text_input("Estimasi Jarak tower ke jaringan PLN", placeholder="Contoh: 150 Meter / 2 KM")
            st.divider()

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
            st.markdown("### 📝 Kelengkapan Data Baterai")
            c_batm1, c_batm2, c_batm3 = st.columns(3)
            with c_batm1: status_battery = st.selectbox("Status Total Baterai", ["Normal", "Degradasi (Drop)", "Rusak / Menggelembung", "Kritis / Hilang"])
            with c_batm2: jml_bat_rusak = st.number_input("Jumlah Baterai Rusak", min_value=0, value=0)
            with c_batm3: tipe_battery = st.text_input("Type / Merk Baterai", placeholder="Contoh: VRLA Shoto 100Ah")
            st.divider()

            b1, b2 = st.columns(2)
            with b1:
                num_bat = st.number_input("Jumlah Baterai (Pengukuran)", min_value=1, value=4)
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
            remark = st.text_area("📝 Remark / Catatan Khusus Laporan:")
            final_status = st.radio("Status Akhir Site:", ["Normal", "Minor Issue", "Major/Critical"])

            if st.button("🚀 UPLOAD & SINKRONKAN DATA", type="primary", use_container_width=True):
                if not site_name: st.error("⚠️ Mohon isi Nama Site!")
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
                            if p.get("tipe") == "Individu": p_res.append({"Tipe": "Individu", "Panel": p["id"], "Voc": p["voc"], "Isc": p["isc"], "Kondisi": p["kondisi"], "URL_Before": upload_image(p["foto_before_obj"]), "URL_After": upload_image(p["foto_after_obj"])})
                            else: p_res.append({"Tipe": "Seri", "Panel": p["id"], "Qty": p["qty"], "Voc": p["voc"], "Isc": p["isc"], "Kondisi": p["kondisi"], "URLs_Before": upload_multiple_images(p["foto_before_objs"]), "URLs_After": upload_multiple_images(p["foto_after_objs"])})
                        
                        b_res = []
                        for b in bat_data: b_res.append({"Baterai": b["id"], "Voltase": b["voltase"], "Suhu": b["suhu"], "Kondisi": b["kondisi"], "URL_Fotos": upload_multiple_images(b["foto_objs"])})

                        report_dict = {
                            "timestamp": datetime.datetime.now().strftime("%Y-%m-%d %H:%M:%S"),
                            "site_name": site_name, "kategori": site_category if standar_form != "SiUPDATE" else "SiUPDATE", 
                            "nop": nop_area, "teknisi": technician_name,
                            "status": final_status, "action": action_taken, "sparepart": sparepart_needed,
                            
                            "status_panel": status_panel, "jml_panel_rusak": jml_panel_rusak, "merk_panel": merk_panel,
                            "status_scc": status_scc, "jml_scc_nok": jml_scc_nok, "merk_scc": merk_scc,
                            "status_battery": status_battery, "jml_bat_rusak": jml_bat_rusak, "tipe_battery": tipe_battery,
                            "possibility_psb": possibility_psb, "jarak_pln": jarak_pln, "remark": remark,
                            
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
                            st.success("✅ BERHASIL! Data Laporan telah diamankan di Google Sheets.")
                        else:
                            st.error("⚠️ Gagal menyambung ke Spreadsheet.")

    # -------------------------------------------------------------------------
    # BLOK B: TAMBAHAN BARU FORM NOP PANGKALAN BUN (SITE VISIT REPORT)
    # -------------------------------------------------------------------------
    elif standar_form == "NOP Pangkalan Bun (Site Visit Report)":
        
        st.info("💡 **Format Khusus Site Visit Report (Pangkalan Bun).** Anda diwajibkan melampirkan foto evidence/bukti di sebelah kanan untuk setiap isian data kelistrikan dan parameter.")
        tabs_pb = st.tabs(["📌 Info Site & Power", "🔌 Rectifier & ENVA", "⚡ Kelistrikan & ACPDB", "⛽ Genset & Log", "📤 Activity & Submit"])
        
        # Kolektor Data & Foto
        pb_payload = {}
        pb_fotos_info = []
        pb_fotos_recti = []
        pb_fotos_kelistrikan = []
        
        with tabs_pb[0]:
            st.markdown("### 📌 INFORMASI SITE")
            pb_te, pb_f1 = input_with_photo("TE Name", key="pb_te")
            pb_sid, pb_f2 = input_with_photo("Site ID", key="pb_sid")
            pb_sname, pb_f3 = input_with_photo("Site Name", key="pb_sname")
            pb_tick, pb_f4 = input_with_photo("Ticket#", key="pb_tick")
            pb_tgl, pb_f5 = input_with_photo("Tanggal On Site", "text", value=datetime.date.today().strftime("%d %B %Y"), key="pb_tgl")
            st.divider()
            st.markdown("### ⚡ CATUAN POWER")
            pb_catuan, pb_f6 = input_with_photo("Catuan Power Utama", "select", options=["GENSET", "PLN", "SPS", "LAINNYA"], key="pb_catuan")
            
            pb_payload.update({"te_name": pb_te, "site_id": pb_sid, "site_name": pb_sname, "ticket": pb_tick, "tgl_onsite": pb_tgl, "power_utama": pb_catuan})
            pb_fotos_info.extend([pb_f1, pb_f2, pb_f3, pb_f4, pb_f5, pb_f6])
            
        with tabs_pb[1]:
            st.markdown("### 🔌 DATA RECTIFIER")
            pb_jml_rect = st.number_input("Jumlah Rectifier", min_value=1, value=2, key="pb_jmlrect")
            
            pb_recti_list = []
            for i in range(int(pb_jml_rect)):
                with st.expander(f"⚙️ Rectifier #{i+1}", expanded=(i==0)):
                    tr, f_tr = input_with_photo("Tipe Rectifier (ZTE, dll)", key=f"tr_{i}")
                    tc, f_tc = input_with_photo("Type Controler (Csu 501B, dll)", key=f"tc_{i}")
                    tm, f_tm = input_with_photo("Type Module (ZXD2000, dll)", key=f"tm_{i}")
                    jb, f_jb = input_with_photo("Jumlah Battery (Bank)", key=f"jb_{i}")
                    bok, f_bok = input_with_photo("Jenis Batt OK", key=f"bok_{i}")
                    bnok, f_bnok = input_with_photo("Jenis Batt NOK", key=f"bnok_{i}")
                    jmo, f_jmo = input_with_photo("Jumlah Module OK", "number", key=f"jmo_{i}")
                    jmn, f_jmn = input_with_photo("Jumlah Module NOK", "number", key=f"jmn_{i}")
                    bcl, f_bcl = input_with_photo("BCL (%)", key=f"bcl_{i}")
                    lc, f_lc = input_with_photo("Load Current (A)", key=f"lc_{i}")
                    dcv, f_dcv = input_with_photo("DC Voltage (V)", key=f"dcv_{i}")
                    lp, f_lp = input_with_photo("LVD Prio", key=f"lp_{i}")
                    lnp, f_lnp = input_with_photo("LVD Non Prio", key=f"lnp_{i}")
                    elp, f_elp = input_with_photo("EQP LVD Prio", key=f"elp_{i}")
                    elnp, f_elnp = input_with_photo("EQP LVD Non Prio", key=f"elnp_{i}")
                    bbt, f_bbt = input_with_photo("Battery Backup Time", key=f"bbt_{i}")
                    
                    pb_recti_list.append({
                        "id": f"Rectifier #{i+1}", "tipe": tr, "controler": tc, "module": tm, "jml_batt": jb, "batt_ok": bok, "batt_nok": bnok,
                        "mod_ok": jmo, "mod_nok": jmn, "bcl": bcl, "load": lc, "dc_v": dcv, "lvd_p": lp, "lvd_np": lnp, "eqp_p": elp, "eqp_np": elnp, "backup": bbt
                    })
                    pb_fotos_recti.extend([f_tr, f_tc, f_tm, f_jb, f_bok, f_bnok, f_jmo, f_jmn, f_bcl, f_lc, f_dcv, f_lp, f_lnp, f_elp, f_elnp, f_bbt])
            
            st.divider()
            st.markdown("### 🚨 ENVA Setelah Simulasi (Bukan Cabut Kabel)")
            opt_enva = ["OK", "NOK", "-"]
            em, f_em = input_with_photo("Main Fail", "select", options=opt_enva, key="em")
            er, f_er = input_with_photo("Rectifier Fail", "select", options=opt_enva, key="er")
            el, f_el = input_with_photo("Low Batt", "select", options=opt_enva, key="el")
            eh, f_eh = input_with_photo("High Temp", "select", options=opt_enva, key="eh")
            egr, f_egr = input_with_photo("Genset Run", "select", options=opt_enva, key="egr")
            egf, f_egf = input_with_photo("Genset Fail", "select", options=opt_enva, key="egf")
            
            pb_payload["enva"] = {"main_fail": em, "rect_fail": er, "low_batt": el, "high_temp": eh, "gen_run": egr, "gen_fail": egf}
            pb_fotos_recti.extend([f_em, f_er, f_el, f_eh, f_egr, f_egf])

        with tabs_pb[2]:
            st.markdown("### ⚡ DATA KELISTRIKAN")
            kd, f_kd = input_with_photo("Daya KWH", key="kd")
            kj, f_kj = input_with_photo("Jumlah Phasa", key="kj")
            kv1, f_kv1 = input_with_photo("Tegangan R-S-T thd N", key="kv1")
            kv2, f_kv2 = input_with_photo("Tegangan Phasa RS-RT-ST", key="kv2")
            ka, f_ka = input_with_photo("Pengukuran Arus R-S-T", key="ka")
            kgn, f_kgn = input_with_photo("Tegangan Pengukuran G-N", key="kgn")
            
            pb_payload["kelistrikan"] = {"daya": kd, "phasa": kj, "v_rst_n": kv1, "v_rs_rt_st": kv2, "arus_rst": ka, "v_gn": kgn}
            pb_fotos_kelistrikan.extend([f_kd, f_kj, f_kv1, f_kv2, f_ka, f_kgn])
            
            st.divider()
            st.markdown("### 🎛 DATA ACPDB")
            aa, f_aa = input_with_photo("Ada ACPDB", "select", options=["Ya", "Tidak", "-"], key="aa")
            av1, f_av1 = input_with_photo("Tegangan R-S-T thd N", key="av1")
            av2, f_av2 = input_with_photo("Tegangan Phasa RS-RT-ST", key="av2")
            aar, f_aar = input_with_photo("Pengukuran Arus R-S-T", key="aar")
            agn, f_agn = input_with_photo("Tegangan Pengukuran G-N", key="agn")
            
            pb_payload["acpdb"] = {"ada": aa, "v_rst_n": av1, "v_rs_rt_st": av2, "arus_rst": aar, "v_gn": agn}
            pb_fotos_kelistrikan.extend([f_aa, f_av1, f_av2, f_aar, f_agn])
            
        with tabs_pb[3]:
            st.markdown("### ⛽ DATA GENSET")
            ga, f_ga = input_with_photo("Ada Genset", "select", options=["Ya", "Tidak"], key="ga")
            gc, f_gc = input_with_photo("Capacity Genset (kVA)", key="gc")
            gs, f_gs = input_with_photo("Genset Status (OK/NOK)", key="gs")
            gco, f_gco = input_with_photo("Config Running", key="gco")
            gt, f_gt = input_with_photo("Kapasitas Tangki", key="gt")
            grh, f_grh = input_with_photo("Running Hour", key="grh")
            gsb, f_gsb = input_with_photo("Sisa BBM", key="gsb")
            
            pb_payload["genset"] = {"ada": ga, "cap": gc, "status": gs, "config": gco, "tangki": gt, "rh": grh, "bbm": gsb}
            pb_fotos_kelistrikan.extend([f_ga, f_gc, f_gs, f_gco, f_gt, f_grh, f_gsb])

        with tabs_pb[4]:
            st.markdown("### 📝 ACTIVITY REPORT & CATATAN")
            pb_act1 = st.multiselect("1. Power & CME Activity", ["PMS/ PMG", "Pembersihan/ PM SPS", "Regional Programm", "NOP Programm", "Site Down Troubleshoot", "PLN Undervolt", "Load Balancing", "Rectifier Troubleshoot", "Battery Troubleshoot/ Claim Warranty", "Pergantian Rectifier Controller (SC/ CSU)", "Mapping Rectifier Controller (SC/ CSU)", "Grounding"])
            pb_act2 = st.multiselect("2. Transport Activity", ["Packet Loss", "NMS Unremote", "On Site VLAN Balancing"])
            pb_act3 = st.multiselect("3. OSA Impact Service Activity", ["Validasi ENVA", "BBU Min Voltage", "RRU Min Voltage", "Cell Down", "Disable PA", "Int. Fault", "Cell Flicker", "BBU Temp", "RRU Temp", "Zero Traffic", "Zero Payload", "VSWR", "Optical port link fault", "GNSS/GPS", "BFAN/FCE"])
            pb_temuan = st.text_area("Temuan/ Catatan:", placeholder="1. \n2. \n3.")
            
            pb_status_akhir = st.radio("Status Akhir Site (Umum):", ["Normal", "Minor Issue", "Major/Critical"])
            
            if st.button("🚀 SUBMIT PANGKALAN BUN REPORT", type="primary", use_container_width=True):
                if not pb_sid: st.error("⚠️ Mohon isi Site ID / Name!")
                else:
                    with st.spinner("⏳ Sedang mengunggah puluhan foto bukti dan menyusun Laporan Site Visit (Ini mungkin memakan waktu)..."):
                        
                        # Upload Image Arrays (Removing None values before upload to save time)
                        up_info = upload_multiple_images([f for f in pb_fotos_info if f is not None], "pbun_info")
                        up_rect = upload_multiple_images([f for f in pb_fotos_recti if f is not None], "pbun_recti")
                        up_kel = upload_multiple_images([f for f in pb_fotos_kelistrikan if f is not None], "pbun_kelistrikan")
                        
                        report_dict = {
                            "timestamp": datetime.datetime.now().strftime("%Y-%m-%d %H:%M:%S"),
                            "site_name": pb_sid + " / " + pb_sname,
                            "kategori": "Pangkalan Bun", 
                            "nop": "Pangkalan Bun", 
                            "teknisi": pb_te,
                            "status": pb_status_akhir, 
                            "action": pb_temuan,
                            
                            # Dictionary Khusus Pangkalan Bun
                            "pb_te": pb_te, "pb_site_id": pb_sid, "pb_ticket": pb_tick, "pb_power": pb_catuan,
                            "pb_rectifiers": pb_recti_list,
                            "pb_enva": pb_payload["enva"],
                            "pb_kelistrikan": pb_payload["kelistrikan"],
                            "pb_acpdb": pb_payload["acpdb"],
                            "pb_genset": pb_payload["genset"],
                            "pb_act_power": pb_act1, "pb_act_trans": pb_act2, "pb_act_osa": pb_act3,
                            
                            # Simpan URL Foto
                            "pb_fotos_info": up_info,
                            "pb_fotos_recti": up_rect,
                            "pb_fotos_kelistrikan": up_kel,
                            
                            # Data dummy agar tidak error di dashboard bawaan
                            "panel_data": [], "battery_data": [], "url_sites": up_info, "total_load": "-", "rect_out_v": "-", "sparepart": "-"
                        }
                        
                        sheet = connect_gsheets()
                        if sheet:
                            row_data = [
                                report_dict['timestamp'], report_dict['site_name'], report_dict['nop'], report_dict['teknisi'], report_dict['status'], json.dumps(report_dict)
                            ]
                            sheet.append_row(row_data)
                            st.cache_data.clear()
                            st.session_state['laporan_db'].append(report_dict)
                            st.success("✅ BERHASIL! Data Site Visit Pangkalan Bun beserta Foto Evidence telah tersimpan.")

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
        col_m1, col_m2, col_m3 = st.columns(3)
        col_m1.metric("Total Laporan Masuk", len(db))
        col_m2.metric("Database Terhubung", "Google Sheets", "Online")
        col_m3.metric("Cloud Storage", "Cloudinary", "Active")
        st.markdown("---")
        
        # ---------------------------------------------------------------------
        # FILTER KATEGORI SITE
        # ---------------------------------------------------------------------
        st.markdown("#### 🗂️ Filter Kategori Site", unsafe_allow_html=True)
        filter_kat = st.radio("Pilih kategori laporan yang ingin ditampilkan:", ["Semua", "SPS", "Site Reguler", "SiUPDATE", "Pangkalan Bun"], horizontal=True)
        
        filtered_indices = []
        for idx, r in enumerate(db):
            kat = r.get('kategori', 'SPS')
            if filter_kat == "Semua" or kat == filter_kat: filtered_indices.append(idx)
                
        if len(filtered_indices) == 0:
            st.info(f"Tidak ada data Laporan untuk kategori: {filter_kat}")
            st.stop()

        # ---------------------------------------------------------------------
        # TABEL EXCEL & PPTX
        # ---------------------------------------------------------------------
        st.markdown("### 📥 Ekspor Laporan Rekapitulasi (Excel & PPTX)")
        summary_list = []
        export_db = [db[idx] for idx in filtered_indices]
        
        for r in export_db:
            panel_issues = [p.get('Panel', 'Panel') for p in r.get('panel_data', []) if p.get('kondisi') and p.get('kondisi') != "Baik"]
            bat_issues = [b.get('Baterai', 'Baterai') for b in r.get('battery_data', []) if b.get('Kondisi') and b.get('Kondisi') != "Normal"]

            summary_list.append({
                "Timestamp": r.get('timestamp', '-'), "Nama Site": r.get('site_name', '-'),
                "Kategori": r.get('kategori', 'SPS'), "NOP": r.get('nop', '-'),
                "Teknisi": r.get('teknisi', '-'), "Status Akhir": r.get('status', '-'),
                "Kondisi Site": r.get('site_cond', '-'), "Fisik Tower": r.get('tower_cond', '-'),
                "Shading Panel": r.get('shading_status', '-'),
                "Status Panel": r.get('status_panel', '-'), "Jumlah Panel Rusak": r.get('jml_panel_rusak', 0),
                "Type Merk Panel": r.get('merk_panel', '-'), "Status Battery": r.get('status_battery', '-'),
                "Jumlah battery Rusak": r.get('jml_bat_rusak', 0), "Type Battery": r.get('tipe_battery', '-'),
                "status SCC": r.get('status_scc', '-'), "Jumlah SCC NOK": r.get('jml_scc_nok', 0),
                "Possibility PSB": r.get('possibility_psb', '-'), "Jarak PLN": r.get('jarak_pln', '-'),
                "SPS Rusak (Detail)": ", ".join(panel_issues) if panel_issues else "Aman",
                "Load Beban (A)": r.get('total_load', '-'), "PLN": r.get('pln_status', '-'),
                "Rectifier": r.get('rect_brand', '-'), "Voltase Recti": r.get('rect_out_v', '-'),
                "Genset": r.get('genset_status', '-'), "Level BBM (%)": r.get('fuel_pct', '-'),
                "Baterai Rusak (Detail)": ", ".join(bat_issues) if bat_issues else "Aman",
                "Action / Pekerjaan": r.get('action', '-'), "Sparepart": r.get('sparepart', '-'), "Remark": r.get('remark', '-')
            })
            
        df_export = pd.DataFrame(summary_list)
        col_dl_ex, col_dl_ppt = st.columns(2)
        with col_dl_ex:
            try:
                excel_buffer = BytesIO()
                with pd.ExcelWriter(excel_buffer, engine='xlsxwriter') as writer:
                    df_export.to_excel(writer, index=False, sheet_name='Database PM')
                    workbook, worksheet = writer.book, writer.sheets['Database PM']
                    
                    header_format = workbook.add_format({'bold': True, 'font_color': 'white', 'bg_color': '#112240', 'border': 1, 'align': 'center', 'valign': 'vcenter'})
                    for col_num, value in enumerate(df_export.columns.values): worksheet.write(0, col_num, value, header_format)
                    cell_center = workbook.add_format({'align': 'center', 'valign': 'vcenter', 'border': 1})
                    cell_left = workbook.add_format({'align': 'left', 'valign': 'vcenter', 'border': 1, 'text_wrap': True})

                    for i, col_name in enumerate(df_export.columns):
                        if col_name in ["Action / Pekerjaan", "Remark", "Possibility PSB", "Type Merk Panel", "SPS Rusak (Detail)", "Baterai Rusak (Detail)"]: worksheet.set_column(i, i, 35, cell_left)
                        else: worksheet.set_column(i, i, 18, cell_center)
                    worksheet.freeze_panes(1, 0)
                    worksheet.autofilter(0, 0, len(df_export), len(df_export.columns) - 1)
                    
                st.download_button(label=f"📊 Download Excel ({filter_kat})", data=excel_buffer.getvalue(), file_name=f"Database_PM_{datetime.date.today()}.xlsx", mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet", type="primary", use_container_width=True)
            except Exception: pass

        with col_dl_ppt:
            if HAS_PPTX:
                pptx_data = build_pptx(export_db)
                st.download_button(label=f"📽 Download Presentasi Report ({filter_kat})", data=pptx_data, file_name=f"Report_PM_KUT_{datetime.date.today()}.pptx", mime="application/vnd.openxmlformats-officedocument.presentationml.presentation", type="primary", use_container_width=True)
            else: st.error("Library `python-pptx` belum ter-install di server.")
        st.divider()

        # ---------------------------------------------------------------------
        # EXPANDER LIST LAPORAN (MENAMPILKAN PALANGKARAYA / PANGKALANBUN / SIUPDATE)
        # ---------------------------------------------------------------------
        for i in reversed(filtered_indices):
            r = db[i]
            site_id = r.get('site_name', 'Unknown')
            kategori_site = r.get('kategori', 'SPS')
            stat = r.get('status', '')
            icon = "🟢" if stat == "Normal" else "🟡" if stat in ["Minor Issue", "Update Progress"] else "🔴"
            
            # WA Teks dinamis
            wa_text = f"*REPORT* ⚡\n📍 *Site:* {site_id} ({kategori_site})\n📅 *Tanggal:* {r.get('timestamp', '-')}\n👷 *Pelaksana:* {r.get('teknisi', '-')}\n\n*Catatan/Temuan:*\n{r.get('action', '-')}"
            wa_url = f"https://wa.me/?text={urllib.parse.quote(wa_text)}"

            with st.expander(f"{icon}  |  {site_id}  |  {kategori_site}  |  {r.get('timestamp', '')}  |  Status: {stat}"):
                
                if kategori_site != "SiUPDATE":
                    if st.checkbox("📄 Buat Berita Acara (PDF Resmi)", key=f"prep_pdf_{i}"):
                        with st.spinner("⏳ Rendering Dokumen PDF Resolusi Tinggi..."):
                            try:
                                pdf_bytes = build_pdf(r)
                                st.download_button(label="📥 Download PDF Berita Acara", data=pdf_bytes, file_name=f"Berita_Acara_{site_id}.pdf", mime="application/pdf", key=f"dl_pdf_{i}", type="primary")
                            except Exception as e: st.error(f"Terjadi kesalahan saat menyusun PDF: {e}")
                
                st.markdown("<hr style='border: 1px solid var(--primary-color); margin: 15px 0;'>", unsafe_allow_html=True)
                c_btn1, c_btn2 = st.columns(2)
                with c_btn1: st.link_button("📱 Share Rangkuman ke WhatsApp", wa_url, use_container_width=True)
                with c_btn2: st.link_button("📈 Buka Database Spreadsheet Target", f"https://docs.google.com/spreadsheets/d/{SHEET_ID}/edit", use_container_width=True)
                
                # --- TAMPILAN KHUSUS NOP PANGKALAN BUN ---
                if kategori_site == "Pangkalan Bun":
                    st.markdown("### 📋 SITE VISIT REPORT (Pangkalan Bun)")
                    st.write(f"**TE Name:** {r.get('pb_te', '-')} | **Ticket:** {r.get('pb_ticket', '-')}")
                    st.write(f"**Catuan Power:** {r.get('pb_power', '-')} | **Status Akhir:** {stat}")
                    
                    tpb1, tpb2, tpb3, tpb4 = st.tabs(["📌 Info & Recti", "⚡ Kelistrikan & ACPDB", "⛽ Genset & ENVA", "📝 Temuan & Activity"])
                    
                    with tpb1:
                        st.markdown("**Data Rectifier:**")
                        for rect in r.get('pb_rectifiers', []):
                            st.write(f"🔹 **{rect.get('id')}** - Tipe: {rect.get('tipe')} | BCL: {rect.get('bcl')}% | Beban: {rect.get('load')}A | Voltase: {rect.get('dc_v')}V")
                        tampilkan_grid_foto(r.get('pb_fotos_info'), "📸 Foto Info Site & Power")
                        tampilkan_grid_foto(r.get('pb_fotos_recti'), "📸 Foto Bukti Rectifier & Module")
                        
                    with tpb2:
                        kel = r.get('pb_kelistrikan', {})
                        acp = r.get('pb_acpdb', {})
                        st.write(f"**Kelistrikan:** Daya KWH: {kel.get('daya', '-')} | Phasa: {kel.get('phasa', '-')} | Teg. G-N: {kel.get('v_gn', '-')}")
                        st.write(f"**ACPDB:** Tersedia: {acp.get('ada', '-')} | Teg. RST thd N: {acp.get('v_rst_n', '-')}")
                        tampilkan_grid_foto(r.get('pb_fotos_kelistrikan'), "📸 Foto Kelistrikan & ACPDB")
                        
                    with tpb3:
                        gen = r.get('pb_genset', {})
                        en = r.get('pb_enva', {})
                        st.write(f"**Genset:** Ada: {gen.get('ada', '-')} | Kapasitas: {gen.get('cap', '-')} | Tangki: {gen.get('tangki', '-')} | Sisa: {gen.get('bbm', '-')}")
                        st.write(f"**ENVA:** Main Fail: {en.get('main_fail', '-')} | Rectifier Fail: {en.get('rect_fail', '-')} | High Temp: {en.get('high_temp', '-')}")
                        
                    with tpb4:
                        st.markdown("**Catatan/Temuan Tim:**")
                        st.info(r.get('action', '-'))
                        st.markdown("**Activity Report:**")
                        if r.get('pb_act_power'): st.write(f"- Power & CME: {', '.join(r.get('pb_act_power'))}")
                        if r.get('pb_act_trans'): st.write(f"- Transport: {', '.join(r.get('pb_act_trans'))}")
                        if r.get('pb_act_osa'): st.write(f"- OSA Impact: {', '.join(r.get('pb_act_osa'))}")

                # --- TAMPILAN KHUSUS SiUPDATE ---
                elif kategori_site == "SiUPDATE":
                    st.markdown(f"**👨‍🔧 Pelaksana / Tim:** {r.get('teknisi', '-')}")
                    st.markdown(f"**📌 Kategori Update:** {r.get('kategori_update', '-')}")
                    st.markdown(f"**🔧 Deskripsi Update:** {r.get('action', '-')}")
                    st.markdown("<br>", unsafe_allow_html=True)
                    tampilkan_grid_foto(r.get('url_sites'), "📸 Dokumentasi Foto SiUPDATE")

                # --- TAMPILAN NOP PALANGKARAYA (Asli) ---
                else:
                    st.markdown(f"**👨‍🔧 Pelaksana (Teknisi):** {r.get('teknisi', '-')} &nbsp;&nbsp;|&nbsp;&nbsp; **⚡ Beban Load:** {r.get('total_load', '-')} A")
                    st.markdown(f"**🔧 Action Dikerjakan:** {r.get('action', '-')}")
                    if r.get('sparepart'): st.warning(f"**📦 Sparepart Diganti:** {r['sparepart']}")
                    if r.get('remark'): st.info(f"**📝 Remark:** {r['remark']}")
                    
                    st.markdown("<br>", unsafe_allow_html=True)
                    
                    ltab1, ltab2, ltab3, ltab4, ltab5, ltab6 = st.tabs(["🏗️ 1. Fisik", "☀️ 2. Panel SPS", "🔌 3. Recti & Genset", "🔋 4. Baterai & Gnd", "📂 5. Datalog", "📈 6. Analisa Power BBU"])
                    
                    with ltab1:
                        st.write(f"- Kondisi Site: {r.get('site_cond', '-')} | Tower: {r.get('tower_cond', '-')} | Shading: {r.get('shading_status', '-')}")
                        tampilkan_grid_foto(r.get('url_sites'), "📸 View Site")
                        tampilkan_grid_foto(r.get('url_shadings'), "📸 Shading")

                    with ltab2:
                        st.write(f"**Status Panel:** {r.get('status_panel', '-')} ({r.get('jml_panel_rusak', 0)} Rusak) | **Merk Panel:** {r.get('merk_panel', '-')}")
                        for p in r.get('panel_data', []):
                            st.markdown(f"**{p.get('Panel', '-')}** | Voc: {p.get('Voc','-')}V | Isc: {p.get('Isc','-')}A")
                            cb, ca = st.columns(2)
                            with cb: tampilkan_grid_foto(p.get('URL_Before') or p.get('URLs_Before'), "📸 KONDISI PANEL")
                            with ca: tampilkan_grid_foto(p.get('URL_After') or p.get('URLs_After'), "✨ PENGUKURAN")

                    with ltab3:
                        c_p1, c_p2 = st.columns(2)
                        with c_p1:
                            st.write(f"**PLN:** {r.get('pln_status', '-')} | **Rectifier:** {r.get('rect_brand', '-')} ({r.get('rect_out_v', '-')}V)")
                            tampilkan_grid_foto(r.get('url_rects'), "📸 Rectifier")
                        with c_p2:
                            st.write(f"**Genset:** {r.get('genset_status', '-')} | **BBM:** {r.get('fuel_pct', '-')}%")
                            tampilkan_grid_foto(r.get('url_gensets'), "📸 Genset")

                    with ltab4:
                        st.markdown("### 🔋 Bank Baterai")
                        for b in r.get('battery_data', []):
                            st.write(f"🔹 **{b.get('Baterai', 'Baterai')}** - Voltase: {b.get('Voltase', '-') } V | Fisik: {b.get('Kondisi', '-')}")
                            tampilkan_grid_foto(b.get('URL_Fotos') or b.get('URL_Foto'))
                        st.markdown("### 🌍 Sistem Grounding")
                        st.write(f"⚡ **Tahanan Grounding:** {r.get('earth_ohm', '-')} Ohm")
                        tampilkan_grid_foto(r.get('url_grds'))

                    with ltab5:
                        st.markdown("📂 **Datalog Universal yang Tersimpan:**")
                        for dl in r.get('datalog_files', []): st.markdown(f"- 🔗 [{dl['name']}]({dl['url']})")

                    with ltab6:
                        st.markdown("### ⚡ Grafik Analisa Power & Voltage")
                        st.info(f"Fitur grafik membaca log excel dari folder repositori lokal Anda ({site_id}.xlsx)")
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
                            except Exception: pass

                # --- FITUR EDIT HANYA UNTUK ADMIN ---
                if st.session_state['role'] == 'Admin' and kategori_site != "Pangkalan Bun":
                    with st.container(border=True):
                        st.markdown("<h4 style='color: var(--primary-color);'>🛠 REVISI KETERANGAN / ACTION</h4>", unsafe_allow_html=True)
                        new_act = st.text_area("Edit Action / Tindakan", r.get('action',''), key=f"ea_{i}")
                        if st.button("💾 Simpan Perubahan Catatan ke Server", key=f"btn_{i}"):
                            with st.spinner("Menyimpan Revisi..."):
                                r['action'] = new_act
                                sheet = connect_gsheets()
                                if sheet:
                                    sheet.update_cell(i + 2, 6, json.dumps(r))
                                    st.cache_data.clear() 
                            st.success("✅ REVISI BERHASIL!")
                            st.rerun()

# =========================================================================
# MENU 3: EXCEL LIVE-EDITOR (MONITORING IMPROVEMENT)
# =========================================================================
elif menu == "📈 Monitoring Improvement":
    render_header_logo()
    st.markdown("<h1 style='text-align: center;'>📈 Master Tracker Improvement</h1>", unsafe_allow_html=True)
    
    file_master = "Monitoring_Availability_Improvement_Visit_SPS_NOP_PLK.xlsx"
    
    if not os.path.exists(file_master):
        st.error(f"❌ File '{file_master}' tidak ditemukan di sistem/server.")
    else:
        tab_dashboard, tab_tim, tab_editor = st.tabs(["📊 Kurva S", "👥 Progress Tim (Done/Berjalan)", "📝 Live Editor Data"])
        
        with tab_dashboard:
            st.markdown("### 📈 S-Curve Target vs Aktual")
            try:
                df_kurva = pd.read_excel(file_master, sheet_name='Kurva S', header=None)
                m1, m2, m3 = st.columns(3)
                m1.metric("Total Site Target", f"{int(df_kurva.iloc[6, 13]) if pd.notna(df_kurva.iloc[6, 13]) else 0} Site")
                m2.metric("Site dengan Plan", f"{int(df_kurva.iloc[8, 13]) if pd.notna(df_kurva.iloc[8, 13]) else 0} Site")
                m3.metric("Minggu Berjalan", f"Week {int(df_kurva.iloc[7, 13]) if pd.notna(df_kurva.iloc[7, 13]) else 0}")
                
                df_chart = df_kurva.iloc[6:58, :11].copy()
                df_chart.columns = ['Week', 'Mulai', 'Label', 'Plan_Minggu', 'Plan_Kumulatif', 'Actual_Minggu', 'Actual_Kumulatif', 'Pct_Plan', 'Pct_Actual', 'Gap', 'Achievement']
                df_chart = df_chart.dropna(subset=['Week', 'Plan_Kumulatif'])
                df_chart['Week'] = "W" + df_chart['Week'].astype(str)
                df_chart.set_index('Week', inplace=True)
                
                col_chart, col_table = st.columns([2, 1])
                with col_chart:
                    st.line_chart(df_chart[['Plan_Kumulatif', 'Actual_Kumulatif']], color=["#FF5252", "#64FFDA"])
                with col_table:
                    st.dataframe(df_chart[['Plan_Kumulatif', 'Actual_Kumulatif', 'Gap']], use_container_width=True)
            except Exception: pass

        with tab_tim:
            st.markdown("### 👥 Dashboard Eksekusi Lapangan")
            try:
                df_track = pd.read_excel(file_master, sheet_name='Tracker Improvement', header=None)
                headers_track = df_track.iloc[3].fillna("").astype(str).tolist()
                df_t = df_track.iloc[4:].copy()
                df_t.columns = headers_track
                
                status_cols = [c for c in headers_track if 'status' in c.lower() or 'progress' in c.lower() or 'keterangan' in c.lower() or 'aktual' in c.lower()]
                sel_col = st.selectbox("Pilih Kolom Status:", headers_track, index=headers_track.index(status_cols[0]) if status_cols else 0)
                
                if sel_col:
                    df_t[sel_col] = df_t[sel_col].astype(str).str.strip().str.upper()
                    counts = df_t[~df_t[sel_col].isin(['NAN', 'NAT', 'NONE', ''])][sel_col].value_counts().reset_index()
                    counts.columns = ['Status', 'Jumlah']
                    
                    if not counts.empty:
                        st.bar_chart(counts.set_index('Status'), color="#64FFDA")
            except Exception: pass

        with tab_editor:
            st.markdown("### 📝 Tabel Master Tracker & Viewer")
            sheet_choice = st.selectbox("Pilih Sheet Excel:", ["Tracker Improvement", "Jadwal Visit SPS", "Dashboard", "Kurva S"])
            try:
                df_raw = pd.read_excel(file_master, sheet_name=sheet_choice, header=None)
                if sheet_choice == "Dashboard": st.dataframe(df_raw, use_container_width=True, height=500)
                else:
                    headers = df_raw.iloc[3].fillna("").astype(str).tolist()
                    df_data = df_raw.iloc[4:].copy()
                    df_data.columns = headers
                    df_data = df_data.reset_index(drop=True)
                    edited_df = st.data_editor(df_data, use_container_width=True, num_rows="dynamic", key=f"editor_{sheet_choice}")
                    
                    if st.button("💾 Simpan Perubahan & Download Excel Utuh", type="primary"):
                        changes = st.session_state[f"editor_{sheet_choice}"]
                        if changes.get("edited_rows") or changes.get("added_rows"):
                            with st.spinner("Menyuntikkan data baru ke Master Excel (Menjaga Format)..."):
                                import openpyxl
                                wb = openpyxl.load_workbook(file_master)
                                ws = wb[sheet_choice]
                                for row_idx_str, col_changes in changes.get("edited_rows", {}).items():
                                    for col_name, new_val in col_changes.items():
                                        if col_name in headers: ws.cell(row=int(row_idx_str) + 5, column=headers.index(col_name) + 1).value = new_val
                                out_buffer = BytesIO()
                                wb.save(out_buffer)
                                st.success("✅ Berhasil! Perubahan telah disuntikkan ke file asli.")
                                st.download_button("📥 Download File Master Terupdate", out_buffer.getvalue(), f"Updated_{file_master}", "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet")
            except Exception: pass

st.markdown("""
    <div class="footer-okta">
        🚀 System Application & Database Management<br>
        <span>Create Data By Okta Pradika</span> © 2026
    </div>
""", unsafe_allow_html=True)
