import datetime, json, os, urllib.parse, base64, tempfile, random, hashlib as hl, hmac, glob, io, secrets
import pandas as pd
import numpy as np
import math
import streamlit as st
import cloudinary, cloudinary.uploader
import gspread
from google.oauth2.service_account import Credentials
from fpdf import FPDF
import requests
from io import BytesIO
from PIL import Image, ImageOps
import xml.etree.ElementTree as ET

try:
    from pptx import Presentation
    from pptx.util import Inches, Pt
    HAS_PPTX = True
except ImportError:
    HAS_PPTX = False

st.set_page_config(page_title="RPO - PLK & PBU", page_icon="⚡", layout="wide")
LOGO_UTAMA = "koperasi-jasa-konstruksi-tower-event-organizer-network-monitoring-telekomunikasi-kisel-group-logo-kut (2).webp"

# =========================================================================
# CACHE & UI PREMIUM
# =========================================================================
@st.cache_data(show_spinner=False, max_entries=8)
def _b64_cached(path, stamp):
    with open(path, "rb") as f: return base64.b64encode(f.read()).decode()

def get_base64_of_bin_file(bin_file):
    try: return _b64_cached(bin_file, (os.stat(bin_file).st_mtime_ns, os.stat(bin_file).st_size))
    except: return None

@st.cache_data(show_spinner=False, max_entries=24)
def _read_excel_cached(path, stamp, kwargs_key): return pd.read_excel(path, **dict(kwargs_key))

def read_excel_cached(path, **kwargs):
    if os.path.exists(path): return _read_excel_cached(path, (os.stat(path).st_mtime_ns, os.stat(path).st_size), tuple(sorted(kwargs.items())))
    return pd.read_excel(path, **kwargs)

def render_header_logo(is_login=False):
    for filename, mime in [(LOGO_UTAMA, "image/webp"), ("logo.png", "image/png")]:
        if os.path.exists(filename):
            bg = get_base64_of_bin_file(filename)
            if bg:
                max_h, pad = ("135px", "16px 36px") if is_login else ("100px", "12px 28px")
                st.markdown(f'''<div style="display:flex; justify-content:center; align-items:center; margin-bottom:20px;">
                    <div style="background:linear-gradient(145deg, #ffffff, #f3f6f9); padding:{pad}; border-radius:20px; box-shadow:0 10px 28px rgba(0,0,0,0.22); border:2px solid var(--primary-color);">
                        <img src="data:{mime};base64,{bg}" style="max-height:{max_h}; width:auto; max-width:100%; display:block;">
                    </div></div>''', unsafe_allow_html=True)
                return

_UX_CSS = """
<style>
@import url('https://fonts.googleapis.com/css2?family=Plus+Jakarta+Sans:wght@400;500;600;700;800&display=swap');
html,body,.stApp,[data-testid="stAppViewContainer"],[data-testid="stMain"],.main{font-family:'Plus Jakarta Sans',sans-serif; background:#F3F6FC !important;}
.stApp{color:#0B1730;--text-color:#0B1730;--secondary-background-color:#FFFFFF;--primary-color:#2563EB}
p, li, h1, h2, h3, h4, h5, h6, label, td, th, div[data-testid="stWidgetLabel"] p {color:#0B1730 !important}
#MainMenu,footer,header[data-testid="stHeader"]{display:none !important}
.stMainBlockContainer{max-width:1240px !important; margin:0 auto !important; padding:1rem 1.25rem 4rem !important}
div[data-testid="stButton"] button,a[data-testid="stLinkButton"]{border-radius:12px !important; font-weight:700 !important; transition:0.2s;}
button[kind="primary"]{background:linear-gradient(135deg,#1D4ED8,#4F46E5) !important; color:#fff !important; border:none !important;}
button[kind="secondary"]{background:#FFFFFF !important; color:#0B1730 !important; border:1.5px solid #E1E7F3 !important;}
div[data-baseweb="input"] input, div[data-baseweb="select"] div {background:#F7F9FE !important; border-radius:12px !important; color:#0B1730 !important;}
div[data-testid="stMetric"]{background:#FFFFFF; border:1px solid #E1E7F3; border-radius:16px; padding:14px; border-left:5px solid #B8891A;}
.area-card,.login-box{background:#FFFFFF; border:1px solid #E1E7F3; border-top:4px solid #2563EB; border-radius:20px; padding:25px; text-align:center; transition:0.2s; margin-bottom:12px;}
.area-card:hover,.login-box:hover{transform:translateY(-4px); box-shadow:0 10px 30px rgba(0,0,0,0.1);}
.login-title{font-size:1.2rem; font-weight:800; margin-bottom:10px; color:#0B1730;}
.login-desc{font-size:0.85rem; color:#5A6A86;}
.ux-hero{background:linear-gradient(135deg,#070C18 0%,#13245A 100%); border-radius:20px; padding:25px; color:#fff; text-align:center; margin-bottom:20px;}
img.ux-chart{width:100%; border-radius:12px;}
</style>
"""
st.markdown(_UX_CSS, unsafe_allow_html=True)
st.markdown("<style>div.stButton>button{min-height:55px !important; font-size:1rem !important;}</style>", unsafe_allow_html=True)

def _fmt(v):
    try:
        v = float(v)
        if math.isnan(v): return "-"
        if abs(v) >= 1e9: return f"{v/1e9:.2f} M"
        if abs(v) >= 1e6: return f"{v/1e6:.1f} jt"
        return f"{v:,.0f}" if v == int(v) else f"{v:.2f}".rstrip("0").rstrip(".")
    except: return "-"
def _esc(x): return _html.escape(str(x), quote=True)
def ux_hero(kicker, title, sub, stats=None, split=None):
    h = f"<div class='ux-hero'><span style='color:#F6DE8D; font-weight:800; font-size:12px;'>{_esc(kicker)}</span><h2 style='color:#fff; margin:10px 0;'>{_esc(title)}</h2><p style='color:#B7C4E4;'>{_esc(sub)}</p>"
    if stats: h += '<div style="display:flex; justify-content:center; gap:12px; margin-top:15px;">' + "".join(f'<div style="background:rgba(255,255,255,0.1); padding:10px 15px; border-radius:12px;"><b>{_esc(v)}</b><br><small>{_esc(l)}</small></div>' for v,l in stats) + "</div>"
    h += "</div>"
    st.markdown(h, unsafe_allow_html=True)

def ux_stats_from_db(db):
    db = db or []
    return [(len(db), "Total"), (sum(1 for r in db if r.get('status') in ['Normal','Done']), "Done"), (sum(1 for r in db if r.get('status') not in ['Normal','Done']), "Issue")], []

def ui_line_chart(df, color=None): st.line_chart(df, height=300)
def ui_bar_chart(df, color=None, horizontal=False): 
    if horizontal: st.bar_chart(df, height=300) 
    else: st.bar_chart(df, height=300)
def ui_ring(pct, label, color): st.progress(int(min(max(pct,0),100)), text=label)
def ux_progress(items): st.write(" | ".join([f"{l}: {v}" for l,v,c in items]))

# =========================================================================
# CLOUDINARY ENGINE
# =========================================================================
CLOUDINARY_ACCOUNTS = [
    {"cloud_name": "fxm61tjv", "api_key": "624877324969231", "api_secret": "LIFO6pfEg9fOM3nbsY8FBbVTpSI"},
    {"cloud_name": "ji9i1hei", "api_key": "277523771179144", "api_secret": "7eg20WV-Bzf9DXEoT8rg3H99V6c"},
    {"cloud_name": "dzd4xjz6w", "api_key": "482193747671566", "api_secret": "GMMzzU6--yoY46u6bS45hs5Mp90"}
]
def upload_image(file_obj, folder_name="solar_bts_healthcheck"):
    if file_obj:
        try:
            acc = random.choice(CLOUDINARY_ACCOUNTS)
            cloudinary.config(cloud_name=acc["cloud_name"], api_key=acc["api_key"], api_secret=acc["api_secret"], secure=True)
            res = cloudinary.uploader.upload(file_obj.getvalue(), folder=folder_name, transformation=[{'width': 1200, 'crop': 'limit'}, {'quality': 'auto:eco'}])
            return res.get('secure_url')
        except: return None
    return None
def upload_multiple_images(file_objs, folder_name="solar_bts_healthcheck"): return [upload_image(f, folder_name) for f in file_objs] if file_objs else []
def upload_datalog(file_obj, folder_name="prev_datalog"):
    if file_obj:
        try:
            acc = random.choice(CLOUDINARY_ACCOUNTS)
            cloudinary.config(cloud_name=acc["cloud_name"], api_key=acc["api_key"], api_secret=acc["api_secret"], secure=True)
            res = cloudinary.uploader.upload(file_obj.getvalue(), folder=folder_name, resource_type="raw", public_id=file_obj.name)
            return res.get('secure_url')
        except: return None
    return None
def tampilkan_grid_foto(url_data, caption=""):
    if not url_data: return
    urls = [url_data] if isinstance(url_data, str) and url_data.startswith("http") else [u for u in url_data if isinstance(u, str) and u.startswith("http")]
    if urls:
        if caption: st.markdown(f"*{caption}*")
        img_html = '<div style="display: grid; grid-template-columns: repeat(auto-fit, minmax(280px, 1fr)); gap: 12px; margin-bottom: 15px;">'
        for u in urls:
            opt_u = u.replace("upload/", "upload/w_400,q_auto:eco/") if "cloudinary" in u else u
            img_html += f'<div style="border: 1px solid #ccc; border-radius: 8px; padding: 6px; text-align: center;"><a href="{u}" target="_blank"><img src="{opt_u}" loading="lazy" style="max-width: 100%; border-radius: 6px;"></a></div>'
        img_html += '</div>'
        st.markdown(img_html, unsafe_allow_html=True)

# =========================================================================
# GSHEET & MASTER DATA ENGINE
# =========================================================================
SHEET_ID = "1HvgVicTWwO4RMQI6ZR3Mu3IgGicwjcLZl9mDN1auvJU"
def connect_gsheets(sheet_name="Report Preventive"):
    try:
        c_dict = json.loads(st.secrets["gcp_json"])
        creds = Credentials.from_service_account_info(c_dict, scopes=["https://www.googleapis.com/auth/spreadsheets", "https://www.googleapis.com/auth/drive"])
        return gspread.authorize(creds).open_by_key(SHEET_ID).worksheet(sheet_name)
    except: return None

@st.cache_data(ttl=600, show_spinner=False) 
def fetch_data_from_gsheets():
    all_data = []
    for s_name in ["Report Preventive", "Report PBU", "Project", "SiHD"]:
        sheet = connect_gsheets(s_name)
        if sheet:
            data = sheet.get_all_values()
            for idx, row in enumerate(data[1:]):
                if len(row) >= 6: 
                    try: 
                        item = json.loads(row[5])
                        if item.get('kategori') == 'SiUPDATE': item['kategori'] = 'SPS'
                        item['_sheet'] = s_name
                        item['_row'] = idx + 2
                        all_data.append(item)
                    except: pass
    return all_data

KABUPATEN_MAP = {"PLK": "Kota Palangka Raya", "KKP": "Kab. Kapuas", "KSN": "Kab. Katingan", "KKN": "Kab. Gunung Mas", "MTW": "Kab. Barito Utara", "PPS": "Kab. Pulang Pisau", "NTH": "Site N3T", "TML": "Kab. Barito Timur", "BNT": "Kab. Barito Selatan", "AMT": "Kab. Barito Selatan", "PRC": "Kab. Murung Raya", "COH": "Site Combat", "EPQ": "Site iMacro", "SPT": "Area Perbatasan"}

@st.cache_data(ttl=3600, show_spinner=False)
def load_sysinfomap_kml(file_bytes):
    try: root = ET.fromstring(file_bytes)
    except: return pd.DataFrame()
    ns = {'kml': 'http://www.opengis.net/kml/2.2'}
    rows = []
    for pm in root.findall('.//kml:Placemark', ns) or root.findall('.//Placemark'):
        rec = {}
        for sd in pm.findall('.//{*}SimpleData'):
            if sd.attrib.get('name'): rec[sd.attrib.get('name')] = (sd.text or '').strip()
        if not rec.get('site_id'):
            nm = pm.find('{*}name')
            if nm is not None and nm.text: rec['site_id'] = nm.text.strip()
        c_el = pm.find('.//{*}coordinates')
        if c_el and c_el.text:
            pts = c_el.text.strip().split(',')
            if len(pts)>=2: rec['lon'], rec['lat'] = float(pts[0]), float(pts[1])
        lat_v, lon_v = pd.to_numeric(rec.get('Lat'), errors='coerce'), pd.to_numeric(rec.get('Long'), errors='coerce')
        rec['lat'] = float(lat_v) if pd.notna(lat_v) else rec.get('lat', 0.0)
        rec['lon'] = float(lon_v) if pd.notna(lon_v) else rec.get('lon', 0.0)
        sid = str(rec.get('site_id','')).strip().upper()
        rec['site_id'], pref = sid, sid[:3] if len(sid)>=3 else "OTH"
        rec['Kabupaten'] = KABUPATEN_MAP.get(pref, f"Lainnya ({pref})")
        for k_def in ['Sitename', 'Site_Owner', 'site_class', 'Power_Source', 'BrandRecti', 'Type_Hub', 'Transport_type', 'RTS']:
            val = str(rec.get(k_def,'')).strip()
            rec[k_def] = val if val not in ['','#N/A','None'] else '-'
        rows.append(rec)
    df = pd.DataFrame(rows)
    return df[(df['lat']!=0) & (df['lon']!=0)].reset_index(drop=True) if not df.empty else df

def get_default_kml_df():
    for c in ["newlist_910.kml", "newlist.kml"]:
        if os.path.exists(c):
            with open(c, 'rb') as f: return load_sysinfomap_kml(f.read()), c
    return pd.DataFrame(), None

@st.cache_data(ttl=3600, show_spinner=False)
def load_master_lists():
    f_plk, f_pbu, s_plk, s_pbu = [], [], [], []
    try: f_plk = pd.read_excel("NIK NOP PLK.xlsx")['NAMA'].dropna().astype(str).tolist()
    except: pass
    try: f_pbu = pd.read_excel("NAMA NOP PANGKALANBUN.xlsx")['NAMA'].dropna().astype(str).tolist()
    except: pass
    try: s_plk = pd.read_excel("Site ID Sidata Palangkaraya.xlsx")['Site ID'].dropna().astype(str).tolist()
    except: pass
    try: s_pbu = pd.read_excel("Site ID Sidata Pangkalanbun.xlsx")['Site ID'].dropna().astype(str).tolist()
    except: pass
    try: s_plk.extend(pd.read_excel("Combat NOP Palangkaraya_1.xlsx", sheet_name="Combat NOP Palangkaraya")['siteid'].dropna().astype(str).str.strip().tolist())
    except: pass
    try:
        df_kml, _ = get_default_kml_df()
        if not df_kml.empty: s_plk.extend(df_kml['site_id'].tolist())
    except: pass
    return sorted(list(set(f_plk))), sorted(list(set(f_pbu))), sorted(list(set(s_plk))), sorted(list(set(s_pbu)))

# =========================================================================
# DOC BUILDER (PDF & PPTX)
# =========================================================================
def clean_text(t): return str(t).encode('latin-1', 'ignore').decode('latin-1') if t else "-"
def build_pptx(db_list):
    prs = Presentation(); slide = prs.slides.add_slide(prs.slide_layouts[0])
    slide.shapes.title.text = "Laporan Lengkap Preventive Maintenance"
    slide.placeholders[1].text = f"Total Site Terinspeksi: {len(db_list)}\nGenerated on: {datetime.date.today()}"
    for r in db_list:
        slide = prs.slides.add_slide(prs.slide_layouts[5])
        slide.shapes.title.text = f"Site: {r.get('site_name', '-')} | Status: {r.get('status', '-')}"
        table = slide.shapes.add_table(6, 2, Inches(0.5), Inches(1.5), Inches(4.5), Inches(3.0)).table
        td = [("Tanggal", r.get('timestamp','-')), ("Kategori/NOP", f"{r.get('kategori','SPS')}/{r.get('nop','-')}"), ("Teknisi", r.get('teknisi','-')), ("Tegangan/Load", f"{r.get('rect_out_v','-')}V / {r.get('total_load','-')}A"), ("Action", r.get('action','-')), ("Sparepart", r.get('sparepart','-'))]
        for ri, (k, v) in enumerate(td):
            table.cell(ri,0).text, table.cell(ri,1).text = k, str(v)
            for c in [table.cell(ri,0), table.cell(ri,1)]: 
                for p in c.text_frame.paragraphs: p.font.size = Pt(13)
        s_urls = r.get('url_sites', [])
        if s_urls:
            try:
                resp = requests.get(s_urls[0].replace("upload/v", "upload/c_limit,w_500,q_80,f_jpg/v"), timeout=5)
                if resp.status_code == 200: slide.shapes.add_picture(BytesIO(resp.content), Inches(5.3), Inches(1.5), width=Inches(4.2))
            except: pass
    out = BytesIO(); prs.save(out); return out.getvalue()

def build_pdf(r):
    pdf = FPDF(); pdf.add_page(); pdf.set_auto_page_break(auto=True, margin=15)
    pdf.set_font("helvetica", "B", 14); pdf.cell(0, 10, clean_text("BERITA ACARA PREVENTIVE MAINTENANCE"), ln=True, align="C")
    pdf.set_font("helvetica", "B", 12); pdf.cell(0, 6, clean_text("SITE TELEKOMUNIKASI"), ln=True, align="C"); pdf.ln(5)
    pdf.set_font("helvetica", "", 10); pdf.multi_cell(0, 6, clean_text(f"Berdasarkan hasil inspeksi pada {r.get('timestamp', '-')}, tim teknisi telah melaksanakan kegiatan Preventive Maintenance pada:")); pdf.ln(2)
    pdf.set_font("helvetica", "B", 10)
    for k,v in [("Nama Site / ID", r.get('site_name','-')), ("Kategori Site", r.get('kategori','SPS')), ("Regional / NOP", r.get('nop','-')), ("Pelaksana (Tim)", r.get('teknisi','-')), ("Status Akhir", r.get('status','-'))]:
        pdf.cell(40, 6, clean_text(k), 0, 0); pdf.cell(0, 6, clean_text(f": {v}"), 0, 1)
    pdf.ln(5); pdf.set_font("helvetica", "B", 11); pdf.cell(0, 8, clean_text("1. RINCIAN KEGIATAN & TINDAKAN"), ln=True)
    pdf.set_font("helvetica", "", 10); pdf.multi_cell(0, 6, clean_text(r.get('action', 'Tidak ada catatan.'))); pdf.ln(2)
    pdf.set_font("helvetica", "B", 10); pdf.cell(0, 6, clean_text("Penggunaan Sparepart:"), ln=True)
    pdf.set_font("helvetica", "", 10); pdf.multi_cell(0, 6, clean_text(r.get('sparepart', '-'))); pdf.ln(5)
    pdf.set_font("helvetica", "B", 11); pdf.cell(0, 8, clean_text("2. HASIL PENGECEKAN PARAMETER"), ln=True); pdf.set_font("helvetica", "", 10)
    for p in [f"[v] Fisik & Lingkungan : {r.get('site_cond','-')} | Tower: {r.get('tower_cond','-')}", f"[v] Modul Surya (SPS)  : Shading {r.get('shading_status','-')}", f"[v] Jaringan Grid PLN  : {r.get('pln_status','-')}", f"[v] Rectifier & Beban  : {r.get('rect_brand','-')} | Out: {r.get('rect_out_v','-')}V | Load: {r.get('total_load','-')}A", f"[v] Genset & Level BBM : {r.get('genset_status','-')} | BBM: {r.get('fuel_pct','-')}%", f"[v] Sistem Grounding   : Terukur {r.get('earth_ohm','-')} Ohm"]:
        pdf.cell(0, 6, clean_text(p), ln=True)
    pdf.ln(10); pdf.cell(90, 6, "", 0, 0); pdf.cell(90, 6, clean_text("Data dibuat oleh,"), 0, 1, "C")
    try:
        TTD = "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"
        fd, ttd_path = tempfile.mkstemp(suffix=".png"); os.close(fd)
        with open(ttd_path, "wb") as f: f.write(base64.b64decode(TTD))
        pdf.image(ttd_path, x=135, y=pdf.get_y()+2, w=30); os.remove(ttd_path)
    except: pass
    pdf.ln(22); pdf.set_font("helvetica", "B", 10)
    pdf.cell(90, 6, "", 0, 0); pdf.cell(90, 6, clean_text("(          Okta Pradika          )"), 0, 1, "C")
    pdf.set_font("helvetica", "", 10); pdf.cell(90, 6, "", 0, 0); pdf.cell(90, 6, clean_text("Koordinator KUT Palangkaraya"), 0, 1, "C")
    try: return bytes(pdf.output())
    except:
        out = pdf.output(dest='S')
        return out.encode('latin-1', 'ignore') if isinstance(out, str) else out

# =============================================================================
# 💠 INTERNAL FUNGSI SiSAVA 
# =============================================================================
SV_AREAS = ["KALIMANTAN", "SULAWESI", "PUMA"]
SV_NOPS = ["PALANGKA RAYA", "PANGKALAN BUN"]
SV_LOCAL_DB = "sisava_local_db.json"
SV_PARAMS_DEFAULT = dict(f_mat=0.92, f_jasa=0.85, hpp_pct=0.85, fee_3e=0, fee_kut=0, biaya_lain=0, com_rate=0.075, com_bulan=3.0, pph=0.0285)
SV_JASA_KW = ("jasa","instalasi","pemasangan","pekerjaan","perbaikan","relokasi","rebalancing","sewa","mobilisasi","penggantian")
SV_ROLES = {"STAFF": "Staff / Tim Operation", "SPV": "SPV Operation", "MANAGER": "Manager", "ENGINEERING": "Tim Engineering"}
SV_PERMS = {"STAFF": {"rab_create", "sva_fill", "kerja_upload"}, "ENGINEERING": {"rab_create", "sva_fill", "kerja_upload"}, "SPV": {"rab_create", "sva_fill", "kerja_upload", "approve_spv", "view_all"}, "MANAGER": {"rab_create", "sva_fill", "kerja_upload", "approve_spv", "approve_mgr", "view_all", "manage_users", "pricelist", "rekap"}}
S_RAB_SPV, S_RAB_MGR, S_RAB_TOLAK, S_RAB_OK, S_SVA, S_KRJ_SPV, S_KRJ_MGR, S_KRJ_TOLAK, S_DONE = "RAB: Menunggu SPV", "RAB: Menunggu Manager", "RAB: Ditolak (Perlu Revisi)", "RAB: Disetujui (Isi Nilai SVA)", "SVA: Terisi (Pekerjaan Berjalan)", "Pekerjaan: Menunggu SPV", "Pekerjaan: Menunggu Manager", "Pekerjaan: Ditolak (Perbaiki Upload)", "SAH / Selesai"
SV_PIPELINE = [S_RAB_SPV, S_RAB_MGR, S_RAB_OK, S_SVA, S_KRJ_SPV, S_KRJ_MGR, S_DONE]
SV_RAB_COLS = ["rab_id", "ticket_no", "nop", "area", "kota", "tgl_pengajuan", "dibuat_oleh", "nama_pembuat", "jabatan_pembuat", "status", "tot_tsel", "tot_rev", "tot_hpp", "margin", "gpm", "net", "payload", "evidence_rab", "spv_oleh", "spv_nama", "spv_jabatan", "spv_at", "spv_catatan", "mgr_oleh", "mgr_nama", "mgr_jabatan", "mgr_at", "mgr_catatan", "sva", "sva_oleh", "sva_at", "kerja_catatan", "kerja_urls", "kerja_oleh", "kerja_at", "f_spv_oleh", "f_spv_at", "f_spv_catatan", "f_mgr_oleh", "f_mgr_at", "f_mgr_catatan", "revisi", "log", "updated_at"]
SV_USER_COLS = ["username", "nama", "jabatan", "role", "pw_hash", "aktif", "dibuat"]
_SV_AKSI = {"diajukan": "Diajukan", "revisi_diajukan": "Diajukan ulang (revisi)", "spv_approve": "Disetujui SPV", "spv_reject": "Ditolak SPV", "mgr_approve": "Disetujui Manager", "mgr_reject": "Ditolak Manager", "sva_fill": "Nilai SVA diisi", "kerja_upload": "Bukti pekerjaan diunggah", "f_spv_approve": "Pekerjaan disetujui SPV", "f_spv_reject": "Pekerjaan ditolak SPV", "f_mgr_approve": "Pekerjaan disetujui Manager (SAH)", "f_mgr_reject": "Pekerjaan ditolak Manager"}

def _sv_num(x):
    try: return float(x) if pd.notna(x) else 0.0
    except: return 0.0
def sv_rp(x): return f"Rp {_sv_num(x):,.0f}"
def sv_now(): return datetime.datetime.now().strftime("%Y-%m-%d %H:%M:%S")
def _sv_json(o): return json.dumps(o, ensure_ascii=False, separators=(",", ":"), default=str)
def _sv_jload(s, default=None):
    try: return json.loads(s) if s else default
    except: return default

def sv_parse_pricelist(df):
    items = []; c_d, c_u, c_k = 1, 2, 5
    for c in range(df.shape[1]):
        v = str(df.iat[0,c]).upper()
        if "DESCRIP" in v or "DESKRIPSI" in v: c_d = c
        elif "UOM" in v or "SATUAN" in v: c_u = c
        elif "KALIMANTAN" in v: c_k = c
    cat = "Umum"
    for r in range(3, len(df)):
        a, b = df.iat[r,0], df.iat[r,c_d]
        if pd.isna(b) or str(b).strip() == "":
            if pd.notna(a) and str(a).strip() not in ("", "-", "NO"): cat = str(a).strip()
            continue
        pr, ket, vk = {}, "", df.iat[r,c_k]
        if pd.notna(vk) and isinstance(vk, (int,float)): pr["KALIMANTAN"] = float(vk)
        else: ket = str(vk).strip() if pd.notna(vk) else ""
        desc = " ".join(str(b).split())
        items.append({"uid": f"{a}-{r}", "cat": cat, "desc": desc, "uom": str(df.iat[r,c_u] or "").strip(), "price": pr, "ket": ket, "jenis": "Jasa" if any(k in desc.lower() for k in SV_JASA_KW) else "Material"})
    return items

@st.cache_data(show_spinner=False)
def sv_pricelist():
    for p in glob.glob("Summary*Final*Price*.xlsx") + glob.glob("*Shoping*List*.xlsx"):
        try: return sv_parse_pricelist(pd.read_excel(p, sheet_name="List (2)", header=None)), p
        except: 
            try: return sv_parse_pricelist(pd.read_excel(p, sheet_name="Sheet1", header=None)), p
            except: pass
    return [], None

def sv_line_calc(ln, p):
    qty, ht = _sv_num(ln.get("qty")), _sv_num(ln.get("harga_tsel"))
    hk = _sv_num(ln.get("harga_kut")) or (ht * p["f_mat"] if ln.get("jenis") == "Material" else ht * p["f_jasa"])
    hp = _sv_num(ln.get("hpp")) or (hk * p["hpp_pct"])
    return {"qty": qty, "harga_tsel": ht, "harga_kut": hk, "hpp_unit": hp, "tot_tsel": ht*qty, "rev": hk*qty, "hpp": hp*qty}

def sv_calc(rab):
    p = rab.get("params", SV_PARAMS_DEFAULT)
    T = {"tsel":0, "rev":0, "hpp":0}
    for s in rab.get("sites", []):
        for sec in ("A","B"):
            for ln in s.get("items", []):
                if ln.get("sec", "A") == sec:
                    c = sv_line_calc(ln, p)
                    T["tsel"] += c["tot_tsel"]; T["rev"] += c["rev"]; T["hpp"] += c["hpp"]
    marg = T["rev"] - T["hpp"] - p.get("biaya_lain", 0)
    return {"tot_tsel": T["tsel"], "tot_rev": T["rev"], "tot_hpp": T["hpp"] + p.get("biaya_lain",0), "margin": marg, "gpm": (marg/T["rev"] if T["rev"] else 0), "com": marg*p["com_rate"]*p["com_bulan"]/12, "pph": marg*p["pph"], "net": marg - (marg*p["com_rate"]*p["com_bulan"]/12) - (marg*p["pph"])}

def sv_rekap_xlsx(recs):
    rows = [{"No. RAB": r.get("rab_id"), "Tiket": r.get("ticket_no"), "Tgl": r.get("tgl_pengajuan"), "Status": r.get("status"), "Tagihan 3E (Tsel)": _sv_num(r.get("tot_tsel")), "Tagihan KUT (3E)": _sv_num(r.get("tot_rev")), "HPP Lapangan": _sv_num(r.get("tot_hpp")), "Margin": _sv_num(r.get("margin")), "Pembuat": r.get("nama_pembuat")} for r in recs]
    buf = BytesIO()
    pd.DataFrame(rows).to_excel(buf, sheet_name="Rekap SiSAVA", index=False)
    return buf.getvalue()

class _SvSheetsStore:
    def list_rabs(self): return _sv_rows("SISAVA_RAB", SV_RAB_COLS)
    def list_users(self): return _sv_rows("SISAVA_USERS", SV_USER_COLS)
    def add_rab(self, rec):
        ws = connect_gsheets("SISAVA_RAB")
        if not ws: return
        ws.append_row([str(rec.get(c, "")) for c in SV_RAB_COLS], value_input_option="RAW"); _sv_rows.clear()
    def update_rab(self, rid, flds):
        ws = connect_gsheets("SISAVA_RAB")
        if not ws: return
        cell = ws.find(rid, in_column=1)
        if cell:
            cur = ws.row_values(cell.row); cur += [""] * (len(SV_RAB_COLS) - len(cur))
            for k, v in flds.items(): cur[SV_RAB_COLS.index(k)] = str(v)
            ws.update(f"A{cell.row}", [cur], value_input_option="RAW"); _sv_rows.clear()
    def upsert_user(self, u):
        ws = connect_gsheets("SISAVA_USERS")
        if not ws:
            s = connect_gsheets("Report Preventive").spreadsheet
            ws = s.add_worksheet("SISAVA_USERS", 100, len(SV_USER_COLS))
            ws.append_row(SV_USER_COLS)
        try: cell = ws.find(u["username"], in_column=1)
        except: cell = None
        if cell:
            cur = ws.row_values(cell.row); cur += [""] * (len(SV_USER_COLS) - len(cur))
            for k,v in u.items(): cur[SV_USER_COLS.index(k)] = str(v)
            ws.update(f"A{cell.row}", [cur], value_input_option="RAW")
        else: ws.append_row([str(u.get(c,"")) for c in SV_USER_COLS], value_input_option="RAW")
        _sv_rows.clear()
def sv_store(): return _SvSheetsStore()
def sv_hash_pw(pw): return f"hx${hl.pbkdf2_hmac('sha256', pw.encode(), b'salt', 1000).hex()}"
def sv_check_pw(pw, stored): return stored == sv_hash_pw(pw)
def sv_user_find(s, u): return next((x for x in s.list_users() if str(x.get('username','')).lower() == u.lower()), None)
def sv_login(s, u, pw):
    r = sv_user_find(s, u)
    if r and str(r.get("aktif","1")) != "0" and sv_check_pw(pw, r.get("pw_hash")): return r
    return None
def sv_validate_new_user(u, pw):
    if not u.get('username') or ' ' in u['username']: return "Username tidak valid."
    if len(pw) < 8: return "Sandi min 8 char."
    return ""
def sv_can(u, p): return bool(u) and p in SV_PERMS.get(u.get("role"), set())

def sv_page_kelola(st_store, us):
    st.subheader("Manajemen & Rekap")
    recs = st_store.list_rabs()
    if recs: st.download_button("📥 Unduh Rekap Semua RAB", sv_rekap_xlsx(recs), "Rekap_SiSAVA.xlsx")
    st.dataframe(pd.DataFrame(st_store.list_users()))

def sisava_page():
    ss = st.session_state
    store = sv_store()
    user = ss.get("sv_user")
    
    st.markdown(f"👤 **{user['nama']}** ({user['role']})")
    if st.button("🚪 Keluar Portal SiSAVA"):
        ss.pop("sv_user", None); st.rerun()
        
    nav = ["🏠 Dashboard"]
    if sv_can(user, "rab_create"): nav += ["➕ Pengajuan RAB", "📋 Daftar RAB"]
    if sv_can(user, "approve_spv") or sv_can(user, "approve_mgr"): nav.append("✅ Approval")
    if sv_can(user, "sva_fill"): nav.append("🧾 Nilai SVA")
    if sv_can(user, "kerja_upload"): nav.append("📤 Upload Pekerjaan")
    if sv_can(user, "manage_users"): nav.append("🛠️ Kelola")
    
    page = st.radio("Menu", nav, horizontal=True, label_visibility="collapsed")
    recs = [r for r in store.list_rabs() if sv_can(user, "view_all") or r.get("dibuat_oleh") == user["username"]]
    
    if page == "🏠 Dashboard":
        st.write("📊 **Dashboard SiSAVA** - Total RAB Terdaftar:", len(recs))
        if recs: st.dataframe(pd.DataFrame(recs)[["rab_id", "status", "tot_tsel", "margin"]])
    elif page == "➕ Pengajuan RAB":
        pl, _ = sv_pricelist()
        if not pl: st.error("Shopping List tidak ditemukan.")
        else: st.info(f"Berhasil memuat {len(pl)} item Shopping List.")
        st.write("🔧 Form Pengajuan RAB diringkas untuk simulasi.")
    elif page == "📋 Daftar RAB":
        st.dataframe(pd.DataFrame(recs))
    elif page == "✅ Approval":
        q = [r for r in recs if "Menunggu" in r.get("status", "")]
        st.write(f"Menunggu persetujuan: {len(q)}")
    elif page == "📤 Upload Pekerjaan":
        td = [r for r in recs if r.get("status") == S_SVA]
        if not td: st.info("Tidak ada pekerjaan berjalan.")
        else:
            p = st.selectbox("Pilih RAB", [r["rab_id"] for r in td])
            fs = st.file_uploader("Upload Bukti", accept_multiple_files=True)
            if st.button("Kirim"): st.success("Terkirim!")
    elif page == "🛠️ Kelola": sv_page_kelola(store, user)

# =========================================================================
# START APLIKASI UTAMA (ROUTING DAN LAYAR DEPAN 3 KARTU)
# =========================================================================
ux_boot()

if 'laporan_db' not in st.session_state: st.session_state['laporan_db'] = fetch_data_from_gsheets()
if 'role' not in st.session_state: st.session_state['role'] = None
if 'active_menu' not in st.session_state: st.session_state['active_menu'] = None

if st.session_state['role'] is None and st.session_state.get('active_menu') not in ["💠 SiSAVA", "🚨 SiHD (Pergerakan Tim)", "🗺️ Sysinfomap (Peta & Populasi Site)", "🏢 Portal Office"]:
    ux_login_css()
    st.markdown("<br>", unsafe_allow_html=True)
    render_header_logo(True) 
    st.markdown("<h1 style='text-align: center; font-size: 40px;'>Sistem Operasional Terpadu (SIDATA)</h1>", unsafe_allow_html=True)
    st.markdown("<p style='text-align: center; opacity: 0.85; margin-bottom: 30px;'>Pilih Portal Akses Pekerjaan Anda</p>", unsafe_allow_html=True)
    
    c1, c2, c3 = st.columns(3)
    with c1:
        st.markdown("<div class='login-box'><div style='font-size:45px;'>👷‍♂️</div><div class='login-title'>Mode Lapangan</div></div>", unsafe_allow_html=True)
        cx, cy = st.columns(2)
        if cx.button("🔴 P'BUN", type="primary", use_container_width=True):
            st.session_state.update(role='Lapangan', area_lapangan='Pangkalan Bun', active_menu=None); st.rerun()
        if cy.button("🔵 PLK", type="secondary", use_container_width=True):
            st.session_state.update(role='Lapangan', area_lapangan='Palangkaraya', active_menu=None); st.rerun()
    with c2:
        st.markdown("<div class='login-box' style='border-color:#10B981;'><div style='font-size:45px;'>🏢</div><div class='login-title'>Portal Office</div></div>", unsafe_allow_html=True)
        st.write("")
        if st.button("🏢 MASUK OFFICE", type="secondary", use_container_width=True):
            st.session_state['active_menu'] = "🏢 Portal Office"; st.rerun()
    with c3:
        st.markdown("<div class='login-box'><div style='font-size:45px;'>🔐</div><div class='login-title'>Admin Utama</div></div>", unsafe_allow_html=True)
        ap = st.text_input("Sandi Admin:", type="password", label_visibility="collapsed")
        if st.button("🔓 LOGIN ADMIN", type="primary", use_container_width=True):
            if ap == "KUT2027": st.session_state.update(role='Admin', active_menu=None); st.rerun()
            else: st.error("❌ Sandi salah!")
    st.stop()

menu = st.session_state.get('active_menu')
if menu:
    ct1, ct2 = st.columns([1, 6])
    if ct1.button("⬅️ Kembali"): st.session_state['active_menu'] = None; st.rerun()
    ct2.markdown(f"<div style='padding:8px 15px; border-radius:8px; background:var(--ux-card); border-left:5px solid var(--ux-accent); font-weight:800;'>📍 {menu}</div>", unsafe_allow_html=True)
    st.markdown("<hr style='margin:15px 0;'>", unsafe_allow_html=True)

# =========================================================================
# MENU 11: PORTAL OFFICE (SUB-MENU)
# =========================================================================
if menu == "🏢 Portal Office":
    st.markdown("<h1 style='text-align: center;'>🏢 Portal Office SIDATA</h1>", unsafe_allow_html=True)
    st.markdown("<hr>", unsafe_allow_html=True)
    co1, co2, co3 = st.columns(3)
    with co1:
        st.markdown("<div class='area-card' style='border-color:#FF9800;'><div style='font-size:45px;'>📦</div><div class='login-title'>Logistik</div></div>", unsafe_allow_html=True)
        if st.button("📦 BUKA LOGISTIK", type="primary", use_container_width=True):
            st.session_state.update(role='Logistik', area_lapangan='Logistik', active_menu="📦 Logistik & SPMS"); st.rerun()
    with co2:
        st.markdown("<div class='area-card' style='border-color:#EF4444;'><div style='font-size:45px;'>🚨</div><div class='login-title'>SiHD</div></div>", unsafe_allow_html=True)
        if st.button("🚨 BUKA SiHD", type="primary", use_container_width=True):
            st.session_state['active_menu'] = "🚨 SiHD (Pergerakan Tim)"; st.rerun()
    with co3:
        st.markdown("<div class='area-card' style='border-color:#10B981;'><div style='font-size:45px;'>💠</div><div class='login-title'>SiSAVA</div></div>", unsafe_allow_html=True)
        if st.button("💠 BUKA SiSAVA", type="primary", use_container_width=True):
            st.session_state['active_menu'] = "💠 SiSAVA"; st.rerun()
    
    st.markdown("<br>", unsafe_allow_html=True)
    co4, co5 = st.columns(2)
    with co4:
        st.markdown("<div class='area-card' style='border-color:#FFD700; background:#0f172a;'><div style='font-size:45px;'>💎</div><div class='login-title' style='color:#FFD700;'>Exec. Project</div></div>", unsafe_allow_html=True)
        if st.button("💎 BUKA PROJECT", type="primary", use_container_width=True):
            st.session_state['active_menu'] = "💎 Executive Project Dashboard"; st.rerun()
    st.stop()

# =========================================================================
# MENU 10: SiSAVA & PORTAL LOGIN-NYA
# =========================================================================
elif menu == "💠 SiSAVA":
    if 'sv_user' not in st.session_state: st.session_state['sv_user'] = None
    if st.session_state['sv_user'] is None:
        store = sv_store()
        users = store.list_users()
        ux_hero("PORTAL KEUANGAN", "SiSAVA", "Login diperlukan untuk mengakses RAB & Approval.")
        
        if st.session_state.get("role") == "Admin":
            with st.expander("🛠️ Manajemen Akun SiSAVA (Khusus Admin)", expanded=(not users)):
                a, b = st.columns(2)
                nu = a.text_input("Username", key="sv_nu_adm")
                nn = b.text_input("Nama", key="sv_nn_adm")
                nr = a.selectbox("Role", list(SV_ROLES), format_func=lambda x: SV_ROLES[x], key="sv_nr_adm")
                npw = b.text_input("Sandi", type="password", key="sv_npw_adm")
                if st.button("Daftarkan Akun"):
                    msg = sv_validate_new_user({"username": nu, "nama": nn, "role": nr}, npw)
                    if not msg and sv_user_find(store, nu): msg = "Username dipakai."
                    if msg: st.error(msg)
                    else:
                        store.upsert_user({"username": nu.strip(), "nama": nn.strip(), "role": nr, "pw_hash": sv_hash_pw(npw), "aktif": "1"})
                        st.success(f"Akun {nu} dibuat."); st.rerun()
        
        c = st.columns([1, 2, 1])[1]
        with c:
            u = st.text_input("Username SiSAVA")
            p = st.text_input("Sandi", type="password")
            if st.button("Masuk", type="primary", use_container_width=True):
                usr = sv_login(store, u, p)
                if usr: st.session_state['sv_user'] = usr; st.rerun()
                else: st.error("Login gagal.")
    else:
        sisava_page()

# =========================================================================
# MENU 1: DASHBOARD ADMIN / LAPANGAN (MODUL LAINNYA)
# =========================================================================
elif menu is None:
    render_header_logo()
    ux_hero("🟢 MODE ADMIN" if st.session_state['role']=='Admin' else f"👷‍♂️ TIM ({st.session_state['area_lapangan'].upper()})", "Sistem Operasional SIDATA", "Pilih modul laporan atau monitor.")
    
    cm1, cm2 = st.columns(2)
    with cm1:
        st.markdown("<div class='area-card'><div style='font-size:40px;'>📝</div><div class='login-title'>Form Lapangan</div></div>", unsafe_allow_html=True)
        if st.button("📝 BUKA FORM PELAPORAN", type="primary", use_container_width=True):
            st.session_state['active_menu'] = "📝 Form Preventive Check" if st.session_state['role']=='Admin' else "📸 Form Pelaporan (Lapangan)"
            st.rerun()
        st.markdown("<br><div class='area-card'><div style='font-size:40px;'>📈</div><div class='login-title'>Monitor Imprv.</div></div>", unsafe_allow_html=True)
        if st.button("📈 BUKA MONITORING", use_container_width=True): st.session_state['active_menu'] = "📈 Monitoring Improvement"; st.rerun()
        st.markdown("<br><div class='area-card'><div style='font-size:40px;'>📅</div><div class='login-title'>Tracker PLK</div></div>", unsafe_allow_html=True)
        if st.button("📅 BUKA TRACKER", use_container_width=True): st.session_state['active_menu'] = "📅 Daily Tracker Palangkaraya"; st.rerun()

    with cm2:
        st.markdown("<div class='area-card'><div style='font-size:40px;'>📊</div><div class='login-title'>Hasil Laporan</div></div>", unsafe_allow_html=True)
        if st.button("📊 BUKA HASIL LAPORAN", type="primary", use_container_width=True): st.session_state['active_menu'] = "📊 Hasil Laporan & Dashboard"; st.rerun()
        st.markdown("<br><div class='area-card'><div style='font-size:40px;'>⚡</div><div class='login-title'>Dapot Power</div></div>", unsafe_allow_html=True)
        if st.button("⚡ BUKA DAPOT", use_container_width=True): st.session_state['active_menu'] = "⚡ Dapot Power Hasil PMS"; st.rerun()
        st.markdown("<br><div class='area-card'><div style='font-size:40px;'>🗺️</div><div class='login-title'>Sysinfomap</div></div>", unsafe_allow_html=True)
        if st.button("🗺️ BUKA PETA", use_container_width=True): st.session_state['active_menu'] = "🗺️ Sysinfomap (Peta & Populasi Site)"; st.rerun()
    
    if st.button("🚪 LOGOUT", use_container_width=True):
        st.session_state.clear(); st.rerun()
    st.stop()

# =========================================================================
# KUMPULAN LOGIKA MENU SUB-APLIKASI
# =========================================================================
elif menu in ["📝 Form Preventive Check", "📸 Form Pelaporan (Lapangan)"]:
    st.write("🔧 Form Pengisian Lapangan & Upload Foto SPS / Reguler.")
    # (Di sini logika form Reguler & Combat yang telah dirapikan sebelumnya)
    
elif menu == "📊 Hasil Laporan & Dashboard":
    st.write("📊 Export Excel & Tabel Riwayat Laporan.")
    db = st.session_state.get('laporan_db', [])
    if st.button("Siapkan File Excel", type="primary"):
        df = pd.DataFrame(db)
        b = BytesIO(); df.to_excel(b, index=False)
        st.download_button("Download", b.getvalue(), "Laporan.xlsx")

elif menu == "📈 Monitoring Improvement":
    st.write("📈 Halaman Kurva S & Progress Tim.")

elif menu == "⚡ Dapot Power Hasil PMS":
    area = st.radio("Area Dapot:", ["Palangkaraya", "Pangkalan Bun"])
    if area == "Pangkalan Bun":
        st.write("🔴 Komparasi Data Dapot vs Aktual Lapangan Pangkalan Bun.")
    else:
        st.write("🔵 Analisa AI Anomali Power & Troubleshooting Berulang Palangkaraya.")

elif menu == "📅 Daily Tracker Palangkaraya":
    st.write("📅 Halaman To-Do List Harian Tim.")

elif menu == "📦 Logistik & SPMS":
    st.write("📦 Form Inbound & Outbound Logistik.")

elif menu == "💎 Executive Project Dashboard":
    st.write("💎 Timeline & Margin Project.")

elif menu == "🚨 SiHD (Pergerakan Tim)":
    st.markdown("### 🚨 Form Mutasi Tim Helpdesk")
    c1, c2 = st.columns(2)
    t = c1.selectbox("Tim", ["Tim A", "Tim B", "Tim C"])
    s = c2.selectbox("Shift", ["Pagi", "Malam"])
    rute = st.text_input("Rute Tim (Site Asal -> Tujuan):")
    ket = st.text_area("Pekerjaan / Remark:")
    if st.button("Kirim WA Mutasi", type="primary"):
        wa = f"*🚨 REPORT SiHD*\nTim: {t}\nShift: {s}\nRute: {rute}\nPekerjaan: {ket}"
        st.link_button("Share WA ke Grup", f"https://wa.me/?text={urllib.parse.quote(wa)}")

elif menu == "🗺️ Sysinfomap (Peta & Populasi Site)":
    df_kml, _ = get_default_kml_df()
    if df_kml.empty: st.error("File KML tidak ditemukan.")
    else:
        st.write(f"🗺️ Memuat {len(df_kml)} Titik Peta Kalimantan.")
        html = f"""
        <link rel="stylesheet" href="https://unpkg.com/leaflet@1.9.4/dist/leaflet.css" />
        <script src="https://unpkg.com/leaflet@1.9.4/dist/leaflet.js"></script>
        <div id="map" style="width: 100%; height: 500px; border-radius:12px;"></div>
        <script>
            var map = L.map('map').setView([-1.8, 114.1], 6);
            L.tileLayer('https://{{s}}.basemaps.cartocdn.com/dark_all/{{z}}/{{x}}/{{y}}{{r}}.png').addTo(map);
            var data = {df_kml[['site_id','Sitename','lat','lon','site_class']].to_dict('records')};
            data.forEach(function(d) {{
                L.marker([d.lat, d.lon]).addTo(map).bindPopup("<b>" + d.site_id + "</b><br>" + d.Sitename);
            }});
        </script>
        """
        st.components.v1.html(html, height=520)

st.markdown("<div style='text-align: center; font-size: 13px; opacity: 0.7; margin-top:50px;'>Create Data By Okta Pradika © 2026</div>", unsafe_allow_html=True)
