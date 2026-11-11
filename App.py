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
# SETUP HALAMAN & FUNGSI LOGO PRESISI (.WEBP)
# -------------------------------------------------------------------------
st.set_page_config(page_title="RPO - PLK & PBU", page_icon="⚡", layout="wide")

LOGO_UTAMA = "koperasi-jasa-konstruksi-tower-event-organizer-network-monitoring-telekomunikasi-kisel-group-logo-kut (2).webp"

def get_base64_of_bin_file(bin_file):
    # Hasil di-cache (kunci = nama file + waktu ubah) supaya logo tidak di-encode ulang di setiap klik
    try:
        stt = os.stat(bin_file)
        return _b64_cached(bin_file, (stt.st_mtime_ns, stt.st_size))
    except Exception:
        return None

def render_header_logo(is_login=False):
    logo_candidates = [
        (LOGO_UTAMA, "image/webp"),
        ("logo.png", "image/png"),
        ("logo.webp", "image/webp")
    ]
    for filename, mime in logo_candidates:
        if os.path.exists(filename):
            bg_base64 = get_base64_of_bin_file(filename)
            if bg_base64:
                max_h = "135px" if is_login else "100px"
                pad = "16px 36px" if is_login else "12px 28px"
                st.markdown(
                    f"""
                    <div style="display: flex; justify-content: center; align-items: center; margin-bottom: 20px;">
                        <div style="background: linear-gradient(145deg, #ffffff, #f3f6f9); padding: {pad}; border-radius: 20px; box-shadow: 0 10px 28px rgba(0,0,0,0.22); border: 2px solid var(--primary-color); display: inline-flex; align-items: center; justify-content: center;">
                            <img src="data:{mime};base64,{bg_base64}" style="max-height: {max_h}; width: auto; max-width: 100%; object-fit: contain; display: block;">
                        </div>
                    </div>
                    """, 
                    unsafe_allow_html=True
                )
                return

# =============================================================================
# 💎 PREMIUM UI KIT — tema mewah • tampilan Web/HP • grafik ringan (tanpa JavaScript)
#   Cara kerja singkat:
#   - Semua gaya dibuat dengan CSS (transform/opacity saja) -> ringan di HP.
#   - Grafik digambar sebagai SVG/HTML murni (bukan Vega/Altair) -> cepat & tanpa lag.
#   - Pilihan tampilan 🖥️ Web | 📱 HP ada di pojok kanan atas; link HP: ?view=hp
# =============================================================================
import html as _html
import math as _math
import numpy as _np
from xml.sax.saxutils import escape as _xesc

UI_THEME = "auto"   # "auto" = ikuti tema Streamlit/perangkat | "light" | "dark"

_PAL_LIGHT = dict(bg="#F3F6FC", card="#FFFFFF", card2="#F7F9FE", line="#E1E7F3", ink="#0B1730", muted="#5A6A86",
                  accent="#2563EB", gold="#B8891A", gold2="#E2B93B", ok="#059669", warn="#D97706", bad="#E11D48",
                  btn1="#1D4ED8", btn2="#4F46E5", glow="rgba(37,99,235,.10)", gglow="rgba(184,137,26,.08)",
                  shadow="0 1px 2px rgba(11,23,48,.06),0 10px 28px rgba(11,23,48,.07)")
_PAL_DARK = dict(bg="#0A1020", card="#111A30", card2="#16213C", line="#243253", ink="#EAF0FF", muted="#93A4C7",
                 accent="#5B9BFF", gold="#E8C547", gold2="#F6DE8D", ok="#34D399", warn="#FBBF24", bad="#FB7185",
                 btn1="#1D4ED8", btn2="#4F46E5", glow="rgba(91,155,255,.12)", gglow="rgba(232,197,71,.06)",
                 shadow="0 1px 2px rgba(0,0,0,.45),0 12px 32px rgba(0,0,0,.38)")

def _theme_mode():
    """Tentukan palet. Kalau config.toml memaksa tema, ikuti; kalau tidak, ikuti mode gelap/terang perangkat."""
    if UI_THEME in ("light", "dark"):
        return UI_THEME
    try:
        base = st.get_option("theme.base")
        if base in ("light", "dark"):
            return base
    except Exception:
        pass
    return "auto"

def _tok(d):
    return ";".join("--ux-%s:%s" % (k, v) for k, v in d.items())

def _root_css(mode):
    if mode == "light": return ":root{" + _tok(_PAL_LIGHT) + "}"
    if mode == "dark":  return ":root{" + _tok(_PAL_DARK) + "}"
    return ":root{" + _tok(_PAL_LIGHT) + "}@media (prefers-color-scheme:dark){:root{" + _tok(_PAL_DARK) + "}}"

def _svg_style(mode):
    def c(d):
        return ".tx{fill:%s}.ik{fill:%s}.gr{stroke:%s}.bs{stroke:%s}" % (d["muted"], d["ink"], d["line"], d["card"])
    if mode == "light": return c(_PAL_LIGHT)
    if mode == "dark":  return c(_PAL_DARK)
    return c(_PAL_LIGHT) + "@media (prefers-color-scheme:dark){" + c(_PAL_DARK) + "}"

# ----------------------------------------------------------------------------
# CSS UTAMA (dipakai Web & HP)
# ----------------------------------------------------------------------------
_UX_CSS = """
<style>
@import url('https://fonts.googleapis.com/css2?family=Plus+Jakarta+Sans:wght@400;500;600;700;800&display=swap');
__ROOT__
html,body,.stApp{font-family:'Plus Jakarta Sans',system-ui,-apple-system,'Segoe UI',Roboto,sans-serif}
.stApp{background:radial-gradient(900px 380px at 10% -8%,var(--ux-glow),transparent 70%),radial-gradient(760px 340px at 94% -6%,var(--ux-gglow),transparent 70%),var(--ux-bg) !important}

/* --- buang elemen bawaan Streamlit yang mengganggu --- */
#MainMenu,footer,header[data-testid="stHeader"],div[data-testid="stToolbar"],div[data-testid="stDecoration"],
[data-testid="stSidebar"],[data-testid="collapsedControl"],[data-testid="stSidebarCollapsedControl"],[data-testid="stExpandSidebarButton"]{display:none !important}

/* --- kanvas --- */
[data-testid="stMainBlockContainer"],.stMainBlockContainer,.main .block-container{max-width:1240px !important;margin:0 auto !important;padding:.5rem 1.25rem 4rem !important}

/* --- indikator loading tipis (hanya saat skrip berjalan) --- */
.stApp[data-test-script-state="running"]::after{content:"";position:fixed;left:0;top:0;height:3px;width:100%;z-index:99999;
  background:linear-gradient(90deg,var(--ux-gold),var(--ux-accent),var(--ux-gold2));transform-origin:0 50%;animation:ux-run 1.1s ease-in-out infinite}
@keyframes ux-run{0%{transform:scaleX(.05)}60%{transform:scaleX(.8)}100%{transform:scaleX(1);opacity:.2}}
@keyframes ux-in{from{opacity:0;transform:translateY(8px)}to{opacity:1;transform:none}}

/* --- tipografi --- */
h1{font-weight:800 !important;letter-spacing:-.02em}
h1::after{content:"";display:block;width:54px;height:3px;border-radius:3px;margin:10px 0 0;background:linear-gradient(90deg,var(--ux-gold),var(--ux-gold2))}
h1[style*="center"]::after{margin-left:auto;margin-right:auto}
h2,h3,h4{font-weight:700 !important;letter-spacing:-.01em}
hr{border:0 !important;height:1px !important;background:linear-gradient(90deg,transparent,var(--ux-line) 15%,var(--ux-line) 85%,transparent) !important;margin:1.4rem 0 !important}

/* --- tombol --- */
div[data-testid="stButton"] button,div[data-testid="stFormSubmitButton"] button,div[data-testid="stDownloadButton"] button,a[data-testid="stLinkButton"]{
  border-radius:12px !important;font-weight:700 !important;letter-spacing:.01em;transition:transform .18s ease,box-shadow .18s ease,border-color .18s ease}
button[kind="primary"],button[data-testid="stBaseButton-primary"],button[data-testid="stBaseButton-primaryFormSubmit"]{
  background:linear-gradient(135deg,var(--ux-btn1),var(--ux-btn2)) !important;color:#fff !important;border:1px solid rgba(255,255,255,.16) !important;
  box-shadow:0 8px 18px -6px rgba(37,99,235,.55),inset 0 1px 0 rgba(255,255,255,.2) !important}
button[kind="secondary"],button[data-testid="stBaseButton-secondary"],button[data-testid="stBaseButton-secondaryFormSubmit"]{
  background:var(--ux-card) !important;color:var(--ux-ink) !important;border:1.5px solid var(--ux-line) !important;box-shadow:0 1px 2px rgba(11,23,48,.05) !important}
div[data-testid="stButton"] button p,div[data-testid="stFormSubmitButton"] button p{color:inherit !important}
button[kind="primary"]:hover,button[data-testid="stBaseButton-primary"]:hover{transform:translateY(-2px);box-shadow:0 12px 22px -6px rgba(37,99,235,.6),inset 0 1px 0 rgba(255,255,255,.22) !important}
button[kind="secondary"]:hover,button[data-testid="stBaseButton-secondary"]:hover{transform:translateY(-1px);border-color:var(--ux-accent) !important;box-shadow:0 8px 18px -8px rgba(37,99,235,.35) !important}
button:active{transform:scale(.985) !important}
button:disabled{opacity:.55 !important;filter:saturate(.6)}

/* --- input --- */
div[data-baseweb="input"],div[data-baseweb="base-input"],div[data-baseweb="textarea"],div[data-baseweb="select"] > div{
  background:var(--ux-card2) !important;border:1.5px solid var(--ux-line) !important;border-radius:12px !important;transition:border-color .15s ease,box-shadow .15s ease}
div[data-baseweb="input"]:focus-within,div[data-baseweb="base-input"]:focus-within,div[data-baseweb="textarea"]:focus-within,div[data-baseweb="select"] > div:focus-within{
  border-color:var(--ux-accent) !important;box-shadow:0 0 0 3px var(--ux-glow) !important}
div[data-testid="stWidgetLabel"] p,div[data-testid="stWidgetLabel"] label{font-weight:700 !important;font-size:.9rem !important}
span[data-baseweb="tag"]{border-radius:8px !important}

/* --- tab --- */
.stTabs [data-baseweb="tab-list"]{gap:4px;border-bottom:1px solid var(--ux-line)}
.stTabs [data-baseweb="tab"]{font-weight:700;border-radius:10px 10px 0 0;padding:10px 16px}
.stTabs [data-baseweb="tab"][aria-selected="true"]{color:var(--ux-accent) !important;background:var(--ux-card)}
.stTabs [data-baseweb="tab-highlight"]{height:3px !important;border-radius:3px;background:linear-gradient(90deg,var(--ux-gold),var(--ux-accent)) !important}

/* --- expander, container border, alert, tabel --- */
div[data-testid="stExpander"] details{border:1px solid var(--ux-line) !important;border-left:4px solid var(--ux-gold) !important;border-radius:14px !important;background:var(--ux-card) !important;box-shadow:var(--ux-shadow);margin-bottom:12px}
div[data-testid="stExpander"] summary{font-weight:700 !important;padding:12px 16px}
div[data-testid="stVerticalBlockBorderWrapper"]{border-radius:16px}
div[data-testid="stAlert"]{border-radius:14px !important}
div[data-testid="stDataFrame"],div[data-testid="stDataEditor"]{border-radius:14px;overflow:hidden;border:1px solid var(--ux-line)}
section[data-testid="stFileUploaderDropzone"]{border-radius:14px !important;border:1.5px dashed var(--ux-line) !important;background:var(--ux-card2) !important}

/* --- KPI (st.metric dibuat jadi kartu premium) --- */
div[data-testid="stMetric"]{position:relative;overflow:hidden;background:var(--ux-card);border:1px solid var(--ux-line);border-radius:16px;padding:14px 16px 12px 20px;box-shadow:var(--ux-shadow);animation:ux-in .35s ease both}
div[data-testid="stMetric"]::before{content:"";position:absolute;left:0;top:0;bottom:0;width:5px;background:linear-gradient(180deg,var(--ux-gold2),var(--ux-gold))}
div[data-testid="stMetricLabel"] p{font-size:.74rem !important;font-weight:700 !important;text-transform:uppercase;letter-spacing:.05em;color:var(--ux-muted) !important}
div[data-testid="stMetric"]{container-type:inline-size}
div[data-testid="stMetricValue"]{font-weight:800 !important;font-size:1.3rem !important;font-size:clamp(1.05rem,7.5cqw,1.7rem) !important;letter-spacing:-.01em;line-height:1.25}
div[data-testid="stMetricValue"] > div{white-space:normal !important;overflow:visible !important;text-overflow:clip !important;word-break:normal !important;overflow-wrap:normal !important}

/* --- kartu menu / login (kelas lama dipertahankan) --- */
.area-card,.login-box{position:relative;background:var(--ux-card);border:1px solid var(--ux-line);border-top:4px solid var(--ux-accent);
  border-radius:20px;padding:28px 22px;text-align:center;box-shadow:var(--ux-shadow);margin-bottom:12px;
  transition:transform .22s ease,box-shadow .22s ease;animation:ux-in .4s ease both}
.area-card:hover,.login-box:hover{transform:translateY(-4px);box-shadow:0 18px 40px -12px rgba(11,23,48,.28)}
.login-title{color:var(--ux-ink);font-size:1.25rem;font-weight:800;margin-bottom:10px;letter-spacing:-.01em}
.login-desc{color:var(--ux-muted);font-size:.9rem;line-height:1.6;margin-bottom:16px}
.area-badge{display:inline-block;padding:5px 14px;border-radius:30px;font-weight:800;font-size:.72rem;letter-spacing:.08em;margin-bottom:12px;
  background:rgba(37,99,235,.12);color:#2563EB;border:1px solid rgba(37,99,235,.55)}
.footer-okta{text-align:center;padding:24px;margin-top:48px;color:var(--ux-muted);font-size:.9rem;border-top:1px solid var(--ux-line)}
.footer-okta span{color:var(--ux-gold);font-weight:800;letter-spacing:1px}

/* --- HERO --- */
.ux-hero{position:relative;overflow:hidden;border-radius:24px;padding:30px 28px 26px;margin:6px 0 22px;color:#EAF0FF;text-align:center;
  background:linear-gradient(135deg,#070C18 0%,#13245A 55%,#0B1226 100%);border:1px solid rgba(226,185,59,.35);
  box-shadow:0 22px 50px -18px rgba(7,12,24,.65);animation:ux-in .45s ease both}
.ux-hero::before{content:"";position:absolute;inset:0;pointer-events:none;background:radial-gradient(520px 220px at 85% -10%,rgba(226,185,59,.20),transparent 70%),radial-gradient(480px 240px at 5% 110%,rgba(91,155,255,.20),transparent 70%)}
.ux-hero > *{position:relative}
.ux-kicker{display:inline-block;padding:6px 16px;border-radius:99px;font-size:.72rem;font-weight:800;letter-spacing:.09em;color:#F6DE8D;background:rgba(226,185,59,.12);border:1px solid rgba(226,185,59,.45)}
.ux-hero-title{font-size:clamp(1.5rem,3.2vw,2.3rem);font-weight:800;letter-spacing:-.02em;margin:14px 0 6px;color:#FFFFFF}
.ux-hero-sub{margin:0 auto;max-width:640px;font-size:.95rem;line-height:1.6;color:#B7C4E4}
.ux-stats{display:flex;flex-wrap:wrap;justify-content:center;gap:12px;margin-top:20px}
.ux-stat{min-width:120px;flex:1 1 120px;max-width:200px;padding:12px 14px;border-radius:16px;background:rgba(255,255,255,.06);border:1px solid rgba(255,255,255,.12)}
.ux-stat b{display:block;font-size:1.55rem;font-weight:800;color:#FFFFFF;line-height:1.1}
.ux-stat span{display:block;margin-top:4px;font-size:.7rem;font-weight:700;letter-spacing:.06em;text-transform:uppercase;color:#9FB1D9}
.ux-split{display:flex;height:9px;border-radius:99px;overflow:hidden;margin:20px auto 8px;max-width:560px;background:rgba(255,255,255,.12)}
.ux-split i{display:block;height:100%}
.ux-legend{display:flex;flex-wrap:wrap;justify-content:center;gap:6px 16px;font-size:.74rem;color:#B7C4E4}
.ux-legend em{display:inline-block;width:9px;height:9px;border-radius:50%;margin-right:6px;vertical-align:baseline;font-style:normal}

/* --- grafik --- */
img.ux-chart{display:block;width:100%;height:auto;border-radius:12px}
.ux-bars{display:flex;flex-direction:column;gap:12px;padding:4px 2px}
.ux-bleg{display:flex;flex-wrap:wrap;gap:6px 16px;font-size:.78rem;font-weight:600;color:var(--ux-muted);margin-bottom:2px}
.ux-bleg em{display:inline-block;width:10px;height:10px;border-radius:3px;margin-right:6px;font-style:normal}
.ux-brow{display:grid;grid-template-columns:minmax(80px,32%) 1fr;gap:6px 12px;align-items:center}
.ux-blbl{font-size:.82rem;font-weight:700;color:var(--ux-ink);overflow-wrap:anywhere;line-height:1.3}
.ux-bset{display:flex;flex-direction:column;gap:5px}
.ux-bline{display:grid;grid-template-columns:1fr auto;gap:8px;align-items:center}
.ux-btrack{height:10px;border-radius:99px;background:var(--ux-line);overflow:hidden}
.ux-btrack i{display:block;height:100%;border-radius:99px;min-width:3px}
.ux-bval{font-size:.76rem;font-weight:800;color:var(--ux-ink);font-variant-numeric:tabular-nums;min-width:48px;text-align:right}
.ux-note{font-size:.76rem;color:var(--ux-muted);margin-top:4px}
.ux-prog{display:flex;height:16px;border-radius:99px;overflow:hidden;background:var(--ux-line);margin:6px 0 10px}
.ux-prog i{display:block;height:100%}
.ux-plegend{display:flex;flex-wrap:wrap;gap:6px 18px;font-size:.82rem;font-weight:600;color:var(--ux-ink)}
.ux-plegend em{display:inline-block;width:10px;height:10px;border-radius:50%;margin-right:7px;font-style:normal}

/* --- saklar tampilan Web | HP --- */
.st-key-view_switch{display:flex;justify-content:flex-end;padding:10px 0 2px}
.st-key-view_switch div[role="radiogroup"]{display:inline-flex;gap:2px;background:var(--ux-line);padding:3px;border-radius:99px}
.st-key-view_switch div[role="radiogroup"] label{margin:0 !important;padding:6px 16px !important;border-radius:99px;cursor:pointer}
.st-key-view_switch div[role="radiogroup"] label > div:first-child{display:none !important}
.st-key-view_switch div[role="radiogroup"] label:has(input:checked){background:var(--ux-card);box-shadow:0 1px 3px rgba(11,23,48,.2)}
.st-key-view_switch div[role="radiogroup"] label p{font-size:.82rem !important;font-weight:800 !important;margin:0 !important}

@media (hover:none){button:hover,.area-card:hover,.login-box:hover{transform:none !important}}
@media (prefers-reduced-motion:reduce){*,*::before,*::after{animation:none !important;transition:none !important}}
</style>
"""

# ----------------------------------------------------------------------------
# CSS KHUSUS MODE HP (satu kolom, area sentuh besar, navigasi atas menempel)
# ----------------------------------------------------------------------------
_UX_CSS_HP = """
<style>
.stApp{background:var(--ux-bg) !important}
[data-testid="stMainBlockContainer"],.stMainBlockContainer,.main .block-container{max-width:480px !important;padding:.3rem .8rem 5rem !important;background:var(--ux-bg)}
@media (min-width:700px){[data-testid="stMainBlockContainer"],.stMainBlockContainer,.main .block-container{min-height:100vh;box-shadow:0 0 0 1px var(--ux-line),0 24px 60px rgba(11,23,48,.14)}}
.st-key-view_switch{justify-content:stretch}
.st-key-view_switch div[role="radiogroup"]{display:flex;width:100%}
.st-key-view_switch div[role="radiogroup"] label{flex:1;justify-content:center;padding:10px 0 !important}

/* kolom ditumpuk; 4 kolom (kartu menu, lat/long) -> 2x2; deretan metrik -> 2 per baris */
div[data-testid="stHorizontalBlock"]{flex-direction:row !important;flex-wrap:wrap !important;gap:.6rem !important}
div[data-testid="stHorizontalBlock"] > div[data-testid="stColumn"],div[data-testid="stHorizontalBlock"] > div[data-testid="column"]{flex:1 1 100% !important;width:100% !important;min-width:100% !important}
div[data-testid="stHorizontalBlock"]:has(> div:nth-child(4):last-child) > div[data-testid="stColumn"],
div[data-testid="stHorizontalBlock"]:has(> div:nth-child(4):last-child) > div[data-testid="column"],
div[data-testid="stHorizontalBlock"]:has([data-testid="stMetric"]) > div[data-testid="stColumn"],
div[data-testid="stHorizontalBlock"]:has([data-testid="stMetric"]) > div[data-testid="column"]{flex:1 1 calc(50% - .3rem) !important;width:calc(50% - .3rem) !important;min-width:calc(50% - .3rem) !important}

/* baris navigasi atas halaman dalam: judul di atas, tombol Kembali & Keluar sebaris, menempel saat scroll */
div[data-testid="stHorizontalBlock"]:has(.st-key-btn_back_main){display:grid !important;grid-template-columns:1fr 1fr;gap:.5rem !important;position:sticky;top:0;z-index:60;background:var(--ux-bg);padding:8px 0 6px;margin-bottom:6px}
div[data-testid="stHorizontalBlock"]:has(.st-key-btn_back_main) > div{width:100% !important;min-width:0 !important}
div[data-testid="stHorizontalBlock"]:has(.st-key-btn_back_main) > div:nth-child(2){grid-column:1 / -1;order:-1}

/* tipografi & kartu */
h1{font-size:1.45rem !important;line-height:1.25 !important}
h2{font-size:1.2rem !important} h3{font-size:1.05rem !important} h4{font-size:1rem !important}
.area-card,.login-box{padding:20px 16px;border-radius:18px}
.login-title{font-size:1.1rem} .login-desc{font-size:.84rem}
.ux-hero{padding:22px 16px 18px;border-radius:20px}
.ux-stat{min-width:0;flex:1 1 calc(50% - 12px);max-width:none}
.ux-stat b{font-size:1.3rem}
.ux-brow{grid-template-columns:1fr}
.watermark-kut img{width:420px !important}
div[data-testid="stMetric"]{padding:12px 12px 10px 16px}
div[data-testid="stMetricValue"]{font-size:1.1rem !important;font-size:clamp(.95rem,7.5cqw,1.25rem) !important}

/* area sentuh >= 48px, input 16px agar iPhone tidak zoom sendiri */
.stApp div.stButton > button,.stApp div[data-testid="stFormSubmitButton"] > button,.stApp div[data-testid="stDownloadButton"] > button{min-height:54px !important;padding:10px 12px !important;font-size:.9rem !important;line-height:1.25 !important}
.stApp div.stButton > button *{white-space:normal !important}
div[data-baseweb="input"] input,div[data-baseweb="base-input"] input,div[data-baseweb="textarea"] textarea,div[data-baseweb="select"] input,.stTextInput input,.stNumberInput input,.stDateInput input,.stTextArea textarea{font-size:16px !important}
div[data-baseweb="input"] input,div[data-baseweb="base-input"] input{padding:12px 14px !important}
div[data-baseweb="select"] > div{min-height:50px !important}
li[role="option"]{min-height:46px !important}
div[data-testid="stCheckbox"] label{min-height:44px;align-items:center}

/* tab bisa digeser, upload selebar layar */
.stTabs [data-baseweb="tab-list"]{overflow-x:auto;flex-wrap:nowrap;scrollbar-width:none}
.stTabs [data-baseweb="tab"]{white-space:nowrap;flex:0 0 auto;padding:10px 14px}
div[data-testid="stExpander"] summary{min-height:48px}
section[data-testid="stFileUploaderDropzone"]{flex-direction:column;align-items:stretch;gap:.6rem;padding:1rem}
section[data-testid="stFileUploaderDropzone"] button{width:100%;min-height:48px}
</style>
"""

# CSS layar login (pengganti blok lama: tombol besar merah/biru sesuai label 🔴 / 🔵)
_UX_CSS_LOGIN = """
<style>
.stApp div.stButton > button{min-height:68px !important;font-size:1.02rem !important;font-weight:800 !important;text-transform:uppercase;letter-spacing:.06em;border-radius:14px !important}
.stApp button[kind="primary"],.stApp button[data-testid="stBaseButton-primary"]{background:linear-gradient(135deg,#F43F5E,#BE123C) !important;color:#fff !important;border:1px solid rgba(255,255,255,.2) !important;box-shadow:0 10px 22px -8px rgba(225,29,72,.6) !important}
.stApp button[kind="secondary"],.stApp button[data-testid="stBaseButton-secondary"]{background:linear-gradient(135deg,#3B82F6,#1D4ED8) !important;color:#fff !important;border:1px solid rgba(255,255,255,.2) !important;box-shadow:0 10px 22px -8px rgba(37,99,235,.6) !important}
</style>
"""

# ----------------------------------------------------------------------------
# BOOT: deteksi perangkat, pasang tema, tampilkan saklar Web | HP
# ----------------------------------------------------------------------------
def _terdeteksi_hp():
    try:
        ua = (st.context.headers.get("User-Agent") or "").lower()
        return any(k in ua for k in ("iphone", "ipod", "mobile"))
    except Exception:
        return False

if "view_mode" not in st.session_state:
    _qp = str(st.query_params.get("view", "")).lower()
    st.session_state.view_mode = _qp if _qp in ("web", "hp") else ("hp" if _terdeteksi_hp() else "web")

def _ganti_tampilan():
    st.session_state.view_mode = "hp" if "HP" in st.session_state.get("_view_pick", "") else "web"
    st.query_params["view"] = st.session_state.view_mode

IS_HP = st.session_state.view_mode == "hp"
UX_MODE = _theme_mode()

def ux_boot():
    """Panggil sekali di awal halaman: pasang CSS tema + saklar tampilan."""
    st.markdown(_UX_CSS.replace("__ROOT__", _root_css(UX_MODE)) + (_UX_CSS_HP if IS_HP else ""), unsafe_allow_html=True)
    try:
        box = st.container(key="view_switch")
    except TypeError:
        box = st.container()
    with box:
        st.radio("Tampilan", ["🖥️ Web", "📱 HP"], index=1 if IS_HP else 0, horizontal=True,
                 key="_view_pick", on_change=_ganti_tampilan, label_visibility="collapsed")

def ux_login_css():
    st.markdown(_UX_CSS_LOGIN, unsafe_allow_html=True)

# ----------------------------------------------------------------------------
# KOMPONEN HTML RINGAN
# ----------------------------------------------------------------------------
def _fmt(v):
    """Angka ringkas: 12,5 jt • 1,25 M • 54.8 • 1,234"""
    try:
        v = float(v)
    except Exception:
        return "-"
    if _math.isnan(v):
        return "-"
    a = abs(v)
    if a >= 1e9: return ("%.2f M" % (v / 1e9)).replace(".00", "")
    if a >= 1e6: return ("%.1f jt" % (v / 1e6)).replace(".0 ", " ")
    if a >= 1e4: return "{:,.0f}".format(v)
    if v == int(v): return "{:,.0f}".format(v)
    return ("%.2f" % v).rstrip("0").rstrip(".")

def _esc(x):
    return _html.escape(str(x), quote=True)

def ux_hero(kicker, title, subtitle="", stats=None, split=None):
    """Banner utama. stats=[(nilai,label)]  split=[(label,jumlah,warna)] untuk bar status tipis."""
    h = '<div class="ux-hero">'
    if kicker: h += '<span class="ux-kicker">%s</span>' % _esc(kicker)
    h += '<div class="ux-hero-title">%s</div>' % _esc(title)
    if subtitle: h += '<p class="ux-hero-sub">%s</p>' % _esc(subtitle)
    if stats:
        h += '<div class="ux-stats">' + "".join('<div class="ux-stat"><b>%s</b><span>%s</span></div>' % (_esc(v), _esc(l)) for v, l in stats) + "</div>"
    if split:
        tot = sum(max(float(n), 0) for _, n, _ in split) or 1.0
        h += '<div class="ux-split">' + "".join('<i style="width:%.2f%%;background:%s"></i>' % (max(float(n), 0) / tot * 100, c) for _, n, c in split if n > 0) + "</div>"
        h += '<div class="ux-legend">' + "".join('<span><em style="background:%s"></em>%s · %s</span>' % (c, _esc(l), _esc(_fmt(n))) for l, n, c in split) + "</div>"
    h += "</div>"
    st.markdown(h, unsafe_allow_html=True)

def ux_stats_from_db(db):
    """Ringkasan cepat dari laporan_db yang sudah ada di memori (tanpa query baru)."""
    db = db or []
    pbu = sum(1 for r in db if r.get("kategori") == "Pangkalan Bun")
    prj = sum(1 for r in db if r.get("kategori") == "Project")
    plk = len(db) - pbu - prj
    ok = sum(1 for r in db if r.get("status") in ("Normal", "Done"))
    bad = sum(1 for r in db if r.get("status") == "Major/Critical")
    mid = len(db) - ok - bad
    stats = [(len(db), "Total Laporan"), (plk, "NOP Palangkaraya"), (pbu, "NOP Pangkalan Bun"), (prj, "Project")]
    split = [("Normal / Done", ok, "#34D399"), ("Minor / Progress", mid, "#FBBF24"), ("Major / Critical", bad, "#FB7185")]
    return stats, split

def ux_progress(items):
    """Bar bertumpuk. items=[(label, jumlah, warna)]"""
    tot = sum(max(float(n), 0) for _, n, _ in items) or 1.0
    h = '<div class="ux-prog">' + "".join('<i style="width:%.2f%%;background:%s"></i>' % (max(float(n), 0) / tot * 100, c) for _, n, c in items if n > 0) + "</div>"
    h += '<div class="ux-plegend">' + "".join('<span><em style="background:%s"></em>%s: <b>%s</b></span>' % (c, _esc(l), _esc(_fmt(n))) for l, n, c in items) + "</div>"
    st.markdown(h, unsafe_allow_html=True)

# ----------------------------------------------------------------------------
# GRAFIK SVG (disajikan sebagai <img> data-URI -> nol JavaScript, nol lag)
# ----------------------------------------------------------------------------
_DEF_COLORS = ["#2563EB", "#E2B93B", "#10B981", "#F43F5E", "#8B5CF6", "#06B6D4"]

def _as_frame(data):
    if isinstance(data, pd.Series):
        data = data.to_frame()
    df = pd.DataFrame(data).copy()
    for c in df.columns:
        df[c] = pd.to_numeric(df[c], errors="coerce")
    return df

# warna neon bawaan (dibuat untuk tema gelap) -> padanan mewah yang tetap kontras di tema terang
_REMAP = {"#64ffda": "#10B981", "#ff5252": "#F43F5E", "#00e676": "#10B981", "#ffc107": "#F59E0B", "#2979ff": "#2563EB",
          "#ffffff": "#E11D48", "#3b82f6": "#2563EB", "#10b981": "#10B981", "#ef4444": "#F43F5E", "#f59e0b": "#F59E0B", "#a855f7": "#8B5CF6"}

def _colors(color, n):
    if color is None: cols = []
    elif isinstance(color, (list, tuple)): cols = list(color)
    else: cols = [color]
    cols = [_REMAP.get(c.lower(), c) if isinstance(c, str) else c for c in cols]
    cols = [c if (isinstance(c, str) and c.startswith("#") and len(c) == 7) else None for c in cols]
    return [(cols[i] if i < len(cols) and cols[i] else _DEF_COLORS[i % len(_DEF_COLORS)]) for i in range(n)]

def _ticks(lo, hi, n=4):
    """Tick 'rapi' dengan paling banyak 6 interval dan rentang sekecil mungkin."""
    if hi <= lo:
        hi = lo + (1.0 if lo == 0 else abs(lo) * 0.1)
    mag = 10 ** _math.floor(_math.log10((hi - lo) / n))
    cands = sorted(m * k for k in (mag / 10, mag, mag * 10) for m in (1, 2, 2.5, 5))
    for step in cands:
        start, end = _math.floor(lo / step) * step, _math.ceil(hi / step) * step
        if round((end - start) / step) <= 6 and (end - start) / step >= 3:
            break
    k = int(round((end - start) / step))
    return [round(start + i * step, 10) for i in range(k + 1)]

def _trunc(s, n):
    s = str(s)
    return s if len(s) <= n else s[: max(n - 1, 1)] + "…"

def _svg_chart(kind, df, colors, height=None, w=None):
    ns, n = df.shape[1], df.shape[0]
    W = w or (400 if IS_HP else 640)
    fs = 13 if IS_HP else 12
    cw = fs * 0.58
    names = [str(c) for c in df.columns]
    # --- legenda (membungkus baris) ---
    lx, ly, rows, leg = 14, 16, 1, []
    for i, nm in enumerate(names):
        t = _trunc(nm, 26)
        wd = 18 + len(t) * cw + 14
        if lx + wd > W - 10 and lx > 14:
            lx, ly, rows = 14, ly + fs + 8, rows + 1
        leg.append((lx, ly, t, colors[i]))
        lx += wd
    if ns == 1 and kind == "bar":
        leg = []; rows = 0
    top = (rows * (fs + 8) + 14) if leg else 18
    bot = fs + 24
    vals = df.values.astype(float)
    finite = vals[_np.isfinite(vals)]
    lo = float(finite.min()) if finite.size else 0.0
    hi = float(finite.max()) if finite.size else 1.0
    if kind == "bar":
        lo = min(lo, 0.0)
    else:
        pad = (hi - lo) * 0.08 if hi > lo else (abs(hi) * 0.05 or 0.5)
        nonneg = lo >= 0
        lo, hi = lo - pad, hi + pad
        if nonneg and lo < 0:
            lo = 0.0
    tk = _ticks(lo, hi)
    ymin, ymax = tk[0], tk[-1]
    left = int(max(len(_fmt(t)) for t in tk) * cw + 18)
    right = 16
    H = int(height) if height else (290 if IS_HP else 300)
    pw, ph = W - left - right, H - top - bot
    X = lambda i: left + (pw / 2 if n == 1 else i * pw / (n - 1))
    Y = lambda v: top + ph * (1 - (v - ymin) / ((ymax - ymin) or 1))
    o = ['<svg xmlns="http://www.w3.org/2000/svg" viewBox="0 0 %d %d" font-family="system-ui,-apple-system,Segoe UI,Roboto,sans-serif" font-size="%s">' % (W, H, fs)]
    o.append("<style>%s</style><defs>" % _svg_style(UX_MODE))
    for i, c in enumerate(colors[:ns]):
        o.append('<linearGradient id="a%d" x1="0" y1="0" x2="0" y2="1"><stop offset="0" stop-color="%s" stop-opacity="%s"/><stop offset="1" stop-color="%s" stop-opacity="%s"/></linearGradient>'
                 % (i, c, ".30" if kind == "line" else "1", c, "0" if kind == "line" else ".62"))
    o.append("</defs>")
    for t in tk:  # grid + label sumbu Y
        y = Y(t)
        o.append('<line class="gr" x1="%d" x2="%d" y1="%.1f" y2="%.1f" stroke-width="1" stroke-dasharray="3 4"/>' % (left, W - right, y, y))
        o.append('<text class="tx" x="%d" y="%.1f" text-anchor="end" dominant-baseline="middle">%s</text>' % (left - 8, y, _xesc(_fmt(t))))
    maxl = 5 if IS_HP else 8
    step = max(1, int(_math.ceil(n / maxl)))
    for i in range(0, n, step):
        if kind == "line":
            xx = X(i)
            anc = "start" if (i == 0 and n > 2) else ("end" if (i + step >= n and n > 2 and i == n - 1) else "middle")
        else:
            xx, anc = left + (i + .5) * pw / n, "middle"
        o.append('<text class="tx" x="%.1f" y="%d" text-anchor="%s">%s</text>' % (xx, H - bot + fs + 8, anc, _xesc(_trunc(df.index[i], 7 if IS_HP else 10))))
    for (x, y, t, c) in leg:
        o.append('<circle cx="%d" cy="%d" r="4.5" fill="%s"/><text class="tx" x="%d" y="%d" dominant-baseline="middle">%s</text>' % (x + 4, y, c, x + 14, y, _xesc(t)))
    badges = []
    if kind == "line":
        for si in range(ns):
            col = df.iloc[:, si].values.astype(float)
            ok = _np.isfinite(col)
            if not ok.any(): continue
            const = ok.sum() >= 3 and _np.nanmax(col) == _np.nanmin(col)
            c = colors[si]
            segs, cur = [], []
            for i in range(n):
                if ok[i]: cur.append((X(i), Y(col[i])))
                elif cur: segs.append(cur); cur = []
            if cur: segs.append(cur)
            if const:
                for sg in segs:
                    o.append('<polyline fill="none" stroke="%s" stroke-width="1.8" stroke-dasharray="6 5" points="%s"/>' % (c, " ".join("%.1f,%.1f" % p for p in sg)))
                continue
            base = Y(ymin)
            for sg in segs:
                pts = " ".join("%.1f,%.1f" % p for p in sg)
                if si == 0 and len(sg) > 1:
                    o.append('<polygon fill="url(#a0)" points="%s %.1f,%.1f %.1f,%.1f"/>' % (pts, sg[-1][0], base, sg[0][0], base))
                o.append('<polyline fill="none" stroke="%s" stroke-width="2.6" stroke-linejoin="round" stroke-linecap="round" points="%s"/>' % (c, pts))
                if len(sg) == 1:
                    o.append('<circle cx="%.1f" cy="%.1f" r="3.6" fill="%s"/>' % (sg[0][0], sg[0][1], c))
            if n <= 24:
                for i in range(n):
                    if ok[i]:
                        o.append('<circle class="bs" cx="%.1f" cy="%.1f" r="3.2" fill="%s" stroke-width="1.6"/>' % (X(i), Y(col[i]), c))
            li = int(_np.where(ok)[0][-1])
            badges.append((Y(col[li]), X(li), col[li], c))
        placed = []
        for (y, x, v, c) in sorted(badges, key=lambda b: b[0]):
            txt = _fmt(v); bw = len(txt) * cw + 14; bh = fs + 7
            by = y - bh - 6
            for py in placed:
                if abs(by - py) < bh + 2: by = py + bh + 3
            placed.append(by)
            bx = min(max(x - bw / 2, left), W - right - bw)
            o.append('<circle cx="%.1f" cy="%.1f" r="5" fill="%s" stroke="#fff" stroke-width="2"/>' % (x, y, c))
            o.append('<rect x="%.1f" y="%.1f" width="%.1f" height="%d" rx="7" fill="%s"/><text x="%.1f" y="%.1f" text-anchor="middle" dominant-baseline="middle" fill="#fff" font-weight="700">%s</text>'
                     % (bx, by, bw, bh, c, bx + bw / 2, by + bh / 2 + .5, _xesc(txt)))
    else:  # bar vertikal
        gw = pw / max(n, 1)
        bwid = max(min(34, gw * 0.72 / ns), 4)
        for i in range(n):
            gx = left + i * gw + (gw - bwid * ns) / 2
            for si in range(ns):
                v = df.iat[i, si]
                if not _np.isfinite(v): continue
                y0, y1 = Y(max(v, 0)), Y(min(v, 0))
                x = gx + si * bwid
                o.append('<rect x="%.1f" y="%.1f" width="%.1f" height="%.1f" rx="3.5" fill="url(#a%d)"/>' % (x + 1, y0, bwid - 2, max(y1 - y0, 1.5), si))
                if bwid >= (22 if not IS_HP else 24) and n <= 10:
                    o.append('<text class="ik" x="%.1f" y="%.1f" text-anchor="middle" font-weight="700" font-size="%s">%s</text>' % (x + bwid / 2, y0 - 5, fs - 1, _xesc(_fmt(v))))
    o.append("</svg>")
    return "".join(o)

def _img(svg, alt):
    import base64 as _b64
    return '<img class="ux-chart" alt="%s" src="data:image/svg+xml;base64,%s">' % (_esc(alt), _b64.b64encode(svg.encode("utf-8")).decode())

def _downsample(df, limit):
    if len(df) <= limit:
        return df
    idx = sorted(set(_np.linspace(0, len(df) - 1, limit).round().astype(int).tolist()))
    return df.iloc[idx]

def ui_line_chart(data, color=None, height=None, **kw):
    """Pengganti st.line_chart: garis + area + lencana nilai terakhir; garis konstan = batas (putus-putus)."""
    df = _as_frame(data)
    if df.empty or df.shape[1] == 0:
        st.info("Belum ada data untuk ditampilkan."); return
    df = _downsample(df, 120 if IS_HP else 220)
    st.markdown(_img(_svg_chart("line", df, _colors(color, df.shape[1]), height), "Grafik garis"), unsafe_allow_html=True)

def ui_bar_chart(data, color=None, height=None, horizontal=None, **kw):
    """Pengganti st.bar_chart: kolom vertikal untuk kategori sedikit, batang horizontal (HTML) untuk kategori banyak."""
    df = _as_frame(data)
    if df.empty or df.shape[1] == 0:
        st.info("Belum ada data untuk ditampilkan."); return
    cols = _colors(color, df.shape[1])
    n = len(df)
    maxlbl = max((len(str(i)) for i in df.index), default=0)
    vertical = (not horizontal) if horizontal is not None else (n <= (6 if IS_HP else 12) and maxlbl <= 12 and df.shape[1] <= 3)
    if vertical:
        st.markdown(_img(_svg_chart("bar", df, cols, height), "Grafik batang"), unsafe_allow_html=True); return
    LIM = 40
    shown = df.head(LIM)
    vmax = float(_np.nanmax(shown.values)) if _np.isfinite(shown.values).any() else 1.0
    pct_mode = all("%" in str(c) for c in df.columns) and vmax <= 100
    scale = 100.0 if pct_mode else (vmax or 1.0)
    h = '<div class="ux-bars">'
    if df.shape[1] > 1:
        h += '<div class="ux-bleg">' + "".join('<span><em style="background:%s"></em>%s</span>' % (cols[i], _esc(c)) for i, c in enumerate(df.columns)) + "</div>"
    for idx, row in shown.iterrows():
        h += '<div class="ux-brow"><div class="ux-blbl">%s</div><div class="ux-bset">' % _esc(idx)
        for i, c in enumerate(df.columns):
            v = row[c]
            w = 0.0 if not _np.isfinite(v) else max(min(float(v) / scale * 100, 100), 0)
            h += '<div class="ux-bline"><div class="ux-btrack"><i style="width:%.1f%%;background:linear-gradient(90deg,%s,%s)"></i></div><div class="ux-bval">%s</div></div>' % (w, cols[i], cols[i] + "AA", _esc(_fmt(v)))
        h += "</div></div>"
    h += "</div>"
    if n > LIM:
        h += '<div class="ux-note">Menampilkan %d dari %d baris. Lihat tabel untuk data lengkap.</div>' % (LIM, n)
    st.markdown(h, unsafe_allow_html=True)

def ui_ring(pct, label="", color="#10B981", size=132):
    """Cincin persentase."""
    pct = 0.0 if pct is None or (isinstance(pct, float) and _math.isnan(pct)) else max(min(float(pct), 100.0), 0.0)
    r, c = 52, 2 * _math.pi * 52
    svg = ('<svg xmlns="http://www.w3.org/2000/svg" viewBox="0 0 132 150" font-family="system-ui,-apple-system,Segoe UI,Roboto,sans-serif"><style>%s</style>'
           '<circle class="gr" cx="66" cy="66" r="%d" fill="none" stroke-width="12"/>'
           '<circle cx="66" cy="66" r="%d" fill="none" stroke="%s" stroke-width="12" stroke-linecap="round" stroke-dasharray="%.1f %.1f" transform="rotate(-90 66 66)"/>'
           '<text class="ik" x="66" y="70" text-anchor="middle" dominant-baseline="middle" font-size="26" font-weight="800">%d%%</text>'
           '<text class="tx" x="66" y="143" text-anchor="middle" font-size="12" font-weight="700">%s</text></svg>') % (_svg_style(UX_MODE), r, r, color, c * pct / 100, c, round(pct), _xesc(_trunc(label, 22)))
    st.markdown('<div style="max-width:%dpx;margin:0 auto">%s</div>' % (size, _img(svg, label or "Persentase")), unsafe_allow_html=True)


# ----------------------------------------------------------------------------
# PERFORMA: baca Excel hanya SEKALI per perubahan file (bukan di setiap klik)
# ----------------------------------------------------------------------------
@st.cache_data(show_spinner=False, max_entries=24)
def _read_excel_cached(path, stamp, kwargs_key):
    return pd.read_excel(path, **dict(kwargs_key))

def read_excel_cached(path, **kwargs):
    """Sama seperti pd.read_excel, tetapi hasilnya di-cache selama file belum berubah.
    File baru/diedit (waktu ubah berubah) otomatis dibaca ulang. Objek file upload tidak di-cache."""
    if isinstance(path, str) and os.path.exists(path):
        stt = os.stat(path)
        return _read_excel_cached(path, (stt.st_mtime_ns, stt.st_size), tuple(sorted(kwargs.items())))
    return pd.read_excel(path, **kwargs)

@st.cache_data(show_spinner=False, max_entries=8)
def _b64_cached(path, stamp):
    with open(path, "rb") as f:
        return base64.b64encode(f.read()).decode()

# -------------------------------------------------------------------------
# FUNGSI PEMBUAT DOKUMEN (PDF & PPTX)
# -------------------------------------------------------------------------
def clean_text(text):
    if not text: return "-"
    return str(text).encode('latin-1', 'ignore').decode('latin-1')

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

    pdf.set_font("helvetica", "", 10)
    pdf.cell(90, 6, "", 0, 0) 
    pdf.cell(90, 6, clean_text("Data dibuat oleh,"), 0, 1, "C")
    try:
        TTD_OKTA_B64 = "iVBORw0KGgoAAAANSUhEUgAAAF0AAABkCAYAAADkBDymAAAVAklEQVR42u2de3hU1bXA197nfeY9yUwykwSIgjxVCvEBioBWW21v7y3fbfp9fWlvq2ktjwq+COpkeAiIQSmUNtQr91L97m28t7VetajViCLPAIogJDzCIyGZZDKZ53mfs+8fWDTykEBmQjDrr+Sbsx/rN/uss9baa58BGJABybYQQlBjYyPXl3Ogv2rQq5a/4dFl7SYA+L8B6DmSzqTDTzNs0cBKz6FYNO8nDO8YgJ5DUTXWpRlaxwD0nD5JdYeuWfG+nAL+qjHHAD5ZM2MD0HMoiqrSGNF9utK/cuaFE9hCOzLTA9BzKF0JnDGNWGLAvOQuGqX8ecLQSWOHkwHoOZLR5S9RXV1SZPLPS1ID5iVHcsPQsS5EGa7rMNIHoOfqtqbUfFk1EoT08Ty+StBlxfDSLCdfArHCV0cEG1cgpa3jfT2Pr5R5yaR1B0Ja52ULvbp2k0BLjmBHpJn2uX2Gk021/PSnU5W+VFYU2aJYQn/vsoP+dE1dPhj4NiZhuhle7xTstg5gkFPG3LQV/7Hxo1n33PxmXymbSMmg6+yJywr60t+/d5tpklKaJltn3jvp489/VlNfvx7tN+9ZXrPhjtkVk/sEvM9DD3G5bMZrl8ODdG1dE7+k5u0fYLBEKfLO83Mqpnz8xWsqysr0X/7ohj+YFi75zbotpX0QjeJE0srwtBzr9yu9unqT0PrJsbs5Vtg5596J277s+jQ2/momje8CIc8BQjnzmMur9opOkfGOnjJK69cuY3X1JkFnpJ9gIBvOBzgAQLhiahRjOrN01VuBXCrqBT3f0LVM+RjUf6GHQiEs0co0iqY+ePRXU/b1pK2imdtVzN6YS0Ulic83CdXnpuWioLuDd34HLOr4Q/dP3dPTtiOCkw8zNFVaU1PP5EpRkzIKDR239lvoS1a9PdowddfjD0y+IJ+3vByZmKDmuJUekrNolIIShjb6J/TnXt7nME24xUkJL13MwLpuHkgpcHXONEVcPgLS3C+hx1o6/tlht22qqCiTLmZgN8prFFimBBDKiaK6roo2wRvtd9CX/e69URhR7Myf3fjRxQ48ffqYNKYoecmfDrpy4KMjhCyX4JCT/Qp6bW0tpRlkKmLQy701uKZbHVRnR9ZL3Gb+5gBrWgx348ThmX4F/VisaIrICofn3Dux19wuQ4b2eJouzLaSyeMtHkIs9fsTkdxvoIdCdbxmqKPjniNv9+bgqUymNSVFPVk3L0yenyAjSeDSkPOCjr3odgB6a7i8vFejOZVv7xIdgi/bSjKsWEwA2qG/QK+prXexLB5aOf2Wbb09uNmUTusKrRJCsurCGHrGTxN8vN9Ab4ukb81k0u8CQK/fnTU19xleN11Q9dLerEamDEUXCyIV6RfQq2s3CciyfNcMSu/OSryCEJElrR1a03xv9blw+d/H1dS81c0NTcmGSEPfbV7U1RF69bo3/HV1dTTAl6R2lVbtRgozTeXl5WbWViHHsxRNOwCgxz704tXbrkjEUzeohs4Egp4Dj9x74+ajUft3j+1vX/yZE0Bwo1Tv9BV7cpbsWlW7x64mkoUMxQYSaSX4/vY3hXwvn5LlwKsAYJwTuolhmKA4XszmBFsjnc2YGD3O6y9ds3GSqqs3O/Jcf+HSWns6nv5GeOUmz7FO3ram5p/kNWtOXnfI3iZghRFHTyjJStEoIQRVLX/D4xRtJYxgC6qq7k+0ddF2u9hOc3CCp+116UNvRe9/OGz9o81ZlV2w6oPBFLbUhx4am9WAotBX4JJkNQAAR8+3zaJn68bIklxGgoHqJ8rHaJ8Gb7U7m13rKMIeRp/bHOEV1Z/QcLK38uihmnrRgcximyAGVVX1LVi+gaEoHmiGaRVF4bBsSBvnzZx0zrK9s0InsjJeBWNHtm/Fg0ePNFkGnHeZW+gPe7x6put7BAnV4U+BAwCUl5eb02Z+tM0uUt1MIQf6YIZGF5RzCYVqWb4gUOxyOItUEzyKrLnlVNJUBTrDYO4oEH3fiJLJ0fJy1CPzS5/llsHLVn9QdN2oW18JZxm6Ly+PsQn2IACc1xeMlfg9ugW1C2ffmPy83QZ4l93RyWFJ7e6lpFKyHyz9S++i0No6HiKQz9j0Qsy6g8jIOOMpw0KGbloWaeE5stvATMf8X9560Xf+GaE/9vR7g2mMuqZORUa2V7pdEIxkRqLO59rHqt+/1rQseeHsW/Z2+yBv+w9sgnskncR+jxd3PzlHGQFVRtu7eWXVmwSJwYWcQAVZxvLTNJ/fFY1raUOLMjpETBvs4gVPbOmvsmNazwgdcfS1mmXuhhzIto/2tLA0J57HIwsh8v43nLz4/BdSFLQoUqMFG8IOkTgLXTG2m3dEM1eWDkENi36zcRLL0AWI4uyqHDe1DIpbFNPGEWo749Vjj/9osgQ5kjPbdEVxsiNcB3MxgUJ3Ifj8jiFfdnx50bNbrpY0kniwoqybfbYF8wcbun7cUFQ/EJxKSal05VMbx4qcUWJ3efO3HRTGe8zYNoOBDKXBFsaQOyvnTO3TxNdp0J9cuTUvnU7Cwu+U5eSbzy8OII5CX5r00ghM0S3uz90j2nrxcMq8g6ZRcWfaMZoTKIdJ8mnMJPfrpr67K8UolKXc+vTD178Al5CcBj2dSQwjmG7J1QSKvB4UiXWcM8Xw5MpP8mKpmCc/j+B5S96eiDi20Od05UumoRAgE2gav0KbMAybzN/27T/8yd9fTK0HKDfvDx0YCiZJIQRAyCUMHbMwyCLWB7maQHukLdGZiHbzONa90Wprb4l6O2JyvoNWhyYNMpm1uVLEio83iHXEzRs7r7mis3n3bpliwcZGk1odTfLLJY0+PHLk4MAdawXfwz+FNsXQgkCg5VICfhr0UIhgmtlWMLzAlrNj3MdbG3WbI+Bd+vzeSbFoxC+w4qATx5oZlmVkrxNFDZ07oacztMCoL/js+YbNLvpbo4klWxuZcArMUgvphcE8atLhFjzIZTMYliXFuqa6a2uPJd5pzAw1LOUAwCUMnfVs8SmaZZZ/LujozXB56Zodzs647td0+xWmga+xELnKwIwfY3psOtP2nsflSKQTdL2sk7QzaLgddnAeOW6IpsWzImuUmFh3ZFKJMgKEyyTjpmoyXwsUeCPxeGKMqvM+t8vItwi+WlOt6JH2zrRlif9qF+ndT1S/Txw2NqLrXFtegRGpKC9LXDLQUzoMAUyOXGyny1Yf9DerZh6SUyWmZhVYIDvnLn2f4njGALAUYtIJwc41GCjzJkjpuFfE9y2YcdO8M/X18JLNt9sEtaFy+sSNAADzV24LcsjcFH5owtZ5y7dPHjek9Jn/Xv/3m2lMvnWsfcgfWX2f5BxZ8vLD946K/fjR3WOLAq6XGK09SlH2YkSTUa3t5nVzl6wXeZstaiHq46pfTTiEclhTeRp0m0gNUTXzwwvp6PGn3p0Ql5hxiu72HYkmKYvC7RyQVoFHDTTrOz7YJkUrKspOC/dnrHidw8R31iAsIyujWZlbd+p/k7kVANbVvELE4wd3cVOnIuOBp7YP15KY3HWHR33zZYp58X/3ZEJ1hD6wfo/bHTAOPlp+W+LzuZ1QTb3IU+YoNS1fP2/p2998ctXmw6pNeSeco0ML3aDzNBMUWeOtC+lIUs0TNqcY8bKkPTx93Hln9Ly8nzIk44w7R6Fn6txdEks9M+/kZnhoVcQeicV8o8f43t27Y9MVBKHDAAByRhdMk4veVARkq8sxpugG2Wzf8rHDMi31ke9dkXr0C/2GT9bs1ANA/apVe+xdZuxmnGBmL165edPcGRPezTb0U5sY99XUMF1JzUw2T4hfSEfVj992dMns6w+Hp4/pUQo1GAiAYZ15ocuqMIIzoeGUpxNrLjE0ZfuMacWdQOExFrEaAADyvbYRALjj6RdPUEnJ5N+bf6uBdaaYoqw4Qsg61/jTp49JPz7rlvWlXvsKAmTwE8s2/GzFite5nEDPi17jzcjxZDh87kn2tgSCAIhiz7jSLR2utSjq8Kn0hMWWEdN4FQDAJrBDebv9GEIIMlIGOW0I2f18nqIbMiEEMrpeyLHseW/R/eQnYzOVMyb+J8cyB2OG45e1tYTKOnQkKgGWouKQY6Hbg6aqGkn4wuZ0KESwpOgerS126NN8OaVZ1o8dbl9nqK6O1nVC3TC4sdOyLEQooTCZ0Bra26JBnfAKAIChQzFH0z3ejJ43a+IGhucO7W7Z8K2sQ+doR4Fos+W81u/OO0GLdegdoSrovto9e4ttblZfufIuFQBgX8ewWwryNN7lbjrWsdFZlFGUWHl5uYkQIhTjKLY73IalJorsnH4EAEDTSRGxLqwCQG9/7TUw2CuXLfvQlt2VDkwe4NxDRwgRRWZOiwuSmWQpZVENJ1c5YRVJnmQpaP3oKwbpeiY2jKFREwDAyxs7HHJapzoTWlOBM2UfN4z8DQAAEyNgs9MXVBodDoctm41qVml1WFahW2ZSsHupPsm+WXCGjRcCX+NotP/kKt9xO0aoIZ7Eke/fMTTB8O4riUr2AQC88eqbNMXoJXZR0xHhbOmWxpaaV1pEVuC9Vw3nLjiyjkTopmPNFpdV6ABgZzuPdvUFdBO6P7PWrq3jKYouDNrHHZ2x4nVOTmqTvIMcb9mduPRPH1ssZnHQO4hvBQAQ8vPsqmZZhmWkgAY9HC7XGhpJniIb0YrvFF1wppRhVZvdrlFZhd7cynQFAndeElWtx+XCwQjj1ooKpDt13zTE4y1zysfEEqnmRMOGgwWmTsz5vziZepbihfkIbIKNxZRd5BkAAMNID+Zo86LK6EQBDy8uciayBv311xs5T7490NMN1t5b6t2HTXQZo02D2bP8j0cCkqwOY+Pyq5Mnh+iGxsMHol2ZUaZpNfwjc4gYzq3pZsrGIE5TuBgAgKRYJSzLX/BRl6XP7XNoulZY7GYOZA36y1v3UwKXEFGOTkWcxvwLkYGkZIZihtrbFon8wulxvxIOTzXS6SCy2ZwspvE4QRRObSVyjBlEFNG60i2morV3nHQXzVKKg2MXOp9MqvMuAPggG4m/U9DpDgUfOnR4P7kEEs+raiN2VsAUGPo1CLPt4eljPgQA2LGjwhg8eNTtYKGhozxjDn6Wfsg4ZFVt5lmHGzCjAgAQgxSxGF8Q9IXVu4qkjDWUvX7Shqz66cmkQgIBf6CvVrphGFZV1ckC1aNNLQGaGIxO1Bv9JO/zm9CEwsIJRdPerKj47HVQol0chilaT6YcZmub2EYIwQSII8+e3+PaxbVrm3jJNO73upwvhbNYCUEDAFw5aRQxDSqvr1a612Pz/u0AsACgdqXgWoEXvs0S9ftz5gzq5sKueuLq0BfbRhMu0Ayrs8QvFzU1w5bK51vzTELkqpl5qfCsns1jX+TEdI7m36mcMbYx61nGIIC+o0NrA0IQ5Di3DADg9QqFaktcAABVVamvmyb39NrFX2s6n7bpDEYURU4QUzeBYkytSxqEMYn2NEc+Z+Gm6bzAts6fNf7tbOuLAQAqKsp0KW0mQlXvUn2x0i0LW/8yxZ2ZuWj/bUDR3rWLR9Seb9uupEaIZbUkU1qMRipuOppyIkSdtz0PrW3i5yyse8ghsl3zZ5W9mAt9T/nppSWmH7yy2BfQTc3Qnvn3Q9diojzidtK7etI2zyUOIUQ8QVGqNChIgk4HN5Lh+fOy56EVHxUr0cgihuJ2Vc3ODfBT5gUAQDch6fKUeuEC6sQv+kFKMZ5Ip3QfC2SzoSZ65DUwlGVLJdBx7yBawQAjLROCNgpt+bJ2v16w6etd8czXaYaqWV45sTGX+p6CzjIetS0iCbkGvuL1Rm7HRvM6iRcqTT15q8/t+6Qn7Qs9nWtLA+xRjrG1dyUy0xDh/BwvntW8zFmwYzBirH8zTCvq8VPzwxVlUq51pj97ICXaaAp8ALAvlxOINEqP2TnQedPgNGxA+IHSeI++tMqr//SPv3/+2EdbYxJd+dA3g8nlD3a/bvqCLaU0oG9rpuqnafw/z867+FPfFw3dkuUOmnWMyOXgM6s+mKWlpZ28zf1JV5f0uI215l5Mf0533laVyOpzf9lZ9cDSnc12kTUSnYZpYTVIY8Lpiv7a6oU3b+7rGPAUdEbk2zRFuSVXAz+yuH6OqiuJpx+76S+LV+/2YB1LT8+dsPVi+uyIxooZGv15qKf4d/s7Do1KqJZfVtTIkIDvlXmzruqAS0ROQU+4+Q5np2kL1RE6nOW69AcX7XoYsKU/+8TNzwEAzL3/mi4A+OvF9qsoRpHgZNunTy9IA8A2uETllMu4pqJMZ2ichG07/Vk1KU9snqXpWmTpI+Of6e2+RR5G8jzfAZe4dKt7UQz1UFw1xwBAr5+5DK1t4qW2zseA4L3LKsf9VzaU0XTijSZj+y916N0O78pqpp7BaGxvD7LgtwdKEifalzEc2pot4IQQihBWcHOZ9n4F/ffhb7QLLEtCq3YM7a0B5j6166aurug8ClDNk7PHZ+334qrWNHhYmpirw1PT/cq8AADwHFOnSuYPAeCiDtaF6gjdvmHr3apuDmMxnrek8vqsvs25LUJ7dEuKQT+Q094NEJ49vt5CWAut/HjKhXb6wJL6q2Mbtz8lsoxWXTl+7pLKG7L++mxNTRc4RF7rl9ABAOS08VtZ0qeFVn5yVU86W/i7hqJfL9jyINGtH3o94urqyvF/zFUZsmRCkQrmif4A/axbRQtqGq9MxtO/Bl1Zu+yxiTvP6QbO3zwMW8xdAGQII7CvLnvk2rdzrcgPHtozB7PG3hcWjV3fb6EDADz65NY8STPuZ0XekjXmDSkTP4oVyfQV+7yaTgoFxj48k0kMUlRLp2n6zZVVZdtRH2yCAAD86OE9T4GJ171QPWrPpQ79nG+f+NQWL3ikeue1GMh1qbh+ta4bBMUMLwKjg3Zw+1wFrr8+WzEiCgCwKtx3imCauJEgRGFAciNrmwj/w0d3r7n77jq+P8z3svj1lw9fPsLrKkkNGTJFG4CeKyUsI99hx1yuDzRkxab3F2ltlUUamW39Zb6XBXSnACUYs1a/uTMvB+jN7Uk6mdaPD0DPoXhc3lK7ncoMQM+hMAzO70xIxwag5zKspiAv3+vWB6DnSAghKJ1KS4lUuqO/zPly8F6w180XYdkhD6z0HMmaNYCTCT1y/XBrAHquZH96M02ILt1zzxB1AHquAqPkG2rQK6/GGBMYkAE5m/w/QHuDWhcmpxEAAAAASUVORK5CYII="
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

# -------------------------------------------------------------------------
# TEMA PREMIUM & SAKLAR TAMPILAN WEB | HP
# -------------------------------------------------------------------------
ux_boot()

# -------------------------------------------------------------------------
# SISTEM LOGIN & ROLE AKSES 
# -------------------------------------------------------------------------
if 'role' not in st.session_state: st.session_state['role'] = None
if 'dashboard_area' not in st.session_state: st.session_state['dashboard_area'] = None

# -------------------------------------------------------------------------
# SISTEM LOGIN & ROLE AKSES (MENU UTAMA INDEPENDEN)
# -------------------------------------------------------------------------
if 'role' not in st.session_state: st.session_state['role'] = None
if 'dashboard_area' not in st.session_state: st.session_state['dashboard_area'] = None

if st.session_state['role'] is None and st.session_state.get('active_menu') not in ["💠 SiSAVA", "🚨 SiHD (Pergerakan Tim)", "🗺️ Sysinfomap (Peta & Populasi Site)", "⚡ SiPower"]:
    
    ux_login_css()
    st.markdown("<br>", unsafe_allow_html=True)
    render_header_logo(is_login=True) 
    st.markdown("<h1 style='text-align: center; font-size: 40px; margin-top: 5px;'>Sistem Operasional Terpadu (SIDATA)</h1>", unsafe_allow_html=True)
    st.markdown("<p style='text-align: center; font-size: 18px; margin-bottom: 35px; letter-spacing: 1px; opacity: 0.85;'>Pilih Portal Akses Pekerjaan Anda Hari Ini</p>", unsafe_allow_html=True)
    
    # --- BARIS 1: LAPANGAN, LOGISTIK, & ADMIN ---
    col_l1, col_l2, col_l3 = st.columns(3)
    
    with col_l1:
        st.markdown("""
            <div class='login-box'>
                <div style='font-size: 45px; margin-bottom: 10px;'>👷‍♂️</div>
                <div class='login-title' style='font-size: 20px;'>Lapangan</div>
                <div class='login-desc' style='font-size: 13px;'>Akses Form Pendataan PM (SPS, Reguler, Pangkalan Bun).</div>
            </div>
        """, unsafe_allow_html=True)
        c_v1, c_v2 = st.columns(2)
        if c_v1.button("🔴 P'BUN", type="primary", use_container_width=True):
            st.session_state['role'], st.session_state['area_lapangan'], st.session_state['active_menu'] = 'Lapangan', 'Pangkalan Bun', None
            st.rerun()
        if c_v2.button("🔵 PLK", type="secondary", use_container_width=True):
            st.session_state['role'], st.session_state['area_lapangan'], st.session_state['active_menu'] = 'Lapangan', 'Palangkaraya', None
            st.rerun()

    with col_l2:
        st.markdown("""
            <div class='login-box' style='border-color: #FF9800;'>
                <div style='font-size: 45px; margin-bottom: 10px;'>📦</div>
                <div class='login-title' style='font-size: 20px;'>Logistik</div>
                <div class='login-desc' style='font-size: 13px;'>Akses stok, In/Out barang, dan Auto-WA Material.</div>
            </div>
        """, unsafe_allow_html=True)
        st.write("")
        if st.button("📦 MASUK LOGISTIK", type="secondary", use_container_width=True):
            st.session_state['role'], st.session_state['area_lapangan'], st.session_state['active_menu'] = 'Logistik', 'Logistik', None
            st.rerun()
            
    with col_l3:
        st.markdown("""
            <div class='login-box'>
                <div style='font-size: 45px; margin-bottom: 10px;'>🔐</div>
                <div class='login-title' style='font-size: 20px;'>Admin</div>
                <div class='login-desc' style='font-size: 13px;'>Akses edit data, Kurva S, Manajemen Akun & Pengaturan Sistem.</div>
            </div>
        """, unsafe_allow_html=True)
        admin_pass = st.text_input("Kata Sandi Admin:", type="password", placeholder="Sandi...", key="pwd_login", label_visibility="collapsed")
        if st.button("🔓 LOGIN ADMIN", type="primary", use_container_width=True):
            if admin_pass == "KUT2027":
                st.session_state['role'], st.session_state['active_menu'] = 'Admin', None
                st.rerun()
            else:
                st.error("❌ Sandi salah!")

    # --- BARIS 2: SiSAVA & SiHD ---
    st.markdown("<hr style='margin: 25px 0;'>", unsafe_allow_html=True)
    c_bot1, c_bot2 = st.columns([1, 1])
    
    with c_bot1:
        st.markdown("""
            <div class='area-card' style='border-top-color: #10B981;'>
                <div style='font-size: 40px; margin-bottom: 8px;'>💠</div>
                <div class='login-title' style='font-size: 18px;'>SiSAVA (Service Variable Activity)</div>
                <div class='login-desc' style='font-size: 13px;'>Pengajuan RAB, validasi Nilai SVA, dan Manajemen Keuangan Project. (Akun dikelola oleh Admin).</div>
            </div>
        """, unsafe_allow_html=True)
        if st.button("💠 MASUK SiSAVA", type="secondary", use_container_width=True):
            st.session_state['active_menu'] = "💠 SiSAVA"
            st.rerun()
            
    with c_bot2:
        st.markdown("""
            <div class='area-card' style='border-top-color: #EF4444;'>
                <div style='font-size: 40px; margin-bottom: 8px;'>🚨</div>
                <div class='login-title' style='font-size: 18px;'>SiHD (Sistem Helpdesk)</div>
                <div class='login-desc' style='font-size: 13px;'>Buku mutasi tim, rute harian, kendala lapangan, dan Auto-WA pergantian shift.</div>
            </div>
        """, unsafe_allow_html=True)
        if st.button("🚨 MASUK SiHD", type="primary", use_container_width=True):
            st.session_state['active_menu'] = "🚨 SiHD (Pergerakan Tim)"
            st.rerun()

    st.markdown("<br><br>", unsafe_allow_html=True)
    st.markdown("<div style='text-align: center; font-size: 13px; opacity: 0.7;'>Secured System Application & Database Management<br><b>Created By Okta Pradika © 2026</b></div>", unsafe_allow_html=True)
    st.stop()
# -------------------------------------------------------------------------
# 🗺️ PARSER KML SYSINFOMAP (IN-MEMORY CACHE — SUPER RINGAN) & MASTER DATA
# -------------------------------------------------------------------------
import xml.etree.ElementTree as _ET

KABUPATEN_MAP = {
    "PLK": "Kota Palangka Raya",
    "KKP": "Kab. Kapuas (Kuala Kapuas)",
    "KSN": "Kab. Katingan (Kasongan)",
    "KKN": "Kab. Gunung Mas (Kuala Kurun)",
    "MTW": "Kab. Barito Utara (Muara Teweh)",
    "PPS": "Kab. Pulang Pisau",
    "NTH": "Site N3T / Remote Area",
    "TML": "Kab. Barito Timur (Tamiang Layang)",
    "BNT": "Kab. Barito Selatan (Buntok)",
    "AMT": "Kab. Barito Selatan (Jenamas)",
    "PRC": "Kab. Murung Raya (Puruk Cahu)",
    "COH": "Site Combat / Khusus",
    "EPQ": "Site Ekspansi / iMacro New",
    "SPT": "Area Perbatasan (Sampit/Lamunti)"
}

@st.cache_data(ttl=3600, show_spinner=False)
def load_sysinfomap_kml(file_path_or_bytes, _stamp=0):
    try:
        if isinstance(file_path_or_bytes, bytes):
            root = _ET.fromstring(file_path_or_bytes)
        else:
            tree = _ET.parse(file_path_or_bytes)
            root = tree.getroot()
    except Exception:
        return pd.DataFrame()

    ns = {'kml': 'http://www.opengis.net/kml/2.2'}
    placemarks = root.findall('.//kml:Placemark', ns)
    if not placemarks:
        placemarks = root.findall('.//Placemark')

    rows = []
    for pm in placemarks:
        rec = {}
        for sd in pm.findall('.//{*}SimpleData'):
            col = sd.attrib.get('name', '')
            if col:
                rec[col] = (sd.text or '').strip()
        if not rec.get('site_id'):
            nm = pm.find('{*}name')
            if nm is not None and nm.text:
                rec['site_id'] = nm.text.strip()

        coord_el = pm.find('.//{*}coordinates')
        c_lon, c_lat = None, None
        if coord_el is not None and coord_el.text:
            parts = coord_el.text.strip().split(',')
            if len(parts) >= 2:
                try:
                    c_lon = float(parts[0])
                    c_lat = float(parts[1])
                except Exception:
                    pass

        lat_val = pd.to_numeric(rec.get('Lat'), errors='coerce')
        lon_val = pd.to_numeric(rec.get('Long'), errors='coerce')
        rec['lat'] = float(lat_val) if pd.notna(lat_val) else (c_lat if c_lat is not None else 0.0)
        rec['lon'] = float(lon_val) if pd.notna(lon_val) else (c_lon if c_lon is not None else 0.0)

        sid = str(rec.get('site_id', '')).strip().upper()
        rec['site_id'] = sid
        pref = sid[:3] if len(sid) >= 3 else "OTH"
        rec['Prefix'] = pref
        rec['Kabupaten'] = KABUPATEN_MAP.get(pref, f"Wilayah Lainnya ({pref})")
        for k_def in ['Sitename', 'Site_Owner', 'site_class', 'Status_site', 'Access', 'RemarkSite', 'BrandRecti', 'Transport_type', 'Type_Hub', 'Power_Source', 'Status_Genset', 'Case_Genset', 'RTS']:
            val_k = str(rec.get(k_def, '') or '').strip()
            rec[k_def] = val_k if val_k not in ['', '#N/A', 'None'] else 'Tidak Terdata'
        rows.append(rec)

    df = pd.DataFrame(rows)
    if not df.empty:
        df = df[(df['lat'] != 0) & (df['lon'] != 0)].reset_index(drop=True)
    return df

def get_default_kml_df():
    candidates = ["newlist_910.kml", "newlist.kml", "Sysinfomap.kml"]
    for c in candidates:
        if os.path.exists(c):
            stt = os.stat(c)
            return load_sysinfomap_kml(c, stt.st_mtime_ns), c
    import glob as _g
    kmls = sorted(_g.glob("*.kml"))
    if kmls:
        stt = os.stat(kmls[0])
        return load_sysinfomap_kml(kmls[0], stt.st_mtime_ns), kmls[0]
    return pd.DataFrame(), None

@st.cache_data(ttl=3600, show_spinner=False)
def load_master_lists():
    fme_plk, fme_pbu, site_plk, site_pbu = [], [], [], []
    try:
        df1 = pd.read_excel("NIK NOP PLK.xlsx")
        if 'NAMA' in df1.columns: fme_plk = df1['NAMA'].dropna().astype(str).tolist()
    except: pass
    
    try:
        df2 = pd.read_excel("NAMA NOP PANGKALANBUN.xlsx")
        if 'NAMA' in df2.columns: fme_pbu = df2['NAMA'].dropna().astype(str).tolist()
    except: pass
    
    try:
        df3 = pd.read_excel("Site ID Sidata Palangkaraya.xlsx")
        if 'Site ID' in df3.columns: site_plk = df3['Site ID'].dropna().astype(str).tolist()
    except: pass
    
    try:
        df4 = pd.read_excel("Site ID Sidata Pangkalanbun.xlsx")
        if 'Site ID' in df4.columns: site_pbu = df4['Site ID'].dropna().astype(str).tolist()
    except: pass

    try:
        df_cb = pd.read_excel("Combat NOP Palangkaraya_1.xlsx", sheet_name="Combat NOP Palangkaraya")
        if 'siteid' in df_cb.columns:
            site_plk.extend(df_cb['siteid'].dropna().astype(str).str.strip().tolist())
    except: pass

    try:
        df_kml_m, _ = get_default_kml_df()
        if not df_kml_m.empty and 'site_id' in df_kml_m.columns:
            kml_formatted = [f"{r['site_id']} - {r['Sitename']}" for _, r in df_kml_m.iterrows()]
            site_plk.extend(kml_formatted)
    except: pass

    return sorted(list(set(fme_plk))), sorted(list(set(fme_pbu))), sorted(list(set(site_plk))), sorted(list(set(site_pbu)))
# -------------------------------------------------------------------------
# MULTI-CLOUDINARY & OPTIMASI GAMBAR (AUTO LOAD-BALANCING & COMPRESSION)
# -------------------------------------------------------------------------
import random

# Tiga Akun Cloudinary Bapak
CLOUDINARY_ACCOUNTS = [
    {
        "cloud_name": "fxm61tjv",
        "api_key": "624877324969231",
        "api_secret": "LIFO6pfEg9fOM3nbsY8FBbVTpSI" # Sesuai codingan lama yg jalan
    },
    {
        "cloud_name": "ji9i1hei",
        "api_key": "277523771179144",
        "api_secret": "7eg20WV-Bzf9DXEoT8rg3H99V6c"
    },
    {
        "cloud_name": "tqrnzopw", # GANTI INI DENGAN CLOUD NAME DARI AKUN 48219...
        "api_key": "482193747671566",
        "api_secret": "GMMzzU6--yoY46u6bS45hs5Mp90"
    }
]

def upload_image(file_obj, folder_name="solar_bts_healthcheck"):
    if file_obj is not None:
        try:
            # Pilih akun secara acak (Random) agar beban / kapasitas terbagi rata
            selected_account = random.choice(CLOUDINARY_ACCOUNTS)
            cloudinary.config(
                cloud_name = selected_account["cloud_name"],
                api_key = selected_account["api_key"],
                api_secret = selected_account["api_secret"],
                secure = True
            )
            
            # Optimasi Kualitas Gambar:
            # Tidak sampai buram, hanya membatasi lebar maksimal layar laptop (1200px)
            # dan kualitas diset ke auto-eco (Tetap tajam namun hemat MB)
            response = cloudinary.uploader.upload(
                file_obj.getvalue(), 
                folder=folder_name,
                transformation=[
                    {'width': 1200, 'crop': 'limit'}, 
                    {'quality': 'auto:eco', 'fetch_format': 'auto'}
                ]
            )
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
            selected_account = random.choice(CLOUDINARY_ACCOUNTS)
            cloudinary.config(
                cloud_name = selected_account["cloud_name"],
                api_key = selected_account["api_key"],
                api_secret = selected_account["api_secret"],
                secure = True
            )
            
            response = cloudinary.uploader.upload(
                file_obj.getvalue(), 
                folder=folder_name, 
                resource_type="raw", 
                public_id=file_obj.name
            )
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
            # OPTIMASI ANTI-LEMOT: Minta Cloudinary mengirim foto resolusi kecil (lebar 400px) untuk tampilan grid dashboard
            opt_u = u.replace("upload/", "upload/w_400,q_auto:eco/") if "cloudinary" in u else u
            img_html += f'<div style="background-color: var(--secondary-background-color); border: 1px solid var(--primary-color); border-radius: 8px; padding: 6px; text-align: center;"><a href="{u}" target="_blank"><img src="{opt_u}" loading="lazy" style="max-width: 100%; height: auto; max-height: 400px; object-fit: contain; border-radius: 6px;"></a></div>'
        img_html += '</div>'
        st.markdown(img_html, unsafe_allow_html=True)
# -------------------------------------------------------------------------
# SETUP DATABASE GOOGLE SHEETS
# -------------------------------------------------------------------------
SHEET_ID = "1HvgVicTWwO4RMQI6ZR3Mu3IgGicwjcLZl9mDN1auvJU"

def connect_gsheets(sheet_name="Report Preventive"):
    try:
        creds_json = st.secrets["gcp_json"]
        creds_dict = json.loads(creds_json)
        scopes = ["https://www.googleapis.com/auth/spreadsheets", "https://www.googleapis.com/auth/drive"]
        creds = Credentials.from_service_account_info(creds_dict, scopes=scopes)
        client = gspread.authorize(creds)
        return client.open_by_key(SHEET_ID).worksheet(sheet_name)
    except Exception as e: return None

@st.cache_data(ttl=600) 
def fetch_data_from_gsheets():
    all_data = []
    
    sheet_prev = connect_gsheets("Report Preventive")
    if sheet_prev:
        data = sheet_prev.get_all_values()
        for idx, row in enumerate(data[1:]):
            if len(row) >= 6: 
                try: 
                    item = json.loads(row[5])
                    if item.get('kategori') == 'SiUPDATE': item['kategori'] = 'SPS'
                    item['_sheet'] = "Report Preventive"
                    item['_row'] = idx + 2
                    all_data.append(item)
                except: pass
                
    sheet_pbu = connect_gsheets("Report PBU")
    if sheet_pbu:
        data_pbu = sheet_pbu.get_all_values()
        for idx, row in enumerate(data_pbu[1:]):
            if len(row) >= 6: 
                try: 
                    item = json.loads(row[5])
                    item['_sheet'] = "Report PBU"
                    item['_row'] = idx + 2
                    all_data.append(item)
                except: pass

    sheet_proj = connect_gsheets("Project")
    if sheet_proj:
        data_proj = sheet_proj.get_all_values()
        for idx, row in enumerate(data_proj[1:]):
            if len(row) >= 6: 
                try: 
                    item = json.loads(row[5])
                    item['_sheet'] = "Project"
                    item['_row'] = idx + 2
                    all_data.append(item)
                except: pass
                
    # --- TAMBAHAN UNTUK SHEET SiHD ---
    sheet_sihd = connect_gsheets("SiHD")
    if sheet_sihd:
        data_sihd = sheet_sihd.get_all_values()
        for idx, row in enumerate(data_sihd[1:]):
            if len(row) >= 6: 
                try: 
                    item = json.loads(row[5])
                    item['_sheet'] = "SiHD"
                    item['_row'] = idx + 2
                    all_data.append(item)
                except: pass
                
    return all_data
if 'laporan_db' not in st.session_state: 
    st.session_state['laporan_db'] = fetch_data_from_gsheets()

# -------------------------------------------------------------------------
# WA GENERATOR PANGKALAN BUN (Otomatis Teks)
# -------------------------------------------------------------------------
def generate_wa_pbu(r):
    text = f"*SITE VISIT REPORT*\n\n"
    text += f"> INFORMASI SITE\n"
    text += f"- FME NAME : {r.get('pb_te', '-')}\n"
    text += f"- Site ID : {r.get('pb_site_id', '-')}\n"
    text += f"- Site Name : {r.get('site_name', '-')}\n"
    text += f"- Ticket# : {r.get('pb_ticket', '-')}\n"
    text += f"- Tanggal On Site : {r.get('timestamp', '-').split(' ')[0]}\n\n"
    text += f"> Catuan Power Utama : {r.get('pb_power', '-')}\n\n"
    
    rects = r.get('pb_rectifiers', [])
    text += f"> DATA RECTIFIER\n"
    text += f"- Jumlah Rectifier: {len(rects)}\n\n"
    
    for idx, rect in enumerate(rects):
        text += f"Rectifier #{idx+1}\n"
        text += f"- Tipe Rectifier : {rect.get('tipe', '-')}\n"
        text += f"- Type Controler : {rect.get('controler', '-')}\n"
        text += f"- Type Module : {rect.get('module', '-')}\n"
        text += f"- Jumlah Battery : {rect.get('jml_batt', '-')}\n"
        text += f"- Jenis Batt OK: {rect.get('batt_ok', '-')}\n"
        text += f"- Jenis Batt NOK: {rect.get('batt_nok', '-')}\n"
        text += f"- Jumlah Module : (OK : {rect.get('mod_ok', '-')}; NOK : {rect.get('mod_nok', '-')})\n"
        text += f"- BCL : {rect.get('bcl', '-')}\n"
        text += f"- Load Current: {rect.get('load', '-')}\n"
        text += f"- DC Voltage:  {rect.get('dc_v', '-')}\n"
        text += f"- LVD Prio:  {rect.get('lvd_p', '-')}\n"
        text += f"- LVD Non Prio : {rect.get('lvd_np', '-')}\n"
        text += f"- EQP LVD Prio: {rect.get('eqp_p', '-')}\n"
        text += f"- EQP LVD Non Prio: {rect.get('eqp_np', '-')}\n"
        text += f"- Battery Backup Time : {rect.get('backup', '-')}\n\n"

    en = r.get('pb_enva', {})
    text += f"> * ENVA Setelah Simulasi (Bukan Cabut Kabel)\n"
    text += f"- Main Fail : {en.get('main_fail', '-')}\n"
    text += f"- Rectifier Fail : {en.get('rect_fail', '-')}\n"
    text += f"- Low Batt : {en.get('low_batt', '-')}\n"
    text += f"- High Temp : {en.get('high_temp', '-')}\n"
    text += f"- Genset Run : {en.get('gen_run', '-')}\n"
    text += f"- Genset Fail : {en.get('gen_fail', '-')}\n\n"

    kel = r.get('pb_kelistrikan', {})
    text += f"> DATA KELISTRIKAN\n"
    text += f"- Daya KWH: {kel.get('daya', '-')}\n"
    text += f"- Jumlah Phasa: {kel.get('phasa', '-')}\n"
    text += f"- Tegangan Pengukuran R-S-T terhadap N: {kel.get('v_rst_n', '-')}\n"
    text += f"- Tegangan Pengukuran Phasa RS-RT-ST: {kel.get('v_rs_rt_st', '-')}\n"
    text += f"- Pengukuran Arus R-S-T: {kel.get('arus_rst', '-')}\n"
    text += f"- Tegangan Pengukuran G-N: {kel.get('v_gn', '-')}\n\n"

    acp = r.get('pb_acpdb', {})
    text += f"> DATA ACPDB\n"
    text += f"- Ada ACPDB: {acp.get('ada', '-')}\n"
    text += f"- Tegangan Pengukuran R-S-T terhadap N: {acp.get('v_rst_n', '-')}\n"
    text += f"- Tegangan Pengukuran Phasa RS-RT-ST: {acp.get('v_rs_rt_st', '-')}\n"
    text += f"- Pengukuran Arus R-S-T: {acp.get('arus_rst', '-')}\n"
    text += f"- Tegangan Pengukuran G-N: {acp.get('v_gn', '-')}\n\n"

    gen = r.get('pb_genset', {})
    text += f"> DATA GENSET\n"
    text += f"- Ada Genset: {gen.get('ada', '-')}\n"
    text += f"- Capacity Genset : {gen.get('cap', '-')}\n"
    text += f"- Genset Status : {gen.get('status', '-')}\n"
    text += f"- Config Running : {gen.get('config', '-')}\n"
    text += f"- Kapasitas Tangki : {gen.get('tangki', '-')}\n"
    text += f"- Running Hour : {gen.get('rh', '-')}\n"
    text += f"- Sisa BBM : {gen.get('bbm', '-')}\n\n"

    text += f"> Activity Report\n"
    text += f"1. Power & CME Activity\n- " + ("\n- ".join(r.get('pb_act_power', [])) if isinstance(r.get('pb_act_power'), list) else str(r.get('pb_act_power','-'))) + "\n\n"
    text += f"2. Transport Activity\n- " + ("\n- ".join(r.get('pb_act_trans', [])) if isinstance(r.get('pb_act_trans'), list) else str(r.get('pb_act_trans','-'))) + "\n\n"
    text += f"3. OSA Impact Service Activity\n- " + ("\n- ".join(r.get('pb_act_osa', [])) if isinstance(r.get('pb_act_osa'), list) else str(r.get('pb_act_osa','-'))) + "\n\n"

    text += f"> Temuan/ Catatan :\n"
    temuan_lines = str(r.get('action', '-')).split('\n')
    for line in temuan_lines:
        if line.strip(): text += f"{line}\n"
        
    text += f"\n\n🔗 Foto & Dokumentasi lengkap tersimpan di Dashboard SPS Portal."
    return text# =============================================================================
# 💠 INTERNAL FUNGSI SiSAVA — Sistem Service Variable Activity
# =============================================================================
import hashlib as _hl
import hmac as _hmac
import glob as _glob
import io as _io
import secrets as _pysecrets
import datetime as _dtm

SV_MENU = "💠 SiSAVA"
SV_AREAS = ["KALIMANTAN", "SULAWESI", "PUMA"]
SV_NOPS = ["PALANGKA RAYA", "PANGKALAN BUN"]
SV_PRICELIST_PATTERNS = ["Summary*Final*Price*.xlsx", "*Shoping*List*.xlsx", "*Shopping*List*.xlsx"]
SV_LOCAL_DB = "sisava_local_db.json"        
SV_ALLOW_SELF_APPROVE = False               

# Parameter hitung sesuai dengan logika Bapak
# - f_mat: Harga KUT = Harga 3E x f_mat
# - hpp_pct: HPP = Harga KUT x hpp_pct
SV_PARAMS_DEFAULT = dict(
    f_mat=0.92, f_jasa=0.85, hpp_pct=0.85, fee_3e=0.0, fee_kut=0.0, biaya_lain=0.0,
    com_rate=0.075, com_bulan=3.0, pph=0.0285
)
SV_JASA_KW = ("jasa", "instalasi", "pemasangan", "pekerjaan", "perbaikan", "relokasi", "rebalancing", "sewa", "mobilisasi",
              "penggantian", "pengecatan", "overhaul", "overhoul", "deinstalasi", "patroli", "pengurusan", "pembayaran", "servis",
              "service", "pembersihan", "pengukuran", "commissioning", "survey", "retrofit", "penarikan", "perawatan",
              "maintenance", "pengelasan", "galian", "pengiriman", "management fee", "transport")

S_RAB_SPV = "RAB: Menunggu SPV"
S_RAB_MGR = "RAB: Menunggu Manager"
S_RAB_TOLAK = "RAB: Ditolak (Perlu Revisi)"
S_RAB_OK = "RAB: Disetujui (Isi Nilai SVA)"
S_SVA = "SVA: Terisi (Pekerjaan Berjalan)"
S_KRJ_SPV = "Pekerjaan: Menunggu SPV"
S_KRJ_MGR = "Pekerjaan: Menunggu Manager"
S_KRJ_TOLAK = "Pekerjaan: Ditolak (Perbaiki Upload)"
S_DONE = "SAH / Selesai"
SV_STATUS_COLOR = {S_RAB_SPV: "#F59E0B", S_RAB_MGR: "#F59E0B", S_RAB_TOLAK: "#E11D48", S_RAB_OK: "#2563EB", S_SVA: "#8B5CF6",
                   S_KRJ_SPV: "#F59E0B", S_KRJ_MGR: "#F59E0B", S_KRJ_TOLAK: "#E11D48", S_DONE: "#059669"}
SV_PIPELINE = [S_RAB_SPV, S_RAB_MGR, S_RAB_OK, S_SVA, S_KRJ_SPV, S_KRJ_MGR, S_DONE]

SV_ROLES = {"STAFF": "Staff / Tim Operation", "SPV": "SPV Operation", "MANAGER": "Manager", "ENGINEERING": "Tim Engineering"}
SV_PERMS = {
    "STAFF":   {"rab_create", "sva_fill", "kerja_upload"},
    "ENGINEERING": {"rab_create", "sva_fill", "kerja_upload"},
    "SPV":     {"rab_create", "sva_fill", "kerja_upload", "approve_spv", "view_all"},
    "MANAGER": {"rab_create", "sva_fill", "kerja_upload", "approve_spv", "approve_mgr", "view_all", "manage_users", "pricelist", "rekap"},
}
SV_URGENSI = ["Normal", "Penting", "Urgent", "Emergency"]
SV_STATUS_TIKET = ["Open", "On Progress", "Menunggu Approval Telkomsel", "Closed"]

SV_RAB_COLS = ["rab_id", "ticket_no", "nop", "area", "kota", "tgl_pengajuan", "dibuat_oleh", "nama_pembuat", "jabatan_pembuat", "status",
               "tot_tsel", "tot_rev", "tot_hpp", "margin", "gpm", "net", "payload", "evidence_rab",
               "spv_oleh", "spv_nama", "spv_jabatan", "spv_at", "spv_catatan", "mgr_oleh", "mgr_nama", "mgr_jabatan", "mgr_at", "mgr_catatan",
               "sva", "sva_oleh", "sva_at", "kerja_catatan", "kerja_urls", "kerja_oleh", "kerja_at",
               "f_spv_oleh", "f_spv_at", "f_spv_catatan", "f_mgr_oleh", "f_mgr_at", "f_mgr_catatan", "revisi", "log", "updated_at"]
SV_USER_COLS = ["username", "nama", "jabatan", "role", "pw_hash", "aktif", "dibuat"]

def _sv_num(x):
    try:
        v = float(x)
        return 0.0 if v != v else v
    except Exception: return 0.0

def sv_rp(x): return "Rp {:,.0f}".format(_sv_num(x))
_BULAN_ID = ["Januari", "Februari", "Maret", "April", "Mei", "Juni", "Juli", "Agustus", "September", "Oktober", "November", "Desember"]
def sv_tgl_id(d):
    if isinstance(d, str): d = _dtm.date.fromisoformat(d[:10])
    return "%d %s %d" % (d.day, _BULAN_ID[d.month - 1], d.year)
def sv_now(): return _dtm.datetime.now().strftime("%Y-%m-%d %H:%M:%S")
def sv_guess_jenis(desc): return "Jasa" if any(k in str(desc).lower() for k in SV_JASA_KW) else "Material"

def sv_find_pricelist():
    import os as _os
    bases = [_os.getcwd()]
    for b in bases:
        for pat in SV_PRICELIST_PATTERNS:
            hits = sorted(_glob.glob(_os.path.join(b, pat)))
            hits = [h for h in hits if "~$" not in h]
            if hits: return hits[0]
    return None

def sv_parse_pricelist(df):
    items, cat = [], "Umum"
    n = len(df)
    
    col_desc = 1
    col_uom = 2
    col_kali = 5
    
    for c in range(df.shape[1]):
        val = str(df.iat[0, c]).upper()
        if "DESCRIP" in val or "DESKRIPSI" in val: col_desc = c
        elif "UOM" in val or "SATUAN" in val: col_uom = c
        elif "KALIMANTAN" in val: col_kali = c

    for r in range(3, n):
        def g(c):
            try: v = df.iat[r, c]
            except: return None
            return None if (v is None or (isinstance(v, float) and v != v)) else v
            
        a = g(0)
        b = g(col_desc)
        
        if b is None or str(b).strip() == "":
            if isinstance(a, str) and a.strip() not in ("", "-", "xxx", "NO"): cat = a.strip()
            continue
            
        price, ket = {}, ""
        v = g(col_kali)
        if isinstance(v, (int, float)) and not isinstance(v, bool): price["KALIMANTAN"] = float(v)
        elif isinstance(v, str) and v.strip() not in ("", "-"): ket = v.strip()
            
        desc = " ".join(str(b).split())
        items.append(dict(uid=f"{a if a is not None else 'x'}-{r}", no=a, cat=cat, desc=desc, uom=str(g(col_uom) or "").strip(),
                          price=price, ket=ket, jenis=sv_guess_jenis(desc)))
    return items

@st.cache_data(show_spinner=False, max_entries=4)
def _sv_load_pricelist(path, stamp):
    try: df = read_excel_cached(path, sheet_name="List (2)", header=None)
    except: 
        try: df = read_excel_cached(path, sheet_name="Sheet1", header=None)
        except: df = pd.read_excel(path, header=None)
    return sv_parse_pricelist(df)

def sv_pricelist():
    import os as _os
    p = sv_find_pricelist()
    if not p: return [], None
    stt = _os.stat(p)
    return _sv_load_pricelist(p, (stt.st_mtime_ns, stt.st_size)), p

def sv_params(p=None):
    d = dict(SV_PARAMS_DEFAULT)
    if p: d.update({k: _sv_num(v) for k, v in p.items() if k in d})
    if d["f_mat"] <= 0: d["f_mat"] = SV_PARAMS_DEFAULT["f_mat"]
    return d

def sv_line_calc(line, p):
    qty, ht = _sv_num(line.get("qty")), _sv_num(line.get("harga_tsel")) # Harga dasar = Tagihan 3E ke Tsel
    hk = line.get("harga_kut") # Tagihan KUT ke 3E
    
    if hk is None or hk == "":
        hk = ht * p["f_mat"] if line.get("jenis") == "Material" else ht * p["f_jasa"]
    hk = _sv_num(hk)
    
    hp = line.get("hpp") # HPP Lapangan
    hp = hk * p["hpp_pct"] if (hp is None or hp == "") else _sv_num(hp)
    
    sub_t = ht * qty; fee_t = sub_t * p["fee_3e"]
    sub_k = hk * qty; fee_k = sub_k * p["fee_kut"]
    
    return dict(qty=qty, harga_tsel=ht, harga_kut=hk, hpp_unit=hp, sub_tsel=sub_t, fee_tsel=fee_t, tot_tsel=sub_t + fee_t,
                sub_kut=sub_k, fee_kut=fee_k, rev=sub_k + fee_k, hpp=hp * qty)

def sv_calc(rab):
    p = sv_params(rab.get("params"))
    out_sites, T = [], dict(tsel=0.0, rev=0.0, hpp=0.0, other_rev=0.0)
    for s in rab.get("sites", []):
        blk = {"site_id": s.get("site_id", ""), "site_name": s.get("site_name", ""), "A": [], "B": [], "sum": {}}
        for sec in ("A", "B"):
            for ln in s.get("items", []):
                if ln.get("sec", "A") != sec: continue
                c = sv_line_calc(ln, p)
                c.update(desc=ln.get("desc", ""), uom=ln.get("uom", ""), jenis=ln.get("jenis", "Material"))
                blk[sec].append(c)
        for sec in ("A", "B"):
            blk["sum"][sec] = {k: sum(c[k] for c in blk[sec]) for k in ("tot_tsel", "rev", "hpp")}
        T["tsel"] += blk["sum"]["A"]["tot_tsel"] + blk["sum"]["B"]["tot_tsel"]
        T["rev"] += blk["sum"]["A"]["rev"] + blk["sum"]["B"]["rev"]
        T["hpp"] += blk["sum"]["A"]["hpp"] + blk["sum"]["B"]["hpp"]
        T["other_rev"] += blk["sum"]["B"]["rev"]
        out_sites.append(blk)
    hpp_all = T["hpp"] + p["biaya_lain"]
    margin = T["rev"] - hpp_all
    gpm = (margin / T["rev"]) if T["rev"] else 0.0
    com = margin * p["com_rate"] * p["com_bulan"] / 12.0
    pph = margin * p["pph"]
    return dict(params=p, sites=out_sites, tot_tsel=T["tsel"], tot_rev=T["rev"], tot_hpp_items=T["hpp"], biaya_lain=p["biaya_lain"],
                tot_hpp=hpp_all, margin=margin, gpm=gpm, com=com, pph=pph, net=margin - com - pph, other_rev=T["other_rev"])

def _sv_prep_image(raw, maxw=1000):
    from PIL import ImageOps
    im = Image.open(_io.BytesIO(raw))
    try: im = ImageOps.exif_transpose(im)
    except: pass
    im = im.convert("RGB")
    if im.width > maxw: im = im.resize((maxw, int(im.height * maxw / im.width)))
    buf = _io.BytesIO()
    im.save(buf, "JPEG", quality=82)
    return buf.getvalue(), im.width, im.height

def sv_build_xlsx(rab, calc, meta=None, images=None):
    import math as _m
    from openpyxl import Workbook
    from openpyxl.styles import Font, PatternFill, Border, Side, Alignment
    from openpyxl.drawing.image import Image as _XLImage
    from openpyxl.workbook.properties import CalcProperties
    meta, images, p = meta or {}, images or [], calc["params"]
    wb = Workbook(); ws = wb.active; ws.title = "RAB"
    ws.sheet_view.showGridLines = False
    BLUE, BLUE_L, PINK, PINK_L, GRAY = "4874CB", "DAE3F5", "E54C5E", "F5B7BF", "D9D9D9"
    thin, med = Side(style="thin"), Side(style="medium")
    MONEY, PCT, QTY = '#,##0;-#,##0;"-"', '0.0%;-0.0%;""', "General"
    def put(ref, v=None, bold=False, size=12, color="000000", fill=None, fmt=None, h=None, vv="center", wrap=False, italic=False, brd=True, ul=False):
        c = ws[ref]
        if v is not None: c.value = v
        c.font = Font(name="Calibri", size=size, bold=bold, color=color, italic=italic, underline="single" if ul else None)
        if fill: c.fill = PatternFill("solid", fgColor=fill)
        if fmt: c.number_format = fmt
        c.alignment = Alignment(horizontal=h, vertical=vv, wrap_text=wrap)
        if brd: c.border = Border(left=thin, right=thin, top=thin, bottom=thin)
        return c
    widths = dict(A=4.6, B=10.1, C=20.7, D=13, E=3.7, F=39.4, G=56.1, H=11.6, I=5.4, J=12, K=5.3, L=10.1, M=12.5, N=11.5, O=6.3, P=11, Q=5.9, R=11, S=12.7, T=12.1, U=13.3)
    for k, v in widths.items(): ws.column_dimensions[k].width = v
    cols = "ABCDEFGHIJKLMNOPQRSTU"
    put("A1", "RENCANA ANGGARAN BIAYA", bold=True, size=20, brd=False); ws.row_dimensions[1].height = 26.25
    put("A3", "NOP", bold=True, size=16, brd=False); put("C3", ": " + str(rab.get("nop", "")), bold=True, size=16, brd=False)
    put("A4", "No. Ticket", bold=True, size=14, brd=False); put("C4", ": " + str(rab.get("ticket", "")), bold=True, size=14, brd=False)
    if meta.get("rab_id"):
        put("N3", "No. Dokumen", size=10, italic=True, brd=False); put("Q3", ": " + meta["rab_id"], size=10, italic=True, brd=False)
        put("N4", "Status", size=10, italic=True, brd=False); put("Q4", ": " + str(meta.get("status", "")), size=10, italic=True, brd=False)
    heads = ["NO", "SITE ID", "SITE NAME", "PEKERJAAN", "SL", "Pekerjaan", "SL", "HARGA 3E (Tsel)", "QTY", "Sub Total", "Fee", "Nilai Fee", "TOTAL",
             "HARGA KUT (3E)", "QTY", "Sub Total", "Fee", "Nilai Fee", "TOTAL REVENUE", "NILAI HPP", "TOTAL HPP"]
    for i, t in enumerate(heads):
        put("%s5" % cols[i], t, bold=True, color="FFFFFF", fill=(BLUE if i < 13 else PINK), h="center", wrap=True, size=12 if i not in (0, 4) else 11)
    ws.row_dimensions[5].height = 40
    r, hdr_rows = 6, []
    for idx, (site, cs) in enumerate(zip(rab.get("sites", []), calc["sites"]), 1):
        start = r
        for sec, title in (("A", "Jasa dan material with SL"), ("B", "Jasa dan material (Other Material ME)")):
            hdr = r; hdr_rows.append(hdr)
            put("E%d" % r, sec, bold=True, fill=BLUE_L, h="center"); put("F%d" % r, title, bold=True, fill=BLUE_L, h="left")
            for c in "GHIJKL": put("%s%d" % (c, r), None, fill=BLUE_L)
            for c in "NOPQR": put("%s%d" % (c, r), None, fill=PINK_L)
            put("T%d" % r, None, fill=PINK_L)
            lines = cs[sec]; r += 1
            first = r
            for ln in lines:
                put("F%d" % r, "Pengadaan material" if ln["jenis"] == "Material" else "Jasa pekerjaan", wrap=True, h="left")
                put("G%d" % r, ln["desc"], wrap=True, h="left")
                put("H%d" % r, round(ln["harga_tsel"], 2), fmt=MONEY); put("I%d" % r, ln["qty"], fmt=QTY, h="center")
                put("J%d" % r, "=H%d*I%d" % (r, r), fmt=MONEY); put("K%d" % r, p["fee_3e"], fmt=PCT, h="center")
                put("L%d" % r, "=J%d*K%d" % (r, r), fmt=MONEY); put("M%d" % r, "=J%d+L%d" % (r, r), fmt=MONEY)
                put("N%d" % r, round(ln["harga_kut"], 2), fmt=MONEY); put("O%d" % r, "=I%d" % r, fmt=QTY, h="center")
                put("P%d" % r, "=N%d*O%d" % (r, r), fmt=MONEY); put("Q%d" % r, p["fee_kut"], fmt=PCT, h="center")
                put("R%d" % r, "=P%d*Q%d" % (r, r), fmt=MONEY); put("S%d" % r, "=P%d+R%d" % (r, r), fmt=MONEY)
                put("T%d" % r, round(ln["hpp_unit"], 2), fmt=MONEY); put("U%d" % r, "=T%d*O%d" % (r, r), fmt=MONEY)
                ws.row_dimensions[r].height = 25 if len(ln["desc"]) <= 60 else 40
                r += 1
            last = r - 1
            if lines:
                put("M%d" % hdr, "=SUM(M%d:M%d)" % (first, last), bold=True, fill=BLUE_L, fmt=MONEY)
                put("S%d" % hdr, "=SUM(S%d:S%d)" % (first, last), bold=True, fill=PINK_L, fmt=MONEY)
                put("U%d" % hdr, "=SUM(U%d:U%d)" % (first, last), bold=True, fill=PINK_L, fmt=MONEY)
            else:
                if sec == "B": put("G%d" % hdr, "Other Material ME", bold=True, fill=BLUE_L, h="left")
                put("M%d" % hdr, 0, bold=True, fill=BLUE_L, fmt=MONEY); put("S%d" % hdr, 0, bold=True, fill=PINK_L, fmt=MONEY)
                put("U%d" % hdr, 0, bold=True, fill=PINK_L, fmt=MONEY)
            ws.row_dimensions[hdr].height = 25
            if sec == "A":                         
                for c in cols[4:]: put("%s%d" % (c, r), None)
                ws.row_dimensions[r].height = 12
                r += 1
        end = r - 1
        for col, val in (("A", idx), ("B", site.get("site_id", "")), ("C", site.get("site_name", "")), ("D", site.get("pekerjaan", ""))):
            ws.merge_cells("%s%d:%s%d" % (col, start, col, end))
            put("%s%d" % (col, start), val, bold=(col in "ABC"), h="center", vv="center" if col != "D" else "top", wrap=True, size=12 if col != "D" else 11)
            for rr in range(start, end + 1): ws["%s%d" % (col, rr)].border = Border(left=thin, right=thin, top=thin if rr == start else None, bottom=thin if rr == end else None)
    
    r_lain = r
    for c in cols[:13]: put("%s%d" % (c, r_lain), None)
    for c in cols[13:]: put("%s%d" % (c, r_lain), None, fill=PINK_L)
    put("T%d" % r_lain, "Biaya lain-lain", fill=PINK_L, h="right", size=10); put("U%d" % r_lain, p["biaya_lain"], fill=PINK_L, fmt=MONEY)
    rt = r_lain + 1
    for c in cols: put("%s%d" % (c, rt), None, bold=True, fill=GRAY)
    put("L%d" % rt, "TOTAL", bold=True, fill=GRAY, h="center")
    put("M%d" % rt, ("=" + "+".join("M%d" % h for h in hdr_rows)) if hdr_rows else 0, bold=True, fill=GRAY, fmt=MONEY)
    ws.merge_cells("N%d:R%d" % (rt, rt)); put("N%d" % rt, "TOTAL", bold=True, fill=GRAY, h="center")
    put("S%d" % rt, ("=" + "+".join("S%d" % h for h in hdr_rows)) if hdr_rows else 0, bold=True, fill=GRAY, fmt=MONEY)
    put("U%d" % rt, ("=" + "+".join(["U%d" % h for h in hdr_rows] + ["U%d" % r_lain])), bold=True, fill=GRAY, fmt=MONEY)
    ws.row_dimensions[rt].height = 35
    def pc(x): return ("%g" % x).replace(".", ",")
    labels = [("Margin", None, "=S%d-U%d" % (rt, rt), MONEY, True),
              ("GPM", None, "=IF(S%d=0,0,U%d/S%d)" % (rt, rt + 1, rt), "0.0%", False),
              ("Cost of Money (%s%% x %s bulan/12 Bulan)" % (pc(p["com_rate"] * 100), pc(p["com_bulan"])), p["com_rate"] * p["com_bulan"] / 12.0, "=U%d*T%d" % (rt + 1, rt + 3), MONEY, False),
              ("PPh", p["pph"], "=T%d*U%d" % (rt + 4, rt + 1), MONEY, False),
              ("Margin Bersih (setelah Cost of Money & PPh)", None, "=U%d-U%d-U%d" % (rt + 1, rt + 3, rt + 4), MONEY, True)]
    for k, (lab, rate, form, fmt, bold) in enumerate(labels, 1):
        rr = rt + k
        for c in cols: put("%s%d" % (c, rr), None, brd=(c in "NOPQRSTU"))
        ws.merge_cells("N%d:S%d" % (rr, rr)); put("N%d" % rr, lab, bold=bold, h="left", wrap=True)
        if rate is not None: put("T%d" % rr, rate, fmt="0.000%", h="center")
        else: put("T%d" % rr, None)
        put("U%d" % rr, form, bold=bold, fmt=fmt)
        ws.row_dimensions[rr].height = 25
    ws.merge_cells("A%d:M%d" % (rt + 1, rt + 4))
    put("A%d" % (rt + 1), "Lampiran Referensi harga (gambar terlampir di bawah)", italic=True, h="center", size=11)
    r = rt + 6
    put("A%d" % r, "Noted: untuk pekerjaan Other material mohon diberikan evidence Refrensi harga nya", italic=True, size=11, brd=False)
    r += 3
    put("A%d" % r, "%s, %s" % (rab.get("kota", ""), sv_tgl_id(rab.get("tanggal") or _dtm.date.today())), size=12, brd=False)
    r += 1
    for col, t in (("A", "Diajukan Oleh"), ("H", "Diketahui Oleh"), ("S", "Disetujui Oleh")): put("%s%d" % (col, r), t, brd=False)
    r += 5
    sig = [("A", meta.get("pembuat_nama") or "(Pemohon)", meta.get("pembuat_jabatan") or ""),
           ("H", meta.get("spv_nama") or "( menunggu persetujuan SPV )", meta.get("spv_jabatan") or "SPV Operation"),
           ("S", meta.get("mgr_nama") or "( menunggu persetujuan Manager )", meta.get("mgr_jabatan") or "Manager")]
    for col, nm, jb in sig:
        put("%s%d" % (col, r), nm, bold=True, brd=False, ul=True); put("%s%d" % (col, r + 1), jb, size=11, brd=False)
    last_text_row = r + 1
    r += 4
    put("A%d" % r, "Lampiran Referensi harga", bold=True, brd=False)
    r += 2
    for raw in images:
        try:
            data, w, h = _sv_prep_image(raw)
            xi = _XLImage(_io.BytesIO(data)); xi.width, xi.height = w, h
            ws.add_image(xi, "B%d" % r)
            r += int(_m.ceil(h / 20.0)) + 2
        except Exception: continue
    ws.print_area = "A1:U%d" % max(last_text_row, r)
    ws.page_setup.orientation = "landscape"; ws.page_setup.fitToWidth = 1; ws.page_setup.fitToHeight = 0
    ws.sheet_properties.pageSetUpPr.fitToPage = True
    wb.calculation = CalcProperties(fullCalcOnLoad=True)
    buf = _io.BytesIO(); wb.save(buf)
    return buf.getvalue()

def sv_rab_filename(rab, rab_id=""):
    import re as _re
    base = "RAB_%s_%s" % (rab.get("ticket") or rab_id or "SVA", (rab.get("sites") or [{}])[0].get("site_name", ""))
    return _re.sub(r"[^A-Za-z0-9_.-]+", "_", base).strip("_")[:120] + ".xlsx"

def _sv_json(o):
    return json.dumps(o, ensure_ascii=False, separators=(",", ":"), default=str)

def _sv_jload(s, default=None):
    try: return json.loads(s) if s else default
    except Exception: return default

class _SvLocalStore:
    kind = "lokal"
    def __init__(self, path=None): self.path = path or SV_LOCAL_DB
    def _read(self):
        try:
            with open(self.path, "r", encoding="utf-8") as f: return json.load(f)
        except Exception: return {"rab": [], "users": []}
    def _write(self, d):
        tmp = self.path + ".tmp"
        with open(tmp, "w", encoding="utf-8") as f: json.dump(d, f, ensure_ascii=False)
        os.replace(tmp, self.path)
    def list_rabs(self): return self._read()["rab"]
    def list_users(self): return self._read()["users"]
    def add_rab(self, rec):
        d = self._read(); d["rab"].append({k: rec.get(k, "") for k in SV_RAB_COLS}); self._write(d)
    def update_rab(self, rab_id, fields):
        d = self._read()
        for r in d["rab"]:
            if r.get("rab_id") == rab_id:
                r.update({k: v for k, v in fields.items() if k in SV_RAB_COLS}); self._write(d); return True
        return False
    def upsert_user(self, u):
        d = self._read()
        for r in d["users"]:
            if r.get("username") == u["username"]:
                r.update(u); self._write(d); return
        d["users"].append({k: u.get(k, "") for k in SV_USER_COLS}); self._write(d)

@st.cache_resource(show_spinner=False, ttl=1800)
def _sv_spreadsheet():
    creds = Credentials.from_service_account_info(json.loads(st.secrets["gcp_json"]),
                                                  scopes=["https://www.googleapis.com/auth/spreadsheets", "https://www.googleapis.com/auth/drive"])
    return gspread.authorize(creds).open_by_key(SHEET_ID)

@st.cache_resource(show_spinner=False, ttl=1800)
def _sv_worksheet(name, cols_key):
    sh = _sv_spreadsheet()
    cols = list(cols_key)
    try: ws = sh.worksheet(name)
    except Exception:
        ws = sh.add_worksheet(title=name, rows=500, cols=max(len(cols), 10) + 2)
        ws.update(range_name="A1", values=[cols])
        return ws
    head = ws.row_values(1)
    if head != cols[:len(head)] or len(head) < len(cols):     
        ws.update(range_name="A1", values=[head + [c for c in cols if c not in head]])
    return ws

@st.cache_data(ttl=12, show_spinner=False)
def _sv_rows(name, cols_key):
    return _sv_worksheet(name, cols_key).get_all_records()

class _SvSheetsStore:
    kind = "Google Sheets"
    def _cols(self, name): return tuple(SV_RAB_COLS if name == "SISAVA_RAB" else SV_USER_COLS)
    def _rows(self, name): return _sv_rows(name, self._cols(name))
    def list_rabs(self): return self._rows("SISAVA_RAB")
    def list_users(self): return self._rows("SISAVA_USERS")
    def _row_values(self, ws, rec, cols):
        head = ws.row_values(1)
        return head, [str(rec.get(c, "")) if rec.get(c, "") is not None else "" for c in head]
    def add_rab(self, rec):
        ws = _sv_worksheet("SISAVA_RAB", self._cols("SISAVA_RAB"))
        head, vals = self._row_values(ws, rec, SV_RAB_COLS)
        ws.append_row(vals, value_input_option="RAW"); _sv_rows.clear()
    def update_rab(self, rab_id, fields):
        ws = _sv_worksheet("SISAVA_RAB", self._cols("SISAVA_RAB"))
        try: cell = ws.find(rab_id, in_column=1)
        except Exception: cell = None
        if not cell: return False
        head = ws.row_values(1); cur = ws.row_values(cell.row)
        cur += [""] * (len(head) - len(cur))
        for k, v in fields.items():
            if k in head:
                cur[head.index(k)] = "" if v is None else str(v)
        ws.update(range_name="A%d" % cell.row, values=[cur], value_input_option="RAW"); _sv_rows.clear()
        return True
    def upsert_user(self, u):
        ws = _sv_worksheet("SISAVA_USERS", self._cols("SISAVA_USERS"))
        try: cell = ws.find(u["username"], in_column=1)
        except Exception: cell = None
        head = ws.row_values(1)
        if cell:                                  
            cur = ws.row_values(cell.row); cur += [""] * (len(head) - len(cur))
            row = dict(zip(head, cur)); row.update({k: v for k, v in u.items() if k in head})
            ws.update(range_name="A%d" % cell.row, values=[[str(row.get(c, "")) for c in head]], value_input_option="RAW")
        else:
            ws.append_row([str(u.get(c, "")) for c in head], value_input_option="RAW")
        _sv_rows.clear()

def sv_store():
    try:
        _sv_spreadsheet()
        return _SvSheetsStore()
    except Exception:
        _sv_spreadsheet.clear()
        return _SvLocalStore()

def sv_hash_pw(pw, salt=None):
    salt = salt or _pysecrets.token_hex(8)
    dk = _hl.pbkdf2_hmac("sha256", pw.encode("utf-8"), salt.encode(), 120000)
    return "pbkdf2$%s$%s" % (salt, dk.hex())

def sv_check_pw(pw, stored):
    try:
        _, salt, hx = str(stored).split("$")
        return _hmac.compare_digest(sv_hash_pw(pw, salt).split("$")[2], hx)
    except Exception: return False

def sv_can(user, perm):
    return bool(user) and perm in SV_PERMS.get(user.get("role"), set())

def sv_user_find(store, username):
    u = str(username).strip().lower()
    for r in store.list_users():
        if str(r.get("username", "")).strip().lower() == u: return r
    return None

def sv_login(store, username, pw):
    r = sv_user_find(store, username)
    if not r or str(r.get("aktif", "1")).strip() in ("0", "False", "false", "nonaktif") or not sv_check_pw(pw, r.get("pw_hash", "")):
        return None
    return {"username": r["username"], "nama": r.get("nama", r["username"]), "jabatan": r.get("jabatan", ""), "role": r.get("role", "STAFF")}

def sv_validate_new_user(u, pw, existing=None):
    if not str(u.get("username", "")).strip() or not str(u.get("nama", "")).strip(): return "Username dan nama wajib diisi."
    if any(ch.isspace() for ch in u["username"]): return "Username tidak boleh mengandung spasi."
    if u.get("role") not in SV_ROLES: return "Role tidak valid."
    if len(pw) < 8: return "Sandi minimal 8 karakter."
    return ""

def sv_log_add(rec, user, aksi, catatan=""):
    log = _sv_jload(rec.get("log"), []) or []
    log.append({"ts": sv_now(), "oleh": user["username"], "nama": user.get("nama", ""), "role": user["role"], "aksi": aksi, "catatan": catatan})
    return log

def sv_next_id(recs, year=None):
    year = year or _dtm.date.today().year
    nums = []
    for r in recs:
        parts = str(r.get("rab_id", "")).split("-")
        if len(parts) == 3 and parts[1] == str(year) and parts[2].isdigit(): nums.append(int(parts[2]))
    return "RAB-%d-%04d" % (year, (max(nums) if nums else 0) + 1)

def sv_transition(rec, action, user, note="", data=None):
    st_, role, uname = rec.get("status"), user.get("role"), user.get("username")
    creator = rec.get("dibuat_oleh")
    def deny(m): return False, m, {}
    need_note = action in ("spv_reject", "mgr_reject", "f_spv_reject", "f_mgr_reject")
    if need_note and not str(note).strip(): return deny("Catatan wajib diisi saat menolak.")
    if action in ("spv_approve", "spv_reject", "f_spv_approve", "f_spv_reject") and not sv_can(user, "approve_spv"): return deny("Akun Anda tidak berwenang sebagai SPV.")
    if action in ("mgr_approve", "mgr_reject", "f_mgr_approve", "f_mgr_reject") and not sv_can(user, "approve_mgr"): return deny("Hanya Manager yang dapat melakukan persetujuan ini.")
    if action in ("spv_approve", "spv_reject", "mgr_approve", "mgr_reject", "f_spv_approve", "f_spv_reject", "f_mgr_approve", "f_mgr_reject") and uname == creator and not SV_ALLOW_SELF_APPROVE: return deny("Anda tidak dapat menyetujui/menolak dokumen yang Anda ajukan sendiri.")
    f = {"updated_at": sv_now()}
    ok_msg = ""
    if action == "spv_approve" and st_ == S_RAB_SPV:
        f.update(status=S_RAB_MGR, spv_oleh=uname, spv_nama=user["nama"], spv_jabatan=user.get("jabatan", ""), spv_at=sv_now(), spv_catatan=note); ok_msg = "RAB diteruskan ke Manager."
    elif action == "spv_reject" and st_ == S_RAB_SPV:
        f.update(status=S_RAB_TOLAK, spv_oleh=uname, spv_nama=user["nama"], spv_at=sv_now(), spv_catatan=note); ok_msg = "RAB dikembalikan untuk revisi."
    elif action == "mgr_approve" and st_ in (S_RAB_MGR, S_RAB_SPV):
        if st_ == S_RAB_SPV: f.update(spv_oleh=uname, spv_nama=user["nama"], spv_jabatan=user.get("jabatan", ""), spv_at=sv_now(), spv_catatan="(disetujui langsung oleh Manager) " + note)
        f.update(status=S_RAB_OK, mgr_oleh=uname, mgr_nama=user["nama"], mgr_jabatan=user.get("jabatan", ""), mgr_at=sv_now(), mgr_catatan=note); ok_msg = "RAB DISETUJUI. Silakan isi Nilai SVA."
    elif action == "mgr_reject" and st_ in (S_RAB_MGR, S_RAB_SPV):
        f.update(status=S_RAB_TOLAK, mgr_oleh=uname, mgr_nama=user["nama"], mgr_at=sv_now(), mgr_catatan=note); ok_msg = "RAB dikembalikan untuk revisi."
    elif action == "sva_fill" and st_ == S_RAB_OK:
        if not sv_can(user, "sva_fill"): return deny("Akun Anda tidak berwenang mengisi Nilai SVA.")
        d = data or {}
        miss = [k for k in ("ticket_no", "status_tiket", "urgensi", "keperluan", "tgl_rencana") if not str(d.get(k, "")).strip()]
        if miss: return deny("Lengkapi: " + ", ".join(miss))
        if _sv_num(d.get("nilai_sva")) <= 0 or _sv_num(d.get("keperluan_dana")) <= 0: return deny("Nilai SVA dan Keperluan Dana harus lebih dari 0.")
        f.update(status=S_SVA, sva=_sv_json(d), sva_oleh=uname, sva_at=sv_now()); ok_msg = "Nilai SVA tersimpan. Pekerjaan dapat berjalan."
    elif action == "kerja_upload" and st_ in (S_SVA, S_KRJ_TOLAK):
        if not sv_can(user, "kerja_upload"): return deny("Akun Anda tidak berwenang mengunggah pekerjaan.")
        d = data or {}
        if not d.get("urls"): return deny("Unggah minimal 1 foto/berkas bukti pekerjaan.")
        f.update(status=S_KRJ_SPV, kerja_urls=_sv_json(d["urls"]), kerja_catatan=note, kerja_oleh=uname, kerja_at=sv_now(),
                 f_spv_oleh="", f_spv_at="", f_spv_catatan="", f_mgr_oleh="", f_mgr_at="", f_mgr_catatan=""); ok_msg = "Bukti pekerjaan terkirim, menunggu SPV."
    elif action == "f_spv_approve" and st_ == S_KRJ_SPV:
        f.update(status=S_KRJ_MGR, f_spv_oleh=uname, f_spv_at=sv_now(), f_spv_catatan=note); ok_msg = "Diteruskan ke Manager."
    elif action == "f_spv_reject" and st_ == S_KRJ_SPV:
        f.update(status=S_KRJ_TOLAK, f_spv_oleh=uname, f_spv_at=sv_now(), f_spv_catatan=note); ok_msg = "Pekerjaan dikembalikan untuk diperbaiki."
    elif action == "f_mgr_approve" and st_ in (S_KRJ_MGR, S_KRJ_SPV):
        if st_ == S_KRJ_SPV: f.update(f_spv_oleh=uname, f_spv_at=sv_now(), f_spv_catatan="(disetujui langsung oleh Manager) " + note)
        f.update(status=S_DONE, f_mgr_oleh=uname, f_mgr_at=sv_now(), f_mgr_catatan=note); ok_msg = "Pekerjaan SAH."
    elif action == "f_mgr_reject" and st_ in (S_KRJ_MGR, S_KRJ_SPV):
        f.update(status=S_KRJ_TOLAK, f_mgr_oleh=uname, f_mgr_at=sv_now(), f_mgr_catatan=note); ok_msg = "Pekerjaan dikembalikan untuk diperbaiki."
    else: return deny("Aksi tidak sesuai dengan status dokumen saat ini (%s)." % st_)
    f["log"] = _sv_json(sv_log_add(rec, user, action, note))
    return True, ok_msg, f

def sv_rab_to_record(rab, calc, user, rab_id, evidence_urls, kota=None):
    return {"rab_id": rab_id, "ticket_no": rab.get("ticket", ""), "nop": rab.get("nop", ""), "area": rab.get("area", ""), "kota": rab.get("kota", ""),
            "tgl_pengajuan": sv_now(), "dibuat_oleh": user["username"], "nama_pembuat": user["nama"], "jabatan_pembuat": user.get("jabatan", ""),
            "status": S_RAB_SPV, "tot_tsel": round(calc["tot_tsel"], 2), "tot_rev": round(calc["tot_rev"], 2), "tot_hpp": round(calc["tot_hpp"], 2),
            "margin": round(calc["margin"], 2), "gpm": round(calc["gpm"], 6), "net": round(calc["net"], 2), "payload": _sv_json(rab),
            "evidence_rab": _sv_json(evidence_urls), "revisi": 0, "updated_at": sv_now()}

def sv_visible(recs, user):
    return list(recs) if sv_can(user, "view_all") else [r for r in recs if r.get("dibuat_oleh") == user.get("username")]

def _sv_badge(status):
    return '<span style="display:inline-block;padding:3px 11px;border-radius:99px;font-size:.74rem;font-weight:800;color:#fff;background:%s">%s</span>' % (
        SV_STATUS_COLOR.get(status, "#64748B"), _esc(status))

def _sv_flash():
    m = st.session_state.pop("sv_flash", None)
    if m: (st.success if m[0] != "!" else st.error)(m.lstrip("!"))

@st.cache_data(ttl=3600, show_spinner=False, max_entries=80)
def _sv_fetch_img(url):
    r = requests.get(url, timeout=10)
    r.raise_for_status(); return r.content

@st.cache_data(ttl=300, show_spinner=False, max_entries=20)
def _sv_xlsx_cached(rab_id, stamp, payload, meta_json, ev_json):
    rab = _sv_jload(payload, {}) or {}
    meta = _sv_jload(meta_json, {}) or {}
    imgs = []
    for u in (_sv_jload(ev_json, []) or [])[:6]:
        try: imgs.append(_sv_fetch_img(u))
        except: pass
    return sv_build_xlsx(rab, sv_calc(rab), meta, imgs)

def sv_record_xlsx(rec):
    meta = {"rab_id": rec.get("rab_id"), "status": rec.get("status"), "pembuat_nama": rec.get("nama_pembuat"), "pembuat_jabatan": rec.get("jabatan_pembuat"),
            "spv_nama": rec.get("spv_nama"), "spv_jabatan": rec.get("spv_jabatan"), "mgr_nama": rec.get("mgr_nama"), "mgr_jabatan": rec.get("mgr_jabatan")}
    return _sv_xlsx_cached(rec.get("rab_id"), str(rec.get("updated_at")), rec.get("payload", ""), _sv_json(meta), rec.get("evidence_rab", ""))

def _sv_autodownload(data, name):
    import base64 as _b64
    safe = "".join(ch for ch in name if ch.isalnum() or ch in "._-")
    js = ("<script>(function(){try{var b=atob('%s');var a=new Uint8Array(b.length);for(var i=0;i<b.length;i++)a[i]=b.charCodeAt(i);"
          "var u=URL.createObjectURL(new Blob([a],{type:'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet'}));"
          "var l=document.createElement('a');l.href=u;l.download='%s';document.body.appendChild(l);l.click();"
          "setTimeout(function(){URL.revokeObjectURL(u);l.remove();},3000);}catch(e){}})();</script>") % (_b64.b64encode(data).decode(), safe)
    st.components.v1.html(js, height=0)

def _sv_new_site():
    return {"uid": _pysecrets.token_hex(3), "site_id": "", "site_name": "", "pekerjaan": "", "items": []}

def _sv_form_reset():
    ss = st.session_state
    for k in [k for k in list(ss.keys()) if str(k).startswith("svw_") or str(k).startswith("svl_")]: ss.pop(k, None)
    ss["sv_form"] = {"nop": SV_NOPS[0], "ticket": "", "area": "KALIMANTAN", "kota": "Balikpapan", "tanggal": _dtm.date.today().isoformat(),
                     "params": dict(SV_PARAMS_DEFAULT), "sites": [_sv_new_site()], "evidence": [], "catatan": "", "edit_id": None, "old_urls": [], "ver": 0}

def _sv_form_load(rec):
    _sv_form_reset()
    rab = _sv_jload(rec.get("payload"), {}) or {}
    f = st.session_state["sv_form"]
    f.update(nop=rab.get("nop", f["nop"]), ticket=rab.get("ticket", ""), area=rab.get("area", f["area"]), kota=rab.get("kota", f["kota"]),
             tanggal=rab.get("tanggal", f["tanggal"]), edit_id=rec["rab_id"], old_urls=_sv_jload(rec.get("evidence_rab"), []) or [])
    f["params"] = sv_params(rab.get("params"))
    f["sites"] = []
    for s in rab.get("sites", []):
        s = dict(s); s["uid"] = _pysecrets.token_hex(3); s.setdefault("items", [])
        for ln in s["items"]: ln.setdefault("sec", "A"); ln.setdefault("harga_tsel", None); ln.setdefault("hpp", None)
        f["sites"].append(s)
    f["sites"] = f["sites"] or [_sv_new_site()]
    st.session_state["sv_nav"] = "➕ Pengajuan RAB"
    st.session_state["sv_flash"] = "Data %s dimuat ke form. Perbaiki lalu ajukan ulang." % rec["rab_id"]

def _sv_rab_from_form(f):
    return {"nop": f["nop"], "ticket": f["ticket"].strip(), "area": f["area"], "kota": f["kota"], "tanggal": f["tanggal"],
            "params": dict(f["params"]), "sites": [{k: s[k] for k in ("site_id", "site_name", "pekerjaan", "items")} for s in f["sites"]]}

def _sv_add_a(uid):
    ss = st.session_state; f = ss["sv_form"]
    s = next((x for x in f["sites"] if x["uid"] == uid), None)
    item = (ss.get("svl_%s" % uid) or {}).get(ss.get("svw_pick_%s" % uid))
    if not s or not item: return
    price = item["price"].get(f["area"])
    s["items"].append({"sec": "A", "jenis": item["jenis"], "desc": item["desc"], "uom": item["uom"], "qty": _sv_num(ss.get("svw_qty_%s" % uid, 1)) or 1.0,
                       "harga_tsel": float(price or 0.0), "harga_kut": None, "hpp": None, "ref": item["uid"]})
    f["ver"] += 1
    if price is None: ss["sv_flash"] = "!Item ini tidak punya harga tetap (%s). Isi harga di tabel." % (item["ket"] or "harga manual")

def _sv_add_b(uid):
    ss = st.session_state; f = ss["sv_form"]
    s = next((x for x in f["sites"] if x["uid"] == uid), None)
    desc = str(ss.get("svw_bd_%s" % uid, "")).strip()
    if not s or not desc: ss["sv_flash"] = "!Isi uraian Other Material terlebih dahulu."; return
    hpp = _sv_num(ss.get("svw_bh_%s" % uid, 0))
    s["items"].append({"sec": "B", "jenis": ss.get("svw_bj_%s" % uid, "Material"), "desc": desc, "uom": str(ss.get("svw_bu_%s" % uid, "")).strip(),
                       "qty": _sv_num(ss.get("svw_bq_%s" % uid, 1)) or 1.0, "harga_tsel": _sv_num(ss.get("svw_bt_%s" % uid, 0)),
                       "harga_kut": _sv_num(ss.get("svw_bk_%s" % uid, 0)), "hpp": hpp if hpp > 0 else None})
    f["ver"] += 1

def _sv_del_site(uid):
    f = st.session_state["sv_form"]
    if len(f["sites"]) > 1: f["sites"] = [s for s in f["sites"] if s["uid"] != uid]; f["ver"] += 1

def _sv_reset_auto(uid):
    f = st.session_state["sv_form"]
    for s in f["sites"]:
        if s["uid"] == uid:
            for ln in s["items"]: ln["harga_kut"] = None; ln["hpp"] = None
    f["ver"] += 1

def _sv_line_editor(f, s, sec, p):
    items, uid = s["items"], s["uid"]
    idxs = [i for i, ln in enumerate(items) if ln.get("sec", "A") == sec]
    if not idxs: st.caption("Belum ada item pada bagian ini."); return
    rows = []
    for i in idxs:
        ln = items[i]; c = sv_line_calc(ln, p)
        rows.append({"Jenis": ln["jenis"], "Uraian": ln["desc"], "UoM": ln.get("uom", ""), "Qty": float(c["qty"]), "Harga 3E→Tsel": round(c["harga_tsel"], 2),
                     "Harga KUT→3E": round(c["harga_kut"], 2), "HPP Lapangan": round(c["hpp_unit"], 2), "Hapus": False})
    df = pd.DataFrame(rows)
    ed = st.data_editor(df, key="svw_ed_%s_%s_%d" % (uid, sec, f["ver"]), hide_index=True, use_container_width=True, num_rows="fixed",
                        disabled=["Uraian", "UoM"],
                        column_config={"Jenis": st.column_config.SelectboxColumn("Jenis", options=["Material", "Jasa"], required=True, width="small"),
                                       "Uraian": st.column_config.TextColumn("Uraian", width="large"),
                                       "Qty": st.column_config.NumberColumn("Qty", min_value=0.0, step=1.0, format="%.2f", width="small"),
                                       "Harga 3E→Tsel": st.column_config.NumberColumn("Harga 3E→Tsel", min_value=0.0, format="%.2f"),
                                       "Harga KUT→3E": st.column_config.NumberColumn("Harga KUT→3E", min_value=0.0, format="%.2f"),
                                       "HPP Lapangan": st.column_config.NumberColumn("HPP Lapangan", min_value=0.0, format="%.2f"),
                                       "Hapus": st.column_config.CheckboxColumn("Hapus", width="small")})
    drop = []
    for k, i in enumerate(idxs):
        er, orig, ln = ed.iloc[k], rows[k], items[i]
        if bool(er["Hapus"]): drop.append(i); continue
        ln["jenis"] = er["Jenis"]; ln["qty"] = _sv_num(er["Qty"]); ln["harga_tsel"] = _sv_num(er["Harga 3E→Tsel"])
        if abs(_sv_num(er["Harga KUT→3E"]) - orig["Harga KUT→3E"]) > 0.005: ln["harga_kut"] = _sv_num(er["Harga KUT→3E"])
        if abs(_sv_num(er["HPP Lapangan"]) - orig["HPP Lapangan"]) > 0.005: ln["hpp"] = _sv_num(er["HPP Lapangan"])
    if drop: s["items"] = [ln for j, ln in enumerate(items) if j not in drop]; f["ver"] += 1; st.rerun()

def _sv_picker_a(f, s, pl):
    uid, ss = s["uid"], st.session_state
    if not pl: st.info("Daftar Shopping List belum terbaca. Gunakan Other Material (bagian B), atau pastikan file Excel Shopping List ada di folder aplikasi."); return
    c1, c2 = st.columns([1, 2])
    cat = c1.selectbox("Kategori", ["Semua kategori"] + sorted({i["cat"] for i in pl}), key="svw_cat_%s" % uid)
    q = c2.text_input("Cari uraian", key="svw_q_%s" % uid, placeholder="mis. kabel twist 4 x 25")
    words = q.lower().split()
    cand = [i for i in pl if (cat == "Semua kategori" or i["cat"] == cat) and all(w in i["desc"].lower() for w in words)]
    shown, labels = cand[:150], {}
    for i in shown:
        harga = sv_rp(i["price"][f["area"]]) if f["area"] in i["price"] else (i["ket"] or "harga manual")
        lab = "%s · %s · %s" % (i["desc"][:90], i["uom"] or "-", harga)
        labels[lab if lab not in labels else lab + " #" + i["uid"]] = i
    ss["svl_%s" % uid] = labels
    st.selectbox("Pilih item — %d ditemukan%s" % (len(cand), " (menampilkan 150 pertama, persempit pencarian)" if len(cand) > 150 else ""),
                 list(labels) or ["(tidak ada hasil)"], key="svw_pick_%s" % uid)
    c3, c4 = st.columns(2)
    c3.number_input("Qty", min_value=0.0, value=1.0, step=1.0, key="svw_qty_%s" % uid)
    c4.markdown("<div style='height:1.9rem'></div>", unsafe_allow_html=True)
    c4.button("➕ Tambah ke RAB", key="svw_add_%s" % uid, on_click=_sv_add_a, args=(uid,), use_container_width=True)

def _sv_picker_b(s):
    uid = s["uid"]
    with st.expander("➕ Tambah Other Material / pekerjaan di luar Shopping List", expanded=False):
        st.text_input("Uraian", key="svw_bd_%s" % uid, placeholder="mis. Kabel NYY 4x10 (harga pasar)")
        c1, c2, c3 = st.columns(3)
        c1.selectbox("Jenis", ["Material", "Jasa"], key="svw_bj_%s" % uid)
        c2.text_input("UoM", key="svw_bu_%s" % uid, placeholder="meter / unit")
        c3.number_input("Qty", min_value=0.0, value=1.0, step=1.0, key="svw_bq_%s" % uid)
        c4, c5, c6 = st.columns(3)
        c4.number_input("Harga 3E→Tsel", min_value=0.0, value=0.0, step=1000.0, key="svw_bt_%s" % uid)
        c5.number_input("Harga KUT→3E (0 = otomatis)", min_value=0.0, value=0.0, step=1000.0, key="svw_bk_%s" % uid)
        c6.number_input("HPP Lapangan (0 = otomatis)", min_value=0.0, value=0.0, step=1000.0, key="svw_bh_%s" % uid)
        st.button("➕ Tambah Other Material", key="svw_addb_%s" % uid, on_click=_sv_add_b, args=(uid,), use_container_width=True)

def _sv_step_info(f, pl_path):
    st.markdown("##### 1️⃣ Informasi RAB")
    c1, c2, c3 = st.columns(3)
    nops = SV_NOPS + ([f["nop"]] if f["nop"] not in SV_NOPS else [])
    f["nop"] = c1.selectbox("NOP", nops, index=nops.index(f["nop"]), key="svw_nop")
    f["ticket"] = c2.text_input("No. Ticket SVA", value=f["ticket"], key="svw_ticket", placeholder="SVA-2026-000000006886")
    f["area"] = c3.selectbox("Area harga Shopping List", SV_AREAS, index=SV_AREAS.index(f["area"]) if f["area"] in SV_AREAS else 0, key="svw_area")
    c4, c5 = st.columns(2)
    f["kota"] = c4.text_input("Kota (untuk tanda tangan)", value=f["kota"], key="svw_kota")
    f["tanggal"] = c5.date_input("Tanggal RAB", value=_dtm.date.fromisoformat(f["tanggal"][:10]), key="svw_tgl").isoformat()
    if f.get("edit_id"): st.info("✏️ Mode REVISI untuk %s. Setelah diajukan, RAB kembali ke SPV." % f["edit_id"])
    if not pl_path: st.warning("File Shopping List tidak ditemukan. Letakkan file Excel di folder yang sama dengan App.py.")
    with st.expander("⚙️ Parameter perhitungan (RAB KUT & HPP)", expanded=False):
        p = f["params"]
        a, b, c = st.columns(3)
        p["f_mat"] = a.number_input("Faktor Material (Hrg KUT = Hrg 3E × x)", 0.10, 2.0, float(p["f_mat"]), 0.01, format="%.3f", key="svw_pfm")
        p["f_jasa"] = b.number_input("Faktor Jasa (Hrg KUT = Hrg 3E × x)", 0.10, 3.0, float(p["f_jasa"]), 0.01, format="%.3f", key="svw_pfj")
        p["hpp_pct"] = c.number_input("HPP Lapangan (% dari Hrg KUT)", 0.0, 200.0, float(p["hpp_pct"]) * 100, 0.5, key="svw_php") / 100.0
        d, e, g = st.columns(3)
        p["fee_3e"] = d.number_input("Fee 3E→Telkomsel (%)", 0.0, 100.0, float(p["fee_3e"]) * 100, 0.5, key="svw_pf3") / 100.0
        p["fee_kut"] = e.number_input("Fee KUT→3E (%)", 0.0, 100.0, float(p["fee_kut"]) * 100, 0.5, key="svw_pfk") / 100.0
        p["biaya_lain"] = g.number_input("Biaya lain-lain (Rp)", 0.0, 1e12, float(p["biaya_lain"]), 500.0, format="%.0f", key="svw_pbl")
        h, i, j = st.columns(3)
        p["com_rate"] = h.number_input("Cost of Money (% per tahun)", 0.0, 100.0, float(p["com_rate"]) * 100, 0.5, key="svw_pcr") / 100.0
        p["com_bulan"] = i.number_input("Cost of Money: jumlah bulan", 0.0, 36.0, float(p["com_bulan"]), 1.0, key="svw_pcb")
        p["pph"] = j.number_input("PPh (% dari margin)", 0.0, 100.0, float(p["pph"]) * 100, 0.05, format="%.2f", key="svw_ppp") / 100.0
        st.caption("Rumus: Harga 3E didapat dari Shopping List. Harga KUT dihitung otomatis berdasarkan Faktor Material (0.92) / Jasa (0.85). HPP dihitung dari persentase KUT (85%). Bisa diedit manual per item di tabel.")

def _sv_step_sites(f, pl):
    st.markdown("##### 2️⃣ Site, Pekerjaan & Item")
    p = sv_params(f["params"])
    for n, s in enumerate(list(f["sites"]), 1):
        uid = s["uid"]
        with st.expander("📍 Site %d" % n, expanded=True):
            c1, c2 = st.columns([1, 2])
            s["site_id"] = c1.text_input("Site ID", value=s["site_id"], key="svw_sid_%s" % uid, placeholder="PRC025").strip().upper()
            s["site_name"] = c2.text_input("Nama Site", value=s["site_name"], key="svw_snm_%s" % uid, placeholder="TUMBANGMASAO").strip().upper()
            s["pekerjaan"] = st.text_area("Pekerjaan", value=s["pekerjaan"], key="svw_job_%s" % uid, height=70, placeholder="mis. Vandalisme cable power BTS and Hybrid SPS and PLN")
            st.markdown("**A. Jasa dan material dengan Shopping List**")
            _sv_picker_a(f, s, pl)
            _sv_line_editor(f, s, "A", p)
            st.markdown("**B. Other Material ME** (di luar Shopping List — wajib lampirkan referensi harga)")
            _sv_picker_b(s)
            _sv_line_editor(f, s, "B", p)
            cc1, cc2 = st.columns(2)
            cc1.button("↺ Kembalikan Harga KUT & HPP ke Otomatis", key="svw_rst_%s" % uid, on_click=_sv_reset_auto, args=(uid,), use_container_width=True)
            if len(f["sites"]) > 1: cc2.button("🗑️ Hapus site ini", key="svw_dels_%s" % uid, on_click=_sv_del_site, args=(uid,), use_container_width=True)
    if st.button("➕ Tambah Site", key="svw_adds", use_container_width=True):
        f["sites"].append(_sv_new_site()); f["ver"] += 1; st.rerun()

def _sv_rab_problems(rab, calc, f, n_files):
    errs = []
    if not str(rab["ticket"]).strip(): errs.append("No. Ticket SVA wajib diisi.")
    if not rab["sites"] or any(not (s["site_id"] and s["site_name"]) for s in rab["sites"]): errs.append("Setiap site wajib punya Site ID dan Nama Site.")
    if any(not s["pekerjaan"].strip() for s in rab["sites"]): errs.append("Isi kolom Pekerjaan pada setiap site.")
    if any(not s["items"] for s in rab["sites"]): errs.append("Setiap site harus memiliki minimal 1 item.")
    if any(_sv_num(ln.get("qty")) <= 0 or _sv_num(ln.get("harga_tsel")) <= 0 for s in rab["sites"] for ln in s["items"]): errs.append("Qty dan Harga Tsel semua item harus lebih dari 0.")
    if calc["other_rev"] > 0 and (n_files + len(f.get("old_urls") or [])) == 0: errs.append("Ada Other Material: lampirkan evidence referensi harga (minimal 1 gambar).")
    return errs

def _sv_step_hitung(f, calc):
    st.markdown("##### 3️⃣ Hitungan Revenue & Margin")
    c1, c2, c3 = st.columns(3)
    c1.metric("Tagihan 3E → Telkomsel", sv_rp(calc["tot_tsel"]))
    c2.metric("Tagihan KUT → 3E", sv_rp(calc["tot_rev"]))
    c3.metric("Total HPP Lapangan (+ biaya lain)", sv_rp(calc["tot_hpp"]))
    d1, d2, d3, d4 = st.columns(4)
    d1.metric("Margin Kotor", sv_rp(calc["margin"]))
    d2.metric("GPM", "%.1f%%" % (calc["gpm"] * 100))
    d3.metric("Cost of Money + PPh", sv_rp(calc["com"] + calc["pph"]))
    d4.metric("Margin Bersih", sv_rp(calc["net"]))
    if calc["tot_rev"] > 0:
        ux_progress([("HPP", calc["tot_hpp_items"], "#2563EB"), ("Biaya lain", calc["biaya_lain"], "#94A3B8"), ("Margin", max(calc["margin"], 0), "#10B981")])
    if calc["margin"] < 0: st.error("⚠️ Margin NEGATIF. Periksa Harga Tsel, Harga KUT, dan HPP sebelum mengajukan.")
    elif calc["tot_rev"] > 0 and calc["gpm"] < 0.05: st.warning("GPM di bawah 5%. Pastikan parameter KUT sudah benar.")
    rows = [{"Site": "%s %s" % (b["site_id"], b["site_name"]), "Ditagih ke Tsel": sv_rp(b["sum"]["A"]["tot_tsel"] + b["sum"]["B"]["tot_tsel"]),
             "Ditagih KUT ke 3E": sv_rp(b["sum"]["A"]["rev"] + b["sum"]["B"]["rev"]), "HPP Lapangan": sv_rp(b["sum"]["A"]["hpp"] + b["sum"]["B"]["hpp"])} for b in calc["sites"]]
    if rows: st.dataframe(pd.DataFrame(rows), use_container_width=True, hide_index=True)
    st.caption("Margin = Tagihan KUT→3E − (Total HPP + biaya lain). GPM = Margin ÷ Tagihan KUT. Margin bersih = Margin − Cost of Money − PPh.")

def _sv_step_ajukan(store, user, f, rab, calc):
    st.markdown("##### 4️⃣ Lampiran & Ajukan")
    files = st.file_uploader("Evidence referensi harga (gambar) — wajib bila ada Other Material", type=["png", "jpg", "jpeg", "webp"],
                             accept_multiple_files=True, key="svw_files")
    if files: f["evidence"] = [(x.name, x.getvalue()) for x in files]
    if f["evidence"]:
        cc1, cc2 = st.columns([3, 1])
        cc1.caption("%d gambar siap dilampirkan (tetap tersimpan walau Anda berpindah menu)." % len(f["evidence"]))
        if cc2.button("Hapus lampiran", key="svw_clr_ev", use_container_width=True):
            f["evidence"] = []; st.session_state.pop("svw_files", None); st.rerun()
    if f.get("old_urls"): st.caption("%d lampiran lama dari pengajuan sebelumnya tetap dipakai." % len(f["old_urls"]))
    f["catatan"] = st.text_area("Catatan untuk approver (opsional)", value=f["catatan"], key="svw_cat", height=70)
    errs = _sv_rab_problems(rab, calc, f, len(f["evidence"]))
    for e in errs: st.warning(e)
    ok = st.checkbox("Saya menyatakan data RAB ini benar.", key="svw_ok")
    label = "🚀 Ajukan Ulang RAB & Unduh Excel" if f.get("edit_id") else "🚀 Ajukan RAB & Unduh Excel"
    if st.button(label, type="primary", use_container_width=True, disabled=bool(errs) or not ok, key="svw_submit"):
        _sv_submit(store, user, f, rab, calc)

def _sv_submit(store, user, f, rab, calc):
    with st.spinner("Menyimpan RAB & membuat Excel..."):
        urls = []
        for name, raw in f["evidence"]:
            class _F:                                   
                def __init__(s, b): s.b = b
                def getvalue(s): return s.b
            u = upload_image(_F(raw), "sisava_rab_evidence")
            if u: urls.append(u)
        all_urls = list(f.get("old_urls") or []) + urls
        try:
            if f.get("edit_id"):
                rab_id = f["edit_id"]
                cur = next((r for r in store.list_rabs() if r.get("rab_id") == rab_id), {})
                rec = sv_rab_to_record(rab, calc, user, rab_id, all_urls)
                upd = {k: rec[k] for k in ("ticket_no", "nop", "area", "kota", "tot_tsel", "tot_rev", "tot_hpp", "margin", "gpm", "net", "payload", "evidence_rab")}
                upd.update(status=S_RAB_SPV, revisi=int(_sv_num(cur.get("revisi"))) + 1, updated_at=sv_now(),
                           spv_oleh="", spv_nama="", spv_jabatan="", spv_at="", spv_catatan="", mgr_oleh="", mgr_nama="", mgr_jabatan="", mgr_at="", mgr_catatan="",
                           log=_sv_json(sv_log_add(cur, user, "revisi_diajukan", f["catatan"])))
                store.update_rab(rab_id, upd)
            else:
                rab_id = sv_next_id(store.list_rabs())
                rec = sv_rab_to_record(rab, calc, user, rab_id, all_urls)
                rec["log"] = _sv_json(sv_log_add({}, user, "diajukan", f["catatan"]))
                store.add_rab(rec)
        except Exception as e:
            st.error("Gagal menyimpan: %s" % e); return
        meta = {"rab_id": rab_id, "status": S_RAB_SPV, "pembuat_nama": user["nama"], "pembuat_jabatan": user.get("jabatan", "")}
        data = sv_build_xlsx(rab, calc, meta, [b for _, b in f["evidence"]])
    st.session_state["sv_last"] = {"id": rab_id, "name": sv_rab_filename(rab, rab_id), "data": data, "tot": calc["tot_tsel"], "margin": calc["margin"]}
    st.session_state["sv_autodl"] = True
    _sv_form_reset()
    st.rerun()

def sv_page_form(store, user, pl, pl_path):
    last = st.session_state.get("sv_last")
    if last:
        st.success("✅ RAB **%s** berhasil diajukan dan tercatat. Menunggu persetujuan SPV Operation." % last["id"])
        if st.session_state.pop("sv_autodl", False): _sv_autodownload(last["data"], last["name"])
        c1, c2 = st.columns([3, 1])
        c1.download_button("📥 Unduh Excel RAB", last["data"], file_name=last["name"],
                           mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet", use_container_width=True, type="primary", key="svw_dl_last")
        if c2.button("Tutup", use_container_width=True, key="svw_close_last"):
            st.session_state.pop("sv_last", None); st.rerun()
    if "sv_form" not in st.session_state: _sv_form_reset()
    f = st.session_state["sv_form"]
    t1, t2, t3, t4 = st.tabs(["1️⃣ Info", "2️⃣ Site & Item", "3️⃣ Hitungan", "4️⃣ Ajukan"])
    with t1: _sv_step_info(f, pl_path)
    with t2: _sv_step_sites(f, pl)
    rab = _sv_rab_from_form(f); calc = sv_calc(rab)
    with t3: _sv_step_hitung(f, calc)
    with t4: _sv_step_ajukan(store, user, f, rab, calc)
    if f.get("edit_id") or any(s["items"] for s in f["sites"]) or f["ticket"]:
        if st.button("🧹 Kosongkan form", key="svw_clear"): _sv_form_reset(); st.rerun()

def _sv_act(rab_id, action, note_key=None, data=None):
    ss = st.session_state; store = sv_store(); user = ss.get("sv_user")
    rec = next((r for r in store.list_rabs() if r.get("rab_id") == rab_id), None)
    if not rec or not user: ss["sv_flash"] = "!Data tidak ditemukan / sesi berakhir."; return
    ok, msg, fields = sv_transition(rec, action, user, ss.get(note_key, "") if note_key else "", data() if callable(data) else data)
    if not ok: ss["sv_flash"] = "!" + msg; return
    try:
        store.update_rab(rab_id, fields); ss["sv_flash"] = "%s (%s)" % (msg, rab_id)
        if note_key: ss[note_key] = ""
        if action == "sva_fill": ss.pop("sv_sva_open", None)
    except Exception as e: ss["sv_flash"] = "!Gagal menyimpan: %s" % e

def _sv_timeline(rec):
    log = _sv_jload(rec.get("log"), []) or []
    if not log: st.caption("Belum ada riwayat."); return
    h = "".join('<div style="display:flex;gap:10px;padding:7px 0;border-bottom:1px solid var(--ux-line)"><div style="min-width:132px;font-size:.78rem;color:var(--ux-muted)">%s</div>'
                '<div style="font-size:.88rem"><b>%s</b> — %s%s</div></div>' % (_esc(x.get("ts", "")), _esc(_SV_AKSI.get(x.get("aksi"), x.get("aksi"))),
                _esc(x.get("nama") or x.get("oleh")), (" · <i>“%s”</i>" % _esc(x["catatan"])) if x.get("catatan") else "") for x in log)
    st.markdown(h, unsafe_allow_html=True)

def _sv_summary(rec):
    c1, c2, c3 = st.columns(3)
    c1.metric("Ditagih ke Telkomsel", sv_rp(rec.get("tot_tsel")))
    c2.metric("Revenue KUT→3E", sv_rp(rec.get("tot_rev")))
    c3.metric("HPP Lapangan", sv_rp(rec.get("tot_hpp")))
    d1, d2, d3 = st.columns(3)
    d1.metric("Margin Kotor", sv_rp(rec.get("margin")))
    d2.metric("GPM", "%.1f%%" % (_sv_num(rec.get("gpm")) * 100))
    d3.metric("Margin bersih", sv_rp(rec.get("net")))

def _sv_rec_header(rec):
    st.markdown("**%s** · Tiket `%s` · NOP %s · diajukan %s oleh **%s**<br>%s" % (_esc(rec.get("rab_id")), _esc(rec.get("ticket_no")), _esc(rec.get("nop")),
                _esc(str(rec.get("tgl_pengajuan"))[:16]), _esc(rec.get("nama_pembuat")), _sv_badge(rec.get("status"))), unsafe_allow_html=True)

def _sv_xlsx_button(rec, key):
    ss = st.session_state; k = "svx_%s" % key
    if ss.get(k, (None,))[0] == str(rec.get("updated_at")):
        data = ss[k][1]
        st.download_button("📥 Unduh Excel RAB", data, file_name=sv_rab_filename(_sv_jload(rec.get("payload"), {}) or {}, rec.get("rab_id")),
                           mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet", key="svdl_%s" % key, use_container_width=True)
    else:
        if st.button("📄 Siapkan Excel RAB", key="svprep_%s" % key, use_container_width=True):
            try: ss[k] = (str(rec.get("updated_at")), sv_record_xlsx(rec)); st.rerun()
            except Exception as e: st.error("Gagal membuat Excel: %s" % e)

def _sv_evidence_view(rec):
    ev = _sv_jload(rec.get("evidence_rab"), []) or []
    if ev: tampilkan_grid_foto(ev, "Referensi harga")
    kr = _sv_jload(rec.get("kerja_urls"), []) or []
    if kr:
        tampilkan_grid_foto([u for u in kr if not str(u).lower().endswith(".pdf")], "Bukti pekerjaan")
        for u in kr:
            if str(u).lower().endswith(".pdf"): st.markdown("[📎 Berkas PDF](%s)" % u)

def _sv_sva_view(rec):
    d = _sv_jload(rec.get("sva"), None)
    if d:
        st.markdown("**Nilai SVA** — Tiket `%s` · Status tiket: %s · Urgensi: **%s**<br>Nilai SVA %s · Keperluan dana %s · Rencana pekerjaan %s<br>Keperluan: %s" % (
            _esc(d.get("ticket_no")), _esc(d.get("status_tiket")), _esc(d.get("urgensi")), sv_rp(d.get("nilai_sva")), sv_rp(d.get("keperluan_dana")),
            _esc(d.get("tgl_rencana")), _esc(d.get("keperluan"))), unsafe_allow_html=True)

def sv_page_dashboard(user, recs):
    n = len(recs); cnt = {s: sum(1 for r in recs if r.get("status") == s) for s in SV_PIPELINE + [S_RAB_TOLAK, S_KRJ_TOLAK]}
    wait = cnt[S_RAB_SPV] + cnt[S_RAB_MGR] + cnt[S_KRJ_SPV] + cnt[S_KRJ_MGR]
    run = cnt[S_RAB_OK] + cnt[S_SVA]; tolak = cnt[S_RAB_TOLAK] + cnt[S_KRJ_TOLAK]
    ux_hero(SV_ROLES.get(user["role"], ""), "Portal SiSAVA", "Sistem Service Variable Activity — RAB, Approval, Nilai SVA, dan Bukti Pekerjaan.",
            stats=[(n, "Total RAB"), (wait, "Menunggu approval"), (run, "Berjalan"), (cnt[S_DONE], "SAH")],
            split=[("Menunggu approval", wait, "#FBBF24"), ("Berjalan", run, "#5B9BFF"), ("SAH", cnt[S_DONE], "#34D399"), ("Ditolak", tolak, "#FB7185")])
    aksi = {"STAFF": cnt[S_RAB_OK] + cnt[S_SVA] + tolak, "ENGINEERING": cnt[S_RAB_OK] + cnt[S_SVA] + tolak, "SPV": cnt[S_RAB_SPV] + cnt[S_KRJ_SPV], "MANAGER": cnt[S_RAB_SPV] + cnt[S_RAB_MGR] + cnt[S_KRJ_SPV] + cnt[S_KRJ_MGR]}[user["role"]]
    appr = [r for r in recs if r.get("status") not in (S_RAB_SPV, S_RAB_MGR, S_RAB_TOLAK)]
    rev, mar = sum(_sv_num(r.get("tot_rev")) for r in appr), sum(_sv_num(r.get("margin")) for r in appr)
    c1, c2, c3, c4 = st.columns(4)
    c1.metric("Perlu tindakan Anda", aksi); c2.metric("Revenue (RAB disetujui)", sv_rp(rev))
    c3.metric("Margin (RAB disetujui)", sv_rp(mar)); c4.metric("GPM", "%.1f%%" % ((mar / rev * 100) if rev else 0))
    if n:
        st.markdown("##### Posisi dokumen per tahap")
        ui_bar_chart(pd.DataFrame({"Jumlah RAB": [cnt[s] for s in SV_PIPELINE]}, index=SV_PIPELINE), color="#2563EB", horizontal=True)
        st.markdown("##### Terbaru")
        last = sorted(recs, key=lambda r: str(r.get("tgl_pengajuan")), reverse=True)[:8]
        st.dataframe(pd.DataFrame([{"No. RAB": r.get("rab_id"), "Tiket": r.get("ticket_no"), "NOP": r.get("nop"), "Status": r.get("status"), "Revenue": sv_rp(r.get("tot_rev")),
                                    "Margin": sv_rp(r.get("margin"))} for r in last]), use_container_width=True, hide_index=True)
    else: st.info("Belum ada RAB. Mulai dari menu ➕ Pengajuan RAB.")

def sv_page_list(user, recs):
    if not recs: st.info("Belum ada RAB."); return
    c1, c2 = st.columns([2, 1])
    q = c1.text_input("Cari (No. RAB / tiket / site / pembuat)", key="svw_lq")
    stt = c2.multiselect("Status", SV_PIPELINE + [S_RAB_TOLAK, S_KRJ_TOLAK], key="svw_ls")
    rows = [r for r in recs if (not stt or r.get("status") in stt) and (not q or q.lower() in (str(r.get("rab_id")) + str(r.get("ticket_no")) + str(r.get("payload"))[:3000] + str(r.get("nama_pembuat"))).lower())]
    rows = sorted(rows, key=lambda r: str(r.get("tgl_pengajuan")), reverse=True)
    st.dataframe(pd.DataFrame([{"No. RAB": r.get("rab_id"), "Tiket": r.get("ticket_no"), "NOP": r.get("nop"), "Pembuat": r.get("nama_pembuat"), "Status": r.get("status"),
                                "Revenue": sv_rp(r.get("tot_rev")), "Margin": sv_rp(r.get("margin")), "GPM": "%.1f%%" % (_sv_num(r.get("gpm")) * 100)} for r in rows]),
                 use_container_width=True, hide_index=True)
    if not rows: return
    ids = [r["rab_id"] for r in rows]
    pick = st.selectbox("Buka detail RAB", ids, key="svw_lpick")
    rec = next(r for r in rows if r["rab_id"] == pick)
    st.markdown("---"); _sv_rec_header(rec); _sv_summary(rec)
    a, b = st.columns(2)
    with a: _sv_xlsx_button(rec, "list_" + pick)
    with b:
        if rec.get("status") == S_RAB_TOLAK and (rec.get("dibuat_oleh") == user["username"] or user["role"] == "MANAGER"):
            st.button("✏️ Revisi RAB ini", key="svw_rev_" + pick, on_click=_sv_form_load, args=(rec,), type="primary", use_container_width=True)
    for lab, a_, c_ in (("SPV", rec.get("spv_nama"), rec.get("spv_catatan")), ("Manager", rec.get("mgr_nama"), rec.get("mgr_catatan"))):
        if a_ and c_: st.markdown("**Catatan %s (%s):** %s" % (lab, _esc(a_), _esc(c_)))
    _sv_sva_view(rec); _sv_evidence_view(rec)
    st.markdown("**Riwayat**"); _sv_timeline(rec)

def sv_page_approval(user, recs):
    can_spv, can_mgr = sv_can(user, "approve_spv"), sv_can(user, "approve_mgr")
    stat = set()
    if can_spv: stat |= {S_RAB_SPV, S_KRJ_SPV}
    if can_mgr: stat |= {S_RAB_SPV, S_RAB_MGR, S_KRJ_SPV, S_KRJ_MGR}
    q = [r for r in recs if r.get("status") in stat and (SV_ALLOW_SELF_APPROVE or r.get("dibuat_oleh") != user["username"])]
    own = [r for r in recs if r.get("status") in stat and r.get("dibuat_oleh") == user["username"] and not SV_ALLOW_SELF_APPROVE]
    st.markdown("##### Menunggu persetujuan Anda (%d)" % len(q))
    if own: st.caption("%d dokumen yang Anda ajukan sendiri tidak muncul di sini (harus disetujui orang lain)." % len(own))
    if not q: st.success("Tidak ada dokumen yang menunggu persetujuan Anda."); return
    for r in sorted(q, key=lambda x: str(x.get("tgl_pengajuan"))):
        rid, s_ = r["rab_id"], r["status"]
        with st.expander("%s · %s · %s" % (rid, r.get("ticket_no"), s_), expanded=False):
            _sv_rec_header(r); _sv_summary(r)
            _sv_xlsx_button(r, "ap_" + rid)
            _sv_sva_view(r); _sv_evidence_view(r)
            st.markdown("**Riwayat**"); _sv_timeline(r)
            nk = "svw_note_%s_%s" % (rid, s_)
            st.text_area("Catatan (wajib bila menolak)", key=nk, height=70)
            if s_ in (S_RAB_SPV, S_RAB_MGR): ya, tidak = ("spv_approve", "spv_reject") if (s_ == S_RAB_SPV and user["role"] == "SPV") else ("mgr_approve", "mgr_reject")
            else: ya, tidak = ("f_spv_approve", "f_spv_reject") if (s_ == S_KRJ_SPV and user["role"] == "SPV") else ("f_mgr_approve", "f_mgr_reject")
            lab = "✅ Setujui" + (" langsung (sebagai Manager)" if (user["role"] == "MANAGER" and s_ in (S_RAB_SPV, S_KRJ_SPV)) else "")
            b1, b2 = st.columns(2)
            b1.button(lab, key="svw_ok_%s_%s" % (rid, s_), type="primary", use_container_width=True, on_click=_sv_act, args=(rid, ya, nk))
            b2.button("↩️ Tolak / minta revisi", key="svw_no_%s_%s" % (rid, s_), use_container_width=True, on_click=_sv_act, args=(rid, tidak, nk))

def sv_page_sva(user, recs):
    todo = [r for r in recs if r.get("status") == S_RAB_OK]
    st.markdown("##### Validasi RAB & pengisian Nilai SVA (%d RAB siap)" % len(todo))
    if not todo: st.info("Belum ada RAB yang disetujui dan menunggu pengisian Nilai SVA."); 
    for r in todo:
        rid = r["rab_id"]
        with st.container(border=True):
            _sv_rec_header(r); _sv_summary(r)
            st.caption("✅ RAB telah disetujui SPV (%s) dan Manager (%s). Nilai di atas sudah tervalidasi." % (r.get("spv_nama") or "-", r.get("mgr_nama") or "-"))
            if sv_can(user, "sva_fill"):
                if st.session_state.get("sv_sva_open") != rid:
                    st.button("🧾 Isi Nilai SVA", key="svw_open_%s" % rid, type="primary", use_container_width=True, on_click=lambda i=rid: st.session_state.update(sv_sva_open=i))
                else:
                    k = lambda x: "svw_sva_%s_%s" % (x, rid)
                    c1, c2 = st.columns(2)
                    c1.text_input("No. Tiket", value=str(r.get("ticket_no", "")), key=k("t"))
                    c2.selectbox("Status Tiket", SV_STATUS_TIKET, key=k("st"))
                    c3, c4 = st.columns(2)
                    c3.number_input("Nilai SVA (Rp)", min_value=0.0, value=float(_sv_num(r.get("tot_tsel"))), step=1000.0, format="%.0f", key=k("n"))
                    c4.number_input("Keperluan Dana (Rp)", min_value=0.0, value=float(_sv_num(r.get("tot_hpp"))), step=1000.0, format="%.0f", key=k("d"))
                    c5, c6 = st.columns(2)
                    c5.selectbox("Urgensi Pekerjaan", SV_URGENSI, key=k("u"))
                    c6.date_input("Rencana tanggal pekerjaan diajukan", value=_dtm.date.today() + _dtm.timedelta(days=1), key=k("tg"))
                    st.text_area("Keperluan / uraian pekerjaan", key=k("kp"), height=80)
                    def _data(rid=rid):
                        g = st.session_state.get
                        return {"ticket_no": str(g("svw_sva_t_%s" % rid, "")).strip(), "status_tiket": g("svw_sva_st_%s" % rid, ""), "nilai_sva": g("svw_sva_n_%s" % rid, 0),
                                "keperluan_dana": g("svw_sva_d_%s" % rid, 0), "urgensi": g("svw_sva_u_%s" % rid, ""), "keperluan": str(g("svw_sva_kp_%s" % rid, "")).strip(),
                                "tgl_rencana": str(g("svw_sva_tg_%s" % rid, ""))}
                    b1, b2 = st.columns(2)
                    b1.button("💾 Simpan Nilai SVA", key="svw_save_%s" % rid, type="primary", use_container_width=True, on_click=_sv_act, args=(rid, "sva_fill", None, _data))
                    b2.button("Batal", key="svw_cancel_%s" % rid, use_container_width=True, on_click=lambda: st.session_state.pop("sv_sva_open", None))
    done = [r for r in recs if r.get("sva")]
    if done:
        st.markdown("##### Nilai SVA yang sudah diisi")
        st.dataframe(pd.DataFrame([{"No. RAB": r["rab_id"], "Tiket": (_sv_jload(r["sva"], {}) or {}).get("ticket_no"), "Status Tiket": (_sv_jload(r["sva"], {}) or {}).get("status_tiket"),
                                    "Urgensi": (_sv_jload(r["sva"], {}) or {}).get("urgensi"), "Nilai SVA": sv_rp((_sv_jload(r["sva"], {}) or {}).get("nilai_sva")),
                                    "Keperluan Dana": sv_rp((_sv_jload(r["sva"], {}) or {}).get("keperluan_dana")), "Rencana": (_sv_jload(r["sva"], {}) or {}).get("tgl_rencana"),
                                    "Status": r["status"]} for r in done]), use_container_width=True, hide_index=True)

def sv_page_kerja(user, recs):
    todo = [r for r in recs if r.get("status") in (S_SVA, S_KRJ_TOLAK)]
    st.markdown("##### Upload bukti pekerjaan (%d menunggu)" % len(todo))
    if not todo: st.info("Tidak ada pekerjaan yang menunggu unggahan bukti."); return
    pick = st.selectbox("Pilih RAB / pekerjaan", [r["rab_id"] for r in todo], format_func=lambda i: "%s · %s" % (i, next(x for x in todo if x["rab_id"] == i).get("ticket_no")), key="svw_kpick")
    r = next(x for x in todo if x["rab_id"] == pick)
    _sv_rec_header(r); _sv_sva_view(r)
    if r.get("status") == S_KRJ_TOLAK: st.warning("Pekerjaan ditolak. Catatan SPV: %s · Catatan Manager: %s" % (r.get("f_spv_catatan") or "-", r.get("f_mgr_catatan") or "-"))
    if not sv_can(user, "kerja_upload"): st.info("Akun Anda tidak berwenang mengunggah pekerjaan."); return
    fs = st.file_uploader("Foto / berkas bukti pekerjaan (sebelum, sesudah, BA, dll.)", type=["png", "jpg", "jpeg", "webp", "pdf"], accept_multiple_files=True, key="svw_kf_" + pick)
    cat = st.text_area("Catatan pekerjaan", key="svw_kc_" + pick, height=80)
    if st.button("📤 Kirim Bukti Pekerjaan", type="primary", use_container_width=True, key="svw_ksend_" + pick, disabled=not fs):
        with st.spinner("Mengunggah bukti..."):
            urls = []
            for fobj in fs:
                u = upload_datalog(fobj, "sisava_pekerjaan") if fobj.name.lower().endswith(".pdf") else upload_image(fobj, "sisava_pekerjaan")
                if u: urls.append(u)
        if len(urls) < len(fs): st.error("Sebagian berkas gagal diunggah (%d dari %d berhasil). Coba lagi." % (len(urls), len(fs)))
        else:
            ss = st.session_state; ss["svw_kc_tmp"] = cat
            _sv_act(pick, "kerja_upload", "svw_kc_tmp", {"urls": urls}); st.rerun()

def sv_page_kelola(store, user):
    t1, t2, t3 = st.tabs(["👥 Akun", "💲 Shopping List", "📊 Rekap"])
    with t1:
        users = store.list_users()
        st.dataframe(pd.DataFrame([{"Username": u.get("username"), "Nama": u.get("nama"), "Jabatan": u.get("jabatan"), "Role": SV_ROLES.get(u.get("role"), u.get("role")),
                                    "Aktif": "Ya" if str(u.get("aktif", "1")) not in ("0", "False", "false") else "Tidak"} for u in users]), use_container_width=True, hide_index=True)
        if users:
            with st.expander("🔑 Reset sandi / nonaktifkan akun", expanded=False):
                un = st.selectbox("Pilih akun", [u["username"] for u in users], key="svw_rpick")
                npw2 = st.text_input("Sandi baru (min. 8 karakter)", type="password", key="svw_rpw")
                x1, x2 = st.columns(2)
                if x1.button("Reset sandi", key="svw_rreset", use_container_width=True):
                    if len(npw2) < 8: st.error("Sandi minimal 8 karakter.")
                    else: store.upsert_user({"username": un, "pw_hash": sv_hash_pw(npw2)}); st.success("Sandi %s direset." % un)
                cur = sv_user_find(store, un) or {}
                aktif = str(cur.get("aktif", "1")) not in ("0", "False", "false")
                managers = [u for u in users if u.get("role") == "MANAGER" and str(u.get("aktif", "1")) not in ("0", "False", "false")]
                if x2.button("Nonaktifkan" if aktif else "Aktifkan", key="svw_rtoggle", use_container_width=True):
                    if aktif and (un == user["username"] or (cur.get("role") == "MANAGER" and len(managers) <= 1)): st.error("Tidak dapat menonaktifkan akun Anda sendiri / Manager aktif terakhir.")
                    else: store.upsert_user({"username": un, "aktif": "0" if aktif else "1"}); st.rerun()
    with t2:
        pl, path = sv_pricelist()
        if not pl: st.warning("File Shopping List belum ditemukan di folder aplikasi."); 
        else:
            st.success("Terbaca %d item dari `%s` (%d berharga tetap)." % (len(pl), os.path.basename(path), sum(1 for i in pl if i["price"])))
            q = st.text_input("Cari uraian", key="svw_plq")
            rows = [{"Kategori": i["cat"], "Uraian": i["desc"], "UoM": i["uom"], "Jenis": i["jenis"], "Kalimantan": i["price"].get("KALIMANTAN"), "Sulawesi": i["price"].get("SULAWESI"),
                     "PUMA": i["price"].get("PUMA"), "Keterangan": i["ket"]} for i in pl if not q or q.lower() in i["desc"].lower()][:400]
            st.dataframe(pd.DataFrame(rows), use_container_width=True, hide_index=True)
            st.caption("Untuk memperbarui harga, ganti file Excel di folder aplikasi lalu deploy ulang (cache otomatis membaca file baru).")
    with t3:
        recs = store.list_rabs()
        st.caption("%d RAB tercatat." % len(recs))
        if recs: st.download_button("📥 Unduh Rekap Seluruh RAB (Excel)", sv_rekap_xlsx(recs), file_name="Rekap_SiSAVA.xlsx",
                               mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet", use_container_width=True, key="svw_rekap")

def sisava_page():
    ss = st.session_state
    store = sv_store()
    _sv_flash()
    user = ss.get("sv_user")
    c1, c2 = st.columns([4, 1])
    c1.markdown("<div style='padding:8px 2px'>👤 <b>%s</b> · %s%s</div>" % (_esc(user["nama"]), _esc(SV_ROLES.get(user["role"], "")), (" · " + _esc(user["jabatan"])) if user.get("jabatan") else ""), unsafe_allow_html=True)
    c2.button("Keluar SiSAVA", key="svw_logout", on_click=lambda: st.session_state.pop("sv_user", None), use_container_width=True)
    nav = ["🏠 Dashboard", "➕ Pengajuan RAB", "📋 Daftar RAB"]
    if sv_can(user, "approve_spv") or sv_can(user, "approve_mgr"): nav.append("✅ Approval")
    nav += ["🧾 Nilai SVA", "📤 Upload Pekerjaan"]
    if sv_can(user, "manage_users"): nav.append("🛠️ Kelola")
    if ss.get("sv_nav") not in nav: ss["sv_nav"] = nav[0]
    page = st.radio("Menu SiSAVA", nav, horizontal=True, key="sv_nav", label_visibility="collapsed")
    try: recs = sv_visible(store.list_rabs(), user)
    except Exception as e: st.error("Gagal membaca database SiSAVA: %s" % e); return
    if page == "🏠 Dashboard": sv_page_dashboard(user, recs)
    elif page == "➕ Pengajuan RAB":
        pl, path = sv_pricelist(); sv_page_form(store, user, pl, path)
    elif page == "📋 Daftar RAB": sv_page_list(user, recs)
    elif page == "✅ Approval": sv_page_approval(user, recs)
    elif page == "🧾 Nilai SVA": sv_page_sva(user, recs)
    elif page == "📤 Upload Pekerjaan": sv_page_kerja(user, recs)
    elif page == "🛠️ Kelola": sv_page_kelola(store, user)
# =============================================================================
if 'active_menu' not in st.session_state:
    st.session_state['active_menu'] = None

# --- KODE WATERMARK GLOBAL (MUNCUL DI MENU UTAMA & DALAM FORM) ---
bg_base64 = get_base64_of_bin_file(LOGO_UTAMA)
if bg_base64:
    st.markdown(
        f"""
        <style>
        .watermark-kut {{
            position: fixed;
            top: 50%;
            left: 50%;
            transform: translate(-50%, -50%);
            opacity: 0.05;
            z-index: 0;
            pointer-events: none;
        }}
        .watermark-kut img {{
            width: 750px;
            max-width: 90vw;
            filter: grayscale(100%);
        }}
        </style>
        <div class="watermark-kut">
            <img src="data:image/webp;base64,{bg_base64}">
        </div>
        """, unsafe_allow_html=True
    )
# -----------------------------------------------------------------

if st.session_state['active_menu'] is None:
    render_header_logo()
    role_badge = "🟢 MODE ADMIN UTAMA" if st.session_state['role'] == 'Admin' else f"👷‍♂️ TIM LAPANGAN ({st.session_state.get('area_lapangan', 'Umum').upper()})"
    _ux_stats, _ux_split = ux_stats_from_db(st.session_state.get("laporan_db", []))
    ux_hero(role_badge, "Sistem Operasional Terpadu SIDATA", "Pilih modul untuk melanjutkan. Data laporan tersinkron langsung dari server.", stats=_ux_stats, split=_ux_split)

    # === PANEL KHUSUS ADMIN UTAMA: MANAJEMEN AKUN SiSAVA ===
    if st.session_state['role'] == 'Admin':
        st.markdown("### 🛠️ Manajemen Akun Pengguna (SiSAVA)")
        st.info("Sebagai Admin Utama, Anda dapat mendaftarkan akun untuk Manager, SPV, Engineering, dan Tim Operasional agar mereka dapat mengakses Portal SiSAVA.")
        
        store = sv_store()
        users = store.list_users()
        
        with st.expander("➕ Buat Akun SiSAVA Baru / Reset Sandi", expanded=True):
            tb_buat, tb_reset = st.tabs(["📝 Buat Akun Baru", "🔑 Reset Sandi"])
            
            with tb_buat:
                a1, a2 = st.columns(2)
                nu_adm = a1.text_input("Username Baru (Tanpa Spasi)", key="svw_nu_adm_main")
                nn_adm = a2.text_input("Nama Lengkap", key="svw_nn_adm_main")
                a3, a4 = st.columns(2)
                nj_adm = a3.text_input("Jabatan Tertulis", placeholder="mis: Koordinator / Engineering", key="svw_nj_adm_main")
                nr_adm = a4.selectbox("Hak Akses (Role)", list(SV_ROLES), format_func=lambda x: SV_ROLES[x], key="svw_nr_adm_main")
                pw_adm = st.text_input("Sandi (min. 8 karakter)", type="password", key="svw_pw_adm_main")
                
                if st.button("💾 Simpan & Daftarkan Akun", type="primary", key="btn_save_akun_main"):
                    msg = sv_validate_new_user({"username": nu_adm.strip(), "nama": nn_adm, "role": nr_adm}, pw_adm)
                    if not msg and sv_user_find(store, nu_adm): msg = "Username tersebut sudah dipakai."
                    if msg: 
                        st.error(msg)
                    else:
                        store.upsert_user({"username": nu_adm.strip(), "nama": nn_adm.strip(), "jabatan": nj_adm.strip(), "role": nr_adm, "pw_hash": sv_hash_pw(pw_adm), "aktif": "1", "dibuat": sv_now()})
                        st.success(f"✅ Akun untuk {nu_adm.strip()} ({SV_ROLES[nr_adm]}) berhasil didaftarkan! Mereka kini bisa masuk melalui Portal SiSAVA.")
                        st.rerun()

            with tb_reset:
                if not users:
                    st.warning("Belum ada akun terdaftar.")
                else:
                    ru_adm = st.selectbox("Pilih Akun yang akan Direset", [x["username"] for x in users], key="svw_ru_adm_main")
                    rpw_adm = st.text_input("Sandi Baru (min. 8 karakter)", type="password", key="svw_rpw_adm_main")
                    if st.button("🔄 Reset Sandi Akun", key="btn_reset_akun_main"):
                        if len(rpw_adm) < 8: 
                            st.error("Sandi minimal 8 karakter.")
                        else: 
                            store.upsert_user({"username": ru_adm, "pw_hash": sv_hash_pw(rpw_adm)})
                            st.success(f"✅ Sandi untuk akun {ru_adm} berhasil direset.")
        
        st.markdown("<hr style='margin: 30px 0;'>", unsafe_allow_html=True)

    # === KOTAK MENU NAVIGASI UTAMA ===
    col_m1, col_m2 = st.columns(2)
    with col_m1:
        if st.session_state['role'] == 'Admin':
            st.markdown("""
                <div class='area-card' style='border-top-color: #007BFF;'>
                    <div style='font-size: 46px; margin-bottom: 8px;'>📝</div>
                    <div class='login-title'>Form Preventive Check</div>
                    <div class='login-desc'>Input laporan Preventive Maintenance Site SPS, Site Reguler Palangkaraya, dan Site Visit Pangkalan Bun.</div>
                </div>
            """, unsafe_allow_html=True)
            if st.button("📝 BUKA FORM PREVENTIVE CHECK", type="primary", use_container_width=True, key="m_form_adm"):
                st.session_state['active_menu'] = "📝 Form Preventive Check"
                st.rerun()
        else:
            st.markdown("""
                <div class='area-card' style='border-top-color: #007BFF;'>
                    <div style='font-size: 46px; margin-bottom: 8px;'>📸</div>
                    <div class='login-title'>Form Pelaporan (Lapangan)</div>
                    <div class='login-desc'>Input data pengerjaan lapangan, pengukuran parameter kelistrikan AC/DC, dan upload foto eviden wajib.</div>
                </div>
            """, unsafe_allow_html=True)
            if st.button("📸 BUKA FORM PELAPORAN LAPANGAN", type="primary", use_container_width=True, key="m_form_lap"):
                st.session_state['active_menu'] = "📸 Form Pelaporan (Lapangan)"
                st.rerun()

        st.markdown("<br>", unsafe_allow_html=True)
        st.markdown("""
            <div class='area-card' style='border-top-color: #00E676;'>
                <div style='font-size: 46px; margin-bottom: 8px;'>📈</div>
                <div class='login-title'>Monitoring Improvement</div>
                <div class='login-desc'>Pantau grafik Kurva S (Plan vs Actual), progress tim lapangan (Done/On Progress), dan Live Editor Tracker.</div>
            </div>
        """, unsafe_allow_html=True)
        if st.button("📈 BUKA MONITORING IMPROVEMENT", type="secondary", use_container_width=True, key="m_mon"):
            st.session_state['active_menu'] = "📈 Monitoring Improvement"
            st.rerun()

    with col_m2:
        st.markdown("""
            <div class='area-card' style='border-top-color: #FF3333;'>
                <div style='font-size: 46px; margin-bottom: 8px;'>📊</div>
                <div class='login-title'>Hasil Laporan & Dashboard</div>
                <div class='login-desc'>Lihat rekapitulasi laporan NOP Palangkaraya & Pangkalan Bun, unduh Excel/PPTX/PDF, dan bagikan ke WhatsApp.</div>
            </div>
        """, unsafe_allow_html=True)
        if st.button("📊 BUKA HASIL LAPORAN & DASHBOARD", type="primary", use_container_width=True, key="m_dash"):
            st.session_state['active_menu'] = "📊 Hasil Laporan & Dashboard"
            st.rerun()

        st.markdown("<br>", unsafe_allow_html=True)
        st.markdown("""
            <div class='area-card' style='border-top-color: #FFC107;'>
                <div style='font-size: 46px; margin-bottom: 8px;'>⚡</div>
                <div class='login-title'>Dapot Power Hasil PMS</div>
                <div class='login-desc'>Executive Dashboard untuk compare bulanan tegangan & beban AC/DC pada site kunjungan rutin maupun 1x kunjungan.</div>
            </div>
        """, unsafe_allow_html=True)
        if st.button("⚡ BUKA DAPOT POWER HASIL PMS", type="secondary", use_container_width=True, key="m_dapot"):
            st.session_state['active_menu'] = "⚡ Dapot Power Hasil PMS"
            st.rerun()

    st.markdown("<hr style='margin: 30px 0;'>", unsafe_allow_html=True)
    c_out1, c_out2, c_out3 = st.columns(3)
    with c_out1:
        st.markdown("""
            <div class='area-card' style='border-top-color: #9C27B0;'>
                <div style='font-size: 42px; margin-bottom: 8px;'>📅</div>
                <div class='login-title' style='font-size: 20px;'>Daily Tracker PLK</div>
                <div class='login-desc' style='font-size: 13px;'>To Do List harian tim Onsite, update status pengerjaan, dan pencatatan FME.</div>
            </div>
        """, unsafe_allow_html=True)
        if st.button("📅 BUKA DAILY TRACKER", type="primary", use_container_width=True, key="m_tracker"):
            st.session_state['active_menu'] = "📅 Daily Tracker Palangkaraya"
            st.rerun()
            
    with c_out2:
        st.markdown("""
            <div class='area-card' style='border-top-color: #EF4444;'>
                <div style='font-size: 42px; margin-bottom: 8px;'>🚨</div>
                <div class='login-title' style='font-size: 20px;'>SiHD (Sistem Helpdesk)</div>
                <div class='login-desc' style='font-size: 13px;'>Buku mutasi elektronik pergerakan tim, rincian pekerjaan, dan WA serah terima shift.</div>
            </div>
        """, unsafe_allow_html=True)
        if st.button("🚨 BUKA SiHD (MUTASI TIM)", type="primary", use_container_width=True, key="m_sihd"):
            st.session_state['active_menu'] = "🚨 SiHD (Pergerakan Tim)"
            st.rerun()

    with c_out3:
        st.markdown("""
            <div class='area-card' style='border-top-color: #10B981;'>
                <div style='font-size: 42px; margin-bottom: 8px;'>💠</div>
                <div class='login-title' style='font-size: 20px;'>Portal SiSAVA</div>
                <div class='login-desc' style='font-size: 13px;'>Sistem Service Variable Activity: Pengajuan RAB, Approval SPV, dan Nilai SVA.</div>
            </div>
        """, unsafe_allow_html=True)
        if st.button("💠 BUKA PORTAL SiSAVA", type="primary", use_container_width=True, key="m_sisava"):
            st.session_state['active_menu'] = "💠 SiSAVA"
            st.rerun()

    st.markdown("<br>", unsafe_allow_html=True)
    c_tool1, c_tool2, c_tool3 = st.columns(3)
    with c_tool1:
        st.markdown("""
            <div class='area-card' style='border-top-color: #FFD700; background: linear-gradient(145deg, #0f172a, #1e293b);'>
                <div style='font-size: 42px; margin-bottom: 8px;'>💎</div>
                <div class='login-title' style='font-size: 20px; color: #FFD700 !important;'>Executive Project</div>
                <div class='login-desc' style='font-size: 13px; color: #cbd5e1 !important;'>Dashboard Khusus PO & SVA: Interactive Timeline, Margin & Live Update.</div>
            </div>
        """, unsafe_allow_html=True)
        if st.button("💎 BUKA EXECUTIVE PROJECT", type="primary", use_container_width=True, key="m_proj_dash"):
            st.session_state['active_menu'] = "💎 Executive Project Dashboard"
            st.rerun()

    with c_tool2:
        if st.session_state['role'] in ['Admin', 'Logistik']:
            st.markdown("""
                <div class='area-card' style='border-top-color: #FF9800;'>
                    <div style='font-size: 42px; margin-bottom: 8px;'>📦</div>
                    <div class='login-title' style='font-size: 20px;'>Logistik & SPMS</div>
                    <div class='login-desc' style='font-size: 13px;'>Manajemen Inbound/Outbound stok material, oli, filter genset, dan gudang.</div>
                </div>
            """, unsafe_allow_html=True)
            if st.button("📦 BUKA LOGISTIK", type="secondary", use_container_width=True, key="m_log"):
                st.session_state['active_menu'] = "📦 Logistik & SPMS"
                st.rerun()
                
    with c_tool3:
        st.markdown("""
            <div class='area-card' style='border-top-color: #06B6D4;'>
                <div style='font-size: 44px; margin-bottom: 6px;'>🗺️</div>
                <div class='login-title' style='font-size: 19px;'>Sysinfomap (Peta Site)</div>
                <div class='login-desc' style='font-size: 13px;'>Peta interaktif 910+ Site, populasi per Kabupaten, dan komparasi Master vs KML.</div>
            </div>
        """, unsafe_allow_html=True)
        if st.button("🗺️ BUKA SYSINFOMAP", type="primary", use_container_width=True, key="m_sysinfomap"):
            st.session_state['active_menu'] = "🗺️ Sysinfomap (Peta & Populasi Site)"
            st.rerun()
            
    st.markdown("<br><hr style='margin: 10px 0;'>", unsafe_allow_html=True)
    if st.button("🚪 KELUAR AKUN", type="secondary", use_container_width=True, key="btn_logout_main"):
        st.session_state['role'] = None
        st.session_state['dashboard_area'] = None
        st.session_state['active_menu'] = None
        st.session_state['proj_dash_unlocked'] = False
        st.session_state['proj_unlocked'] = False
        st.session_state.pop('sv_user', None)
        st.rerun()
    st.stop()
else:
    menu = st.session_state['active_menu']
    c_top1, c_top2, c_top3 = st.columns([2.2, 5.6, 2.2])
    with c_top1:
        if st.button("⬅️ KEMBALI KE MENU UTAMA", type="secondary", use_container_width=True, key="btn_back_main"):
            st.session_state['active_menu'] = None
            st.session_state['dashboard_area'] = None
            st.rerun()
    with c_top2:
        st.markdown(
            f"<div style='text-align: center; padding: 10px 16px; border-radius: 10px; background-color: var(--secondary-background-color); border: 1px solid var(--primary-color); font-weight: 800; font-size: 16px;'>📍 Menu Aktif: {menu}</div>",
            unsafe_allow_html=True
        )
    with c_top3:
        if st.button("🚪 Keluar Akun (Log Out)", use_container_width=True, key="btn_logout_top"):
            st.session_state['role'] = None
            st.session_state['dashboard_area'] = None
            st.session_state['active_menu'] = None
            st.rerun()
    st.markdown("<hr style='margin: 15px 0 25px 0;'>", unsafe_allow_html=True)


# =========================================================================
# MENU 1: FORM PENGISIAN LENGKAP
# =========================================================================
if menu in ["📝 Form Preventive Check", "📸 Form Pelaporan (Lapangan)"]:
    
    # Ambil Data Master untuk Dropdown
    fme_plk_list, fme_pbu_list, site_plk_list, site_pbu_list = load_master_lists()
    
    if menu == "📸 Form Pelaporan (Lapangan)":
        if st.session_state.get('area_lapangan') == 'Pangkalan Bun':
            st.markdown("<h1 style='text-align: center;'>📸 Form Pelaporan (Pangkalan Bun)</h1>", unsafe_allow_html=True)
            st.info("Sistem telah mengunci akses Anda ke Form Pangkalan Bun berdasarkan pilihan di layar login.")
            standar_form = "Pangkalan Bun"
            nop_area_default = "Pangkalan Bun"
        else:
            st.markdown("<h1 style='text-align: center;'>📸 Form Pelaporan (Standar/Palangkaraya)</h1>", unsafe_allow_html=True)
            st.info("Sistem telah mengunci akses Anda ke Form Standar berdasarkan pilihan di layar login.")
            standar_form = "Standar"
            nop_area_default = "Palangkaraya"
    else:
        st.markdown("<h1 style='text-align: center;'>⚡ Form Preventive Maintenance (Admin)</h1>", unsafe_allow_html=True)
        pilih_form = st.radio("Pilih Format Laporan:", ["Standar (Palangkaraya/Reguler)", "Pangkalan Bun"], horizontal=True)
        standar_form = "Pangkalan Bun" if "Pangkalan Bun" in pilih_form else "Standar"
        nop_area_default = "Palangkaraya"

    # -------------------------------------------------------------------------
    # BLOK A: FORM STANDAR PALANGKARAYA (DIPISAH SPS VS REGULER TAPI 1 DATABASE)
    # -------------------------------------------------------------------------
    if standar_form == "Standar":
        
        st.info("💡 Data dan Lampiran (Foto & Datalog) Anda akan dienkripsi dan dikirim langsung ke Cloud Server.")
        site_category = st.radio("Kategori Site (Wajib Pilih):", ["SPS", "Site Reguler", "Project"], horizontal=True)
        st.markdown("<hr style='margin:10px 0;'>", unsafe_allow_html=True)
        
        if site_category == "SPS":
            tabs = st.tabs(["📌 1. Info Site", "☀️ 2. SPS Panel", "🔌 3. PLN & Recti", "⛽ 4. Genset & BBM", "🔋 5. Baterai & Gnd", "📤 6. Upload & Submit"])

            with tabs[0]:
                st.markdown("#### Kategori Kunjungan: SPS")
                c1, c2 = st.columns(2)
                with c1:
                    if site_plk_list:
                        site_name = st.selectbox("Pilih Site ID", [""] + site_plk_list)
                    else:
                        site_name = st.text_input("Nama / ID Site", placeholder="Contoh: BTS-PKY-001")
                        
                    opsi_nop = ["Palangkaraya", "Pangkalan Bun", "Tarakan", "Pontianak", "Lainnya"]
                    idx_default = opsi_nop.index(nop_area_default) if nop_area_default in opsi_nop else 0
                    nop_area = st.selectbox("NOP Area", opsi_nop, index=idx_default)
                    check_date = st.date_input("Tanggal Pengecekan", value=datetime.date.today())
                with c2:
                    if fme_plk_list:
                        technician_name = st.multiselect("Pilih Nama Pelaksana / FME (Bisa >1):", fme_plk_list)
                    else:
                        tek_str = st.text_input("Nama Pelaksana / Teknisi")
                        technician_name = [tek_str] if tek_str else []
                        
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
                final_status = st.radio("Status Akhir Site:", ["Normal", "Minor Issue", "Major/Critical", "Update Progress"])

                if st.button("🚀 UPLOAD & SINKRONKAN DATA SPS", type="primary", use_container_width=True):
                    if not site_name:
                        st.error("⚠️ Mohon pilih Nama Site!")
                    elif not technician_name:
                        st.error("⚠️ Mohon pilih Pelaksana / FME!")
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

                            tek_str_final = ", ".join(technician_name)
                            report_dict = {
                                "timestamp": datetime.datetime.now().strftime("%Y-%m-%d %H:%M:%S"),
                                "site_name": site_name, 
                                "kategori": "SPS", 
                                "nop": nop_area, 
                                "teknisi": tek_str_final,
                                "status": final_status, 
                                "action": action_taken, 
                                "sparepart": sparepart_needed,
                                
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
                            
                            sheet = connect_gsheets("Report Preventive")
                            if sheet:
                                row_data = [
                                    report_dict.get('timestamp', ''), report_dict.get('site_name', ''),
                                    report_dict.get('nop', ''), report_dict.get('teknisi', ''),
                                    report_dict.get('status', ''), json.dumps(report_dict)
                                ]
                                sheet.insert_row(row_data, index=2)
                                st.cache_data.clear()
                                st.session_state['laporan_db'].append(report_dict)
                                st.success("✅ BERHASIL! Data SPS telah diamankan ke Baris Teratas Spreadsheet.")
                            else:
                                st.error("⚠️ Gagal menyambung ke Server Database.")

        elif site_category == "Site Reguler":
            # -------------------------------------------------------------------------
            # BLOK A.2: FORM KHUSUS SITE REGULER & AUDIT TOWER
            # -------------------------------------------------------------------------
            tabs_reg = st.tabs(["📌 1. Info & Problem", "⚡ 2. Action & Parameter", "🏗️ 3. Audit Fisik Tower", "📤 4. Submit"])
            cond_data = {}
            cond_files = {}
            
            with tabs_reg[0]:
                st.markdown("#### Kategori Kunjungan: Site Reguler (Maintenance / Audit Tower)")
                c1, c2 = st.columns(2)
                with c1:
                    if site_plk_list:
                        site_name_reg = st.selectbox("Pilih Site ID (Wajib):", [""] + site_plk_list, key="reg_site")
                    else:
                        site_name_reg = st.text_input("Nama / ID Site (Wajib):", key="reg_site")
                        
                    opsi_nop = ["Palangkaraya", "Pangkalan Bun", "Tarakan", "Pontianak", "Lainnya"]
                    idx_default = opsi_nop.index(nop_area_default) if nop_area_default in opsi_nop else 0
                    nop_area_reg = st.selectbox("NOP Area:", opsi_nop, index=idx_default, key="reg_nop")
                with c2:
                    if fme_plk_list:
                        teknisi_reg = st.multiselect("Pilih Nama FME (Wajib, Bisa >1):", fme_plk_list, key="reg_fme")
                    else:
                        t_str = st.text_input("Nama FME (Wajib):", key="reg_fme")
                        teknisi_reg = [t_str] if t_str else []
                    
                    role_fme_reg = st.selectbox("Role FME:", ["TE", "MBP", "CME"], key="reg_role")
                    check_date_reg = st.date_input("Tanggal Kunjungan:", value=datetime.date.today(), key="reg_date")
                
                st.divider()
                st.markdown("### 🚨 Problem Alarm / Tujuan Visit")
                prob_alarm = st.selectbox(
                    "Pilih Problem Utama / Tujuan:", 
                    ["Down power", "Fliker power", "Fliker Site Down", "Transmisi Problem", "cell down", "Audit Fisik / Kondisi Tower"]
                )
                foto_prob = st.file_uploader("📸 Foto Problem / Kondisi Awal (WAJIB DILAMPIRKAN):", accept_multiple_files=True, key="f_prob")

            with tabs_reg[1]:
                st.markdown("### 🔧 Action & Perbaikan")
                
                base_act_opts = [
                    "Replace RRU", "Replace Optik", "Replace SFP", "Replace Module APR", 
                    "Teg PLN Drop pelaporan PLN", "Reset Hardware", "MCB Trip", "Add Baterai",
                    "Cleaning Fan BTS", "Cleaning FAN Recti", "Cleaning FAN Perangkat TRM", "Cleaning FAN Shelter",
                    "Troubleshoot Genset / ATS", "Pengecekan Vertikalitas Tower", "Inspeksi & Kencangkan Baut/Skur"
                ]
                existing_actions = set(base_act_opts)
                if 'laporan_db' in st.session_state and st.session_state['laporan_db']:
                    for r_db in st.session_state['laporan_db']:
                        if r_db.get('kategori') == 'Site Reguler' and r_db.get('action_reguler'):
                            acts = r_db.get('action_reguler')
                            if isinstance(acts, list):
                                for a in acts: existing_actions.add(a)
                            elif isinstance(acts, str):
                                existing_actions.add(acts)
                
                act_opts = sorted(list(existing_actions))
                act_sel = st.multiselect("Pilih Action Dikerjakan (WAJIB PILIH, Bisa Multiple):", act_opts)
                
                custom_act = st.text_input("➕ Tambah Action Manual (Ketik di sini jika tidak ada di pilihan atas):", placeholder="Contoh: Ganti Filter Oli Mesin...")
                if custom_act:
                    if custom_act not in act_sel:
                        act_sel.append(custom_act)
                        
                foto_act = st.file_uploader("📸 Foto Action / Perbaikan (WAJIB DILAMPIRKAN):", accept_multiple_files=True, key="f_act")

                show_cond_power = True 
                show_cond_genset = any("Genset" in item or "ATS" in item for item in act_sel)
                
                if show_cond_power or show_cond_genset:
                    st.divider()
                    st.warning("⚠️ **WAJIB DIISI 100%:** Seluruh kolom nilai ukur AC/DC dan foto eviden di bawah ini wajib dilengkapi!")
                
                if show_cond_power:
                    st.markdown("#### ⚡ A. Pengukuran AC/DC & Eviden Foto")
                    cc1, cc2, cc3 = st.columns(3)
                    with cc1:
                        cond_data['v_rn'] = st.text_input("Teg AC R-N [V] *", placeholder="Contoh: 220")
                        cond_files['f_v_rn'] = st.file_uploader("📸 Foto Eviden Teg AC R-N *", key="f_v_rn")
                        cond_data['i_r'] = st.text_input("Load AC R [A] *", placeholder="Contoh: 12.5")
                        cond_files['f_i_r'] = st.file_uploader("📸 Foto Eviden Load AC R *", key="f_i_r")
                    with cc2:
                        cond_data['v_sn'] = st.text_input("Teg AC S-N [V] *", placeholder="Contoh: 221")
                        cond_files['f_v_sn'] = st.file_uploader("📸 Foto Eviden Teg AC S-N *", key="f_v_sn")
                        cond_data['i_s'] = st.text_input("Load AC S [A] *", placeholder="Contoh: 11.8")
                        cond_files['f_i_s'] = st.file_uploader("📸 Foto Eviden Load AC S *", key="f_i_s")
                    with cc3:
                        cond_data['v_tn'] = st.text_input("Teg AC T-N [V] *", placeholder="Contoh: 219")
                        cond_files['f_v_tn'] = st.file_uploader("📸 Foto Eviden Teg AC T-N *", key="f_v_tn")
                        cond_data['i_t'] = st.text_input("Load AC T [A] *", placeholder="Contoh: 10.4")
                        cond_files['f_i_t'] = st.file_uploader("📸 Foto Eviden Load AC T *", key="f_i_t")

                    st.markdown("**Parameter DC & Modul**")
                    cc7, cc8 = st.columns(2)
                    with cc7:
                        cond_data['load_dc'] = st.text_input("Load DC Total [A] *", placeholder="Contoh: 45.2")
                        cond_files['f_ldc'] = st.file_uploader("📸 Foto Eviden Load DC *", key="f_ldc")
                        cond_data['set_bcl'] = st.text_input("Settingan BCL [%] *", placeholder="Contoh: 100%")
                        cond_files['f_bcl'] = st.file_uploader("📸 Foto Eviden Settingan BCL *", key="f_bcl")
                    with cc8:
                        cond_data['set_cap'] = st.text_input("Settingan Capacity [Ah] *", placeholder="Contoh: 400Ah")
                        cond_files['f_cap'] = st.file_uploader("📸 Foto Eviden Settingan Cap *", key="f_cap")
                        cond_data['mod_rusak'] = st.text_input("Module Kontroler Rusak * (Isi 'Aman/0' jika normal)")
                        cond_files['f_mod'] = st.file_uploader("📸 Foto Eviden Module *", key="f_mod")

                if show_cond_genset:
                    st.markdown("#### ⚙️ B. Pengecekan Mesin Genset / ATS & Eviden Foto")
                    cg1, cg2, cg3 = st.columns(3)
                    with cg1:
                        cond_data['gen_v'] = st.text_input("Tegangan Out Genset [V] *", placeholder="Cth: 220")
                        cond_files['f_gen_v'] = st.file_uploader("📸 Foto Tegangan Genset *", key="f_gen_v")
                        cond_data['gen_i'] = st.text_input("Load Arus Genset [A] *", placeholder="Cth: 25.4")
                        cond_files['f_gen_i'] = st.file_uploader("📸 Foto Arus Genset *", key="f_gen_i")
                    with cg2:
                        cond_data['gen_ats'] = st.selectbox("Status Panel ATS *", ["Auto Normal", "Manual", "Rusak / Bypass"])
                        cond_files['f_gen_ats'] = st.file_uploader("📸 Foto Panel ATS *", key="f_gen_ats")
                        cond_data['gen_issue'] = st.selectbox("Issue Mesin Genset *", ["Normal / Aman", "Hunting / Suara Kasar", "Overheat", "Bocor Oli / BBM", "Mati Total / Tidak Bisa Start"])
                        cond_files['f_gen_issue'] = st.file_uploader("📸 Foto Fisik Mesin *", key="f_gen_issue")
                    with cg3:
                        cond_data['gen_hm'] = st.text_input("Hour Meter (HM) *", placeholder="Cth: 1250.5")
                        cond_files['f_gen_hm'] = st.file_uploader("📸 Foto Hour Meter *", key="f_gen_hm")
                        cond_data['gen_bbm'] = st.text_input("Estimasi Sisa BBM [Liter] *", placeholder="Cth: 150")
                        cond_files['f_gen_bbm'] = st.file_uploader("📸 Foto Tangki/Sisa BBM *", key="f_gen_bbm")

            with tabs_reg[2]:
                st.markdown("### 🏗️ Audit Fisik & Perkuatan Tower (Mencegah Insiden Roboh)")
                st.info("💡 Jika site ini adalah **Combat** atau diindikasi mengalami kemiringan/kerusakan struktur, **WAJIB** isi form ini.")
                
                c_tw1, c_tw2 = st.columns(2)
                with c_tw1:
                    cond_data['tower_type'] = st.selectbox("Tipe Tower", ["SST (Self Supporting Tower)", "Guyed Mast (Tarikan Seling)", "Monopole", "Combat", "Rooftop"])
                    cond_data['tower_vert'] = st.selectbox("Kondisi Kemiringan (Vertikalitas)", ["Aman (Tegak Lurus)", "Miring Ringan (<2 Derajat)", "Miring Kritis / Bahaya"])
                    cond_data['tower_skur'] = st.selectbox("Kondisi Tali Skur / Seling (Khusus Combat/Guyed)", ["Aman & Kencang", "Kendor", "Karat / Putus Sebagian", "Tidak Ada / Bukan Guyed"])
                with c_tw2:
                    cond_data['tower_baut'] = st.selectbox("Kondisi Baut & Struktur Besi", ["Aman (Lengkap)", "Ada Baut Lepas / Hilang", "Besi Keropos / Patah"])
                    cond_data['tower_base'] = st.selectbox("Kondisi Pondasi (Base)", ["Kokoh", "Tanah Ambles / Erosi", "Retak"])
                    cond_data['tower_remark'] = st.text_area("Catatan Kerusakan Struktur:", placeholder="Jelaskan detail kerusakan pada tower (jika ada)...")

                st.markdown("**📸 Lampiran Eviden Fisik Tower:**")
                cond_files['f_tw_full'] = st.file_uploader("📸 Foto Tower Full Body (Bawah sampai Atas)", accept_multiple_files=True, key="f_tw_full")
                cond_files['f_tw_defect'] = st.file_uploader("📸 Foto Titik Kerusakan (Seling Kendor/Baut Lepas/Pondasi Retak)", accept_multiple_files=True, key="f_tw_defect")

                cond_data['note_action'] = st.text_area("📝 NOTE / REMARK Action Lapangan (Keseluruhan) *:", placeholder="Jelaskan secara spesifik hasil pengerjaan atau temuan akhir...", key="reg_note_action")

            with tabs_reg[3]:
                st.markdown("### 📤 Konfirmasi & Submit Laporan")
                final_status_reg = st.radio("Status Akhir Site:", ["Normal", "Minor Issue", "Major/Critical", "Update Progress"], key="stat_reg")
                
                if st.button("🚀 UPLOAD & SINKRONKAN DATA REGULER", type="primary", use_container_width=True):
                    # Validasi Ketat
                    item_kurang = []
                    if not str(site_name_reg).strip(): item_kurang.append("Site ID")
                    if not teknisi_reg: item_kurang.append("Nama FME")
                    if not foto_prob: item_kurang.append("Foto Problem Alarm")
                    if not act_sel: item_kurang.append("Pilihan Action Dikerjakan")
                    if not foto_act: item_kurang.append("Foto Action / Perbaikan Utama")
                    
                    wajib_teks = []
                    wajib_foto = []
                    
                    if show_cond_power:
                        wajib_teks.extend([
                            ('v_rn', 'Isian Teg AC R-N'), ('v_sn', 'Isian Teg AC S-N'), ('v_tn', 'Isian Teg AC T-N'),
                            ('i_r', 'Isian Load AC R'), ('i_s', 'Isian Load AC S'), ('i_t', 'Isian Load AC T'),
                            ('load_dc', 'Isian Load DC Total'), ('set_bcl', 'Isian Settingan BCL'),
                            ('set_cap', 'Isian Capacity'), ('mod_rusak', 'Isian Module Kontroler')
                        ])
                        wajib_foto.extend([
                            ('f_v_rn', 'Foto Teg AC R-N'), ('f_v_sn', 'Foto Teg AC S-N'), ('f_v_tn', 'Foto Teg AC T-N'),
                            ('f_i_r', 'Foto Load AC R'), ('f_i_s', 'Foto Load AC S'), ('f_i_t', 'Foto Load AC T'),
                            ('f_ldc', 'Foto Load DC'), ('f_bcl', 'Foto BCL'), ('f_cap', 'Foto Capacity'), ('f_mod', 'Foto Module')
                        ])
                        
                    wajib_teks.append(('note_action', 'Isian Note / Remark Action'))
                    
                    for k_teks, label_teks in wajib_teks:
                        if not str(cond_data.get(k_teks, '')).strip(): item_kurang.append(label_teks)
                    for k_foto, label_foto in wajib_foto:
                        if cond_files.get(k_foto) is None: item_kurang.append(label_foto)

                    if item_kurang:
                        st.error(f"⚠️ **LAPORAN DITOLAK!** Tim wajib melengkapi seluruh data dan foto sesuai Action yang dipilih. Data yang masih kosong:\n- " + "\n- ".join(item_kurang))
                    else:
                        with st.spinner("⏳ Sedang memproses dan mengamankan seluruh foto eviden ke Cloud..."):
                            url_probs = upload_multiple_images(foto_prob, "reg_prob")
                            url_acts = upload_multiple_images(foto_act, "reg_act")
                            url_tw_full = upload_multiple_images(cond_files.get('f_tw_full'), "reg_tower")
                            url_tw_defect = upload_multiple_images(cond_files.get('f_tw_defect'), "reg_tower")
                            
                            cond_data['url_tw_full'] = url_tw_full
                            cond_data['url_tw_defect'] = url_tw_defect
                            
                            if show_cond_power:
                                for k_f in ['v_rn', 'v_sn', 'v_tn', 'i_r', 'i_s', 'i_t', 'ldc', 'bcl', 'cap', 'mod']:
                                    cond_data[f'url_{k_f}'] = upload_image(cond_files.get(f'f_{k_f}'), "reg_param")
                                cond_data['volt_ac_rst'] = f"{cond_data['v_rn']} / {cond_data['v_sn']} / {cond_data['v_tn']}"
                                cond_data['load_ac_rst'] = f"{cond_data['i_r']} / {cond_data['i_s']} / {cond_data['i_t']}"
                                
                            if show_cond_genset:
                                for k_g in ['gen_v', 'gen_i', 'gen_ats', 'gen_issue', 'gen_hm', 'gen_bbm']:
                                    cond_data[f'url_{k_g}'] = upload_image(cond_files.get(f'f_{k_g}'), "reg_param")
                            
                            tek_str_reg = ", ".join(teknisi_reg)
                            report_dict = {
                                "timestamp": datetime.datetime.now().strftime("%Y-%m-%d %H:%M:%S"),
                                "site_name": site_name_reg,
                                "kategori": "Site Reguler",
                                "nop": nop_area_reg,
                                "teknisi": tek_str_reg,
                                "role_fme": role_fme_reg,
                                "status": final_status_reg,
                                "action": f"Problem: {prob_alarm} | Action: {', '.join(act_sel)} | Note: {cond_data['note_action']}",
                                "sparepart": "-",
                                "problem_alarm": prob_alarm,
                                "action_reguler": act_sel,
                                "url_problem": url_probs,
                                "url_action": url_acts,
                                "cond_data": cond_data,
                                "pln_status": "-", "rect_brand": "-", "rect_out_v": "-", 
                                "total_load": cond_data.get('load_dc', '-'),
                                "url_sites": url_probs + url_acts + url_tw_full + url_tw_defect, 
                                "extras_fisik": [], "extras_panel": [], "extras_baterai": [], "extras_elektrikal": []
                            }
                            
                            sheet = connect_gsheets("Report Preventive")
                            if sheet:
                                row_data = [
                                    report_dict.get('timestamp', ''), report_dict.get('site_name', ''),
                                    report_dict.get('nop', ''), report_dict.get('teknisi', ''),
                                    report_dict.get('status', ''), json.dumps(report_dict)
                                ]
                                sheet.insert_row(row_data, index=2)
                                st.cache_data.clear()
                                st.session_state['laporan_db'].append(report_dict)
                                st.success("✅ BERHASIL! Data Laporan Site Reguler (beserta Audit Tower) telah diamankan ke Spreadsheet.")
                            else:
                                st.error("⚠️ Gagal menyambung ke Server Database.")
        elif site_category == "Project":
            # -------------------------------------------------------------------------
            # BLOK A.3: FORM KHUSUS PROJECT (PO & SVA) DENGAN PASSWORD
            # -------------------------------------------------------------------------
            if 'proj_unlocked' not in st.session_state:
                st.session_state['proj_unlocked'] = False
                
            if not st.session_state['proj_unlocked'] and st.session_state['role'] != 'Admin':
                with st.container(border=True):
                    st.markdown("<h3 style='text-align: center; color: #9C27B0;'>🔒 Akses Form Project Terkunci</h3>", unsafe_allow_html=True)
                    st.markdown("<p style='text-align: center;'>Form ini khusus untuk tim pengerjaan Project (PO / SVA).</p>", unsafe_allow_html=True)
                    proj_pass = st.text_input("Masukkan Kata Sandi Khusus Project:", type="password", key="pwd_proj")
                    if st.button("🔓 BUKA FORM PROJECT", type="primary", use_container_width=True):
                        import base64
                        if base64.b64encode(proj_pass.encode()).decode() == "WXVpMjAyNSM=":
                            st.session_state['proj_unlocked'] = True
                            st.rerun()
                        else:
                            st.error("❌ Kata sandi salah! Akses ditolak.")
            else:
                tabs_proj = st.tabs(["📌 1. Info & Multi-Site PO/SVA", "📸 2. Dokumentasi (Opsional Jika Belum Mulai)", "📤 3. Submit Project"])
                proj_data = {}
                proj_files = {}

                with tabs_proj[0]:
                    st.markdown("#### Kategori Kunjungan: Project Lapangan")
                    jenis_project = st.selectbox("Pilih Jenis Project:", ["PO", "SVA"])
                    proj_data['jenis_project'] = jenis_project

                    if fme_plk_list:
                        teknisi_proj = st.multiselect("Nama FME / Tim (Wajib):", fme_plk_list, key="proj_fme")
                    else:
                        t_str = st.text_input("Nama FME / Tim (Wajib):", key="proj_fme")
                        teknisi_proj = [t_str] if t_str else []

                    opsi_nop = ["Palangkaraya", "Pangkalan Bun", "Tarakan", "Pontianak", "Lainnya"]
                    nop_area_proj = st.selectbox("NOP Area:", opsi_nop, index=0, key="proj_nop")
                    st.divider()

                    if jenis_project == "PO":
                        st.markdown("### 📋 Detail Kontrak Purchase Order (PO)")
                        c_po1, c_po2 = st.columns(2)
                        with c_po1:
                            proj_data['no_po'] = st.text_input("Nomor PO *")
                            proj_data['mitra'] = st.text_input("Mitra Pemberi Kerja *", placeholder="Cth: Telkomsel / Huawei / Kisel...")
                        with c_po2:
                            nilai_po_str = st.text_input("Total Nilai PO (Rp) *", value="0", help="Ketik angka saja, misal: 50000000")
                            proj_data['nilai_po'] = nilai_po_str
                            tipe_pekerjaan = st.selectbox("Program Utama PO *", ["Relokasi NE", "Instalasi Imacro", "Perbaikan / Modernisasi", "Lainnya"])
                            proj_data['tipe_pekerjaan'] = tipe_pekerjaan

                        st.markdown("### 🗺️ Tabel Multi-Site, Rincian Pekerjaan, Timeline & Alokasi Biaya")
                        st.info("💡 **Klik baris di bawah (atau klik tanda `+` di baris paling bawah tabel)** untuk menambahkan banyak rute site sekaligus, termasuk **Dana Tersupport (Cair)** dan **Alasan jika Mandek**!")
                        
                        df_po = pd.DataFrame([
                            {
                                "Site Asal": "",
                                "Site Tujuan": "",
                                "Pekerjaan Spesifik": "",
                                "Budget Per Site": 0,
                                "Dana Tersupport": 0,
                                "Start Date": datetime.date.today().strftime("%Y-%m-%d"),
                                "End Date": datetime.date.today().strftime("%Y-%m-%d"),
                                "Progress (%)": 0,
                                "Status": "Not Started",
                                "Kategori Kendala": "Lancar / Normal",
                                "Remark": ""
                            }
                        ])
                        
                        edited_po = st.data_editor(
                            df_po, 
                            num_rows="dynamic", 
                            use_container_width=True,
                            column_config={
                                "Site Asal": st.column_config.TextColumn("Site Asal (Dari)"),
                                "Site Tujuan": st.column_config.TextColumn("Site Tujuan (Ke)"),
                                "Pekerjaan Spesifik": st.column_config.TextColumn("Pekerjaan di Site"),
                                "Budget Per Site": st.column_config.NumberColumn("Total Budget Site (Rp)", min_value=0, default=0, step=100000),
                                "Dana Tersupport": st.column_config.NumberColumn("Dana Tersupport / Cair (Rp)", min_value=0, default=0, step=100000),
                                "Start Date": st.column_config.TextColumn("Tgl Mulai (YYYY-MM-DD)"),
                                "End Date": st.column_config.TextColumn("Target Selesai (YYYY-MM-DD)"),
                                "Progress (%)": st.column_config.NumberColumn("Progress (%)", min_value=0, max_value=100, default=0),
                                "Status": st.column_config.SelectboxColumn("Status Site", options=["Not Started", "On Progress", "Done", "Pending Issue"]),
                                "Kategori Kendala": st.column_config.SelectboxColumn("Kenapa Mandek / Kendala", options=[
                                    "Lancar / Normal",
                                    "Mandek - Pending Keuangan (Dana Belum Cair)",
                                    "Mandek - Pending Material / Sparepart",
                                    "Mandek - Izin Site / Community Issue",
                                    "Mandek - Kendala Teknis / Cuaca"
                                ]),
                                "Remark": st.column_config.TextColumn("Detail Kendala / Catatan")
                            },
                            key="tabel_po_input"
                        )
                        
                        try:
                            val_po = float(str(nilai_po_str).replace('.', '').replace(',', ''))
                        except:
                            val_po = 0.0
                            
                        df_clean_po = edited_po.copy()
                        if not df_clean_po.empty:
                            df_clean_po["Budget Per Site"] = pd.to_numeric(df_clean_po["Budget Per Site"], errors='coerce').fillna(0.0)
                            df_clean_po["Dana Tersupport"] = pd.to_numeric(df_clean_po["Dana Tersupport"], errors='coerce').fillna(0.0)
                            df_clean_po["Progress (%)"] = pd.to_numeric(df_clean_po["Progress (%)"], errors='coerce').fillna(0.0)
                            
                        total_cost = df_clean_po["Budget Per Site"].sum() if not df_clean_po.empty else 0.0
                        total_sup = df_clean_po["Dana Tersupport"].sum() if not df_clean_po.empty else 0.0
                        pending_fin = max(total_cost - total_sup, 0.0)
                        margin = val_po - total_cost
                        avg_progress = df_clean_po["Progress (%)"].mean() if not df_clean_po.empty else 0.0
                        
                        cm1, cm2, cm3, cm4, cm5 = st.columns(5)
                        cm1.metric("💰 Total Nilai PO", f"Rp {val_po:,.0f}")
                        cm2.metric("💸 Total Budget Site", f"Rp {total_cost:,.0f}")
                        cm3.metric("✅ Dana Tersupport", f"Rp {total_sup:,.0f}")
                        cm4.metric("🚨 Pending Keuangan", f"Rp {pending_fin:,.0f}")
                        cm5.metric("📈 Estimasi Margin", f"Rp {margin:,.0f}", delta=f"{avg_progress:.0f}% Progress")
                        
                        proj_data['tabel_pekerjaan'] = df_clean_po.astype(str).to_dict('records')
                        proj_data['total_cost'] = float(total_cost)
                        proj_data['total_support'] = float(total_sup)
                        proj_data['pending_finance'] = float(pending_fin)
                        proj_data['margin'] = float(margin)
                        proj_data['avg_progress'] = float(avg_progress)

                        if tipe_pekerjaan == "Instalasi Imacro":
                            st.markdown("**📌 Checklist Milestone Imacro:**")
                            c_chk1, c_chk2 = st.columns(2)
                            with c_chk1:
                                p1 = st.checkbox("Dismantle")
                                p2 = st.checkbox("Instalasi")
                                p3 = st.checkbox("Penarikan Optik")
                            with c_chk2:
                                p4 = st.checkbox("Material Terkirim/Lengkap")
                                p5 = st.checkbox("PSB PLN")
                            proj_data['progress_imacro'] = [k for k, v in zip(["Dismantle", "Instalasi", "Penarikan Optik", "Material", "PSB PLN"], [p1, p2, p3, p4, p5]) if v]

                    elif jenis_project == "SVA":
                        st.markdown("### 📋 Detail Site Verification / Audit (SVA)")
                        c_sva1, c_sva2 = st.columns(2)
                        with c_sva1:
                            proj_data['site_sva'] = st.selectbox("Pilih Site ID *", [""] + site_plk_list, key="sva_site") if site_plk_list else st.text_input("Site ID *")
                            proj_data['no_tiket'] = st.text_input("Nomor Tiket SVA *")
                        with c_sva2:
                            tgl_sva_val = st.date_input("Tanggal Pengerjaan SVA *", value=datetime.date.today())
                            proj_data['tgl_sva'] = tgl_sva_val.strftime("%Y-%m-%d")
                            proj_data['status_sva'] = st.selectbox("Status SVA", ["Not Started", "On Progress", "Done", "Pending Issue"])

                        proj_data['temuan_sva'] = st.text_area("Temuan / Hasil Audit Lapangan:", placeholder="Catat detail temuan SVA di sini...")

                with tabs_proj[1]:
                    st.markdown("### 📸 Lampiran Dokumentasi Project")
                    st.info("💡 ** Catatan Penting:** Jika pekerjaan belum dimulai (*Not Started*) atau baru perencanaan PO, foto di bawah ini **BOLEH DIKOSONGKAN** dan dapat di-update menyusul di *Executive Project Dashboard*.")

                    if jenis_project == "PO":
                        foto_awal = st.file_uploader("📸 Foto Kondisi Awal (Before)", accept_multiple_files=True, key="p_awal")
                        foto_proses = st.file_uploader("📸 Foto Proses Pengerjaan (On Progress)", accept_multiple_files=True, key="p_pros")
                        foto_akhir = st.file_uploader("📸 Foto Hasil Akhir (After) / BAST", accept_multiple_files=True, key="p_akhir")
                        proj_files['foto_awal'] = foto_awal
                        proj_files['foto_proses'] = foto_proses
                        proj_files['foto_akhir'] = foto_akhir
                    else:
                        foto_sva_1 = st.file_uploader("📸 Foto View Site & Perangkat", accept_multiple_files=True, key="sva_1")
                        foto_sva_2 = st.file_uploader("📸 Foto Temuan SVA / Defect", accept_multiple_files=True, key="sva_2")
                        proj_files['foto_sva_1'] = foto_sva_1
                        proj_files['foto_sva_2'] = foto_sva_2

                with tabs_proj[2]:
                    st.markdown("### 📤 Konfirmasi & Submit Laporan Project")
                    proj_data['note_akhir'] = st.text_area("📝 Catatan Akhir Project:", placeholder="Kendala atau catatan penting lainnya...")
                    final_status_proj = st.radio("Status Keseluruhan Project:", ["Not Started", "On Progress", "Done", "Pending"], key="stat_proj")

                    if st.button("🚀 SIMPAN & SINKRONKAN PROJECT KE CLOUD", type="primary", use_container_width=True):
                        item_kurang = []
                        if not teknisi_proj: item_kurang.append("Nama FME/Tim")

                        if jenis_project == "PO":
                            if not proj_data.get('no_po'): item_kurang.append("Nomor PO")
                            if not proj_data.get('mitra'): item_kurang.append("Mitra Pemberi Kerja")
                            if final_status_proj == "Done" and not proj_files.get('foto_akhir') and not proj_files.get('foto_proses'):
                                item_kurang.append("Minimal Foto Proses / Akhir (Wajib karena status dipilih Done)")
                        else:
                            if not proj_data.get('site_sva'): item_kurang.append("Site ID SVA")
                            if not proj_data.get('no_tiket'): item_kurang.append("Nomor Tiket SVA")

                        if item_kurang:
                            st.error(f"⚠️ **LAPORAN DITOLAK!** Mohon lengkapi isian berikut:\n- " + "\n- ".join(item_kurang))
                        else:
                            with st.spinner("⏳ Sedang memproses dan mengamankan data Project ke Cloud..."):
                                url_awal, url_proses, url_akhir = [], [], []
                                url_sva1, url_sva2 = [], []

                                if jenis_project == "PO":
                                    url_awal = upload_multiple_images(proj_files.get('foto_awal'), "proj_po")
                                    url_proses = upload_multiple_images(proj_files.get('foto_proses'), "proj_po")
                                    url_akhir = upload_multiple_images(proj_files.get('foto_akhir'), "proj_po")
                                else:
                                    url_sva1 = upload_multiple_images(proj_files.get('foto_sva_1'), "proj_sva")
                                    url_sva2 = upload_multiple_images(proj_files.get('foto_sva_2'), "proj_sva")

                                proj_data['url_awal'] = url_awal
                                proj_data['url_proses'] = url_proses
                                proj_data['url_akhir'] = url_akhir
                                proj_data['url_sva1'] = url_sva1
                                proj_data['url_sva2'] = url_sva2

                                tek_str_proj = ", ".join(teknisi_proj)
                                nama_site_db = proj_data.get('site_sva') if jenis_project == "SVA" else f"PO: {proj_data.get('no_po')}"

                                report_dict = {
                                    "timestamp": datetime.datetime.now().strftime("%Y-%m-%d %H:%M:%S"),
                                    "site_name": nama_site_db,
                                    "kategori": "Project",
                                    "nop": nop_area_proj,
                                    "teknisi": tek_str_proj,
                                    "status": final_status_proj,
                                    "action": f"Project {jenis_project}",
                                    "sparepart": "-",
                                    "proj_data": proj_data,
                                    "url_sites": url_awal + url_sva1,
                                    "url_action": url_proses + url_akhir + url_sva2,
                                    "pln_status": "-", "rect_brand": "-", "rect_out_v": "-",
                                    "total_load": "-", "cond_data": {}
                                }

                                sheet = connect_gsheets("Project")
                                if sheet:
                                    row_data = [
                                        report_dict.get('timestamp', ''), report_dict.get('site_name', ''),
                                        report_dict.get('nop', ''), report_dict.get('teknisi', ''),
                                        report_dict.get('status', ''), json.dumps(report_dict, default=str)
                                    ]
                                    try:
                                        sheet.insert_row(row_data, index=2)
                                    except Exception:
                                        sheet.insert_row(["Timestamp", "Site Name", "NOP", "Teknisi", "Status", "Data JSON"], index=1)
                                        sheet.insert_row(row_data, index=2)
                                        
                                    st.cache_data.clear()
                                    st.session_state['laporan_db'] = fetch_data_from_gsheets()
                                    st.success("✅ BERHASIL! Data Project telah tersimpan dan siap dipantau di menu 💎 Executive Project.")
                                else:
                                    st.error("⚠️ Gagal menyambung ke Sheet 'Project'. Pastikan sheet bernama 'Project' sudah dibuat di Spreadsheet Utama Anda.")

    # -------------------------------------------------------------------------
    # BLOK B: FORM PANGKALAN BUN (LENGKAP DENGAN UPLOAD FOTO & AUTO WA)
    # -------------------------------------------------------------------------
    elif standar_form == "Pangkalan Bun":
        
        st.info("💡 **Format Khusus Site Visit Report (Pangkalan Bun).** Silakan isi parameter di bawah ini dan lampirkan foto dokumentasi pada setiap bagian. Seluruh foto akan tersimpan di Cloud & muncul di Dashboard.")
        tabs_pb = st.tabs(["📌 Info Site & Power", "🔌 Rectifier & ENVA", "⚡ Kelistrikan & ACPDB", "⛽ Genset & Log", "📤 Activity & Submit"])
        
        pb_payload = {}
        
        with tabs_pb[0]:
            st.markdown("### 📌 INFORMASI SITE")
            c1, c2 = st.columns(2)
            with c1:
                if fme_pbu_list:
                    pb_te = st.multiselect("FME NAME (Bisa >1):", fme_pbu_list, key="pb_te")
                else:
                    pt_str = st.text_input("FME NAME", key="pb_te")
                    pb_te = [pt_str] if pt_str else []
                    
                if site_pbu_list:
                    pb_sid = st.selectbox("Pilih Site ID", [""] + site_pbu_list, key="pb_sid")
                else:
                    pb_sid = st.text_input("Site ID", key="pb_sid")
                    
            with c2:
                pb_tick = st.text_input("Ticket#")
                pb_tgl = st.text_input("Tanggal On Site", value=datetime.date.today().strftime("%d %B %Y"))
                pb_catuan = st.selectbox("Catuan Power Utama", ["GENSET", "PLN", "SPS", "LAINNYA"])
                
            pb_foto_site = st.file_uploader("📸 Foto Site / Kondisi Umum (Bisa >1)", accept_multiple_files=True, key="pb_foto_site")
            
            tek_pb_str = ", ".join(pb_te) if pb_te else "-"
            pb_payload.update({"te_name": tek_pb_str, "site_id": pb_sid, "site_name": pb_sid, "ticket": pb_tick, "tgl_onsite": pb_tgl, "power_utama": pb_catuan})
            
        with tabs_pb[1]:
            st.markdown("### 🔌 DATA RECTIFIER")
            pb_jml_rect = st.number_input("Jumlah Rectifier", min_value=1, value=2)
            
            pb_recti_list = []
            for i in range(int(pb_jml_rect)):
                with st.expander(f"⚙️ Rectifier #{i+1}", expanded=(i==0)):
                    c1, c2 = st.columns(2)
                    with c1:
                        tr = st.text_input("Tipe Rectifier (ZTE, dll)", key=f"tr_{i}")
                        tc = st.text_input("Type Controler (Csu 501B, dll)", key=f"tc_{i}")
                        tm = st.text_input("Type Module (ZXD2000, dll)", key=f"tm_{i}")
                        jb = st.text_input("Jumlah Battery (Bank)", key=f"jb_{i}")
                        bok = st.text_input("Jenis Batt OK", key=f"bok_{i}")
                        bnok = st.text_input("Jenis Batt NOK", key=f"bnok_{i}")
                        jmo = st.text_input("Jumlah Module OK", key=f"jmo_{i}")
                        jmn = st.text_input("Jumlah Module NOK", key=f"jmn_{i}")
                    with c2:
                        bcl = st.text_input("BCL (%)", key=f"bcl_{i}")
                        lc = st.text_input("Load Current (A)", key=f"lc_{i}")
                        dcv = st.text_input("DC Voltage (V)", key=f"dcv_{i}")
                        lp = st.text_input("LVD Prio", key=f"lp_{i}")
                        lnp = st.text_input("LVD Non Prio", key=f"lnp_{i}")
                        elp = st.text_input("EQP LVD Prio", key=f"elp_{i}")
                        elnp = st.text_input("EQP LVD Non Prio", key=f"elnp_{i}")
                        bbt = st.text_input("Battery Backup Time", key=f"bbt_{i}")
                    
                    pb_recti_list.append({
                        "tipe": tr, "controler": tc, "module": tm, "jml_batt": jb, "batt_ok": bok, "batt_nok": bnok,
                        "mod_ok": jmo, "mod_nok": jmn, "bcl": bcl, "load": lc, "dc_v": dcv, "lvd_p": lp, "lvd_np": lnp, "eqp_p": elp, "eqp_np": elnp, "backup": bbt
                    })
            
            st.divider()
            st.markdown("### 🚨 ENVA Setelah Simulasi (Bukan Cabut Kabel)")
            opt_enva = ["OK", "NOK", "-"]
            c1, c2 = st.columns(2)
            with c1:
                em = st.selectbox("Main Fail", opt_enva, key="em")
                er = st.selectbox("Rectifier Fail", opt_enva, key="er")
                el = st.selectbox("Low Batt", opt_enva, key="el")
            with c2:
                eh = st.selectbox("High Temp", opt_enva, key="eh")
                egr = st.selectbox("Genset Run", opt_enva, key="egr")
                egf = st.selectbox("Genset Fail", opt_enva, key="egf")
                
            pb_foto_recti = st.file_uploader("📸 Foto Rectifier, Baterai & ENVA (Bisa >1)", accept_multiple_files=True, key="pb_foto_recti")
            pb_payload["enva"] = {"main_fail": em, "rect_fail": er, "low_batt": el, "high_temp": eh, "gen_run": egr, "gen_fail": egf}

        with tabs_pb[2]:
            st.markdown("### ⚡ DATA KELISTRIKAN")
            c1, c2 = st.columns(2)
            with c1:
                kd = st.text_input("Daya KWH", key="kd")
                kj = st.text_input("Jumlah Phasa", key="kj")
                kv1 = st.text_input("Tegangan R-S-T thd N", key="kv1")
            with c2:
                kv2 = st.text_input("Tegangan Phasa RS-RT-ST", key="kv2")
                ka = st.text_input("Pengukuran Arus R-S-T", key="ka")
                kgn = st.text_input("Tegangan Pengukuran G-N", key="kgn")
            pb_payload["kelistrikan"] = {"daya": kd, "phasa": kj, "v_rst_n": kv1, "v_rs_rt_st": kv2, "arus_rst": ka, "v_gn": kgn}
            
            st.divider()
            st.markdown("### 🎛 DATA ACPDB")
            c3, c4 = st.columns(2)
            with c3:
                aa = st.selectbox("Ada ACPDB", ["Ya", "Tidak", "-"], key="aa")
                av1 = st.text_input("Tegangan R-S-T thd N (ACPDB)", key="av1")
                av2 = st.text_input("Tegangan Phasa RS-RT-ST (ACPDB)", key="av2")
            with c4:
                aar = st.text_input("Pengukuran Arus R-S-T (ACPDB)", key="aar")
                agn = st.text_input("Tegangan Pengukuran G-N (ACPDB)", key="agn")
            
            pb_foto_listrik = st.file_uploader("📸 Foto Kelistrikan & ACPDB (Bisa >1)", accept_multiple_files=True, key="pb_foto_listrik")
            pb_payload["acpdb"] = {"ada": aa, "v_rst_n": av1, "v_rs_rt_st": av2, "arus_rst": aar, "v_gn": agn}
            
        with tabs_pb[3]:
            st.markdown("### ⛽ DATA GENSET")
            c1, c2 = st.columns(2)
            with c1:
                ga = st.selectbox("Ada Genset", ["Ya", "Tidak"], key="ga")
                gc = st.text_input("Capacity Genset (kVA)", key="gc")
                gs = st.text_input("Genset Status (OK/NOK)", key="gs")
                gco = st.text_input("Config Running", key="gco")
            with c2:
                gt = st.text_input("Kapasitas Tangki", key="gt")
                grh = st.text_input("Running Hour", key="grh")
                gsb = st.text_input("Sisa BBM", key="gsb")
                
            pb_foto_genset = st.file_uploader("📸 Foto Genset & Tangki BBM (Bisa >1)", accept_multiple_files=True, key="pb_foto_genset")
            pb_payload["genset"] = {"ada": ga, "cap": gc, "status": gs, "config": gco, "tangki": gt, "rh": grh, "bbm": gsb}

        with tabs_pb[4]:
            st.markdown("### 📝 ACTIVITY REPORT & CATATAN")
            pb_act1 = st.multiselect("1. Power & CME Activity", ["PMS/ PMG", "Pembersihan/ PM SPS", "Regional Programm", "NOP Programm", "Site Down Troubleshoot", "PLN Undervolt", "Load Balancing", "Rectifier Troubleshoot", "Battery Troubleshoot/ Claim Warranty", "Pergantian Rectifier Controller (SC/ CSU)", "Mapping Rectifier Controller (SC/ CSU)", "Grounding"])
            pb_act2 = st.multiselect("2. Transport Activity", ["Packet Loss", "NMS Unremote", "On Site VLAN Balancing"])
            pb_act3 = st.multiselect("3. OSA Impact Service Activity", ["Validasi ENVA", "BBU Min Voltage", "RRU Min Voltage", "Cell Down", "Disable PA", "Int. Fault", "Cell Flicker", "BBU Temp", "RRU Temp", "Zero Traffic", "Zero Payload", "VSWR", "Optical port link fault", "GNSS/GPS", "BFAN/FCE"])
            pb_temuan = st.text_area("Temuan/ Catatan (Gunakan Enter untuk list ke bawah):", placeholder="1. \n2. \n3.")
            
            pb_foto_act = st.file_uploader("📸 Foto Temuan / Action Lapangan (Bisa >1)", accept_multiple_files=True, key="pb_foto_act")
            pb_status_akhir = st.radio("Status Akhir Site:", ["Normal", "Minor Issue", "Major/Critical"])
            
            if st.button("🚀 SUBMIT SITE VISIT & BUKA WA", type="primary", use_container_width=True):
                if not pb_sid: st.error("⚠️ Mohon pilih Site ID!")
                elif not pb_te: st.error("⚠️ Mohon pilih FME NAME!")
                else:
                    with st.spinner("⏳ Sedang mengunggah foto ke Cloud & menyimpan data Laporan..."):
                        
                        url_pb_site = upload_multiple_images(pb_foto_site, "pb_site")
                        url_pb_recti = upload_multiple_images(pb_foto_recti, "pb_recti")
                        url_pb_listrik = upload_multiple_images(pb_foto_listrik, "pb_listrik")
                        url_pb_genset = upload_multiple_images(pb_foto_genset, "pb_genset")
                        url_pb_act = upload_multiple_images(pb_foto_act, "pb_action")
                        
                        report_dict = {
                            "timestamp": datetime.datetime.now().strftime("%Y-%m-%d %H:%M:%S"),
                            "site_name": pb_sid,
                            "kategori": "Pangkalan Bun", 
                            "nop": "Pangkalan Bun", 
                            "teknisi": tek_pb_str,
                            "status": pb_status_akhir, 
                            "action": pb_temuan,
                            
                            "pb_te": tek_pb_str, "pb_site_id": pb_sid, "pb_ticket": pb_tick, "pb_power": pb_catuan,
                            "pb_rectifiers": pb_recti_list,
                            "pb_enva": pb_payload["enva"],
                            "pb_kelistrikan": pb_payload["kelistrikan"],
                            "pb_acpdb": pb_payload["acpdb"],
                            "pb_genset": pb_payload["genset"],
                            "pb_act_power": pb_act1, "pb_act_trans": pb_act2, "pb_act_osa": pb_act3,
                            
                            "pb_url_site": url_pb_site,
                            "pb_url_recti": url_pb_recti,
                            "pb_url_listrik": url_pb_listrik,
                            "pb_url_genset": url_pb_genset,
                            "pb_url_act": url_pb_act,
                            
                            "panel_data": [], "battery_data": [], "url_sites": [], "total_load": "-", "rect_out_v": "-", "sparepart": "-"
                        }
                        
                        sheet_pbu = connect_gsheets("Report PBU")
                        if sheet_pbu:
                            row_data = [
                                report_dict['timestamp'], report_dict['site_name'], report_dict['nop'], 
                                report_dict['teknisi'], report_dict['status'], json.dumps(report_dict)
                            ]
                            sheet_pbu.insert_row(row_data, index=2)
                            st.cache_data.clear()
                            st.session_state['laporan_db'] = fetch_data_from_gsheets()
                            
                            wa_text = generate_wa_pbu(report_dict)
                            wa_url = f"https://wa.me/?text={urllib.parse.quote(wa_text)}"
                            
                            st.success("✅ BERHASIL! Data Visit & Dokumentasi Foto Pangkalan Bun telah tersimpan di Baris Teratas.")
                            st.markdown("### 📱 Langkah Terakhir: Bagikan Laporan")
                            st.markdown("Silakan tekan tombol di bawah ini untuk langsung membagikan rangkuman laporan Anda ke Grup WhatsApp.")
                            st.link_button("🚀 BUKA WHATSAPP & KIRIM LAPORAN", wa_url, type="primary")
                            
                        else:
                            st.error("⚠️ Gagal menyambung ke Server Database.")


# =========================================================================
# MENU 2: HASIL LAPORAN & DASHBOARD (DENGAN PILIHAN AREA ELEGAN)
# =========================================================================
elif menu == "📊 Hasil Laporan & Dashboard":
    
    st.markdown("<h1 style='text-align: center;'>📊 Executive Dashboard & Analytics Report</h1>", unsafe_allow_html=True)
    st.markdown("<p style='text-align: center; opacity: 0.8; margin-bottom: 25px;'>Sistem Monitoring Terpadu Hasil Preventive Maintenance & Site Visit</p>", unsafe_allow_html=True)
    
    db = st.session_state['laporan_db']

    if not db:
        st.warning("⚠ Belum ada data di Server / Koneksi Sedang Proses.")
    else:
        idx_plk = [idx for idx, r in enumerate(db) if r.get('kategori') not in ['Pangkalan Bun', 'Project']]
        idx_pbu = [idx for idx, r in enumerate(db) if r.get('kategori') == 'Pangkalan Bun']

        # ---------------------------------------------------------------------
        # LAYAR PILIHAN AREA DASHBOARD (AGAR TIDAK NUMPUK)
        # ---------------------------------------------------------------------
        if st.session_state.get('dashboard_area') is None:
            col_m1, col_m2, col_m3 = st.columns(3)
            with col_m1: st.metric("Total Seluruh Laporan", f"{len(db)} Site")
            with col_m2: st.metric("Laporan NOP Palangkaraya", f"{len(idx_plk)} Site")
            with col_m3: st.metric("Laporan NOP Pangkalan Bun", f"{len(idx_pbu)} Site")
            
            st.markdown("<hr style='margin: 22px 0;'>", unsafe_allow_html=True)
            st.markdown("<h3 style='text-align: center; margin-bottom: 25px;'>📍 Pilih Area Operasional untuk Membuka Daftar Site</h3>", unsafe_allow_html=True)

            col_d1, col_d_space, col_d2 = st.columns([4.5, 0.4, 4.5])

            with col_d1:
                st.markdown(f"""
                    <div class='area-card' style='border-top-color: #007BFF;'>
                        <div style='font-size: 50px; margin-bottom: 10px;'>🗼</div>
                        <div class='area-badge'>AREA OPERASIONAL 1</div>
                        <div class='login-title'>NOP PALANGKARAYA</div>
                        <div class='login-desc'>Daftar Site Preventive Maintenance (SPS & Reguler), Laporan Project, Dokumentasi Foto Modul Surya, Baterai, Datalog & Berita Acara PDF.</div>
                        <div style='font-size: 20px; font-weight: 800; color: #007BFF; margin-bottom: 8px;'>📊 {len(idx_plk)} Laporan Tersedia</div>
                    </div>
                """, unsafe_allow_html=True)
                if st.button("🔵 BUKA DASHBOARD PALANGKARAYA", type="secondary", use_container_width=True, key="btn_open_plk"):
                    st.session_state['dashboard_area'] = "Palangkaraya"
                    st.rerun()

            with col_d2:
                st.markdown(f"""
                    <div class='area-card' style='border-top-color: #FF3333;'>
                        <div style='font-size: 50px; margin-bottom: 10px;'>📡</div>
                        <div class='area-badge' style='color: #FF3333; border-color: #FF3333; background-color: rgba(255,51,51,0.12);'>AREA OPERASIONAL 2</div>
                        <div class='login-title'>NOP PANGKALAN BUN</div>
                        <div class='login-desc'>Daftar Site Visit Report FME, Parameter Rectifier, Simulasi ENVA, Kelistrikan, ACPDB, Genset & Dokumentasi Foto Lapangan.</div>
                        <div style='font-size: 20px; font-weight: 800; color: #FF3333; margin-bottom: 8px;'>📊 {len(idx_pbu)} Laporan Tersedia</div>
                    </div>
                """, unsafe_allow_html=True)
                if st.button("🔴 BUKA DASHBOARD PANGKALAN BUN", type="primary", use_container_width=True, key="btn_open_pbu"):
                    st.session_state['dashboard_area'] = "Pangkalan Bun"
                    st.rerun()

            st.stop()

        # ---------------------------------------------------------------------
        # JIKA AREA SUDAH DIPILIH -> TAMPILKAN NAVIGASI KEMBALI & DAFTAR SITE
        # ---------------------------------------------------------------------
        aktif_area = st.session_state.get('dashboard_area')
        
        c_nav1, c_nav2 = st.columns([1.5, 3.5])
        with c_nav1:
            if st.button("⬅️ Ganti Area Operasional Dashboard", use_container_width=True):
                st.session_state['dashboard_area'] = None
                st.rerun()
        with c_nav2:
            warna_area = "#007BFF" if aktif_area == "Palangkaraya" else "#FF3333"
            st.markdown(
                f"<div style='padding: 8px 18px; border-radius: 10px; border-left: 6px solid {warna_area}; background-color: var(--secondary-background-color); font-size: 18px; font-weight: 800;'>📍 Menampilkan Data NOP {aktif_area.upper()}</div>",
                unsafe_allow_html=True
            )

        st.markdown("<hr style='margin: 18px 0;'>", unsafe_allow_html=True)

        # Filter berdasarkan Area yang dipilih + Pencarian Cepat
        if aktif_area == "Palangkaraya":
            c_flt1, c_flt2 = st.columns([2, 2])
            with c_flt1:
                filter_kat = st.radio("Sub-Kategori Site Palangkaraya:", ["Semua (Palangkaraya)", "SPS", "Site Reguler"], horizontal=True)
            with c_flt2:
                cari_site = st.text_input("🔍 Cari Nama Site / Teknisi:", placeholder="Ketik ID atau Nama Site...")
                
            filtered_indices = []
            for idx in idx_plk:
                r = db[idx]
                kat = r.get('kategori', 'SPS')
                match_kat = (filter_kat == "Semua (Palangkaraya)") or (kat == filter_kat)
                match_cari = True
                if cari_site.strip():
                    q = cari_site.strip().lower()
                    match_cari = (q in str(r.get('site_name', '')).lower()) or (q in str(r.get('teknisi', '')).lower())
                if match_kat and match_cari:
                    filtered_indices.append(idx)
        else:
            filter_kat = "Pangkalan Bun"
            cari_site = st.text_input("🔍 Cari Site ID / Nama Site / FME Name (Pangkalan Bun):", placeholder="Ketik Site ID atau Nama FME...")
            filtered_indices = []
            for idx in idx_pbu:
                r = db[idx]
                match_cari = True
                if cari_site.strip():
                    q = cari_site.strip().lower()
                    match_cari = (q in str(r.get('site_name', '')).lower()) or (q in str(r.get('pb_te', '')).lower()) or (q in str(r.get('pb_ticket', '')).lower())
                if match_cari:
                    filtered_indices.append(idx)

        # Metrik Ringkasan Area Aktif
        export_db = [db[idx] for idx in filtered_indices]
        jml_normal = sum(1 for x in export_db if x.get('status') in ['Normal', 'Done'])
        jml_issue = len(export_db) - jml_normal

        cm1, cm2, cm3 = st.columns(3)
        cm1.metric(f"Total Laporan ({aktif_area})", f"{len(export_db)} Berkas")
        cm2.metric("🟢 Status Normal / Done", f"{jml_normal} Berkas")
        cm3.metric("🟡/🔴 Issue / Progress", f"{jml_issue} Berkas")

        if len(filtered_indices) == 0:
            st.info(f"Tidak ada data laporan yang cocok untuk filter saat ini di area {aktif_area}.")
            st.stop()

        st.markdown("<br>", unsafe_allow_html=True)

        # ---------------------------------------------------------------------
        # TABEL EXCEL FLAT & POWERPOINT (SESUAI AREA AKTIF)
        # ---------------------------------------------------------------------
        with st.container(border=True):
            st.markdown(f"#### 📥 Ekspor Rekapitulasi Laporan — {aktif_area}")
            st.info("💡 Untuk mempercepat kinerja Dashboard, file Excel dan PPTX akan dirakit otomatis saat Anda menekan tombol di bawah ini.")
            
            if st.button("🛠️ Siapkan File Excel & PPTX untuk Diunduh", type="primary", use_container_width=True):
                with st.spinner("⏳ Sedang merakit dokumen resolusi tinggi, mohon tunggu sebentar..."):
                    summary_list = []
                    for r in export_db:
                        if r.get('kategori') == 'Pangkalan Bun':
                            kel = r.get('pb_kelistrikan', {})
                            acp = r.get('pb_acpdb', {})
                            enva = r.get('pb_enva', {})
                            gen = r.get('pb_genset', {})
                            summary_list.append({
                                "Timestamp": r.get('timestamp', '-'), "Nama Site": r.get('site_name', '-'),
                                "Kategori": "Pangkalan Bun", "FME NAME": r.get('pb_te', '-'),
                                "Ticket#": r.get('pb_ticket', '-'), "Catuan Power": r.get('pb_power', '-'),
                                "Status Akhir": r.get('status', '-'), "Main Fail (ENVA)": enva.get('main_fail', '-'),
                                "Rect Fail (ENVA)": enva.get('rect_fail', '-'), "Low Batt (ENVA)": enva.get('low_batt', '-'),
                                "High Temp (ENVA)": enva.get('high_temp', '-'), "Genset Run (ENVA)": enva.get('gen_run', '-'),
                                "Genset Fail (ENVA)": enva.get('gen_fail', '-'), "KWH Daya": kel.get('daya', '-'),
                                "KWH Phasa": kel.get('phasa', '-'), "Tegangan RST-N": kel.get('v_rst_n', '-'),
                                "Tegangan RS-RT-ST": kel.get('v_rs_rt_st', '-'), "Arus RST": kel.get('arus_rst', '-'),
                                "Tegangan G-N": kel.get('v_gn', '-'), "ACPDB Ada": acp.get('ada', '-'),
                                "ACPDB RST-N": acp.get('v_rst_n', '-'), "ACPDB RS-RT-ST": acp.get('v_rs_rt_st', '-'),
                                "ACPDB Arus": acp.get('arus_rst', '-'), "ACPDB G-N": acp.get('v_gn', '-'),
                                "Genset Ada": gen.get('ada', '-'), "Genset Cap": gen.get('cap', '-'),
                                "Genset Status": gen.get('status', '-'), "Genset Config": gen.get('config', '-'),
                                "Genset Tangki": gen.get('tangki', '-'), "Genset RH": gen.get('rh', '-'),
                                "Genset BBM": gen.get('bbm', '-'), "Power & CME Act": ", ".join(r.get('pb_act_power', [])) if isinstance(r.get('pb_act_power'), list) else str(r.get('pb_act_power','-')),
                                "Transport Act": ", ".join(r.get('pb_act_trans', [])) if isinstance(r.get('pb_act_trans'), list) else str(r.get('pb_act_trans','-')),
                                "OSA Act": ", ".join(r.get('pb_act_osa', [])) if isinstance(r.get('pb_act_osa'), list) else str(r.get('pb_act_osa','-')),
                                "Catatan/Temuan": r.get('action', '-')
                            })
                        else:
                            cond_d = r.get('cond_data', {})
                            pd_d = r.get('proj_data', {})
                            summary_list.append({
                                "Timestamp": r.get('timestamp', '-'), "Nama Site": r.get('site_name', '-'),
                                "Kategori": r.get('kategori', 'SPS'), "NOP": r.get('nop', '-'),
                                "Teknisi / FME": r.get('teknisi', '-'), "Status Akhir": r.get('status', '-'),
                                "Load Beban (A)": r.get('total_load', '-'), "PLN": r.get('pln_status', '-'),
                                "Rectifier (Merek)": r.get('rect_brand', '-'), "Voltase Recti (V)": r.get('rect_out_v', '-'),
                                "Genset": r.get('genset_status', '-'), "Level BBM (%)": r.get('fuel_pct', '-'),
                                "Action / Pekerjaan": r.get('action', '-'), "Sparepart Diganti": r.get('sparepart', '-'),
                                "Remark": r.get('remark', '-'),
                                "Reguler - Teg AC (RN/SN/TN)": f"{cond_d.get('v_rn','-')} / {cond_d.get('v_sn','-')} / {cond_d.get('v_tn','-')}",
                                "Reguler - Arus AC (R/S/T)": f"{cond_d.get('i_r','-')} / {cond_d.get('i_s','-')} / {cond_d.get('i_t','-')}",
                                "Reguler - Load DC": cond_d.get('load_dc', '-'),
                                "Reguler - Problem Alarm": r.get('problem_alarm', '-'),
                                "Reguler - Note/Remark": cond_d.get('note_action', '-')
                            })
                    
                    df_export = pd.DataFrame(summary_list)
                    col_dl_ex, col_dl_ppt = st.columns(2)
                    with col_dl_ex:
                        try:
                            excel_buffer = BytesIO()
                            with pd.ExcelWriter(excel_buffer, engine='xlsxwriter') as writer:
                                df_export.to_excel(writer, index=False, sheet_name='Database PM')
                            st.download_button(label=f"📊 Download Excel ({aktif_area})", data=excel_buffer.getvalue(), file_name=f"Database_PM_{aktif_area}_{datetime.date.today()}.xlsx", mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet", type="primary", use_container_width=True)
                        except:
                            st.download_button(label="📥 Download Laporan (CSV)", data=df_export.to_csv(index=False).encode('utf-8'), file_name=f"Database_PM_{datetime.date.today()}.csv", mime="text/csv", type="secondary")

                    with col_dl_ppt:
                        if aktif_area == "Palangkaraya" and HAS_PPTX:
                            pptx_data = build_pptx([x for x in export_db if x.get('kategori') not in ['Pangkalan Bun', 'Site Reguler', 'Project']])
                            st.download_button(label="📽 Download Presentasi (PPTX)", data=pptx_data, file_name=f"Report_PM_KUT_{datetime.date.today()}.pptx", mime="application/vnd.openxmlformats-officedocument.presentationml.presentation", type="primary", use_container_width=True)
                        elif aktif_area == "Pangkalan Bun":
                            st.info("💡 Rekap parameter Pangkalan Bun diunduh via Excel.")

        st.markdown("<br>", unsafe_allow_html=True)
        tab_chart, tab_list = st.tabs(["📈 Analisa & Grafik Laporan", "📋 Daftar Detail Laporan Site"])
        
        with tab_chart:
            st.markdown("### 📈 Executive Analytics & Repeated Issue Tracker")
            st.info("💡 Sistem merekap frekuensi kunjungan lapangan untuk mendeteksi *Repeated Issue* dan merangkum histori problem beserta tren beban power pada site tersebut.")
            
            site_counts = {}
            power_data = []
            for r in export_db:
                sn = str(r.get('site_name', 'Unknown')).upper().strip()
                site_counts[sn] = site_counts.get(sn, 0) + 1
                
                prob = str(r.get('problem_alarm') or r.get('action') or "-")
                rem = str(r.get('remark') or r.get('cond_data', {}).get('note_action') or "-")
                
                v_ac = 0.0; v_dc = 0.0; i_dc = 0.0
                kat = r.get('kategori')
                if kat == 'Site Reguler':
                    cd = r.get('cond_data', {})
                    v_ac = _sv_num(cd.get('v_rn', 0))
                    i_dc = _sv_num(cd.get('load_dc', 0))
                elif kat == 'SPS':
                    v_dc = _sv_num(r.get('rect_out_v', 0))
                    i_dc = _sv_num(r.get('total_load', 0))
                elif kat == 'Pangkalan Bun':
                    kel = r.get('pb_kelistrikan', {})
                    v_ac = _sv_num(kel.get('v_rst_n', 0))
                    rects = r.get('pb_rectifiers', [])
                    if rects:
                        v_dc = _sv_num(rects[0].get('dc_v', 0))
                        i_dc = _sv_num(rects[0].get('load', 0))
                        
                power_data.append({
                    "Site": sn, "Tanggal": str(r.get('timestamp',''))[:10], 
                    "Kategori": kat, "Status": r.get('status', '-'),
                    "Problem": prob[:70] + ("..." if len(prob)>70 else ""), 
                    "Remark": rem[:70] + ("..." if len(rem)>70 else ""), 
                    "V_AC (Volt)": v_ac, "V_DC (Volt)": v_dc, "Load_DC (Ampere)": i_dc
                })
            
            df_power = pd.DataFrame(power_data)
            
            c_a1, c_a2 = st.columns([1, 1.5])
            with c_a1:
                with st.container(border=True):
                    st.markdown("**📊 Top 10 Site Paling Sering Dikunjungi**")
                    df_sc = pd.DataFrame(list(site_counts.items()), columns=["Site", "Jumlah Visit"]).sort_values("Jumlah Visit", ascending=False).head(10)
                    ui_bar_chart(df_sc.set_index("Site"), color="#3B82F6")
            
            with c_a2:
                with st.container(border=True):
                    st.markdown("**🛑 Rekap Status Akhir Keseluruhan**")
                    df_stat = df_power['Status'].value_counts().reset_index()
                    df_stat.columns = ["Status", "Jumlah"]
                    ui_bar_chart(df_stat.set_index("Status"), color="#F59E0B", horizontal=True)
            
            st.markdown("#### ⚡ Analisa Tren Power & Histori Perbaikan Site")
            repeated = [s for s, c in site_counts.items() if c > 1]
            if not repeated:
                st.success("✅ Belum ada site dengan kunjungan lebih dari 1 kali (Tidak ada Repeated Issue).")
                pilihan = st.selectbox("📌 Pilih Site untuk melihat histori:", list(site_counts.keys()))
            else:
                st.warning(f"🚨 Terdapat {len(repeated)} Site yang terindikasi mengalami gangguan berulang (Visit > 1x).")
                pilihan = st.selectbox("📌 Pilih Site Berulang (Repeated Issue) untuk Analisa:", repeated, format_func=lambda x: f"{x} ({site_counts[x]}x Kunjungan)")
                
            if pilihan:
                df_pilih = df_power[df_power['Site'] == pilihan].sort_values("Tanggal").copy()
                
                cp1, cp2 = st.columns(2)
                with cp1:
                    with st.container(border=True):
                        st.markdown(f"**⚡ Tren Tegangan (Volt)**")
                        ui_line_chart(df_pilih.set_index("Tanggal")[['V_AC (Volt)', 'V_DC (Volt)']], color=["#EF4444", "#10B981"])
                        st.caption("Merah = AC | Hijau = DC")
                with cp2:
                    with st.container(border=True):
                        st.markdown(f"**🔋 Tren Beban (Ampere)**")
                        ui_bar_chart(df_pilih.set_index("Tanggal")[['Load_DC (Ampere)']], color="#F59E0B")
                        
                st.markdown("**📋 Histori Kendala & Tindakan:**")
                st.dataframe(df_pilih[['Tanggal', 'Kategori', 'Status', 'Problem', 'Remark']], use_container_width=True, hide_index=True)

        with tab_list:
            st.markdown("### 📋 Daftar Laporan Site Terinspeksi")
            # ---------------------------------------------------------------------
            # TAMPILAN KARTU UI MODERN & URUTAN TERBARU DI ATAS
            # ---------------------------------------------------------------------
            for i in filtered_indices:
                r = db[i]
                site_id = r.get('site_name', 'Unknown')
                kategori_site = r.get('kategori', 'SPS')
                stat = r.get('status', '')
                
                if stat == "Normal" or stat == "Done":
                    warna_status = "#00E676" 
                    icon = "🟢"
                elif stat in ["Minor Issue", "Update Progress", "On Progress"]:
                    warna_status = "#FFC107" 
                    icon = "🟡"
                else:
                    warna_status = "#FF5252" 
                    icon = "🔴"
                
                is_new_reguler = kategori_site == "Site Reguler" and ("problem_alarm" in r or "cond_data" in r)
                is_project = kategori_site == "Project"
                
                if kategori_site == "Pangkalan Bun":
                    wa_text = generate_wa_pbu(r)
                elif is_project:
                    pd_data = r.get('proj_data', {})
                    wa_text = f"*REPORT PROJECT ({pd_data.get('jenis_project', '')})* ⚡\n📍 *Project:* {site_id} ({r.get('nop', '-')})\n📅 *Tanggal:* {r.get('timestamp', '-')}\n👷 *Tim/FME:* {r.get('teknisi', '-')}\n📊 *Status Laporan:* {r.get('status', '-')}\n"
                    if pd_data.get('jenis_project') == "PO":
                        wa_text += f"\n*📋 DETAIL PO:*\n- No PO: {pd_data.get('no_po', '-')}\n- Mitra: {pd_data.get('mitra', '-')}\n- Pekerjaan Utama: {pd_data.get('tipe_pekerjaan', '-')}\n- Progress Keseluruhan: {pd_data.get('avg_progress', 0):.1f}%\n"
                        tabel_p = pd_data.get('tabel_pekerjaan', [])
                        if tabel_p:
                            wa_text += f"\n*🗺️ RINCIAN MULTI-SITE ({len(tabel_p)} Rute/Item):*\n"
                            for idx, tp in enumerate(tabel_p):
                                asal = tp.get('Site Asal', '')
                                tuj = tp.get('Site Tujuan', '')
                                pek = tp.get('Pekerjaan Spesifik', '')
                                prog = tp.get('Progress (%)', 0)
                                stat_tp = tp.get('Status', '')
                                rute = f"{asal} ➡️ {tuj}" if asal and tuj else (asal or tuj or "Site Project")
                                wa_text += f"  {idx+1}. {rute} | {pek} ({prog}% - {stat_tp})\n"
                        if pd_data.get('tipe_pekerjaan') == "Instalasi Imacro":
                            wa_text += f"\n*🛠️ Progress Imacro:* {', '.join(pd_data.get('progress_imacro', []))}\n"
                    elif pd_data.get('jenis_project') == "SVA":
                        wa_text += f"\n*📋 DETAIL SVA:*\n- No Tiket: {pd_data.get('no_tiket', '-')}\n- Temuan: {pd_data.get('temuan_sva', '-')}\n"
                    if pd_data.get('note_akhir'):
                        wa_text += f"\n*📝 NOTE:* {pd_data.get('note_akhir')}\n"
                elif is_new_reguler:
                    wa_text = f"*BERITA ACARA SITE REGULER* ⚡\n📍 *Site:* {site_id} ({r.get('nop', '-')})\n📅 *Tanggal:* {r.get('timestamp', '-')}\n👷 *FME:* {r.get('teknisi', '-')} ({r.get('role_fme', '-')})\n📊 *Status:* {r.get('status', '-')}\n\n*🚨 PROBLEM ALARM:*\n{r.get('problem_alarm', '-')}\n\n*🔧 ACTION / PERBAIKAN:*\n{', '.join(r.get('action_reguler', [])) if isinstance(r.get('action_reguler'), list) else r.get('action_reguler', '-')}\n"
                    cd = r.get('cond_data', {})
                    if cd:
                        teks_info = ""
                        if cd.get('v_rn'):
                            v_rst = cd.get('volt_ac_rst') or f"{cd.get('v_rn','-')}/{cd.get('v_sn','-')}/{cd.get('v_tn','-')}"
                            i_rst = cd.get('load_ac_rst') or f"{cd.get('i_r','-')}/{cd.get('i_s','-')}/{cd.get('i_t','-')}"
                            teks_info += f"**⚡ Power AC/DC:** Teg AC ({v_rst} V) | Arus AC ({i_rst} A) | Load DC ({cd.get('load_dc','-')} A) | BCL ({cd.get('set_bcl','-')}) | Cap ({cd.get('set_cap','-')})\n\n"
                        if cd.get('gen_v'):
                            teks_info += f"**⚙️ Genset/ATS:** Teg Out ({cd.get('gen_v','-')} V) | Arus ({cd.get('gen_i','-')} A) | ATS ({cd.get('gen_ats','-')}) | Issue ({cd.get('gen_issue','-')}) | HM ({cd.get('gen_hm','-')}) | BBM ({cd.get('gen_bbm','-')} L)\n\n"
                        if teks_info: st.info(teks_info)
                        if cd.get('note_action'):
                            st.warning(f"**📝 NOTE/REMARK Lapangan:** {cd.get('note_action')}")
                else:
                    wa_text = f"*BERITA ACARA PREVENTIVE MAINTENANCE* ⚡\n📍 *Site:* {site_id} ({kategori_site} - {r.get('nop', '-')})\n📅 *Tanggal:* {r.get('timestamp', '-')}\n👷 *Pelaksana:* {r.get('teknisi', '-')}\n📊 *Status:* {r.get('status', '-')}\n\n*RINCIAN TINDAKAN:*\n{r.get('action', '-')}\n\n*POWER & LOAD:*\n- PLN: {r.get('pln_status', '-')}\n- Rectifier: {r.get('rect_brand', '-')} ({r.get('rect_out_v', '-')}V)\n- Load BTS: {r.get('total_load', '-')} A\n\n*SPAREPART:*\n{r.get('sparepart', '-')}"
                
                wa_url = f"https://wa.me/?text={urllib.parse.quote(wa_text)}"
    
                with st.container(border=False):
                    fme_tampil = r.get('teknisi', '-') if kategori_site != "Pangkalan Bun" else r.get('pb_te', '-')
                    st.markdown(f"""
                        <div style="padding: 18px 20px; border-radius: 12px; border-left: 8px solid {warna_status}; background-color: var(--secondary-background-color); margin-bottom: 12px; box-shadow: 0 6px 16px rgba(0,0,0,0.08); border-top: 1px solid rgba(128,128,128,0.1); border-right: 1px solid rgba(128,128,128,0.1); border-bottom: 1px solid rgba(128,128,128,0.1);">
                            <div style="display: flex; justify-content: space-between; align-items: center; margin-bottom: 12px;">
                                <h3 style="margin: 0; padding: 0; font-size: 22px; font-weight: 800;">{icon} {site_id}</h3>
                                <span style="font-size: 13px; font-weight: 700; color: #ffffff; background-color: {warna_status}; padding: 4px 12px; border-radius: 20px; box-shadow: 0 2px 4px rgba(0,0,0,0.2);">📅 {r.get('timestamp', '')}</span>
                            </div>
                            <div style="font-size: 15px; margin-bottom: 8px; display: flex; align-items: center; gap: 8px; flex-wrap: wrap;">
                                <span style="background-color: rgba(128,128,128,0.1); padding: 4px 10px; border-radius: 6px;"><b>📝 Kategori:</b> {kategori_site} ({r.get('nop', '-')})</span>
                                <span style="background-color: rgba(128,128,128,0.1); padding: 4px 10px; border-radius: 6px;"><b>👷 Tim:</b> {fme_tampil}</span>
                            </div>
                            <div style="font-size: 15px; border-top: 1px dashed rgba(128,128,128,0.3); padding-top: 8px; margin-top: 8px;">
                                <b>📊 Status Akhir:</b> <span style="color: {warna_status}; font-weight: 900; font-size: 16px;">{stat}</span>
                            </div>
                        </div>
                    """, unsafe_allow_html=True)
                    
                    c_btn1, c_btn2 = st.columns(2)
                    with c_btn1:
                        st.link_button("📱 Share Rangkuman Laporan (WA)", wa_url, use_container_width=True)
                    with c_btn2:
                        if kategori_site != "Pangkalan Bun" and not is_new_reguler and not is_project:
                            if st.checkbox("📄 Buat PDF Resmi", key=f"prep_pdf_{i}"):
                                with st.spinner("⏳ Rendering Dokumen PDF Resolusi Tinggi..."):
                                    try:
                                        pdf_bytes = build_pdf(r)
                                        st.download_button(label="📥 Download PDF", data=pdf_bytes, file_name=f"Berita_Acara_{site_id}.pdf", mime="application/pdf", key=f"dl_pdf_{i}", type="primary", use_container_width=True)
                                    except Exception as e:
                                        st.error(f"Gagal memproses PDF: {e}")
                        else:
                            st.button("📄 Format Tanpa PDF", disabled=True, key=f"dl_pdf_dis_{i}", use_container_width=True)
    
                    with st.expander("📂 BUKA DETAIL & DOKUMENTASI LAPORAN"):
                        if kategori_site == "Pangkalan Bun":
                            st.markdown("### 📋 SITE VISIT REPORT (Pangkalan Bun)")
                            st.write(f"**FME NAME:** {r.get('pb_te', '-')} | **Ticket:** {r.get('pb_ticket', '-')}")
                            st.write(f"**Catuan Power:** {r.get('pb_power', '-')} | **Status Akhir:** {stat}")
                            
                            tpb1, tpb2, tpb3, tpb4 = st.tabs(["📌 Info & Recti", "⚡ Kelistrikan & ACPDB", "⛽ Genset & ENVA", "📝 Temuan & Activity"])
                            with tpb1:
                                st.markdown("**Data Rectifier:**")
                                for rect in r.get('pb_rectifiers', []): st.write(f"🔹 **{rect.get('tipe')}** | BCL: {rect.get('bcl')}% | Beban: {rect.get('load')}A | Voltase: {rect.get('dc_v')}V")
                                st.markdown("<br>", unsafe_allow_html=True)
                                tampilkan_grid_foto(r.get('pb_url_site', []), "📸 Foto Site / Kondisi Umum")
                                tampilkan_grid_foto(r.get('pb_url_recti', []), "📸 Foto Rectifier & ENVA")
                            with tpb2:
                                kel = r.get('pb_kelistrikan', {})
                                acp = r.get('pb_acpdb', {})
                                st.write(f"**Kelistrikan:** Daya KWH: {kel.get('daya', '-')} | Phasa: {kel.get('phasa', '-')} | Teg. G-N: {kel.get('v_gn', '-')}")
                                st.write(f"**ACPDB:** Tersedia: {acp.get('ada', '-')} | Teg. RST thd N: {acp.get('v_rst_n', '-')}")
                                st.markdown("<br>", unsafe_allow_html=True)
                                tampilkan_grid_foto(r.get('pb_url_listrik', []), "📸 Foto Kelistrikan & ACPDB")
                            with tpb3:
                                gen = r.get('pb_genset', {})
                                en = r.get('pb_enva', {})
                                st.write(f"**Genset:** Ada: {gen.get('ada', '-')} | Kapasitas: {gen.get('cap', '-')} | Tangki: {gen.get('tangki', '-')} | Sisa: {gen.get('bbm', '-')}")
                                st.write(f"**ENVA:** Main Fail: {en.get('main_fail', '-')} | Rectifier Fail: {en.get('rect_fail', '-')} | High Temp: {en.get('high_temp', '-')}")
                                st.markdown("<br>", unsafe_allow_html=True)
                                tampilkan_grid_foto(r.get('pb_url_genset', []), "📸 Foto Genset / Tangki")
                            with tpb4:
                                st.markdown("**Catatan/Temuan Tim:**")
                                st.info(r.get('action', '-'))
                                st.markdown("**Activity Report:**")
                                if r.get('pb_act_power'): st.write(f"- Power & CME: {', '.join(r.get('pb_act_power'))}")
                                if r.get('pb_act_trans'): st.write(f"- Transport: {', '.join(r.get('pb_act_trans'))}")
                                if r.get('pb_act_osa'): st.write(f"- OSA Impact: {', '.join(r.get('pb_act_osa'))}")
                                st.markdown("<br>", unsafe_allow_html=True)
                                tampilkan_grid_foto(r.get('pb_url_act', []), "📸 Foto Temuan / Action Lapangan")
    
                        elif is_project:
                            st.markdown("### 📋 REPORT PROJECT")
                            pd_data = r.get('proj_data', {})
                            st.markdown(f"**👨‍🔧 Tim:** {r.get('teknisi', '-')} &nbsp;&nbsp;|&nbsp;&nbsp; **🏷️ Kategori:** Project {pd_data.get('jenis_project', '-')}")
                            
                            if pd_data.get('jenis_project') == 'PO':
                                st.info(f"**📋 Detail PO:** No PO: {pd_data.get('no_po','-')} | Mitra: {pd_data.get('mitra','-')} | Pekerjaan Utama: {pd_data.get('tipe_pekerjaan','-')}")
                                val_po = float(str(pd_data.get('nilai_po', '0')).replace('.', '').replace(',', '')) if pd_data.get('nilai_po') else 0.0
                                t_cost = pd_data.get('total_cost', 0.0)
                                t_margin = pd_data.get('margin', 0.0)
                                avg_prog = pd_data.get('avg_progress', 0.0)
                                c_mar1, c_mar2, c_mar3, c_mar4 = st.columns(4)
                                c_mar1.metric("💰 Nilai PO Total", f"Rp {val_po:,.0f}")
                                c_mar2.metric("💸 Total Cost Pelaksanaan", f"Rp {t_cost:,.0f}")
                                c_mar3.metric("📈 Profit Margin", f"Rp {t_margin:,.0f}")
                                c_mar4.metric("📊 Rata-Rata Progress", f"{avg_prog:.1f}%")
                                tabel_p = pd_data.get('tabel_pekerjaan', [])
                                if tabel_p:
                                    st.markdown("#### 🗺️ Rincian Timeline, Progress & Cost Per-Site")
                                    st.dataframe(pd.DataFrame(tabel_p), use_container_width=True, hide_index=True)
                                if pd_data.get('tipe_pekerjaan') == 'Instalasi Imacro':
                                    st.warning(f"**🛠️ Progress Checklists:** {', '.join(pd_data.get('progress_imacro', []))}")
                            elif pd_data.get('jenis_project') == 'SVA':
                                st.info(f"**📋 Detail SVA:** No Tiket: {pd_data.get('no_tiket','-')} | Tanggal: {pd_data.get('tgl_sva','-')} | Temuan: {pd_data.get('temuan_sva','-')}")
                                
                            st.markdown("<br>", unsafe_allow_html=True)
                            if pd_data.get('jenis_project') == 'PO':
                                t1, t2, t3 = st.tabs(["📸 Foto Before", "📸 Foto On Progress", "📸 Foto After / BAST"])
                                with t1: tampilkan_grid_foto(pd_data.get('url_awal', []), "Kondisi Awal")
                                with t2: tampilkan_grid_foto(pd_data.get('url_proses', []), "Proses Pengerjaan")
                                with t3: tampilkan_grid_foto(pd_data.get('url_akhir', []), "Hasil Akhir (BAST)")
                            else:
                                t1, t2 = st.tabs(["📸 View Site & Perangkat", "📸 Temuan / Defect"])
                                with t1: tampilkan_grid_foto(pd_data.get('url_sva1', []), "Foto View Site & Perangkat")
                                with t2: tampilkan_grid_foto(pd_data.get('url_sva2', []), "Foto Temuan Defect")
                                
                        elif is_new_reguler:
                            st.markdown("### 📋 SITE REGULER REPORT")
                            st.markdown(f"**👨‍🔧 FME:** {r.get('teknisi', '-')} ({r.get('role_fme', '-')}) &nbsp;&nbsp;|&nbsp;&nbsp; **🚨 Problem:** {r.get('problem_alarm', '-')}")
                            st.markdown(f"**🔧 Action Dikerjakan:** {', '.join(r.get('action_reguler', [])) if isinstance(r.get('action_reguler'), list) else r.get('action_reguler', '-')}")
                            
                            cd = r.get('cond_data', {})
                            if cd:
                                teks_info = ""
                                if cd.get('v_rn'):
                                    v_rst = cd.get('volt_ac_rst') or f"{cd.get('v_rn','-')}/{cd.get('v_sn','-')}/{cd.get('v_tn','-')}"
                                    i_rst = cd.get('load_ac_rst') or f"{cd.get('i_r','-')}/{cd.get('i_s','-')}/{cd.get('i_t','-')}"
                                    teks_info += f"**⚡ Power AC/DC:** Teg AC ({v_rst} V) | Arus AC ({i_rst} A) | Load DC ({cd.get('load_dc','-')} A) | BCL ({cd.get('set_bcl','-')}) | Cap ({cd.get('set_cap','-')})\n\n"
                                if cd.get('gen_v'):
                                    teks_info += f"**⚙️ Genset/ATS:** Teg Out ({cd.get('gen_v','-')} V) | Arus ({cd.get('gen_i','-')} A) | ATS ({cd.get('gen_ats','-')}) | Issue ({cd.get('gen_issue','-')}) | HM ({cd.get('gen_hm','-')}) | BBM ({cd.get('gen_bbm','-')} L)\n\n"
                                if teks_info: st.info(teks_info)
                                if cd.get('note_action'):
                                    st.warning(f"**📝 NOTE/REMARK Lapangan:** {cd.get('note_action')}")
                                
                            st.markdown("<br>", unsafe_allow_html=True)
                            t1, t2, t3 = st.tabs(["📸 Problem Alarm", "📸 Action Perbaikan", "📸 Eviden Parameter"])
                            with t1: tampilkan_grid_foto(r.get('url_problem', []), "Dokumentasi Problem Alarm")
                            with t2: tampilkan_grid_foto(r.get('url_action', []), "Dokumentasi Perbaikan / Action")
                            with t3: 
                                if cd:
                                    eviden_list = [
                                        (cd.get('url_v_rn'), f"Teg AC R-N: {cd.get('v_rn', '-')} V"),
                                        (cd.get('url_v_sn'), f"Teg AC S-N: {cd.get('v_sn', '-')} V"),
                                        (cd.get('url_v_tn'), f"Teg AC T-N: {cd.get('v_tn', '-')} V"),
                                        (cd.get('url_i_r'), f"Load AC R: {cd.get('i_r', '-')} A"),
                                        (cd.get('url_i_s'), f"Load AC S: {cd.get('i_s', '-')} A"),
                                        (cd.get('url_i_t'), f"Load AC T: {cd.get('i_t', '-')} A"),
                                        (cd.get('url_ldc'), f"Load DC Total: {cd.get('load_dc', '-')} A"),
                                        (cd.get('url_bcl'), f"Set BCL: {cd.get('set_bcl', '-')}"),
                                        (cd.get('url_cap'), f"Set Capacity: {cd.get('set_cap', '-')}"),
                                        (cd.get('url_mod'), f"Module Kontroler: {cd.get('mod_rusak', '-')}"),
                                        (cd.get('url_gen_v'), f"Tegangan Out Genset: {cd.get('gen_v', '-')} V"),
                                        (cd.get('url_gen_i'), f"Load Arus Genset: {cd.get('gen_i', '-')} A"),
                                        (cd.get('url_gen_ats'), f"Panel ATS: {cd.get('gen_ats', '-')}"),
                                        (cd.get('url_gen_issue'), f"Kondisi Mesin: {cd.get('gen_issue', '-')}"),
                                        (cd.get('url_gen_hm'), f"Hour Meter: {cd.get('gen_hm', '-')} HM"),
                                        (cd.get('url_gen_bbm'), f"Sisa BBM: {cd.get('gen_bbm', '-')} Liter")
                                    ]
                                    ada_foto = False
                                    for u_ev, cap_ev in eviden_list:
                                        if u_ev:
                                            ada_foto = True
                                            tampilkan_grid_foto([u_ev], f"📌 {cap_ev}")
                                    if not ada_foto: st.write("Belum ada foto parameter pengukuran pada laporan ini.")
                                else:
                                    st.write("Tidak ada parameter pengukuran yang dilaporkan.")
                        else:
                            st.warning("⏱️ Data di bawah ini direkam menggunakan Form Versi Lama (sebelum ada pemisahan layout).")
                            st.markdown(f"**👨‍🔧 Pelaksana (Teknisi):** {r.get('teknisi', '-')} &nbsp;&nbsp;|&nbsp;&nbsp; **⚡ Beban Load:** {r.get('total_load', '-')} A")
                            st.markdown(f"**🔧 Action Dikerjakan:** {r.get('action', '-')}")
                            if r.get('sparepart'): st.warning(f"**📦 Sparepart Diganti:** {r['sparepart']}")
                            if r.get('remark'): st.info(f"**📝 Remark:** {r['remark']}")
                            st.markdown("<br>", unsafe_allow_html=True)
                            
                            ltab1, ltab2, ltab3, ltab4, ltab5 = st.tabs(["🏗️ 1. Fisik", "☀️ 2. Panel SPS", "🔌 3. Recti & Genset", "🔋 4. Baterai & Gnd", "📂 5. Datalog"])
                            with ltab1:
                                st.write(f"- Kondisi Site: {r.get('site_cond', '-')} | Tower: {r.get('tower_cond', '-')} | Shading: {r.get('shading_status', '-')}")
                                tampilkan_grid_foto(r.get('url_sites'), "📸 View Site")
                                tampilkan_grid_foto(r.get('url_shadings'), "📸 Shading")
                                tampilkan_grid_foto(r.get('extras_fisik'), "📸 Tambahan")
                            with ltab2:
                                st.write(f"**Status Panel:** {r.get('status_panel', '-')} ({r.get('jml_panel_rusak', 0)} Rusak) | **Merk Panel:** {r.get('merk_panel', '-')}")
                                st.write(f"**Status SCC:** {r.get('status_scc', '-')} ({r.get('jml_scc_nok', 0)} Rusak) | **Merk SCC:** {r.get('merk_scc', '-')}")
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
                                    st.write(f"**Possibility PSB:** {r.get('possibility_psb', '-')} | **Jarak PLN:** {r.get('jarak_pln', '-')}")
                                    tampilkan_grid_foto(r.get('url_rects'), "📸 Rectifier")
                                with c_p2:
                                    st.write(f"**Genset:** {r.get('genset_status', '-')} | **BBM:** {r.get('fuel_pct', '-')}%")
                                    tampilkan_grid_foto(r.get('url_gensets'), "📸 Genset")
                                st.divider()
                                tampilkan_grid_foto(r.get('extras_elektrikal'), "📸 Tambahan Mesin/Elektrikal")
                            with ltab4:
                                st.markdown("### 🔋 Bank Baterai")
                                st.write(f"**Status Baterai Keseluruhan:** {r.get('status_battery', '-')} ({r.get('jml_bat_rusak', 0)} Rusak) | **Merk Baterai:** {r.get('tipe_battery', '-')}")
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
    
                        if st.session_state['role'] == 'Admin':
                            st.markdown("<br>", unsafe_allow_html=True)
                            with st.container(border=True):
                                if is_project and r.get('proj_data', {}).get('jenis_project') == 'PO':
                                    st.markdown("<h4 style='color: var(--primary-color);'>🛠 UPDATE PROGRESS PROJECT & FOTO SUSULAN</h4>", unsafe_allow_html=True)
                                    st.info("💡 Update progress tabel bisa dilakukan di menu Executive Project Dashboard.")
                                else:
                                    st.markdown("<h4 style='color: var(--primary-color);'>🛠 REVISI NAMA SITE / FME</h4>", unsafe_allow_html=True)
                                    ce1, ce2 = st.columns(2)
                                    with ce1: new_sreg = st.text_input("Edit Nama Site", r.get('site_name',''), key=f"esreg_{i}")
                                    with ce2: new_freg = st.text_input("Edit Nama FME / Tim", r.get('teknisi',''), key=f"efreg_{i}")
                                    
                                    if st.button("💾 Simpan Perubahan", key=f"breg_{i}"):
                                        with st.spinner("Menyimpan..."):
                                            r['site_name'], r['teknisi'] = new_sreg, new_freg
                                            sheet = connect_gsheets(r.get('_sheet', 'Report Preventive'))
                                            if sheet:
                                                sheet.update_cell(r.get('_row', i + 2), 2, new_sreg)
                                                sheet.update_cell(r.get('_row', i + 2), 4, new_freg)
                                                sheet.update_cell(r.get('_row', i + 2), 6, json.dumps(r))
                                                st.cache_data.clear() 
                                        st.success("✅ REVISI BERHASIL!")
                                        st.rerun()
# =========================================================================
# MENU 3: EXCEL LIVE-EDITOR (MONITORING IMPROVEMENT)
# =========================================================================
elif menu == "📈 Monitoring Improvement":
    st.markdown("<h1 style='text-align: center;'>📈 Master Tracker Improvement</h1>", unsafe_allow_html=True)
    
    file_master = "Monitoring_Availability_Improvement_Visit_SPS_NOP_PLK.xlsx"
    
    if not os.path.exists(file_master):
        st.error(f"❌ File '{file_master}' tidak ditemukan di sistem/server. Pastikan Anda telah meletakkan file tersebut satu folder dengan aplikasi.")
    else:
        tab_dashboard, tab_tim, tab_editor = st.tabs(["📊 Kurva S", "👥 Progress Tim (Done/Berjalan)", "📝 Live Editor Data"])
        
        with tab_dashboard:
            st.markdown("### 📈 S-Curve Target vs Aktual")
            st.info("💡 Data grafik ini dibaca secara Real-Time dari **Sheet 'Kurva S'** di file Excel Master Anda.")
            
            try:
                df_kurva = read_excel_cached(file_master, sheet_name='Kurva S', header=None)
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
                    ui_line_chart(df_chart[['Plan_Kumulatif', 'Actual_Kumulatif']], color=["#FF5252", "#64FFDA"])
                    
                with col_table:
                    st.markdown("**Tabel Gap Mingguan**")
                    st.dataframe(df_chart[['Plan_Kumulatif', 'Actual_Kumulatif', 'Gap']], use_container_width=True)
            except Exception as e:
                st.error(f"Gagal memvisualisasikan Grafik S-Curve: {e}")

        with tab_tim:
            st.markdown("### 👥 Dashboard Eksekusi Lapangan")
            st.info("💡 Sistem akan membaca Sheet **Tracker Improvement** untuk menghitung berapa site yang sudah Selesai (Done) dan masih dalam pengerjaan (On Progress).")
            
            try:
                df_track = read_excel_cached(file_master, sheet_name='Tracker Improvement', header=None)
                headers_track = df_track.iloc[3].fillna("").astype(str).tolist()
                df_t = df_track.iloc[4:].copy()
                df_t.columns = headers_track
                
                status_cols = [c for c in headers_track if 'status' in c.lower() or 'progress' in c.lower() or 'keterangan' in c.lower() or 'aktual' in c.lower()]
                default_idx = headers_track.index(status_cols[0]) if status_cols else 0
                
                st.markdown("**Pilih kolom pada Excel yang berisi status pengerjaan tim:**")
                sel_col = st.selectbox("Pilih Kolom Status:", headers_track, index=default_idx)
                
                if sel_col:
                    df_t[sel_col] = df_t[sel_col].astype(str).str.strip().str.upper()
                    df_valid = df_t[~df_t[sel_col].isin(['NAN', 'NAT', 'NONE', ''])]
                    
                    counts = df_valid[sel_col].value_counts().reset_index()
                    counts.columns = ['Status', 'Jumlah']
                    
                    if not counts.empty:
                        done_kws = ['DONE', 'SELESAI', 'CLOSE', 'OK', 'COMPLETED']
                        prog_kws = ['PROGRESS', 'BERJALAN', 'ON GOING', 'OPEN', 'PENDING', 'ON PROGRESS']
                        
                        jml_done = counts[counts['Status'].apply(lambda x: any(k in x for k in done_kws))]['Jumlah'].sum()
                        jml_prog = counts[counts['Status'].apply(lambda x: any(k in x for k in prog_kws))]['Jumlah'].sum()
                        jml_lain = counts['Jumlah'].sum() - (jml_done + jml_prog)
                        
                        st.markdown("#### 🎯 Hasil Kalkulasi Progress")
                        c_t1, c_t2, c_t3 = st.columns(3)
                        c_t1.metric("✅ TOTAL SELESAI (Done)", f"{jml_done} Site")
                        c_t2.metric("⏳ SEDANG BERJALAN (On Progress)", f"{jml_prog} Site")
                        c_t3.metric("📝 Lainnya / Belum Mulai", f"{jml_lain} Site")
                        _tot_t = max(jml_done + jml_prog + jml_lain, 1)
                        c_rg1, c_rg2 = st.columns([1, 2])
                        with c_rg1:
                            ui_ring(jml_done / _tot_t * 100, "Selesai", "#10B981")
                        with c_rg2:
                            ux_progress([("Selesai", jml_done, "#10B981"), ("On Progress", jml_prog, "#F59E0B"), ("Lainnya / Belum Mulai", jml_lain, "#94A3B8")])
                        
                        st.markdown("#### 📊 Grafik Breakdown Status")
                        ui_bar_chart(counts.set_index('Status'), color="#64FFDA")
                        
                        with st.expander("Tampilkan Rekap Tabel Mentah Status", expanded=False):
                            st.dataframe(counts, use_container_width=True)
                    else:
                        st.warning(f"Data pada kolom '{sel_col}' kosong atau seluruhnya terbaca sebagai cell kosong.")
            except Exception as e:
                st.error(f"Gagal memproses data tracker tim: {e}")

        with tab_editor:
            st.markdown("### 📝 Tabel Master Tracker & Viewer")
            st.write("Klik ganda (*double click*) pada sel tabel untuk merubah isi data secara instan.")
            
            sheet_options = ["Tracker Improvement", "Jadwal Visit SPS", "Dashboard", "Kurva S"]
            sheet_choice = st.selectbox("Pilih Sheet Excel yang Ingin Diedit / Dilihat:", sheet_options)
            
            try:
                df_raw = read_excel_cached(file_master, sheet_name=sheet_choice, header=None)
                
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

# =========================================================================
# MENU 4: DAPOT POWER HASIL PMS (MENU TERPISAH KHUSUS EXECUTIVE DASHBOARD)
# =========================================================================
elif menu == "⚡ Dapot Power Hasil PMS":
    st.markdown("<h1 style='text-align: center;'>⚡ Executive Dapot Power</h1>", unsafe_allow_html=True)
    st.markdown("<p style='text-align: center; opacity: 0.8;'>Power Infrastructure Performance Dashboard — Analisa & Komparasi Master vs Aktual</p>", unsafe_allow_html=True)

    dapot_file = "Dapot_Power_Executive_Dashboard.xlsx"
    db = st.session_state.get('laporan_db', [])

    st.markdown("<br>", unsafe_allow_html=True)
    area_dapot = st.radio("📍 Pilih Area Dapot Power:", ["🔵 Palangkaraya", "🔴 Pangkalan Bun"], horizontal=True)
    st.markdown("<hr style='margin: 5px 0 20px 0;'>", unsafe_allow_html=True)

    # =========================================================
    # DAPOT KHUSUS PANGKALAN BUN
    # =========================================================
    if area_dapot == "🔴 Pangkalan Bun":
        st.markdown("### 🔴 Komparasi Dapot Power vs Aktual Lapangan (Pangkalan Bun)")
        st.info("💡 Sistem membandingkan data master dari sheet **'dapot pbun'** dengan hasil laporan (Onsite) terbaru dari tim lapangan untuk melihat gap / selisih aktual.")
        
        if not os.path.exists(dapot_file):
            st.error(f"❌ File **`{dapot_file}`** belum ditemukan.")
        else:
            try:
                df_dapot_pbu = read_excel_cached(dapot_file, sheet_name='dapot pbun')
                
                # Cari kolom yang mengandung kata 'site' sebagai kunci ID
                site_col = next((c for c in df_dapot_pbu.columns if 'site' in str(c).lower() or 'id' in str(c).lower()), df_dapot_pbu.columns[0])
                df_dapot_pbu[site_col] = df_dapot_pbu[site_col].astype(str).str.strip().str.upper()
                
                # Filter DB hanya untuk laporan Pangkalan Bun terbaru
                db_pbu = [r for r in db if r.get('kategori') == 'Pangkalan Bun']
                latest_onsite = {}
                for r in db_pbu:
                    sid = str(r.get('pb_site_id', '')).strip().upper()
                    if sid:
                        latest_onsite[sid] = r # Dictionary otomatis mereplace key sama dgn data paling akhir
                        
                komparasi_data = []
                for _, row_dapot in df_dapot_pbu.iterrows():
                    sid = str(row_dapot[site_col]).strip().upper()
                    onsite_r = latest_onsite.get(sid)
                    
                    row_dict = {"Site ID (Master)": sid}
                    
                    # Ambil beberapa kolom dari Master Dapot sebagai info target (Max 4 kolom agar tidak terlalu panjang)
                    for col in [c for c in df_dapot_pbu.columns if c != site_col][:4]:
                        row_dict[f"Master: {col}"] = row_dapot[col]
                    
                    if onsite_r:
                        rects = onsite_r.get('pb_rectifiers', [])
                        aktual_rect = rects[0].get('tipe', '-') if rects else '-'
                        aktual_load = rects[0].get('load', '-') if rects else '-'
                        aktual_modul = rects[0].get('mod_ok', '-') if rects else '-'
                        aktual_bat = rects[0].get('jml_batt', '-') if rects else '-'
                        
                        gen = onsite_r.get('pb_genset', {})
                        aktual_genset = gen.get('cap', '-')
                        
                        tgl_aktual = str(onsite_r.get('timestamp', ''))[:10]
                        fme = onsite_r.get('pb_te', '-')
                        status_visit = "✅ Sudah Visit"
                        act_power = onsite_r.get('pb_act_power', [])
                        act = ", ".join(act_power) if isinstance(act_power, list) else str(act_power)
                    else:
                        aktual_rect = "-"
                        aktual_load = "-"
                        aktual_modul = "-"
                        aktual_bat = "-"
                        aktual_genset = "-"
                        tgl_aktual = "-"
                        fme = "-"
                        act = "-"
                        status_visit = "❌ Belum Visit"
                        
                    row_dict.update({
                        "Status Aktual": status_visit,
                        "Tgl Visit Lapangan": tgl_aktual,
                        "Tim FME": fme,
                        "Aktual Rectifier": aktual_rect,
                        "Aktual Load DC (A)": aktual_load,
                        "Aktual Modul OK": aktual_modul,
                        "Aktual Baterai": aktual_bat,
                        "Aktual Genset": aktual_genset,
                        "Activity": act
                    })
                    komparasi_data.append(row_dict)
                    
                df_komp = pd.DataFrame(komparasi_data)
                
                tot_dapot = len(df_komp)
                tot_visit = len(df_komp[df_komp['Status Aktual'] == "✅ Sudah Visit"])
                
                c_m1, c_m2, c_m3 = st.columns(3)
                c_m1.metric("Total Site (Master Dapot PBU)", f"{tot_dapot} Site")
                c_m2.metric("Tersinkron Data Onsite (Sudah Visit)", f"{tot_visit} Site")
                c_m3.metric("Gap Belum Di-Visit", f"{tot_dapot - tot_visit} Site")
                
                tab_k1, tab_k2 = st.tabs(["📊 Tabel Komparasi Master vs Onsite", "🗂️ Live Editor Master Dapot PBUN"])
                
                with tab_k1:
                    st.dataframe(df_komp, use_container_width=True, hide_index=True)
                    st.download_button(
                        "📥 Download Data Komparasi (.csv)",
                        data=df_komp.to_csv(index=False).encode('utf-8'),
                        file_name=f"Komparasi_Dapot_PBU_{datetime.date.today()}.csv",
                        mime="text/csv",
                        type="primary",
                        use_container_width=True
                    )
                    
                with tab_k2:
                    st.write("Klik dua kali (*double-click*) pada sel tabel di bawah untuk merevisi data Master Dapot PBUN secara langsung. Jangan lupa klik tombol Simpan di bawah.")
                    edited_pbun = st.data_editor(df_dapot_pbu, use_container_width=True, num_rows="dynamic")
                    if st.button("💾 Simpan Perubahan Dapot PBUN ke Excel", type="primary"):
                        with st.spinner("Menyimpan ke Excel Master..."):
                            import openpyxl
                            wb = openpyxl.load_workbook(dapot_file)
                            ws = wb['dapot pbun']
                            ws.delete_rows(1, ws.max_row)
                            
                            for c_idx, c_name in enumerate(edited_pbun.columns, 1):
                                ws.cell(row=1, column=c_idx).value = c_name
                                
                            for r_idx, row in edited_pbun.iterrows():
                                for c_idx, c_name in enumerate(edited_pbun.columns, 1):
                                    ws.cell(row=r_idx+2, column=c_idx).value = row[c_name]
                                    
                            buf = BytesIO()
                            wb.save(buf)
                            st.success("✅ Dapot Pangkalan Bun berhasil diperbarui di Master Excel!")
                            st.download_button("📥 Download Dapot Terupdate", data=buf.getvalue(), file_name=f"Updated_Dapot_PBU_{datetime.date.today()}.xlsx", mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet")
                            
            except Exception as e:
                st.error(f"⚠️ Terjadi kesalahan saat membaca sheet 'dapot pbun': {e}. Pastikan nama sheet di dalam file Excel 'Dapot_Power_Executive_Dashboard.xlsx' benar-benar 'dapot pbun'.")
        
        # Hentikan eksekusi di sini agar tidak menjalankan logika Palangkaraya di bawahnya
        st.stop()

    # =========================================================
    # DAPOT PALANGKARAYA (KODE EXISTING)
    # =========================================================
    tab_comp, tab_rekap, tab_reguler, tab_master = st.tabs([
        "📈 Compare Bulanan PMS (Master Excel)", 
        "📋 Daftar Site Rutin vs 1x Visit", 
        "🚨 Analisa Anomali Site Reguler (Live DB)",
        "🗂️ Master Data Konsolidasi & Live Editor"
    ])

    # =========================================================
    # TAB BARU: ANALISA SITE REGULER (TROUBLESHOOTING BERULANG)
    # =========================================================
    with tab_reguler:
        st.markdown("### 🚨 Analisa Troubleshooting Berulang & Anomali Power (Site Reguler)")
        st.info("💡 Sistem melacak histori perbaikan Site Reguler secara otomatis. Site dengan kunjungan > 1 kali (*Repeated Issue*) akan dianalisa untuk menemukan **'Ultimate Solution'**.")
        
        # Ambil data Site Reguler dari laporan_db
        db_reg = [r for r in db if r.get('kategori') == 'Site Reguler']
        
        if not db_reg:
            st.warning("Belum ada data kunjungan Site Reguler di database.")
        else:
            reg_data = []
            for r in db_reg:
                cd = r.get('cond_data', {})
                reg_data.append({
                    "Timestamp": r.get('timestamp', ''),
                    "Site Name": str(r.get('site_name', '')).upper().strip(),
                    "NOP": r.get('nop', '-'),
                    "Problem": r.get('problem_alarm', '-'),
                    "Action": ", ".join(r.get('action_reguler', [])) if isinstance(r.get('action_reguler'), list) else r.get('action_reguler', '-'),
                    "Teknisi": r.get('teknisi', '-'),
                    "V_RN": pd.to_numeric(cd.get('v_rn'), errors='coerce'),
                    "V_SN": pd.to_numeric(cd.get('v_sn'), errors='coerce'),
                    "V_TN": pd.to_numeric(cd.get('v_tn'), errors='coerce'),
                    "Load_DC": pd.to_numeric(cd.get('load_dc'), errors='coerce'),
                    "Remark": cd.get('note_action', '-')
                })
            df_reg = pd.DataFrame(reg_data)
            df_reg['Timestamp_DT'] = pd.to_datetime(df_reg['Timestamp'], errors='coerce')
            df_reg = df_reg.sort_values('Timestamp_DT')
            
            # Hitung frekuensi kunjungan
            visit_counts_reg = df_reg['Site Name'].value_counts()
            repeated_sites = visit_counts_reg[visit_counts_reg > 1].index.tolist()
            
            cr1, cr2, cr3 = st.columns(3)
            cr1.metric("Total Kunjungan Reguler (Troubleshoot)", f"{len(df_reg)} Visit")
            cr2.metric("Total Unik Site Terlibat", f"{len(visit_counts_reg)} Site")
            cr3.metric("🚨 Site Repeated Issue (>1x Visit)", f"{len(repeated_sites)} Site")
            
            st.markdown("---")
            if not repeated_sites:
                st.success("✅ Luar Biasa! Belum ada site yang mengalami kerusakan atau kunjungan berulang.")
                st.dataframe(df_reg[['Timestamp', 'Site Name', 'Problem', 'Teknisi']], use_container_width=True, hide_index=True)
            else:
                c_sel1, c_sel2 = st.columns([1, 2])
                with c_sel1:
                    site_pilihan = st.selectbox("📌 Pilih Site Berulang untuk Dianalisa:", repeated_sites, format_func=lambda x: f"{x} ({visit_counts_reg[x]}x Kunjungan)")
                
                with c_sel2:
                    st.markdown(f"**Histori Problem Alarm pada {site_pilihan}:**")
                    df_site = df_reg[df_reg['Site Name'] == site_pilihan].copy()
                    st.write(" ➔ ".join(df_site['Problem'].tolist()))
                
                # Visualisasi Trend Power (Menggunakan UX UI Premium)
                st.markdown(f"#### 📈 Tren Parameter Power AC & DC: {site_pilihan}")
                df_chart = df_site.copy()
                df_chart['Waktu'] = df_chart['Timestamp'].str[:10]  # Ambil YYYY-MM-DD
                df_chart = df_chart.set_index('Waktu')
                
                c_ch1, c_ch2 = st.columns(2)
                with c_ch1:
                    with st.container(border=True):
                        st.markdown("**⚡ Fluktuasi Tegangan AC (Fasa R, S, T)**")
                        ui_line_chart(df_chart[['V_RN', 'V_SN', 'V_TN']], color=["#EF4444", "#F59E0B", "#3B82F6"])
                with c_ch2:
                    with st.container(border=True):
                        st.markdown("**🔋 Tren Beban DC (Load Current)**")
                        ui_bar_chart(df_chart[['Load_DC']], color="#10B981")
                
                # Auto Generate Ultimate Solution (Rule-based AI Expert)
                avg_load = df_site['Load_DC'].mean()
                min_ac = df_site[['V_RN', 'V_SN', 'V_TN']].min().min()
                max_ac = df_site[['V_RN', 'V_SN', 'V_TN']].max().max()
                
                solusi = ""
                if min_ac < 190:
                    solusi += "🚨 **ANOMALI AC (UNDERVOLTAGE):** Tegangan AC pernah drop parah di bawah 190V. Hal ini menyebabkan modul rectifier gagal mensuplai DC dengan maksimal, sehingga baterai sering terkuras.\n"
                    solusi += "💡 **ULTIMATE SOLUTION:** Segera koordinasi dengan PLN untuk penormalan tap trafo, atau ajukan pemasangan AVR/Stabilizer 3 Fasa kapasitas besar di site ini.\n\n"
                elif max_ac > 245:
                    solusi += "🚨 **ANOMALI AC (OVERVOLTAGE):** Tegangan AC melonjak abnormal melebihi 245V. Kondisi ini memicu proteksi *Over Voltage Protection* (OVP) modul rectifier sehingga modul mematikan diri (*shutdown*).\n"
                    solusi += "💡 **ULTIMATE SOLUTION:** Pemasangan surge arrester yang valid dan koordinasi PLN untuk memecah / menyeimbangkan beban fasa jaringan luar.\n\n"
                
                if avg_load > 60:
                    solusi += f"🚨 **BEBAN DC OVERLOAD:** Rata-rata beban harian terbaca sangat tinggi ({avg_load:.1f}A). Jika kapasitas baterai (Ah) dan jumlah modul rectifier pas-pasan, site akan sangat mudah *down* walau PLN hanya kedip sebentar.\n"
                    solusi += "💡 **ULTIMATE SOLUTION:** Segera buatkan pengajuan RAB (via SiSAVA) untuk **Upgrade Kapasitas Baterai (Add Battery Bank)** dan penambahan Modul Rectifier agar rasio beban vs *backup time* kembali ideal.\n\n"
                
                problem_gabungan = " ".join(df_site['Problem'].tolist()).lower()
                if "cell down" in problem_gabungan or "down power" in problem_gabungan:
                    solusi += "🚨 **REPEATED DOWN POWER / CELL DOWN:** Kasus mati total / sektor mati berulang kali.\n"
                    solusi += "💡 **ULTIMATE SOLUTION:** Audit total resistansi sistem *grounding*, cek suhu ruang baterai (AC mati / kipas shelter rusak mempercepat baterai soak), dan periksa degradasi sel baterai secara individual (dummy load test).\n\n"
                    
                if not solusi:
                    solusi += "✅ **PARAMETER POWER TERLIHAT STABIL:** Tegangan PLN dan beban DC terbaca dalam ambang batas wajar. Masalah berulang kemungkinan murni disebabkan oleh faktor transmisi, optik, atau perangkat CME (*hardware bug*).\n"
                    solusi += "💡 **ULTIMATE SOLUTION:** Lakukan *preventive maintenance* fisik pada rute kabel FO, bersihkan modul-modul radio (RRU/BBU) dari debu yang menyumbat pembuangan panas, dan perketat seal *weatherproofing*.\n\n"
                    
                st.markdown("#### 🤖 Analisa Sistem & Ultimate Solution")
                st.info(solusi)
                
                st.markdown("#### 📋 Rincian Histori Pekerjaan (Tabel Log)")
                st.dataframe(df_site[['Timestamp', 'Teknisi', 'Problem', 'Action', 'V_RN', 'V_SN', 'V_TN', 'Load_DC', 'Remark']], use_container_width=True, hide_index=True)

    # =========================================================
    # TAB EXCEL MASTER LAMA
    # =========================================================
    if not os.path.exists(dapot_file):
        with tab_comp: st.error(f"❌ File **`{dapot_file}`** belum ditemukan. Pastikan file Excel master tersebut sudah di-upload.")
    else:
        try:
            df_kpi = read_excel_cached(dapot_file, sheet_name='DB_KPI')
            df_trend = read_excel_cached(dapot_file, sheet_name='DB_TREND')

            df_kpi['SITE_ID'] = df_kpi['SITE_ID'].astype(str).str.strip().str.upper()
            df_trend['SITE_ID'] = df_trend['SITE_ID'].astype(str).str.strip().str.upper()
            df_trend['BULAN'] = df_trend['BULAN'].astype(str).str.strip()

            visit_counts = df_trend['SITE_ID'].value_counts()
            sites_rutin = visit_counts[visit_counts > 1].index.tolist()
            sites_sekali = visit_counts[visit_counts == 1].index.tolist()
            all_visited_sites = visit_counts.index.tolist()

            with tab_comp:
                st.markdown("### 🔍 Filter & Pilih Site untuk Compare Bulanan (Master Excel)")
                
                c_f1, c_f2 = st.columns([1.5, 2])
                with c_f1:
                    kat_kunjungan = st.radio("Filter Kategori Kunjungan Site:", ["🔄 Dikunjungi Rutin (>1 Bulan / Bisa Compare)", "📍 Pernah Dikunjungi (1x Kunjungan)", "🌐 Semua Site Terdata"])
                
                if "Rutin" in kat_kunjungan: pilihan_sites = sites_rutin
                elif "1x" in kat_kunjungan: pilihan_sites = sites_sekali
                else: pilihan_sites = all_visited_sites

                with c_f2:
                    format_opsi = {s: f"{s} ({visit_counts.get(s, 0)}x Kunjungan)" for s in pilihan_sites}
                    selected_site = st.selectbox("Pilih Site ID untuk Melihat Grafik & KPI:", options=pilihan_sites, format_func=lambda x: format_opsi.get(x, x))

                if selected_site:
                    kpi_row = df_kpi[df_kpi['SITE_ID'] == selected_site]
                    st.markdown(f"#### 📌 Status KPI Terakhir — Site **{selected_site}**")
                    
                    k1, k2, k3, k4, k5, k6 = st.columns(6)
                    if not kpi_row.empty:
                        kr = kpi_row.iloc[0]
                        k1.metric("Last Inspection", str(kr.get('TGL_UPDATE', '-')))
                        k2.metric("Status PLN", str(kr.get('STATUS_PLN', '-')).upper())
                        k3.metric("Total Baterai", f"{kr.get('TOTAL_BATERAI', 0)} Bank")
                        k4.metric("Total Modul", f"{kr.get('TOTAL_MODUL', 0)} Pcs")
                        k5.metric("Latest Voltage AC", f"{kr.get('TEGANGAN_AC', 0)} V")
                        k6.metric("Latest Load DC", f"{kr.get('LOAD_DC', 0)} A")
                    else:
                        st.info("Data ringkasan di sheet DB_KPI belum tersedia untuk Site ini.")

                    st.markdown("---")

                    df_site_trend = df_trend[df_trend['SITE_ID'] == selected_site].sort_values(by='BULAN').copy()
                    
                    if not df_site_trend.empty:
                        for col_num in ['VAC', 'VDC', 'IAC', 'LDC']:
                            df_site_trend[col_num] = pd.to_numeric(df_site_trend[col_num], errors='coerce').fillna(0.0)

                        df_site_trend['Delta_VAC'] = df_site_trend['VAC'].diff().fillna(0.0).round(2)
                        df_site_trend['Delta_VDC'] = df_site_trend['VDC'].diff().fillna(0.0).round(2)
                        df_site_trend['Delta_IAC'] = df_site_trend['IAC'].diff().fillna(0.0).round(2)
                        df_site_trend['Delta_LDC'] = df_site_trend['LDC'].diff().fillna(0.0).round(2)

                        df_plot = df_site_trend.set_index('BULAN')

                        st.markdown(f"#### 📊 Grafik Perbandingan Bulanan AC & DC — **{selected_site}** ({len(df_site_trend)} Bulan Kunjungan)")
                        
                        col_g1, col_g2 = st.columns(2)
                        with col_g1:
                            with st.container(border=True):
                                st.markdown("**⚡ Perbandingan Tegangan: AC (VAC) vs DC (VDC) per Bulan**")
                                ui_line_chart(df_plot[['VAC', 'VDC']], color=["#FF5252", "#00E676"])
                                st.caption("Merah = Tegangan AC (Volt) | Hijau = Tegangan DC Rectifier (Volt)")
                                
                        with col_g2:
                            with st.container(border=True):
                                st.markdown("**🔌 Perbandingan Arus/Beban: Arus AC (IAC) vs Load DC (LDC) per Bulan**")
                                ui_line_chart(df_plot[['IAC', 'LDC']], color=["#FFC107", "#2979FF"])
                                st.caption("Kuning = Arus AC (Ampere) | Biru = Beban Load DC (Ampere)")

                        st.markdown("#### 📋 Tabel Histori & Selisih (Compare) Antar Bulan")
                        st.dataframe(df_site_trend[['BULAN', 'VAC', 'Delta_VAC', 'VDC', 'Delta_VDC', 'IAC', 'Delta_IAC', 'LDC', 'Delta_LDC']], use_container_width=True, hide_index=True)
                    else:
                        st.warning("Data trend bulanan tidak ditemukan untuk site ini.")

            with tab_rekap:
                st.markdown("### 📋 Rekapitulasi Frekuensi Kunjungan Site PMS")
                
                rekap_df = df_trend.groupby('SITE_ID').agg(
                    Jumlah_Kunjungan=('BULAN', 'count'),
                    Daftar_Bulan_Kunjungan=('BULAN', lambda x: ", ".join(sorted(x.astype(str).unique()))),
                    Rata_VAC=('VAC', 'mean'), Rata_VDC=('VDC', 'mean'), Rata_IAC=('IAC', 'mean'), Rata_LDC=('LDC', 'mean')
                ).reset_index()

                rekap_df['Kategori_Kunjungan'] = rekap_df['Jumlah_Kunjungan'].apply(lambda x: "Dikunjungi Rutin (>1x)" if x > 1 else "Pernah Dikunjungi (1x)")
                rekap_df = rekap_df.sort_values(by='Jumlah_Kunjungan', ascending=False).reset_index(drop=True)
                rekap_df[['Rata_VAC', 'Rata_VDC', 'Rata_IAC', 'Rata_LDC']] = rekap_df[['Rata_VAC', 'Rata_VDC', 'Rata_IAC', 'Rata_LDC']].round(2)

                filter_tabel = st.radio("Tampilkan Tabel:", ["Semua Site", "Hanya Site Dikunjungi Rutin (>1x)", "Hanya Site Pernah Dikunjungi (1x)"], horizontal=True)
                
                if "Rutin" in filter_tabel: tampil_rekap = rekap_df[rekap_df['Jumlah_Kunjungan'] > 1]
                elif "1x" in filter_tabel: tampil_rekap = rekap_df[rekap_df['Jumlah_Kunjungan'] == 1]
                else: tampil_rekap = rekap_df

                st.dataframe(tampil_rekap, use_container_width=True, height=450)

            with tab_master:
                st.markdown("### 🗂️ Explorer & Live Editor Database Dapot Power")
                pilih_sheet_dapot = st.selectbox("Pilih Sheet dari File Dapot_Power_Executive_Dashboard.xlsx:", ["DB_TREND", "DB_KPI", "Master_Data_Konsolidasi"])
                
                df_dapot_sheet = read_excel_cached(dapot_file, sheet_name=pilih_sheet_dapot)
                edited_dapot = st.data_editor(df_dapot_sheet, use_container_width=True, num_rows="dynamic", key=f"ed_dapot_{pilih_sheet_dapot}")

                if st.button("💾 Simpan Perubahan & Download File Dapot Power", type="primary"):
                    state_dapot = st.session_state[f"ed_dapot_{pilih_sheet_dapot}"]
                    if state_dapot.get("edited_rows") or state_dapot.get("added_rows"):
                        with st.spinner("Menyuntikkan pembaruan ke file Dapot_Power_Executive_Dashboard.xlsx..."):
                            import openpyxl
                            wb_d = openpyxl.load_workbook(dapot_file)
                            ws_d = wb_d[pilih_sheet_dapot]
                            headers_d = list(df_dapot_sheet.columns)

                            for r_idx_str, c_changes in state_dapot.get("edited_rows", {}).items():
                                xl_row = int(r_idx_str) + 2
                                for c_name, n_val in c_changes.items():
                                    if c_name in headers_d: ws_d.cell(row=xl_row, column=headers_d.index(c_name) + 1).value = n_val

                            for add_r in state_dapot.get("added_rows", []):
                                xl_row = ws_d.max_row + 1
                                for c_name, n_val in add_r.items():
                                    if c_name in headers_d: ws_d.cell(row=xl_row, column=headers_d.index(c_name) + 1).value = n_val

                            buf_d = BytesIO(); wb_d.save(buf_d)
                            st.success("✅ Perubahan berhasil disimpan ke dalam struktur asli Excel!")
                            st.download_button("📥 Download Dapot_Power_Executive_Dashboard.xlsx Terupdate", data=buf_d.getvalue(), file_name=f"Updated_{dapot_file}", mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet", type="primary")
                    else:
                        st.warning("Belum ada sel yang diubah pada tabel di atas.")

        except Exception as e:
            st.error(f"Terjadi kesalahan saat memuat Dapot Power Executive Dashboard: {e}")
# =========================================================================
# MENU 5: DAILY TRACKER (TO DO LIST NOP PALANGKARAYA)
# =========================================================================
elif menu == "📅 Daily Tracker Palangkaraya":
    st.markdown("<h1 style='text-align: center;'>📅 Daily Tracker & To Do List (Palangkaraya)</h1>", unsafe_allow_html=True)
    st.markdown("<p style='text-align: center; opacity: 0.8;'>Update rencana kerja harian, realisasi site, dan pencocokan otomatis dengan Master Data Sidata.</p>", unsafe_allow_html=True)
    
    file_todo = "To Do List NOP Palangkaraya.xlsx"
    
    if not os.path.exists(file_todo):
        st.error(f"❌ File '{file_todo}' tidak ditemukan di server. Pastikan file tersebut sudah di-upload satu folder dengan aplikasi `App.py`.")
    else:
        fme_plk_list, _, site_plk_list, _ = load_master_lists()
        
        tab_harian, tab_database = st.tabs(["📝 Form Update Tim Onsite", "🗂️ Live Editor Master To-Do List"])
        
        with tab_harian:
            st.info("💡 **Khusus Tim Onsite / FME:** Data Site ID dan Nama FME di bawah ini sudah otomatis membaca dari Repositori Sidata. Silakan pilih dan laporkan progress harian.")
            
            c_td1, c_td2 = st.columns(2)
            with c_td1:
                pilih_site_td = st.selectbox("📍 Pilih Site ID Target (Auto-sync):", [""] + site_plk_list, key="td_site")
                pilih_fme_td = st.selectbox("👷 Nama FME / Tim Onsite:", [""] + fme_plk_list, key="td_fme")
            with c_td2:
                tanggal_td = st.date_input("📅 Tanggal Pengerjaan:", value=datetime.date.today())
                status_td = st.selectbox("📊 Update Status Harian:", ["On Progress", "Done", "Pending", "Cancel / Reschedule"])
            
            catatan_td = st.text_area("📝 Keterangan / Remarks Lapangan:", placeholder="Ketik keterangan mengenai pengerjaan hari ini atau alasan jika pending...")
            
            if st.button("💾 Simpan Update ke To Do List", type="primary", use_container_width=True):
                if not pilih_site_td or not pilih_fme_td:
                    st.error("⚠️ Site ID dan Nama FME wajib dipilih!")
                else:
                    with st.spinner("Menyimpan ke file Excel To Do List..."):
                        try:
                            import openpyxl
                            wb_todo = openpyxl.load_workbook(file_todo)
                            
                            sheet_target = "to do list" if "to do list" in wb_todo.sheetnames else wb_todo.sheetnames[0]
                            ws_todo = wb_todo[sheet_target]
                            
                            next_row = ws_todo.max_row + 1
                            
                            ws_todo.cell(row=next_row, column=11).value = tanggal_td.strftime("%Y-%m-%d") # Start Date
                            ws_todo.cell(row=next_row, column=3).value = pilih_site_td # Site ID
                            ws_todo.cell(row=next_row, column=4).value = pilih_fme_td # FME Name
                            ws_todo.cell(row=next_row, column=13).value = status_td # Status Task
                            ws_todo.cell(row=next_row, column=14).value = catatan_td # Remark
                            
                            wb_todo.save(file_todo)
                            st.success(f"✅ Laporan Diterima! Data kunjungan Site {pilih_site_td} oleh {pilih_fme_td} telah ter-update di file Tracker.")
                        except Exception as e:
                            st.error(f"Terjadi kesalahan saat menyimpan ke Excel: {e}")

        with tab_database:
            st.markdown("### 🗂️ Tabel Master To Do List (Bisa di-edit langsung oleh Admin)")
            st.write("Klik dua kali pada tabel di bawah (khususnya bagian kolom Status) untuk mengubah data secara cepat.")
            
            try:
                sheet_tujuan = "to do list"
                try:
                    df_aktif = read_excel_cached(file_todo, sheet_name=sheet_tujuan, header=2)
                    if df_aktif.empty or 'Site ID' not in df_aktif.columns:
                        df_aktif = read_excel_cached(file_todo, sheet_name=sheet_tujuan)
                except:
                    df_aktif = read_excel_cached(file_todo, sheet_name=0)
                
                edited_todo = st.data_editor(df_aktif, use_container_width=True, num_rows="dynamic", key="ed_todo_master")
                
                if st.button("💾 Simpan Perubahan Excel Permanen", type="primary"):
                    state_todo = st.session_state["ed_todo_master"]
                    if state_todo.get("edited_rows") or state_todo.get("added_rows"):
                        with st.spinner("Menyuntikkan pembaruan formasi langsung ke file Excel..."):
                            import openpyxl
                            wb_t = openpyxl.load_workbook(file_todo)
                            sheet_target = "to do list" if "to do list" in wb_t.sheetnames else wb_t.sheetnames[0]
                            ws_t = wb_t[sheet_target]
                            headers_t = list(df_aktif.columns)

                            offset_row = 3
                            
                            for r_idx_str, c_changes in state_todo.get("edited_rows", {}).items():
                                xl_row = int(r_idx_str) + 1 + offset_row
                                for c_name, n_val in c_changes.items():
                                    if c_name in headers_t:
                                        col_index = headers_t.index(c_name) + 1
                                        ws_t.cell(row=xl_row, column=col_index).value = n_val

                            for add_r in state_todo.get("added_rows", []):
                                xl_row = ws_t.max_row + 1
                                for c_name, n_val in add_r.items():
                                    if c_name in headers_t:
                                        col_index = headers_t.index(c_name) + 1
                                        ws_t.cell(row=xl_row, column=col_index).value = n_val

                            buf_t = BytesIO()
                            wb_t.save(buf_t)
                            st.success("✅ Perubahan berhasil disimpan permanen ke Master To Do List!")
                            st.download_button(
                                label="📥 Download File Tracker Terupdate",
                                data=buf_t.getvalue(),
                                file_name=f"Updated_To_Do_List_PLK.xlsx",
                                mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                                type="primary"
                            )
                    else:
                        st.warning("Belum ada data pada tabel yang Bapak ubah.")
                        
            except Exception as e:
                st.error(f"Gagal memproses Live Editor To Do List: {e}.")

# =========================================================================
# MENU 6: LOGISTIK & SPMS (MANAJEMEN STOK GUDANG)
# =========================================================================
elif menu == "📦 Logistik & SPMS":
    st.markdown("<h1 style='text-align: center;'>📦 Manajemen Logistik & SPMS</h1>", unsafe_allow_html=True)
    st.markdown("<p style='text-align: center; opacity: 0.8;'>Sistem Pencatatan Barang Masuk (Inbound), Barang Keluar (Outbound), dan Auto-Kalkulasi Stok.</p>", unsafe_allow_html=True)
    
    sheet_logistik = connect_gsheets("Data Logistik")
    if sheet_logistik is None:
        st.error("⚠️ **PENTING:** Silakan buat Sheet baru bernama **'Data Logistik'** (tanpa tanda kutip) di dalam file Google Spreadsheet Database Utama Anda.")
        st.stop()
        
    data_log = sheet_logistik.get_all_values()
    headers_standar = ["Tanggal", "Tipe Transaksi", "Kategori", "Merek / Nama Barang", "Untuk Genset/Mesin", "Jumlah", "PIC / Pengambil", "Site ID", "Keterangan"]
    
    # AUTO-FIX: Cek jika sheet kosong atau baris 1 berisi sel kosong/spasi gaib
    if not data_log or len(data_log) == 0 or not any(str(c).strip() for c in data_log[0]):
        sheet_logistik.clear() # Bersihkan sisa format/spasi di Sheet
        sheet_logistik.insert_row(headers_standar, index=1)
        data_log = [headers_standar]
        
    header_db = data_log[0]
    data_rows = data_log[1:]
    
    # Ratakan panjang baris dengan panjang header agar tidak Error di Pandas
    data_rows_padded = [row + [''] * (len(header_db) - len(row)) for row in data_rows]
    
    df_log = pd.DataFrame(data_rows_padded, columns=header_db)
    
    # Proteksi Darurat jika nama kolom masih belum terbaca sempurna
    if 'Tipe Transaksi' not in df_log.columns: df_log['Tipe Transaksi'] = ''
    if 'Merek / Nama Barang' not in df_log.columns: df_log['Merek / Nama Barang'] = ''
    if 'Jumlah' not in df_log.columns: df_log['Jumlah'] = '0'
    if 'Kategori' not in df_log.columns: df_log['Kategori'] = ''
    
    tab_in, tab_out, tab_dash = st.tabs(["📥 Inbound (Barang Masuk)", "📤 Outbound (Ambil Barang)", "📊 Dashboard Stok & Ekspor"])
    
    kategori_opsi = ["Filter Oli", "Filter Solar", "Filter Udara", "Oli Mesin", "Water Separator", "Sparepart Genset", "Modul Kontroler", "Material CME", "Lainnya"]
    
    with tab_in:
        st.markdown("### 📥 Form Barang Masuk Gudang (Inbound)")
        with st.container(border=True):
            c1, c2 = st.columns(2)
            with c1:
                tgl_in = st.date_input("📅 Tanggal Masuk", key="in_tgl")
                kat_in = st.selectbox("📂 Kategori Barang", kategori_opsi, key="in_kat")
                nama_in = st.text_input("🏷️ Merek / Nama Barang Lengkap (Wajib)", placeholder="Cth: Filter Oli Donaldson / Oli Meditran SX", key="in_nama")
            with c2:
                genset_in = st.text_input("⚙️ Diperuntukkan Untuk Tipe Genset", placeholder="Cth: Perkins 20kVA / Kubota / Umum", key="in_gen")
                qty_in = st.number_input("📦 Jumlah Masuk (Pcs/Liter)", min_value=1, value=1, key="in_qty")
                pic_in = st.text_input("👨‍💼 PIC / Penerima Gudang", placeholder="Nama penerima barang", key="in_pic")
            ket_in = st.text_input("📝 Keterangan (Opsional)", placeholder="Cth: DO dari pusat / Pembelian lokal", key="in_ket")
            
            if st.button("💾 SIMPAN DATA MASUK (INBOUND)", type="primary", use_container_width=True, key="btn_in"):
                if not nama_in: st.error("⚠️ Merek/Nama Barang Wajib diisi!")
                else:
                    with st.spinner("Menyimpan stok ke Database..."):
                        row_in = [tgl_in.strftime("%Y-%m-%d"), "INBOUND", kat_in, nama_in, genset_in, qty_in, pic_in, "-", ket_in]
                        sheet_logistik.insert_row(row_in, index=2)
                    st.success(f"✅ Berhasil! {qty_in} qty '{nama_in}' telah ditambahkan ke dalam stok gudang.")
                    st.rerun()
                    
    with tab_out:
        st.markdown("### 📤 Form Pengambilan Barang (Outbound)")
        with st.container(border=True):
            st.info("💡 **Catatan:** Pilih Merek Barang yang sudah terdaftar di stok, atau ketik manual jika tidak ada.")
            c3, c4 = st.columns(2)
            with c3:
                tgl_out = st.date_input("📅 Tanggal Pengambilan", key="out_tgl")
                kat_out = st.selectbox("📂 Kategori Barang Keluar", kategori_opsi, key="out_kat")
                
                stok_tersedia = df_log[df_log['Tipe Transaksi'] == 'INBOUND']['Merek / Nama Barang'].unique().tolist() if not df_log.empty else []
                if stok_tersedia:
                    nama_out = st.selectbox("🏷️ Pilih Merek / Nama Barang dari Stok", ["- Ketik Manual / Barang Baru -"] + stok_tersedia, key="out_nama_sel")
                    if nama_out == "- Ketik Manual / Barang Baru -":
                        nama_out = st.text_input("Ketik Merek / Nama Barang:", key="out_nama_txt")
                else:
                    nama_out = st.text_input("🏷️ Merek / Nama Barang Keluar", key="out_nama_txt")
                    
            with c4:
                qty_out = st.number_input("📦 Jumlah Diambil (Pcs/Liter)", min_value=1, value=1, key="out_qty")
                fme_plk_list, _, site_plk_list, _ = load_master_lists()
                pic_out = st.selectbox("👷 Nama Tim / FME Pengambil", [""] + fme_plk_list, key="out_pic") if fme_plk_list else st.text_input("👷 FME / Tim Pengambil")
                site_out = st.selectbox("📍 Digunakan Untuk Site Target", [""] + site_plk_list, key="out_site") if site_plk_list else st.text_input("📍 Digunakan Untuk Site")
                
            ket_out = st.text_input("📝 Keterangan Pengerjaan", placeholder="Cth: PM Genset / Perbaikan Breakdown...", key="out_ket")
            
            if st.button("🚀 SIMPAN PENGAMBILAN & POTONG STOK", type="primary", use_container_width=True, key="btn_out"):
                if not nama_out or not pic_out or not site_out: 
                    st.error("⚠️ Nama Barang, Nama FME, dan Site Target wajib diisi agar bisa dilacak!")
                else:
                    with st.spinner("Memotong stok dan merekam data..."):
                        row_out = [tgl_out.strftime("%Y-%m-%d"), "OUTBOUND", kat_out, nama_out, "-", qty_out, pic_out, site_out, ket_out]
                        sheet_logistik.insert_row(row_out, index=2)
                    st.success(f"✅ Pengambilan tercatat! {qty_out} qty '{nama_out}' telah dikeluarkan untuk Site {site_out}.")
                    st.rerun()
                    
    with tab_dash:
        st.markdown("### 📊 Kalkulasi Stok Otomatis (In vs Out)")
        if df_log.empty:
            st.info("Belum ada data transaksi Logistik di dalam Database.")
        else:
            try:
                df_calc = df_log.copy()
                df_calc['Jumlah'] = pd.to_numeric(df_calc['Jumlah'], errors='coerce').fillna(0)
                
                # Kalkulasi Inbound
                df_in = df_calc[df_calc['Tipe Transaksi'] == 'INBOUND'].groupby(['Kategori', 'Merek / Nama Barang', 'Untuk Genset/Mesin'])['Jumlah'].sum().reset_index()
                df_in.rename(columns={'Jumlah': 'Total Masuk'}, inplace=True)
                
                # Kalkulasi Outbound
                df_out = df_calc[df_calc['Tipe Transaksi'] == 'OUTBOUND'].groupby(['Kategori', 'Merek / Nama Barang'])['Jumlah'].sum().reset_index()
                df_out.rename(columns={'Jumlah': 'Total Keluar'}, inplace=True)
                
                # Gabungkan & Hitung Sisa
                df_stok = pd.merge(df_in, df_out, on=['Kategori', 'Merek / Nama Barang'], how='left').fillna(0)
                df_stok['Sisa Stok Tersedia'] = df_stok['Total Masuk'] - df_stok['Total Keluar']
                
                # Highlight Stok Habis
                def color_stock(val):
                    color = '#FF5252' if val <= 0 else '#00E676'
                    return f'color: {color}; font-weight: bold;'
                
                st.dataframe(df_stok.style.map(color_stock, subset=['Sisa Stok Tersedia']), use_container_width=True)
                
                st.markdown("### 📥 Ekspor & Rekap Laporan Logistik")
                c_dl1, c_dl2 = st.columns(2)
                with c_dl1:
                    csv_riwayat = df_log.to_csv(index=False).encode('utf-8')
                    st.download_button("📥 Download Riwayat Keluar-Masuk (CSV)", data=csv_riwayat, file_name=f"Riwayat_In_Out_Logistik_{datetime.date.today()}.csv", mime="text/csv", use_container_width=True)
                with c_dl2:
                    csv_stok = df_stok.to_csv(index=False).encode('utf-8')
                    st.download_button("📊 Download Ringkasan Stok (CSV)", data=csv_stok, file_name=f"Sisa_Stok_Logistik_{datetime.date.today()}.csv", mime="text/csv", type="primary", use_container_width=True)
                    
            except Exception as e:
                st.error(f"Terjadi kesalahan saat mengalkulasi stok: {e}")
# =========================================================================
# MENU 7: EXECUTIVE PROJECT DASHBOARD (TERPISAH & TERPROTEKSI SANDI)
# =========================================================================
elif menu == "💎 Executive Project Dashboard":
    if 'proj_dash_unlocked' not in st.session_state:
        st.session_state['proj_dash_unlocked'] = False

    if not st.session_state['proj_dash_unlocked']:
        st.markdown("<br>", unsafe_allow_html=True)
        c_lk1, c_lk2, c_lk3 = st.columns([2.5, 5, 2.5])
        with c_lk2:
            st.markdown("""
                <div style="background: linear-gradient(145deg, #0f172a, #1e293b); padding: 40px 30px; border-radius: 20px; border: 2px solid #FFD700; box-shadow: 0 15px 35px rgba(0,0,0,0.4); text-align: center; color: white;">
                    <div style="font-size: 60px; margin-bottom: 10px;">🔐</div>
                    <h2 style="color: #FFD700 !important; margin-bottom: 8px; font-weight: 900; letter-spacing: 1px;">EXECUTIVE PROJECT VAULT</h2>
                    <p style="color: #cbd5e1; font-size: 14px; margin-bottom: 25px;">Area Rahasia Finansial PO, SVA, Profit Margin & Interactive Timeline.<br>Masukkan kata sandi otorisasi untuk membuka dashboard.</p>
                </div>
            """, unsafe_allow_html=True)
            st.write("")
            pwd_vault = st.text_input("Kata Sandi Otorisasi Project:", type="password", placeholder="Masukkan kata sandi rahasia...", key="input_pwd_vault")
            if st.button("🔓 BUKA EXECUTIVE DASHBOARD", type="primary", use_container_width=True):
                import base64
                if base64.b64encode(pwd_vault.encode()).decode() == "WXVpMjAyNSM=":
                    st.session_state['proj_dash_unlocked'] = True
                    st.rerun()
                else:
                    st.error("❌ Kata sandi otorisasi salah! Akses ke Executive Project ditolak.")
        st.stop()

    # DESAIN LUXURY INTERNATIONAL PMO DASHBOARD
    st.markdown("""
        <style>
        .lux-header {
            background: linear-gradient(135deg, #090d16 0%, #172554 50%, #0f172a 100%);
            padding: 30px 35px;
            border-radius: 20px;
            border: 1px solid rgba(255, 215, 0, 0.35);
            box-shadow: 0 12px 30px rgba(0,0,0,0.25);
            color: white;
            margin-bottom: 24px;
        }
        .lux-kpi {
            background: linear-gradient(145deg, #111827, #1f2937);
            border: 1px solid rgba(255,255,255,0.1);
            border-left: 5px solid #FFD700;
            border-radius: 16px;
            padding: 18px 20px;
            color: white;
            box-shadow: 0 8px 20px rgba(0,0,0,0.18);
            margin-bottom: 15px;
        }
        .lux-kpi-title { font-size: 11px; text-transform: uppercase; letter-spacing: 1.2px; color: #9ca3af; font-weight: 700; margin-bottom: 6px; }
        .lux-kpi-val { font-size: 21px; font-weight: 900; color: #ffffff; letter-spacing: 0.5px; }
        .lux-kpi-sub { font-size: 11px; color: #10b981; font-weight: 700; margin-top: 4px; }
        .tl-card {
            background: var(--secondary-background-color);
            border: 1px solid rgba(128,128,128,0.25);
            border-radius: 14px;
            padding: 16px 20px;
            margin-bottom: 12px;
            box-shadow: 0 4px 12px rgba(0,0,0,0.06);
        }
        .badge-done { background: rgba(16, 185, 129, 0.15); color: #10b981; border: 1px solid #10b981; padding: 3px 10px; border-radius: 20px; font-size: 12px; font-weight: 800; }
        .badge-prog { background: rgba(245, 158, 11, 0.15); color: #f59e0b; border: 1px solid #f59e0b; padding: 3px 10px; border-radius: 20px; font-size: 12px; font-weight: 800; }
        .badge-wait { background: rgba(107, 114, 128, 0.15); color: #9ca3af; border: 1px solid #9ca3af; padding: 3px 10px; border-radius: 20px; font-size: 12px; font-weight: 800; }
        .badge-issue { background: rgba(239, 68, 68, 0.15); color: #ef4444; border: 1px solid #ef4444; padding: 3px 10px; border-radius: 20px; font-size: 12px; font-weight: 800; }
        </style>
    """, unsafe_allow_html=True)

    c_hd1, c_hd2 = st.columns([8, 2])
    with c_hd1:
        st.markdown("""
            <div class="lux-header">
                <div style="font-size: 12px; color: #FFD700; font-weight: 800; letter-spacing: 2px; text-transform: uppercase;">💎 Enterprise Portfolio & Financial Intelligence</div>
                <h1 style="color: #ffffff !important; font-size: 32px; font-weight: 900; margin: 5px 0 8px 0;">Executive Project & Timeline Command Center</h1>
                <div style="color: #cbd5e1; font-size: 14px;">Monitoring Real-Time Purchase Order (PO), Grafik Budget vs Dana Tersupport, Pending Keuangan & Analisa Project Mandek.</div>
            </div>
        """, unsafe_allow_html=True)
    with c_hd2:
        st.write("")
        if st.button("🔒 Kunci Vault", use_container_width=True):
            st.session_state['proj_dash_unlocked'] = False
            st.rerun()
        if st.button("🔄 Sync Cloud", type="primary", use_container_width=True):
            st.cache_data.clear()
            st.session_state['laporan_db'] = fetch_data_from_gsheets()
            st.rerun()

    db_all = st.session_state.get('laporan_db', [])
    proj_indices = [idx for idx, r in enumerate(db_all) if r.get('kategori') == 'Project']

    if not proj_indices:
        st.info("💡 Belum ada data Project (PO / SVA) yang di-submit. Silakan input melalui menu Form Pelaporan -> Kategori: Project.")
    else:
        total_po_val = 0.0
        total_cost_val = 0.0
        total_support_val = 0.0
        total_routes = 0
        routes_done = 0
        routes_mandek = 0
        all_timeline_rows = []

        for idx in proj_indices:
            r_p = db_all[idx]
            pd_p = r_p.get('proj_data', {})
            if pd_p.get('jenis_project') == 'PO':
                try:
                    v_po = float(str(pd_p.get('nilai_po', '0')).replace('.', '').replace(',', ''))
                except:
                    v_po = 0.0
                total_po_val += v_po
                
                tbl_p = pd_p.get('tabel_pekerjaan', [])
                if not tbl_p:
                    tbl_p = [{
                        "Site Asal": r_p.get('site_name', 'PO Utama'),
                        "Site Tujuan": "-",
                        "Pekerjaan Spesifik": pd_p.get('tipe_pekerjaan', 'Project PO'),
                        "Budget Per Site": v_po,
                        "Dana Tersupport": 0,
                        "Start Date": r_p.get('timestamp', '')[:10],
                        "End Date": r_p.get('timestamp', '')[:10],
                        "Progress (%)": pd_p.get('avg_progress', 0),
                        "Status": r_p.get('status', 'Not Started'),
                        "Kategori Kendala": "Lancar / Normal",
                        "Remark": pd_p.get('note_akhir', '-')
                    }]

                for r_i, row_t in enumerate(tbl_p):
                    asal_s = str(row_t.get('Site Asal', '')).strip()
                    tuj_s = str(row_t.get('Site Tujuan', '')).strip()
                    pek_s = str(row_t.get('Pekerjaan Spesifik', '')).strip()
                    try:
                        c_site = float(str(row_t.get('Budget Per Site', '0')).replace('.', '').replace(',', ''))
                    except:
                        c_site = 0.0
                    try:
                        sup_site = float(str(row_t.get('Dana Tersupport', '0')).replace('.', '').replace(',', ''))
                    except:
                        sup_site = 0.0

                    # Fallback jika budget per site masih 0 pada baris tunggal agar grafik langsung tampil proporsional
                    if c_site == 0 and len(tbl_p) == 1 and v_po > 0:
                        c_site = v_po

                    if not asal_s or asal_s == 'nan':
                        asal_s = f"Site-{r_i+1}"
                    if not pek_s or pek_s == 'nan':
                        pek_s = str(pd_p.get('tipe_pekerjaan', 'Pekerjaan PO'))

                    total_routes += 1
                    total_cost_val += c_site
                    total_support_val += sup_site
                    pending_fin_site = max(c_site - sup_site, 0.0)
                    
                    st_site = str(row_t.get('Status', 'Not Started') or 'Not Started')
                    kendala_site = str(row_t.get('Kategori Kendala', 'Lancar / Normal') or 'Lancar / Normal')
                    if kendala_site == 'nan':
                        kendala_site = 'Lancar / Normal'
                        
                    try:
                        pr_site = float(row_t.get('Progress (%)', 0) or 0)
                    except:
                        pr_site = 0.0

                    # Auto-deteksi jika dana belum tersupport 100% dan status belum Done
                    if st_site != 'Done' and pr_site < 100 and pending_fin_site > 0 and kendala_site == 'Lancar / Normal':
                        kendala_site = "Mandek - Pending Keuangan (Dana Belum Cair)"
                        
                    if st_site == 'Done' or pr_site >= 100:
                        routes_done += 1
                    elif st_site == 'Pending Issue' or "Mandek" in kendala_site or pending_fin_site > 0:
                        routes_mandek += 1
                        
                    rute_label = f"{asal_s} ➔ {tuj_s}" if tuj_s and tuj_s not in ['-', 'nan', ''] else asal_s
                    all_timeline_rows.append({
                        "No PO": str(pd_p.get('no_po', '-')),
                        "Mitra": str(pd_p.get('mitra', '-')),
                        "Program": str(pd_p.get('tipe_pekerjaan', '-')),
                        "Site Asal": asal_s,
                        "Site Tujuan": tuj_s if tuj_s not in ['nan', ''] else '-',
                        "Rute Site": f"[{pd_p.get('no_po', '-')}] {rute_label}",
                        "Pekerjaan Spesifik": pek_s,
                        "Total Budget (Rp)": c_site,
                        "Dana Tersupport (Rp)": sup_site,
                        "Pending Keuangan (Rp)": pending_fin_site,
                        "Start Date": str(row_t.get('Start Date', '') or r_p.get('timestamp', '')[:10]),
                        "End Date": str(row_t.get('End Date', '') or r_p.get('timestamp', '')[:10]),
                        "Progress (%)": pr_site,
                        "Status": st_site,
                        "Kategori Kendala": kendala_site,
                        "Remark": str(row_t.get('Remark', '') or pd_p.get('note_akhir', '-') or '-'),
                        "PIC Tim": r_p.get('teknisi', '-')
                    })

        total_margin_val = total_po_val - total_cost_val
        total_pending_fin = max(total_cost_val - total_support_val, 0.0)
        support_ratio = (total_support_val / total_cost_val * 100) if total_cost_val > 0 else 0.0
        margin_pct = (total_margin_val / total_po_val * 100) if total_po_val > 0 else 0.0
        route_completion_pct = (routes_done / total_routes * 100) if total_routes > 0 else 0.0

        # 6 LUXURY KPI CARDS
        k1, k2, k3 = st.columns(3)
        with k1:
            st.markdown(f"""
                <div class="lux-kpi" style="border-left-color: #3b82f6;">
                    <div class="lux-kpi-title">💰 Total Nilai Kontrak PO</div>
                    <div class="lux-kpi-val">Rp {total_po_val:,.0f}</div>
                    <div class="lux-kpi-sub" style="color: #60a5fa;">{len(proj_indices)} Kontrak Project Terdaftar</div>
                </div>
            """, unsafe_allow_html=True)
        with k2:
            st.markdown(f"""
                <div class="lux-kpi" style="border-left-color: #f59e0b;">
                    <div class="lux-kpi-title">💸 Total Budget / RAB Rute</div>
                    <div class="lux-kpi-val">Rp {total_cost_val:,.0f}</div>
                    <div class="lux-kpi-sub" style="color: #fbbf24;">Alokasi untuk {total_routes} Rute Site</div>
                </div>
            """, unsafe_allow_html=True)
        with k3:
            warna_m = "#10b981" if total_margin_val >= 0 else "#ef4444"
            st.markdown(f"""
                <div class="lux-kpi" style="border-left-color: {warna_m};">
                    <div class="lux-kpi-title">📈 Estimasi Net Margin PO</div>
                    <div class="lux-kpi-val" style="color: {warna_m};">Rp {total_margin_val:,.0f}</div>
                    <div class="lux-kpi-sub" style="color: {warna_m};">Profitability: {margin_pct:.1f}% dari Nilai PO</div>
                </div>
            """, unsafe_allow_html=True)

        k4, k5, k6 = st.columns(3)
        with k4:
            st.markdown(f"""
                <div class="lux-kpi" style="border-left-color: #10b981;">
                    <div class="lux-kpi-title">✅ Dana Sudah Tersupport (Cair)</div>
                    <div class="lux-kpi-val" style="color: #10b981;">Rp {total_support_val:,.0f}</div>
                    <div class="lux-kpi-sub" style="color: #34d399;">{support_ratio:.1f}% dari Total Budget Rute</div>
                </div>
            """, unsafe_allow_html=True)
        with k5:
            st.markdown(f"""
                <div class="lux-kpi" style="border-left-color: #ef4444;">
                    <div class="lux-kpi-title">🚨 Pending Keuangan (Belum Cair)</div>
                    <div class="lux-kpi-val" style="color: #ef4444;">Rp {total_pending_fin:,.0f}</div>
                    <div class="lux-kpi-sub" style="color: #f87171;">Dana tertahan / belum disupport</div>
                </div>
            """, unsafe_allow_html=True)
        with k6:
            st.markdown(f"""
                <div class="lux-kpi" style="border-left-color: #a855f7;">
                    <div class="lux-kpi-title">🎯 Clearance & Bottleneck</div>
                    <div class="lux-kpi-val">{routes_done}/{total_routes} Clear | <span style="color:#ef4444;">{routes_mandek} Tertahan</span></div>
                    <div class="lux-kpi-sub" style="color: #c084fc;">{route_completion_pct:.1f}% Rute Site Selesai</div>
                </div>
            """, unsafe_allow_html=True)

        tab_tl, tab_po_list, tab_export = st.tabs([
            "📊 Grafik Budget vs Support, Timeline & Analisa Mandek",
            "📂 Klik Nomor PO / SVA (Update Dana Support, Timeline & Foto)",
            "📥 Executive Financial Export"
        ])

        with tab_tl:
            if not all_timeline_rows:
                st.info("Belum ada rincian PO yang dimasukkan.")
            else:
                df_tl = pd.DataFrame(all_timeline_rows)

                # =====================================================================
                # 1. GRAFIK KOMPARASI TOTAL BUDGET VS DANA TERSUPPORT VS PENDING KEUANGAN
                # =====================================================================
                st.markdown("### 📊 1. Grafik Komparasi Budget vs Dana Tersupport & Pending Keuangan")
                cg_1, cg_2 = st.columns([1.4, 1])
                with cg_1:
                    with st.container(border=True):
                        st.markdown("**💰 Komparasi Total Budget vs Sudah Tersupport vs Pending Keuangan (Per Rute Site)**")
                        df_chart_fin = df_tl.set_index("Rute Site")[["Total Budget (Rp)", "Dana Tersupport (Rp)", "Pending Keuangan (Rp)"]]
                        ui_bar_chart(df_chart_fin, color=["#3b82f6", "#10b981", "#ef4444"])
                        st.caption("Biru = Total Budget Site | Hijau = Dana Sudah Tersupport | Merah = Pending Keuangan (Belum Cair)")

                with cg_2:
                    with st.container(border=True):
                        st.markdown("**🛑 Analisa Bottleneck: Kenapa Project Mandek / Tertahan?**")
                        df_kendala = df_tl["Kategori Kendala"].value_counts().reset_index()
                        df_kendala.columns = ["Alasan / Kendala", "Jumlah Site"]
                        ui_bar_chart(df_kendala.set_index("Alasan / Kendala"), color="#f59e0b")
                        
                        df_stuck = df_tl[(df_tl["Pending Keuangan (Rp)"] > 0) | (df_tl["Status"] != "Done")]
                        if not df_stuck.empty:
                            st.markdown("**🚨 Daftar Rute Site Tertahan / Butuh Support Keuangan:**")
                            for _, r_st in df_stuck.head(5).iterrows():
                                st.markdown(
                                    f"<div style='font-size:12px; padding:6px 10px; margin-bottom:4px; border-left:4px solid #ef4444; background:rgba(239,68,68,0.08); border-radius:6px;'>"
                                    f"<b>{r_st['Rute Site']}</b> — Pending Keuangan: <b style='color:#ef4444;'>Rp {r_st['Pending Keuangan (Rp)']:,.0f}</b><br>"
                                    f"📌 <i>Kendala: {r_st['Kategori Kendala']} | Remark: {r_st['Remark']}</i></div>",
                                    unsafe_allow_html=True
                                )

                st.markdown("<hr style='margin: 20px 0;'>", unsafe_allow_html=True)

                # =====================================================================
                # 2. VISUAL TIMELINE GANTT & ABSENSI PENGERJAAN SITE
                # =====================================================================
                st.markdown("### 🗓️ 2. Visual Timeline Chart & Absensi Progress Per-Site")
                cf1, cf2, cf3 = st.columns(3)
                with cf1:
                    flt_po = st.selectbox("🔍 Filter Nomor PO:", ["Semua PO"] + sorted(df_tl["No PO"].unique().tolist()))
                with cf2:
                    flt_stat = st.selectbox("📊 Filter Status Site:", ["Semua Status", "Done", "On Progress", "Not Started", "Pending Issue"])
                with cf3:
                    flt_cari = st.text_input("🔎 Cari Site Asal / Tujuan / Kendala:", placeholder="Ketik Site ID atau Kendala...")

                df_tl_view = df_tl.copy()
                if flt_po != "Semua PO":
                    df_tl_view = df_tl_view[df_tl_view["No PO"] == flt_po]
                if flt_stat != "Semua Status":
                    df_tl_view = df_tl_view[df_tl_view["Status"] == flt_stat]
                if flt_cari.strip():
                    q_tl = flt_cari.strip().lower()
                    df_tl_view = df_tl_view[
                        df_tl_view["Rute Site"].str.lower().str.contains(q_tl) |
                        df_tl_view["Pekerjaan Spesifik"].str.lower().str.contains(q_tl) |
                        df_tl_view["Kategori Kendala"].str.lower().str.contains(q_tl)
                    ]

                with st.container(border=True):
                    st.markdown("**📈 Grafik Perbandingan Progress Fisik (%) vs Persentase Support Keuangan (%)**")
                    df_prog_cmp = df_tl_view.copy()
                    df_prog_cmp["Support Keuangan (%)"] = df_prog_cmp.apply(
                        lambda x: min((x["Dana Tersupport (Rp)"] / x["Total Budget (Rp)"] * 100), 100.0) if x["Total Budget (Rp)"] > 0 else 0.0,
                        axis=1
                    ).round(1)
                    ui_bar_chart(df_prog_cmp.set_index("Rute Site")[["Progress (%)", "Support Keuangan (%)"]], color=["#a855f7", "#10b981"])
                    st.caption("Ungu = Progress Fisik Lapangan (%) | Hijau = Persentase Dana Sudah Disupport Keuangan (%)")

                for _, row_item in df_tl_view.iterrows():
                    prog_val = float(row_item["Progress (%)"])
                    b_tot = float(row_item["Total Budget (Rp)"])
                    b_sup = float(row_item["Dana Tersupport (Rp)"])
                    b_pend = float(row_item["Pending Keuangan (Rp)"])
                    sup_pct = min((b_sup / b_tot * 100), 100.0) if b_tot > 0 else 0.0

                    stat_val = row_item["Status"]
                    if stat_val == "Done" or prog_val >= 100:
                        b_cls = "badge-done"
                        bar_col = "#10b981"
                    elif stat_val == "On Progress":
                        b_cls = "badge-prog"
                        bar_col = "#f59e0b"
                    elif stat_val == "Pending Issue" or "Mandek" in str(row_item["Kategori Kendala"]):
                        b_cls = "badge-issue"
                        bar_col = "#ef4444"
                    else:
                        b_cls = "badge-wait"
                        bar_col = "#6b7280"

                    warna_kendala = "#ef4444" if ("Mandek" in str(row_item["Kategori Kendala"]) or b_pend > 0) and stat_val != "Done" else "#10b981"

                    st.markdown(f"""
                        <div class="tl-card" style="border-left: 6px solid {bar_col};">
                            <div style="display:flex; justify-content:space-between; align-items:center; flex-wrap:wrap; gap:8px; margin-bottom:8px;">
                                <div>
                                    <span style="font-size:12px; font-weight:800; color:#3b82f6; background:rgba(59,130,246,0.12); padding:3px 8px; border-radius:6px; margin-right:8px;">PO: {row_item['No PO']}</span>
                                    <span style="font-size:17px; font-weight:900;">📍 {row_item['Site Asal']} ➡️ {row_item['Site Tujuan']}</span>
                                </div>
                                <div style="display:flex; gap:8px; align-items:center;">
                                    <span style="font-size:12px; font-weight:800; color:{warna_kendala}; background:rgba(239,68,68,0.1); padding:3px 10px; border-radius:20px; border:1px solid {warna_kendala};">🚦 {row_item['Kategori Kendala']}</span>
                                    <span class="{b_cls}">{stat_val} ({prog_val:.0f}%)</span>
                                </div>
                            </div>
                            <div style="font-size:13px; opacity:0.92; margin-bottom:10px; display:flex; gap:16px; flex-wrap:wrap;">
                                <span>🛠️ <b>Pekerjaan:</b> {row_item['Pekerjaan Spesifik']}</span>
                                <span>💰 <b>Total Budget:</b> Rp {b_tot:,.0f}</span>
                                <span style="color:#10b981; font-weight:700;">✅ <b>Tersupport:</b> Rp {b_sup:,.0f} ({sup_pct:.0f}%)</span>
                                <span style="color:#ef4444; font-weight:700;">🚨 <b>Pending Keuangan:</b> Rp {b_pend:,.0f}</span>
                                <span>🗓️ <b>Timeline:</b> {row_item['Start Date']} ➔ {row_item['End Date']}</span>
                            </div>
                            <div style="font-size:11px; font-weight:700; margin-bottom:2px;">🛠️ Progress Fisik Lapangan ({prog_val:.0f}%)</div>
                            <div style="width:100%; background:rgba(128,128,128,0.2); height:8px; border-radius:10px; overflow:hidden; margin-bottom:6px;">
                                <div style="width:{min(max(prog_val, 0), 100)}%; background:{bar_col}; height:100%; border-radius:10px;"></div>
                            </div>
                            <div style="font-size:11px; font-weight:700; margin-bottom:2px;">💳 Pencairan / Support Keuangan ({sup_pct:.0f}%)</div>
                            <div style="width:100%; background:rgba(128,128,128,0.2); height:8px; border-radius:10px; overflow:hidden; margin-bottom:6px;">
                                <div style="width:{min(max(sup_pct, 0), 100)}%; background:#10b981; height:100%; border-radius:10px;"></div>
                            </div>
                            <div style="font-size:12px; opacity:0.8;">📝 <b>Keterangan / Kenapa Mandek:</b> {row_item['Remark']} | 👷 <b>PIC:</b> {row_item['PIC Tim']}</div>
                        </div>
                    """, unsafe_allow_html=True)

        with tab_po_list:
            st.markdown("### 📂 Pilih Nomor PO / SVA untuk Update Dana Support, Timeline & Foto")
            opsi_proj_select = {
                idx: f"[{db_all[idx].get('proj_data',{}).get('jenis_project','PO')}] {db_all[idx].get('site_name','Project')} — Mitra: {db_all[idx].get('proj_data',{}).get('mitra','-')} ({db_all[idx].get('status','-')})"
                for idx in proj_indices
            }
            sel_proj_idx = st.selectbox("📌 Klik / Pilih Nomor PO atau Tiket SVA:", options=proj_indices, format_func=lambda x: opsi_proj_select[x])

            if sel_proj_idx is not None:
                r = db_all[sel_proj_idx]
                pd_data = r.get('proj_data', {})

                st.markdown("<hr style='margin: 15px 0;'>", unsafe_allow_html=True)
                if pd_data.get('jenis_project') == 'PO':
                    val_po = float(str(pd_data.get('nilai_po', '0')).replace('.', '').replace(',', '')) if pd_data.get('nilai_po') else 0.0
                    t_cost = float(pd_data.get('total_cost', 0.0) or 0.0)
                    t_sup = float(pd_data.get('total_support', 0.0) or 0.0)
                    t_margin = float(pd_data.get('margin', 0.0) or 0.0)

                    c_m1, c_m2, c_m3, c_m4, c_m5 = st.columns(5)
                    c_m1.metric("💰 Nilai PO", f"Rp {val_po:,.0f}")
                    c_m2.metric("💸 Total Budget Site", f"Rp {t_cost:,.0f}")
                    c_m3.metric("✅ Dana Tersupport", f"Rp {t_sup:,.0f}")
                    c_m4.metric("🚨 Pending Keuangan", f"Rp {max(t_cost - t_sup, 0.0):,.0f}")
                    c_m5.metric("📈 Net Margin PO", f"Rp {t_margin:,.0f}")

                    st.markdown("#### 🛠️ Live Editor: Update Rute Site, Dana Tersupport, Alasan Mandek & Progress (%)")
                    st.info("💡 Klik dua kali pada sel tabel di bawah ini untuk mengisi **Dana Tersupport (Rp)** dari keuangan, memilih **Kategori Kendala (Kenapa Mandek)**, atau mengupdate tanggal & progress.")

                    df_edit_po = pd.DataFrame(pd_data.get('tabel_pekerjaan', []))
                    cols_standar = ["Site Asal", "Site Tujuan", "Pekerjaan Spesifik", "Budget Per Site", "Dana Tersupport", "Start Date", "End Date", "Progress (%)", "Status", "Kategori Kendala", "Remark"]
                    if df_edit_po.empty:
                        df_edit_po = pd.DataFrame(columns=cols_standar)
                    else:
                        for c_st in cols_standar:
                            if c_st not in df_edit_po.columns:
                                if c_st in ["Budget Per Site", "Dana Tersupport", "Progress (%)"]:
                                    df_edit_po[c_st] = 0
                                elif c_st == "Kategori Kendala":
                                    df_edit_po[c_st] = "Lancar / Normal"
                                else:
                                    df_edit_po[c_st] = ""
                        df_edit_po = df_edit_po[cols_standar]
                        df_edit_po["Budget Per Site"] = pd.to_numeric(df_edit_po["Budget Per Site"], errors='coerce').fillna(0)
                        df_edit_po["Dana Tersupport"] = pd.to_numeric(df_edit_po["Dana Tersupport"], errors='coerce').fillna(0)
                        df_edit_po["Progress (%)"] = pd.to_numeric(df_edit_po["Progress (%)"], errors='coerce').fillna(0)

                    updated_po_table = st.data_editor(
                        df_edit_po,
                        num_rows="dynamic",
                        use_container_width=True,
                        column_config={
                            "Site Asal": st.column_config.TextColumn("Site Asal (Dari)"),
                            "Site Tujuan": st.column_config.TextColumn("Site Tujuan (Ke)"),
                            "Pekerjaan Spesifik": st.column_config.TextColumn("Pekerjaan"),
                            "Budget Per Site": st.column_config.NumberColumn("Total Budget (Rp)", min_value=0, default=0, step=100000),
                            "Dana Tersupport": st.column_config.NumberColumn("Dana Tersupport (Rp)", min_value=0, default=0, step=100000),
                            "Start Date": st.column_config.TextColumn("Mulai (YYYY-MM-DD)"),
                            "End Date": st.column_config.TextColumn("Selesai (YYYY-MM-DD)"),
                            "Progress (%)": st.column_config.NumberColumn("Progress (%)", min_value=0, max_value=100, default=0),
                            "Status": st.column_config.SelectboxColumn("Status", options=["Not Started", "On Progress", "Done", "Pending Issue"]),
                            "Kategori Kendala": st.column_config.SelectboxColumn("Kenapa Mandek / Kendala", options=[
                                "Lancar / Normal",
                                "Mandek - Pending Keuangan (Dana Belum Cair)",
                                "Mandek - Pending Material / Sparepart",
                                "Mandek - Izin Site / Community Issue",
                                "Mandek - Kendala Teknis / Cuaca"
                            ]),
                            "Remark": st.column_config.TextColumn("Detail Kendala / Catatan")
                        },
                        key=f"lux_edit_po_{sel_proj_idx}"
                    )

                    c_up1, c_up2 = st.columns(2)
                    with c_up1:
                        stat_opts_p = ["Not Started", "On Progress", "Done", "Pending"]
                        cur_st_p = r.get('status', 'On Progress')
                        new_status_proj = st.selectbox("Status Keseluruhan PO:", stat_opts_p, index=stat_opts_p.index(cur_st_p) if cur_st_p in stat_opts_p else 1, key=f"lux_stat_{sel_proj_idx}")
                        new_nilai_po = st.text_input("Revisi Nilai PO Total (Rp):", value=str(pd_data.get('nilai_po', '0')), key=f"lux_valpo_{sel_proj_idx}")
                    with c_up2:
                        new_note_proj = st.text_area("Update Catatan / Remark PO:", value=str(pd_data.get('note_akhir', '')), key=f"lux_note_{sel_proj_idx}")

                    st.markdown("#### 📸 Upload Dokumentasi Foto Susulan (Saat Pekerjaan Dimulai / Selesai)")
                    cu_f1, cu_f2, cu_f3 = st.columns(3)
                    with cu_f1:
                        up_p_awal = st.file_uploader("➕ Tambah Foto Before", accept_multiple_files=True, key=f"lux_f_awal_{sel_proj_idx}")
                    with cu_f2:
                        up_p_pros = st.file_uploader("➕ Tambah Foto On Progress", accept_multiple_files=True, key=f"lux_f_pros_{sel_proj_idx}")
                    with cu_f3:
                        up_p_akh = st.file_uploader("➕ Tambah Foto After / BAST", accept_multiple_files=True, key=f"lux_f_akh_{sel_proj_idx}")

                    if st.button("💾 SIMPAN UPDATE PO & KALKULASI ULANG MARGIN", type="primary", use_container_width=True, key=f"btn_save_lux_{sel_proj_idx}"):
                        with st.spinner("⏳ Menyimpan perubahan tabel, dana support & mengunggah foto susulan ke Cloud..."):
                            r['proj_data']['tabel_pekerjaan'] = updated_po_table.astype(str).to_dict('records')
                            r['proj_data']['nilai_po'] = new_nilai_po
                            r['proj_data']['note_akhir'] = new_note_proj
                            try:
                                v_po_new = float(str(new_nilai_po).replace('.', '').replace(',', ''))
                            except:
                                v_po_new = 0.0
                            t_cost_new = pd.to_numeric(updated_po_table["Budget Per Site"], errors='coerce').fillna(0).sum() if not updated_po_table.empty else 0.0
                            t_sup_new = pd.to_numeric(updated_po_table["Dana Tersupport"], errors='coerce').fillna(0).sum() if not updated_po_table.empty else 0.0
                            avg_p_new = pd.to_numeric(updated_po_table["Progress (%)"], errors='coerce').fillna(0).mean() if not updated_po_table.empty else 0.0

                            r['proj_data']['total_cost'] = float(t_cost_new)
                            r['proj_data']['total_support'] = float(t_sup_new)
                            r['proj_data']['pending_finance'] = float(max(t_cost_new - t_sup_new, 0.0))
                            r['proj_data']['margin'] = float(v_po_new - t_cost_new)
                            r['proj_data']['avg_progress'] = float(avg_p_new)
                            r['status'] = new_status_proj

                            if up_p_awal:
                                r['proj_data']['url_awal'] = r['proj_data'].get('url_awal', []) + upload_multiple_images(up_p_awal, "proj_po")
                            if up_p_pros:
                                r['proj_data']['url_proses'] = r['proj_data'].get('url_proses', []) + upload_multiple_images(up_p_pros, "proj_po")
                            if up_p_akh:
                                r['proj_data']['url_akhir'] = r['proj_data'].get('url_akhir', []) + upload_multiple_images(up_p_akh, "proj_po")

                            sheet = connect_gsheets(r.get('_sheet', 'Project'))
                            if sheet:
                                sheet.update_cell(r.get('_row', sel_proj_idx + 2), 5, new_status_proj)
                                sheet.update_cell(r.get('_row', sel_proj_idx + 2), 6, json.dumps(r, default=str))
                                st.cache_data.clear()
                        st.success("✅ BERHASIL! Data Rute, Dana Tersupport, Timeline, Margin & Foto PO telah diperbarui.")
                        st.rerun()

                    st.markdown("#### 🖼️ Galeri Dokumentasi PO Saat Ini")
                    gt1, gt2, gt3 = st.tabs(["📸 Foto Kondisi Awal (Before)", "📸 Foto Pengerjaan (On Progress)", "📸 Foto Selesai (After / BAST)"])
                    with gt1: tampilkan_grid_foto(pd_data.get('url_awal', []), "Dokumentasi Kondisi Awal")
                    with gt2: tampilkan_grid_foto(pd_data.get('url_proses', []), "Dokumentasi On Progress")
                    with gt3: tampilkan_grid_foto(pd_data.get('url_akhir', []), "Dokumentasi Akhir / BAST")

                else:
                    st.info(f"**📋 Detail SVA:** Site: {pd_data.get('site_sva','-')} | No Tiket: {pd_data.get('no_tiket','-')} | Tanggal: {pd_data.get('tgl_sva','-')} | Temuan: {pd_data.get('temuan_sva','-')}")
                    gt_s1, gt_s2 = st.tabs(["📸 View Site & Perangkat", "📸 Temuan / Defect SVA"])
                    with gt_s1: tampilkan_grid_foto(pd_data.get('url_sva1', []), "Foto View Site & Perangkat")
                    with gt_s2: tampilkan_grid_foto(pd_data.get('url_sva2', []), "Foto Temuan Defect")

        with tab_export:
            st.markdown("### 📥 Unduh Rekapitulasi Finansial & Timeline Project (Excel)")
            if all_timeline_rows:
                df_exp_proj = pd.DataFrame(all_timeline_rows)
                st.dataframe(df_exp_proj, use_container_width=True)
                try:
                    buf_p = BytesIO()
                    with pd.ExcelWriter(buf_p, engine='xlsxwriter') as writer:
                        df_exp_proj.to_excel(writer, index=False, sheet_name='Executive_Project')
                    st.download_button("💎 Download Executive Project Report (.xlsx)", data=buf_p.getvalue(), file_name=f"Executive_Project_Report_{datetime.date.today()}.xlsx", mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet", type="primary", use_container_width=True)
                except Exception:
                    csv_p = df_exp_proj.to_csv(index=False).encode('utf-8')
                    st.download_button("💎 Download Executive Project Report (.csv)", data=csv_p, file_name=f"Executive_Project_Report_{datetime.date.today()}.csv", mime="text/csv", type="primary", use_container_width=True)
            else:
                st.info("Belum ada rincian tabel rute PO untuk diekspor.")# -------------------------------------------------------------------------
# =========================================================================
# MENU 9: SYSINFOMAP (PETA GEOSPASIAL, POPULASI KABUPATEN & KOMPARASI SITE)
# =========================================================================
elif menu == "🗺️ Sysinfomap (Peta & Populasi Site)":
    st.markdown("<h1 style='text-align: center;'>🗺️ Sysinfomap — Geospatial & Site Population Intelligence</h1>", unsafe_allow_html=True)
    st.markdown("<p style='text-align: center; opacity: 0.85;'>Pemetaan Interaktif 910+ Site Telekomunikasi, Distribusi Populasi per Kabupaten, dan Komparasi Otomatis Database SIDATA</p>", unsafe_allow_html=True)

    df_kml_base, kml_filename = get_default_kml_df()

    with st.expander("📂 Sumber Data Peta KML / Update File KML Baru (Opsional)", expanded=df_kml_base.empty):
        st.info("💡 Sistem otomatis membaca file **`newlist_910.kml`** di server dengan teknologi *In-Memory Cache* (sangat ringan & tanpa loading ulang). Anda juga dapat mengunggah file `.kml` terbaru di sini untuk pembaruan data instan.")
        up_kml = st.file_uploader("Unggah File KML Baru (.kml):", type=["kml", "xml"], key="up_sysinfomap_kml")
        if up_kml is not None:
            df_kml_base = load_sysinfomap_kml(up_kml.getvalue(), len(up_kml.getvalue()))
            kml_filename = up_kml.name

    if df_kml_base.empty:
        st.error("⚠️ File KML (`newlist_910.kml`) belum ditemukan di folder aplikasi. Pastikan file `newlist_910.kml` sudah diletakkan satu folder dengan `App.py` atau unggah melalui kotak di atas.")
        st.stop()

    # INTEGRASI & KOMPARASI REAL-TIME DENGAN DATA KUNJUNGAN SIDATA & MASTER EXCEL
    db_live = st.session_state.get('laporan_db', [])
    visited_map = {}
    for r in db_live:
        if r.get('kategori') in ['Project', 'SiHD']:
            continue
        raw_sn = str(r.get('pb_site_id') or r.get('site_name') or '').strip().upper()
        if not raw_sn:
            continue
        sid_key = raw_sn[:6] if len(raw_sn) >= 6 and raw_sn[:3].isalpha() and raw_sn[3:6].isdigit() else raw_sn.split()[0].split('-')[0]
        if sid_key not in visited_map:
            visited_map[sid_key] = {
                "count": 0,
                "last_date": str(r.get('timestamp', '-'))[:10],
                "last_status": r.get('status', '-'),
                "last_tek": r.get('teknisi') or r.get('pb_te') or '-',
                "last_action": r.get('action', '-')
            }
        visited_map[sid_key]["count"] += 1

    master_excel_ids = set()
    for xl_file in ["Site ID Sidata Palangkaraya.xlsx", "Site ID Sidata Pangkalanbun.xlsx"]:
        if os.path.exists(xl_file):
            try:
                df_xl_m = read_excel_cached(xl_file)
                if 'Site ID' in df_xl_m.columns:
                    for s_item in df_xl_m['Site ID'].dropna().astype(str):
                        s_clean = s_item.strip().upper()
                        if s_clean:
                            master_excel_ids.add(s_clean[:6] if len(s_clean) >= 6 and s_clean[:3].isalpha() and s_clean[3:6].isdigit() else s_clean)
            except Exception:
                pass

    combat_excel_ids = set()
    if os.path.exists("Combat NOP Palangkaraya_1.xlsx"):
        try:
            df_cb_xl = read_excel_cached("Combat NOP Palangkaraya_1.xlsx", sheet_name="Combat NOP Palangkaraya")
            if 'siteid' in df_cb_xl.columns:
                combat_excel_ids = set(df_cb_xl['siteid'].dropna().astype(str).str.strip().str.upper())
                master_excel_ids.update(combat_excel_ids)
        except Exception:
            pass

    df_map = df_kml_base.copy()
    df_map['Jumlah_Visit_PM'] = df_map['site_id'].map(lambda x: visited_map.get(x, {}).get('count', 0))
    df_map['Status_Kunjungan'] = df_map['Jumlah_Visit_PM'].apply(lambda c: f"Sudah Dikunjungi ({c}x)" if c > 0 else "Belum Dikunjungi")
    df_map['Status_PM_Terakhir'] = df_map['site_id'].map(lambda x: visited_map.get(x, {}).get('last_status', 'Belum Ada Laporan'))
    df_map['Tgl_Visit_Terakhir'] = df_map['site_id'].map(lambda x: visited_map.get(x, {}).get('last_date', '-'))
    df_map['Teknisi_Terakhir'] = df_map['site_id'].map(lambda x: visited_map.get(x, {}).get('last_tek', '-'))
    df_map['Status_Master_Excel'] = df_map['site_id'].apply(lambda x: "Terdaftar di Master" if x in master_excel_ids else "Site Baru (Hanya di KML)")
    df_map['Terdaftar_Combat'] = df_map['site_id'].apply(lambda x: "Ya (Site Combat)" if (x in combat_excel_ids or x.startswith("COH")) else "Reguler / Macro")

    tot_kml = len(df_map)
    tot_kab = df_map['Kabupaten'].nunique()
    tot_visited = int((df_map['Jumlah_Visit_PM'] > 0).sum())
    tot_unvisited = tot_kml - tot_visited
    pct_visited = (tot_visited / tot_kml * 100) if tot_kml > 0 else 0.0

    ux_hero(
        f"SUMBER DATA: {kml_filename or 'KML AKTIF'}",
        "Sysinfomap Geospatial Command Center",
        "Peta Populasi Site Kalimantan, Filter Multi-Dimensi (Kabupaten, RTS, Owner, Power, Class), & Komparasi Real-Time Database SIDATA.",
        stats=[
            (f"{tot_kml} Site", "Total Populasi KML"),
            (f"{tot_kab} Wilayah", "Cakupan Kabupaten/Area"),
            (f"{tot_visited} Site ({pct_visited:.1f}%)", "Sudah Terinspeksi PM"),
            (f"{tot_unvisited} Site", "Belum Dikunjungi (Target)")
        ],
        split=[
            ("Sudah Dikunjungi PM", tot_visited, "#10B981"),
            ("Belum Dikunjungi", tot_unvisited, "#3B82F6")
        ]
    )

    with st.container(border=True):
        st.markdown("#### 🔍 Filter Geospasial & Populasi Site")
        f1, f2, f3, f4 = st.columns(4)
        with f1:
            opsi_kab = ["Semua Kabupaten/Wilayah"] + sorted(df_map['Kabupaten'].unique().tolist())
            sel_kab = st.selectbox("🏛️ Kabupaten / Wilayah:", opsi_kab, key="sim_kab")
        with f2:
            opsi_cls = ["Semua Kelas"] + sorted(df_map['site_class'].unique().tolist())
            sel_cls = st.selectbox("⭐ Site Class:", opsi_cls, key="sim_cls")
        with f3:
            opsi_rts = ["Semua RTS"] + sorted(df_map['RTS'].unique().tolist())
            sel_rts = st.selectbox("👷 PIC RTS:", opsi_rts, key="sim_rts")
        with f4:
            opsi_vis = ["Semua Status Visit", "Sudah Dikunjungi PM", "Belum Dikunjungi", "Site Combat (COH)"]
            sel_vis = st.selectbox("📋 Filter Komparasi / Visit:", opsi_vis, key="sim_vis")

        f5, f6, f7 = st.columns([1.2, 1.2, 1.6])
        with f5:
            opsi_own = ["Semua Owner"] + sorted(df_map['Site_Owner'].unique().tolist())
            sel_own = st.selectbox("🏢 Tower Owner:", opsi_own, key="sim_own")
        with f6:
            opsi_pwr = ["Semua Power Source"] + sorted(df_map['Power_Source'].unique().tolist())
            sel_pwr = st.selectbox("⚡ Power Source:", opsi_pwr, key="sim_pwr")
        with f7:
            cari_sim = st.text_input("🔎 Cari Cepat Site ID / Sitename / Rectifier:", placeholder="Ketik misal: PLK001, BNT, ZTE, HUAWEI...", key="sim_search")

    df_view = df_map.copy()
    if sel_kab != "Semua Kabupaten/Wilayah":
        df_view = df_view[df_view['Kabupaten'] == sel_kab]
    if sel_cls != "Semua Kelas":
        df_view = df_view[df_view['site_class'] == sel_cls]
    if sel_rts != "Semua RTS":
        df_view = df_view[df_view['RTS'] == sel_rts]
    if sel_own != "Semua Owner":
        df_view = df_view[df_view['Site_Owner'] == sel_own]
    if sel_pwr != "Semua Power Source":
        df_view = df_view[df_view['Power_Source'] == sel_pwr]
    if sel_vis == "Sudah Dikunjungi PM":
        df_view = df_view[df_view['Jumlah_Visit_PM'] > 0]
    elif sel_vis == "Belum Dikunjungi":
        df_view = df_view[df_view['Jumlah_Visit_PM'] == 0]
    elif sel_vis == "Site Combat (COH)":
        df_view = df_view[df_view['Terdaftar_Combat'] == "Ya (Site Combat)"]

    if cari_sim.strip():
        q_s = cari_sim.strip().lower()
        df_view = df_view[
            df_view['site_id'].str.lower().str.contains(q_s) |
            df_view['Sitename'].str.lower().str.contains(q_s) |
            df_view['BrandRecti'].str.lower().str.contains(q_s) |
            df_view['Type_Hub'].str.lower().str.contains(q_s)
        ]

    m_s1, m_s2, m_s3, m_s4 = st.columns(4)
    m_s1.metric("📍 Site Sesuai Filter", f"{len(df_view)} Site")
    m_s2.metric("✅ Sudah Dikunjungi (Filter)", f"{int((df_view['Jumlah_Visit_PM'] > 0).sum())} Site")
    m_s3.metric("⏳ Belum Dikunjungi (Filter)", f"{int((df_view['Jumlah_Visit_PM'] == 0).sum())} Site")
    m_s4.metric("🗼 Tower Owner Terbanyak", df_view['Site_Owner'].mode()[0] if not df_view.empty else "-")

    tab_peta, tab_pop, tab_comp, tab_detail = st.tabs([
        "🗺️ Peta Interaktif & Navigasi",
        "📊 Populasi Kabupaten & Statistik Infrastruktur",
        "🔄 Komparasi Database & Gap Kunjungan",
        "📋 Tabel Data Lengkap & Inspector Site"
    ])

    with tab_peta:
        st.markdown("#### 🗺️ Interactive Tower Radar Map & Live Navigation")
        if df_view.empty:
            st.warning("Tidak ada site yang cocok dengan filter saat ini.")
        else:
            c_mp_ctrl1, c_mp_ctrl2, c_mp_ctrl3 = st.columns([1.4, 1.3, 1.3])
            with c_mp_ctrl1:
                mode_warna = st.radio(
                    "🎨 Mode Warna Ikon Tower:",
                    ["Kelas Site (Diamond/Platinum/Gold/Silver/Bronze)", "Status Kunjungan PM (Hijau = Sudah Visit, Merah = Belum)"],
                    horizontal=True, key="sim_color_mode"
                )
            with c_mp_ctrl2:
                map_theme = st.selectbox(
                    "🛰️ Tema Peta (Base Layer):",
                    ["Dark NOC Command Center (Mewah)", "Citra Satelit (Esri World Imagery)", "Peta Jalan Terang (Street Map)"],
                    key="sim_map_theme"
                )
            with c_mp_ctrl3:
                anim_mode = st.selectbox(
                    "✨ Efek Animasi & Cluster Tower:",
                    ["Radar Pulse Aktif + Clustering Pintar", "Semua Tower Tampil Langsung (Tanpa Cluster)", "Radar Pulse Khusus Site Prioritas / Belum Visit"],
                    key="sim_anim_mode"
                )

            c_mp1, c_mp2 = st.columns([2.35, 1.15])
            with c_mp2:
                with st.container(border=True):
                    st.markdown("#### 📌 Inspector & Auto Fly-To Site")
                    opsi_ins = ["🎯 Tampilkan Semua (Auto-Center)"] + df_view['site_id'].tolist()
                    map_label_ins = {"🎯 Tampilkan Semua (Auto-Center)": "🎯 Tampilkan Semua Site di Peta"}
                    for _, r_row in df_view.iterrows():
                        map_label_ins[r_row['site_id']] = f"🗼 {r_row['site_id']} — {r_row['Sitename']} ({r_row['Kabupaten']})"
                    
                    pilih_sid_map = st.selectbox(
                        "Pilih Site ID (Peta otomatis Zoom/Fly-To ke Tower):",
                        opsi_ins, format_func=lambda x: map_label_ins.get(x, x), key="sim_pick_site"
                    )
                    
                    if pilih_sid_map and pilih_sid_map != "🎯 Tampilkan Semua (Auto-Center)":
                        s_info = df_view[df_view['site_id'] == pilih_sid_map].iloc[0]
                    else:
                        s_info = df_view.iloc[0]

                    gmaps_url = f"https://www.google.com/maps?q={s_info['lat']},{s_info['lon']}"
                    st.markdown(f"""
                        <div style="padding:14px 16px; border-radius:14px; background:var(--ux-card2); border-left:5px solid #2563EB; margin-bottom:12px; font-size:13.5px; line-height:1.65; box-shadow:0 4px 12px rgba(0,0,0,0.06);">
                            <div style="font-size:16px; font-weight:800; margin-bottom:6px; color:var(--ux-ink);">🗼 {s_info['site_id']} · {s_info['Sitename']}</div>
                            <div>🏛️ <b>Kabupaten:</b> {s_info['Kabupaten']}</div>
                            <div>📍 <b>Koordinat:</b> <code>{s_info['lat']:.5f}, {s_info['lon']:.5f}</code></div>
                            <div>🏢 <b>Owner & Class:</b> {s_info['Site_Owner']} (<b>{s_info['site_class']}</b>)</div>
                            <div>⚡ <b>Power Source:</b> {s_info['Power_Source']}</div>
                            <div>🔌 <b>Rectifier & Hub:</b> {s_info['BrandRecti']} · {s_info['Type_Hub']}</div>
                            <div>📡 <b>Transport & Akses:</b> {s_info['Transport_type']} · {s_info['Access']}</div>
                            <div>👷 <b>PIC RTS:</b> {s_info['RTS']}</div>
                            <hr style="margin:8px 0;">
                            <div>📋 <b>Status PM di SIDATA:</b> <b>{s_info['Status_Kunjungan']}</b></div>
                            <div>🗓️ <b>Visit Terakhir:</b> {s_info['Tgl_Visit_Terakhir']} ({s_info['Status_PM_Terakhir']})</div>
                        </div>
                    """, unsafe_allow_html=True)
                    st.link_button("🧭 BUKA NAVIGASI GOOGLE MAPS KE SITE INI", gmaps_url, type="primary", use_container_width=True)

            with c_mp1:
                color_map_cls = {
                    "Diamond": "#a855f7",
                    "Platinum": "#ef4444",
                    "Gold": "#f59e0b",
                    "Silver": "#38bdf8",
                    "Bronze": "#10b981"
                }
                is_visit_mode = "Status Kunjungan" in mode_warna
                use_cluster = "Tanpa Cluster" not in anim_mode
                pulse_only_prio = "Khusus Site Prioritas" in anim_mode

                markers_payload = []
                for _, r_m in df_view.iterrows():
                    vis_cnt = int(r_m['Jumlah_Visit_PM'])
                    s_cls = str(r_m['site_class'])
                    if is_visit_mode:
                        col_hex = "#10b981" if vis_cnt > 0 else "#ef4444"
                    else:
                        col_hex = color_map_cls.get(s_cls, "#3b82f6")

                    if pulse_only_prio:
                        do_pulse = (vis_cnt == 0) or (s_cls in ["Diamond", "Platinum", "Gold"]) or str(r_m['site_id']).startswith("COH")
                    else:
                        do_pulse = True

                    markers_payload.append({
                        "id": str(r_m['site_id']),
                        "name": str(r_m['Sitename']),
                        "kab": str(r_m['Kabupaten']),
                        "cls": s_cls,
                        "own": str(r_m['Site_Owner']),
                        "pwr": str(r_m['Power_Source']),
                        "rect": str(r_m['BrandRecti']),
                        "hub": str(r_m['Type_Hub']),
                        "trm": str(r_m['Transport_type']),
                        "rts": str(r_m['RTS']),
                        "vis": str(r_m['Status_Kunjungan']),
                        "ldate": str(r_m['Tgl_Visit_Terakhir']),
                        "lstat": str(r_m['Status_PM_Terakhir']),
                        "lat": round(float(r_m['lat']), 6),
                        "lon": round(float(r_m['lon']), 6),
                        "color": col_hex,
                        "pulse": 1 if do_pulse else 0,
                        "is_combat": 1 if str(r_m['Terdaftar_Combat']) == "Ya (Site Combat)" else 0
                    })

                if pilih_sid_map and pilih_sid_map != "🎯 Tampilkan Semua (Auto-Center)":
                    focus_lat = float(s_info['lat'])
                    focus_lon = float(s_info['lon'])
                    focus_zoom = 14
                    focus_id = str(s_info['site_id'])
                else:
                    focus_lat = float(df_view['lat'].mean())
                    focus_lon = float(df_view['lon'].mean())
                    focus_zoom = 7
                    focus_id = ""

                if "Satelit" in map_theme:
                    tile_url = "https://server.arcgisonline.com/ArcGIS/rest/services/World_Imagery/MapServer/tile/{z}/{y}/{x}"
                    tile_sub = ""
                elif "Terang" in map_theme:
                    tile_url = "https://{s}.basemaps.cartocdn.com/rastertiles/voyager/{z}/{x}/{y}{r}.png"
                    tile_sub = "abcd"
                else:
                    tile_url = "https://{s}.basemaps.cartocdn.com/dark_all/{z}/{x}/{y}{r}.png"
                    tile_sub = "abcd"

                markers_json = json.dumps(markers_payload, ensure_ascii=False)
                
                leaflet_html = f"""
                <!DOCTYPE html>
                <html>
                <head>
                    <meta charset="utf-8" />
                    <link rel="stylesheet" href="https://unpkg.com/leaflet@1.9.4/dist/leaflet.css" />
                    <link rel="stylesheet" href="https://unpkg.com/leaflet.markercluster@1.5.3/dist/MarkerCluster.css" />
                    <link rel="stylesheet" href="https://unpkg.com/leaflet.markercluster@1.5.3/dist/MarkerCluster.Default.css" />
                    <script src="https://unpkg.com/leaflet@1.9.4/dist/leaflet.js"></script>
                    <script src="https://unpkg.com/leaflet.markercluster@1.5.3/dist/leaflet.markercluster.js"></script>
                    <style>
                        body {{ margin: 0; padding: 0; font-family: 'Plus Jakarta Sans', system-ui, -apple-system, sans-serif; background: transparent; }}
                        #map {{ width: 100%; height: 560px; border-radius: 16px; border: 1.5px solid rgba(59,130,246,0.35); box-shadow: 0 12px 32px rgba(0,0,0,0.22); overflow: hidden; }}
                        .tw-wrap {{ position: relative; width: 38px; height: 38px; display: flex; align-items: center; justify-content: center; transition: transform 0.2s ease; }}
                        .tw-wrap:hover {{ transform: scale(1.28) translateY(-3px); z-index: 9999 !important; }}
                        .tw-pulse {{
                            position: absolute; width: 36px; height: 36px; border-radius: 50%;
                            animation: radarPulse 2.1s infinite ease-out; pointer-events: none;
                        }}
                        .tw-pulse2 {{
                            position: absolute; width: 36px; height: 36px; border-radius: 50%;
                            animation: radarPulse 2.1s infinite ease-out 1.05s; pointer-events: none;
                        }}
                        @keyframes radarPulse {{
                            0% {{ transform: scale(0.45); opacity: 0.9; }}
                            100% {{ transform: scale(2.15); opacity: 0; }}
                        }}
                        .tw-core {{
                            position: relative; width: 28px; height: 28px; border-radius: 50%;
                            display: flex; align-items: center; justify-content: center;
                            box-shadow: 0 0 12px rgba(0,0,0,0.65), inset 0 1px 2px rgba(255,255,255,0.45);
                            border: 2px solid #ffffff;
                        }}
                        .tw-lbl {{
                            position: absolute; bottom: -15px; background: rgba(15,23,42,0.88); color: #f8fafc;
                            font-size: 9.5px; font-weight: 800; padding: 1px 5px; border-radius: 4px;
                            white-space: nowrap; border: 1px solid rgba(255,255,255,0.2); pointer-events: none;
                            letter-spacing: 0.3px; box-shadow: 0 2px 4px rgba(0,0,0,0.4);
                        }}
                        .leaflet-popup-content-wrapper {{
                            background: linear-gradient(145deg, #0f172a, #1e293b) !important;
                            color: #f8fafc !important; border-radius: 14px !important;
                            border: 1px solid rgba(255,215,0,0.4); box-shadow: 0 14px 30px rgba(0,0,0,0.5) !important;
                        }}
                        .leaflet-popup-tip {{ background: #1e293b !important; }}
                        .pop-card {{ min-width: 245px; font-size: 12.5px; line-height: 1.55; }}
                        .pop-head {{ font-size: 14.5px; font-weight: 800; color: #38bdf8; margin-bottom: 6px; border-bottom: 1px solid rgba(255,255,255,0.12); padding-bottom: 5px; }}
                        .pop-badge {{ display: inline-block; padding: 2px 8px; border-radius: 99px; font-size: 10.5px; font-weight: 800; color: #fff; margin-bottom: 6px; }}
                        .pop-btn {{
                            display: block; text-align: center; margin-top: 9px; padding: 7px 10px;
                            background: linear-gradient(135deg, #2563eb, #4f46e5); color: #ffffff !important;
                            text-decoration: none; font-weight: 800; border-radius: 8px; font-size: 11.5px;
                        }}
                        .pop-btn:hover {{ filter: brightness(1.15); }}
                        .map-legend {{
                            background: rgba(15, 23, 42, 0.88); backdrop-filter: blur(6px);
                            color: #f8fafc; padding: 10px 13px; border-radius: 12px;
                            border: 1px solid rgba(255,255,255,0.18); font-size: 11.5px; line-height: 1.6;
                            box-shadow: 0 6px 18px rgba(0,0,0,0.35);
                        }}
                        .map-legend i {{ display: inline-block; width: 10px; height: 10px; border-radius: 50%; margin-right: 6px; }}
                    </style>
                </head>
                <body>
                    <div id="map"></div>
                    <script>
                        const sites = {markers_json};
                        const focusLat = {focus_lat};
                        const focusLon = {focus_lon};
                        const focusZoom = {focus_zoom};
                        const focusId = "{focus_id}";
                        const useCluster = {"true" if use_cluster else "false"};

                        const map = L.map('map', {{
                            center: [focusLat, focusLon],
                            zoom: focusZoom,
                            zoomControl: true,
                            preferCanvas: true
                        }});

                        L.tileLayer('{tile_url}', {{
                            subdomains: '{tile_sub}',
                            maxZoom: 19,
                            attribution: '⚡ SIDATA Geospatial Engine'
                        }}).addTo(map);

                        function makeTowerIcon(s) {{
                            const pulseHtml = s.pulse ? `
                                <div class="tw-pulse" style="border: 2.2px solid ${{s.color}}; background: ${{s.color}}28;"></div>
                                <div class="tw-pulse2" style="border: 1.5px solid ${{s.color}};"></div>
                            ` : '';
                            const svgAntenna = s.is_combat ? `
                                <svg width="16" height="16" viewBox="0 0 24 24" fill="none" stroke="#ffffff" stroke-width="2.3" stroke-linecap="round" stroke-linejoin="round">
                                    <path d="M12 2L15 8L21 9L16.5 14L18 20L12 17L6 20L7.5 14L3 9L9 8L12 2Z"/>
                                </svg>
                            ` : `
                                <svg width="16" height="16" viewBox="0 0 24 24" fill="none" stroke="#ffffff" stroke-width="2.2" stroke-linecap="round" stroke-linejoin="round">
                                    <path d="M4.9 19.1C1 15.2 1 8.8 4.9 4.9"/>
                                    <path d="M7.8 16.2c-2.3-2.3-2.3-6.1 0-8.5"/>
                                    <circle cx="12" cy="12" r="2"/>
                                    <path d="M16.2 7.8c2.3 2.3 2.3 6.1 0 8.5"/>
                                    <path d="M19.1 4.9C23 8.8 23 15.1 19.1 19"/>
                                    <line x1="12" y1="14" x2="9" y2="22"/>
                                    <line x1="12" y1="14" x2="15" y2="22"/>
                                    <line x1="9.8" y1="19" x2="14.2" y2="19"/>
                                </svg>
                            `;
                            return L.divIcon({{
                                className: '',
                                iconSize: [38, 38],
                                iconAnchor: [19, 19],
                                popupAnchor: [0, -18],
                                html: `
                                    <div class="tw-wrap">
                                        ${{pulseHtml}}
                                        <div class="tw-core" style="background: ${{s.color}};">
                                            ${{svgAntenna}}
                                        </div>
                                        <div class="tw-lbl">${{s.id}}</div>
                                    </div>
                                `
                            }});
                        }}

                        const group = useCluster ? L.markerClusterGroup({{
                            showCoverageOnHover: false,
                            maxClusterRadius: 42,
                            spiderfyOnMaxZoom: true
                        }}) : L.layerGroup();

                        let targetMarker = null;
                        sites.forEach(s => {{
                            const m = L.marker([s.lat, s.lon], {{ icon: makeTowerIcon(s) }});
                            const popupHtml = `
                                <div class="pop-card">
                                    <div class="pop-head">🗼 ${{s.id}} — ${{s.name}}</div>
                                    <span class="pop-badge" style="background:${{s.color}};">Class: ${{s.cls}}</span>
                                    <span class="pop-badge" style="background:#334155;">${{s.own}}</span>
                                    <div>🏛️ <b>Wilayah:</b> ${{s.kab}}</div>
                                    <div>⚡ <b>Power:</b> ${{s.pwr}}</div>
                                    <div>🔌 <b>Rectifier:</b> ${{s.rect}} (${{s.hub}})</div>
                                    <div>📡 <b>Transport:</b> ${{s.trm}}</div>
                                    <div>👷 <b>PIC RTS:</b> ${{s.rts}}</div>
                                    <div>📋 <b>Status Visit:</b> ${{s.vis}}</div>
                                    <div>🗓️ <b>Last PM:</b> ${{s.ldate}} (${{s.lstat}})</div>
                                    <a class="pop-btn" href="https://www.google.com/maps?q=${{s.lat}},${{s.lon}}" target="_blank">🧭 Navigasi Google Maps</a>
                                </div>
                            `;
                            m.bindPopup(popupHtml);
                            group.addLayer(m);
                            if (focusId && s.id === focusId) {{
                                targetMarker = m;
                            }}
                        }});

                        map.addLayer(group);

                        if (targetMarker) {{
                            setTimeout(() => {{
                                if (useCluster && group.zoomToShowLayer) {{
                                    group.zoomToShowLayer(targetMarker, () => targetMarker.openPopup());
                                }} else {{
                                    targetMarker.openPopup();
                                }}
                            }}, 350);
                        }}

                        const legend = L.control({{ position: 'bottomright' }});
                        legend.onAdd = function() {{
                            const div = L.DomUtil.create('div', 'map-legend');
                            if ({"true" if is_visit_mode else "false"}) {{
                                div.innerHTML = `
                                    <b>📡 Status Kunjungan PM</b><br>
                                    <i style="background:#10b981"></i> Sudah Dikunjungi PM<br>
                                    <i style="background:#ef4444"></i> Belum Dikunjungi (Target)
                                `;
                            }} else {{
                                div.innerHTML = `
                                    <b>🗼 Kelas Tower (Site Class)</b><br>
                                    <i style="background:#a855f7"></i> Diamond &nbsp;
                                    <i style="background:#ef4444"></i> Platinum<br>
                                    <i style="background:#f59e0b"></i> Gold &nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;
                                    <i style="background:#38bdf8"></i> Silver<br>
                                    <i style="background:#10b981"></i> Bronze
                                `;
                            }}
                            return div;
                        }};
                        legend.addTo(map);
                    </script>
                </body>
                </html>
                """
                st.components.v1.html(leaflet_html, height=575)
    with tab_pop:
        st.markdown("#### 📊 Distribusi Populasi Site Berdasarkan Kabupaten & Infrastruktur")
        if df_view.empty:
            st.info("Tidak ada data untuk divisualisasikan.")
        else:
            cp_1, cp_2 = st.columns(2)
            with cp_1:
                with st.container(border=True):
                    st.markdown("**🏛️ Populasi Site per Kabupaten / Kategori Wilayah**")
                    df_kab_cnt = df_view['Kabupaten'].value_counts().to_frame("Jumlah Site")
                    ui_bar_chart(df_kab_cnt, color="#2563EB", horizontal=True)
            with cp_2:
                with st.container(border=True):
                    st.markdown("**🏢 Populasi Berdasarkan Tower Owner (Mitra Tower)**")
                    df_own_cnt = df_view['Site_Owner'].value_counts().head(10).to_frame("Jumlah Site")
                    ui_bar_chart(df_own_cnt, color="#10B981", horizontal=True)

            cp_3, cp_4 = st.columns(2)
            with cp_3:
                with st.container(border=True):
                    st.markdown("**⚡ Distribusi Catuan Daya (Power Source)**")
                    df_pwr_cnt = df_view['Power_Source'].value_counts().head(8).to_frame("Jumlah Site")
                    ui_bar_chart(df_pwr_cnt, color="#F59E0B", horizontal=True)
            with cp_4:
                with st.container(border=True):
                    st.markdown("**⭐ Distribusi Kelas Site (Site Class) & Brand Rectifier**")
                    df_cls_cnt = df_view['site_class'].value_counts().to_frame("Jumlah Site")
                    ui_bar_chart(df_cls_cnt, color="#8B5CF6")
                    df_rect_cnt = df_view['BrandRecti'].value_counts().head(6).to_frame("Jumlah Site")
                    ui_bar_chart(df_rect_cnt, color="#EC4899", horizontal=True)

            st.markdown("##### 📋 Tabel Rekapitulasi Populasi & Progres Kunjungan per Kabupaten")
            rekap_kab = df_view.groupby('Kabupaten').agg(
                Total_Site=('site_id', 'count'),
                Sudah_Dikunjungi=('Jumlah_Visit_PM', lambda x: int((x > 0).sum())),
                Belum_Dikunjungi=('Jumlah_Visit_PM', lambda x: int((x == 0).sum())),
                Site_Platinum_Diamond=('site_class', lambda x: int(x.isin(['Platinum', 'Diamond']).sum())),
                Site_Gold=('site_class', lambda x: int((x == 'Gold').sum())),
                Site_SPS=('RemarkSite', lambda x: int(x.str.contains('SPS', case=False, na=False).sum()))
            ).reset_index()
            rekap_kab['Coverage_PM (%)'] = (rekap_kab['Sudah_Dikunjungi'] / rekap_kab['Total_Site'] * 100).round(1)
            rekap_kab = rekap_kab.sort_values(by='Total_Site', ascending=False).reset_index(drop=True)
            st.dataframe(rekap_kab, use_container_width=True, hide_index=True)

    with tab_comp:
        st.markdown("#### 🔄 Komparasi Otomatis: Data KML Sysinfomap vs Database Existing SIDATA")
        st.info("💡 Modul ini membandingkan seluruh **Site ID di Sysinfomap (KML)** dengan **Master Excel Sidata**, **File Combat NOP Palangkaraya**, serta **Database Laporan Kunjungan Tim (Google Sheets)** sehingga terlihat jelas site mana yang baru, yang sudah terinspeksi, dan yang belum tersentuh.")

        kml_set = set(df_map['site_id'].unique())
        baru_di_kml = sorted(list(kml_set - master_excel_ids))
        sinkron_master = sorted(list(kml_set.intersection(master_excel_ids)))
        hanya_di_master = sorted(list(master_excel_ids - kml_set))

        cc_1, cc_2, cc_3, cc_4 = st.columns(4)
        cc_1.metric("🌐 Total Site di KML", f"{len(kml_set)} Site")
        cc_2.metric("🔗 Sinkron dgn Master Excel", f"{len(sinkron_master)} Site")
        cc_3.metric("🆕 Site Baru (Ada di KML, Belum di Excel)", f"{len(baru_di_kml)} Site")
        cc_4.metric("📂 Hanya Ada di Master/Combat Excel", f"{len(hanya_di_master)} Site")

        sub_c1, sub_c2, sub_c3 = st.tabs([
            f"✅ Site Sudah Dikunjungi PM ({tot_visited})",
            f"⏳ Daftar Antrean Site Belum Dikunjungi ({tot_unvisited})",
            f"🆕 Audit Perbedaan Master Excel vs KML ({len(baru_di_kml)})"
        ])

        with sub_c1:
            df_done_pm = df_map[df_map['Jumlah_Visit_PM'] > 0][['site_id', 'Sitename', 'Kabupaten', 'site_class', 'RTS', 'Jumlah_Visit_PM', 'Tgl_Visit_Terakhir', 'Status_PM_Terakhir', 'Teknisi_Terakhir']].sort_values(by='Jumlah_Visit_PM', ascending=False)
            if df_done_pm.empty:
                st.warning("Belum ada site dari KML yang tercatat di laporan kunjungan PM.")
            else:
                st.dataframe(df_done_pm, use_container_width=True, hide_index=True)

        with sub_c2:
            df_unv = df_map[df_map['Jumlah_Visit_PM'] == 0][['site_id', 'Sitename', 'Kabupaten', 'site_class', 'Site_Owner', 'Power_Source', 'RTS', 'Access']]
            st.dataframe(df_unv, use_container_width=True, hide_index=True)
            st.download_button(
                "📥 Download Daftar Site Belum Dikunjungi (.csv)",
                data=df_unv.to_csv(index=False).encode('utf-8'),
                file_name=f"Unvisited_Sites_Sysinfomap_{datetime.date.today()}.csv",
                mime="text/csv",
                key="dl_unv_sim"
            )

        with sub_c3:
            col_au1, col_au2 = st.columns(2)
            with col_au1:
                st.markdown(f"**🆕 Site Ada di KML Tapi Belum Masuk Excel Master ({len(baru_di_kml)} Site):**")
                if baru_di_kml:
                    df_new_kml = df_map[df_map['site_id'].isin(baru_di_kml)][['site_id', 'Sitename', 'Kabupaten', 'site_class', 'Status_site', 'RTS']]
                    st.dataframe(df_new_kml, use_container_width=True, hide_index=True)
                else:
                    st.success("✅ Seluruh Site ID KML sudah terdaftar di Master Excel!")
            with col_au2:
                st.markdown(f"**📂 Site Ada di Excel Master/Combat Tapi Tidak Ada di KML ({len(hanya_di_master)} Site):**")
                if hanya_di_master:
                    st.dataframe(pd.DataFrame({"Site ID (Hanya di Master/Combat)": hanya_di_master}), use_container_width=True, hide_index=True)
                else:
                    st.success("✅ Tidak ada selisih Site ID!")

    with tab_detail:
        st.markdown("#### 📋 Explorer Tabel Data Sysinfomap Lengkap")
        cols_tampil = ['site_id', 'Sitename', 'Kabupaten', 'site_class', 'Site_Owner', 'Status_site', 'Power_Source', 'BrandRecti', 'Transport_type', 'Type_Hub', 'Access', 'RTS', 'Status_Kunjungan', 'Tgl_Visit_Terakhir', 'lat', 'lon']
        st.dataframe(df_view[cols_tampil], use_container_width=True, hide_index=True, height=480)
        st.download_button(
            "📥 Download Data Sysinfomap Terfilter (.csv)",
            data=df_view[cols_tampil].to_csv(index=False).encode('utf-8'),
            file_name=f"Database_Sysinfomap_{datetime.date.today()}.csv",
            mime="text/csv",
            type="primary",
            use_container_width=True,
            key="dl_full_sim"
        )
# =========================================================================
# MENU 10: SiSAVA (SISTEM SERVICE VARIABLE ACTIVITY)
# =========================================================================
elif menu == "💠 SiSAVA":
    if 'sv_user' not in st.session_state:
        st.session_state['sv_user'] = None
    
    if st.session_state['sv_user'] is None:
        store = sv_store()
        ux_hero("AKUN SiSAVA", "Portal SiSAVA", "Sistem Pengajuan RAB, Nilai SVA, Approval Berjenjang, dan Bukti Pekerjaan Lapangan.")
        
        locked = st.session_state.get("sv_lock_until", 0) - _dtm.datetime.now().timestamp()
        
        st.markdown("<h3 style='text-align:center;'>Silakan Login ke Portal SiSAVA</h3>", unsafe_allow_html=True)
        c = st.columns([1.5, 2, 1.5])[1]
        with c:
            with st.container(border=True):
                u = st.text_input("Username SiSAVA", key="svw_lu")
                p = st.text_input("Kata Sandi", type="password", key="svw_lp")
                if locked > 0: 
                    st.error(f"Terlalu banyak percobaan gagal. Coba lagi dalam {int(locked)} detik.")
                
                st.write("")
                if st.button("🔓 Masuk Portal SiSAVA", type="primary", use_container_width=True, disabled=locked > 0):
                    usr = sv_login(store, u, p)
                    if usr:
                        st.session_state["sv_user"] = usr
                        st.session_state["sv_fail"] = 0
                        st.session_state["sv_nav"] = "🏠 Dashboard"
                        st.rerun()
                    else:
                        st.session_state["sv_fail"] = st.session_state.get("sv_fail", 0) + 1
                        if st.session_state["sv_fail"] >= 5: 
                            st.session_state["sv_lock_until"] = _dtm.datetime.now().timestamp() + 60
                            st.session_state["sv_fail"] = 0
                        st.error("Username atau sandi salah, atau akun nonaktif.")
    else:
        sisava_page()
# FOOTER HAK CIPTA OKTA PRADIKA
# -------------------------------------------------------------------------
st.markdown("""
    <div class="footer-okta">
        System Application & Database Management<br>
        <span>Create Data By Okta Pradika</span> © 2026
    </div>
""", unsafe_allow_html=True)
