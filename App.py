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

UI_THEME = "light"  # "light" (disarankan: teks selalu terbaca) | "dark" | "auto" (ikut mode gelap/terang perangkat)

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
html,body,[data-testid="stAppViewContainer"],[data-testid="stMain"],.main{background:var(--ux-bg) !important}
.stApp{color:var(--ux-ink);--text-color:var(--ux-ink);--secondary-background-color:var(--ux-card);--background-color:var(--ux-bg);--primary-color:var(--ux-accent)}

/* --- KUNCI WARNA TEKS: terbaca di tema Streamlit terang maupun gelap, di laptop mode apa pun ---
   (elemen yang sudah punya warna inline, mis. kartu gelap Executive, sengaja tidak disentuh) */
.stApp p:not([style*="color"]),.stApp li:not([style*="color"]),.stApp h1:not([style*="color"]),.stApp h2:not([style*="color"]),.stApp h3:not([style*="color"]),
.stApp h4:not([style*="color"]),.stApp h5:not([style*="color"]),.stApp h6:not([style*="color"]),.stApp label,.stApp [data-testid="stText"],.stApp td,.stApp th,
div[data-testid="stWidgetLabel"] p,div[data-testid="stWidgetLabel"] label,div[data-testid="stWidgetLabel"] span,
div[data-testid="stCheckbox"] label span,div[data-testid="stCheckbox"] p,div[data-testid="stToggle"] p,
div[role="radiogroup"] label p,div[role="radiogroup"] label span,
div[data-testid="stExpander"] summary,div[data-testid="stExpander"] summary p,div[data-testid="stExpander"] summary span,
div[data-testid="stAlert"] p,div[data-testid="stAlert"] li,div[data-testid="stAlert"] span,
div[data-testid="stMetricValue"],div[data-testid="stMetricValue"] div,
section[data-testid="stFileUploaderDropzone"] span,section[data-testid="stFileUploaderDropzone"] small,section[data-testid="stFileUploaderDropzone"] p{color:var(--ux-ink) !important}
div[data-testid="stCaptionContainer"],div[data-testid="stCaptionContainer"] p,.stApp small{color:var(--ux-muted) !important}
.stTabs [data-baseweb="tab"]{color:var(--ux-muted) !important}
.stTabs [data-baseweb="tab"] p{color:inherit !important}
.stTabs [data-baseweb="tab"][aria-selected="true"]{color:var(--ux-accent) !important}
div[data-baseweb="input"] input,div[data-baseweb="base-input"] input,div[data-baseweb="textarea"] textarea,.stTextInput input,.stNumberInput input,.stDateInput input,.stTextArea textarea{
  color:var(--ux-ink) !important;-webkit-text-fill-color:var(--ux-ink) !important;caret-color:var(--ux-ink) !important;background:transparent !important}
input:disabled,textarea:disabled{opacity:1 !important;-webkit-text-fill-color:var(--ux-ink) !important;background:var(--ux-line) !important}
input::placeholder,textarea::placeholder{color:var(--ux-muted) !important;-webkit-text-fill-color:var(--ux-muted) !important;opacity:1 !important}
div[data-baseweb="select"] span,div[data-baseweb="select"] div,div[data-baseweb="select"] input{color:var(--ux-ink) !important;-webkit-text-fill-color:var(--ux-ink) !important}
div[data-baseweb="select"] svg,div[data-testid="stNumberInput"] button svg{fill:var(--ux-ink) !important;color:var(--ux-ink) !important}
div[data-testid="stNumberInput"] button{background:var(--ux-card2) !important;color:var(--ux-ink) !important}
span[data-baseweb="tag"]{background:rgba(37,99,235,.12) !important;border:1px solid rgba(37,99,235,.35) !important}
span[data-baseweb="tag"] span{color:var(--ux-accent) !important;-webkit-text-fill-color:var(--ux-accent) !important;font-weight:700}
a[data-testid^="stBaseLinkButton"],a[data-testid="stLinkButton"]{background:var(--ux-card) !important;color:var(--ux-ink) !important;border:1.5px solid var(--ux-line) !important}
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
.stApp div[data-testid="stButton"] button p,.stApp div[data-testid="stFormSubmitButton"] button p,.stApp div[data-testid="stDownloadButton"] button p,
.stApp section[data-testid="stFileUploaderDropzone"] button p,.stApp section[data-testid="stFileUploaderDropzone"] button span{color:inherit !important}
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
    if subtitle: h += '<div class="ux-hero-sub">%s</div>' % _esc(subtitle)
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
        TTD_OKTA_B64 = "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"
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

def sv_rekap_xlsx(recs):
    rows = [{"No. RAB": r.get("rab_id"), "Tiket": r.get("ticket_no"), "NOP": r.get("nop"), "Tgl Pengajuan": r.get("tgl_pengajuan"),
             "Pembuat": r.get("nama_pembuat"), "Status": r.get("status"), "Total ke Telkomsel": _sv_num(r.get("tot_tsel")),
             "Total Revenue KUT": _sv_num(r.get("tot_rev")), "Total HPP": _sv_num(r.get("tot_hpp")), "Margin": _sv_num(r.get("margin")),
             "GPM": _sv_num(r.get("gpm")), "Margin Bersih": _sv_num(r.get("net")), "SPV": r.get("spv_nama"), "Manager": r.get("mgr_nama")} for r in recs]
    buf = _io.BytesIO()
    with pd.ExcelWriter(buf, engine="openpyxl") as xw:
        pd.DataFrame(rows).to_excel(xw, sheet_name="Rekap SiSAVA", index=False)
    return buf.getvalue()

# ----------------------------------------------------------------------------
# PENYIMPANAN: Google Sheets (utama) / file JSON lokal (cadangan)
# ----------------------------------------------------------------------------
def _sv_json(o):
    return _json.dumps(o, ensure_ascii=False, separators=(",", ":"), default=str)

def _sv_jload(s, default=None):
    try:
        return _json.loads(s) if s else default
    except Exception:
        return default

class _SvLocalStore:
    kind = "lokal"
    def __init__(self, path=None):
        self.path = path or SV_LOCAL_DB
    def _read(self):
        try:
            with open(self.path, "r", encoding="utf-8") as f:
                return _json.load(f)
        except Exception:
            return {"rab": [], "users": []}
    def _write(self, d):
        tmp = self.path + ".tmp"
        with open(tmp, "w", encoding="utf-8") as f:
            _json.dump(d, f, ensure_ascii=False)
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
    try:
        ws = sh.worksheet(name)
    except Exception:
        ws = sh.add_worksheet(title=name, rows=500, cols=max(len(cols), 10) + 2)
        ws.update(range_name="A1", values=[cols])
        return ws
    head = ws.row_values(1)
    if head != cols[:len(head)] or len(head) < len(cols):      # skema bertambah -> lengkapi header di kanan
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
        for k in ("payload", "evidence_rab", "log", "sva"):
            if len(str(rec.get(k, ""))) > 45000: raise ValueError("Data '%s' terlalu besar untuk satu sel Google Sheets." % k)
        ws = _sv_worksheet("SISAVA_RAB", self._cols("SISAVA_RAB"))
        head, vals = self._row_values(ws, rec, SV_RAB_COLS)
        ws.append_row(vals, value_input_option="RAW"); _sv_rows.clear()
    def update_rab(self, rab_id, fields):
        ws = _sv_worksheet("SISAVA_RAB", self._cols("SISAVA_RAB"))
        try:
            cell = ws.find(rab_id, in_column=1)
        except Exception:
            cell = None
        if not cell: return False
        head = ws.row_values(1); cur = ws.row_values(cell.row)
        cur += [""] * (len(head) - len(cur))
        for k, v in fields.items():
            if k in head:
                if len(str(v)) > 45000: raise ValueError("Data '%s' terlalu besar untuk satu sel Google Sheets." % k)
                cur[head.index(k)] = "" if v is None else str(v)
        ws.update(range_name="A%d" % cell.row, values=[cur], value_input_option="RAW"); _sv_rows.clear()
        return True
    def upsert_user(self, u):
        ws = _sv_worksheet("SISAVA_USERS", self._cols("SISAVA_USERS"))
        try:
            cell = ws.find(u["username"], in_column=1)
        except Exception:
            cell = None
        head = ws.row_values(1)
        if cell:                                   # pembaruan sebagian: gabungkan dengan isi lama, jangan menimpa kolom lain
            cur = ws.row_values(cell.row); cur += [""] * (len(head) - len(cur))
            row = dict(zip(head, cur)); row.update({k: v for k, v in u.items() if k in head})
            ws.update(range_name="A%d" % cell.row, values=[[str(row.get(c, "")) for c in head]], value_input_option="RAW")
        else:
            ws.append_row([str(u.get(c, "")) for c in head], value_input_option="RAW")
        _sv_rows.clear()

def sv_store():
    """Google Sheets bila tersambung, kalau tidak -> file lokal (tidak permanen di Streamlit Cloud)."""
    try:
        _sv_spreadsheet()
        return _SvSheetsStore()
    except Exception:
        _sv_spreadsheet.clear()
        return _SvLocalStore()

# ----------------------------------------------------------------------------
# AKUN & SANDI (PBKDF2-SHA256 + garam; sandi tidak pernah disimpan polos)
# ----------------------------------------------------------------------------
def sv_hash_pw(pw, salt=None):
    salt = salt or _pysecrets.token_hex(8)
    dk = _hl.pbkdf2_hmac("sha256", pw.encode("utf-8"), salt.encode(), 120000)
    return "pbkdf2$%s$%s" % (salt, dk.hex())

def sv_check_pw(pw, stored):
    try:
        _, salt, hx = str(stored).split("$")
        return _hmac.compare_digest(sv_hash_pw(pw, salt).split("$")[2], hx)
    except Exception:
        return False

def sv_can(user, perm):
    return bool(user) and perm in SV_PERMS.get(user.get("role"), set())

def sv_user_find(store, username):
    u = str(username).strip().lower()
    for r in store.list_users():
        if str(r.get("username", "")).strip().lower() == u:
            return r
    return None

def sv_login(store, username, pw):
    r = sv_user_find(store, username)
    if not r or str(r.get("aktif", "1")).strip() in ("0", "False", "false", "nonaktif") or not sv_check_pw(pw, r.get("pw_hash", "")):
        return None
    return {"username": r["username"], "nama": r.get("nama", r["username"]), "jabatan": r.get("jabatan", ""), "role": r.get("role", "STAFF")}

def sv_validate_new_user(u, pw, existing=None):
    if not str(u.get("username", "")).strip() or not str(u.get("nama", "")).strip():
        return "Username dan nama wajib diisi."
    if any(ch.isspace() for ch in u["username"]):
        return "Username tidak boleh mengandung spasi."
    if u.get("role") not in SV_ROLES:
        return "Role tidak valid."
    if len(pw) < 8:
        return "Sandi minimal 8 karakter."
    return ""

# ----------------------------------------------------------------------------
# ALUR APPROVAL (aturan ada di sini, bukan di tampilan -> tidak bisa dilewati lewat tombol)
# ----------------------------------------------------------------------------
def sv_log_add(rec, user, aksi, catatan=""):
    log = _sv_jload(rec.get("log"), []) or []
    log.append({"ts": sv_now(), "oleh": user["username"], "nama": user.get("nama", ""), "role": user["role"], "aksi": aksi, "catatan": catatan})
    return log

def sv_next_id(recs, year=None):
    year = year or _dtm.date.today().year
    nums = []
    for r in recs:
        parts = str(r.get("rab_id", "")).split("-")
        if len(parts) == 3 and parts[1] == str(year) and parts[2].isdigit():
            nums.append(int(parts[2]))
    return "RAB-%d-%04d" % (year, (max(nums) if nums else 0) + 1)

def sv_transition(rec, action, user, note="", data=None):
    """Kembalikan (ok, pesan, fields). Tidak menulis ke database; pemanggil yang menyimpan."""
    st_, role, uname = rec.get("status"), user.get("role"), user.get("username")
    creator = rec.get("dibuat_oleh")
    def deny(m): return False, m, {}
    need_note = action in ("spv_reject", "mgr_reject", "f_spv_reject", "f_mgr_reject")
    if need_note and not str(note).strip():
        return deny("Catatan wajib diisi saat menolak.")
    if action in ("spv_approve", "spv_reject", "f_spv_approve", "f_spv_reject") and not sv_can(user, "approve_spv"):
        return deny("Akun Anda tidak berwenang sebagai SPV.")
    if action in ("mgr_approve", "mgr_reject", "f_mgr_approve", "f_mgr_reject") and not sv_can(user, "approve_mgr"):
        return deny("Hanya Manager yang dapat melakukan persetujuan ini.")
    if action in ("spv_approve", "spv_reject", "mgr_approve", "mgr_reject", "f_spv_approve", "f_spv_reject", "f_mgr_approve", "f_mgr_reject") \
            and uname == creator and not SV_ALLOW_SELF_APPROVE:
        return deny("Anda tidak dapat menyetujui/menolak dokumen yang Anda ajukan sendiri.")
    f = {"updated_at": sv_now()}
    ok_msg = ""
    if action == "spv_approve" and st_ == S_RAB_SPV:
        f.update(status=S_RAB_MGR, spv_oleh=uname, spv_nama=user["nama"], spv_jabatan=user.get("jabatan", ""), spv_at=sv_now(), spv_catatan=note); ok_msg = "RAB diteruskan ke Manager."
    elif action == "spv_reject" and st_ == S_RAB_SPV:
        f.update(status=S_RAB_TOLAK, spv_oleh=uname, spv_nama=user["nama"], spv_at=sv_now(), spv_catatan=note); ok_msg = "RAB dikembalikan untuk revisi."
    elif action == "mgr_approve" and st_ in (S_RAB_MGR, S_RAB_SPV):
        if st_ == S_RAB_SPV:
            f.update(spv_oleh=uname, spv_nama=user["nama"], spv_jabatan=user.get("jabatan", ""), spv_at=sv_now(), spv_catatan="(disetujui langsung oleh Manager) " + note)
        f.update(status=S_RAB_OK, mgr_oleh=uname, mgr_nama=user["nama"], mgr_jabatan=user.get("jabatan", ""), mgr_at=sv_now(), mgr_catatan=note); ok_msg = "RAB DISETUJUI. Silakan isi Nilai SVA."
    elif action == "mgr_reject" and st_ in (S_RAB_MGR, S_RAB_SPV):
        f.update(status=S_RAB_TOLAK, mgr_oleh=uname, mgr_nama=user["nama"], mgr_at=sv_now(), mgr_catatan=note); ok_msg = "RAB dikembalikan untuk revisi."
    elif action == "sva_fill" and st_ == S_RAB_OK:
        if not sv_can(user, "sva_fill"): return deny("Akun Anda tidak berwenang mengisi Nilai SVA.")
        d = data or {}
        miss = [k for k in ("ticket_no", "status_tiket", "urgensi", "keperluan", "tgl_rencana") if not str(d.get(k, "")).strip()]
        if miss: return deny("Lengkapi: " + ", ".join(miss))
        if _sv_num(d.get("nilai_sva")) <= 0 or _sv_num(d.get("keperluan_dana")) <= 0:
            return deny("Nilai SVA dan Keperluan Dana harus lebih dari 0.")
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
        if st_ == S_KRJ_SPV:
            f.update(f_spv_oleh=uname, f_spv_at=sv_now(), f_spv_catatan="(disetujui langsung oleh Manager) " + note)
        f.update(status=S_DONE, f_mgr_oleh=uname, f_mgr_at=sv_now(), f_mgr_catatan=note); ok_msg = "Pekerjaan SAH."
    elif action == "f_mgr_reject" and st_ in (S_KRJ_MGR, S_KRJ_SPV):
        f.update(status=S_KRJ_TOLAK, f_mgr_oleh=uname, f_mgr_at=sv_now(), f_mgr_catatan=note); ok_msg = "Pekerjaan dikembalikan untuk diperbaiki."
    else:
        return deny("Aksi tidak sesuai dengan status dokumen saat ini (%s)." % st_)
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


# ----------------------------------------------------------------------------
# TAMPILAN SiSAVA — bagian bersama
# ----------------------------------------------------------------------------
def _sv_badge(status):
    return '<span style="display:inline-block;padding:3px 11px;border-radius:99px;font-size:.74rem;font-weight:800;color:#fff;background:%s">%s</span>' % (
        SV_STATUS_COLOR.get(status, "#64748B"), _esc(status))

def _sv_flash():
    m = st.session_state.pop("sv_flash", None)
    if m:
        (st.success if m[0] != "!" else st.error)(m.lstrip("!"))

@st.cache_data(ttl=3600, show_spinner=False, max_entries=80)
def _sv_fetch_img(url):
    r = requests.get(url, timeout=10)
    r.raise_for_status()
    return r.content

@st.cache_data(ttl=300, show_spinner=False, max_entries=20)
def _sv_xlsx_cached(rab_id, stamp, payload, meta_json, ev_json):
    rab = _sv_jload(payload, {}) or {}
    meta = _sv_jload(meta_json, {}) or {}
    imgs = []
    for u in (_sv_jload(ev_json, []) or [])[:6]:
        try:
            imgs.append(_sv_fetch_img(u))
        except Exception:
            pass
    return sv_build_xlsx(rab, sv_calc(rab), meta, imgs)

def sv_record_xlsx(rec):
    meta = {"rab_id": rec.get("rab_id"), "status": rec.get("status"), "pembuat_nama": rec.get("nama_pembuat"), "pembuat_jabatan": rec.get("jabatan_pembuat"),
            "spv_nama": rec.get("spv_nama"), "spv_jabatan": rec.get("spv_jabatan"), "mgr_nama": rec.get("mgr_nama"), "mgr_jabatan": rec.get("mgr_jabatan")}
    return _sv_xlsx_cached(rec.get("rab_id"), str(rec.get("updated_at")), rec.get("payload", ""), _sv_json(meta), rec.get("evidence_rab", ""))

def _sv_autodownload(data, name):
    """Picu unduhan otomatis lewat iframe; tombol manual tetap disediakan sebagai cadangan."""
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
    for k in [k for k in list(ss.keys()) if str(k).startswith("svw_") or str(k).startswith("svl_")]:
        ss.pop(k, None)
    ss["sv_form"] = {"nop": SV_NOPS[0], "ticket": "", "area": "KALIMANTAN", "kota": "Balikpapan", "tanggal": _dtm.date.today().isoformat(),
                     "params": dict(SV_PARAMS_DEFAULT), "sites": [_sv_new_site()], "evidence": [], "catatan": "", "edit_id": None, "old_urls": [], "ver": 0}

def _sv_form_load(rec):
    """Muat RAB yang ditolak ke form untuk direvisi."""
    _sv_form_reset()
    rab = _sv_jload(rec.get("payload"), {}) or {}
    f = st.session_state["sv_form"]
    f.update(nop=rab.get("nop", f["nop"]), ticket=rab.get("ticket", ""), area=rab.get("area", f["area"]), kota=rab.get("kota", f["kota"]),
             tanggal=rab.get("tanggal", f["tanggal"]), edit_id=rec["rab_id"], old_urls=_sv_jload(rec.get("evidence_rab"), []) or [])
    f["params"] = sv_params(rab.get("params"))
    f["sites"] = []
    for s in rab.get("sites", []):
        s = dict(s); s["uid"] = _pysecrets.token_hex(3); s.setdefault("items", [])
        for ln in s["items"]:
            ln.setdefault("sec", "A"); ln.setdefault("harga_tsel", None); ln.setdefault("hpp", None)
        f["sites"].append(s)
    f["sites"] = f["sites"] or [_sv_new_site()]
    st.session_state["sv_nav"] = "➕ Pengajuan RAB"
    st.session_state["sv_flash"] = "Data %s dimuat ke form. Perbaiki lalu ajukan ulang." % rec["rab_id"]

def _sv_rab_from_form(f):
    return {"nop": f["nop"], "ticket": f["ticket"].strip(), "area": f["area"], "kota": f["kota"], "tanggal": f["tanggal"],
            "params": dict(f["params"]), "sites": [{k: s[k] for k in ("site_id", "site_name", "pekerjaan", "items")} for s in f["sites"]]}

# ----------------------------------------------------------------------------
# FORM PENGAJUAN RAB
# ----------------------------------------------------------------------------
def _sv_add_a(uid):
    ss = st.session_state; f = ss["sv_form"]
    s = next((x for x in f["sites"] if x["uid"] == uid), None)
    item = (ss.get("svl_%s" % uid) or {}).get(ss.get("svw_pick_%s" % uid))
    if not s or not item:
        return
    price = item["price"].get(f["area"])
    s["items"].append({"sec": "A", "jenis": item["jenis"], "desc": item["desc"], "uom": item["uom"], "qty": _sv_num(ss.get("svw_qty_%s" % uid, 1)) or 1.0,
                       "harga_kut": float(price or 0.0), "harga_tsel": None, "hpp": None, "ref": item["uid"]})
    f["ver"] += 1
    if price is None:
        ss["sv_flash"] = "!Item ini tidak punya harga tetap (%s). Isi harga di tabel." % (item["ket"] or "harga manual")

def _sv_add_b(uid):
    ss = st.session_state; f = ss["sv_form"]
    s = next((x for x in f["sites"] if x["uid"] == uid), None)
    desc = str(ss.get("svw_bd_%s" % uid, "")).strip()
    if not s or not desc:
        ss["sv_flash"] = "!Isi uraian Other Material terlebih dahulu."; return
    hpp = _sv_num(ss.get("svw_bh_%s" % uid, 0))
    s["items"].append({"sec": "B", "jenis": ss.get("svw_bj_%s" % uid, "Material"), "desc": desc, "uom": str(ss.get("svw_bu_%s" % uid, "")).strip(),
                       "qty": _sv_num(ss.get("svw_bq_%s" % uid, 1)) or 1.0, "harga_kut": _sv_num(ss.get("svw_bk_%s" % uid, 0)),
                       "harga_tsel": None, "hpp": hpp if hpp > 0 else None})
    f["ver"] += 1

def _sv_del_site(uid):
    f = st.session_state["sv_form"]
    if len(f["sites"]) > 1:
        f["sites"] = [s for s in f["sites"] if s["uid"] != uid]; f["ver"] += 1

def _sv_reset_auto(uid):
    f = st.session_state["sv_form"]
    for s in f["sites"]:
        if s["uid"] == uid:
            for ln in s["items"]:
                ln["harga_tsel"] = None; ln["hpp"] = None
    f["ver"] += 1

def _sv_line_editor(f, s, sec, p):
    items, uid = s["items"], s["uid"]
    idxs = [i for i, ln in enumerate(items) if ln.get("sec", "A") == sec]
    if not idxs:
        st.caption("Belum ada item pada bagian ini."); return
    rows = []
    for i in idxs:
        ln = items[i]; c = sv_line_calc(ln, p)
        rows.append({"Jenis": ln["jenis"], "Uraian": ln["desc"], "UoM": ln.get("uom", ""), "Qty": float(c["qty"]), "Harga KUT→3E": round(c["harga_kut"], 2),
                     "Harga 3E→Tsel": round(c["harga_tsel"], 2), "HPP/unit": round(c["hpp_unit"], 2), "Hapus": False})
    df = pd.DataFrame(rows)
    ed = st.data_editor(df, key="svw_ed_%s_%s_%d" % (uid, sec, f["ver"]), hide_index=True, use_container_width=True, num_rows="fixed",
                        disabled=["Uraian", "UoM"],
                        column_config={"Jenis": st.column_config.SelectboxColumn("Jenis", options=["Material", "Jasa"], required=True, width="small"),
                                       "Uraian": st.column_config.TextColumn("Uraian", width="large"),
                                       "Qty": st.column_config.NumberColumn("Qty", min_value=0.0, step=1.0, format="%.2f", width="small"),
                                       "Harga KUT→3E": st.column_config.NumberColumn("Harga KUT→3E", min_value=0.0, format="%.2f"),
                                       "Harga 3E→Tsel": st.column_config.NumberColumn("Harga 3E→Tsel", min_value=0.0, format="%.2f"),
                                       "HPP/unit": st.column_config.NumberColumn("HPP/unit", min_value=0.0, format="%.2f"),
                                       "Hapus": st.column_config.CheckboxColumn("Hapus", width="small")})
    drop = []
    for k, i in enumerate(idxs):
        er, orig, ln = ed.iloc[k], rows[k], items[i]
        if bool(er["Hapus"]):
            drop.append(i); continue
        ln["jenis"] = er["Jenis"]; ln["qty"] = _sv_num(er["Qty"]); ln["harga_kut"] = _sv_num(er["Harga KUT→3E"])
        if abs(_sv_num(er["Harga 3E→Tsel"]) - orig["Harga 3E→Tsel"]) > 0.005: ln["harga_tsel"] = _sv_num(er["Harga 3E→Tsel"])
        if abs(_sv_num(er["HPP/unit"]) - orig["HPP/unit"]) > 0.005: ln["hpp"] = _sv_num(er["HPP/unit"])
    if drop:
        s["items"] = [ln for j, ln in enumerate(items) if j not in drop]; f["ver"] += 1; st.rerun()

def _sv_picker_a(f, s, pl):
    uid, ss = s["uid"], st.session_state
    if not pl:
        st.info("Daftar Shopping List belum terbaca. Gunakan Other Material (bagian B), atau letakkan file Excel Shopping List di folder aplikasi."); return
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
        c4, c5 = st.columns(2)
        c4.number_input("Harga KUT→3E (per unit)", min_value=0.0, value=0.0, step=1000.0, key="svw_bk_%s" % uid)
        c5.number_input("HPP per unit (0 = otomatis)", min_value=0.0, value=0.0, step=1000.0, key="svw_bh_%s" % uid)
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
    if f.get("edit_id"):
        st.info("✏️ Mode REVISI untuk %s. Setelah diajukan, RAB kembali ke SPV." % f["edit_id"])
    if not pl_path:
        st.warning("File Shopping List tidak ditemukan. Letakkan file Excel *Summary Final Price Shopping List* di folder yang sama dengan App.py.")
    with st.expander("⚙️ Parameter perhitungan (asumsi dari contoh RAB, boleh diubah)", expanded=False):
        p = f["params"]
        a, b, c = st.columns(3)
        p["f_mat"] = a.number_input("Faktor material (harga KUT→3E = x × harga 3E→Tsel)", 0.10, 2.0, float(p["f_mat"]), 0.01, format="%.3f", key="svw_pfm")
        p["f_jasa"] = b.number_input("Faktor jasa (harga 3E→Tsel = harga KUT→3E × x)", 0.10, 3.0, float(p["f_jasa"]), 0.001, format="%.3f", key="svw_pfj")
        p["hpp_pct"] = c.number_input("HPP default (% dari harga KUT→3E)", 0.0, 200.0, float(p["hpp_pct"]) * 100, 0.5, key="svw_php") / 100.0
        d, e, g = st.columns(3)
        p["fee_3e"] = d.number_input("Fee 3E→Telkomsel (%)", 0.0, 100.0, float(p["fee_3e"]) * 100, 0.5, key="svw_pf3") / 100.0
        p["fee_kut"] = e.number_input("Fee KUT→3E (%)", 0.0, 100.0, float(p["fee_kut"]) * 100, 0.5, key="svw_pfk") / 100.0
        p["biaya_lain"] = g.number_input("Biaya lain-lain (Rp)", 0.0, 1e12, float(p["biaya_lain"]), 500.0, format="%.0f", key="svw_pbl")
        h, i, j = st.columns(3)
        p["com_rate"] = h.number_input("Cost of Money (% per tahun)", 0.0, 100.0, float(p["com_rate"]) * 100, 0.5, key="svw_pcr") / 100.0
        p["com_bulan"] = i.number_input("Cost of Money: jumlah bulan", 0.0, 36.0, float(p["com_bulan"]), 1.0, key="svw_pcb")
        p["pph"] = j.number_input("PPh (% dari margin)", 0.0, 100.0, float(p["pph"]) * 100, 0.05, format="%.2f", key="svw_ppp") / 100.0
        st.caption("Nilai awal berasal dari contoh RAB (material: 0,92; jasa: 1,177; HPP 85%; Cost of Money 7,5% × 3/12; PPh 2,85%). "
                   "Harga 3E→Telkomsel dan HPP per baris juga bisa diubah langsung di tabel item.")

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
            cc1.button("↺ Kembalikan harga 3E→Tsel & HPP ke otomatis", key="svw_rst_%s" % uid, on_click=_sv_reset_auto, args=(uid,), use_container_width=True)
            if len(f["sites"]) > 1:
                cc2.button("🗑️ Hapus site ini", key="svw_dels_%s" % uid, on_click=_sv_del_site, args=(uid,), use_container_width=True)
    if st.button("➕ Tambah Site", key="svw_adds", use_container_width=True):
        f["sites"].append(_sv_new_site()); f["ver"] += 1; st.rerun()

def _sv_rab_problems(rab, calc, f, n_files):
    errs = []
    if not str(rab["ticket"]).strip(): errs.append("No. Ticket SVA wajib diisi.")
    if not rab["sites"] or any(not (s["site_id"] and s["site_name"]) for s in rab["sites"]): errs.append("Setiap site wajib punya Site ID dan Nama Site.")
    if any(not s["pekerjaan"].strip() for s in rab["sites"]): errs.append("Isi kolom Pekerjaan pada setiap site.")
    if any(not s["items"] for s in rab["sites"]): errs.append("Setiap site harus memiliki minimal 1 item.")
    if any(_sv_num(ln.get("qty")) <= 0 or _sv_num(ln.get("harga_kut")) <= 0 for s in rab["sites"] for ln in s["items"]): errs.append("Qty dan Harga KUT→3E semua item harus lebih dari 0.")
    if calc["other_rev"] > 0 and (n_files + len(f.get("old_urls") or [])) == 0: errs.append("Ada Other Material: lampirkan evidence referensi harga (minimal 1 gambar).")
    return errs

def _sv_step_hitung(f, calc):
    st.markdown("##### 3️⃣ Hitungan Revenue & Margin")
    c1, c2, c3 = st.columns(3)
    c1.metric("Total ditagihkan 3E → Telkomsel", sv_rp(calc["tot_tsel"]))
    c2.metric("Total Revenue KUT → 3E", sv_rp(calc["tot_rev"]))
    c3.metric("Total HPP (+ biaya lain)", sv_rp(calc["tot_hpp"]))
    d1, d2, d3, d4 = st.columns(4)
    d1.metric("Margin", sv_rp(calc["margin"]))
    d2.metric("GPM", "%.1f%%" % (calc["gpm"] * 100))
    d3.metric("Cost of Money + PPh", sv_rp(calc["com"] + calc["pph"]))
    d4.metric("Margin bersih", sv_rp(calc["net"]))
    if calc["tot_rev"] > 0:
        ux_progress([("HPP", calc["tot_hpp_items"], "#2563EB"), ("Biaya lain", calc["biaya_lain"], "#94A3B8"), ("Margin", max(calc["margin"], 0), "#10B981")])
    if calc["margin"] < 0:
        st.error("⚠️ Margin NEGATIF. Periksa harga KUT→3E, HPP, atau biaya lain sebelum mengajukan.")
    elif calc["tot_rev"] > 0 and calc["gpm"] < 0.05:
        st.warning("GPM di bawah 5%. Pastikan harga dan HPP sudah benar.")
    rows = [{"Site": "%s %s" % (b["site_id"], b["site_name"]), "Ditagih ke Tsel": sv_rp(b["sum"]["A"]["tot_tsel"] + b["sum"]["B"]["tot_tsel"]),
             "Revenue KUT": sv_rp(b["sum"]["A"]["rev"] + b["sum"]["B"]["rev"]), "HPP": sv_rp(b["sum"]["A"]["hpp"] + b["sum"]["B"]["hpp"])} for b in calc["sites"]]
    if rows:
        st.dataframe(pd.DataFrame(rows), use_container_width=True, hide_index=True)
    st.caption("Margin = Total Revenue KUT→3E − (Total HPP + biaya lain). GPM = Margin ÷ Revenue. Margin bersih = Margin − Cost of Money − PPh.")

def _sv_step_ajukan(store, user, f, rab, calc):
    st.markdown("##### 4️⃣ Lampiran & Ajukan")
    files = st.file_uploader("Evidence referensi harga (gambar) — wajib bila ada Other Material", type=["png", "jpg", "jpeg", "webp"],
                             accept_multiple_files=True, key="svw_files")
    if files:
        f["evidence"] = [(x.name, x.getvalue()) for x in files]
    if f["evidence"]:
        cc1, cc2 = st.columns([3, 1])
        cc1.caption("%d gambar siap dilampirkan (tetap tersimpan walau Anda berpindah menu)." % len(f["evidence"]))
        if cc2.button("Hapus lampiran", key="svw_clr_ev", use_container_width=True):
            f["evidence"] = []; st.session_state.pop("svw_files", None); st.rerun()
    if f.get("old_urls"):
        st.caption("%d lampiran lama dari pengajuan sebelumnya tetap dipakai." % len(f["old_urls"]))
    f["catatan"] = st.text_area("Catatan untuk approver (opsional)", value=f["catatan"], key="svw_cat", height=70)
    errs = _sv_rab_problems(rab, calc, f, len(f["evidence"]))
    for e in errs:
        st.warning(e)
    ok = st.checkbox("Saya menyatakan data RAB ini benar.", key="svw_ok")
    label = "🚀 Ajukan Ulang RAB & Unduh Excel" if f.get("edit_id") else "🚀 Ajukan RAB & Unduh Excel"
    if st.button(label, type="primary", use_container_width=True, disabled=bool(errs) or not ok, key="svw_submit"):
        _sv_submit(store, user, f, rab, calc)

def _sv_submit(store, user, f, rab, calc):
    with st.spinner("Menyimpan RAB & membuat Excel..."):
        urls = []
        for name, raw in f["evidence"]:
            class _F:                                   # pembungkus kecil supaya cocok dengan upload_image()
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
        if st.session_state.pop("sv_autodl", False):
            _sv_autodownload(last["data"], last["name"])
        c1, c2 = st.columns([3, 1])
        c1.download_button("📥 Unduh Excel RAB (jika unduhan otomatis diblokir browser)", last["data"], file_name=last["name"],
                           mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet", use_container_width=True, type="primary", key="svw_dl_last")
        if c2.button("Tutup", use_container_width=True, key="svw_close_last"):
            st.session_state.pop("sv_last", None); st.rerun()
    if "sv_form" not in st.session_state:
        _sv_form_reset()
    f = st.session_state["sv_form"]
    t1, t2, t3, t4 = st.tabs(["1️⃣ Info", "2️⃣ Site & Item", "3️⃣ Hitungan", "4️⃣ Ajukan"])
    with t1: _sv_step_info(f, pl_path)
    with t2: _sv_step_sites(f, pl)
    rab = _sv_rab_from_form(f); calc = sv_calc(rab)
    with t3: _sv_step_hitung(f, calc)
    with t4: _sv_step_ajukan(store, user, f, rab, calc)
    if f.get("edit_id") or any(s["items"] for s in f["sites"]) or f["ticket"]:
        if st.button("🧹 Kosongkan form", key="svw_clear"):
            _sv_form_reset(); st.rerun()


# ----------------------------------------------------------------------------
# AKSI, DETAIL, DAN HALAMAN LAIN
# ----------------------------------------------------------------------------
def _sv_act(rab_id, action, note_key=None, data=None):
    ss = st.session_state; store = sv_store(); user = ss.get("sv_user")
    rec = next((r for r in store.list_rabs() if r.get("rab_id") == rab_id), None)
    if not rec or not user:
        ss["sv_flash"] = "!Data tidak ditemukan / sesi berakhir."; return
    ok, msg, fields = sv_transition(rec, action, user, ss.get(note_key, "") if note_key else "", data() if callable(data) else data)
    if not ok:
        ss["sv_flash"] = "!" + msg; return
    try:
        store.update_rab(rab_id, fields); ss["sv_flash"] = "%s (%s)" % (msg, rab_id)
        if note_key: ss[note_key] = ""
        if action == "sva_fill": ss.pop("sv_sva_open", None)
    except Exception as e:
        ss["sv_flash"] = "!Gagal menyimpan: %s" % e

_SV_AKSI = {"diajukan": "Diajukan", "revisi_diajukan": "Diajukan ulang (revisi)", "spv_approve": "Disetujui SPV", "spv_reject": "Ditolak SPV",
            "mgr_approve": "Disetujui Manager", "mgr_reject": "Ditolak Manager", "sva_fill": "Nilai SVA diisi", "kerja_upload": "Bukti pekerjaan diunggah",
            "f_spv_approve": "Pekerjaan disetujui SPV", "f_spv_reject": "Pekerjaan ditolak SPV", "f_mgr_approve": "Pekerjaan disetujui Manager (SAH)",
            "f_mgr_reject": "Pekerjaan ditolak Manager"}

def _sv_timeline(rec):
    log = _sv_jload(rec.get("log"), []) or []
    if not log:
        st.caption("Belum ada riwayat."); return
    h = "".join('<div style="display:flex;gap:10px;padding:7px 0;border-bottom:1px solid var(--ux-line)"><div style="min-width:132px;font-size:.78rem;color:var(--ux-muted)">%s</div>'
                '<div style="font-size:.88rem"><b>%s</b> — %s%s</div></div>' % (_esc(x.get("ts", "")), _esc(_SV_AKSI.get(x.get("aksi"), x.get("aksi"))),
                _esc(x.get("nama") or x.get("oleh")), (" · <i>“%s”</i>" % _esc(x["catatan"])) if x.get("catatan") else "") for x in log)
    st.markdown(h, unsafe_allow_html=True)

def _sv_summary(rec):
    c1, c2, c3 = st.columns(3)
    c1.metric("Ditagih ke Telkomsel", sv_rp(rec.get("tot_tsel")))
    c2.metric("Revenue KUT→3E", sv_rp(rec.get("tot_rev")))
    c3.metric("HPP", sv_rp(rec.get("tot_hpp")))
    d1, d2, d3 = st.columns(3)
    d1.metric("Margin", sv_rp(rec.get("margin")))
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
            try:
                ss[k] = (str(rec.get("updated_at")), sv_record_xlsx(rec)); st.rerun()
            except Exception as e:
                st.error("Gagal membuat Excel: %s" % e)

def _sv_evidence_view(rec):
    ev = _sv_jload(rec.get("evidence_rab"), []) or []
    if ev:
        tampilkan_grid_foto(ev, "Referensi harga")
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
    ux_hero(SV_ROLES.get(user["role"], ""), "SiSAVA Dashboard", "Pantau seluruh siklus RAB sampai pekerjaan SAH.",
            stats=[(n, "Total RAB"), (wait, "Menunggu approval"), (run, "Berjalan"), (cnt[S_DONE], "SAH")],
            split=[("Menunggu approval", wait, "#FBBF24"), ("Berjalan", run, "#5B9BFF"), ("SAH", cnt[S_DONE], "#34D399"), ("Ditolak", tolak, "#FB7185")])
    aksi = {"STAFF": cnt[S_RAB_OK] + cnt[S_SVA] + tolak, "SPV": cnt[S_RAB_SPV] + cnt[S_KRJ_SPV], "MANAGER": cnt[S_RAB_SPV] + cnt[S_RAB_MGR] + cnt[S_KRJ_SPV] + cnt[S_KRJ_MGR]}[user["role"]]
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
    else:
        st.info("Belum ada RAB. Mulai dari menu ➕ Pengajuan RAB.")

def sv_page_list(user, recs):
    if not recs:
        st.info("Belum ada RAB."); return
    c1, c2 = st.columns([2, 1])
    q = c1.text_input("Cari (No. RAB / tiket / site / pembuat)", key="svw_lq")
    stt = c2.multiselect("Status", SV_PIPELINE + [S_RAB_TOLAK, S_KRJ_TOLAK], key="svw_ls")
    rows = [r for r in recs if (not stt or r.get("status") in stt) and (not q or q.lower() in (str(r.get("rab_id")) + str(r.get("ticket_no")) + str(r.get("payload"))[:3000] + str(r.get("nama_pembuat"))).lower())]
    rows = sorted(rows, key=lambda r: str(r.get("tgl_pengajuan")), reverse=True)
    st.dataframe(pd.DataFrame([{"No. RAB": r.get("rab_id"), "Tiket": r.get("ticket_no"), "NOP": r.get("nop"), "Pembuat": r.get("nama_pembuat"), "Status": r.get("status"),
                                "Revenue": sv_rp(r.get("tot_rev")), "Margin": sv_rp(r.get("margin")), "GPM": "%.1f%%" % (_sv_num(r.get("gpm")) * 100)} for r in rows]),
                 use_container_width=True, hide_index=True)
    if not rows:
        return
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
        if a_ and c_:
            st.markdown("**Catatan %s (%s):** %s" % (lab, _esc(a_), _esc(c_)))
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
    if own:
        st.caption("%d dokumen yang Anda ajukan sendiri tidak muncul di sini (harus disetujui orang lain)." % len(own))
    if not q:
        st.success("Tidak ada dokumen yang menunggu persetujuan Anda."); return
    for r in sorted(q, key=lambda x: str(x.get("tgl_pengajuan"))):
        rid, s_ = r["rab_id"], r["status"]
        with st.expander("%s · %s · %s" % (rid, r.get("ticket_no"), s_), expanded=False):
            _sv_rec_header(r); _sv_summary(r)
            _sv_xlsx_button(r, "ap_" + rid)
            _sv_sva_view(r); _sv_evidence_view(r)
            st.markdown("**Riwayat**"); _sv_timeline(r)
            nk = "svw_note_%s_%s" % (rid, s_)
            st.text_area("Catatan (wajib bila menolak)", key=nk, height=70)
            if s_ in (S_RAB_SPV, S_RAB_MGR):
                ya, tidak = ("spv_approve", "spv_reject") if (s_ == S_RAB_SPV and user["role"] == "SPV") else ("mgr_approve", "mgr_reject")
            else:
                ya, tidak = ("f_spv_approve", "f_spv_reject") if (s_ == S_KRJ_SPV and user["role"] == "SPV") else ("f_mgr_approve", "f_mgr_reject")
            lab = "✅ Setujui" + (" langsung (sebagai Manager)" if (user["role"] == "MANAGER" and s_ in (S_RAB_SPV, S_KRJ_SPV)) else "")
            b1, b2 = st.columns(2)
            b1.button(lab, key="svw_ok_%s_%s" % (rid, s_), type="primary", use_container_width=True, on_click=_sv_act, args=(rid, ya, nk))
            b2.button("↩️ Tolak / minta revisi", key="svw_no_%s_%s" % (rid, s_), use_container_width=True, on_click=_sv_act, args=(rid, tidak, nk))

def sv_page_sva(user, recs):
    todo = [r for r in recs if r.get("status") == S_RAB_OK]
    st.markdown("##### Validasi RAB & pengisian Nilai SVA (%d RAB siap)" % len(todo))
    if not todo:
        st.info("Belum ada RAB yang disetujui dan menunggu pengisian Nilai SVA."); 
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
    if not todo:
        st.info("Tidak ada pekerjaan yang menunggu unggahan bukti."); return
    pick = st.selectbox("Pilih RAB / pekerjaan", [r["rab_id"] for r in todo], format_func=lambda i: "%s · %s" % (i, next(x for x in todo if x["rab_id"] == i).get("ticket_no")), key="svw_kpick")
    r = next(x for x in todo if x["rab_id"] == pick)
    _sv_rec_header(r); _sv_sva_view(r)
    if r.get("status") == S_KRJ_TOLAK:
        st.warning("Pekerjaan ditolak. Catatan SPV: %s · Catatan Manager: %s" % (r.get("f_spv_catatan") or "-", r.get("f_mgr_catatan") or "-"))
    if not sv_can(user, "kerja_upload"):
        st.info("Akun Anda tidak berwenang mengunggah pekerjaan."); return
    fs = st.file_uploader("Foto / berkas bukti pekerjaan (sebelum, sesudah, BA, dll.)", type=["png", "jpg", "jpeg", "webp", "pdf"], accept_multiple_files=True, key="svw_kf_" + pick)
    cat = st.text_area("Catatan pekerjaan", key="svw_kc_" + pick, height=80)
    if st.button("📤 Kirim Bukti Pekerjaan", type="primary", use_container_width=True, key="svw_ksend_" + pick, disabled=not fs):
        with st.spinner("Mengunggah bukti..."):
            urls = []
            for fobj in fs:
                u = upload_datalog(fobj, "sisava_pekerjaan") if fobj.name.lower().endswith(".pdf") else upload_image(fobj, "sisava_pekerjaan")
                if u: urls.append(u)
        if len(urls) < len(fs):
            st.error("Sebagian berkas gagal diunggah (%d dari %d berhasil). Coba lagi." % (len(urls), len(fs)))
        else:
            ss = st.session_state; ss["svw_kc_tmp"] = cat
            _sv_act(pick, "kerja_upload", "svw_kc_tmp", {"urls": urls}); st.rerun()

def sv_page_kelola(store, user):
    t1, t2, t3 = st.tabs(["👥 Akun", "💲 Shopping List", "📊 Rekap"])
    with t1:
        users = store.list_users()
        st.dataframe(pd.DataFrame([{"Username": u.get("username"), "Nama": u.get("nama"), "Jabatan": u.get("jabatan"), "Role": SV_ROLES.get(u.get("role"), u.get("role")),
                                    "Aktif": "Ya" if str(u.get("aktif", "1")) not in ("0", "False", "false") else "Tidak"} for u in users]), use_container_width=True, hide_index=True)
        with st.expander("➕ Tambah akun", expanded=False):
            a, b = st.columns(2)
            nu = a.text_input("Username", key="svw_nu"); nn = b.text_input("Nama lengkap", key="svw_nn")
            c, d = st.columns(2)
            nj = c.text_input("Jabatan (tampil di tanda tangan Excel)", key="svw_nj", placeholder="mis. Spv Support Reg. Kalimantan")
            nr = d.selectbox("Role", list(SV_ROLES), format_func=lambda x: SV_ROLES[x], key="svw_nr")
            npw = st.text_input("Sandi awal (min. 8 karakter)", type="password", key="svw_npw")
            if st.button("Simpan akun baru", key="svw_nsave", type="primary"):
                msg = sv_validate_new_user({"username": nu.strip(), "nama": nn, "role": nr}, npw)
                if not msg and sv_user_find(store, nu): msg = "Username sudah dipakai."
                if msg: st.error(msg)
                else:
                    store.upsert_user({"username": nu.strip(), "nama": nn.strip(), "jabatan": nj.strip(), "role": nr, "pw_hash": sv_hash_pw(npw), "aktif": "1", "dibuat": sv_now()})
                    st.success("Akun %s dibuat." % nu.strip()); st.rerun()
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
                    if aktif and (un == user["username"] or (cur.get("role") == "MANAGER" and len(managers) <= 1)):
                        st.error("Tidak dapat menonaktifkan akun Anda sendiri / Manager aktif terakhir.")
                    else:
                        store.upsert_user({"username": un, "aktif": "0" if aktif else "1"}); st.rerun()
    with t2:
        pl, path = sv_pricelist()
        if not pl:
            st.warning("File Shopping List belum ditemukan di folder aplikasi."); 
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
        if recs:
            st.download_button("📥 Unduh Rekap Seluruh RAB (Excel)", sv_rekap_xlsx(recs), file_name="Rekap_SiSAVA.xlsx",
                               mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet", use_container_width=True, key="svw_rekap")

# ----------------------------------------------------------------------------
# LOGIN SiSAVA & HALAMAN UTAMA
# ----------------------------------------------------------------------------
def _sv_logout():
    st.session_state.pop("sv_user", None)

def _sv_login_view(store):
    ss = st.session_state
    ux_hero("AKUN SiSAVA", "SiSAVA", "Sistem Service Variable Activity — RAB, Approval, Nilai SVA, dan Bukti Pekerjaan.")
    users = store.list_users()
    if not users:
        st.warning("Belum ada akun SiSAVA.")
        if ss.get("role") != "Admin":
            st.info("Setup akun pertama hanya dapat dilakukan oleh Admin SiDATA. Silakan keluar, lalu masuk sebagai Admin."); return
        st.markdown("##### Buat akun Manager pertama")
        a, b = st.columns(2)
        u = a.text_input("Username", key="svw_su"); n = b.text_input("Nama lengkap", key="svw_sn")
        j = st.text_input("Jabatan", key="svw_sj", placeholder="Mgr. Business Operation Reg. Kalimantan")
        p1 = st.text_input("Sandi (min. 8 karakter)", type="password", key="svw_sp")
        if st.button("Buat akun Manager", type="primary", key="svw_sgo"):
            msg = sv_validate_new_user({"username": u.strip(), "nama": n, "role": "MANAGER"}, p1)
            if msg: st.error(msg)
            else:
                store.upsert_user({"username": u.strip(), "nama": n.strip(), "jabatan": j.strip(), "role": "MANAGER", "pw_hash": sv_hash_pw(p1), "aktif": "1", "dibuat": sv_now()})
                st.success("Akun Manager dibuat. Silakan login."); st.rerun()
        return
    locked = ss.get("sv_lock_until", 0) - _dtm.datetime.now().timestamp()
    c = st.columns([1, 2, 1])[1]
    with c:
        u = st.text_input("Username", key="svw_lu"); p = st.text_input("Sandi", type="password", key="svw_lp")
        if locked > 0:
            st.error("Terlalu banyak percobaan gagal. Coba lagi dalam %d detik." % int(locked))
        if st.button("🔓 Masuk SiSAVA", type="primary", use_container_width=True, key="svw_lgo", disabled=locked > 0):
            usr = sv_login(store, u, p)
            if usr:
                ss["sv_user"] = usr; ss["sv_fail"] = 0; ss["sv_nav"] = "🏠 Dashboard"; st.rerun()
            else:
                ss["sv_fail"] = ss.get("sv_fail", 0) + 1
                if ss["sv_fail"] >= 5:
                    ss["sv_lock_until"] = _dtm.datetime.now().timestamp() + 60; ss["sv_fail"] = 0
                st.error("Username atau sandi salah, atau akun nonaktif.")
        if ss.get("role") == "Admin":
            with st.expander("Lupa sandi? (khusus Admin SiDATA)"):
                ru = st.selectbox("Akun", [x["username"] for x in users], key="svw_ru"); rp = st.text_input("Sandi baru (min. 8 karakter)", type="password", key="svw_rp")
                if st.button("Reset sandi", key="svw_rgo"):
                    if len(rp) < 8: st.error("Sandi minimal 8 karakter.")
                    else: store.upsert_user({"username": ru, "pw_hash": sv_hash_pw(rp)}); st.success("Sandi %s direset." % ru)

def sisava_page():
    ss = st.session_state
    store = sv_store()
    _sv_flash()
    if store.kind == "lokal":
        st.warning("⚠️ Google Sheets belum tersambung (st.secrets 'gcp_json'). Data disimpan sementara di file lokal dan bisa hilang saat aplikasi restart.")
    user = ss.get("sv_user")
    if not user:
        _sv_login_view(store); return
    c1, c2 = st.columns([4, 1])
    c1.markdown("<div style='padding:8px 2px'>👤 <b>%s</b> · %s%s</div>" % (_esc(user["nama"]), _esc(SV_ROLES.get(user["role"], "")), (" · " + _esc(user["jabatan"])) if user.get("jabatan") else ""), unsafe_allow_html=True)
    c2.button("Keluar SiSAVA", key="svw_logout", on_click=_sv_logout, use_container_width=True)
    nav = ["🏠 Dashboard", "➕ Pengajuan RAB", "📋 Daftar RAB"]
    if sv_can(user, "approve_spv") or sv_can(user, "approve_mgr"): nav.append("✅ Approval")
    nav += ["🧾 Nilai SVA", "📤 Upload Pekerjaan"]
    if sv_can(user, "manage_users"): nav.append("🛠️ Kelola")
    if ss.get("sv_nav") not in nav: ss["sv_nav"] = nav[0]
    page = st.radio("Menu SiSAVA", nav, horizontal=True, key="sv_nav", label_visibility="collapsed")
    try:
        recs = sv_visible(store.list_rabs(), user)
    except Exception as e:
        st.error("Gagal membaca database SiSAVA: %s" % e); return
    if page == "🏠 Dashboard": sv_page_dashboard(user, recs)
    elif page == "➕ Pengajuan RAB":
        pl, path = sv_pricelist(); sv_page_form(store, user, pl, path)
    elif page == "📋 Daftar RAB": sv_page_list(user, recs)
    elif page == "✅ Approval": sv_page_approval(user, recs)
    elif page == "🧾 Nilai SVA": sv_page_sva(user, recs)
    elif page == "📤 Upload Pekerjaan": sv_page_kerja(user, recs)
    elif page == "🛠️ Kelola": sv_page_kelola(store, user)
