import datetime
import pandas as pd
import streamlit as st
import json
import os
import urllib.parse
import gspread
from google.oauth2.service_account import Credentials
import requests
from io import BytesIO
import base64
import tempfile
import re
import zipfile
import time
import importlib.util as _imp_util

# Cek ketersediaan PPTX tanpa memaksa import berat saat halaman login dibuka
HAS_PPTX = _imp_util.find_spec("pptx") is not None

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
    from pptx import Presentation
    from pptx.util import Inches, Pt
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
    from fpdf import FPDF
    from PIL import Image
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
# SISTEM LOGIN & ROLE AKSES (3 KARTU UTAMA: LAPANGAN, OFFICE, ADMIN)
# -------------------------------------------------------------------------
if 'role' not in st.session_state: st.session_state['role'] = None
if 'dashboard_area' not in st.session_state: st.session_state['dashboard_area'] = None
if 'active_menu' not in st.session_state: st.session_state['active_menu'] = None

OFFICE_ALLOWED_MENUS = [
    "🏢 Portal Office",
    "📦 Logistik & SPMS",
    "🚨 SiHD (Pergerakan Tim)",
    "💠 SiSAVA",
    "💎 Executive Project Dashboard",
    "📋 Program NOP & Task Force",
    "🛡️ SiDAVID (Verifikasi Internal)",
    "🗺️ Sysinfomap (Peta & Populasi Site)",
    "🚧 Form & Peta Hard Akses",
    "📅 Daily Tracker Palangkaraya",
    "👥 Kelola Master Tim (PLK & PBU)"
]

if st.session_state['role'] is None and st.session_state.get('active_menu') not in OFFICE_ALLOWED_MENUS:
    ux_login_css()
    st.markdown("<br>", unsafe_allow_html=True)
    render_header_logo(is_login=True)
    st.markdown("<h1 style='text-align: center; font-size: 40px; margin-top: 5px;'>Sistem Operasional Terpadu (SIDATA)</h1>", unsafe_allow_html=True)
    st.markdown("<p style='text-align: center; font-size: 18px; margin-bottom: 35px; letter-spacing: 1px; opacity: 0.85;'>Pilih Portal Akses Pekerjaan Anda Hari Ini</p>", unsafe_allow_html=True)

    col_v, col_l, col_a = st.columns([3, 3, 3])

    with col_v:
        st.markdown("""
            <div class='login-box'>
                <div style='font-size: 45px; margin-bottom: 10px;'>👷‍♂️</div>
                <div class='login-title' style='font-size: 20px;'>Mode Lapangan</div>
                <div class='login-desc' style='font-size: 13px;'>Akses Form Pendataan PM (SPS, Reguler & Audit Tower, Pangkalan Bun) untuk Teknisi (FME).</div>
            </div>
        """, unsafe_allow_html=True)
        c_v1, c_v2 = st.columns(2)
        if c_v1.button("🔴 P'BUN", type="primary", use_container_width=True, key="btn_login_pbu"):
            st.session_state['role'] = 'Lapangan'
            st.session_state['area_lapangan'] = 'Pangkalan Bun'
            st.session_state['active_menu'] = None
            st.rerun()
        if c_v2.button("🔵 PLK", type="secondary", use_container_width=True, key="btn_login_plk"):
            st.session_state['role'] = 'Lapangan'
            st.session_state['area_lapangan'] = 'Palangkaraya'
            st.session_state['active_menu'] = None
            st.rerun()

    with col_l:
        st.markdown("""
            <div class='login-box' style='border-color: #10B981;'>
                <div style='font-size: 45px; margin-bottom: 10px;'>🏢</div>
                <div class='login-title' style='font-size: 20px;'>Portal Office</div>
                <div class='login-desc' style='font-size: 13px;'>Sub-Sistem: Logistik, Mutasi Tim (SiHD), SiSAVA, Executive Project, Program NOP & SiDAVID.</div>
            </div>
        """, unsafe_allow_html=True)
        st.write("")
        if st.button("🏢 MASUK PORTAL OFFICE", type="secondary", use_container_width=True, key="btn_login_office"):
            st.session_state['active_menu'] = "🏢 Portal Office"
            st.rerun()

    with col_a:
        st.markdown("""
            <div class='login-box'>
                <div style='font-size: 45px; margin-bottom: 10px;'>🔐</div>
                <div class='login-title' style='font-size: 20px;'>Mode Admin Utama</div>
                <div class='login-desc' style='font-size: 13px;'>Akses Revisi, Dapot Power, Sysinfomap, Manajemen Akun SiSAVA & Seluruh Portal Office.</div>
            </div>
        """, unsafe_allow_html=True)
        admin_pass = st.text_input("Kata Sandi Admin:", type="password", placeholder="Sandi...", key="pwd_login_main", label_visibility="collapsed")
        if st.button("🔓 LOGIN ADMIN", type="primary", use_container_width=True, key="btn_login_admin"):
            if admin_pass == "KUT2027":
                st.session_state['role'] = 'Admin'
                st.session_state['active_menu'] = None
                st.rerun()
            else:
                st.error("❌ Sandi salah!")

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

MASTER_TIM_ROLES = [
    "FME / TE",
    "MBP",
    "CME",
    "Leader Cluster",
    "Koordinator",
    "Admin Operational",
    "Admin NOP",
    "Helpdesk (SiHD)",
    "Tim Project / SVA",
    "Helper / Driver",
    "Lainnya"
]

COLS_MASTER_TIM_PLK = ["NAMA", "ROLE", "CLUSTER", "STATUS_AKTIF", "NO_HP", "LAT_HB", "LON_HB", "UPDATED_AT"]
COLS_MASTER_TIM_PBU = ["NAMA", "NOP", "ROLE", "STATUS_AKTIF", "NO_HP", "Latitude Berangkat", "Longtitude Berangkat", "UPDATED_AT"]

def _seed_local_tim_df(nop_key):
    import glob as _gm
    now_s = datetime.datetime.now().strftime("%Y-%m-%d %H:%M")
    if nop_key == "PLK":
        rows = []
        for pf in ["NIK NOP PLK.xlsx"] + sorted(_gm.glob("*NIK NOP PLK*.xlsx")):
            if os.path.exists(pf):
                try:
                    df1 = read_excel_cached(pf)
                    if "NAMA" in df1.columns:
                        c_role = next((c for c in df1.columns if "role" in str(c).lower() or "jabatan" in str(c).lower()), None)
                        c_cl = next((c for c in df1.columns if "cluster" in str(c).lower() or "area" in str(c).lower()), None)
                        for _, rw in df1.iterrows():
                            nm = str(rw.get("NAMA", "")).strip()
                            if nm and nm.lower() not in ("nan", "none", ""):
                                rows.append({
                                    "NAMA": nm,
                                    "ROLE": str(rw.get(c_role, "FME / TE")).strip() if c_role and pd.notna(rw.get(c_role)) else "FME / TE",
                                    "CLUSTER": str(rw.get(c_cl, "Palangkaraya")).strip() if c_cl and pd.notna(rw.get(c_cl)) else "Palangkaraya",
                                    "STATUS_AKTIF": "Aktif",
                                    "NO_HP": "",
                                    "LAT_HB": "",
                                    "LON_HB": "",
                                    "UPDATED_AT": now_s
                                })
                        break
                except Exception:
                    pass
        return pd.DataFrame(rows, columns=COLS_MASTER_TIM_PLK)
    else:
        rows = []
        for pf in ["NAMA NOP PANGKALANBUN (2).xlsx", "NAMA NOP PANGKALANBUN.xlsx"] + sorted(_gm.glob("*NAMA NOP PANGKALANBUN*.xlsx")):
            if os.path.exists(pf):
                try:
                    df2 = read_excel_cached(pf)
                    if "NAMA" in df2.columns:
                        c_lat = next((c for c in df2.columns if "lat" in str(c).lower()), None)
                        c_lon = next((c for c in df2.columns if "lon" in str(c).lower()), None)
                        c_role = next((c for c in df2.columns if "role" in str(c).lower() or "jabatan" in str(c).lower()), None)
                        for _, rw in df2.iterrows():
                            nm = str(rw.get("NAMA", "")).strip()
                            if nm and nm.lower() not in ("nan", "none", ""):
                                rows.append({
                                    "NAMA": nm,
                                    "NOP": str(rw.get("NOP", "NOP PANGKALAN BUN")).strip() if pd.notna(rw.get("NOP")) else "NOP PANGKALAN BUN",
                                    "ROLE": str(rw.get(c_role, "FME / TE")).strip() if c_role and pd.notna(rw.get(c_role)) else "FME / TE",
                                    "STATUS_AKTIF": "Aktif",
                                    "NO_HP": "",
                                    "Latitude Berangkat": str(rw.get(c_lat, "")).strip() if c_lat and pd.notna(rw.get(c_lat)) else "",
                                    "Longtitude Berangkat": str(rw.get(c_lon, "")).strip() if c_lon and pd.notna(rw.get(c_lon)) else "",
                                    "UPDATED_AT": now_s
                                })
                        break
                except Exception:
                    pass
        return pd.DataFrame(rows, columns=COLS_MASTER_TIM_PBU)

def load_master_tim_df(nop_key="PLK"):
    """Membaca Master Tim (PLK -> 'Master_Tim_PLK', PBU -> 'Master_HB_Tim_PBU') dengan fallback Excel lokal."""
    sh_name = "Master_Tim_PLK" if nop_key == "PLK" else "Master_HB_Tim_PBU"
    cols = COLS_MASTER_TIM_PLK if nop_key == "PLK" else COLS_MASTER_TIM_PBU
    df_c, ok_c = gsheet_read_df(sh_name, cols)
    src_lbl = f"Google Sheets ({sh_name})"
    if not df_c.empty:
        # Pastikan baris kosong tersaring & kolom standar lengkap
        df_c = df_c[df_c["NAMA"].astype(str).str.strip().str.lower().isin(["", "nan", "none"]) == False].copy()
        for col_std in cols:
            if col_std not in df_c.columns:
                df_c[col_std] = "Aktif" if col_std == "STATUS_AKTIF" else ("FME / TE" if col_std == "ROLE" else "")
        df_c.loc[df_c["STATUS_AKTIF"].astype(str).str.strip() == "", "STATUS_AKTIF"] = "Aktif"
        df_c.loc[df_c["ROLE"].astype(str).str.strip() == "", "ROLE"] = "FME / TE"
        return df_c[cols].reset_index(drop=True), src_lbl, ok_c
    df_loc = _seed_local_tim_df(nop_key)
    return df_loc, "Excel Lokal (Belum Disimpan ke Cloud)", ok_c

def load_master_tim_names(nop_key="PLK", include_inactive=False):
    df_t, _, _ = load_master_tim_df(nop_key)
    if df_t.empty:
        return []
    if not include_inactive and "STATUS_AKTIF" in df_t.columns:
        df_t = df_t[~df_t["STATUS_AKTIF"].astype(str).str.strip().str.lower().isin(["nonaktif", "resign", "keluar", "tidak aktif", "0", "false"])]
    return sorted(list({str(x).strip() for x in df_t["NAMA"].dropna().tolist() if str(x).strip()}))

@st.cache_data(ttl=3600, show_spinner=False)
def load_site_master_plk():
    """Memuat daftar Site ID + Koordinat Master NOP Palangkaraya (di-cache 1 jam)."""
    site_set = set()
    coords_plk = {}
    try:
        if os.path.exists("Site ID Sidata Palangkaraya.xlsx"):
            df3 = read_excel_cached("Site ID Sidata Palangkaraya.xlsx")
            if 'Site ID' in df3.columns:
                c_lat = next((c for c in df3.columns if "lat" in str(c).lower()), None)
                c_lon = next((c for c in df3.columns if "lon" in str(c).lower()), None)
                c_nm = next((c for c in df3.columns if "name" in str(c).lower() or "nama" in str(c).lower()), None)
                for _, rw in df3.iterrows():
                    s_raw = str(rw.get('Site ID', '')).strip()
                    if s_raw and s_raw.lower() not in ("nan", "none"):
                        site_set.add(s_raw)
                        sid_u = _norm_site_id(s_raw)
                        if c_lat and c_lon:
                            lv = pd.to_numeric(rw.get(c_lat), errors="coerce")
                            ov = pd.to_numeric(rw.get(c_lon), errors="coerce")
                            if pd.notna(lv) and pd.notna(ov):
                                coords_plk[sid_u] = (float(lv), float(ov), str(rw.get(c_nm, sid_u)).strip() if c_nm else sid_u)
    except Exception:
        pass

    try:
        if os.path.exists("Combat NOP Palangkaraya_1.xlsx"):
            df_cb = read_excel_cached("Combat NOP Palangkaraya_1.xlsx", sheet_name="Combat NOP Palangkaraya")
            if 'siteid' in df_cb.columns:
                for val in df_cb['siteid'].dropna().astype(str).str.strip().tolist():
                    if val:
                        site_set.add(val)
    except Exception:
        pass

    try:
        df_kml_m, _ = get_default_kml_df()
        if not df_kml_m.empty and 'site_id' in df_kml_m.columns:
            for _, r in df_kml_m.iterrows():
                sid_k = str(r.get('site_id', '')).strip().upper()
                sn_k = str(r.get('Sitename', sid_k)).strip()
                if sid_k:
                    site_set.add(f"{sid_k} - {sn_k}")
                    lat_k = pd.to_numeric(r.get('lat'), errors="coerce")
                    lon_k = pd.to_numeric(r.get('lon'), errors="coerce")
                    if pd.notna(lat_k) and pd.notna(lon_k):
                        coords_plk[sid_k] = (float(lat_k), float(lon_k), sn_k)
    except Exception:
        pass

    return sorted(list(site_set)), coords_plk

@st.cache_data(ttl=3600, show_spinner=False)
def load_site_master_pbu():
    """Memuat daftar Site ID + Koordinat Master NOP Pangkalan Bun (di-cache 1 jam)."""
    import glob as _gm
    site_set = set()
    coords_pbu = {}
    COLS_SITE_PBU = ["Site ID", "Latitude Tujuan", "Longtitude Tujuan"]
    df_s, _ = gsheet_read_df("Master_Site_PBU", COLS_SITE_PBU)
    if df_s.empty:
        for p_spbu in ["Site ID Sidata Pangkalanbun (1).xlsx", "Site ID Sidata Pangkalanbun.xlsx"] + sorted(_gm.glob("*Site ID Sidata Pangkalanbun*.xlsx")):
            if os.path.exists(p_spbu):
                try:
                    df_s = read_excel_cached(p_spbu)
                    break
                except Exception:
                    pass
    if not df_s.empty:
        c_sid = next((c for c in df_s.columns if "site" in str(c).lower()), df_s.columns[0])
        c_lat = next((c for c in df_s.columns if "lat" in str(c).lower()), None)
        c_lon = next((c for c in df_s.columns if "lon" in str(c).lower()), None)
        c_nm = next((c for c in df_s.columns if "name" in str(c).lower() or "nama" in str(c).lower()), None)
        for _, rw in df_s.iterrows():
            sid = str(rw.get(c_sid, '')).strip().upper()
            if sid and sid not in ("NAN", "NONE", ""):
                site_set.add(sid)
                if c_lat and c_lon:
                    lv = pd.to_numeric(rw.get(c_lat), errors="coerce")
                    ov = pd.to_numeric(rw.get(c_lon), errors="coerce")
                    if pd.notna(lv) and pd.notna(ov):
                        nm_s = str(rw.get(c_nm, sid)).strip() if (c_nm and pd.notna(rw.get(c_nm))) else sid
                        coords_pbu[sid] = (float(lv), float(ov), nm_s)
    return sorted(list(site_set)), coords_pbu

def load_master_lists():
    """Backward-compatible wrapper yang langsung mengambil Master Tim dinamis dari Cloud (PLK & PBU) + Master Site ter-cache."""
    fme_plk = load_master_tim_names("PLK")
    fme_pbu = load_master_tim_names("PBU")
    site_plk, _ = load_site_master_plk()
    site_pbu, _ = load_site_master_pbu()
    return fme_plk, fme_pbu, site_plk, site_pbu
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
            import cloudinary
            import cloudinary.uploader
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
            import cloudinary
            import cloudinary.uploader
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

_GS_LAST_OK = {}
_SV_LAST_OK = {}

def _is_quota_429(err):
    msg = str(err or "")
    return ("429" in msg) or ("Quota exceeded" in msg) or ("RATE_LIMIT" in msg) or ("503" in msg)

def _gsheet_retry(fn, *args, **kwargs):
    """Menjalankan panggilan API Google Sheets dengan backoff otomatis jika terkena limit 429/503."""
    delays = (1.5, 3.0, 6.0)
    last_err = None
    for attempt in range(len(delays) + 1):
        try:
            return fn(*args, **kwargs)
        except Exception as e:
            last_err = e
            if _is_quota_429(e) and attempt < len(delays):
                time.sleep(delays[attempt])
                continue
            raise last_err

@st.cache_resource(ttl=3600, show_spinner=False)
def get_cached_spreadsheet():
    """Satu koneksi Spreadsheet terpusat (di-cache 1 jam) agar seluruh modul tidak memanggil open_by_key berulang-ulang."""
    creds_json = st.secrets["gcp_json"]
    creds_dict = json.loads(creds_json)
    scopes = ["https://www.googleapis.com/auth/spreadsheets", "https://www.googleapis.com/auth/drive"]
    creds = Credentials.from_service_account_info(creds_dict, scopes=scopes)
    client = gspread.authorize(creds)
    return _gsheet_retry(client.open_by_key, SHEET_ID)

@st.cache_resource(ttl=3600, show_spinner=False)
def _get_cached_ws_only(sheet_name):
    sh = get_cached_spreadsheet()
    return _gsheet_retry(sh.worksheet, sheet_name)

def connect_gsheets(sheet_name="Report Preventive"):
    try:
        return _get_cached_ws_only(sheet_name)
    except Exception:
        return None

@st.cache_data(ttl=600, show_spinner=False)
def fetch_data_from_gsheets():
    """Membaca 4 sheet operasional utama dalam 1x batch request (hemat 11 read request per siklus)."""
    all_data = []
    sheet_names = ["Report Preventive", "Report PBU", "Project", "SiHD"]
    try:
        sh = get_cached_spreadsheet()
        ranges = [f"'{sn}'!A:F" for sn in sheet_names]
        batch_res = _gsheet_retry(sh.values_batch_get, ranges)
        v_ranges = batch_res.get("valueRanges", [])
        for sn, vr in zip(sheet_names, v_ranges):
            rows = vr.get("values", [])
            for idx, row in enumerate(rows[1:]):
                if len(row) >= 6:
                    try:
                        item = json.loads(row[5])
                        if sn == "Report Preventive" and item.get('kategori') == 'SiUPDATE':
                            item['kategori'] = 'SPS'
                        item['_sheet'] = sn
                        item['_row'] = idx + 2
                        all_data.append(item)
                    except Exception:
                        pass
        _GS_LAST_OK["__laporan_db__"] = all_data
        return all_data
    except Exception as e_batch:
        # Fallback aman jika salah satu sheet (misal 'SiHD') belum dibuat di Spreadsheet
        if _is_quota_429(e_batch) and "__laporan_db__" in _GS_LAST_OK:
            return _GS_LAST_OK["__laporan_db__"]
        for sn in sheet_names:
            ws_s = connect_gsheets(sn)
            if not ws_s:
                continue
            try:
                data_s = _gsheet_retry(ws_s.get_all_values)
                for idx, row in enumerate(data_s[1:]):
                    if len(row) >= 6:
                        try:
                            item = json.loads(row[5])
                            if sn == "Report Preventive" and item.get('kategori') == 'SiUPDATE':
                                item['kategori'] = 'SPS'
                            item['_sheet'] = sn
                            item['_row'] = idx + 2
                            all_data.append(item)
                        except Exception:
                            pass
            except Exception:
                pass
        if all_data:
            _GS_LAST_OK["__laporan_db__"] = all_data
        return all_data or _GS_LAST_OK.get("__laporan_db__", [])
def get_laporan_db(force_refresh=False):
    """Memuat database laporan secara lazy (hanya ketika menu yang membutuhkan dibuka)."""
    if force_refresh:
        fetch_data_from_gsheets.clear()
        st.session_state['laporan_db'] = fetch_data_from_gsheets()
    elif 'laporan_db' not in st.session_state:
        st.session_state['laporan_db'] = fetch_data_from_gsheets()
    return st.session_state['laporan_db']

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
    return text


# =========================================================================
# 🛡️ HELPER UNIVERSAL GOOGLE SHEETS (SiHD, PROGRAM NOP & SiDAVID)
# =========================================================================
@st.cache_resource(ttl=3600, show_spinner=False)
def _gs_cached_ws_and_head(sheet_name, default_headers_tuple):
    sh = get_cached_spreadsheet()
    default_headers = list(default_headers_tuple)
    try:
        ws = _gsheet_retry(sh.worksheet, sheet_name)
    except Exception as e_ws:
        if _is_quota_429(e_ws):
            raise
        ws = _gsheet_retry(sh.add_worksheet, title=sheet_name, rows=1000, cols=max(len(default_headers) + 4, 20))
        _gsheet_retry(ws.update, range_name="A1", values=[default_headers])
        return ws, default_headers
    head = _gsheet_retry(ws.row_values, 1)
    if not head or not any(str(c).strip() for c in head):
        _gsheet_retry(ws.update, range_name="A1", values=[default_headers])
        head = default_headers
    return ws, head

def gsheet_get_or_create_ws(sheet_name, default_headers):
    try:
        ws, _ = _gs_cached_ws_and_head(sheet_name, tuple(default_headers))
        return ws
    except Exception:
        return None

@st.cache_data(ttl=300, show_spinner=False)
def gsheet_read_df(sheet_name, default_headers):
    try:
        ws, _ = _gs_cached_ws_and_head(sheet_name, tuple(default_headers))
        vals = _gsheet_retry(ws.get_all_values)
    except Exception:
        if sheet_name in _GS_LAST_OK:
            return _GS_LAST_OK[sheet_name].copy(), True
        return pd.DataFrame(columns=default_headers), False
    if not vals or len(vals) < 1:
        df_e = pd.DataFrame(columns=default_headers)
        _GS_LAST_OK[sheet_name] = df_e
        return df_e, True
    head = vals[0]
    if not any(str(c).strip() for c in head):
        df_e = pd.DataFrame(columns=default_headers)
        _GS_LAST_OK[sheet_name] = df_e
        return df_e, True
    clean_head = []
    seen_h = {}
    for idx_h, h_val in enumerate(head):
        h_s = str(h_val).strip() or f"Kolom_{idx_h+1}"
        if h_s in seen_h:
            seen_h[h_s] += 1
            h_s = f"{h_s}_{seen_h[h_s]}"
        else:
            seen_h[h_s] = 0
        clean_head.append(h_s)
    rows = [r[:len(clean_head)] + [""] * max(0, len(clean_head) - len(r)) for r in vals[1:]]
    df = pd.DataFrame(rows, columns=clean_head)
    for col in default_headers:
        if col not in df.columns:
            df[col] = ""
    _GS_LAST_OK[sheet_name] = df
    return df, True

def gsheet_overwrite_ws(sheet_name, df):
    df_clean = df.fillna("").astype(str)
    headers = [str(c) for c in df_clean.columns.tolist()]
    ws = gsheet_get_or_create_ws(sheet_name, headers)
    if ws is None:
        return False
    _gsheet_retry(ws.clear)
    payload = [headers] + df_clean.values.tolist()
    if len(payload) > ws.row_count or len(headers) > ws.col_count:
        _gsheet_retry(ws.resize, rows=max(len(payload) + 100, ws.row_count), cols=max(len(headers) + 4, ws.col_count))
    _gsheet_retry(ws.update, range_name="A1", values=payload, value_input_option="RAW")
    _GS_LAST_OK[sheet_name] = df_clean
    _gs_cached_ws_and_head.clear()
    gsheet_read_df.clear()
    return True

def gsheet_append_row(sheet_name, default_headers, row_dict):
    try:
        ws, head = _gs_cached_ws_and_head(sheet_name, tuple(default_headers))
    except Exception:
        return False
    row_vals = [str(row_dict.get(c, "")) for c in head]
    _gsheet_retry(ws.append_row, row_vals, value_input_option="RAW")
    gsheet_read_df.clear()
    return True

def sidavid_eval_expired(punya_status, exp_str):
    if str(punya_status).strip().lower() != "punya":
        return "❌ Belum Punya"
    try:
        exp_dt = pd.to_datetime(exp_str, errors="coerce")
        if pd.isna(exp_dt):
            return "⚠️ Tgl Expired Kosong"
        sisa_hari = (exp_dt.date() - datetime.date.today()).days
        if sisa_hari < 0:
            return f"🔴 Expired ({abs(sisa_hari)} hari lalu)"
        elif sisa_hari <= 60:
            return f"🟡 Segera Expired ({sisa_hari} hari lagi)"
        return f"🟢 Aktif ({sisa_hari} hari)"
    except Exception:
        return "⚠️ Cek Tanggal"

def _norm_site_id(raw_val):
    s = str(raw_val or "").strip().upper()
    if not s or s in ("NAN", "NONE", "-"):
        return ""
    if len(s) >= 6 and s[:3].isalpha() and s[3:6].isdigit():
        return s[:6]
    return s.split()[0].split("-")[0]

def build_onsite_history_index(db_list):
    """Membangun index riwayat onsite & power dari seluruh laporan SIDATA per Site ID."""
    idx_map = {}
    for r in (db_list or []):
        kat = r.get("kategori", "")
        if kat in ("Project", "SiHD"):
            continue
        raw_sn = r.get("pb_site_id") or r.get("site_name") or ""
        sid = _norm_site_id(raw_sn)
        if not sid:
            continue
        ts = str(r.get("timestamp", ""))
        prob = str(r.get("problem_alarm") or r.get("action") or "-")
        rem = str(r.get("remark") or r.get("cond_data", {}).get("note_action") or "-")
        tek = str(r.get("teknisi") or r.get("pb_te") or "-")
        stat = str(r.get("status") or "-")

        v_ac, v_dc, load_dc = None, None, None
        if kat == "Site Reguler":
            cd = r.get("cond_data", {})
            v_vals = [pd.to_numeric(cd.get(k), errors="coerce") for k in ("v_rn", "v_sn", "v_tn")]
            v_ok = [v for v in v_vals if pd.notna(v)]
            v_ac = round(float(sum(v_ok) / len(v_ok)), 1) if v_ok else None
            ld = pd.to_numeric(cd.get("load_dc"), errors="coerce")
            load_dc = float(ld) if pd.notna(ld) else None
        elif kat == "SPS":
            vd = pd.to_numeric(r.get("rect_out_v"), errors="coerce")
            ld = pd.to_numeric(r.get("total_load"), errors="coerce")
            v_dc = float(vd) if pd.notna(vd) else None
            load_dc = float(ld) if pd.notna(ld) else None
        elif kat == "Pangkalan Bun":
            kel = r.get("pb_kelistrikan", {})
            va_raw = str(kel.get("v_rst_n", "")).replace(",", ".")
            va_nums = re.findall(r"\d+\.?\d*", va_raw)
            if va_nums:
                v_ac = round(float(va_nums[0]), 1)
            rects = r.get("pb_rectifiers", [])
            if rects:
                vd = pd.to_numeric(rects[0].get("dc_v"), errors="coerce")
                ld = pd.to_numeric(rects[0].get("load"), errors="coerce")
                v_dc = float(vd) if pd.notna(vd) else None
                load_dc = float(ld) if pd.notna(ld) else None

        entry = {
            "Timestamp": ts,
            "Tanggal": ts[:10],
            "Kategori": kat,
            "Teknisi": tek,
            "Status": stat,
            "Problem": prob,
            "Remark": rem,
            "V_AC": v_ac,
            "V_DC": v_dc,
            "Load_DC": load_dc
        }
        idx_map.setdefault(sid, []).append(entry)

    for sid in idx_map:
        idx_map[sid].sort(key=lambda x: x["Timestamp"])
    return idx_map

def _sanitize_xlsx_bytes(file_or_path):
    """Menghapus tag XML autoFilter yang memicu bug openpyxl pada beberapa file Excel."""
    if isinstance(file_or_path, str):
        with open(file_or_path, "rb") as _f:
            raw = _f.read()
    elif isinstance(file_or_path, bytes):
        raw = file_or_path
    else:
        if hasattr(file_or_path, "seek"):
            file_or_path.seek(0)
        raw = file_or_path.read()
        if hasattr(file_or_path, "seek"):
            file_or_path.seek(0)
    out_buf = BytesIO()
    with zipfile.ZipFile(BytesIO(raw), "r") as zin, zipfile.ZipFile(out_buf, "w", zipfile.ZIP_DEFLATED) as zout:
        for item in zin.infolist():
            data = zin.read(item.filename)
            if item.filename.startswith("xl/worksheets/") and item.filename.endswith(".xml"):
                data = re.sub(rb"<autoFilter\b[^>]*>.*?</autoFilter>", b"", data, flags=re.DOTALL)
                data = re.sub(rb"<autoFilter\b[^>]*/>", b"", data)
            zout.writestr(item, data)
    out_buf.seek(0)
    return out_buf

def get_excel_sheet_names_safe(file_or_path):
    try:
        if hasattr(file_or_path, "seek"):
            file_or_path.seek(0)
        return pd.ExcelFile(file_or_path).sheet_names
    except Exception:
        return pd.ExcelFile(_sanitize_xlsx_bytes(file_or_path)).sheet_names

def read_excel_safe_raw(file_or_path, sheet_name=0, header=0):
    try:
        if hasattr(file_or_path, "seek"):
            file_or_path.seek(0)
        return pd.read_excel(file_or_path, sheet_name=sheet_name, header=header)
    except Exception:
        return pd.read_excel(_sanitize_xlsx_bytes(file_or_path), sheet_name=sheet_name, header=header)

@st.cache_data(ttl=3600, show_spinner=False, max_entries=16)
def _parse_excel_smart_header_cached(path, stamp, sheet_name, target_keywords):
    return _parse_excel_smart_header_core(path, sheet_name, target_keywords)

def _parse_excel_smart_header_core(file_or_path, sheet_name, target_keywords):
    df_raw = read_excel_safe_raw(file_or_path, sheet_name=sheet_name, header=None)
    if df_raw is None or df_raw.empty:
        return pd.DataFrame()
    hdr_row = 0
    for r_idx in range(min(15, len(df_raw))):
        row_vals = [str(v).strip().lower() for v in df_raw.iloc[r_idx].tolist()]
        if any(any(kw in cell for kw in target_keywords) for cell in row_vals):
            hdr_row = r_idx
            break
    cols = []
    seen = {}
    for c_i, val in enumerate(df_raw.iloc[hdr_row].tolist()):
        cname = str(val).strip() if pd.notna(val) and str(val).strip() else f"Kolom_{c_i+1}"
        if cname in seen:
            seen[cname] += 1
            cname = f"{cname}_{seen[cname]}"
        else:
            seen[cname] = 0
        cols.append(cname)
    df_out = df_raw.iloc[hdr_row + 1:].copy()
    df_out.columns = cols
    site_col = next((c for c in cols if any(kw in c.lower() for kw in target_keywords)), cols[0])
    df_out = df_out.dropna(subset=[site_col])
    df_out = df_out[~df_out[site_col].astype(str).str.strip().str.lower().isin(["", "nan", "none", "total", "site id"])]
    return df_out.reset_index(drop=True)

def parse_excel_smart_header(file_or_path, sheet_name, target_keywords=("site id", "site_id", "siteid")):
    """Membaca Excel secara aman (anti-crash AutoFilter openpyxl) dan mendeteksi baris header otomatis + cache untuk file lokal."""
    if isinstance(file_or_path, str) and os.path.exists(file_or_path):
        stt = os.stat(file_or_path)
        return _parse_excel_smart_header_cached(file_or_path, (stt.st_mtime_ns, stt.st_size), sheet_name, tuple(target_keywords)).copy()
    return _parse_excel_smart_header_core(file_or_path, sheet_name, tuple(target_keywords))

# =========================================================================
# 🛰️ HELPER GEOSPASIAL & TIER ROUTE TRACKING NOP PANGKALAN BUN
# =========================================================================
def haversine_km(lat1, lon1, lat2, lon2):
    """Menghitung jarak garis lurus (km) antara dua koordinat GPS (Haversine)."""
    try:
        lat1, lon1, lat2, lon2 = float(lat1), float(lon1), float(lat2), float(lon2)
    except Exception:
        return None
    R = 6371.0
    p1, p2 = _math.radians(lat1), _math.radians(lat2)
    dp = _math.radians(lat2 - lat1)
    dl = _math.radians(lon2 - lon1)
    a = _math.sin(dp / 2.0) ** 2 + _math.cos(p1) * _math.cos(p2) * _math.sin(dl / 2.0) ** 2
    return round(2.0 * R * _math.asin(_math.sqrt(min(1.0, max(0.0, a)))), 2)

def _find_first_existing(candidates, glob_pat=None):
    import glob as _gf
    for c in candidates:
        if os.path.exists(c):
            return c
    if glob_pat:
        hits = [h for h in sorted(_gf.glob(glob_pat)) if "~$" not in h]
        if hits:
            return hits[-1]
    return None

def load_pbu_coords_master():
    """Memuat Master LongLat Site Tujuan Pangkalan Bun & HB Tim Pangkalan Bun (Cloud GSheets -> Fallback Excel lokal)."""
    COLS_SITE_PBU = ["Site ID", "Latitude Tujuan", "Longtitude Tujuan"]
    COLS_HB_PBU = ["NAMA", "NOP", "Latitude Berangkat", "Longtitude Berangkat"]

    df_site_c, _ = gsheet_read_df("Master_Site_PBU", COLS_SITE_PBU)
    df_hb_c, _ = gsheet_read_df("Master_HB_Tim_PBU", COLS_HB_PBU)

    src_site = "Google Sheets (Master_Site_PBU)"
    src_hb = "Google Sheets (Master_HB_Tim_PBU)"

    if df_site_c.empty:
        f_site = _find_first_existing(
            ["Site ID Sidata Pangkalanbun (1).xlsx", "Site ID Sidata Pangkalanbun.xlsx"],
            "*Site ID Sidata Pangkalanbun*.xlsx"
        )
        if f_site:
            try:
                df_site_c = read_excel_safe_raw(f_site, sheet_name=0, header=0)
                src_site = f"Excel Lokal ({os.path.basename(f_site)})"
            except Exception:
                df_site_c = pd.DataFrame(columns=COLS_SITE_PBU)

    if df_hb_c.empty:
        f_hb = _find_first_existing(
            ["NAMA NOP PANGKALANBUN (2).xlsx", "NAMA NOP PANGKALANBUN.xlsx"],
            "*NAMA NOP PANGKALANBUN*.xlsx"
        )
        if f_hb:
            try:
                df_hb_c = read_excel_safe_raw(f_hb, sheet_name=0, header=0)
                src_hb = f"Excel Lokal ({os.path.basename(f_hb)})"
            except Exception:
                df_hb_c = pd.DataFrame(columns=COLS_HB_PBU)

    site_coords = {}
    if not df_site_c.empty:
        c_sid = next((c for c in df_site_c.columns if "site" in str(c).lower()), df_site_c.columns[0])
        c_lat = next((c for c in df_site_c.columns if "lat" in str(c).lower()), None)
        c_lon = next((c for c in df_site_c.columns if "lon" in str(c).lower()), None)
        if c_lat and c_lon:
            for _, r in df_site_c.iterrows():
                sid = str(r.get(c_sid, "")).strip().upper()
                lat = pd.to_numeric(r.get(c_lat), errors="coerce")
                lon = pd.to_numeric(r.get(c_lon), errors="coerce")
                if sid and sid not in ("NAN", "NONE", "") and pd.notna(lat) and pd.notna(lon):
                    site_coords[sid] = (float(lat), float(lon))

    fme_coords = {}
    if not df_hb_c.empty:
        c_nm = next((c for c in df_hb_c.columns if "nama" in str(c).lower()), df_hb_c.columns[0])
        c_lat = next((c for c in df_hb_c.columns if "lat" in str(c).lower()), None)
        c_lon = next((c for c in df_hb_c.columns if "lon" in str(c).lower()), None)
        if c_lat and c_lon:
            for _, r in df_hb_c.iterrows():
                nm = str(r.get(c_nm, "")).strip().upper()
                lat = pd.to_numeric(r.get(c_lat), errors="coerce")
                lon = pd.to_numeric(r.get(c_lon), errors="coerce")
                if nm and nm not in ("NAN", "NONE", "") and pd.notna(lat) and pd.notna(lon):
                    fme_coords[nm] = (float(lat), float(lon))

    return df_site_c, df_hb_c, site_coords, fme_coords, src_site, src_hb

def _match_fme_coord(name_str, fme_coords):
    nm = str(name_str or "").strip().upper()
    if not nm or nm in ("NAN", "NONE", "-"):
        return None, ""
    if nm in fme_coords:
        return fme_coords[nm], nm
    for k, v in fme_coords.items():
        if nm in k or k in nm:
            return v, k
    return None, nm

def _clean_cell_str(v):
    if v is None:
        return ""
    s = str(v).strip()
    return "" if s.lower() in ("nan", "nat", "none") else s

def _fmt_date_short(v):
    s = _clean_cell_str(v)
    if not s:
        return "-"
    return s[:10]

COLS_HARD_ACCESS_PBU = [
    "Record_ID", "Timestamp", "Site_ID", "Site_Name", "Lokasi_Kendala",
    "Kategori_Hard_Access", "Tingkat_Kesulitan", "Status_Akses",
    "Lat_Kendala", "Lon_Kendala", "Estimasi_Tempuh", "Kendaraan_Wajib",
    "PIC_Input", "Keterangan_Detail", "Foto_URLs"
]

COLS_SITE_HARD_AKSES = [
    "Record_ID", "Timestamp", "NOP", "Site_ID", "Site_Name",
    "Titik_Start_Pos", "Lat_Start", "Lon_Start",
    "Lokasi_Kendala", "Lat_Kendala", "Lon_Kendala",
    "Lat_Site", "Lon_Site",
    "Kategori_Hard_Access", "Tingkat_Kesulitan", "Status_Akses",
    "Rute_Wajib_Moda", "Estimasi_Biaya_Sewa", "Estimasi_Waktu_Tempuh",
    "Rekomendasi_Rute_Detail", "PIC_Input", "Keterangan_Detail", "Foto_URLs"
]

HARD_ACCESS_KATEGORI = [
    "Jalan Lumpur / Amblas / Rusak Parah",
    "Perbaikan Jalan / Alat Berat / Akses Tidak Tembus",
    "Jembatan Putus / Rusak / Rawan",
    "Banjir / Akses Sungai & Perahu (Klotok)",
    "Izin Perusahaan / Corporate / Perkebunan Sawit",
    "Portal / Community Issue / Penutupan Warga",
    "Medan Ekstrem / Hutan / Wajib 4WD atau Trail",
    "Lainnya"
]

HARD_ACCESS_MODA_WAJIB = [
    "🏍️ Sewa Motor Trail",
    "🚙 Wajib 4WD / Double Cabin (4x4)",
    "🚤 Sewa Klotok / Perahu Sungai",
    "⛴️ Ferry / Rakit Penyeberangan",
    "🚜 Escort / Bantuan Alat Berat",
    "🥾 Jalan Kaki + Sewa Porter",
    "📄 Wajib Izin Corporate / Perkebunan",
    "🛠️ Rute Alternatif / Memutar",
    "🚗 Mobil Standar (Hati-hati)"
]

HARD_ACCESS_LEVEL = [
    "🔴 Ekstrem / Tidak Tembus",
    "🟠 Berat / Wajib Trail, 4WD, atau Klotok",
    "🟡 Sedang / Terbatas"
]

HARD_ACCESS_STATUS = [
    "🔴 Blocked / Tidak Tembus",
    "🟡 Terbatas / Wajib Moda Khusus",
    "🟢 Normal / Sudah Tembus"
]

def _parse_url_list_field(raw_val):
    if not raw_val:
        return []
    if isinstance(raw_val, list):
        return [str(u).strip() for u in raw_val if str(u).strip().startswith("http")]
    s = str(raw_val).strip()
    if not s or s.lower() in ("nan", "none", "[]"):
        return []
    try:
        parsed = json.loads(s)
        if isinstance(parsed, list):
            return [str(u).strip() for u in parsed if str(u).strip().startswith("http")]
    except Exception:
        pass
    return [p.strip() for p in re.split(r"[\n,;]+", s) if p.strip().startswith("http")]

def load_site_hard_akses(nop_filter=None, site_coords_plk=None, site_coords_pbu=None):
    """Membaca data Hard Akses dari Google Sheets tab 'Site Hard Akses' (plus sinkron bacaan lama 'Hard_Access_PBU' jika ada)."""
    site_coords_plk = site_coords_plk or {}
    site_coords_pbu = site_coords_pbu or {}
    df_sha, ok_sha = gsheet_read_df("Site Hard Akses", COLS_SITE_HARD_AKSES)

    all_items = []
    seen_rids = set()

    def _coord_lookup(nop_str, sid_u):
        is_pbu = "pangkalan" in str(nop_str).lower() or "pbu" in str(nop_str).lower()
        src = site_coords_pbu if is_pbu else site_coords_plk
        if sid_u in src:
            v = src[sid_u]
            return (v[0], v[1]) if isinstance(v, (tuple, list)) else (None, None)
        # Fallback cross lookup
        other = site_coords_plk if is_pbu else site_coords_pbu
        if sid_u in other:
            v = other[sid_u]
            return (v[0], v[1]) if isinstance(v, (tuple, list)) else (None, None)
        return (None, None)

    if not df_sha.empty:
        for _, rw in df_sha.iterrows():
            sid = _clean_cell_str(rw.get("Site_ID")).upper()
            if not sid:
                continue
            rid = _clean_cell_str(rw.get("Record_ID")) or f"SHA-{sid}"
            nop_v = _clean_cell_str(rw.get("NOP")) or "Pangkalan Bun"
            if nop_filter and nop_filter.lower() not in nop_v.lower():
                continue

            lat_k = pd.to_numeric(rw.get("Lat_Kendala"), errors="coerce")
            lon_k = pd.to_numeric(rw.get("Lon_Kendala"), errors="coerce")
            lat_s = pd.to_numeric(rw.get("Lat_Site"), errors="coerce")
            lon_s = pd.to_numeric(rw.get("Lon_Site"), errors="coerce")
            lat_st = pd.to_numeric(rw.get("Lat_Start"), errors="coerce")
            lon_st = pd.to_numeric(rw.get("Lon_Start"), errors="coerce")

            lk_lat, lk_lon = _coord_lookup(nop_v, _norm_site_id(sid))
            if pd.isna(lat_s) or pd.isna(lon_s):
                lat_s, lon_s = lk_lat, lk_lon
            if pd.isna(lat_k) or pd.isna(lon_k):
                lat_k, lon_k = lat_s, lon_s

            photos = _parse_url_list_field(rw.get("Foto_URLs"))
            moda_raw = _clean_cell_str(rw.get("Rute_Wajib_Moda")) or "-"
            biaya_num = pd.to_numeric(rw.get("Estimasi_Biaya_Sewa"), errors="coerce")
            biaya_val = float(biaya_num) if pd.notna(biaya_num) else 0.0

            dist_start_kendala = haversine_km(lat_st, lon_st, lat_k, lon_k) if (pd.notna(lat_st) and pd.notna(lon_st) and pd.notna(lat_k) and pd.notna(lon_k)) else None
            dist_kendala_site = haversine_km(lat_k, lon_k, lat_s, lon_s) if (pd.notna(lat_k) and pd.notna(lon_k) and pd.notna(lat_s) and pd.notna(lon_s)) else None
            total_route_km = round((dist_start_kendala or 0.0) + (dist_kendala_site or 0.0), 2)

            item = {
                "record_id": rid,
                "timestamp": _clean_cell_str(rw.get("Timestamp"))[:16] or "-",
                "nop": nop_v,
                "site_id": _norm_site_id(sid) or sid,
                "site_id_full": sid,
                "site_name": _clean_cell_str(rw.get("Site_Name")) or sid,
                "titik_start": _clean_cell_str(rw.get("Titik_Start_Pos")) or "-",
                "lat_start": float(lat_st) if pd.notna(lat_st) else None,
                "lon_start": float(lon_st) if pd.notna(lon_st) else None,
                "lokasi": _clean_cell_str(rw.get("Lokasi_Kendala")) or "-",
                "lat": float(lat_k) if pd.notna(lat_k) else None,
                "lon": float(lon_k) if pd.notna(lon_k) else None,
                "lat_site": float(lat_s) if pd.notna(lat_s) else None,
                "lon_site": float(lon_s) if pd.notna(lon_s) else None,
                "kategori": _clean_cell_str(rw.get("Kategori_Hard_Access")) or "-",
                "tingkat": _clean_cell_str(rw.get("Tingkat_Kesulitan")) or HARD_ACCESS_LEVEL[1],
                "status": _clean_cell_str(rw.get("Status_Akses")) or HARD_ACCESS_STATUS[0],
                "kendaraan": moda_raw,
                "biaya_sewa": biaya_val,
                "estimasi": _clean_cell_str(rw.get("Estimasi_Waktu_Tempuh")) or "-",
                "rekomendasi_rute": _clean_cell_str(rw.get("Rekomendasi_Rute_Detail")) or "-",
                "pic": _clean_cell_str(rw.get("PIC_Input")) or "-",
                "keterangan": _clean_cell_str(rw.get("Keterangan_Detail")) or "-",
                "photos": photos,
                "dist_total_km": total_route_km,
                "source_sheet": "Site Hard Akses"
            }
            all_items.append(item)
            seen_rids.add(rid)

    # Backward compatibility: gabungkan data lama dari 'Hard_Access_PBU' jika ada (hanya saat filter None atau Pangkalan Bun)
    if nop_filter is None or "pangkalan" in str(nop_filter).lower():
        try:
            df_old, ok_old = gsheet_read_df("Hard_Access_PBU", COLS_HARD_ACCESS_PBU)
            if ok_old and not df_old.empty:
                for _, rw in df_old.iterrows():
                    sid = _clean_cell_str(rw.get("Site_ID")).upper()
                    rid = _clean_cell_str(rw.get("Record_ID")) or f"OLD-{sid}"
                    if not sid or rid in seen_rids:
                        continue
                    lat_k = pd.to_numeric(rw.get("Lat_Kendala"), errors="coerce")
                    lon_k = pd.to_numeric(rw.get("Lon_Kendala"), errors="coerce")
                    lk_lat, lk_lon = _coord_lookup("Pangkalan Bun", _norm_site_id(sid))
                    if pd.isna(lat_k) or pd.isna(lon_k):
                        lat_k, lon_k = lk_lat, lk_lon
                    all_items.append({
                        "record_id": rid,
                        "timestamp": _clean_cell_str(rw.get("Timestamp"))[:16] or "-",
                        "nop": "Pangkalan Bun",
                        "site_id": _norm_site_id(sid) or sid,
                        "site_id_full": sid,
                        "site_name": _clean_cell_str(rw.get("Site_Name")) or sid,
                        "titik_start": "-",
                        "lat_start": None, "lon_start": None,
                        "lokasi": _clean_cell_str(rw.get("Lokasi_Kendala")) or "-",
                        "lat": float(lat_k) if pd.notna(lat_k) else None,
                        "lon": float(lon_k) if pd.notna(lon_k) else None,
                        "lat_site": lk_lat, "lon_site": lk_lon,
                        "kategori": _clean_cell_str(rw.get("Kategori_Hard_Access")) or "-",
                        "tingkat": _clean_cell_str(rw.get("Tingkat_Kesulitan")) or HARD_ACCESS_LEVEL[1],
                        "status": _clean_cell_str(rw.get("Status_Akses")) or HARD_ACCESS_STATUS[0],
                        "kendaraan": _clean_cell_str(rw.get("Kendaraan_Wajib")) or "-",
                        "biaya_sewa": 0.0,
                        "estimasi": _clean_cell_str(rw.get("Estimasi_Tempuh")) or "-",
                        "rekomendasi_rute": "-",
                        "pic": _clean_cell_str(rw.get("PIC_Input")) or "-",
                        "keterangan": _clean_cell_str(rw.get("Keterangan_Detail")) or "-",
                        "photos": _parse_url_list_field(rw.get("Foto_URLs")),
                        "dist_total_km": 0.0,
                        "source_sheet": "Hard_Access_PBU"
                    })
                    seen_rids.add(rid)
        except Exception:
            pass

    ha_by_site = {}
    for it in all_items:
        ha_by_site.setdefault(it["site_id"], []).append(it)
    return df_sha, ok_sha, all_items, ha_by_site

def load_pbu_hard_access_records(site_coords=None):
    """Wrapper untuk Program NOP Pangkalan Bun agar otomatis tersinkron dengan sheet 'Site Hard Akses'."""
    return load_site_hard_akses(nop_filter="Pangkalan Bun", site_coords_pbu=site_coords)

def render_hard_access_route_map_html(ha_list, map_theme="Dark NOC Command Center", show_routes=True):
    """Peta Interaktif Premium khusus Rute & Lokasi Hard Akses (Palangkaraya & Pangkalan Bun) + Foto & Moda Sewa Wajib."""
    if "Satelit" in map_theme:
        tile_url = "https://server.arcgisonline.com/ArcGIS/rest/services/World_Imagery/MapServer/tile/{z}/{y}/{x}"
        tile_sub = ""
    elif "Terang" in map_theme:
        tile_url = "https://{s}.basemaps.cartocdn.com/rastertiles/voyager/{z}/{x}/{y}{r}.png"
        tile_sub = "abcd"
    else:
        tile_url = "https://{s}.basemaps.cartocdn.com/dark_all/{z}/{x}/{y}{r}.png"
        tile_sub = "abcd"

    ha_json = json.dumps(ha_list or [], ensure_ascii=False)
    return f"""
    <!DOCTYPE html>
    <html>
    <head>
        <meta charset="utf-8" />
        <link rel="stylesheet" href="https://unpkg.com/leaflet@1.9.4/dist/leaflet.css" />
        <script src="https://unpkg.com/leaflet@1.9.4/dist/leaflet.js"></script>
        <style>
            body {{ margin:0; padding:0; font-family:'Plus Jakarta Sans',system-ui,sans-serif; background:transparent; }}
            #hamap {{ width:100%; height:590px; border-radius:16px; border:1.5px solid rgba(249,115,22,0.55); box-shadow:0 14px 34px rgba(0,0,0,0.28); overflow:hidden; }}
            .pin-start {{ width:32px; height:32px; border-radius:50%; background:linear-gradient(135deg,#2563eb,#1d4ed8); border:2.5px solid #fff; color:#fff; display:flex; align-items:center; justify-content:center; font-size:15px; box-shadow:0 0 14px rgba(37,99,235,0.85); }}
            .pin-ha {{ width:36px; height:36px; border-radius:11px; border:2.5px solid #fef08a; color:#fff; display:flex; align-items:center; justify-content:center; font-size:17px; box-shadow:0 0 18px rgba(249,115,22,0.95); animation:pulseHA 1.8s infinite; }}
            .pin-site {{ width:32px; height:32px; border-radius:50%; background:linear-gradient(135deg,#8b5cf6,#6d28d9); border:2.5px solid #fff; color:#fff; display:flex; align-items:center; justify-content:center; font-size:15px; box-shadow:0 0 14px rgba(139,92,246,0.9); }}
            @keyframes pulseHA {{ 0% {{ box-shadow:0 0 0 0 rgba(249,115,22,0.8); }} 70% {{ box-shadow:0 0 0 15px rgba(249,115,22,0); }} 100% {{ box-shadow:0 0 0 0 rgba(249,115,22,0); }} }}
            .pin-lbl {{ position:absolute; top:36px; left:50%; transform:translateX(-50%); background:rgba(15,23,42,0.92); color:#f8fafc; font-size:9.5px; font-weight:800; padding:2px 7px; border-radius:5px; white-space:nowrap; border:1px solid rgba(255,255,255,0.22); pointer-events:none; }}
            .moda-pill {{ background:rgba(15,23,42,0.94); color:#fef08a; font-size:10px; font-weight:800; padding:2px 8px; border-radius:99px; white-space:nowrap; border:1.5px solid #f97316; box-shadow:0 2px 8px rgba(0,0,0,0.45); pointer-events:none; }}
            @keyframes dashFlow {{ from {{ stroke-dashoffset: 36; }} to {{ stroke-dashoffset: 0; }} }}
            .flow-ha {{ animation: dashFlow 1.5s linear infinite; }}
            .leaflet-popup-content-wrapper {{ background:linear-gradient(145deg,#0f172a,#1e293b) !important; color:#f8fafc !important; border-radius:14px !important; border:1px solid rgba(249,115,22,0.55); max-width:340px; }}
            .leaflet-popup-tip {{ background:#1e293b !important; }}
            .pc {{ min-width:255px; font-size:12px; line-height:1.55; }}
            .pc-h {{ font-size:13.5px; font-weight:800; color:#fb923c; border-bottom:1px solid rgba(255,255,255,0.14); padding-bottom:5px; margin-bottom:6px; }}
            .pc-moda {{ margin:6px 0; padding:6px 8px; border-radius:8px; background:rgba(249,115,22,0.16); border:1px solid rgba(249,115,22,0.45); color:#fed7aa; font-weight:700; }}
            .pc-rmk {{ margin-top:6px; padding:6px 8px; border-radius:8px; background:rgba(255,255,255,0.08); border-left:3px solid #38bdf8; }}
            .pc-photos {{ display:grid; grid-template-columns:repeat(2,1fr); gap:5px; margin-top:7px; }}
            .pc-photos img {{ width:100%; height:72px; object-fit:cover; border-radius:6px; border:1px solid rgba(255,255,255,0.22); }}
            .pc-btn {{ display:block; text-align:center; margin-top:8px; padding:6px; border-radius:7px; background:#ea580c; color:#fff !important; text-decoration:none; font-weight:800; font-size:11px; }}
            .map-leg {{ background:rgba(15,23,42,0.92); color:#f8fafc; padding:10px 13px; border-radius:12px; border:1px solid rgba(255,255,255,0.2); font-size:11px; line-height:1.7; }}
        </style>
    </head>
    <body>
        <div id="hamap"></div>
        <script>
            const items = {ha_json};
            const showRoutes = {"true" if show_routes else "false"};
            const map = L.map('hamap', {{ center: [-2.15, 113.0], zoom: 7 }});
            L.tileLayer('{tile_url}', {{ subdomains: '{tile_sub}', maxZoom: 19, attribution: '⚡ SIDATA Hard Akses Route Map' }}).addTo(map);
            const bounds = [];

            function makePin(cls, iconTxt, lblTxt, bg) {{
                const st = bg ? `background:${{bg}};` : '';
                return L.divIcon({{
                    className: '', iconSize: [36, 36], iconAnchor: [18, 18], popupAnchor: [0, -18],
                    html: `<div style="position:relative;"><div class="${{cls}}" style="${{st}}">${{iconTxt}}</div><div class="pin-lbl">${{lblTxt}}</div></div>`
                }});
            }}

            function curvePts(p1, p2, bend) {{
                const dLat = p2[0] - p1[0], dLon = p2[1] - p1[1];
                if (Math.hypot(dLat, dLon) < 0.0005) return [p1, p2];
                const mLat = (p1[0] + p2[0]) / 2 - dLon * bend;
                const mLon = (p1[1] + p2[1]) / 2 + dLat * bend;
                const out = [];
                for (let i = 0; i <= 26; i++) {{
                    const t = i / 26, inv = 1 - t;
                    out.push([inv*inv*p1[0] + 2*inv*t*mLat + t*t*p2[0], inv*inv*p1[1] + 2*inv*t*mLon + t*t*p2[1]]);
                }}
                return out;
            }}

            function drawRouteSeg(p1, p2, col, bend, badgeTxt) {{
                const pts = curvePts(p1, p2, bend);
                L.polyline(pts, {{ color: col, weight: 9, opacity: 0.22 }}).addTo(map);
                L.polyline(pts, {{ color: col, weight: 4, opacity: 0.94 }}).addTo(map);
                L.polyline(pts, {{ color: '#ffffff', weight: 1.8, opacity: 0.7, dashArray: '6, 12', className: 'flow-ha' }}).addTo(map);
                const arrIdx = Math.floor(pts.length * 0.65);
                const pa1 = pts[Math.max(0, arrIdx - 1)], pa2 = pts[Math.min(pts.length - 1, arrIdx + 1)];
                const deg = Math.atan2(pa2[1] - pa1[1], pa2[0] - pa1[0]) * 180 / Math.PI;
                L.marker(pts[arrIdx], {{
                    interactive: false,
                    icon: L.divIcon({{
                        className: '', iconSize: [18, 18], iconAnchor: [9, 9],
                        html: `<div style="transform:rotate(${{deg}}deg);"><svg width="16" height="16" viewBox="0 0 16 16"><path d="M8 1 L14 13 L8 10 L2 13 Z" fill="${{col}}" stroke="#0f172a" stroke-width="1.2"/></svg></div>`
                    }})
                }}).addTo(map);
                if (badgeTxt && items.length <= 12) {{
                    const mid = pts[Math.floor(pts.length / 2)];
                    L.marker(mid, {{
                        interactive: false,
                        icon: L.divIcon({{ className: '', iconSize: [0, 0], html: `<div class="moda-pill" style="border-color:${{col}};transform:translate(-50%,-50%);">${{badgeTxt}}</div>` }})
                    }}).addTo(map);
                }}
            }}

            items.forEach(ha => {{
                const hasHA = ha.lat !== null && ha.lon !== null;
                const hasSt = ha.lat_start !== null && ha.lon_start !== null;
                const hasSite = ha.lat_site !== null && ha.lon_site !== null;
                const stLow = (ha.status || '').toLowerCase();
                const isNorm = stLow.includes('normal') || (stLow.includes('tembus') && !stLow.includes('tidak'));
                const isMid = stLow.includes('terbatas') || stLow.includes('moda');
                const colMain = isNorm ? '#10b981' : (isMid ? '#f59e0b' : '#ef4444');
                const haGrad = isNorm ? 'linear-gradient(135deg,#10b981,#047857)' : (isMid ? 'linear-gradient(135deg,#f59e0b,#d97706)' : 'linear-gradient(135deg,#f97316,#dc2626)');

                if (showRoutes && hasSt && hasHA) {{
                    bounds.push([ha.lat_start, ha.lon_start]);
                    L.marker([ha.lat_start, ha.lon_start], {{ icon: makePin('pin-start', '🏠', 'POS: ' + (ha.titik_start || 'Start')) }}).addTo(map)
                        .bindPopup(`<div class="pc"><div class="pc-h" style="color:#38bdf8;">🏠 TITIK START / POS TERAKHIR</div><div><b>Pos:</b> ${{ha.titik_start}}</div><div><b>Menuju Site:</b> ${{ha.site_id}} (${{ha.nop}})</div><div><b>Moda Wajib:</b> ${{ha.kendaraan}}</div></div>`);
                    drawRouteSeg([ha.lat_start, ha.lon_start], [ha.lat, ha.lon], colMain, 0.16, (ha.kendaraan || '').slice(0, 26));
                }}

                if (showRoutes && hasHA && hasSite && Math.hypot(ha.lat - ha.lat_site, ha.lon - ha.lon_site) > 0.001) {{
                    bounds.push([ha.lat_site, ha.lon_site]);
                    L.marker([ha.lat_site, ha.lon_site], {{ icon: makePin('pin-site', '🗼', 'SITE: ' + ha.site_id) }}).addTo(map)
                        .bindPopup(`<div class="pc"><div class="pc-h" style="color:#c084fc;">🗼 SITE TUJUAN: ${{ha.site_id}}</div><div><b>Site Name:</b> ${{ha.site_name}}</div><div><b>NOP:</b> ${{ha.nop}}</div><a class="pc-btn" style="background:#7c3aed;" href="https://www.google.com/maps?q=${{ha.lat_site}},${{ha.lon_site}}" target="_blank">🧭 Navigasi ke Site</a></div>`);
                    drawRouteSeg([ha.lat, ha.lon], [ha.lat_site, ha.lon_site], '#a855f7', -0.15, '');
                }}

                if (hasHA) {{
                    bounds.push([ha.lat, ha.lon]);
                    let phHtml = '';
                    if (ha.photos && ha.photos.length > 0) {{
                        phHtml = '<div class="pc-photos">' + ha.photos.slice(0, 4).map(u => {{
                            const th = u.includes('cloudinary') ? u.replace('upload/', 'upload/w_320,q_auto:eco/') : u;
                            return `<a href="${{u}}" target="_blank"><img src="${{th}}" loading="lazy"></a>`;
                        }}).join('') + '</div>';
                    }}
                    const biayaTxt = ha.biaya_sewa > 0 ? `Rp ${{Number(ha.biaya_sewa).toLocaleString('id-ID')}}` : '-';
                    L.marker([ha.lat, ha.lon], {{ icon: makePin('pin-ha', '🚧', `${{ha.site_id}} (${{ha.nop.includes('Pangkalan') ? 'PBU' : 'PLK'}})`, haGrad) }}).addTo(map)
                        .bindPopup(`
                            <div class="pc">
                                <div class="pc-h">🚧 HARD AKSES: ${{ha.site_id}} — ${{ha.site_name}}</div>
                                <div>🏢 <b>NOP:</b> ${{ha.nop}} | <b>Status:</b> ${{ha.status}}</div>
                                <div>📍 <b>Lokasi Kendala:</b> ${{ha.lokasi}}</div>
                                <div>🏷️ <b>Kategori:</b> ${{ha.kategori}}</div>
                                <div class="pc-moda">🚙 Moda Wajib / Sewa: ${{ha.kendaraan}}<br>💰 Est. Biaya Sewa: ${{biayaTxt}} | ⏱️ ${{ha.estimasi}}</div>
                                <div class="pc-rmk"><b>🛣️ Rute Wajib:</b> ${{ha.rekomendasi_rute}}<br><b>📝 Detail:</b> ${{ha.keterangan}}</div>
                                ${{phHtml}}
                                <a class="pc-btn" href="https://www.google.com/maps?q=${{ha.lat}},${{ha.lon}}" target="_blank">🧭 Buka Titik Hard Akses di Maps</a>
                            </div>
                        `);
                }}
            }});

            if (bounds.length > 0) {{
                map.fitBounds(bounds, {{ padding: [45, 45], maxZoom: 12 }});
            }}

            const leg = L.control({{ position: 'bottomright' }});
            leg.onAdd = function() {{
                const d = L.DomUtil.create('div', 'map-leg');
                d.innerHTML = `
                    <b>🚧 Legenda Peta Rute Hard Akses</b><br>
                    🏠 Titik Start / Pos Terakhir &nbsp;|&nbsp; 🗼 Site Tujuan<br>
                    🚧 Titik Kendala Hard Akses (Klik untuk Foto)<br>
                    <b style="color:#ef4444">━</b> Blocked &nbsp; <b style="color:#f59e0b">━</b> Wajib Trail/4WD/Klotok &nbsp; <b style="color:#10b981">━</b> Normal
                `;
                return d;
            }};
            leg.addTo(map);
        </script>
    </body>
    </html>
    """

# =========================================================================
# 📅 ENGINE DAILY TRACKER PLK (LEADER CLUSTER + AUTO-CLOSED INTEGRASI SIDATA)
# =========================================================================
COLS_DAILY_TRACKER_PLK = [
    "Tracker_ID", "Tanggal", "Cluster", "Leader_Cluster",
    "Plan_Site_ID", "Program_Kerja", "Plan_Remark",
    "Tim_Tier_1", "Tim_Tier_2", "Tim_Tier_3",
    "Tgl_Tracker_Tier_1", "Tracker_Tier_1_Site_ID", "Tracker_Tier_1_Remark",
    "Tgl_Tracker_Tier_2", "Tracker_Tier_2_Site_ID", "Tracker_Tier_2_Remark",
    "Tgl_Tracker_Tier_3", "Tracker_Tier_3_Site_ID", "Tracker_Tier_3_Remark",
    "Updated_At"
]

DAILY_PROGRAM_KERJA_OPTS = ["PM", "Tshoot", "Task Force", "Backup Power", "Audit Tower", "Project / SVA", "Lainnya"]
DAILY_CLUSTER_PLK_OPTS = [
    "Palangkaraya Kota", "Kapuas - Pulang Pisau", "Katingan - Kasongan",
    "Gunung Mas - Kuala Kurun", "Barito Utara - Muara Teweh",
    "Barito Selatan - Buntok", "Barito Timur - Tamiang Layang",
    "Murung Raya - Puruk Cahu", "Lainnya"
]

def build_sidata_date_site_index(db_list):
    """Membangun index cepat (site_id_norm, tanggal_YYYY_MM_DD) -> list report SIDATA untuk auto-closed Daily Tracker."""
    idx = {}
    for r in (db_list or []):
        if r.get("kategori") == "SiHD" or r.get("_sheet") == "SiHD":
            continue
        ts = str(r.get("timestamp", "")).strip()
        tgl = ts[:10]
        if len(tgl) != 10:
            continue
        raw_sn = r.get("pb_site_id") or r.get("site_name") or ""
        sid_n = _norm_site_id(raw_sn)
        if not sid_n:
            continue
        entry = {
            "jam": ts[11:16] if len(ts) >= 16 else "-",
            "teknisi": str(r.get("teknisi") or r.get("pb_te") or "-"),
            "kategori": str(r.get("kategori") or "-"),
            "status": str(r.get("status") or "-")
        }
        idx.setdefault((sid_n, tgl), []).append(entry)
    return idx

def evaluate_daily_tracker_with_sidata(df_dt, db_list):
    """Mencocokkan Plan Site & Tracker Tier #1/#2/#3 dengan report aktual tim di SIDATA untuk menentukan status CLOSED vs OPEN."""
    if df_dt is None or df_dt.empty:
        return pd.DataFrame()

    sidata_idx = build_sidata_date_site_index(db_list)
    eval_rows = []

    for _, rw in df_dt.iterrows():
        tgl_plan = _fmt_date_short(rw.get("Tanggal"))
        plan_sid_raw = _clean_cell_str(rw.get("Plan_Site_ID"))
        plan_sid = _norm_site_id(plan_sid_raw)

        tiers_filled = []
        for k_t in (1, 2, 3):
            t_sid_raw = _clean_cell_str(rw.get(f"Tracker_Tier_{k_t}_Site_ID"))
            if t_sid_raw and t_sid_raw != "-":
                t_tgl = _fmt_date_short(rw.get(f"Tgl_Tracker_Tier_{k_t}"))
                if t_tgl == "-":
                    t_tgl = tgl_plan
                tiers_filled.append({
                    "tier": k_t,
                    "tgl": t_tgl,
                    "site_raw": t_sid_raw,
                    "site_id": _norm_site_id(t_sid_raw),
                    "remark": _clean_cell_str(rw.get(f"Tracker_Tier_{k_t}_Remark")) or "-"
                })

        plan_hits = sidata_idx.get((plan_sid, tgl_plan), []) if (plan_sid and tgl_plan != "-") else []
        plan_ok = len(plan_hits) > 0

        tier_results = []
        for tf in tiers_filled:
            th = sidata_idx.get((tf["site_id"], tf["tgl"]), []) if (tf["site_id"] and tf["tgl"] != "-") else []
            tier_results.append((tf, th))

        bukti_parts = []
        missing_sites = []

        if plan_sid:
            if plan_ok:
                bukti_parts.append(f"✅ Plan {plan_sid} ({plan_hits[-1]['jam']} · {plan_hits[-1]['teknisi']})")
            elif not tiers_filled:
                bukti_parts.append(f"❌ Plan {plan_sid} (Belum Input)")
                missing_sites.append(f"{plan_sid} ({tgl_plan})")
            else:
                bukti_parts.append(f"🔀 Plan {plan_sid} (Dialihkan ke Tier)")

        for tf, th in tier_results:
            if th:
                bukti_parts.append(f"✅ T#{tf['tier']} {tf['site_id']} ({th[-1]['jam']} · {th[-1]['teknisi']})")
            else:
                bukti_parts.append(f"❌ T#{tf['tier']} {tf['site_id']} (Belum Input)")
                missing_sites.append(f"T#{tf['tier']} {tf['site_id']} ({tf['tgl']})")

        if not tiers_filled:
            exec_cat = "Sesuai Plan"
            if plan_ok:
                auto_status = "✅ CLOSED (Terlapor di SIDATA)"
            else:
                auto_status = "🔴 OPEN — Wajib Follow Up Tim"
        else:
            exec_cat = "🔀 Dialihkan Urgent (" + " ➔ ".join(f"T#{t['tier']}:{t['site_id']}" for t in tiers_filled) + ")"
            tier_ok_cnt = sum(1 for _, th in tier_results if len(th) > 0)
            if tier_ok_cnt == len(tier_results):
                auto_status = "✅ CLOSED (Plan & Semua Tier Terlapor)" if plan_ok else "✅ CLOSED (Divert Tier Terlapor di SIDATA)"
            elif tier_ok_cnt > 0 or plan_ok:
                auto_status = "🟡 PARTIAL OPEN — Follow Up Tim"
            else:
                auto_status = "🔴 OPEN — Wajib Follow Up Tim"

        tim_list = [x for x in [_clean_cell_str(rw.get("Tim_Tier_1")), _clean_cell_str(rw.get("Tim_Tier_2")), _clean_cell_str(rw.get("Tim_Tier_3"))] if x and x != "-"]
        row_out = dict(rw)
        row_out["Tim_Bertugas"] = ", ".join(tim_list) if tim_list else "-"
        row_out["Eksekusi_Lapangan"] = exec_cat
        row_out["Status_Auto_SIDATA"] = auto_status
        row_out["Bukti_Report_SIDATA"] = " | ".join(bukti_parts) if bukti_parts else "-"
        row_out["Site_Belum_Input_SIDATA"] = ", ".join(missing_sites) if missing_sites else "Lengkap ✅"
        eval_rows.append(row_out)

    return pd.DataFrame(eval_rows)

def analyze_pbu_timeplan_routes(df_tp, site_coords, fme_coords, onsite_idx=None, hard_access_by_site=None, road_factor=1.35):
    """Analisa menyeluruh Timeplan Pangkalan Bun: HB Tim -> Site Target, tarikan Tier #1->#2->#3, jarak km, Hard Access & remark."""
    onsite_idx = onsite_idx or {}
    hard_access_by_site = hard_access_by_site or {}
    cols = list(df_tp.columns)

    def find_col(patterns, exclude=()):
        for pat in patterns:
            for c in cols:
                cl = str(c).strip().lower()
                if any(ex in cl for ex in exclude):
                    continue
                if pat == cl or pat in cl:
                    return c
        return None

    c_no = find_col(["no"])
    c_sid = find_col(["site id", "site_id"], exclude=("tracker",))
    c_sname = find_col(["site name", "site_name"])
    c_pic = find_col(["pic task", "pic"])
    c_t1 = find_col(["tier tim #1", "tier tim 1"])
    c_t2 = find_col(["tier tim #2", "tier tim 2"])
    c_t3 = find_col(["tier tim #3", "tier tim 3"])
    c_cls = find_col(["site class", "class"])
    c_rca = find_col(["rca"])
    c_act = find_col(["action detail", "action"])
    c_pdt = find_col(["plan date"])
    c_tcl = find_col(["target closed"])
    c_sta = find_col(["status activity", "status"])
    c_bya = find_col(["biaya"])
    c_rplan = find_col(["remark (target", "remark target"])
    c_rhasil = find_col(["remark hasil"])

    tracker_triples = []
    for n_tr in (1, 2, 3):
        ct_tgl = find_col([f"tracker #{n_tr} tgl", f"tracker {n_tr} tgl"])
        ct_sid = find_col([f"tracker #{n_tr} site id", f"tracker {n_tr} site id"])
        ct_rem = None
        if ct_sid is not None:
            idx_s = cols.index(ct_sid)
            if idx_s + 1 < len(cols) and "remark" in str(cols[idx_s + 1]).strip().lower():
                ct_rem = cols[idx_s + 1]
        if ct_rem is None:
            if n_tr == 1:
                ct_rem = next((c for c in cols if str(c).strip().lower() == "remark"), None)
            else:
                ct_rem = next((c for c in cols if str(c).strip().lower() in (f"remark_{n_tr-1}", f"remark.{n_tr-1}")), None)
        tracker_triples.append((n_tr, ct_tgl, ct_sid, ct_rem))

    summary_rows = []
    map_routes = []
    today_dt = datetime.date.today()

    for row_i, r in df_tp.iterrows():
        target_sid = _clean_cell_str(r.get(c_sid) if c_sid else "").upper()
        if not target_sid:
            continue
        no_val = _clean_cell_str(r.get(c_no)) if c_no else str(row_i + 1)
        sname = _clean_cell_str(r.get(c_sname)) if c_sname else "-"
        pic_task = _clean_cell_str(r.get(c_pic)) if c_pic else "-"
        tim1 = _clean_cell_str(r.get(c_t1)) if c_t1 else ""
        tim2 = _clean_cell_str(r.get(c_t2)) if c_t2 else ""
        tim3 = _clean_cell_str(r.get(c_t3)) if c_t3 else ""
        s_class = _clean_cell_str(r.get(c_cls)) if c_cls else "-"
        rca = _clean_cell_str(r.get(c_rca)) if c_rca else "-"
        action_det = _clean_cell_str(r.get(c_act)) if c_act else "-"
        plan_dt = _fmt_date_short(r.get(c_pdt)) if c_pdt else "-"
        target_cl = _fmt_date_short(r.get(c_tcl)) if c_tcl else "-"
        status_act = _clean_cell_str(r.get(c_sta)) if c_sta else "Open"
        biaya_val = pd.to_numeric(r.get(c_bya), errors="coerce") if c_bya else 0.0
        biaya_val = float(biaya_val) if pd.notna(biaya_val) else 0.0
        rem_plan = _clean_cell_str(r.get(c_rplan)) if c_rplan else "-"
        rem_hasil = _clean_cell_str(r.get(c_rhasil)) if c_rhasil else "-"

        hb_coord, hb_matched_name = _match_fme_coord(tim1, fme_coords)
        tgt_coord = site_coords.get(target_sid)

        dist_hb_tgt = haversine_km(hb_coord[0], hb_coord[1], tgt_coord[0], tgt_coord[1]) if (hb_coord and tgt_coord) else None
        est_road_tgt = round(dist_hb_tgt * road_factor, 1) if dist_hb_tgt is not None else None

        prev_coord = hb_coord
        prev_label = f"HB {tim1 or 'Tim'}"
        total_act_km = 0.0
        waypoints = []
        notes_non_site = []
        visited_sids = []

        tr_flat = {}
        for n_tr, ct_tgl, ct_sid, ct_rem in tracker_triples:
            t_tgl = _fmt_date_short(r.get(ct_tgl)) if ct_tgl else "-"
            t_sid_raw = _clean_cell_str(r.get(ct_sid)) if ct_sid else ""
            t_rem = _clean_cell_str(r.get(ct_rem)) if ct_rem else ""
            leg_km = None

            if t_sid_raw:
                t_sid_u = t_sid_raw.upper()
                sc = site_coords.get(t_sid_u)
                if sc:
                    visited_sids.append(t_sid_u)
                    if prev_coord:
                        leg_km = haversine_km(prev_coord[0], prev_coord[1], sc[0], sc[1])
                        total_act_km += leg_km
                    waypoints.append({
                        "tier": n_tr,
                        "tgl": t_tgl,
                        "site_id": t_sid_u,
                        "lat": sc[0],
                        "lon": sc[1],
                        "from_label": prev_label,
                        "leg_km": leg_km if leg_km is not None else 0.0,
                        "est_road_km": round((leg_km or 0.0) * road_factor, 1),
                        "remark": t_rem or (rem_hasil if t_sid_u == target_sid else "-")
                    })
                    prev_coord = sc
                    prev_label = f"Tier #{n_tr} ({t_sid_u})"
                else:
                    notes_non_site.append(f"Tier #{n_tr} ({t_tgl}): {t_sid_raw}" + (f" — {t_rem}" if t_rem else ""))

            tr_flat[f"Tier #{n_tr} Tgl"] = t_tgl
            tr_flat[f"Tier #{n_tr} Site/Info"] = t_sid_raw or "-"
            tr_flat[f"Tier #{n_tr} Jarak (km)"] = leg_km if leg_km is not None else "-"
            tr_flat[f"Tier #{n_tr} Remark"] = t_rem or "-"

        est_road_act = round(total_act_km * road_factor, 1) if total_act_km > 0 else 0.0

        # Deteksi Hard Access dari Sheet Hard_Access_PBU maupun teks catatan akses di Timeplan
        ha_recs_tgt = hard_access_by_site.get(target_sid, [])
        all_txt_check = " ".join([rem_hasil, rem_plan] + notes_non_site).lower()
        has_access_kw = any(k in all_txt_check for k in ("akses", "tidak tembus", "jalan", "lumpur", "banjir", "jembatan", "perahu", "corporate"))
        if ha_recs_tgt:
            ha_latest = ha_recs_tgt[-1]
            ha_status_lbl = f"🚧 {ha_latest['status']} — {ha_latest['kategori']} ({len(ha_latest['photos'])} foto)"
        elif has_access_kw:
            ha_status_lbl = "⚠️ Terindikasi Kendala Akses (Belum di-Report Hard Access)"
        else:
            ha_status_lbl = "✅ Normal"

        st_low = status_act.lower()
        visited_target = target_sid in visited_sids
        overdue = False
        try:
            tcl_dt = pd.to_datetime(target_cl, errors="coerce")
            if pd.notna(tcl_dt) and tcl_dt.date() < today_dt and "close" not in st_low:
                overdue = True
        except Exception:
            pass

        if "close" in st_low or "done" in st_low or "selesai" in st_low:
            analisa_target = "✅ Tercapai (Closed)"
        elif ha_recs_tgt and ("blocked" in ha_recs_tgt[-1]["status"].lower() or "tidak tembus" in ha_recs_tgt[-1]["status"].lower()):
            analisa_target = f"🚧 Hard Access Blocked ({ha_recs_tgt[-1]['kategori'][:42]})"
        elif visited_target:
            alasan_short = rem_hasil if rem_hasil != "-" else (waypoints[-1]["remark"] if waypoints else "On Progress di site target")
            analisa_target = f"🔄 On Target — Belum Closed ({alasan_short[:55]})"
        elif visited_sids:
            divert_str = " ➔ ".join(visited_sids)
            extra_note = f" | {notes_non_site[0]}" if notes_non_site else ""
            analisa_target = f"🔀 Dialihkan ke {divert_str} ({round(total_act_km, 1)} km){extra_note[:45]}"
        elif notes_non_site:
            analisa_target = f"🚧 Kendala Lapangan ({notes_non_site[0][:55]})"
        elif has_access_kw:
            analisa_target = f"🚧 Terkendala Akses ({rem_hasil[:48]})"
        elif overdue:
            analisa_target = "⚠️ Melewati Target — Belum Ada Tracker"
        else:
            analisa_target = "📅 Terjadwal / Waiting Eksekusi"

        rute_chain = f"HB ({tim1 or '-'})"
        if waypoints:
            rute_chain += " ➔ " + " ➔ ".join([f"T#{w['tier']}:{w['site_id']} ({w['leg_km']}km)" for w in waypoints])
        if notes_non_site:
            rute_chain += " ➔ [" + "; ".join(notes_non_site) + "]"
        if not waypoints and not notes_non_site:
            rute_chain = "Belum ada tarikan tracker"

        hist_sid = onsite_idx.get(_norm_site_id(target_sid), [])
        onsite_cnt = len(hist_sid)
        last_onsite_tgl = hist_sid[-1]["Tanggal"] if hist_sid else "-"
        last_onsite_pwr = "-"
        if hist_sid:
            lh = hist_sid[-1]
            parts_p = []
            if lh.get("V_AC") is not None: parts_p.append(f"VAC:{lh['V_AC']}V")
            if lh.get("V_DC") is not None: parts_p.append(f"VDC:{lh['V_DC']}V")
            if lh.get("Load_DC") is not None: parts_p.append(f"Load:{lh['Load_DC']}A")
            last_onsite_pwr = " | ".join(parts_p) if parts_p else "-"

        rec_sum = {
            "No": no_val,
            "SITE ID Target": target_sid,
            "SITE NAME": sname,
            "Class": s_class,
            "Status": status_act,
            "Status Hard Access": ha_status_lbl,
            "PIC Task": pic_task,
            "Tier Tim #1": tim1 or "-",
            "Tier Tim #2": tim2 or "-",
            "Tier Tim #3": tim3 or "-",
            "RCA": rca,
            "Action Detail": action_det,
            "Plan Date": plan_dt,
            "Target Closed": target_cl,
            "HB->Target (km)": dist_hb_tgt if dist_hb_tgt is not None else 0.0,
            "Est Jalan HB->Target (km)": est_road_tgt if est_road_tgt is not None else 0.0,
            "Total Jarak Aktual (km)": round(total_act_km, 2),
            "Est Jalan Aktual (km)": est_road_act,
            "Tarikan Rute Aktual": rute_chain,
            "Analisa Target / Divert": analisa_target,
            "Tier #1 Tgl": tr_flat["Tier #1 Tgl"],
            "Tier #1 Site/Info": tr_flat["Tier #1 Site/Info"],
            "Tier #1 Jarak (km)": tr_flat["Tier #1 Jarak (km)"],
            "Tier #1 Remark": tr_flat["Tier #1 Remark"],
            "Tier #2 Tgl": tr_flat["Tier #2 Tgl"],
            "Tier #2 Site/Info": tr_flat["Tier #2 Site/Info"],
            "Tier #2 Jarak (km)": tr_flat["Tier #2 Jarak (km)"],
            "Tier #2 Remark": tr_flat["Tier #2 Remark"],
            "Tier #3 Tgl": tr_flat["Tier #3 Tgl"],
            "Tier #3 Site/Info": tr_flat["Tier #3 Site/Info"],
            "Tier #3 Jarak (km)": tr_flat["Tier #3 Jarak (km)"],
            "Tier #3 Remark": tr_flat["Tier #3 Remark"],
            "Biaya (Rp)": biaya_val,
            "Remark Hasil Pekerjaan": rem_hasil,
            "Onsite SIDATA (x)": onsite_cnt,
            "Last Onsite SIDATA": last_onsite_tgl,
            "Last Power SIDATA": last_onsite_pwr
        }
        summary_rows.append(rec_sum)

        map_routes.append({
            "key": f"#{no_val} | {target_sid} — {sname} ({tim1 or pic_task})",
            "no": no_val,
            "target_sid": target_sid,
            "sname": sname,
            "s_class": s_class,
            "status": status_act,
            "hard_access_status": ha_status_lbl,
            "pic": pic_task,
            "tim1": tim1 or "-",
            "tim2": tim2 or "-",
            "tim3": tim3 or "-",
            "rca": rca,
            "action": action_det,
            "plan_dt": plan_dt,
            "target_cl": target_cl,
            "biaya": biaya_val,
            "rem_plan": rem_plan,
            "rem_hasil": rem_hasil,
            "analisa": analisa_target,
            "hb_lat": hb_coord[0] if hb_coord else None,
            "hb_lon": hb_coord[1] if hb_coord else None,
            "hb_name": hb_matched_name or tim1 or "-",
            "tgt_lat": tgt_coord[0] if tgt_coord else None,
            "tgt_lon": tgt_coord[1] if tgt_coord else None,
            "dist_hb_tgt": dist_hb_tgt,
            "est_road_tgt": est_road_tgt,
            "total_act_km": round(total_act_km, 2),
            "est_road_act": est_road_act,
            "waypoints": waypoints,
            "notes_non_site": notes_non_site
        })

    return pd.DataFrame(summary_rows), map_routes

def build_pbu_management_report_xlsx(df_summary, df_site_master, df_hb_master, df_hard_access=None):
    """Membangun laporan Excel multi-sheet (5 Sheet) berformat profesional untuk pengiriman ke Manajemen."""
    from openpyxl import Workbook
    from openpyxl.styles import Font, PatternFill, Alignment, Border, Side
    from openpyxl.utils import get_column_letter

    wb = Workbook()
    NAVY, SLATE, ZEBRA, GREEN_BG, AMBER_BG, RED_BG = "1E3A5F", "334155", "F8FAFC", "DCFCE7", "FEF3C7", "FEE2E2"
    thin = Side(style="thin", color="CBD5E1")
    brd = Border(left=thin, right=thin, top=thin, bottom=thin)

    ws1 = wb.active
    ws1.title = "Ringkasan_Manajemen"
    ws1.sheet_view.showGridLines = True

    ws1["A1"] = "LAPORAN EKSEKUTIF PROGRAM NOP, TIER TRACKING & HARD ACCESS — NOP PANGKALAN BUN"
    ws1["A1"].font = Font(name="Calibri", size=15, bold=True, color=NAVY)
    ws1["A2"] = f"Tanggal Generate: {datetime.datetime.now().strftime('%Y-%m-%d %H:%M')} | Sumber: SIDATA Program NOP + Master LongLat + Report Hard Access PBU"
    ws1["A2"].font = Font(name="Calibri", size=10, italic=True, color="64748B")

    tot_t = len(df_summary)
    cnt_cl = int(df_summary["Status"].astype(str).str.lower().str.contains("close|done|selesai").sum()) if tot_t else 0
    cnt_ip = int(df_summary["Status"].astype(str).str.lower().str.contains("progress").sum()) if tot_t else 0
    cnt_op = tot_t - cnt_cl - cnt_ip
    tot_by = float(pd.to_numeric(df_summary["Biaya (Rp)"], errors="coerce").fillna(0).sum()) if tot_t else 0.0
    tot_km_plan = float(pd.to_numeric(df_summary["HB->Target (km)"], errors="coerce").fillna(0).sum()) if tot_t else 0.0
    tot_km_act = float(pd.to_numeric(df_summary["Total Jarak Aktual (km)"], errors="coerce").fillna(0).sum()) if tot_t else 0.0
    cnt_div = int(df_summary["Analisa Target / Divert"].astype(str).str.contains("Dialihkan").sum()) if tot_t else 0
    cnt_ha = int(df_summary["Status Hard Access"].astype(str).str.contains("🚧|⚠️").sum()) if (tot_t and "Status Hard Access" in df_summary.columns) else 0
    tot_ha_rep = len(df_hard_access) if (df_hard_access is not None and not df_hard_access.empty) else 0

    kpi_rows = [
        ("Total Site Target Program", f"{tot_t} Site"),
        ("Status Closed (Tercapai)", f"{cnt_cl} Site ({(cnt_cl/tot_t*100) if tot_t else 0:.1f}%)"),
        ("Status In Progress", f"{cnt_ip} Site"),
        ("Status Open (Belum Closed)", f"{cnt_op} Site"),
        ("Task Dialihkan ke Site Emergency/Backup", f"{cnt_div} Task"),
        ("Site Program Terdampak Hard Access / Kendala Akses", f"{cnt_ha} Site (Total Report Hard Access PBU: {tot_ha_rep} Lokasi)"),
        ("Total Jarak Plan (HB -> Target, Garis Lurus)", f"{tot_km_plan:,.1f} km (Est. Jalan: {tot_km_plan*1.35:,.1f} km)"),
        ("Total Jarak Aktual Ditempuh Tim (Tier #1-#3)", f"{tot_km_act:,.1f} km (Est. Jalan: {tot_km_act*1.35:,.1f} km)"),
        ("Total Biaya Operasional Tercatat", f"Rp {tot_by:,.0f}")
    ]
    ws1["A4"] = "INDIKATOR KPI MANAJEMEN"; ws1["B4"] = "NILAI / CAPAIAN"
    for col_c in ("A4", "B4"):
        ws1[col_c].font = Font(bold=True, color="FFFFFF")
        ws1[col_c].fill = PatternFill("solid", fgColor=NAVY)
        ws1[col_c].border = brd
    for idx_k, (k_lbl, k_val) in enumerate(kpi_rows, 5):
        ws1[f"A{idx_k}"] = k_lbl; ws1[f"B{idx_k}"] = k_val
        ws1[f"A{idx_k}"].font = Font(bold=True)
        ws1[f"A{idx_k}"].border = brd; ws1[f"B{idx_k}"].border = brd
        if idx_k % 2 == 1:
            ws1[f"A{idx_k}"].fill = PatternFill("solid", fgColor=ZEBRA)
            ws1[f"B{idx_k}"].fill = PatternFill("solid", fgColor=ZEBRA)

    r_start = 16
    ws1[f"A{r_start}"] = "REKAP PERGERAKAN & BEBAN JARAK PER TIM (TIER TIM #1)"
    ws1[f"A{r_start}"].font = Font(size=12, bold=True, color=NAVY)
    if not df_summary.empty:
        grp_tim = df_summary.groupby("Tier Tim #1").agg(
            Jumlah_Task=("SITE ID Target", "count"),
            Closed=("Status", lambda x: int(x.astype(str).str.lower().str.contains("close|done|selesai").sum())),
            Total_Jarak_Plan_km=("HB->Target (km)", "sum"),
            Total_Jarak_Aktual_km=("Total Jarak Aktual (km)", "sum"),
            Est_Jalan_Aktual_km=("Est Jalan Aktual (km)", "sum"),
            Total_Biaya_Rp=("Biaya (Rp)", "sum")
        ).reset_index().sort_values("Total_Jarak_Aktual_km", ascending=False)
        hdrs_g = list(grp_tim.columns)
        for ci, h_txt in enumerate(hdrs_g, 1):
            c = ws1.cell(row=r_start + 1, column=ci, value=h_txt)
            c.font = Font(bold=True, color="FFFFFF"); c.fill = PatternFill("solid", fgColor=SLATE); c.border = brd
        for ri, (_, rw) in enumerate(grp_tim.iterrows(), r_start + 2):
            for ci, h_txt in enumerate(hdrs_g, 1):
                val = rw[h_txt]
                c = ws1.cell(row=ri, column=ci, value=round(val, 2) if isinstance(val, float) else val)
                c.border = brd
                if "Biaya" in h_txt: c.number_format = '#,##0'
                elif "km" in h_txt: c.number_format = '#,##0.0'

    ws1.column_dimensions["A"].width = 46
    ws1.column_dimensions["B"].width = 54
    for col_l in ("C", "D", "E", "F", "G"):
        ws1.column_dimensions[col_l].width = 22

    # Sheet 2: Detail Rute Tier & Jarak
    ws2 = wb.create_sheet("Detail_Rute_Tier_Jarak")
    cols_s2 = list(df_summary.columns)
    for ci, cname in enumerate(cols_s2, 1):
        c = ws2.cell(row=1, column=ci, value=cname)
        c.font = Font(bold=True, color="FFFFFF", size=10)
        c.fill = PatternFill("solid", fgColor=NAVY)
        c.alignment = Alignment(horizontal="center", vertical="center", wrap_text=True)
        c.border = brd
    ws2.row_dimensions[1].height = 28
    ws2.freeze_panes = "D2"

    for ri, (_, rw) in enumerate(df_summary.iterrows(), 2):
        st_txt = str(rw.get("Status", "")).lower()
        ha_txt = str(rw.get("Status Hard Access", ""))
        row_fill = GREEN_BG if "close" in st_txt else (RED_BG if "Blocked" in ha_txt else (AMBER_BG if "progress" in st_txt else (ZEBRA if ri % 2 == 0 else None)))
        for ci, cname in enumerate(cols_s2, 1):
            val = rw[cname]
            c = ws2.cell(row=ri, column=ci, value=val)
            c.border = brd
            c.alignment = Alignment(vertical="top", wrap_text=True)
            if row_fill:
                c.fill = PatternFill("solid", fgColor=row_fill)
            if "Biaya" in cname:
                c.number_format = '#,##0'
            elif "(km)" in cname and isinstance(val, (int, float)):
                c.number_format = '#,##0.00'
    for ci, cname in enumerate(cols_s2, 1):
        ws2.column_dimensions[get_column_letter(ci)].width = 30 if ("Remark" in cname or "Rute" in cname or "Analisa" in cname or "Action" in cname or "Hard Access" in cname) else 16

    # Sheet 3: Report Hard Access PBU
    if df_hard_access is not None and not df_hard_access.empty:
        ws3 = wb.create_sheet("Report_Hard_Access_PBU")
        ha_cols = list(df_hard_access.columns)
        for ci, cn in enumerate(ha_cols, 1):
            c = ws3.cell(row=1, column=ci, value=str(cn))
            c.font = Font(bold=True, color="FFFFFF"); c.fill = PatternFill("solid", fgColor="B91C1C"); c.border = brd
        for ri, (_, rw) in enumerate(df_hard_access.iterrows(), 2):
            for ci, cn in enumerate(ha_cols, 1):
                c = ws3.cell(row=ri, column=ci, value=rw[cn])
                c.border = brd
                c.alignment = Alignment(vertical="top", wrap_text=True)
        for ci, cn in enumerate(ha_cols, 1):
            ws3.column_dimensions[get_column_letter(ci)].width = 34 if ("Lokasi" in cn or "Keterangan" in cn or "Foto" in cn) else 20

    # Sheet 4 & 5: Master LongLat Snapshot
    for sh_nm, df_ref in [("Master_LongLat_Site", df_site_master), ("Master_HB_Tim", df_hb_master)]:
        if df_ref is not None and not df_ref.empty:
            ws_r = wb.create_sheet(sh_nm)
            r_cols = list(df_ref.columns)
            for ci, cn in enumerate(r_cols, 1):
                c = ws_r.cell(row=1, column=ci, value=str(cn))
                c.font = Font(bold=True, color="FFFFFF"); c.fill = PatternFill("solid", fgColor=SLATE); c.border = brd
            for ri, (_, rw) in enumerate(df_ref.iterrows(), 2):
                for ci, cn in enumerate(r_cols, 1):
                    ws_r.cell(row=ri, column=ci, value=rw[cn]).border = brd
            for ci in range(1, len(r_cols) + 1):
                ws_r.column_dimensions[get_column_letter(ci)].width = 24

    buf = BytesIO()
    wb.save(buf)
    return buf.getvalue()

def render_pbu_route_map_html(routes_list, hard_access_list=None, map_theme="Dark NOC Command Center", show_plan=True, show_actual=True, show_hard_access=True, show_dist_labels=True):
    """Peta Interaktif Premium Leaflet: Garis Melengkung (Bezier) + Panah Arah + Badge Jarak (km) + Pin Hard Access ber-Foto."""
    hard_access_list = hard_access_list or []
    if "Satelit" in map_theme:
        tile_url = "https://server.arcgisonline.com/ArcGIS/rest/services/World_Imagery/MapServer/tile/{z}/{y}/{x}"
        tile_sub = ""
    elif "Terang" in map_theme:
        tile_url = "https://{s}.basemaps.cartocdn.com/rastertiles/voyager/{z}/{x}/{y}{r}.png"
        tile_sub = "abcd"
    else:
        tile_url = "https://{s}.basemaps.cartocdn.com/dark_all/{z}/{x}/{y}{r}.png"
        tile_sub = "abcd"

    payload_json = json.dumps(routes_list, ensure_ascii=False)
    ha_json = json.dumps(hard_access_list, ensure_ascii=False)

    html_map = f"""
    <!DOCTYPE html>
    <html>
    <head>
        <meta charset="utf-8" />
        <link rel="stylesheet" href="https://unpkg.com/leaflet@1.9.4/dist/leaflet.css" />
        <script src="https://unpkg.com/leaflet@1.9.4/dist/leaflet.js"></script>
        <style>
            body {{ margin:0; padding:0; font-family:'Plus Jakarta Sans',system-ui,sans-serif; background:transparent; }}
            #pbumap {{ width:100%; height:600px; border-radius:16px; border:1.5px solid rgba(56,189,248,0.5); box-shadow:0 14px 34px rgba(0,0,0,0.28); overflow:hidden; }}
            .pin-hb {{
                width:34px; height:34px; border-radius:50%; background:linear-gradient(135deg,#2563eb,#1d4ed8);
                border:2.5px solid #fff; color:#fff; display:flex; align-items:center; justify-content:center;
                font-size:15px; box-shadow:0 0 16px rgba(37,99,235,0.9);
            }}
            .pin-tgt {{
                width:34px; height:34px; border-radius:50%; background:linear-gradient(135deg,#ef4444,#b91c1c);
                border:2.5px solid #fde047; color:#fff; display:flex; align-items:center; justify-content:center;
                font-size:15px; box-shadow:0 0 18px rgba(239,68,68,0.95); animation:pulseTgt 2s infinite;
            }}
            .pin-ha {{
                width:34px; height:34px; border-radius:10px; background:linear-gradient(135deg,#f97316,#dc2626);
                border:2.5px solid #fef08a; color:#fff; display:flex; align-items:center; justify-content:center;
                font-size:16px; box-shadow:0 0 18px rgba(249,115,22,0.95); animation:pulseHA 1.8s infinite;
            }}
            @keyframes pulseTgt {{
                0% {{ box-shadow:0 0 0 0 rgba(239,68,68,0.75); }}
                70% {{ box-shadow:0 0 0 14px rgba(239,68,68,0); }}
                100% {{ box-shadow:0 0 0 0 rgba(239,68,68,0); }}
            }}
            @keyframes pulseHA {{
                0% {{ box-shadow:0 0 0 0 rgba(249,115,22,0.8); }}
                70% {{ box-shadow:0 0 0 15px rgba(249,115,22,0); }}
                100% {{ box-shadow:0 0 0 0 rgba(249,115,22,0); }}
            }}
            .pin-tr {{
                width:30px; height:30px; border-radius:50%; border:2.5px solid #fff; color:#fff;
                font-size:11.5px; font-weight:900; display:flex; align-items:center; justify-content:center;
                box-shadow:0 4px 12px rgba(0,0,0,0.6);
            }}
            .pin-lbl {{
                position:absolute; top:35px; left:50%; transform:translateX(-50%);
                background:rgba(15,23,42,0.92); color:#f8fafc; font-size:9.5px; font-weight:800;
                padding:2px 7px; border-radius:5px; white-space:nowrap; border:1px solid rgba(255,255,255,0.22);
                pointer-events:none; box-shadow:0 2px 6px rgba(0,0,0,0.4);
            }}
            .dist-pill {{
                background:rgba(15,23,42,0.92); color:#fff; font-size:10px; font-weight:800;
                padding:2px 7px; border-radius:99px; white-space:nowrap; border:1.5px solid #38bdf8;
                box-shadow:0 2px 8px rgba(0,0,0,0.45); pointer-events:none; display:inline-flex; align-items:center; gap:3px;
            }}
            .arrow-badge {{
                width:20px; height:20px; display:flex; align-items:center; justify-content:center;
                pointer-events:none;
            }}
            @keyframes dashFlow {{
                from {{ stroke-dashoffset: 36; }}
                to {{ stroke-dashoffset: 0; }}
            }}
            .flow-line {{
                animation: dashFlow 1.6s linear infinite;
            }}
            .leaflet-popup-content-wrapper {{
                background:linear-gradient(145deg,#0f172a,#1e293b) !important;
                color:#f8fafc !important; border-radius:14px !important; border:1px solid rgba(56,189,248,0.48);
                max-width:320px;
            }}
            .leaflet-popup-tip {{ background:#1e293b !important; }}
            .pc {{ min-width:245px; font-size:12px; line-height:1.55; }}
            .pc-h {{ font-size:13.5px; font-weight:800; color:#38bdf8; border-bottom:1px solid rgba(255,255,255,0.14); padding-bottom:5px; margin-bottom:6px; }}
            .pc-h-ha {{ font-size:13.5px; font-weight:800; color:#fb923c; border-bottom:1px solid rgba(255,255,255,0.14); padding-bottom:5px; margin-bottom:6px; }}
            .pc-rmk {{ margin-top:6px; padding:6px 8px; border-radius:8px; background:rgba(255,255,255,0.08); border-left:3px solid #f59e0b; }}
            .pc-photos {{ display:grid; grid-template-columns:repeat(2,1fr); gap:5px; margin-top:7px; }}
            .pc-photos img {{ width:100%; height:68px; object-fit:cover; border-radius:6px; border:1px solid rgba(255,255,255,0.22); }}
            .pc-btn {{ display:block; text-align:center; margin-top:8px; padding:6px; border-radius:7px; background:#2563eb; color:#fff !important; text-decoration:none; font-weight:800; font-size:11px; }}
            .map-leg {{
                background:rgba(15,23,42,0.92); color:#f8fafc; padding:10px 13px; border-radius:12px;
                border:1px solid rgba(255,255,255,0.2); font-size:11px; line-height:1.7;
            }}
            .line-dash {{ display:inline-block; width:22px; height:0; border-top:3px dashed #f59e0b; vertical-align:middle; margin-right:6px; }}
            .line-act {{ display:inline-block; width:22px; height:0; border-top:3.5px solid #06b6d4; vertical-align:middle; margin-right:6px; }}
        </style>
    </head>
    <body>
        <div id="pbumap"></div>
        <script>
            const routes = {payload_json};
            const hardAccess = {ha_json};
            const showPlan = {"true" if show_plan else "false"};
            const showActual = {"true" if show_actual else "false"};
            const showHardAccess = {"true" if show_hard_access else "false"};
            const showDistLabels = {"true" if show_dist_labels else "false"};
            const tierColors = {{ 1: "#06b6d4", 2: "#a855f7", 3: "#10b981" }};

            const map = L.map('pbumap', {{ center: [-2.35, 111.6], zoom: 8 }});
            L.tileLayer('{tile_url}', {{ subdomains: '{tile_sub}', maxZoom: 19, attribution: '⚡ SIDATA PBU Tier & Hard Access Maps' }}).addTo(map);

            const bounds = [];

            function makeIcon(cls, txt, lbl, bg) {{
                const styleBg = bg ? `background:${{bg}};` : '';
                return L.divIcon({{
                    className: '',
                    iconSize: [34, 34],
                    iconAnchor: [17, 17],
                    popupAnchor: [0, -18],
                    html: `<div style="position:relative;"><div class="${{cls}}" style="${{styleBg}}">${{txt}}</div><div class="pin-lbl">${{lbl}}</div></div>`
                }});
            }}

            // Generator kurva Quadratic Bezier mulus agar garis Plan & Aktual tidak bertumpuk
            function makeCurvedPoints(p1, p2, bend) {{
                const lat1 = p1[0], lon1 = p1[1], lat2 = p2[0], lon2 = p2[1];
                const dLat = lat2 - lat1, dLon = lon2 - lon1;
                const dist = Math.sqrt(dLat * dLat + dLon * dLon);
                if (dist < 0.0005) return [p1, p2];
                const midLat = (lat1 + lat2) / 2;
                const midLon = (lon1 + lon2) / 2;
                // Vektor tegak lurus untuk kelengkungan
                const ctrlLat = midLat - dLon * bend;
                const ctrlLon = midLon + dLat * bend;
                const pts = [];
                const steps = 28;
                for (let i = 0; i <= steps; i++) {{
                    const t = i / steps;
                    const inv = 1 - t;
                    const la = inv * inv * lat1 + 2 * inv * t * ctrlLat + t * t * lat2;
                    const lo = inv * inv * lon1 + 2 * inv * t * ctrlLon + t * t * lon2;
                    pts.push([la, lo]);
                }}
                return pts;
            }}

            function addArrowAndLabel(pts, color, labelTxt) {{
                if (!pts || pts.length < 4) return;
                const midIdx = Math.floor(pts.length / 2);
                const pA = pts[Math.max(0, midIdx - 1)];
                const pB = pts[Math.min(pts.length - 1, midIdx + 1)];
                const angleRad = Math.atan2(pB[1] - pA[1], pB[0] - pA[0]);
                const deg = (angleRad * 180 / Math.PI);

                // Panah arah di posisi 65% segmen
                const arrIdx = Math.floor(pts.length * 0.65);
                const pa1 = pts[Math.max(0, arrIdx - 1)];
                const pa2 = pts[Math.min(pts.length - 1, arrIdx + 1)];
                const arrDeg = Math.atan2(pa2[1] - pa1[1], pa2[0] - pa1[0]) * 180 / Math.PI;
                const arrowIcon = L.divIcon({{
                    className: '',
                    iconSize: [20, 20],
                    iconAnchor: [10, 10],
                    html: `<div class="arrow-badge" style="transform:rotate(${{arrDeg}}deg);"><svg width="16" height="16" viewBox="0 0 16 16"><path d="M8 1 L14 13 L8 10 L2 13 Z" fill="${{color}}" stroke="#0f172a" stroke-width="1.2"/></svg></div>`
                }});
                L.marker(pts[arrIdx], {{ icon: arrowIcon, interactive: false }}).addTo(map);

                if (showDistLabels && labelTxt) {{
                    const lblIcon = L.divIcon({{
                        className: '',
                        iconSize: [0, 0],
                        html: `<div class="dist-pill" style="border-color:${{color}}; transform:translate(-50%,-50%);">${{labelTxt}}</div>`
                    }});
                    L.marker(pts[midIdx], {{ icon: lblIcon, interactive: false }}).addTo(map);
                }}
            }}

            routes.forEach((rt, rIdx) => {{
                const hasHB = rt.hb_lat !== null && rt.hb_lon !== null;
                const hasTgt = rt.tgt_lat !== null && rt.tgt_lon !== null;

                if (hasHB) {{
                    bounds.push([rt.hb_lat, rt.hb_lon]);
                    const mHB = L.marker([rt.hb_lat, rt.hb_lon], {{ icon: makeIcon('pin-hb', '🏠', 'HB: ' + rt.tim1) }}).addTo(map);
                    mHB.bindPopup(`
                        <div class="pc">
                            <div class="pc-h">🏠 HOMEBASE TIM: ${{rt.tim1}}</div>
                            <div>👷 <b>PIC Task:</b> ${{rt.pic}}</div>
                            <div>📍 <b>Koordinat HB:</b> ${{rt.hb_lat.toFixed(5)}}, ${{rt.hb_lon.toFixed(5)}}</div>
                            <div>🎯 <b>Site Target:</b> ${{rt.target_sid}} (${{rt.sname}})</div>
                            <div>📏 <b>Jarak ke Target:</b> ${{rt.dist_hb_tgt !== null ? rt.dist_hb_tgt + ' km (Est. Jalan ' + rt.est_road_tgt + ' km)' : '-'}}</div>
                            <div>🚗 <b>Total Jarak Aktual:</b> ${{rt.total_act_km}} km (Est. Jalan ${{rt.est_road_act}} km)</div>
                        </div>
                    `);
                }}

                if (hasTgt) {{
                    bounds.push([rt.tgt_lat, rt.tgt_lon]);
                    const mTgt = L.marker([rt.tgt_lat, rt.tgt_lon], {{ icon: makeIcon('pin-tgt', '🎯', 'TGT: ' + rt.target_sid) }}).addTo(map);
                    mTgt.bindPopup(`
                        <div class="pc">
                            <div class="pc-h">🎯 TARGET #${{rt.no}}: ${{rt.target_sid}} — ${{rt.sname}}</div>
                            <div>⭐ <b>Class:</b> ${{rt.s_class}} | <b>Status:</b> ${{rt.status}}</div>
                            <div>🚧 <b>Akses Site:</b> ${{rt.hard_access_status || 'Normal'}}</div>
                            <div>🔧 <b>RCA / Action:</b> ${{rt.rca}} — ${{rt.action}}</div>
                            <div>📅 <b>Plan / Target Closed:</b> ${{rt.plan_dt}} ➔ ${{rt.target_cl}}</div>
                            <div>📏 <b>Jarak dari HB (${{rt.tim1}}):</b> ${{rt.dist_hb_tgt !== null ? rt.dist_hb_tgt + ' km (Est. Jalan ' + rt.est_road_tgt + ' km)' : '-'}}</div>
                            <div>💰 <b>Biaya:</b> Rp ${{Number(rt.biaya || 0).toLocaleString('id-ID')}}</div>
                            <div class="pc-rmk"><b>🚦 Analisa:</b> ${{rt.analisa}}<br><b>📝 Hasil:</b> ${{rt.rem_hasil}}</div>
                            <a class="pc-btn" href="https://www.google.com/maps?q=${{rt.tgt_lat}},${{rt.tgt_lon}}" target="_blank">🧭 Buka Google Maps</a>
                        </div>
                    `);
                }}

                // 1. Garis Plan Melengkung (Sisi kiri: bend = -0.16)
                if (showPlan && hasHB && hasTgt) {{
                    const planPts = makeCurvedPoints([rt.hb_lat, rt.hb_lon], [rt.tgt_lat, rt.tgt_lon], -0.16);
                    // Glow bawah
                    L.polyline(planPts, {{ color: '#f59e0b', weight: 7, opacity: 0.16 }}).addTo(map);
                    const planLine = L.polyline(planPts, {{
                        color: '#f59e0b', weight: 3, opacity: 0.9, dashArray: '10, 8'
                    }}).addTo(map);
                    planLine.bindTooltip(`📐 PLAN: HB ${{rt.tim1}} ➔ ${{rt.target_sid}} | ${{rt.dist_hb_tgt}} km (Est. Jalan ${{rt.est_road_tgt}} km)`, {{ sticky: true }});
                    addArrowAndLabel(planPts, '#f59e0b', routes.length <= 5 ? `📐 ${{rt.dist_hb_tgt}} km` : '');
                }}

                // 2. Garis Aktual Melengkung + Flow Animation (HB -> T#1 -> T#2 -> T#3)
                if (showActual && rt.waypoints && rt.waypoints.length > 0) {{
                    let prevPt = hasHB ? [rt.hb_lat, rt.hb_lon] : null;
                    rt.waypoints.forEach((wp, wIdx) => {{
                        bounds.push([wp.lat, wp.lon]);
                        const colT = tierColors[wp.tier] || '#06b6d4';

                        if (prevPt) {{
                            const bendVal = (wIdx % 2 === 0) ? 0.18 : -0.18;
                            const segPts = makeCurvedPoints(prevPt, [wp.lat, wp.lon], bendVal);
                            // Glow neon
                            L.polyline(segPts, {{ color: colT, weight: 9, opacity: 0.22 }}).addTo(map);
                            // Garis solid utama
                            const segLine = L.polyline(segPts, {{ color: colT, weight: 4, opacity: 0.95 }}).addTo(map);
                            // Garis animasi aliran arah
                            L.polyline(segPts, {{ color: '#ffffff', weight: 1.8, opacity: 0.65, dashArray: '6, 12', className: 'flow-line' }}).addTo(map);
                            segLine.bindTooltip(`🚗 TIER #${{wp.tier}} (${{wp.tgl}}): ${{wp.from_label}} ➔ ${{wp.site_id}} | ${{wp.leg_km}} km (Est. Jalan ${{wp.est_road_km}} km) | ${{wp.remark}}`, {{ sticky: true }});
                            addArrowAndLabel(segPts, colT, routes.length <= 6 ? `T${{wp.tier}}: ${{wp.leg_km}} km` : '');
                        }}

                        const mWp = L.marker([wp.lat, wp.lon], {{
                            icon: makeIcon('pin-tr', 'T' + wp.tier, `T#${{wp.tier}}: ${{wp.site_id}} (${{wp.leg_km}}km)`, colT)
                        }}).addTo(map);
                        mWp.bindPopup(`
                            <div class="pc">
                                <div class="pc-h">📍 KUNJUNGAN TIER #${{wp.tier}} — ${{wp.site_id}}</div>
                                <div>📅 <b>Tanggal Kunjungan:</b> ${{wp.tgl}}</div>
                                <div>👷 <b>Tim (Task #${{rt.no}} ${{rt.target_sid}}):</b> ${{rt.tim1}}</div>
                                <div>📏 <b>Jarak dari ${{wp.from_label}}:</b> ${{wp.leg_km}} km (Est. Jalan ${{wp.est_road_km}} km)</div>
                                <div class="pc-rmk"><b>🔧 Dikerjakan / Remark Tier #${{wp.tier}}:</b><br>${{wp.remark}}</div>
                                <a class="pc-btn" href="https://www.google.com/maps?q=${{wp.lat}},${{wp.lon}}" target="_blank">🧭 Buka Google Maps</a>
                            </div>
                        `);
                        prevPt = [wp.lat, wp.lon];
                    }});
                }}
            }});

            // 3. Layer Pin Hard Access Pangkalan Bun (+ Preview Foto langsung di Popup!)
            if (showHardAccess && hardAccess && hardAccess.length > 0) {{
                hardAccess.forEach(ha => {{
                    if (ha.lat === null || ha.lon === null) return;
                    bounds.push([ha.lat, ha.lon]);
                    const isNormal = (ha.status || '').toLowerCase().includes('normal') || (ha.status || '').toLowerCase().includes('tembus') && !(ha.status || '').toLowerCase().includes('tidak');
                    const haBg = isNormal ? 'linear-gradient(135deg,#10b981,#047857)' : 'linear-gradient(135deg,#f97316,#dc2626)';
                    const mHA = L.marker([ha.lat, ha.lon], {{
                        icon: makeIcon('pin-ha', '🚧', 'HA: ' + ha.site_id, haBg)
                    }}).addTo(map);

                    let photoHtml = '';
                    if (ha.photos && ha.photos.length > 0) {{
                        photoHtml = '<div class="pc-photos">' + ha.photos.slice(0, 4).map(u => {{
                            const thumb = u.includes('cloudinary') ? u.replace('upload/', 'upload/w_320,q_auto:eco/') : u;
                            return `<a href="${{u}}" target="_blank"><img src="${{thumb}}" loading="lazy"></a>`;
                        }}).join('') + '</div>';
                    }}

                    mHA.bindPopup(`
                        <div class="pc">
                            <div class="pc-h-ha">🚧 HARD ACCESS: ${{ha.site_id}} (${{ha.site_name}})</div>
                            <div>📍 <b>Lokasi Kendala:</b> ${{ha.lokasi}}</div>
                            <div>🏷️ <b>Kategori:</b> ${{ha.kategori}}</div>
                            <div>⚠️ <b>Tingkat / Status:</b> ${{ha.tingkat}} | <b>${{ha.status}}</b></div>
                            <div>🚙 <b>Kendaraan Wajib:</b> ${{ha.kendaraan}} | ⏱️ ${{ha.estimasi}}</div>
                            <div class="pc-rmk"><b>📝 Detail Akses:</b> ${{ha.keterangan}}<br><span style="opacity:0.75;font-size:10.5px;">Input: ${{ha.pic}} (${{ha.timestamp}})</span></div>
                            ${{photoHtml}}
                            <a class="pc-btn" style="background:#ea580c;" href="https://www.google.com/maps?q=${{ha.lat}},${{ha.lon}}" target="_blank">🧭 Navigasi Titik Hard Access</a>
                        </div>
                    `);
                }});
            }}

            if (bounds.length > 0) {{
                map.fitBounds(bounds, {{ padding: [48, 48], maxZoom: 12 }});
            }}

            const leg = L.control({{ position: 'bottomright' }});
            leg.onAdd = function() {{
                const div = L.DomUtil.create('div', 'map-leg');
                div.innerHTML = `
                    <b>🗺️ Legenda Tier & Hard Access Maps (PBU)</b><br>
                    🏠 Homebase (HB) Tim Tier #1 &nbsp;|&nbsp; 🎯 Site ID Target<br>
                    🚧 Titik Hard Access (Klik untuk Foto Lokasi)<br>
                    <span class="line-dash"></span> Rute Plan Melengkung (HB ➔ Target)<br>
                    <span class="line-act"></span> Rute Aktual + Panah Arah (HB ➔ T1 ➔ T2 ➔ T3)<br>
                    <b style="color:#06b6d4">● T1</b> &nbsp; <b style="color:#a855f7">● T2</b> &nbsp; <b style="color:#10b981">● T3</b> Kunjungan per Tier
                `;
                return div;
            }};
            leg.addTo(map);
        </script>
    </body>
    </html>
    """
    return html_map

# =============================================================================
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
    from PIL import Image, ImageOps
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
def sv_rekap_xlsx(recs):
    rows = [{"No. RAB": r.get("rab_id"), "Tiket": r.get("ticket_no"), "NOP": r.get("nop"), "Tgl Pengajuan": r.get("tgl_pengajuan"),
             "Pembuat": r.get("nama_pembuat"), "Status": r.get("status"), "Total ke Telkomsel": _sv_num(r.get("tot_tsel")),
             "Total Revenue KUT": _sv_num(r.get("tot_rev")), "Total HPP Lapangan": _sv_num(r.get("tot_hpp")), "Margin": _sv_num(r.get("margin")),
             "GPM": _sv_num(r.get("gpm")), "Margin Bersih": _sv_num(r.get("net")), "SPV": r.get("spv_nama"), "Manager": r.get("mgr_nama")} for r in recs]
    buf = _io.BytesIO()
    with pd.ExcelWriter(buf, engine="openpyxl") as xw:
        pd.DataFrame(rows).to_excel(xw, sheet_name="Rekap SiSAVA", index=False)
    return buf.getvalue()
    
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
    def add_rabs_batch(self, recs):
        if not recs: return 0
        d = self._read()
        for rec in recs:
            d["rab"].append({k: rec.get(k, "") for k in SV_RAB_COLS})
        self._write(d)
        return len(recs)
    def update_rab(self, rab_id, fields):
        d = self._read()
        for r in d["rab"]:
            if r.get("rab_id") == rab_id:
                r.update({k: v for k, v in fields.items() if k in SV_RAB_COLS}); self._write(d); return True
        return False
    def update_rabs_batch(self, updates_map):
        if not updates_map: return 0
        d = self._read(); cnt = 0
        for r in d["rab"]:
            rid = r.get("rab_id")
            if rid in updates_map:
                r.update({k: v for k, v in updates_map[rid].items() if k in SV_RAB_COLS})
                cnt += 1
        if cnt: self._write(d)
        return cnt
    def upsert_user(self, u):
        d = self._read()
        for r in d["users"]:
            if r.get("username") == u["username"]:
                r.update(u); self._write(d); return
        d["users"].append({k: u.get(k, "") for k in SV_USER_COLS}); self._write(d)

def _sv_spreadsheet():
    return get_cached_spreadsheet()

@st.cache_resource(show_spinner=False, ttl=3600)
def _sv_worksheet_and_head(name, cols_key):
    sh = get_cached_spreadsheet()
    cols = list(cols_key)
    try:
        ws = _gsheet_retry(sh.worksheet, name)
    except Exception as e_w:
        if _is_quota_429(e_w):
            raise
        ws = _gsheet_retry(sh.add_worksheet, title=name, rows=500, cols=max(len(cols), 10) + 2)
        _gsheet_retry(ws.update, range_name="A1", values=[cols])
        return ws, cols
    head = _gsheet_retry(ws.row_values, 1)
    if not head or not any(str(c).strip() for c in head):
        _gsheet_retry(ws.update, range_name="A1", values=[cols])
        return ws, cols
    missing = [c for c in cols if c not in head]
    if missing:
        head = head + missing
        _gsheet_retry(ws.update, range_name="A1", values=[head])
    return ws, head

def _sv_worksheet(name, cols_key):
    ws, _ = _sv_worksheet_and_head(name, cols_key)
    return ws

@st.cache_data(ttl=300, show_spinner=False)
def _sv_rows(name, cols_key):
    try:
        ws, _ = _sv_worksheet_and_head(name, cols_key)
        recs = _gsheet_retry(ws.get_all_records)
        _SV_LAST_OK[name] = recs
        st.session_state.pop("sv_quota_warn", None)
        return recs
    except Exception as e:
        if _is_quota_429(e) and name in _SV_LAST_OK:
            st.session_state["sv_quota_warn"] = "⚠️ Kuota baca Google Sheets sempat padat (429). Menampilkan data cache terakhir — klik 🔄 Refresh 30 detik lagi."
            return _SV_LAST_OK[name]
        raise

class _SvSheetsStore:
    kind = "Google Sheets"
    def _cols(self, name): return tuple(SV_RAB_COLS if name == "SISAVA_RAB" else SV_USER_COLS)
    def _rows(self, name): return _sv_rows(name, self._cols(name))
    def list_rabs(self): return self._rows("SISAVA_RAB")
    def list_users(self): return self._rows("SISAVA_USERS")
    def add_rab(self, rec):
        ws, head = _sv_worksheet_and_head("SISAVA_RAB", self._cols("SISAVA_RAB"))
        vals = [str(rec.get(c, "")) if rec.get(c, "") is not None else "" for c in head]
        _gsheet_retry(ws.append_row, vals, value_input_option="RAW")
        _sv_rows.clear()
    def add_rabs_batch(self, recs):
        if not recs: return 0
        ws, head = _sv_worksheet_and_head("SISAVA_RAB", self._cols("SISAVA_RAB"))
        rows_payload = [[str(rec.get(c, "")) if rec.get(c, "") is not None else "" for c in head] for rec in recs]
        _gsheet_retry(ws.append_rows, rows_payload, value_input_option="RAW")
        _sv_rows.clear()
        return len(rows_payload)
    def update_rabs_batch(self, updates_map):
        """Memperbarui banyak RAB sekaligus hanya dengan 1x read + 1x batch_update (anti-429!)."""
        if not updates_map: return 0
        ws, head = _sv_worksheet_and_head("SISAVA_RAB", self._cols("SISAVA_RAB"))
        all_vals = _gsheet_retry(ws.get_all_values)
        if not all_vals: return 0
        head_live = all_vals[0] if any(str(c).strip() for c in all_vals[0]) else head
        batch_data = []
        updated_cnt = 0
        for row_idx, row_vals in enumerate(all_vals[1:], start=2):
            rid = str(row_vals[0]).strip() if row_vals else ""
            if rid in updates_map:
                cur = list(row_vals) + [""] * max(0, len(head_live) - len(row_vals))
                for k, v in updates_map[rid].items():
                    if k in head_live:
                        cur[head_live.index(k)] = "" if v is None else str(v)
                batch_data.append({"range": f"A{row_idx}", "values": [cur]})
                updated_cnt += 1
        if batch_data:
            _gsheet_retry(ws.batch_update, batch_data, value_input_option="RAW")
            _sv_rows.clear()
        return updated_cnt
    def update_rab(self, rab_id, fields):
        return self.update_rabs_batch({rab_id: fields}) > 0
    def upsert_user(self, u):
        ws, head = _sv_worksheet_and_head("SISAVA_USERS", self._cols("SISAVA_USERS"))
        all_vals = _gsheet_retry(ws.get_all_values)
        head_live = all_vals[0] if (all_vals and any(str(c).strip() for c in all_vals[0])) else head
        target_row = None
        target_vals = None
        uname_l = str(u["username"]).strip().lower()
        for row_idx, row_vals in enumerate(all_vals[1:], start=2):
            if row_vals and str(row_vals[0]).strip().lower() == uname_l:
                target_row = row_idx
                target_vals = row_vals
                break
        if target_row is not None:
            cur = list(target_vals) + [""] * max(0, len(head_live) - len(target_vals))
            row = dict(zip(head_live, cur))
            row.update({k: v for k, v in u.items() if k in head_live})
            _gsheet_retry(ws.update, range_name=f"A{target_row}", values=[[str(row.get(c, "")) for c in head_live]], value_input_option="RAW")
        else:
            _gsheet_retry(ws.append_row, [str(u.get(c, "")) for c in head_live], value_input_option="RAW")
        _sv_rows.clear()

def sv_store():
    try:
        if "gcp_json" not in st.secrets:
            return _SvLocalStore()
        get_cached_spreadsheet()
        return _SvSheetsStore()
    except Exception as e:
        if _is_quota_429(e):
            # Jangan hapus cache resource saat 429; tetap pakai _SvSheetsStore agar fallback cache SiSAVA jalan
            return _SvSheetsStore()
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

_SV_AKSI = {"diajukan": "Diajukan", "revisi_diajukan": "Diajukan ulang (revisi)", "spv_approve": "Disetujui SPV", "spv_reject": "Ditolak SPV",
            "mgr_approve": "Disetujui Manager", "mgr_reject": "Ditolak Manager", "sva_fill": "Nilai SVA diisi", "kerja_upload": "Bukti pekerjaan diunggah",
            "f_spv_approve": "Pekerjaan disetujui SPV", "f_spv_reject": "Pekerjaan ditolak SPV", "f_mgr_approve": "Pekerjaan disetujui Manager (SAH)",
            "f_mgr_reject": "Pekerjaan ditolak Manager"}
        
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


# =============================================================================
# 📥 MODUL IMPORT EXCEL & INPUT ALL MASSAL SiSAVA -> GOOGLE SHEETS (SISAVA_RAB)
# =============================================================================
SV_ALL_STATUSES = [
    S_RAB_SPV, S_RAB_MGR, S_RAB_TOLAK, S_RAB_OK, S_SVA,
    S_KRJ_SPV, S_KRJ_MGR, S_KRJ_TOLAK, S_DONE
]

def sv_build_template_input_all_xlsx():
    """Membuat file Template Excel siap pakai untuk Input All / Import Massal SiSAVA."""
    from openpyxl import Workbook
    from openpyxl.styles import Font, PatternFill, Alignment, Border, Side
    from openpyxl.utils import get_column_letter

    wb = Workbook()
    ws = wb.active
    ws.title = "Input_All_SiSAVA"
    NAVY, SUB_BG, EX_BG = "1E3A5F", "E2E8F0", "F8FAFC"
    thin = Side(style="thin", color="CBD5E1")
    brd = Border(left=thin, right=thin, top=thin, bottom=thin)

    headers = [
        "No. Ticket", "NOP", "Area", "Kota", "Tanggal Pengajuan",
        "Site ID", "Site Name", "Pekerjaan", "Jenis (Material/Jasa)",
        "Qty", "UoM", "Harga 3E (Tsel)", "Harga KUT (Opsional)",
        "HPP Lapangan (Opsional)", "Biaya Lain (Opsional)",
        "Status SiSAVA", "Nilai SVA (Opsional)", "Keperluan Dana (Opsional)",
        "Urgensi", "Rencana Tgl Kerja", "Catatan"
    ]
    for ci, h_txt in enumerate(headers, 1):
        c = ws.cell(row=1, column=ci, value=h_txt)
        c.font = Font(name="Calibri", size=11, bold=True, color="FFFFFF")
        c.fill = PatternFill("solid", fgColor=NAVY)
        c.alignment = Alignment(horizontal="center", vertical="center", wrap_text=True)
        c.border = brd
    ws.row_dimensions[1].height = 30
    ws.freeze_panes = "C2"

    today_iso = _dtm.date.today().isoformat()
    sample_rows = [
        [
            "SVA-2026-000101", "PALANGKA RAYA", "KALIMANTAN", "Palangka Raya", today_iso,
            "PLK025", "PALANGKA_KOTA", "Penggantian Kabel Power NYY 4x16 + Instalasi", "Material",
            1, "lot", 15000000, "", "", 0,
            S_RAB_SPV, "", "", "Urgent", today_iso, "Contoh 1: Harga KUT & HPP kosong -> dihitung otomatis (92% & 85%)"
        ],
        [
            "SVA-2026-000102", "PANGKALAN BUN", "KALIMANTAN", "Pangkalan Bun", today_iso,
            "PBU067", "MELATA", "Jasa Troubleshoot & Replace Modul Rectifier", "Jasa",
            1, "ls", 8500000, 7225000, 5800000, 250000,
            S_SVA, 8500000, 5800000, "Emergency", today_iso, "Contoh 2: Nilai KUT, HPP & SVA diisi manual"
        ]
    ]
    for ri, row_vals in enumerate(sample_rows, 2):
        for ci, val in enumerate(row_vals, 1):
            c = ws.cell(row=ri, column=ci, value=val)
            c.fill = PatternFill("solid", fgColor=EX_BG)
            c.border = brd
            if ci in (12, 13, 14, 15, 17, 18) and isinstance(val, (int, float)):
                c.number_format = "#,##0"

    for ci, h_txt in enumerate(headers, 1):
        ws.column_dimensions[get_column_letter(ci)].width = 32 if ("Pekerjaan" in h_txt or "Catatan" in h_txt) else (22 if "Status" in h_txt or "Ticket" in h_txt else 17)

    # Sheet Petunjuk
    ws_p = wb.create_sheet("Petunjuk_Pengisian")
    guide = [
        ("PETUNJUK INPUT ALL EXCEL SiSAVA", ""),
        ("1. Kolom Wajib", "Minimal isi 'No. Ticket' (atau 'Site ID'), 'Pekerjaan', dan 'Harga 3E (Tsel)'."),
        ("2. Perhitungan Otomatis", "Jika 'Harga KUT' dikosongkan -> otomatis Harga 3E x 0.92 (Material) atau x 0.85 (Jasa). Jika 'HPP Lapangan' dikosongkan -> otomatis 85% dari Harga KUT. Margin, GPM, Cost of Money & PPh dihitung otomatis oleh sistem."),
        ("3. Multi-Item / Multi-Site", "Boleh mengisi banyak baris dengan 'No. Ticket' yang sama jika ingin digabungkan menjadi 1 dokumen RAB (pilih mode 'Gabung per No. Ticket' saat upload)."),
        ("4. Pilihan Status SiSAVA", " | ".join(SV_ALL_STATUSES)),
        ("5. Pilihan NOP", "PALANGKA RAYA | PANGKALAN BUN"),
        ("6. Format Excel RAB Resmi", "Anda juga dapat langsung mengunggah file Excel RAB hasil export SiSAVA ('RENCANA ANGGARAN BIAYA') tanpa perlu mengubah format!")
    ]
    for ri, (k_g, v_g) in enumerate(guide, 1):
        ws_p.cell(row=ri, column=1, value=k_g).font = Font(bold=True, color=NAVY if ri == 1 else "000000", size=13 if ri == 1 else 11)
        ws_p.cell(row=ri, column=2, value=v_g)
    ws_p.column_dimensions["A"].width = 28
    ws_p.column_dimensions["B"].width = 95

    buf = _io.BytesIO()
    wb.save(buf)
    return buf.getvalue()

def _sv_parse_num_flex(v):
    if v is None or (isinstance(v, float) and v != v):
        return None
    if isinstance(v, (int, float)) and not isinstance(v, bool):
        return float(v)
    s = str(v).strip()
    if not s or s.lower() in ("nan", "none", "-", ""):
        return None
    s_clean = re.sub(r"(?i)rp\.?\s*", "", s).replace(" ", "")
    # Tangani format pemisah ribuan Indonesia (1.500.000) vs desimal (1500000.50)
    if "," in s_clean and "." in s_clean:
        if s_clean.rfind(",") > s_clean.rfind("."):
            s_clean = s_clean.replace(".", "").replace(",", ".")
        else:
            s_clean = s_clean.replace(",", "")
    elif s_clean.count(".") > 1:
        s_clean = s_clean.replace(".", "")
    elif s_clean.count(",") > 1:
        s_clean = s_clean.replace(",", "")
    elif "," in s_clean:
        parts = s_clean.split(",")
        s_clean = s_clean.replace(",", "") if len(parts[-1]) == 3 else s_clean.replace(",", ".")
    try:
        return float(s_clean)
    except Exception:
        return None

def _sv_detect_official_rab_sheet(df_raw):
    """Deteksi apakah sheet Excel adalah format RAB resmi SiSAVA (RENCANA ANGGARAN BIAYA)."""
    if df_raw is None or df_raw.empty or df_raw.shape[1] < 13:
        return False
    top_text = " ".join(str(x).upper() for x in df_raw.iloc[:8, :15].values.flatten() if pd.notna(x))
    return ("RENCANA ANGGARAN BIAYA" in top_text) and ("HARGA 3E" in top_text or "TOTAL REVENUE" in top_text or "NILAI HPP" in top_text)

def sv_parse_official_rab_excel(df_raw, default_nop="PALANGKA RAYA", default_area="KALIMANTAN", default_kota="Palangka Raya"):
    """Parser khusus untuk Excel RAB resmi SiSAVA (format 21 kolom A..U)."""
    nop, ticket, kota = default_nop, "", default_kota
    for r_i in range(min(8, len(df_raw))):
        a_val = str(df_raw.iat[r_i, 0]).strip().upper() if pd.notna(df_raw.iat[r_i, 0]) else ""
        c_val = str(df_raw.iat[r_i, 2]).strip() if df_raw.shape[1] > 2 and pd.notna(df_raw.iat[r_i, 2]) else ""
        c_clean = c_val.lstrip(":").strip()
        if a_val == "NOP" and c_clean:
            nop = c_clean.upper()
        elif "TICKET" in a_val and c_clean:
            ticket = c_clean

    # Cari baris header tabel (baris dengan 'SITE ID')
    hdr_row = 4
    for r_i in range(min(12, len(df_raw))):
        row_u = [str(x).strip().upper() for x in df_raw.iloc[r_i].tolist() if pd.notna(x)]
        if "SITE ID" in row_u and ("QTY" in row_u or "SL" in row_u):
            hdr_row = r_i
            break

    sites = []
    cur_site = None
    cur_sec = "A"
    biaya_lain = 0.0

    for r_i in range(hdr_row + 1, len(df_raw)):
        def _g(ci):
            if ci >= df_raw.shape[1]: return None
            v = df_raw.iat[r_i, ci]
            return None if (v is None or (isinstance(v, float) and v != v)) else v

        col_a, col_b, col_c, col_d = _g(0), _g(1), _g(2), _g(3)
        col_e, col_f, col_g = _g(4), _g(5), _g(6)
        col_h, col_i, col_n, col_t, col_u = _g(7), _g(8), _g(13), _g(19), _g(20)

        t_str = str(col_t or "").strip().lower()
        l_str = str(_g(11) or "").strip().upper()
        if "biaya lain" in t_str:
            bl = _sv_parse_num_flex(col_u)
            if bl is not None:
                biaya_lain = bl
            continue
        if l_str == "TOTAL" or str(col_n or "").strip().upper() == "TOTAL":
            break

        if col_b is not None and str(col_b).strip() and str(col_b).strip().upper() not in ("SITE ID", "NONE"):
            cur_site = {
                "site_id": str(col_b).strip().upper(),
                "site_name": str(col_c or col_b).strip().upper(),
                "pekerjaan": str(col_d or "Pekerjaan SVA").strip(),
                "items": []
            }
            sites.append(cur_site)

        e_u = str(col_e or "").strip().upper()
        if e_u in ("A", "B"):
            cur_sec = e_u
            continue

        desc = str(col_g or "").strip()
        qty_v = _sv_parse_num_flex(col_i)
        ht_v = _sv_parse_num_flex(col_h)
        if desc and qty_v and qty_v > 0 and ht_v is not None and ht_v > 0:
            if cur_site is None:
                cur_site = {"site_id": "SITE-1", "site_name": "SITE-1", "pekerjaan": "Pekerjaan SVA", "items": []}
                sites.append(cur_site)
            f_low = str(col_f or "").strip().lower()
            jenis = "Jasa" if "jasa" in f_low else ("Material" if "material" in f_low else sv_guess_jenis(desc))
            hk_v = _sv_parse_num_flex(col_n)
            hp_v = _sv_parse_num_flex(col_t)
            cur_site["items"].append({
                "sec": cur_sec,
                "jenis": jenis,
                "desc": desc,
                "uom": "unit",
                "qty": float(qty_v),
                "harga_tsel": float(ht_v),
                "harga_kut": float(hk_v) if (hk_v is not None and hk_v > 0) else None,
                "hpp": float(hp_v) if (hp_v is not None and hp_v > 0) else None
            })

    sites = [s for s in sites if s["items"]]
    if not sites:
        return []

    params = dict(SV_PARAMS_DEFAULT)
    params["biaya_lain"] = biaya_lain
    rab = {
        "nop": nop,
        "ticket": ticket or f"SVA-IMP-{_dtm.date.today().strftime('%Y%m%d')}",
        "area": default_area,
        "kota": kota,
        "tanggal": _dtm.date.today().isoformat(),
        "params": params,
        "sites": sites
    }
    return [{"rab": rab, "status": S_RAB_SPV, "sva": None, "tgl_pengajuan": sv_now(), "catatan": "Import dari Excel RAB Resmi"}]

def sv_parse_bulk_table_to_entries(df_in, default_nop="PALANGKA RAYA", default_area="KALIMANTAN", default_kota="Palangka Raya", default_status=S_RAB_SPV, group_by_ticket=False):
    """Mengonversi DataFrame tabel Input All menjadi daftar objek siap simpan ke SISAVA_RAB."""
    if df_in is None or df_in.empty:
        return []

    cols = list(df_in.columns)
    def _pick_col(keywords, exclude=()):
        for kw in keywords:
            for c in cols:
                cl = str(c).strip().lower()
                if any(ex in cl for ex in exclude):
                    continue
                if kw == cl or kw in cl:
                    return c
        return None

    c_rab_id = _pick_col(["rab_id", "no. rab", "no rab"])
    c_ticket = _pick_col(["no. ticket", "no ticket", "ticket", "tiket"])
    c_nop = _pick_col(["nop"])
    c_area = _pick_col(["area"])
    c_kota = _pick_col(["kota", "city"])
    c_tgl = _pick_col(["tanggal pengajuan", "tgl pengajuan", "tgl_pengajuan", "tanggal", "tgl"])
    c_sid = _pick_col(["site id", "site_id", "siteid"])
    c_sname = _pick_col(["site name", "site_name", "nama site"])
    c_pek = _pick_col(["pekerjaan", "uraian", "activity", "deskripsi", "keperluan", "item"])
    c_jns = _pick_col(["jenis"])
    c_qty = _pick_col(["qty", "jumlah", "vol"])
    c_uom = _pick_col(["uom", "satuan"])
    c_tsel = _pick_col(["harga 3e", "total ke telkomsel", "total telkomsel", "tot_tsel", "nilai 3e", "harga tsel", "tsel"], exclude=("kut", "hpp"))
    c_kut = _pick_col(["harga kut", "revenue kut", "total revenue", "tot_rev", "revenue", "kut"])
    c_hpp = _pick_col(["hpp lapangan", "total hpp", "nilai hpp", "tot_hpp", "hpp"])
    c_bl = _pick_col(["biaya lain", "biaya_lain"])
    c_stat = _pick_col(["status sisava", "status rab", "status"], exclude=("tiket",))
    c_sva_n = _pick_col(["nilai sva", "sva"], exclude=("status", "ticket", "tiket"))
    c_dana = _pick_col(["keperluan dana", "dana"])
    c_urg = _pick_col(["urgensi"])
    c_st_tk = _pick_col(["status tiket", "status_tiket"])
    c_rnc = _pick_col(["rencana tgl", "tgl rencana", "tgl_rencana"])
    c_cat = _pick_col(["catatan", "remark", "keterangan", "note"])
    c_pem = _pick_col(["pembuat", "nama_pembuat", "pic"])

    def _s(row, col, default=""):
        if not col: return default
        v = row.get(col)
        if v is None or (isinstance(v, float) and v != v): return default
        st_v = str(v).strip()
        return default if st_v.lower() in ("nan", "nat", "none", "") else st_v

    def _match_status(raw_st):
        if not raw_st: return default_status
        rl = raw_st.strip().lower()
        for opt in SV_ALL_STATUSES:
            if opt.lower() == rl:
                return opt
        if "sah" in rl or "selesai" in rl or "done" in rl or "close" in rl:
            return S_DONE
        if "sva" in rl and ("terisi" in rl or "berjalan" in rl or "ok" in rl):
            return S_SVA
        if "disetujui" in rl or "approve" in rl:
            return S_RAB_OK
        if "manager" in rl or "mgr" in rl:
            return S_RAB_MGR
        if "tolak" in rl or "revisi" in rl:
            return S_RAB_TOLAK
        return default_status

    parsed_rows = []
    for idx_r, row in df_in.iterrows():
        ticket = _s(row, c_ticket)
        sid = _s(row, c_sid).upper()
        sname = _s(row, c_sname).upper() or sid
        pek = _s(row, c_pek)
        ht = _sv_parse_num_flex(row.get(c_tsel)) if c_tsel else None
        hk = _sv_parse_num_flex(row.get(c_kut)) if c_kut else None
        hp = _sv_parse_num_flex(row.get(c_hpp)) if c_hpp else None

        # Lewati baris kosong total
        if not ticket and not sid and not pek and ht is None and hk is None:
            continue

        # Jika Harga 3E kosong tapi Revenue KUT ada, back-calculate Harga 3E
        if (ht is None or ht <= 0) and (hk is not None and hk > 0):
            ht = round(hk / SV_PARAMS_DEFAULT["f_mat"], 2)
        if ht is None:
            ht = 0.0

        qty = _sv_parse_num_flex(row.get(c_qty)) if c_qty else None
        qty = qty if (qty is not None and qty > 0) else 1.0
        jns_raw = _s(row, c_jns)
        jenis = "Jasa" if "jasa" in jns_raw.lower() else ("Material" if "mat" in jns_raw.lower() else sv_guess_jenis(pek))

        # Jika Excel berupa rekap total (qty=1) dan hk/hp diisi sebagai angka total, unit = nilai langsung
        hk_unit = (hk / qty) if (hk is not None and hk > 0) else None
        hp_unit = (hp / qty) if (hp is not None and hp > 0) else None

        bl = _sv_parse_num_flex(row.get(c_bl)) if c_bl else 0.0
        bl = bl if (bl is not None and bl > 0) else 0.0

        tgl_raw = _s(row, c_tgl, _dtm.date.today().isoformat())
        tgl_iso = tgl_raw[:10] if len(tgl_raw) >= 10 and tgl_raw[4:5] == "-" else _dtm.date.today().isoformat()

        sva_val = _sv_parse_num_flex(row.get(c_sva_n)) if c_sva_n else None
        dana_val = _sv_parse_num_flex(row.get(c_dana)) if c_dana else None
        sva_dict = None
        if (sva_val is not None and sva_val > 0) or (dana_val is not None and dana_val > 0):
            sva_dict = {
                "ticket_no": ticket or sid or f"IMP-{idx_r+1}",
                "status_tiket": _s(row, c_st_tk, "Open"),
                "nilai_sva": float(sva_val or ht * qty),
                "keperluan_dana": float(dana_val or (hp if hp else ht * qty * 0.78)),
                "urgensi": _s(row, c_urg, "Normal"),
                "keperluan": pek or "Pekerjaan SVA",
                "tgl_rencana": _s(row, c_rnc, tgl_iso)[:10]
            }

        parsed_rows.append({
            "existing_rab_id": _s(row, c_rab_id),
            "ticket": ticket or f"SVA-IMP-{_dtm.date.today().strftime('%Y%m%d')}-{idx_r+1:03d}",
            "nop": _s(row, c_nop, default_nop).upper(),
            "area": _s(row, c_area, default_area).upper(),
            "kota": _s(row, c_kota, default_kota),
            "tanggal": tgl_iso,
            "site_id": sid or f"SITE-{idx_r+1:03d}",
            "site_name": sname or sid or f"SITE-{idx_r+1:03d}",
            "pekerjaan": pek or "Pekerjaan Service Variable Activity",
            "jenis": jenis,
            "uom": _s(row, c_uom, "lot"),
            "qty": float(qty),
            "harga_tsel": float(ht),
            "harga_kut": float(hk_unit) if hk_unit is not None else None,
            "hpp": float(hp_unit) if hp_unit is not None else None,
            "biaya_lain": float(bl),
            "status": _match_status(_s(row, c_stat)),
            "sva": sva_dict,
            "catatan": _s(row, c_cat, "Import Excel Input All"),
            "pembuat_nama": _s(row, c_pem)
        })

    entries = []
    if group_by_ticket:
        grouped = {}
        order_keys = []
        for pr in parsed_rows:
            gk = pr["ticket"].upper()
            if gk not in grouped:
                grouped[gk] = []
                order_keys.append(gk)
            grouped[gk].append(pr)

        for gk in order_keys:
            grp = grouped[gk]
            first = grp[0]
            sites_map = {}
            site_order = []
            bl_tot = 0.0
            for item_r in grp:
                bl_tot += item_r["biaya_lain"]
                sk = item_r["site_id"]
                if sk not in sites_map:
                    sites_map[sk] = {
                        "site_id": item_r["site_id"],
                        "site_name": item_r["site_name"],
                        "pekerjaan": item_r["pekerjaan"],
                        "items": []
                    }
                    site_order.append(sk)
                sites_map[sk]["items"].append({
                    "sec": "A",
                    "jenis": item_r["jenis"],
                    "desc": item_r["pekerjaan"],
                    "uom": item_r["uom"],
                    "qty": item_r["qty"],
                    "harga_tsel": item_r["harga_tsel"],
                    "harga_kut": item_r["harga_kut"],
                    "hpp": item_r["hpp"]
                })
            params = dict(SV_PARAMS_DEFAULT)
            params["biaya_lain"] = bl_tot
            rab = {
                "nop": first["nop"],
                "ticket": first["ticket"],
                "area": first["area"],
                "kota": first["kota"],
                "tanggal": first["tanggal"],
                "params": params,
                "sites": [sites_map[sk] for sk in site_order]
            }
            entries.append({
                "existing_rab_id": first["existing_rab_id"],
                "rab": rab,
                "status": first["status"],
                "sva": first["sva"],
                "tgl_pengajuan": first["tanggal"] + " 08:00:00",
                "catatan": first["catatan"],
                "pembuat_nama": first["pembuat_nama"]
            })
    else:
        for pr in parsed_rows:
            params = dict(SV_PARAMS_DEFAULT)
            params["biaya_lain"] = pr["biaya_lain"]
            rab = {
                "nop": pr["nop"],
                "ticket": pr["ticket"],
                "area": pr["area"],
                "kota": pr["kota"],
                "tanggal": pr["tanggal"],
                "params": params,
                "sites": [{
                    "site_id": pr["site_id"],
                    "site_name": pr["site_name"],
                    "pekerjaan": pr["pekerjaan"],
                    "items": [{
                        "sec": "A",
                        "jenis": pr["jenis"],
                        "desc": pr["pekerjaan"],
                        "uom": pr["uom"],
                        "qty": pr["qty"],
                        "harga_tsel": pr["harga_tsel"],
                        "harga_kut": pr["harga_kut"],
                        "hpp": pr["hpp"]
                    }]
                }]
            }
            entries.append({
                "existing_rab_id": pr["existing_rab_id"],
                "rab": rab,
                "status": pr["status"],
                "sva": pr["sva"],
                "tgl_pengajuan": pr["tanggal"] + " 08:00:00",
                "catatan": pr["catatan"],
                "pembuat_nama": pr["pembuat_nama"]
            })
    return entries

def sv_commit_entries_to_store(store, user, entries, dup_mode="skip", source_label="Import Excel"):
    """Menyimpan sekumpulan RAB hasil Import Excel / Input All secara batch ke Google Sheets (SISAVA_RAB)."""
    existing_all = store.list_rabs()
    by_ticket = {str(r.get("ticket_no", "")).strip().upper(): r for r in existing_all if str(r.get("ticket_no", "")).strip()}
    by_id = {str(r.get("rab_id", "")).strip().upper(): r for r in existing_all if str(r.get("rab_id", "")).strip()}

    year = _dtm.date.today().year
    nums = []
    for r in existing_all:
        parts = str(r.get("rab_id", "")).split("-")
        if len(parts) == 3 and parts[1] == str(year) and parts[2].isdigit():
            nums.append(int(parts[2]))
    next_seq = (max(nums) if nums else 0) + 1

    to_insert = []
    to_update_map = {}
    updated_cnt = 0
    skipped_cnt = 0

    for ent in entries:
        rab = ent["rab"]
        calc = sv_calc(rab)
        t_key = str(rab.get("ticket", "")).strip().upper()
        ex_id = str(ent.get("existing_rab_id", "")).strip().upper()
        matched_rec = by_id.get(ex_id) if ex_id else by_ticket.get(t_key)

        if matched_rec and dup_mode == "skip":
            skipped_cnt += 1
            continue

        maker_name = ent.get("pembuat_nama") or user["nama"]
        status_final = ent.get("status") or S_RAB_SPV
        sva_obj = ent.get("sva")

        if matched_rec and dup_mode == "update":
            rid = matched_rec["rab_id"]
            upd_fields = {
                "ticket_no": rab.get("ticket", ""),
                "nop": rab.get("nop", ""),
                "area": rab.get("area", ""),
                "kota": rab.get("kota", ""),
                "status": status_final,
                "tot_tsel": round(calc["tot_tsel"], 2),
                "tot_rev": round(calc["tot_rev"], 2),
                "tot_hpp": round(calc["tot_hpp"], 2),
                "margin": round(calc["margin"], 2),
                "gpm": round(calc["gpm"], 6),
                "net": round(calc["net"], 2),
                "payload": _sv_json(rab),
                "updated_at": sv_now(),
                "log": _sv_json(sv_log_add(matched_rec, user, "import_update", f"{source_label}: {ent.get('catatan', '')}"))
            }
            if sva_obj:
                upd_fields.update(sva=_sv_json(sva_obj), sva_oleh=user["username"], sva_at=sv_now())
            to_update_map[rid] = upd_fields
            continue

        rab_id = "RAB-%d-%04d" % (year, next_seq)
        next_seq += 1
        rec = sv_rab_to_record(rab, calc, user, rab_id, [])
        rec["nama_pembuat"] = maker_name
        rec["status"] = status_final
        if ent.get("tgl_pengajuan"):
            rec["tgl_pengajuan"] = ent["tgl_pengajuan"]
        if sva_obj:
            rec["sva"] = _sv_json(sva_obj)
            rec["sva_oleh"] = user["username"]
            rec["sva_at"] = sv_now()
        rec["log"] = _sv_json(sv_log_add({}, user, "import_excel", f"{source_label}: {ent.get('catatan', '')}"))
        to_insert.append(rec)
        if t_key:
            by_ticket[t_key] = rec

    if to_update_map:
        updated_cnt = store.update_rabs_batch(to_update_map)
    inserted_cnt = 0
    if to_insert:
        inserted_cnt = store.add_rabs_batch(to_insert)

    # Catat jejak audit import ke Google Sheets tab 'SISAVA_IMPORT_LOG'
    try:
        gsheet_append_row(
            "SISAVA_IMPORT_LOG",
            ["Timestamp", "User", "Sumber_File", "Mode_Duplikat", "Baru_Masuk", "Diupdate", "Dilewati"],
            {
                "Timestamp": sv_now(),
                "User": f"{user.get('nama', '')} ({user.get('username', '')})",
                "Sumber_File": source_label,
                "Mode_Duplikat": dup_mode,
                "Baru_Masuk": inserted_cnt,
                "Diupdate": updated_cnt,
                "Dilewati": skipped_cnt
            }
        )
    except Exception:
        pass

    return inserted_cnt, updated_cnt, skipped_cnt

def sv_page_import_all(store, user):
    st.markdown("### 📥 Import Excel & Input All Massal SiSAVA (Auto-Save Spreadsheet)")
    st.info("💡 Tim Admin dapat mengunggah **Excel rekap banyak tiket sekaligus**, mengunggah **file Excel RAB resmi SiSAVA**, atau mengetik/paste massal langsung di tabel. Semua data otomatis terhitung (3E ➔ KUT ➔ HPP ➔ Margin) dan tersimpan permanen ke Google Sheets tab **`SISAVA_RAB`**.")

    t_imp_xl, t_imp_grid, t_imp_db = st.tabs([
        "📂 1. Upload Excel (Input All / RAB Resmi)",
        "⌨️ 2. Input All Langsung di Tabel (Copy-Paste)",
        "🗂️ 3. Bulk Edit Database Spreadsheet (`SISAVA_RAB`)"
    ])

    # ---------------------------------------------------------------------
    # SUB-TAB 1: UPLOAD EXCEL -> PARSE -> PREVIEW -> SAVE BATCH KE GSHEETS
    # ---------------------------------------------------------------------
    with t_imp_xl:
        c_tpl1, c_tpl2 = st.columns([2, 1])
        with c_tpl1:
            st.markdown("**Langkah 1:** Unggah file Excel/CSV milik Tim Admin, atau gunakan Template resmi Input All di sebelah kanan.")
        with c_tpl2:
            st.download_button(
                "📄 Download Template Input All (.xlsx)",
                data=sv_build_template_input_all_xlsx(),
                file_name="Template_Input_All_SiSAVA.xlsx",
                mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                use_container_width=True,
                key="svw_dl_tpl_all"
            )

        up_sv_xl = st.file_uploader(
            "📂 Pilih File Excel / CSV SiSAVA (.xlsx, .xls, .csv):",
            type=["xlsx", "xls", "csv"],
            key="svw_up_excel_all"
        )

        c_opt1, c_opt2, c_opt3, c_opt4 = st.columns(4)
        def_nop_imp = c_opt1.selectbox("Default NOP (jika kosong):", SV_NOPS, key="svw_imp_nop")
        def_area_imp = c_opt2.selectbox("Default Area:", SV_AREAS, key="svw_imp_area")
        def_stat_imp = c_opt3.selectbox("Default Status RAB:", SV_ALL_STATUSES, index=0, key="svw_imp_stat")
        dup_opt_lbl = c_opt4.selectbox(
            "Jika No. Ticket sudah ada:",
            ["Lewati (Skip Duplikat)", "Update Data Lama", "Tetap Tambah Baru (Append)"],
            key="svw_imp_dup"
        )
        dup_mode = "skip" if "Lewati" in dup_opt_lbl else ("update" if "Update" in dup_opt_lbl else "append")

        if up_sv_xl is not None:
            try:
                is_official_rab = False
                df_raw_xl = None
                df_flat_xl = None

                if up_sv_xl.name.lower().endswith(".csv"):
                    df_flat_xl = pd.read_csv(up_sv_xl).dropna(how="all").reset_index(drop=True)
                else:
                    sh_names = get_excel_sheet_names_safe(up_sv_xl)
                    pick_sh = st.selectbox("Pilih Sheet Excel yang akan diimpor:", sh_names, key="svw_imp_sh")
                    df_raw_xl = read_excel_safe_raw(up_sv_xl, sheet_name=pick_sh, header=None)
                    is_official_rab = _sv_detect_official_rab_sheet(df_raw_xl)

                    if not is_official_rab:
                        df_flat_xl = parse_excel_smart_header(
                            up_sv_xl,
                            pick_sh,
                            target_keywords=("ticket", "tiket", "site id", "site_id", "pekerjaan", "harga 3e", "no. rab", "rab_id")
                        )

                if is_official_rab:
                    st.success("✅ Terdeteksi otomatis: **Format File Excel RAB Resmi SiSAVA (`RENCANA ANGGARAN BIAYA`)**!")
                    entries = sv_parse_official_rab_excel(df_raw_xl, default_nop=def_nop_imp, default_area=def_area_imp)
                    for e_it in entries:
                        e_it["status"] = def_stat_imp
                else:
                    grp_ticket = st.checkbox(
                        "🔗 Gabungkan baris yang memiliki 'No. Ticket' sama menjadi 1 dokumen RAB (Multi-Site / Multi-Item)",
                        value=False,
                        key="svw_imp_grp"
                    )
                    st.markdown(f"**Preview & Edit Cepat Data Excel ({len(df_flat_xl)} baris terbaca — bisa diedit langsung sebelum disimpan):**")
                    df_flat_ed = st.data_editor(df_flat_xl, use_container_width=True, num_rows="dynamic", key="svw_imp_flat_ed")
                    entries = sv_parse_bulk_table_to_entries(
                        df_flat_ed,
                        default_nop=def_nop_imp,
                        default_area=def_area_imp,
                        default_status=def_stat_imp,
                        group_by_ticket=grp_ticket
                    )

                if not entries:
                    st.warning("⚠️ Tidak ada baris data RAB yang valid ditemukan di sheet ini. Pastikan kolom minimal berisi No. Ticket / Site ID / Pekerjaan / Harga 3E.")
                else:
                    # Ringkasan kalkulasi otomatis sebelum masuk Spreadsheet
                    prev_rows = []
                    sum_tsel = sum_rev = sum_hpp = sum_mar = 0.0
                    for idx_e, ent in enumerate(entries, 1):
                        rab_e = ent["rab"]
                        calc_e = sv_calc(rab_e)
                        sum_tsel += calc_e["tot_tsel"]
                        sum_rev += calc_e["tot_rev"]
                        sum_hpp += calc_e["tot_hpp"]
                        sum_mar += calc_e["margin"]
                        sids_str = ", ".join(s.get("site_id", "-") for s in rab_e.get("sites", []))
                        prev_rows.append({
                            "No": idx_e,
                            "No. Ticket": rab_e.get("ticket"),
                            "NOP": rab_e.get("nop"),
                            "Site ID": sids_str,
                            "Jml Item": sum(len(s.get("items", [])) for s in rab_e.get("sites", [])),
                            "Status": ent.get("status"),
                            "Tagihan 3E→Tsel": sv_rp(calc_e["tot_tsel"]),
                            "Revenue KUT→3E": sv_rp(calc_e["tot_rev"]),
                            "Total HPP": sv_rp(calc_e["tot_hpp"]),
                            "Margin": sv_rp(calc_e["margin"]),
                            "GPM": "%.1f%%" % (calc_e["gpm"] * 100),
                            "Nilai SVA": sv_rp((ent.get("sva") or {}).get("nilai_sva", 0)) if ent.get("sva") else "-"
                        })

                    st.markdown("#### 📊 Ringkasan Otomatis Data yang Akan Disimpan ke Spreadsheet (`SISAVA_RAB`)")
                    km1, km2, km3, km4, km5 = st.columns(5)
                    km1.metric("Jumlah Dokumen RAB", f"{len(entries)} RAB")
                    km2.metric("Total Tagihan Tsel", sv_rp(sum_tsel))
                    km3.metric("Total Revenue KUT", sv_rp(sum_rev))
                    km4.metric("Total HPP", sv_rp(sum_hpp))
                    km5.metric("Total Margin", sv_rp(sum_mar), delta="GPM %.1f%%" % ((sum_mar / sum_rev * 100) if sum_rev else 0))

                    st.dataframe(pd.DataFrame(prev_rows), use_container_width=True, hide_index=True)

                    if st.button(f"🚀 SIMPAN SEMUA ({len(entries)} RAB) KE GOOGLE SHEETS (SISAVA_RAB)", type="primary", use_container_width=True, key="svw_btn_commit_xl"):
                        with st.spinner("Menyimpan seluruh data ke Google Sheets tab 'SISAVA_RAB'..."):
                            ins_c, upd_c, skp_c = sv_commit_entries_to_store(store, user, entries, dup_mode=dup_mode, source_label=up_sv_xl.name)
                        st.session_state["sv_flash"] = f"✅ Import Excel selesai! Baru masuk: {ins_c} RAB | Diperbarui: {upd_c} RAB | Dilewati (duplikat): {skp_c} RAB."
                        st.rerun()
            except Exception as e:
                st.error(f"Gagal membaca atau memproses file Excel: {e}")

    # ---------------------------------------------------------------------
    # SUB-TAB 2: INPUT ALL LANGSUNG DI TABEL (TANPA UPLOAD FILE)
    # ---------------------------------------------------------------------
    with t_imp_grid:
        st.markdown("#### ⌨️ Input All Cepat (Ketik atau Copy-Paste Banyak Baris Sekaligus)")
        st.caption("💡 Kosongkan `Harga KUT` dan `HPP Lapangan` jika ingin dihitung otomatis (92% Material / 85% Jasa, HPP 85% KUT). Klik tanda `+` di baris bawah tabel untuk menambah baris baru.")

        if "sv_quick_df" not in st.session_state:
            st.session_state["sv_quick_df"] = pd.DataFrame([{
                "No. Ticket": "",
                "NOP": SV_NOPS[0],
                "Site ID": "",
                "Site Name": "",
                "Pekerjaan": "",
                "Jenis": "Material",
                "Qty": 1.0,
                "Harga 3E (Tsel)": 0.0,
                "Harga KUT (0=Auto)": 0.0,
                "HPP Lapangan (0=Auto)": 0.0,
                "Status SiSAVA": S_RAB_SPV,
                "Nilai SVA (0=Kosong)": 0.0,
                "Catatan": ""
            }])

        ed_quick = st.data_editor(
            st.session_state["sv_quick_df"],
            use_container_width=True,
            num_rows="dynamic",
            column_config={
                "NOP": st.column_config.SelectboxColumn("NOP", options=SV_NOPS, default=SV_NOPS[0]),
                "Jenis": st.column_config.SelectboxColumn("Jenis", options=["Material", "Jasa"], default="Material"),
                "Qty": st.column_config.NumberColumn("Qty", min_value=0.0, step=1.0, default=1.0),
                "Harga 3E (Tsel)": st.column_config.NumberColumn("Harga 3E (Tsel)", min_value=0.0, step=10000.0, format="%.0f"),
                "Harga KUT (0=Auto)": st.column_config.NumberColumn("Harga KUT (0=Auto)", min_value=0.0, step=10000.0, format="%.0f"),
                "HPP Lapangan (0=Auto)": st.column_config.NumberColumn("HPP Lapangan (0=Auto)", min_value=0.0, step=10000.0, format="%.0f"),
                "Status SiSAVA": st.column_config.SelectboxColumn("Status SiSAVA", options=SV_ALL_STATUSES, default=S_RAB_SPV),
                "Nilai SVA (0=Kosong)": st.column_config.NumberColumn("Nilai SVA (0=Kosong)", min_value=0.0, step=10000.0, format="%.0f")
            },
            key="svw_quick_input_ed"
        )

        cq1, cq2 = st.columns([2, 1])
        grp_q = cq1.checkbox("🔗 Gabungkan baris dengan No. Ticket sama menjadi 1 RAB", value=False, key="svw_q_grp")
        if cq2.button("🧹 Reset Tabel Input", use_container_width=True, key="svw_q_rst"):
            st.session_state.pop("sv_quick_df", None)
            st.rerun()

        if st.button("💾 SIMPAN INPUT ALL TABEL KE GOOGLE SHEETS (SISAVA_RAB)", type="primary", use_container_width=True, key="svw_btn_save_quick"):
            df_q_norm = ed_quick.rename(columns={
                "Harga KUT (0=Auto)": "Harga KUT",
                "HPP Lapangan (0=Auto)": "HPP Lapangan",
                "Nilai SVA (0=Kosong)": "Nilai SVA"
            })
            q_entries = sv_parse_bulk_table_to_entries(df_q_norm, group_by_ticket=grp_q)
            if not q_entries:
                st.error("⚠️ Belum ada baris valid untuk disimpan. Isi minimal No. Ticket / Site ID dan Harga 3E.")
            else:
                with st.spinner("Menyimpan Input All ke Google Sheets..."):
                    ins_q, upd_q, skp_q = sv_commit_entries_to_store(store, user, q_entries, dup_mode="append", source_label="Input All Tabel")
                st.session_state.pop("sv_quick_df", None)
                st.session_state["sv_flash"] = f"✅ Berhasil menyimpan {ins_q} dokumen RAB dari Input All Tabel ke Google Sheets!"
                st.rerun()

    # ---------------------------------------------------------------------
    # SUB-TAB 3: BULK EDIT & SINKRONISASI DATABASE SISAVA_RAB
    # ---------------------------------------------------------------------
    with t_imp_db:
        st.markdown("#### 🗂️ Database Spreadsheet `SISAVA_RAB` — Live Update Status & Export")
        recs_all_db = store.list_rabs() if sv_can(user, "view_all") else sv_visible(store.list_rabs(), user)
        if not recs_all_db:
            st.info("Belum ada data RAB di Spreadsheet `SISAVA_RAB`.")
        else:
            rows_db_v = [{
                "rab_id": r.get("rab_id", ""),
                "ticket_no": r.get("ticket_no", ""),
                "nop": r.get("nop", ""),
                "tgl_pengajuan": str(r.get("tgl_pengajuan", ""))[:16],
                "nama_pembuat": r.get("nama_pembuat", ""),
                "status": r.get("status", S_RAB_SPV),
                "tot_tsel": _sv_num(r.get("tot_tsel")),
                "tot_rev": _sv_num(r.get("tot_rev")),
                "tot_hpp": _sv_num(r.get("tot_hpp")),
                "margin": _sv_num(r.get("margin"))
            } for r in recs_all_db]
            df_db_ed = pd.DataFrame(rows_db_v)

            ed_db = st.data_editor(
                df_db_ed,
                use_container_width=True,
                hide_index=True,
                disabled=["rab_id", "tgl_pengajuan"],
                column_config={
                    "rab_id": st.column_config.TextColumn("No. RAB"),
                    "ticket_no": st.column_config.TextColumn("No. Ticket"),
                    "nop": st.column_config.SelectboxColumn("NOP", options=SV_NOPS),
                    "status": st.column_config.SelectboxColumn("Status SiSAVA", options=SV_ALL_STATUSES),
                    "tot_tsel": st.column_config.NumberColumn("Total Tsel (Rp)", format="%.0f"),
                    "tot_rev": st.column_config.NumberColumn("Revenue KUT (Rp)", format="%.0f"),
                    "tot_hpp": st.column_config.NumberColumn("HPP (Rp)", format="%.0f"),
                    "margin": st.column_config.NumberColumn("Margin (Rp)", format="%.0f")
                },
                key="svw_bulk_db_ed"
            )

            cb_db1, cb_db2 = st.columns(2)
            with cb_db1:
                if st.button("💾 Simpan Perubahan Massal ke Google Sheets", type="primary", use_container_width=True, key="svw_btn_bulk_upd"):
                    if not sv_can(user, "view_all"):
                        st.warning("⚠️ Hanya Admin Utama / Manager / SPV yang dapat melakukan bulk update status database.")
                    else:
                        updates_batch_map = {}
                        for i_row in range(len(ed_db)):
                            orig_r = df_db_ed.iloc[i_row]
                            new_r = ed_db.iloc[i_row]
                            diff_fields = {}
                            for col_chk in ("ticket_no", "nop", "nama_pembuat", "status", "tot_tsel", "tot_rev", "tot_hpp", "margin"):
                                if str(orig_r[col_chk]) != str(new_r[col_chk]):
                                    diff_fields[col_chk] = new_r[col_chk]
                            if diff_fields:
                                rev_n = _sv_num(new_r["tot_rev"])
                                mar_n = _sv_num(new_r["margin"])
                                diff_fields["gpm"] = round((mar_n / rev_n) if rev_n else 0.0, 6)
                                diff_fields["updated_at"] = sv_now()
                                updates_batch_map[str(new_r["rab_id"])] = diff_fields
                        changed_cnt = 0
                        if updates_batch_map:
                            with st.spinner(f"Menyinkronkan {len(updates_batch_map)} perubahan sekaligus ke Google Sheets..."):
                                changed_cnt = store.update_rabs_batch(updates_batch_map)
                        if changed_cnt > 0:
                            st.session_state["sv_flash"] = f"✅ {changed_cnt} baris RAB berhasil diperbarui di Google Sheets!"
                            st.rerun()
                        else:
                            st.info("Tidak ada sel yang diubah.")
            with cb_db2:
                st.download_button(
                    "📥 Download Rekap Seluruh SiSAVA (.xlsx)",
                    data=sv_rekap_xlsx(recs_all_db),
                    file_name=f"Rekap_SiSAVA_{_dtm.date.today()}.xlsx",
                    mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                    use_container_width=True,
                    key="svw_dl_db_all"
                )

def sisava_page():
    ss = st.session_state
    store = sv_store()
    _sv_flash()
    
    if store.kind == "lokal":
        st.warning("⚠️ Google Sheets belum tersambung (st.secrets 'gcp_json'). Data disimpan sementara di file lokal dan bisa hilang saat aplikasi restart.")
        
    user = ss.get("sv_user")
    if not user:
        st.warning("⚠️ Sesi SiSAVA telah berakhir. Silakan login kembali.")
        st.session_state['sv_user'] = None
        st.rerun()
        return
        
    c1, c_rf, c2 = st.columns([3.5, 1, 1])
    c1.markdown("<div style='padding:8px 2px'>👤 <b>%s</b> · %s%s</div>" % (_esc(user["nama"]), _esc(SV_ROLES.get(user["role"], "")), (" · " + _esc(user["jabatan"])) if user.get("jabatan") else ""), unsafe_allow_html=True)
    if c_rf.button("🔄 Refresh", key="svw_refresh_btn", use_container_width=True):
        _sv_rows.clear()
        st.rerun()
    c2.button("Keluar Portal", key="svw_logout", on_click=lambda: st.session_state.pop("sv_user", None), use_container_width=True)
    if st.session_state.get("sv_quota_warn"):
        st.warning(st.session_state["sv_quota_warn"])
    
    # --- PENGATURAN MENU BERDASARKAN ROLE / HAK AKSES ---
    nav = ["🏠 Dashboard"]
    
    if sv_can(user, "rab_create"):
        nav.append("➕ Pengajuan RAB")
        nav.append("📥 Import Excel (Input All)")
        nav.append("📋 Daftar RAB")
        
    if sv_can(user, "approve_spv") or sv_can(user, "approve_mgr"):
        nav.append("✅ Approval")
        
    if sv_can(user, "sva_fill"):
        nav.append("🧾 Nilai SVA")
        
    if sv_can(user, "kerja_upload"):
        nav.append("📤 Upload Pekerjaan")
        
    if sv_can(user, "manage_users"):
        nav.append("🛠️ Kelola")
        
    if ss.get("sv_nav") not in nav:
        ss["sv_nav"] = nav[0]
        
    page = st.radio("Menu Navigasi SiSAVA", nav, horizontal=True, key="sv_nav", label_visibility="collapsed")
    
    try:
        recs = sv_visible(store.list_rabs(), user)
    except Exception as e:
        if _is_quota_429(e):
            st.warning("⏳ Server Google Sheets sedang padat (Limit 60 baca/menit tercapai). Tunggu ±30–60 detik lalu klik tombol **🔄 Refresh** di kanan atas.")
        else:
            st.error("Gagal membaca database SiSAVA: %s" % e)
        return
        
    if page == "🏠 Dashboard":
        sv_page_dashboard(user, recs)
    elif page == "➕ Pengajuan RAB":
        pl, path = sv_pricelist()
        sv_page_form(store, user, pl, path)
    elif page == "📥 Import Excel (Input All)":
        sv_page_import_all(store, user)
    elif page == "📋 Daftar RAB":
        sv_page_list(user, recs)
    elif page == "✅ Approval":
        sv_page_approval(user, recs)
    elif page == "🧾 Nilai SVA":
        sv_page_sva(user, recs)
    elif page == "📤 Upload Pekerjaan":
        sv_page_kerja(user, recs)
    elif page == "🛠️ Kelola":
        sv_page_kelola(store, user)
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
    if 'laporan_db' in st.session_state:
        _ux_stats, _ux_split = ux_stats_from_db(st.session_state.get("laporan_db", []))
        ux_hero(role_badge, "Sistem Operasional Terpadu SIDATA", "Pilih modul untuk melanjutkan. Data laporan tersinkron langsung dari server.", stats=_ux_stats, split=_ux_split)
    else:
        ux_hero(role_badge, "Sistem Operasional Terpadu SIDATA", "Mode Fast-Load Aktif ⚡ — Pilih modul di bawah untuk mulai bekerja tanpa menunggu sinkronisasi penuh.")
        c_fl1, c_fl2, c_fl3 = st.columns([2, 2, 2])
        if c_fl2.button("🔄 Muat Statistik Live Laporan", use_container_width=True, key="btn_lazy_stats_home"):
            with st.spinner("Menyinkronkan ringkasan laporan..."):
                get_laporan_db(force_refresh=True)
            st.rerun()

    # === PANEL KHUSUS ADMIN UTAMA: MANAJEMEN AKUN SiSAVA (ON-DEMAND) & PINTU PORTAL OFFICE ===
    if st.session_state['role'] == 'Admin':
        c_adm_p1, c_adm_p2 = st.columns([2.2, 1.3])
        with c_adm_p1:
            st.markdown("### 🛠️ Manajemen Akun Pengguna (SiSAVA)")
            open_akun_mgr = st.toggle("🔓 Buka Panel Buat Akun / Reset Sandi SiSAVA", value=False, key="tgl_adm_sisava_mgr")
            if open_akun_mgr:
                store = sv_store()
                users = store.list_users()
                with st.container(border=True):
                    tb_buat, tb_reset = st.tabs(["📝 Buat Akun Baru", "🔑 Reset Sandi"])
                    with tb_buat:
                        a1, a2 = st.columns(2)
                        nu_adm = a1.text_input("Username Baru (Tanpa Spasi)", key="svw_nu_adm_main")
                        nn_adm = a2.text_input("Nama Lengkap", key="svw_nn_adm_main")
                        a3, a4 = st.columns(2)
                        nj_adm = a3.text_input("Jabatan Tertulis", placeholder="mis: Engineering Project Control / Koordinator", key="svw_nj_adm_main")
                        nr_adm = a4.selectbox("Hak Akses (Role)", list(SV_ROLES), format_func=lambda x: SV_ROLES[x], key="svw_nr_adm_main")
                        pw_adm = st.text_input("Sandi (min. 8 karakter)", type="password", key="svw_pw_adm_main")
                        if st.button("💾 Simpan & Daftarkan Akun", type="primary", key="btn_save_akun_main"):
                            msg = sv_validate_new_user({"username": nu_adm.strip(), "nama": nn_adm, "role": nr_adm}, pw_adm)
                            if not msg and sv_user_find(store, nu_adm): msg = "Username tersebut sudah dipakai."
                            if msg:
                                st.error(msg)
                            else:
                                store.upsert_user({"username": nu_adm.strip(), "nama": nn_adm.strip(), "jabatan": nj_adm.strip(), "role": nr_adm, "pw_hash": sv_hash_pw(pw_adm), "aktif": "1", "dibuat": sv_now()})
                                st.success(f"✅ Akun untuk {nu_adm.strip()} ({SV_ROLES[nr_adm]}) berhasil didaftarkan!")
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
        with c_adm_p2:
            st.markdown("### 🏢 Akses Cepat Office")
            st.markdown("""
                <div class='area-card' style='border-top-color: #10B981; padding: 18px 16px;'>
                    <div style='font-size: 36px; margin-bottom: 6px;'>🏢</div>
                    <div class='login-title' style='font-size: 18px;'>Portal Office</div>
                    <div class='login-desc' style='font-size: 12.5px; margin-bottom: 8px;'>Logistik, SiHD, SiSAVA, Executive Project, Program NOP & SiDAVID.</div>
                </div>
            """, unsafe_allow_html=True)
            c_ao1, c_ao2 = st.columns(2)
            if c_ao1.button("🏢 PORTAL OFFICE", type="primary", use_container_width=True, key="m_adm_to_office"):
                st.session_state['active_menu'] = "🏢 Portal Office"
                st.rerun()
            if c_ao2.button("👥 KELOLA MASTER TIM", type="secondary", use_container_width=True, key="m_adm_to_mastertim"):
                st.session_state['active_menu'] = "👥 Kelola Master Tim (PLK & PBU)"
                st.rerun()

        st.markdown("<hr style='margin: 22px 0;'>", unsafe_allow_html=True)

    # === KOTAK MENU UTAMA (BERSIH & FOKUS OPERASIONAL LAPANGAN / ADMIN) ===
    col_m1, col_m2 = st.columns(2)
    with col_m1:
        if st.session_state['role'] == 'Admin':
            st.markdown("""
                <div class='area-card' style='border-top-color: #007BFF;'>
                    <div style='font-size: 46px; margin-bottom: 8px;'>📝</div>
                    <div class='login-title'>Form Preventive Check</div>
                    <div class='login-desc'>Input laporan Preventive Maintenance Site SPS, Site Reguler & Audit Tower Palangkaraya, serta Site Visit Pangkalan Bun.</div>
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
                    <div class='login-desc'>Input data pengerjaan lapangan, pengukuran parameter kelistrikan AC/DC, audit fisik tower & upload foto eviden wajib.</div>
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
                <div class='login-desc'>Lihat rekapitulasi laporan NOP Palangkaraya & Pangkalan Bun, analisa grafik power, unduh Excel/PPTX/PDF & bagikan WA.</div>
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
                <div class='login-desc'>Executive Dashboard compare bulanan AC/DC, AI Analisa Repeated Issue (PLK) & Komparasi Dapot vs Aktual (PBU).</div>
            </div>
        """, unsafe_allow_html=True)
        if st.button("⚡ BUKA DAPOT POWER HASIL PMS", type="secondary", use_container_width=True, key="m_dapot"):
            st.session_state['active_menu'] = "⚡ Dapot Power Hasil PMS"
            st.rerun()

    st.markdown("<hr style='margin: 26px 0;'>", unsafe_allow_html=True)
    c_out1, c_out2, c_out3, c_out4 = st.columns(4)
    with c_out1:
        st.markdown("""
            <div class='area-card' style='border-top-color: #F97316;'>
                <div style='font-size: 40px; margin-bottom: 8px;'>🚧</div>
                <div class='login-title' style='font-size: 18px;'>Form Hard Akses</div>
                <div class='login-desc' style='font-size: 12.5px;'>Input & Peta Rute Hard Akses (PBU & PLK), Moda Wajib Trail/4WD/Klotok & Foto.</div>
            </div>
        """, unsafe_allow_html=True)
        if st.button("🚧 FORM & PETA HARD AKSES", type="primary", use_container_width=True, key="m_hard_akses"):
            st.session_state['active_menu'] = "🚧 Form & Peta Hard Akses"
            st.rerun()

    with c_out2:
        st.markdown("""
            <div class='area-card' style='border-top-color: #9C27B0;'>
                <div style='font-size: 40px; margin-bottom: 8px;'>📅</div>
                <div class='login-title' style='font-size: 18px;'>Daily Tracker PLK</div>
                <div class='login-desc' style='font-size: 12.5px;'>Planning Harian Leader Cluster, Tracker Tier Urgent & Auto-Closed SIDATA.</div>
            </div>
        """, unsafe_allow_html=True)
        if st.button("📅 BUKA DAILY TRACKER", type="primary", use_container_width=True, key="m_tracker"):
            st.session_state['active_menu'] = "📅 Daily Tracker Palangkaraya"
            st.rerun()

    with c_out3:
        st.markdown("""
            <div class='area-card' style='border-top-color: #06B6D4;'>
                <div style='font-size: 40px; margin-bottom: 8px;'>🗺️</div>
                <div class='login-title' style='font-size: 18px;'>Sysinfomap (Peta Site)</div>
                <div class='login-desc' style='font-size: 12.5px;'>Peta radar interaktif 910+ Site, populasi Kabupaten & komparasi Master vs KML.</div>
            </div>
        """, unsafe_allow_html=True)
        if st.button("🗺️ BUKA SYSINFOMAP", type="primary", use_container_width=True, key="m_sysinfomap"):
            st.session_state['active_menu'] = "🗺️ Sysinfomap (Peta & Populasi Site)"
            st.rerun()

    with c_out4:
        st.markdown("""
            <div class='area-card' style='border-top-color: #64748B;'>
                <div style='font-size: 40px; margin-bottom: 8px;'>🚪</div>
                <div class='login-title' style='font-size: 18px;'>Keluar Akun</div>
                <div class='login-desc' style='font-size: 12.5px;'>Kembali ke layar utama 3 Portal (Lapangan, Office, Admin) dengan aman.</div>
            </div>
        """, unsafe_allow_html=True)
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
    _is_office_sub = menu in [
        "📦 Logistik & SPMS",
        "🚨 SiHD (Pergerakan Tim)",
        "💠 SiSAVA",
        "💎 Executive Project Dashboard",
        "📋 Program NOP & Task Force",
        "🛡️ SiDAVID (Verifikasi Internal)",
        "🚧 Form & Peta Hard Akses",
        "📅 Daily Tracker Palangkaraya",
        "👥 Kelola Master Tim (PLK & PBU)"
    ]
    c_top1, c_top2, c_top3 = st.columns([2.3, 5.4, 2.3])
    with c_top1:
        if _is_office_sub and st.session_state.get('role') is None:
            if st.button("⬅️ KEMBALI KE PORTAL OFFICE", type="secondary", use_container_width=True, key="btn_back_main"):
                st.session_state['active_menu'] = "🏢 Portal Office"
                st.session_state['dashboard_area'] = None
                st.rerun()
        elif menu == "🏢 Portal Office" and st.session_state.get('role') is None:
            if st.button("⬅️ KEMBALI KE LAYAR DEPAN", type="secondary", use_container_width=True, key="btn_back_main"):
                st.session_state['active_menu'] = None
                st.session_state['dashboard_area'] = None
                st.rerun()
        else:
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
            st.session_state['proj_dash_unlocked'] = False
            st.session_state['proj_unlocked'] = False
            st.session_state.pop('sv_user', None)
            st.rerun()
    st.markdown("<hr style='margin: 15px 0 25px 0;'>", unsafe_allow_html=True)


# =========================================================================
# MENU 1: FORM PENGISIAN LENGKAP
# =========================================================================
if menu in ["📝 Form Preventive Check", "📸 Form Pelaporan (Lapangan)"]:
    
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

    # Mode Ringan: hanya memuat Master Tim & Master Site sesuai area form aktif (tidak membaca KML PLK jika buka PBU!)
    if standar_form == "Pangkalan Bun":
        fme_pbu_list = load_master_tim_names("PBU")
        site_pbu_list, _ = load_site_master_pbu()
        fme_plk_list, site_plk_list = [], []
    else:
        fme_plk_list = load_master_tim_names("PLK")
        site_plk_list, _ = load_site_master_plk()
        fme_pbu_list, site_pbu_list = [], []

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
                                fetch_data_from_gsheets.clear()
                                if 'laporan_db' in st.session_state: st.session_state['laporan_db'].append(report_dict)
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
                                fetch_data_from_gsheets.clear()
                                if 'laporan_db' in st.session_state: st.session_state['laporan_db'].append(report_dict)
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
                                        
                                    fetch_data_from_gsheets.clear()
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
                            fetch_data_from_gsheets.clear()
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
    
    db = get_laporan_db()

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
        m2_view_mode = st.radio(
            "Mode Tampilan Dashboard (Hanya memuat tab aktif agar cepat):",
            ["📋 Daftar Detail Laporan Site", "📈 Analisa & Grafik Laporan"],
            horizontal=True,
            key="m2_dash_view_mode",
            label_visibility="collapsed"
        )
        
        if m2_view_mode == "📈 Analisa & Grafik Laporan":
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

        else:
            st.markdown("### 📋 Daftar Laporan Site Terinspeksi")
            c_pg1, c_pg2, c_pg3 = st.columns([1.4, 1.4, 2.2])
            per_page = c_pg1.selectbox("Tampilkan per halaman:", [10, 15, 25, 50], index=1, key="dash_per_page")
            tot_pages = max(1, _math.ceil(len(filtered_indices) / per_page))
            cur_page = c_pg2.number_input(f"Halaman (1 s/d {tot_pages}):", min_value=1, max_value=tot_pages, value=1, step=1, key="dash_cur_page")
            p_start = (int(cur_page) - 1) * int(per_page)
            p_end = min(p_start + int(per_page), len(filtered_indices))
            c_pg3.markdown(f"<div style='padding-top:30px;font-size:13px;opacity:0.85;'>⚡ Menampilkan laporan <b>{p_start+1}–{p_end}</b> dari <b>{len(filtered_indices)}</b> laporan (Mode Ringan)</div>", unsafe_allow_html=True)
            # ---------------------------------------------------------------------
            # TAMPILAN KARTU UI MODERN & URUTAN TERBARU DI ATAS (TERPAGINASI)
            # ---------------------------------------------------------------------
            for i in filtered_indices[p_start:p_end]:
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
                                                fetch_data_from_gsheets.clear() 
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
    db = get_laporan_db()

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
# MENU 5: DAILY TRACKER PLK (LEADER CLUSTER PLANNING + AUTO-CLOSED SIDATA)
# =========================================================================
elif menu == "📅 Daily Tracker Palangkaraya":
    st.markdown("<h1 style='text-align: center;'>📅 Daily Tracker & Leader Cluster Planning (PLK)</h1>", unsafe_allow_html=True)
    st.markdown("<p style='text-align: center; opacity: 0.85;'>Planning harian Leader Cluster, pencatatan pengalihan Urgent (Tracker Tier #1/#2/#3) & Deteksi Auto-Closed dari Report Tim di SIDATA.</p>", unsafe_allow_html=True)

    fme_plk_df, src_tim_plk, _ = load_master_tim_df("PLK")
    fme_plk_active = load_master_tim_names("PLK")
    leader_candidates = []
    if not fme_plk_df.empty:
        ldr_rows = fme_plk_df[fme_plk_df["ROLE"].astype(str).str.lower().str.contains("leader|koordinator")]
        leader_candidates = sorted(list({str(x).strip() for x in ldr_rows["NAMA"].dropna().tolist() if str(x).strip()}))
    leader_options = list(dict.fromkeys(leader_candidates + fme_plk_active))
    site_plk_list, _ = load_site_master_plk()

    df_dt_cloud, ok_dt_c = gsheet_read_df("Daily_Tracker_PLK", COLS_DAILY_TRACKER_PLK)

    dt_nav = st.radio(
        "Navigasi Daily Tracker PLK:",
        [
            "🎯 Monitoring Auto-Closed SIDATA & Follow Up",
            "📝 Input Planning Harian (Leader Cluster)",
            "🔀 Update Tracker Tier Urgent & Live Editor",
            "📂 Arsip Excel To Do List Lama"
        ],
        horizontal=True,
        key="dt_plk_nav",
        label_visibility="collapsed"
    )
    st.markdown("<hr style='margin: 8px 0 18px 0;'>", unsafe_allow_html=True)

    # ---------------------------------------------------------------------
    # SUB-1: MONITORING AUTO-CLOSED SIDATA & FOLLOW UP LEADER
    # ---------------------------------------------------------------------
    if dt_nav == "🎯 Monitoring Auto-Closed SIDATA & Follow Up":
        c_top_dt1, c_top_dt2 = st.columns([3.5, 1.2])
        c_top_dt1.info("💡 **Integrasi Otomatis SIDATA:** Jika `Plan Site ID` dan/atau `Tracker Tier #1/#2/#3` sudah terdeteksi masuk report tim di SIDATA pada tanggal tersebut, status otomatis menjadi **✅ CLOSED**. Jika belum ada report, status tetap **🔴 OPEN / 🟡 PARTIAL OPEN** agar Leader segera *follow up* tim.")
        if c_top_dt2.button("🔄 Sinkron Report SIDATA", use_container_width=True, key="btn_sync_dt_sidata"):
            with st.spinner("Menyegarkan data SIDATA & Daily Tracker..."):
                get_laporan_db(force_refresh=True)
                gsheet_read_df.clear()
            st.rerun()

        if df_dt_cloud.empty:
            st.warning("Belum ada data Planning di sheet `Daily_Tracker_PLK`. Silakan mulai input di tab **📝 Input Planning Harian (Leader Cluster)**.")
        else:
            db_live_dt = get_laporan_db()
            df_eval = evaluate_daily_tracker_with_sidata(df_dt_cloud, db_live_dt)

            # Filter Tanggal, Leader, Cluster, Status
            cf1, cf2, cf3, cf4 = st.columns(4)
            tgl_opts = ["Semua Tanggal"] + sorted(df_eval["Tanggal"].astype(str).unique().tolist(), reverse=True)
            pick_tgl_f = cf1.selectbox("📅 Filter Tanggal:", tgl_opts, key="dt_flt_tgl")
            ldr_opts = ["Semua Leader"] + sorted(df_eval["Leader_Cluster"].astype(str).unique().tolist())
            pick_ldr_f = cf2.selectbox("👤 Filter Leader Cluster:", ldr_opts, key="dt_flt_ldr")
            cl_opts = ["Semua Cluster"] + sorted(df_eval["Cluster"].astype(str).unique().tolist())
            pick_cl_f = cf3.selectbox("📍 Filter Cluster:", cl_opts, key="dt_flt_cl")
            st_opts = ["Semua Status", "✅ CLOSED", "🟡 PARTIAL OPEN", "🔴 OPEN"]
            pick_st_f = cf4.selectbox("🚦 Filter Status SIDATA:", st_opts, key="dt_flt_st")

            df_f = df_eval.copy()
            if pick_tgl_f != "Semua Tanggal":
                df_f = df_f[df_f["Tanggal"].astype(str) == pick_tgl_f]
            if pick_ldr_f != "Semua Leader":
                df_f = df_f[df_f["Leader_Cluster"].astype(str) == pick_ldr_f]
            if pick_cl_f != "Semua Cluster":
                df_f = df_f[df_f["Cluster"].astype(str) == pick_cl_f]
            if pick_st_f == "✅ CLOSED":
                df_f = df_f[df_f["Status_Auto_SIDATA"].str.contains("CLOSED")]
            elif pick_st_f == "🟡 PARTIAL OPEN":
                df_f = df_f[df_f["Status_Auto_SIDATA"].str.contains("PARTIAL")]
            elif pick_st_f == "🔴 OPEN":
                df_f = df_f[df_f["Status_Auto_SIDATA"].str.contains("🔴 OPEN")]

            tot_p = len(df_f)
            cnt_cl = int(df_f["Status_Auto_SIDATA"].str.contains("CLOSED").sum()) if tot_p else 0
            cnt_part = int(df_f["Status_Auto_SIDATA"].str.contains("PARTIAL").sum()) if tot_p else 0
            cnt_op = int(df_f["Status_Auto_SIDATA"].str.contains("🔴 OPEN").sum()) if tot_p else 0
            cnt_div = int(df_f["Eksekusi_Lapangan"].str.contains("Dialihkan").sum()) if tot_p else 0

            km1, km2, km3, km4, km5 = st.columns(5)
            km1.metric("📋 Total Planning", f"{tot_p} Task")
            km2.metric("✅ CLOSED (Terlapor)", f"{cnt_cl} ({(cnt_cl/tot_p*100) if tot_p else 0:.0f}%)")
            km3.metric("🟡 Partial Open", f"{cnt_part} Task")
            km4.metric("🔴 OPEN (Follow Up)", f"{cnt_op} Task")
            km5.metric("🔀 Dialihkan Urgent", f"{cnt_div} Task")

            cols_view = [
                "Tanggal", "Cluster", "Leader_Cluster", "Plan_Site_ID", "Program_Kerja",
                "Tim_Bertugas", "Eksekusi_Lapangan", "Status_Auto_SIDATA",
                "Site_Belum_Input_SIDATA", "Bukti_Report_SIDATA",
                "Tracker_Tier_1_Site_ID", "Tracker_Tier_1_Remark",
                "Tracker_Tier_2_Site_ID", "Tracker_Tier_2_Remark",
                "Tracker_Tier_3_Site_ID", "Tracker_Tier_3_Remark"
            ]
            cols_view = [c for c in cols_view if c in df_f.columns]
            st.markdown("#### 📊 Tabel Monitoring Realisasi Harian vs Inputan SIDATA")
            st.dataframe(df_f[cols_view], use_container_width=True, hide_index=True, height=400)

            # Daftar Tugas Follow Up Leader + Auto-WA Reminder ke Tim
            df_open_fu = df_f[~df_f["Status_Auto_SIDATA"].str.contains("CLOSED")]
            if not df_open_fu.empty:
                st.markdown(f"#### 🚨 Daftar Follow Up Leader — Belum Terdeteksi Input SIDATA ({len(df_open_fu)} Task)")
                for idx_fu, r_fu in df_open_fu.head(20).iterrows():
                    wa_rem = (
                        f"*🔔 FOLLOW UP INPUT LAPORAN SIDATA — NOP PALANGKARAYA*\n"
                        f"📅 *Tanggal:* {r_fu['Tanggal']}\n"
                        f"👤 *Leader Cluster:* {r_fu['Leader_Cluster']} ({r_fu['Cluster']})\n"
                        f"👷 *Tim Bertugas:* {r_fu['Tim_Bertugas']}\n"
                        f"📋 *Program:* {r_fu['Program_Kerja']} | *Plan Awal:* {r_fu['Plan_Site_ID']}\n"
                        f"🚦 *Status:* {r_fu['Status_Auto_SIDATA']}\n"
                        f"⚠️ *Site Belum Terdeteksi Input di SIDATA:* {r_fu['Site_Belum_Input_SIDATA']}\n\n"
                        f"Mohon tim segera menyelesaikan input laporan & eviden foto di aplikasi SIDATA hari ini agar status otomatis *CLOSED*. Terima kasih!"
                    )
                    with st.container(border=True):
                        cfu1, cfu2 = st.columns([3.2, 1.3])
                        cfu1.markdown(
                            f"**{r_fu['Tanggal']} · {r_fu['Cluster']}** | Leader: **{r_fu['Leader_Cluster']}** | Tim: **{r_fu['Tim_Bertugas']}**<br>"
                            f"🎯 Plan: `{r_fu['Plan_Site_ID']}` ({r_fu['Program_Kerja']}) · {r_fu['Eksekusi_Lapangan']}<br>"
                            f"🚨 **Wajib Follow Up Input SIDATA:** `{r_fu['Site_Belum_Input_SIDATA']}`",
                            unsafe_allow_html=True
                        )
                        cfu2.link_button("📱 Kirim Reminder WA Tim", f"https://wa.me/?text={urllib.parse.quote(wa_rem)}", use_container_width=True)
            else:
                st.success("🎉 Mantap! Seluruh task pada filter ini sudah **CLOSED** (terdeteksi report masuk di SIDATA).")

            # Export Excel Rekap Daily Tracker
            if st.button("🛠️ Siapkan Download Rekap Daily Tracker PLK (.xlsx)", use_container_width=True, key="btn_prep_dt_xl"):
                buf_dt = BytesIO()
                with pd.ExcelWriter(buf_dt, engine="openpyxl") as xw:
                    df_f[cols_view].to_excel(xw, sheet_name="Daily_Tracker_PLK", index=False)
                st.session_state["dt_plk_xlsx_bytes"] = buf_dt.getvalue()
            if st.session_state.get("dt_plk_xlsx_bytes"):
                st.download_button(
                    "📥 Download Rekap Daily Tracker + Auto-Closed (.xlsx)",
                    data=st.session_state["dt_plk_xlsx_bytes"],
                    file_name=f"Daily_Tracker_PLK_{datetime.date.today()}.xlsx",
                    mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                    type="primary",
                    use_container_width=True,
                    key="dl_dt_plk_xl"
                )

    # ---------------------------------------------------------------------
    # SUB-2: FORM INPUT PLANNING HARIAN OLEH LEADER CLUSTER
    # ---------------------------------------------------------------------
    elif dt_nav == "📝 Input Planning Harian (Leader Cluster)":
        st.markdown("#### 📝 Form Input Planning Harian Tim oleh Leader Cluster")
        st.info("💡 Leader mengisi rencana kerja pagi hari (`Plan Site ID`, `Program Kerja`, `Tim Tier #1/#2/#3`). Jika di tengah hari tim ternyata dialihkan ke Site lain yang lebih urgent, isi bagian **Tracker Tier #1/#2/#3** (bisa diisi sekarang atau di-update menyusul di tab **🔀 Update Tracker Tier Urgent**).")

        with st.container(border=True):
            c_i1, c_i2, c_i3 = st.columns(3)
            dt_tgl = c_i1.date_input("📅 Tanggal Planning *:", value=datetime.date.today(), key="dt_in_tgl")
            dt_cl = c_i2.selectbox("📍 Cluster *:", DAILY_CLUSTER_PLK_OPTS, key="dt_in_cl")
            dt_ldr_pick = c_i3.selectbox("👤 Nama Leader Cluster *:", ["-- Ketik Manual --"] + leader_options, key="dt_in_ldr_p")
            if dt_ldr_pick == "-- Ketik Manual --":
                dt_ldr = c_i3.text_input("Ketik Nama Leader:", key="dt_in_ldr_m").strip()
            else:
                dt_ldr = dt_ldr_pick

            c_i4, c_i5, c_i6 = st.columns([1.6, 1.1, 1.3])
            dt_plan_site = c_i4.selectbox("🎯 Plan Site ID (Master PLK) *:", [""] + site_plk_list, key="dt_in_psite")
            dt_prog = c_i5.selectbox("⚙️ Program Kerja *:", DAILY_PROGRAM_KERJA_OPTS, key="dt_in_prog")
            dt_plan_rmk = c_i6.text_input("📝 Remark Plan Awal:", placeholder="Cth: PM SPS / Cek Recti Drop", key="dt_in_prmk")

            st.markdown("**👷 Penugasan Tim Lapangan (Diambil dari Master Tim PLK Dinamis):**")
            c_t1, c_t2, c_t3 = st.columns(3)
            dt_tim1 = c_t1.selectbox("Tim Tier #1 (Utama) *:", [""] + fme_plk_active, key="dt_in_t1")
            dt_tim2 = c_t2.selectbox("Tim Tier #2 (Opsional):", [""] + fme_plk_active, key="dt_in_t2")
            dt_tim3 = c_t3.selectbox("Tim Tier #3 (Opsional):", [""] + fme_plk_active, key="dt_in_t3")

            with st.expander("🔀 Opsional: Isi Tracker Tier #1/#2/#3 (Jika Tim Dialihkan ke Site Urgent Lain)", expanded=False):
                st.caption("Contoh: Plan awal Site A, tetapi ada Site C & Site D lebih urgent pada tanggal tersebut -> pilih Site C di Tracker Tier #1 dan Site D di Tracker Tier #2 beserta Remark pekerjaannya.")
                tr_inputs = {}
                for k_t in (1, 2, 3):
                    ct1, ct2, ct3 = st.columns([1, 1.6, 1.8])
                    tr_inputs[f"tgl_{k_t}"] = ct1.date_input(f"Tgl Tracker Tier #{k_t}:", value=dt_tgl, key=f"dt_in_ttgl_{k_t}")
                    tr_inputs[f"site_{k_t}"] = ct2.selectbox(f"Tracker Tier #{k_t} Site ID:", [""] + site_plk_list, key=f"dt_in_tsite_{k_t}")
                    tr_inputs[f"rmk_{k_t}"] = ct3.text_input(f"Tracker Tier #{k_t} Remark:", placeholder=f"Pekerjaan urgent di Tier #{k_t}...", key=f"dt_in_trmk_{k_t}")

            if st.button("🚀 SIMPAN PLANNING LEADER KE CLOUD (`Daily_Tracker_PLK`)", type="primary", use_container_width=True, key="btn_save_dt_plk"):
                if not dt_ldr or not dt_plan_site or not dt_tim1:
                    st.error("⚠️ Nama Leader Cluster, Plan Site ID, dan Tim Tier #1 wajib diisi!")
                else:
                    rec_dt = {
                        "Tracker_ID": f"DT-PLK-{dt_tgl.strftime('%Y%m%d')}-{datetime.datetime.now().strftime('%H%M%S')}",
                        "Tanggal": dt_tgl.strftime("%Y-%m-%d"),
                        "Cluster": dt_cl,
                        "Leader_Cluster": dt_ldr,
                        "Plan_Site_ID": dt_plan_site,
                        "Program_Kerja": dt_prog,
                        "Plan_Remark": dt_plan_rmk.strip() or "-",
                        "Tim_Tier_1": dt_tim1,
                        "Tim_Tier_2": dt_tim2,
                        "Tim_Tier_3": dt_tim3,
                        "Tgl_Tracker_Tier_1": tr_inputs["tgl_1"].strftime("%Y-%m-%d") if tr_inputs["site_1"] else "",
                        "Tracker_Tier_1_Site_ID": tr_inputs["site_1"],
                        "Tracker_Tier_1_Remark": tr_inputs["rmk_1"].strip(),
                        "Tgl_Tracker_Tier_2": tr_inputs["tgl_2"].strftime("%Y-%m-%d") if tr_inputs["site_2"] else "",
                        "Tracker_Tier_2_Site_ID": tr_inputs["site_2"],
                        "Tracker_Tier_2_Remark": tr_inputs["rmk_2"].strip(),
                        "Tgl_Tracker_Tier_3": tr_inputs["tgl_3"].strftime("%Y-%m-%d") if tr_inputs["site_3"] else "",
                        "Tracker_Tier_3_Site_ID": tr_inputs["site_3"],
                        "Tracker_Tier_3_Remark": tr_inputs["rmk_3"].strip(),
                        "Updated_At": datetime.datetime.now().strftime("%Y-%m-%d %H:%M:%S")
                    }
                    if gsheet_append_row("Daily_Tracker_PLK", COLS_DAILY_TRACKER_PLK, rec_dt):
                        st.success(f"✅ Planning Leader **{dt_ldr}** untuk `{dt_plan_site}` ({dt_prog}) oleh **{dt_tim1}** berhasil tersimpan ke Cloud!")
                        st.rerun()
                    else:
                        st.error("⚠️ Gagal menyimpan ke Google Sheets tab `Daily_Tracker_PLK`.")

    # ---------------------------------------------------------------------
    # SUB-3: UPDATE TRACKER TIER URGENT (SITE C, D, DST.) & LIVE EDITOR
    # ---------------------------------------------------------------------
    elif dt_nav == "🔀 Update Tracker Tier Urgent & Live Editor":
        if df_dt_cloud.empty:
            st.warning("Belum ada data Planning untuk di-update.")
        else:
            st.markdown("#### 🔀 Catat Pengalihan Urgent (Tracker Tier #1 / #2 / #3) pada Planning yang Ada")
            st.caption("Pilih Planning pagi yang ternyata dialihkan ke Site Urgent (Site C, Site D, dst.), lalu simpan agar langsung terdeteksi otomatis terhadap laporan SIDATA.")

            opts_dt_edit = [
                f"{r['Tracker_ID']} | {r['Tanggal']} | Plan: {r['Plan_Site_ID']} | Tim: {r['Tim_Tier_1']} ({r['Leader_Cluster']})"
                for _, r in df_dt_cloud.iloc[::-1].iterrows()
            ]
            pick_dt_lbl = st.selectbox("📌 Pilih Baris Planning untuk Di-update Tracker Tier-nya:", opts_dt_edit, key="dt_upd_pick")
            pick_tid = pick_dt_lbl.split(" | ")[0].strip()
            row_cur = df_dt_cloud[df_dt_cloud["Tracker_ID"].astype(str) == pick_tid].iloc[0]

            with st.container(border=True):
                st.markdown(f"**Plan Awal:** `{row_cur['Plan_Site_ID']}` ({row_cur['Program_Kerja']}) · Tanggal `{row_cur['Tanggal']}` · Tim: **{row_cur['Tim_Tier_1']}**")
                upd_tr = {}
                for k_t in (1, 2, 3):
                    cu1, cu2, cu3 = st.columns([1, 1.7, 1.8])
                    cur_tgl_s = _fmt_date_short(row_cur.get(f"Tgl_Tracker_Tier_{k_t}"))
                    def_d = pd.to_datetime(cur_tgl_s, errors="coerce")
                    def_d_val = def_d.date() if pd.notna(def_d) else (pd.to_datetime(row_cur["Tanggal"], errors="coerce").date() if pd.notna(pd.to_datetime(row_cur["Tanggal"], errors="coerce")) else datetime.date.today())
                    upd_tr[f"tgl_{k_t}"] = cu1.date_input(f"Tgl Tracker #{k_t}:", value=def_d_val, key=f"dt_u_tgl_{k_t}_{pick_tid}")
                    cur_s = _clean_cell_str(row_cur.get(f"Tracker_Tier_{k_t}_Site_ID"))
                    s_opts_u = [""] + site_plk_list + ([cur_s] if cur_s and cur_s not in site_plk_list else [])
                    upd_tr[f"site_{k_t}"] = cu2.selectbox(f"Tracker Tier #{k_t} Site ID:", s_opts_u, index=s_opts_u.index(cur_s) if cur_s in s_opts_u else 0, key=f"dt_u_site_{k_t}_{pick_tid}")
                    upd_tr[f"rmk_{k_t}"] = cu3.text_input(f"Tracker Tier #{k_t} Remark:", value=_clean_cell_str(row_cur.get(f"Tracker_Tier_{k_t}_Remark")), key=f"dt_u_rmk_{k_t}_{pick_tid}")

                if st.button("💾 SIMPAN UPDATE TRACKER TIER KE CLOUD", type="primary", use_container_width=True, key=f"btn_dt_upd_tier_{pick_tid}"):
                    df_new_dt = df_dt_cloud.copy()
                    mask_t = df_new_dt["Tracker_ID"].astype(str) == pick_tid
                    for k_t in (1, 2, 3):
                        s_val = upd_tr[f"site_{k_t}"]
                        df_new_dt.loc[mask_t, f"Tgl_Tracker_Tier_{k_t}"] = upd_tr[f"tgl_{k_t}"].strftime("%Y-%m-%d") if s_val else ""
                        df_new_dt.loc[mask_t, f"Tracker_Tier_{k_t}_Site_ID"] = s_val
                        df_new_dt.loc[mask_t, f"Tracker_Tier_{k_t}_Remark"] = upd_tr[f"rmk_{k_t}"].strip()
                    df_new_dt.loc[mask_t, "Updated_At"] = datetime.datetime.now().strftime("%Y-%m-%d %H:%M:%S")
                    if gsheet_overwrite_ws("Daily_Tracker_PLK", df_new_dt):
                        st.success(f"✅ Tracker Tier untuk `{pick_tid}` berhasil diperbarui di Cloud!")
                        st.rerun()

            st.markdown("---")
            with st.expander("🗂️ Live Editor Massal Tabel `Daily_Tracker_PLK` (Edit / Hapus Baris)", expanded=False):
                ed_dt_all = st.data_editor(
                    df_dt_cloud,
                    use_container_width=True,
                    num_rows="dynamic",
                    column_config={
                        "Program_Kerja": st.column_config.SelectboxColumn("Program Kerja", options=DAILY_PROGRAM_KERJA_OPTS),
                        "Cluster": st.column_config.SelectboxColumn("Cluster", options=DAILY_CLUSTER_PLK_OPTS)
                    },
                    key="ed_dt_plk_cloud"
                )
                if st.button("💾 Simpan Perubahan Massal `Daily_Tracker_PLK`", type="primary", use_container_width=True, key="btn_save_ed_dt_all"):
                    if gsheet_overwrite_ws("Daily_Tracker_PLK", ed_dt_all):
                        st.success("✅ Database `Daily_Tracker_PLK` berhasil diperbarui!")
                        st.rerun()

    # ---------------------------------------------------------------------
    # SUB-4: ARSIP EXCEL TO DO LIST LAMA (OPSIONAL)
    # ---------------------------------------------------------------------
    else:
        file_todo = "To Do List NOP Palangkaraya.xlsx"
        if not os.path.exists(file_todo):
            st.info(f"ℹ️ File arsip lokal `{file_todo}` tidak ada di server. Seluruh Daily Tracker aktif kini tersimpan permanen di Google Sheets tab **`Daily_Tracker_PLK`**.")
        else:
            try:
                df_old_todo = read_excel_cached(file_todo, sheet_name=0)
                st.dataframe(df_old_todo, use_container_width=True, hide_index=True)
            except Exception as e:
                st.error(f"Gagal membaca arsip Excel lama: {e}")

# =========================================================================
# MENU 6: LOGISTIK & SPMS (MANAJEMEN STOK GUDANG)
# =========================================================================
elif menu == "📦 Logistik & SPMS":
    st.markdown("<h1 style='text-align: center;'>📦 Manajemen Logistik & SPMS</h1>", unsafe_allow_html=True)
    st.markdown("<p style='text-align: center; opacity: 0.8;'>Sistem Pencatatan Barang Masuk (Inbound), Barang Keluar (Outbound), dan Auto-Kalkulasi Stok.</p>", unsafe_allow_html=True)
    
    headers_standar = ["Tanggal", "Tipe Transaksi", "Kategori", "Merek / Nama Barang", "Untuk Genset/Mesin", "Jumlah", "PIC / Pengambil", "Site ID", "Keterangan"]
    df_log, ok_log = gsheet_read_df("Data Logistik", headers_standar)
    sheet_logistik = gsheet_get_or_create_ws("Data Logistik", headers_standar)
    if not ok_log or sheet_logistik is None:
        st.error("⚠️ **PENTING:** Gagal menyambung ke sheet **'Data Logistik'** di Spreadsheet Utama Anda.")
        st.stop()
    
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
                        _gsheet_retry(sheet_logistik.insert_row, row_in, index=2); gsheet_read_df.clear()
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
                        _gsheet_retry(sheet_logistik.insert_row, row_out, index=2); gsheet_read_df.clear()
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
            fetch_data_from_gsheets.clear()
            st.session_state['laporan_db'] = fetch_data_from_gsheets()
            st.rerun()

    db_all = get_laporan_db()
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
                                fetch_data_from_gsheets.clear()
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
                st.info("Belum ada rincian tabel rute PO untuk diekspor.")

# =========================================================================
# MENU 8: SiHD (SISTEM HELPDESK — MUTASI PERGERAKAN TIM & SERAH TERIMA SHIFT)
# =========================================================================
elif menu == "🚨 SiHD (Pergerakan Tim)":
    st.markdown("<h1 style='text-align: center;'>🚨 SiHD — Sistem Helpdesk & Mutasi Pergerakan Tim</h1>", unsafe_allow_html=True)
    st.markdown("<p style='text-align: center; opacity: 0.85;'>Buku Mutasi Elektronik Pergerakan Tim Lapangan (Palangkaraya & Pangkalan Bun) serta Auto-WA Serah Terima Shift</p>", unsafe_allow_html=True)

    fme_plk_list, fme_pbu_list, site_plk_list, site_pbu_list = load_master_lists()
    db_sihd = [r for r in get_laporan_db() if r.get('_sheet') == 'SiHD' or r.get('kategori') == 'SiHD']

    tot_sihd = len(db_sihd)
    today_str = datetime.date.today().strftime("%Y-%m-%d")
    today_sihd = sum(1 for r in db_sihd if str(r.get('timestamp', '')).startswith(today_str))
    open_sihd = sum(1 for r in db_sihd if str(r.get('status', '')).lower() in ['on progress', 'open', 'pending', 'otw / berangkat'])

    ux_hero(
        "HELPDESK COMMAND CENTER",
        "SiHD — Mutasi & Serah Terima Shift",
        "Rekam pergerakan FME secara real-time ke Google Sheets tab 'SiHD' dan kirim ringkasan serah terima shift ke WhatsApp dalam 1 klik.",
        stats=[
            (f"{tot_sihd} Log", "Total Catatan Mutasi"),
            (f"{today_sihd} Log", "Pergerakan Hari Ini"),
            (f"{open_sihd} Aktif", "Masih Open / On Progress")
        ]
    )

    tab_hd_in, tab_hd_log = st.tabs(["📝 Input Mutasi & Pergerakan Tim", "📋 Buku Mutasi & WA Serah Terima Shift"])

    with tab_hd_in:
        st.markdown("### 📝 Form Pencatatan Pergerakan Tim")
        with st.container(border=True):
            ch1, ch2, ch3 = st.columns(3)
            with ch1:
                hd_nop = st.selectbox("Area NOP:", ["Palangkaraya", "Pangkalan Bun"], key="hd_nop")
                hd_shift = st.selectbox("Shift Jaga:", ["Pagi (08:00 - 16:00)", "Siang (16:00 - 00:00)", "Malam (00:00 - 08:00)"], key="hd_shift")
                hd_pic = st.text_input("Nama Helpdesk / Dispatcher *:", placeholder="Ketik nama HD jaga...", key="hd_pic")
            with ch2:
                opsi_fme_hd = fme_plk_list if hd_nop == "Palangkaraya" else fme_pbu_list
                if opsi_fme_hd:
                    hd_tim = st.multiselect("Pilih Tim / FME Bergerak *:", opsi_fme_hd, key="hd_tim_sel")
                else:
                    t_hd_m = st.text_input("Nama Tim / FME *:", key="hd_tim_txt")
                    hd_tim = [t_hd_m] if t_hd_m.strip() else []
                opsi_site_hd = site_plk_list if hd_nop == "Palangkaraya" else site_pbu_list
                if opsi_site_hd:
                    hd_sites = st.multiselect("Site Target / Rute Kunjungan:", opsi_site_hd, key="hd_sites_sel")
                else:
                    hd_sites = []
                hd_site_manual = st.text_input("Atau Ketik Site/Rute Manual:", placeholder="Cth: PLK001 -> PLK025", key="hd_site_man")
            with ch3:
                hd_kat = st.selectbox("Kategori Kegiatan:", ["Troubleshoot Down / Alarm", "Preventive Maintenance (PM)", "Refuel / BBM Genset", "Audit Tower / Combat", "Project / Relokasi", "Pengambilan Material / Logistik", "Standby / Patroli"], key="hd_kat")
                hd_tiket = st.text_input("No. Tiket / Incident (Opsional):", placeholder="Cth: INC-2026-xxxx", key="hd_tiket")
                hd_status = st.selectbox("Status Pergerakan:", ["OTW / Berangkat", "On Site / Pengerjaan", "Done / Clear", "Pending / Butuh Follow-Up"], key="hd_status")

            hd_ket = st.text_area("📝 Rincian Pekerjaan / Kendala / Info Handover *:", placeholder="Jelaskan kronologi, tindakan tim, atau hal yang wajib dipantau shift berikutnya...", key="hd_ket")

            if st.button("🚀 SIMPAN MUTASI SiHD & SIAPKAN WA", type="primary", use_container_width=True, key="btn_save_sihd"):
                site_gabung = ", ".join(hd_sites + ([hd_site_manual.strip()] if hd_site_manual.strip() else []))
                if not hd_pic.strip() or not hd_tim or not hd_ket.strip():
                    st.error("⚠️ Nama Helpdesk, Tim/FME, dan Rincian Pekerjaan wajib diisi!")
                else:
                    with st.spinner("Menyimpan mutasi ke Google Sheets tab 'SiHD'..."):
                        rec_hd = {
                            "timestamp": datetime.datetime.now().strftime("%Y-%m-%d %H:%M:%S"),
                            "site_name": site_gabung or "-",
                            "kategori": "SiHD",
                            "nop": hd_nop,
                            "teknisi": ", ".join(hd_tim),
                            "status": hd_status,
                            "action": hd_ket.strip(),
                            "hd_shift": hd_shift,
                            "hd_pic": hd_pic.strip(),
                            "hd_kat_giat": hd_kat,
                            "hd_tiket": hd_tiket.strip() or "-",
                            "_sheet": "SiHD"
                        }
                        ws_hd = gsheet_get_or_create_ws("SiHD", ["Timestamp", "Site Name", "NOP", "Teknisi", "Status", "Data JSON"])
                        if ws_hd:
                            row_hd = [rec_hd["timestamp"], rec_hd["site_name"], rec_hd["nop"], rec_hd["teknisi"], rec_hd["status"], json.dumps(rec_hd, ensure_ascii=False)]
                            ws_hd.insert_row(row_hd, index=2)
                            fetch_data_from_gsheets.clear()
                            st.session_state['laporan_db'] = fetch_data_from_gsheets()
                            st.success("✅ Catatan mutasi SiHD berhasil disimpan ke Cloud!")
                        else:
                            get_laporan_db().insert(0, rec_hd)
                            st.warning("⚠️ Tersimpan di sesi lokal (koneksi Google Sheets sedang offline).")

                        wa_hd = (
                            f"*🚨 MUTASI PERGERAKAN TIM (SiHD)*\n"
                            f"📍 *NOP:* {hd_nop} | *Shift:* {hd_shift}\n"
                            f"🗓️ *Waktu:* {rec_hd['timestamp']}\n"
                            f"🎧 *Helpdesk:* {rec_hd['hd_pic']}\n"
                            f"👷 *Tim FME:* {rec_hd['teknisi']}\n"
                            f"🗼 *Site / Rute:* {rec_hd['site_name']}\n"
                            f"🏷️ *Kegiatan:* {hd_kat} (Tiket: {rec_hd['hd_tiket']})\n"
                            f"📊 *Status:* *{hd_status}*\n\n"
                            f"📝 *Keterangan / Handover:*\n{rec_hd['action']}"
                        )
                        st.link_button("📱 KIRIM LAPORAN MUTASI INI KE WHATSAPP", f"https://wa.me/?text={urllib.parse.quote(wa_hd)}", type="primary", use_container_width=True)

    with tab_hd_log:
        st.markdown("### 📋 Rekap Buku Mutasi & Generator WA Serah Terima Shift")
        if not db_sihd:
            st.info("Belum ada catatan mutasi tim pada tab 'SiHD'. Silakan input melalui tab pertama.")
        else:
            rows_hd = []
            for r in db_sihd:
                rows_hd.append({
                    "Waktu": r.get("timestamp", "-"),
                    "Tanggal": str(r.get("timestamp", ""))[:10],
                    "NOP": r.get("nop", "-"),
                    "Shift": r.get("hd_shift", "-"),
                    "Helpdesk": r.get("hd_pic", "-"),
                    "Tim FME": r.get("teknisi", "-"),
                    "Site / Rute": r.get("site_name", "-"),
                    "Kegiatan": r.get("hd_kat_giat", "-"),
                    "Tiket": r.get("hd_tiket", "-"),
                    "Status": r.get("status", "-"),
                    "Keterangan": r.get("action", "-")
                })
            df_hd = pd.DataFrame(rows_hd).sort_values("Waktu", ascending=False)

            cfh1, cfh2, cfh3 = st.columns(3)
            flt_hd_tgl = cfh1.selectbox("Filter Tanggal:", ["Semua Tanggal"] + sorted(df_hd["Tanggal"].unique().tolist(), reverse=True), key="flt_hd_tgl")
            flt_hd_nop = cfh2.selectbox("Filter NOP:", ["Semua NOP", "Palangkaraya", "Pangkalan Bun"], key="flt_hd_nop")
            flt_hd_shf = cfh3.selectbox("Filter Shift:", ["Semua Shift"] + sorted(df_hd["Shift"].unique().tolist()), key="flt_hd_shf")

            df_hd_v = df_hd.copy()
            if flt_hd_tgl != "Semua Tanggal": df_hd_v = df_hd_v[df_hd_v["Tanggal"] == flt_hd_tgl]
            if flt_hd_nop != "Semua NOP": df_hd_v = df_hd_v[df_hd_v["NOP"] == flt_hd_nop]
            if flt_hd_shf != "Semua Shift": df_hd_v = df_hd_v[df_hd_v["Shift"] == flt_hd_shf]

            st.dataframe(df_hd_v.drop(columns=["Tanggal"]), use_container_width=True, hide_index=True)

            # Generate WA Serah Terima Shift Gabungan
            wa_handover = f"*📋 SERAH TERIMA SHIFT HELPDESK (SiHD)*\n📍 Area: {flt_hd_nop} | Tanggal: {flt_hd_tgl} | Shift: {flt_hd_shf}\nTotal Pergerakan: {len(df_hd_v)} Aktivitas\n\n"
            for idx_h, (_, rh) in enumerate(df_hd_v.head(20).iterrows(), 1):
                wa_handover += f"{idx_h}. *[{rh['Status']}]* {rh['Site / Rute']} | Tim: {rh['Tim FME']}\n   ↳ {rh['Kegiatan']}: {rh['Keterangan']}\n"
            wa_handover += "\n_Generated via SIDATA SiHD_"

            c_hw1, c_hw2 = st.columns(2)
            with c_hw1:
                st.link_button("📱 KIRIM WA SERAH TERIMA SHIFT (GABUNGAN)", f"https://wa.me/?text={urllib.parse.quote(wa_handover)}", type="primary", use_container_width=True)
            with c_hw2:
                st.download_button("📥 Download Buku Mutasi SiHD (.csv)", data=df_hd_v.to_csv(index=False).encode("utf-8"), file_name=f"Buku_Mutasi_SiHD_{datetime.date.today()}.csv", mime="text/csv", use_container_width=True)

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
    db_live = get_laporan_db()
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
        ux_hero("AKUN SiSAVA", "Portal SiSAVA", "Sistem Pengajuan RAB, Import Excel Input All, Nilai SVA, Approval Berjenjang, dan Bukti Pekerjaan Lapangan.")
        
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

                if st.session_state.get("role") == "Admin":
                    st.markdown("<hr style='margin: 12px 0;'>", unsafe_allow_html=True)
                    if st.button("⚡ Masuk Cepat sebagai Admin Utama (Akses Input All & Manager)", type="secondary", use_container_width=True, key="svw_admin_quick"):
                        st.session_state["sv_user"] = {
                            "username": "admin_utama",
                            "nama": "Admin Utama SIDATA",
                            "jabatan": "Engineering Project Control",
                            "role": "MANAGER"
                        }
                        st.session_state["sv_nav"] = "📥 Import Excel (Input All)"
                        st.rerun()
    else:
        sisava_page()
# =========================================================================

# =========================================================================
# MENU 11: PORTAL OFFICE (6 SUB-MODUL TERORGANISIR: LOGISTIK, SiHD, SiSAVA, PROJECT, PROGRAM NOP, SiDAVID)
# =========================================================================
elif menu == "🏢 Portal Office":
    st.markdown("<h1 style='text-align: center;'>🏢 Portal Office SIDATA</h1>", unsafe_allow_html=True)
    st.markdown("<p style='text-align: center; opacity: 0.85;'>Pusat Kendali Back-Office: Logistik, Mutasi Helpdesk (SiHD), Keuangan RAB (SiSAVA), Executive Project, Program NOP & SiDAVID.</p>", unsafe_allow_html=True)
    st.markdown("<hr style='margin: 15px 0 28px 0;'>", unsafe_allow_html=True)

    c_off1, c_off2, c_off3 = st.columns(3)
    with c_off1:
        st.markdown("""
            <div class='area-card' style='border-top-color: #FF9800;'>
                <div style='font-size: 44px; margin-bottom: 10px;'>📦</div>
                <div class='login-title' style='font-size: 19px;'>Gudang & Logistik</div>
                <div class='login-desc' style='font-size: 13px;'>Pengelolaan Inbound/Outbound stok material, oli, filter genset, dan SPMS.</div>
            </div>
        """, unsafe_allow_html=True)
        if st.button("📦 BUKA LOGISTIK", type="primary", use_container_width=True, key="off_btn_log"):
            st.session_state['active_menu'] = "📦 Logistik & SPMS"
            st.rerun()

    with c_off2:
        st.markdown("""
            <div class='area-card' style='border-top-color: #EF4444;'>
                <div style='font-size: 44px; margin-bottom: 10px;'>🚨</div>
                <div class='login-title' style='font-size: 19px;'>SiHD (Sistem Helpdesk)</div>
                <div class='login-desc' style='font-size: 13px;'>Buku mutasi elektronik tim, rute harian & Auto-WA Serah Terima Shift.</div>
            </div>
        """, unsafe_allow_html=True)
        if st.button("🚨 BUKA SiHD", type="primary", use_container_width=True, key="off_btn_sihd"):
            st.session_state['active_menu'] = "🚨 SiHD (Pergerakan Tim)"
            st.rerun()

    with c_off3:
        st.markdown("""
            <div class='area-card' style='border-top-color: #10B981;'>
                <div style='font-size: 44px; margin-bottom: 10px;'>💠</div>
                <div class='login-title' style='font-size: 19px;'>Portal SiSAVA</div>
                <div class='login-desc' style='font-size: 13px;'>Pengajuan RAB (3E ➔ KUT ➔ HPP), Approval berjenjang, Nilai SVA & Bukti Kerja.</div>
            </div>
        """, unsafe_allow_html=True)
        if st.button("💠 BUKA SiSAVA", type="primary", use_container_width=True, key="off_btn_sisava"):
            st.session_state['active_menu'] = "💠 SiSAVA"
            st.rerun()

    st.markdown("<br>", unsafe_allow_html=True)
    c_off4, c_off5, c_off6 = st.columns(3)
    with c_off4:
        st.markdown("""
            <div class='area-card' style='border-top-color: #FFD700; background: linear-gradient(145deg, #0f172a, #1e293b);'>
                <div style='font-size: 44px; margin-bottom: 10px;'>💎</div>
                <div class='login-title' style='font-size: 19px; color: #FFD700 !important;'>Executive Project</div>
                <div class='login-desc' style='font-size: 13px; color: #cbd5e1 !important;'>Dashboard Khusus PO & SVA: Budget vs Support, Timeline & Analisa Mandek.</div>
            </div>
        """, unsafe_allow_html=True)
        if st.button("💎 BUKA EXECUTIVE PROJECT", type="primary", use_container_width=True, key="off_btn_proj"):
            st.session_state['active_menu'] = "💎 Executive Project Dashboard"
            st.rerun()

    with c_off5:
        st.markdown("""
            <div class='area-card' style='border-top-color: #38BDF8;'>
                <div style='font-size: 44px; margin-bottom: 10px;'>📋</div>
                <div class='login-title' style='font-size: 19px;'>Program NOP & Task Force</div>
                <div class='login-desc' style='font-size: 13px;'>Upload & simpan program NOP PLK/PBU, checklist target, availability & history power.</div>
            </div>
        """, unsafe_allow_html=True)
        if st.button("📋 BUKA PROGRAM NOP", type="primary", use_container_width=True, key="off_btn_prognop"):
            st.session_state['active_menu'] = "📋 Program NOP & Task Force"
            st.rerun()

    with c_off6:
        st.markdown("""
            <div class='area-card' style='border-top-color: #8B5CF6;'>
                <div style='font-size: 44px; margin-bottom: 10px;'>🛡️</div>
                <div class='login-title' style='font-size: 19px;'>SiDAVID (Verifikasi Internal)</div>
                <div class='login-desc' style='font-size: 13px;'>Verifikasi Mateline, Sertifikat WAH/BA/BE, Join/Keluar Tim & Validasi Standby.</div>
            </div>
        """, unsafe_allow_html=True)
        if st.button("🛡️ BUKA SiDAVID", type="primary", use_container_width=True, key="off_btn_sidavid"):
            st.session_state['active_menu'] = "🛡️ SiDAVID (Verifikasi Internal)"
            st.rerun()

    st.markdown("<br>", unsafe_allow_html=True)
    c_off7, c_off8, c_off9 = st.columns(3)
    with c_off7:
        st.markdown("""
            <div class='area-card' style='border-top-color: #F97316;'>
                <div style='font-size: 44px; margin-bottom: 10px;'>🚧</div>
                <div class='login-title' style='font-size: 19px;'>Form & Peta Hard Akses</div>
                <div class='login-desc' style='font-size: 13px;'>Report & Peta Rute Hard Akses PBU & PLK, Moda Wajib (Trail/4WD/Klotok) & Eviden Foto.</div>
            </div>
        """, unsafe_allow_html=True)
        if st.button("🚧 BUKA FORM HARD AKSES", type="primary", use_container_width=True, key="off_btn_hardakses"):
            st.session_state['active_menu'] = "🚧 Form & Peta Hard Akses"
            st.rerun()

    with c_off8:
        st.markdown("""
            <div class='area-card' style='border-top-color: #9C27B0;'>
                <div style='font-size: 44px; margin-bottom: 10px;'>📅</div>
                <div class='login-title' style='font-size: 19px;'>Daily Tracker PLK</div>
                <div class='login-desc' style='font-size: 13px;'>Planning Harian Leader Cluster, Pengalihan Tier Urgent & Auto-Closed Report SIDATA.</div>
            </div>
        """, unsafe_allow_html=True)
        if st.button("📅 BUKA DAILY TRACKER PLK", type="primary", use_container_width=True, key="off_btn_dailyplk"):
            st.session_state['active_menu'] = "📅 Daily Tracker Palangkaraya"
            st.rerun()

    with c_off9:
        st.markdown("""
            <div class='area-card' style='border-top-color: #0EA5E9;'>
                <div style='font-size: 44px; margin-bottom: 10px;'>👥</div>
                <div class='login-title' style='font-size: 19px;'>Kelola Master Tim</div>
                <div class='login-desc' style='font-size: 13px;'>Admin mandiri tambah/hapus tim baru & lama beserta Role untuk Master PLK & PBU.</div>
            </div>
        """, unsafe_allow_html=True)
        if st.button("👥 KELOLA MASTER TIM", type="primary", use_container_width=True, key="off_btn_mastertim"):
            st.session_state['active_menu'] = "👥 Kelola Master Tim (PLK & PBU)"
            st.rerun()


# =========================================================================
# MENU 12: PROGRAM NOP & TASK FORCE (FAST LAZY LOAD + HARD ACCESS PBU + TIER MAPS)
# =========================================================================
elif menu == "📋 Program NOP & Task Force":
    st.markdown("<h1 style='text-align: center;'>📋 Program NOP & Task Force Analytics</h1>", unsafe_allow_html=True)
    st.markdown("<p style='text-align: center; opacity: 0.85;'>Tier Route Maps (HB ➔ Target & Tier #1/#2/#3), Report Hard Access + Foto (Admin), Checklist Target & Sinkronisasi History Onsite SIDATA</p>", unsafe_allow_html=True)

    pilih_area_prog = st.radio(
        "📍 Pilih Area Program NOP (Mode Ringan — hanya memuat data area yang dipilih):",
        ["🔴 NOP Pangkalan Bun (Tier Maps, Hard Access & Timeplan)", "🔵 NOP Palangkaraya (Task Force & Availability)"],
        horizontal=True,
        key="prog_nop_area_pick"
    )
    st.markdown("<hr style='margin: 8px 0 18px 0;'>", unsafe_allow_html=True)

    def _enrich_with_sidata(df_in, site_col_name, onsite_idx):
        df_e = df_in.copy()
        cnts, ldates, lstats, lprobs, lvacs, lvdcs, lloads = [], [], [], [], [], [], []
        for _, rw in df_e.iterrows():
            sid_n = _norm_site_id(rw.get(site_col_name, ""))
            hist = onsite_idx.get(sid_n, [])
            if hist:
                last_h = hist[-1]
                cnts.append(len(hist))
                ldates.append(last_h["Tanggal"])
                lstats.append(last_h["Status"])
                lprobs.append(str(last_h["Problem"])[:80])
                lvacs.append(f"{last_h['V_AC']} V" if last_h["V_AC"] is not None else "-")
                lvdcs.append(f"{last_h['V_DC']} V" if last_h["V_DC"] is not None else "-")
                lloads.append(f"{last_h['Load_DC']} A" if last_h["Load_DC"] is not None else "-")
            else:
                cnts.append(0); ldates.append("-"); lstats.append("Belum Onsite"); lprobs.append("-"); lvacs.append("-"); lvdcs.append("-"); lloads.append("-")
        df_e["Onsite_SIDATA (x)"] = cnts
        df_e["Tgl_Onsite_Terakhir"] = ldates
        df_e["Status_Onsite_SIDATA"] = lstats
        df_e["Problem_Onsite_Terakhir"] = lprobs
        df_e["Last_VAC"] = lvacs
        df_e["Last_VDC"] = lvdcs
        df_e["Last_Load_DC"] = lloads
        return df_e

    def _render_site_history_inspector(df_prog_base, site_col_name, key_prefix, onsite_idx):
        st.markdown("#### 🔍 Inspector History Onsite & Grafik Power Site Program (Dari Database SIDATA)")
        sids_all = sorted(df_prog_base[site_col_name].dropna().astype(str).str.strip().str.upper().unique().tolist())
        if not sids_all:
            st.info("Belum ada Site ID pada daftar program.")
            return
        pick_s = st.selectbox(
            "📌 Pilih Site ID dari Program untuk melihat riwayat onsite & tren power:",
            sids_all,
            format_func=lambda x: f"{x} ({len(onsite_idx.get(_norm_site_id(x), []))}x Record Onsite di SIDATA)",
            key=f"insp_site_{key_prefix}"
        )
        sid_norm = _norm_site_id(pick_s)
        hist_list = onsite_idx.get(sid_norm, [])
        if not hist_list:
            st.warning(f"Site **{pick_s}** belum memiliki record kunjungan onsite di database SIDATA.")
        else:
            df_h = pd.DataFrame(hist_list)
            st.success(f"✅ Ditemukan **{len(df_h)}x kunjungan onsite** untuk site **{pick_s}** di database SIDATA.")
            ch_p1, ch_p2 = st.columns(2)
            with ch_p1:
                with st.container(border=True):
                    st.markdown("**⚡ Riwayat Tegangan AC / DC (Volt)**")
                    df_hv = df_h.set_index("Tanggal")[["V_AC", "V_DC"]].dropna(how="all")
                    if not df_hv.empty:
                        ui_line_chart(df_hv, color=["#EF4444", "#10B981"])
                    else:
                        st.caption("Belum ada angka pengukuran tegangan pada riwayat onsite site ini.")
            with ch_p2:
                with st.container(border=True):
                    st.markdown("**🔋 Riwayat Beban DC (Ampere)**")
                    df_hl = df_h.set_index("Tanggal")[["Load_DC"]].dropna(how="all")
                    if not df_hl.empty:
                        ui_bar_chart(df_hl, color="#F59E0B")
                    else:
                        st.caption("Belum ada angka pengukuran beban DC pada riwayat onsite site ini.")
            st.dataframe(df_h[["Timestamp", "Kategori", "Teknisi", "Status", "Problem", "Remark", "V_AC", "V_DC", "Load_DC"]], use_container_width=True, hide_index=True)

    # =====================================================================
    # AREA 1: NOP PANGKALAN BUN (TIER MAPS, HARD ACCESS ADMIN & ANALISA)
    # =====================================================================
    if "Pangkalan Bun" in pilih_area_prog:
        COLS_PROG_PBU = [
            "No", "SITE ID", "SITE NAME", "PIC TASK", "Tier Tim #1", "Tier Tim #2", "Tier Tim #3",
            "SITE CLASS", "RCA", "Action Detail", "Plan Date", "Target Closed", "Status Activity",
            "Biaya (Diisi Koordinator)", "Remark (Target & Plan)",
            "Tracker #1 Tgl", "Tracker #1 Site ID", "Remark",
            "Tracker #2 Tgl", "Tracker #2 Site ID", "Remark_1",
            "Tracker #3 Tgl", "Tracker #3 Site ID", "Remark_2",
            "Remark Hasil Pekerjaan"
        ]
        df_pbu_cloud, ok_pbu_c = gsheet_read_df("Program_NOP_PBU", COLS_PROG_PBU)
        local_pbu_file = _find_first_existing(
            ["Timeplan_Target_Closed_Biaya_Nov2026 (1).xlsx", "Timeplan_Target_Closed_Biaya_Nov2026 pangkalanbun.xlsx"],
            "*Timeplan_Target_Closed_Biaya*.xlsx"
        )
        if df_pbu_cloud.empty and local_pbu_file:
            try:
                df_pbu_cloud = parse_excel_smart_header(local_pbu_file, "Timeplan & Tracking")
            except Exception:
                pass

        df_site_m_pbu, df_hb_m_pbu, site_coords_pbu, fme_coords_pbu, src_site_pbu, src_hb_pbu = load_pbu_coords_master()
        df_ha_pbu, ok_ha_pbu, ha_list_pbu, ha_by_site_pbu = load_pbu_hard_access_records(site_coords_pbu)

        with st.expander("📤 Upload & Update Data Cloud Pangkalan Bun (Timeplan Tier + LongLat Site + LongLat HB Tim)", expanded=df_pbu_cloud.empty):
            st.caption(f"📡 Sumber Koordinat Aktif — **Site Tujuan:** `{src_site_pbu}` ({len(site_coords_pbu)} site) | **HB Tim:** `{src_hb_pbu}` ({len(fme_coords_pbu)} personel) | **Hard Access:** `{len(ha_list_pbu)} lokasi`")
            u_t1, u_t2, u_t3 = st.tabs([
                "📋 1. Upload Timeplan & Tier Tracking",
                "📍 2. Update LongLat Site (`Site ID Sidata Pangkalanbun`)",
                "🏠 3. Update LongLat HB Tim (`NAMA NOP PANGKALANBUN`)"
            ])

            with u_t1:
                up_pbu = st.file_uploader("Unggah File Excel Timeplan & Tier Pangkalan Bun (.xlsx / .csv):", type=["xlsx", "xls", "csv"], key="up_prog_pbu")
                mode_pbu = st.radio("Mode Simpan ke Cloud (`Program_NOP_PBU`):", ["Replace (Ganti List Program)", "Append (Tambah ke List Lama)"], horizontal=True, key="mode_prog_pbu")
                if up_pbu is not None:
                    try:
                        if up_pbu.name.lower().endswith(".csv"):
                            df_new_pbu = pd.read_csv(up_pbu)
                        else:
                            sheets_pbu = get_excel_sheet_names_safe(up_pbu)
                            def_sh_b = sheets_pbu.index("Timeplan & Tracking") if "Timeplan & Tracking" in sheets_pbu else 0
                            sh_pbu = st.selectbox("Pilih Sheet sumber:", sheets_pbu, index=def_sh_b, key="sh_prog_pbu")
                            df_new_pbu = parse_excel_smart_header(up_pbu, sh_pbu)
                        st.write(f"Preview ({len(df_new_pbu)} task):")
                        st.dataframe(df_new_pbu.head(10), use_container_width=True, hide_index=True)
                        if st.button("💾 SIMPAN TIMEPLAN PANGKALAN BUN KE CLOUD", type="primary", use_container_width=True, key="btn_save_prog_pbu"):
                            df_save_b = pd.concat([df_pbu_cloud, df_new_pbu], ignore_index=True) if ("Append" in mode_pbu and not df_pbu_cloud.empty) else df_new_pbu
                            if gsheet_overwrite_ws("Program_NOP_PBU", df_save_b):
                                st.success(f"✅ {len(df_save_b)} task Program NOP Pangkalan Bun tersimpan ke tab `Program_NOP_PBU`!")
                                st.rerun()
                            else:
                                st.error("⚠️ Gagal menyimpan ke Google Sheets.")
                    except Exception as e:
                        st.error(f"Gagal membaca Excel Timeplan Pangkalan Bun: {e}")

            with u_t2:
                st.info("💡 Unggah file `Site ID Sidata Pangkalanbun.xlsx` (kolom: `Site ID`, `Latitude Tujuan`, `Longtitude Tujuan`) untuk memperbarui koordinat site tujuan.")
                up_site_ll = st.file_uploader("Upload Excel LongLat Site Tujuan PBU:", type=["xlsx", "xls", "csv"], key="up_site_ll_pbu")
                if up_site_ll is not None:
                    try:
                        df_s_up = pd.read_csv(up_site_ll) if up_site_ll.name.lower().endswith(".csv") else read_excel_safe_raw(up_site_ll, sheet_name=0, header=0)
                        df_s_up = df_s_up.dropna(how="all").reset_index(drop=True)
                        st.dataframe(df_s_up.head(10), use_container_width=True, hide_index=True)
                        if st.button("💾 SIMPAN MASTER LONGLAT SITE PBU KE CLOUD (`Master_Site_PBU`)", type="primary", use_container_width=True, key="btn_save_site_ll_pbu"):
                            if gsheet_overwrite_ws("Master_Site_PBU", df_s_up):
                                st.success(f"✅ {len(df_s_up)} koordinat Site Pangkalan Bun berhasil diperbarui di Cloud!")
                                st.rerun()
                    except Exception as e:
                        st.error(f"Gagal membaca file Site LongLat: {e}")

            with u_t3:
                st.info("💡 Unggah file `NAMA NOP PANGKALANBUN.xlsx` (kolom: `NAMA`, `NOP`, `Latitude Berangkat`, `Longtitude Berangkat`) untuk memperbarui koordinat Homebase (HB) Tim.")
                up_hb_ll = st.file_uploader("Upload Excel LongLat HB Tim PBU:", type=["xlsx", "xls", "csv"], key="up_hb_ll_pbu")
                if up_hb_ll is not None:
                    try:
                        df_hb_up = pd.read_csv(up_hb_ll) if up_hb_ll.name.lower().endswith(".csv") else read_excel_safe_raw(up_hb_ll, sheet_name=0, header=0)
                        df_hb_up = df_hb_up.dropna(how="all").reset_index(drop=True)
                        st.dataframe(df_hb_up.head(10), use_container_width=True, hide_index=True)
                        if st.button("💾 SIMPAN MASTER HB TIM PBU KE CLOUD (`Master_HB_Tim_PBU`)", type="primary", use_container_width=True, key="btn_save_hb_ll_pbu"):
                            if gsheet_overwrite_ws("Master_HB_Tim_PBU", df_hb_up):
                                st.success(f"✅ {len(df_hb_up)} koordinat HB Tim Pangkalan Bun berhasil diperbarui di Cloud!")
                                st.rerun()
                    except Exception as e:
                        st.error(f"Gagal membaca file HB Tim: {e}")

        # Hanya muat history onsite jika sudah di sesi atau saat sub-menu History dibuka
        onsite_idx_pbu = build_onsite_history_index(st.session_state.get("laporan_db", []))
        if not df_pbu_cloud.empty:
            scol_pbu = next((c for c in df_pbu_cloud.columns if "site id" in c.lower() and "tracker" not in c.lower()), df_pbu_cloud.columns[0])
            stcol_pbu = next((c for c in df_pbu_cloud.columns if "status" in c.lower()), None)
            df_pbu_routes, map_routes_pbu = analyze_pbu_timeplan_routes(
                df_pbu_cloud, site_coords_pbu, fme_coords_pbu,
                onsite_idx=onsite_idx_pbu, hard_access_by_site=ha_by_site_pbu
            )
            tot_pbu = len(df_pbu_routes)
            cnt_closed_pbu = int(df_pbu_routes["Status"].astype(str).str.lower().str.contains("close|done|selesai").sum())
            cnt_inprog_pbu = int(df_pbu_routes["Status"].astype(str).str.lower().str.contains("progress").sum())
            cnt_open_pbu = tot_pbu - cnt_closed_pbu - cnt_inprog_pbu
            pct_pbu = (cnt_closed_pbu / tot_pbu * 100) if tot_pbu else 0.0
            cnt_divert_pbu = int(df_pbu_routes["Analisa Target / Divert"].astype(str).str.contains("Dialihkan").sum())
            cnt_ha_pbu = int(df_pbu_routes["Status Hard Access"].astype(str).str.contains("🚧|⚠️").sum())
            tot_km_plan_pbu = float(df_pbu_routes["HB->Target (km)"].sum())
            tot_km_act_pbu = float(df_pbu_routes["Total Jarak Aktual (km)"].sum())
            tot_biaya_pbu = float(df_pbu_routes["Biaya (Rp)"].sum())

            kpb1, kpb2, kpb3, kpb4, kpb5, kpb6 = st.columns(6)
            kpb1.metric("Total Site Target", f"{tot_pbu} Site")
            kpb2.metric("✅ Closed", f"{cnt_closed_pbu} ({pct_pbu:.1f}%)")
            kpb3.metric("🔀 Divert / Open", f"{cnt_divert_pbu} Divert / {cnt_open_pbu} Open")
            kpb4.metric("🚧 Hard Access", f"{cnt_ha_pbu} Task", delta=f"{len(ha_list_pbu)} Laporan Lokasi")
            kpb5.metric("🚗 Total Jarak Aktual", f"{tot_km_act_pbu:,.1f} km", delta=f"Est. Jalan {tot_km_act_pbu*1.35:,.0f} km")
            kpb6.metric("💰 Total Biaya", f"Rp {tot_biaya_pbu:,.0f}")
        else:
            scol_pbu = "SITE ID"; stcol_pbu = "Status Activity"
            df_pbu_routes, map_routes_pbu = pd.DataFrame(), []
            tot_pbu = cnt_closed_pbu = cnt_inprog_pbu = cnt_open_pbu = cnt_divert_pbu = cnt_ha_pbu = 0
            pct_pbu = tot_km_plan_pbu = tot_km_act_pbu = tot_biaya_pbu = 0.0

        # Sub-navigasi Radio (Hanya merender 1 modul aktif -> SUPER RINGAN!)
        sub_pbu_nav = st.radio(
            "Sub-Menu Program NOP Pangkalan Bun:",
            [
                "🗺️ Peta Akses & Tier Maps",
                "🚧 Report Hard Access (Admin & Foto)",
                "📊 Analisa Manajemen & Export",
                "✏️ Live Editor Timeplan & Tier",
                "⚡ History Onsite & Power SIDATA"
            ],
            horizontal=True,
            key="pbu_sub_nav",
            label_visibility="collapsed"
        )
        st.markdown("<hr style='margin: 8px 0 18px 0;'>", unsafe_allow_html=True)

        # -----------------------------------------------------------------
        # SUB-MODUL 1: PETA AKSES & TIER MAPS (GARIS MELENGKUNG + PANAH + HARD ACCESS)
        # -----------------------------------------------------------------
        if sub_pbu_nav == "🗺️ Peta Akses & Tier Maps":
            if df_pbu_routes.empty and not ha_list_pbu:
                st.warning("Belum ada data Timeplan maupun Report Hard Access Pangkalan Bun.")
            else:
                st.markdown("#### 🗺️ Tier Route & Hard Access Command Map — NOP Pangkalan Bun")
                st.info("💡 **Garis Melengkung Putus-Putus Oranye** = Rute Plan (`HB Tim ➔ Site Target`). **Garis Melengkung Neon + Panah Bergerak (`T1 Cyan ➔ T2 Ungu ➔ T3 Hijau`)** = Tarikan Aktual Kunjungan Tim per Tier. **Pin 🚧** = Titik Hard Access lengkap dengan preview foto kondisi jalan/akses.")

                cm_c1, cm_c2, cm_c3 = st.columns([1.6, 1.2, 1.4])
                with cm_c1:
                    mode_peta_pbu = st.radio(
                        "Mode Fokus Peta:",
                        ["🎯 Fokus 1 Task", "👷 Per Tim (Tier #1)", "🚧 Fokus Hard Access", "🌐 Semua Rute"],
                        horizontal=True,
                        key="pbu_map_mode"
                    )
                with cm_c2:
                    tema_peta_pbu = st.selectbox(
                        "🛰️ Base Layer Peta:",
                        ["Dark NOC Command Center", "Citra Satelit (Esri)", "Peta Jalan Terang (Street Map)"],
                        key="pbu_map_theme"
                    )
                with cm_c3:
                    cl_1, cl_2 = st.columns(2)
                    chk_plan_line = cl_1.checkbox("📐 Garis Plan", value=True, key="pbu_chk_plan")
                    chk_act_line = cl_2.checkbox("🚗 Garis Tier", value=True, key="pbu_chk_act")
                    chk_ha_pins = cl_1.checkbox("🚧 Pin Hard Access", value=True, key="pbu_chk_ha")
                    chk_dist_lbl = cl_2.checkbox("🏷️ Label km", value=True, key="pbu_chk_lbl")

                selected_routes = map_routes_pbu
                selected_ha = ha_list_pbu
                focus_route = map_routes_pbu[0] if map_routes_pbu else None

                if mode_peta_pbu == "🎯 Fokus 1 Task" and map_routes_pbu:
                    pick_task_key = st.selectbox("📌 Pilih Task / Site Target (Auto-Zoom ke rute tim):", [r["key"] for r in map_routes_pbu], key="pbu_pick_task_map")
                    selected_routes = [r for r in map_routes_pbu if r["key"] == pick_task_key]
                    focus_route = selected_routes[0] if selected_routes else None
                elif mode_peta_pbu == "👷 Per Tim (Tier #1)" and map_routes_pbu:
                    tim_list_pbu = sorted(list({r["tim1"] for r in map_routes_pbu if r["tim1"] != "-"}))
                    pick_tim_map = st.selectbox("👷 Pilih Nama Tim (Tier Tim #1):", tim_list_pbu, key="pbu_pick_tim_map")
                    selected_routes = [r for r in map_routes_pbu if r["tim1"] == pick_tim_map]
                    focus_route = selected_routes[0] if selected_routes else None
                elif mode_peta_pbu == "🚧 Fokus Hard Access":
                    ha_sids_set = set(ha_by_site_pbu.keys())
                    selected_routes = [r for r in map_routes_pbu if (r["target_sid"] in ha_sids_set) or ("🚧" in r.get("hard_access_status", "")) or ("⚠️" in r.get("hard_access_status", ""))]
                    focus_route = selected_routes[0] if selected_routes else None
                    if not selected_routes and not ha_list_pbu:
                        st.info("Belum ada data Hard Access yang tercatat. Admin dapat menginputnya di tab **🚧 Report Hard Access (Admin & Foto)**.")

                c_mp_l, c_mp_r = st.columns([2.25, 1.25])
                with c_mp_l:
                    st.components.v1.html(
                        render_pbu_route_map_html(
                            selected_routes,
                            hard_access_list=selected_ha,
                            map_theme=tema_peta_pbu,
                            show_plan=chk_plan_line,
                            show_actual=chk_act_line,
                            show_hard_access=chk_ha_pins,
                            show_dist_labels=chk_dist_lbl
                        ),
                        height=615
                    )

                with c_mp_r:
                    with st.container(border=True):
                        st.markdown("#### 📌 Kronologi Rute, Remark & Hard Access")
                        if mode_peta_pbu != "🎯 Fokus 1 Task" and len(selected_routes) > 1:
                            pick_side = st.selectbox("Detail Task:", [r["key"] for r in selected_routes], key="pbu_side_detail_pick")
                            focus_route = next((r for r in selected_routes if r["key"] == pick_side), selected_routes[0])

                        if focus_route:
                            st.markdown(f"""
                                <div style="padding:12px 14px; border-radius:12px; background:var(--ux-card2); border-left:5px solid #EF4444; font-size:13px; line-height:1.6; margin-bottom:10px;">
                                    <div style="font-size:15px; font-weight:800;">🎯 #{focus_route['no']} · {focus_route['target_sid']} ({focus_route['s_class']})</div>
                                    <div><b>Site Name:</b> {focus_route['sname']}</div>
                                    <div><b>Status:</b> <b>{focus_route['status']}</b> | <b>Biaya:</b> Rp {focus_route['biaya']:,.0f}</div>
                                    <div><b>Akses:</b> {focus_route.get('hard_access_status', '✅ Normal')}</div>
                                    <div><b>Action:</b> {focus_route['rca']} — {focus_route['action']}</div>
                                    <div><b>Plan ➔ Target:</b> {focus_route['plan_dt']} ➔ {focus_route['target_cl']}</div>
                                    <hr style="margin:6px 0;">
                                    <div>🏠 <b>HB Tim #1:</b> {focus_route['tim1']}</div>
                                    <div>📏 <b>Jarak HB ➔ Target:</b> <b>{focus_route['dist_hb_tgt'] if focus_route['dist_hb_tgt'] is not None else '-'} km</b> (Est. Jalan: {focus_route['est_road_tgt'] if focus_route['est_road_tgt'] is not None else '-'} km)</div>
                                    <div>🚗 <b>Total Jarak Aktual:</b> <b>{focus_route['total_act_km']} km</b> (Est. Jalan: {focus_route['est_road_act']} km)</div>
                                    <div style="margin-top:4px;"><b>🚦 Analisa:</b> {focus_route['analisa']}</div>
                                </div>
                            """, unsafe_allow_html=True)

                            st.markdown("**👣 Jejak Kunjungan per Tier:**")
                            if focus_route["waypoints"]:
                                tier_col_map = {1: "#06b6d4", 2: "#a855f7", 3: "#10b981"}
                                for wp in focus_route["waypoints"]:
                                    tc_w = tier_col_map.get(wp["tier"], "#06b6d4")
                                    st.markdown(
                                        f"<div style='padding:8px 10px; border-radius:8px; background:rgba(6,182,212,0.08); border-left:4px solid {tc_w}; margin-bottom:6px; font-size:12.5px;'>"
                                        f"<b>Tier #{wp['tier']} ({wp['tgl']}) ➔ {wp['site_id']}</b><br>"
                                        f"📏 Jarak: <b>{wp['leg_km']} km</b> (Est. Jalan {wp['est_road_km']} km) dari {wp['from_label']}<br>"
                                        f"🔧 <b>Remark:</b> {wp['remark']}</div>",
                                        unsafe_allow_html=True
                                    )
                            else:
                                st.caption("Belum ada titik kunjungan Tier berkoordinat untuk task ini.")

                            if focus_route["notes_non_site"]:
                                for nt in focus_route["notes_non_site"]:
                                    st.warning(f"🚧 {nt}")

                            if focus_route["rem_hasil"] and focus_route["rem_hasil"] != "-":
                                st.info(f"📝 **Remark Hasil Pekerjaan:** {focus_route['rem_hasil']}")

                            # Tampilkan foto Hard Access langsung jika site target terdaftar di Report Hard Access!
                            ha_for_focus = ha_by_site_pbu.get(focus_route["target_sid"], [])
                            if ha_for_focus:
                                st.markdown("**🚧 Dokumentasi Hard Access Site Ini:**")
                                for ha_f in ha_for_focus:
                                    st.caption(f"📍 {ha_f['lokasi']} | {ha_f['kategori']} | {ha_f['kendaraan']}")
                                    if ha_f["photos"]:
                                        tampilkan_grid_foto(ha_f["photos"], f"Foto Kondisi Akses ({ha_f['timestamp']})")

        # -----------------------------------------------------------------
        # SUB-MODUL 2: REPORT HARD ACCESS SITE PANGKALAN BUN (INPUT ADMIN + FOTO)
        # -----------------------------------------------------------------
        elif sub_pbu_nav == "🚧 Report Hard Access (Admin & Foto)":
            st.markdown("#### 🚧 Report Kondisi Akses Site & Lokasi Hard Access — NOP Pangkalan Bun")
            st.info("💡 Modul khusus pelaporan lokasi-lokasi **Hard Access** (jalan amblas/lumpur, jembatan putus, perbaikan jalan, banjir/akses perahu, izin corporate) lengkap dengan titik GPS dan **Foto Eviden** yang langsung muncul di Peta Tier Maps & Laporan Manajemen.")

            is_admin_ha = (st.session_state.get("role") == "Admin") or bool(st.session_state.get("pbu_ha_admin_unlocked", False))

            if not is_admin_ha:
                with st.expander("🔐 Login Admin untuk Input / Edit Report Hard Access", expanded=False):
                    c_ha_l1, c_ha_l2 = st.columns([2, 1])
                    pwd_ha = c_ha_l1.text_input("Masukkan Kata Sandi Admin:", type="password", key="pbu_ha_pwd", label_visibility="collapsed", placeholder="Kata sandi Admin...")
                    if c_ha_l2.button("🔓 Buka Mode Input Admin", type="primary", use_container_width=True, key="btn_unlock_ha_adm"):
                        if pwd_ha == "KUT2027":
                            st.session_state["pbu_ha_admin_unlocked"] = True
                            st.rerun()
                        else:
                            st.error("❌ Kata sandi Admin salah!")
            else:
                with st.expander("➕ Input Report Hard Access Site Baru (Admin + Upload Foto)", expanded=df_ha_pbu.empty):
                    all_pbu_sids = sorted(list(site_coords_pbu.keys()))
                    # Prioritaskan site yang ada di Timeplan di bagian paling atas dropdown
                    tp_sids = [r["target_sid"] for r in map_routes_pbu] if map_routes_pbu else []
                    ordered_sids = list(dict.fromkeys(tp_sids + all_pbu_sids))

                    cha1, cha2, cha3 = st.columns(3)
                    with cha1:
                        ha_sid_pick = st.selectbox("📍 Pilih Site ID Pangkalan Bun *:", ["-- Ketik Manual --"] + ordered_sids, key="ha_sid_pick")
                        if ha_sid_pick == "-- Ketik Manual --":
                            ha_sid = st.text_input("Ketik Site ID *:", key="ha_sid_man").strip().upper()
                        else:
                            ha_sid = ha_sid_pick
                        ha_sname_def = next((r["sname"] for r in map_routes_pbu if r["target_sid"] == ha_sid), ha_sid)
                        ha_sname = st.text_input("Nama Site:", value=ha_sname_def, key="ha_sname")
                        ha_lokasi = st.text_input("🗺️ Lokasi / Titik Hard Access *:", placeholder="Cth: Km 12 Jl. Poros Kuala Kuayan (Perbaikan Jalan)", key="ha_lokasi")

                    with cha2:
                        ha_kat = st.selectbox("🏷️ Kategori Hard Access *:", HARD_ACCESS_KATEGORI, key="ha_kat")
                        ha_lvl = st.selectbox("⚠️ Tingkat Kesulitan *:", HARD_ACCESS_LEVEL, key="ha_lvl")
                        ha_stat = st.selectbox("🚦 Status Akses Saat Ini *:", HARD_ACCESS_STATUS, key="ha_stat")

                    with cha3:
                        def_lat, def_lon = site_coords_pbu.get(ha_sid, (-2.35, 111.60))
                        ha_lat = st.number_input("Latitude Lokasi Kendala:", value=float(def_lat), format="%.6f", key=f"ha_lat_{ha_sid}")
                        ha_lon = st.number_input("Longitude Lokasi Kendala:", value=float(def_lon), format="%.6f", key=f"ha_lon_{ha_sid}")
                        ha_kend = st.selectbox("🚙 Kendaraan / Moda Wajib:", ["4x4 / Hilux Double Cabin", "Motor Trail", "Perahu / Kelotok + Darat", "Jalan Kaki + Porter", "Izin Khusus Corporate", "Mobil Standar"], key="ha_kend")

                    cha4, cha5 = st.columns([1.3, 1.7])
                    ha_est = cha4.text_input("⏱️ Estimasi Waktu Tempuh / Tambahan Biaya:", placeholder="Cth: 5 Jam PP / Sewa Perahu Rp 500rb", key="ha_est")
                    ha_pic = cha5.text_input("👨‍💼 Nama Admin / Koordinator Pelapor *:", value="Admin Pangkalan Bun", key="ha_pic")
                    ha_ket = st.text_area("📝 Keterangan Detail Kondisi Akses & Mitigasi *:", placeholder="Jelaskan kondisi jalan, titik rusak, rekomendasi rute alternatif, atau kebutuhan alat/izin...", key="ha_ket")
                    ha_photos_up = st.file_uploader("📸 Upload Foto-Foto Kondisi Lokasi Hard Access (Bisa Banyak Foto):", type=["png", "jpg", "jpeg", "webp"], accept_multiple_files=True, key="ha_photos_up")

                    if st.button("🚀 SIMPAN REPORT HARD ACCESS & UPLOAD FOTO KE CLOUD", type="primary", use_container_width=True, key="btn_save_ha_pbu"):
                        if not ha_sid or not ha_lokasi.strip() or not ha_ket.strip():
                            st.error("⚠️ Site ID, Lokasi/Titik Hard Access, dan Keterangan Detail wajib diisi!")
                        else:
                            with st.spinner("⏳ Mengunggah foto ke Cloudinary & menyimpan ke Google Sheets (`Hard_Access_PBU`)..."):
                                urls_ha = upload_multiple_images(ha_photos_up, "hard_access_pbu") if ha_photos_up else []
                                rec_id = f"HA-PBU-{datetime.datetime.now().strftime('%Y%m%d%H%M%S')}"
                                row_ha = {
                                    "Record_ID": rec_id,
                                    "Timestamp": datetime.datetime.now().strftime("%Y-%m-%d %H:%M:%S"),
                                    "Site_ID": ha_sid,
                                    "Site_Name": ha_sname.strip() or ha_sid,
                                    "Lokasi_Kendala": ha_lokasi.strip(),
                                    "Kategori_Hard_Access": ha_kat,
                                    "Tingkat_Kesulitan": ha_lvl,
                                    "Status_Akses": ha_stat,
                                    "Lat_Kendala": f"{ha_lat:.6f}",
                                    "Lon_Kendala": f"{ha_lon:.6f}",
                                    "Estimasi_Tempuh": ha_est.strip() or "-",
                                    "Kendaraan_Wajib": ha_kend,
                                    "PIC_Input": ha_pic.strip() or "Admin",
                                    "Keterangan_Detail": ha_ket.strip(),
                                    "Foto_URLs": json.dumps(urls_ha, ensure_ascii=False)
                                }
                                row_sha_pbu = {
                                    "Record_ID": rec_id,
                                    "Timestamp": row_ha["Timestamp"],
                                    "NOP": "Pangkalan Bun",
                                    "Site_ID": ha_sid,
                                    "Site_Name": row_ha["Site_Name"],
                                    "Titik_Start_Pos": "-",
                                    "Lat_Start": "",
                                    "Lon_Start": "",
                                    "Lokasi_Kendala": row_ha["Lokasi_Kendala"],
                                    "Lat_Kendala": row_ha["Lat_Kendala"],
                                    "Lon_Kendala": row_ha["Lon_Kendala"],
                                    "Lat_Site": f"{def_lat:.6f}",
                                    "Lon_Site": f"{def_lon:.6f}",
                                    "Kategori_Hard_Access": ha_kat,
                                    "Tingkat_Kesulitan": ha_lvl,
                                    "Status_Akses": ha_stat,
                                    "Rute_Wajib_Moda": ha_kend,
                                    "Estimasi_Biaya_Sewa": "0",
                                    "Estimasi_Waktu_Tempuh": row_ha["Estimasi_Tempuh"],
                                    "Rekomendasi_Rute_Detail": "-",
                                    "PIC_Input": row_ha["PIC_Input"],
                                    "Keterangan_Detail": row_ha["Keterangan_Detail"],
                                    "Foto_URLs": row_ha["Foto_URLs"]
                                }
                                if gsheet_append_row("Site Hard Akses", COLS_SITE_HARD_AKSES, row_sha_pbu):
                                    st.success(f"✅ Report Hard Access **{ha_sid}** ({len(urls_ha)} foto) berhasil disimpan ke sheet **`Site Hard Akses`** & otomatis tampil di Peta!")
                                    st.rerun()
                                else:
                                    st.error("⚠️ Gagal menyimpan ke Google Sheets tab `Site Hard Akses`.")

            # Ringkasan & Galeri Daftar Lokasi Hard Access Pangkalan Bun
            st.markdown("---")
            if not ha_list_pbu:
                st.info("Belum ada data lokasi Hard Access Pangkalan Bun yang tercatat. Admin dapat menambahkan lokasi + foto melalui form di atas.")
            else:
                cnt_blk = sum(1 for h in ha_list_pbu if "blocked" in h["status"].lower() or "tidak tembus" in h["status"].lower())
                cnt_trb = sum(1 for h in ha_list_pbu if "terbatas" in h["status"].lower())
                cnt_nrm = len(ha_list_pbu) - cnt_blk - cnt_trb
                tot_ha_fotos = sum(len(h["photos"]) for h in ha_list_pbu)

                mh1, mh2, mh3, mh4 = st.columns(4)
                mh1.metric("🚧 Total Lokasi Hard Access", f"{len(ha_list_pbu)} Lokasi")
                mh2.metric("🔴 Blocked / Tidak Tembus", f"{cnt_blk} Lokasi")
                mh3.metric("🟡 Terbatas / 🟢 Normal", f"{cnt_trb} Terbatas / {cnt_nrm} Normal")
                mh4.metric("📸 Total Foto Eviden", f"{tot_ha_fotos} Foto")

                fh1, fh2 = st.columns([1.6, 1.4])
                q_ha = fh1.text_input("🔍 Cari Site ID / Lokasi / Kendala:", key="q_ha_pbu")
                flt_ha_st = fh2.selectbox("Filter Status Akses:", ["Semua Status"] + HARD_ACCESS_STATUS, key="flt_ha_st_pbu")

                ha_filtered = [
                    h for h in ha_list_pbu
                    if (flt_ha_st == "Semua Status" or h["status"] == flt_ha_st)
                    and (not q_ha.strip() or q_ha.strip().lower() in f"{h['site_id']} {h['site_name']} {h['lokasi']} {h['kategori']} {h['keterangan']}".lower())
                ]

                for idx_h, ha in enumerate(reversed(ha_filtered)):
                    border_c = "#10B981" if ("normal" in ha["status"].lower()) else ("#F59E0B" if "terbatas" in ha["status"].lower() else "#EF4444")
                    with st.container(border=True):
                        st.markdown(f"""
                            <div style="border-left:6px solid {border_c}; padding-left:14px;">
                                <div style="display:flex; justify-content:space-between; flex-wrap:wrap; gap:8px; align-items:center;">
                                    <span style="font-size:17px; font-weight:800;">🚧 {ha['site_id']} — {ha['site_name']}</span>
                                    <span style="background:{border_c}; color:#fff; padding:3px 10px; border-radius:99px; font-size:12px; font-weight:800;">{ha['status']}</span>
                                </div>
                                <div style="font-size:13px; margin-top:6px; line-height:1.6;">
                                    📍 <b>Lokasi Kendala:</b> {ha['lokasi']} &nbsp;|&nbsp; 🏷️ <b>Kategori:</b> {ha['kategori']}<br>
                                    ⚠️ <b>Kesulitan:</b> {ha['tingkat']} &nbsp;|&nbsp; 🚙 <b>Moda Wajib:</b> {ha['kendaraan']} &nbsp;|&nbsp; ⏱️ <b>Estimasi:</b> {ha['estimasi']}<br>
                                    📝 <b>Detail:</b> {ha['keterangan']}<br>
                                    <span style="opacity:0.75; font-size:12px;">👨‍💼 Input oleh: {ha['pic']} · 📅 {ha['timestamp']} · 🌐 Koordinat: {ha['lat']}, {ha['lon']}</span>
                                </div>
                            </div>
                        """, unsafe_allow_html=True)

                        c_ha_act1, c_ha_act2 = st.columns(2)
                        if ha["lat"] is not None and ha["lon"] is not None:
                            c_ha_act1.link_button("🧭 Buka Titik di Google Maps", f"https://www.google.com/maps?q={ha['lat']},{ha['lon']}", use_container_width=True)
                        wa_ha_txt = (
                            f"*🚧 REPORT HARD ACCESS SITE — NOP PANGKALAN BUN*\n"
                            f"📍 *Site:* {ha['site_id']} ({ha['site_name']})\n"
                            f"🗺️ *Lokasi Kendala:* {ha['lokasi']}\n"
                            f"🏷️ *Kategori:* {ha['kategori']}\n"
                            f"⚠️ *Kesulitan:* {ha['tingkat']} | *Status:* {ha['status']}\n"
                            f"🚙 *Kendaraan Wajib:* {ha['kendaraan']} | ⏱️ {ha['estimasi']}\n"
                            f"📝 *Detail:* {ha['keterangan']}\n"
                            f"🌐 *Maps:* https://www.google.com/maps?q={ha['lat']},{ha['lon']}\n"
                            f"📸 *Jumlah Foto Eviden:* {len(ha['photos'])} Foto (tersimpan di SIDATA)"
                        )
                        c_ha_act2.link_button("📱 Share Report Hard Access ke WA", f"https://wa.me/?text={urllib.parse.quote(wa_ha_txt)}", use_container_width=True)

                        if ha["photos"]:
                            tampilkan_grid_foto(ha["photos"], f"📸 Dokumentasi Kondisi Akses ({len(ha['photos'])} Foto)")
                        else:
                            st.caption("Belum ada foto dokumentasi pada laporan ini.")

                        # Admin: Tambah Foto Susulan & Update Status Langsung per Kartu
                        if is_admin_ha:
                            with st.expander(f"🛠️ Admin: Tambah Foto Susulan / Update Status ({ha['site_id']})", expanded=False):
                                c_u1, c_u2 = st.columns(2)
                                cur_st_i = HARD_ACCESS_STATUS.index(ha["status"]) if ha["status"] in HARD_ACCESS_STATUS else 0
                                new_ha_st = c_u1.selectbox("Update Status Akses:", HARD_ACCESS_STATUS, index=cur_st_i, key=f"upd_ha_st_{ha['record_id']}_{idx_h}")
                                new_ha_ket = c_u2.text_input("Update Keterangan:", value=ha["keterangan"], key=f"upd_ha_ket_{ha['record_id']}_{idx_h}")
                                add_ha_pics = st.file_uploader("➕ Tambah Foto Susulan:", type=["png", "jpg", "jpeg", "webp"], accept_multiple_files=True, key=f"add_ha_pic_{ha['record_id']}_{idx_h}")
                                if st.button("💾 Simpan Update & Foto Susulan", type="primary", key=f"btn_upd_ha_{ha['record_id']}_{idx_h}"):
                                    with st.spinner("Mengunggah foto & memperbarui Cloud..."):
                                        extra_urls = upload_multiple_images(add_ha_pics, "hard_access_pbu") if add_ha_pics else []
                                        merged_urls = ha["photos"] + extra_urls
                                        df_ha_upd = df_ha_pbu.copy()
                                        mask_rid = df_ha_upd["Record_ID"].astype(str) == str(ha["record_id"])
                                        if mask_rid.any():
                                            df_ha_upd.loc[mask_rid, "Status_Akses"] = new_ha_st
                                            df_ha_upd.loc[mask_rid, "Keterangan_Detail"] = new_ha_ket
                                            df_ha_upd.loc[mask_rid, "Foto_URLs"] = json.dumps(merged_urls, ensure_ascii=False)
                                            if gsheet_overwrite_ws("Hard_Access_PBU", df_ha_upd):
                                                st.success("✅ Data & Foto Susulan Hard Access berhasil diperbarui!")
                                                st.rerun()

                if is_admin_ha:
                    with st.expander("🗂️ Live Editor Tabel Database `Hard_Access_PBU` (Admin)", expanded=False):
                        ed_ha = st.data_editor(
                            df_ha_pbu,
                            use_container_width=True,
                            num_rows="dynamic",
                            column_config={
                                "Status_Akses": st.column_config.SelectboxColumn("Status Akses", options=HARD_ACCESS_STATUS),
                                "Tingkat_Kesulitan": st.column_config.SelectboxColumn("Tingkat Kesulitan", options=HARD_ACCESS_LEVEL)
                            },
                            key="ed_ha_pbu_tbl"
                        )
                        if st.button("💾 Simpan Perubahan Tabel Hard Access ke Cloud", type="primary", use_container_width=True, key="btn_save_ed_ha_tbl"):
                            if gsheet_overwrite_ws("Hard_Access_PBU", ed_ha):
                                st.success("✅ Tabel `Hard_Access_PBU` berhasil diperbarui!")
                                st.rerun()

        # -----------------------------------------------------------------
        # SUB-MODUL 3: ANALISA MANAJEMEN & EXPORT (ON-DEMAND EXCEL 5-SHEET)
        # -----------------------------------------------------------------
        elif sub_pbu_nav == "📊 Analisa Manajemen & Export":
            if df_pbu_routes.empty:
                st.warning("Belum ada data Program NOP Pangkalan Bun untuk dianalisa.")
            else:
                st.markdown("#### 📊 Executive Analytics & Validasi Laporan Manajemen — NOP Pangkalan Bun")

                # Auto-Insight Ringkasan Eksekutif Manajemen
                tim_grp_sum = df_pbu_routes.groupby("Tier Tim #1").agg(
                    Task=("SITE ID Target", "count"),
                    Closed=("Status", lambda x: int(x.astype(str).str.lower().str.contains("close|done|selesai").sum())),
                    Plan_km=("HB->Target (km)", "sum"),
                    Aktual_km=("Total Jarak Aktual (km)", "sum"),
                    Biaya=("Biaya (Rp)", "sum")
                ).sort_values("Aktual_km", ascending=False)
                top_tim_nm = tim_grp_sum.index[0] if not tim_grp_sum.empty else "-"
                top_tim_km = float(tim_grp_sum.iloc[0]["Aktual_km"]) if not tim_grp_sum.empty else 0.0
                far_row = df_pbu_routes.loc[df_pbu_routes["HB->Target (km)"].idxmax()] if not df_pbu_routes.empty else None
                far_txt = f"**{far_row['SITE ID Target']}** ({far_row['HB->Target (km)']} km dari HB {far_row['Tier Tim #1']})" if far_row is not None else "-"

                st.info(
                    f"📌 **Executive Insight Pangkalan Bun:** Dari **{tot_pbu} site target**, pencapaian Closed saat ini **{cnt_closed_pbu} site ({pct_pbu:.1f}%)**. "
                    f"Sebanyak **{cnt_divert_pbu} task** dialihkan sementara ke site emergency/backup dan **{cnt_ha_pbu} site** teridentifikasi kendala Hard Access. "
                    f"Mobilitas tertinggi tercatat pada tim **{top_tim_nm}** (**{top_tim_km:,.1f} km** aktual / Est. Jalan {top_tim_km*1.35:,.0f} km). "
                    f"Target terjauh dari Homebase: {far_txt}."
                )

                ca_b1, ca_b2 = st.columns(2)
                with ca_b1:
                    with st.container(border=True):
                        st.markdown("**🚗 Total Jarak Plan vs Aktual per Tim Tier #1 (km)**")
                        ui_bar_chart(tim_grp_sum[["Plan_km", "Aktual_km"]].rename(columns={"Plan_km": "Plan HB➔Target (km)", "Aktual_km": "Aktual Tempuh (km)"}), color=["#F59E0B", "#06B6D4"])
                        st.caption("Oranye = Jarak Plan (HB➔Target) | Cyan = Total Jarak Aktual Tempuh Tim (Tier #1-#3)")
                with ca_b2:
                    with st.container(border=True):
                        st.markdown("**🚦 Akar Penyebab Status Task (Closed vs Divert vs Hard Access)**")
                        def _short_cat(s):
                            if "Tercapai" in s: return "✅ Tercapai (Closed)"
                            if "Hard Access" in s or "Akses" in s: return "🚧 Kendala Hard Access"
                            if "On Target" in s: return "🔄 On Target (Waiting Material/Swap)"
                            if "Dialihkan" in s: return "🔀 Dialihkan ke Site Emergency"
                            if "Kendala" in s: return "⚠️ Kendala Lapangan Lain"
                            if "Melewati" in s: return "🔴 Overdue / Belum Jalan"
                            return "📅 Terjadwal / Waiting"
                        df_cat_an = df_pbu_routes["Analisa Target / Divert"].apply(_short_cat).value_counts().to_frame("Jumlah Task")
                        ui_bar_chart(df_cat_an, color="#EF4444", horizontal=True)

                with st.expander("👷 Lihat Tabel Rekap Beban Kerja, Jarak & Biaya per Tim Tier #1", expanded=True):
                    df_t_show = tim_grp_sum.reset_index()
                    df_t_show["Est. Jalan Aktual (km)"] = (df_t_show["Aktual_km"] * 1.35).round(1)
                    df_t_show["Plan_km"] = df_t_show["Plan_km"].round(1)
                    df_t_show["Aktual_km"] = df_t_show["Aktual_km"].round(1)
                    st.dataframe(df_t_show, use_container_width=True, hide_index=True)

                st.markdown("#### 📋 Tabel Validasi Detail Jarak, Hard Access & Remark per Tier (Siap Audit Manajemen)")
                cols_mgmt_view = [
                    "No", "SITE ID Target", "SITE NAME", "Class", "Status", "Status Hard Access", "Tier Tim #1",
                    "HB->Target (km)", "Est Jalan HB->Target (km)",
                    "Total Jarak Aktual (km)", "Est Jalan Aktual (km)",
                    "Analisa Target / Divert", "Tarikan Rute Aktual",
                    "Tier #1 Tgl", "Tier #1 Site/Info", "Tier #1 Remark",
                    "Tier #2 Tgl", "Tier #2 Site/Info", "Tier #2 Remark",
                    "Tier #3 Tgl", "Tier #3 Site/Info", "Tier #3 Remark",
                    "Biaya (Rp)", "Remark Hasil Pekerjaan"
                ]
                st.dataframe(df_pbu_routes[cols_mgmt_view], use_container_width=True, hide_index=True, height=440)

                wa_mgmt = (
                    f"*📊 LAPORAN DETAIL PROGRAM NOP, TIER TRACKING & HARD ACCESS — PANGKALAN BUN*\n"
                    f"📅 Tanggal: {datetime.date.today().strftime('%d %B %Y')}\n"
                    f"🎯 Target: {tot_pbu} Site | ✅ Closed: {cnt_closed_pbu} ({pct_pbu:.1f}%) | 🔄 In Progress: {cnt_inprog_pbu} | ⏳ Open: {cnt_open_pbu}\n"
                    f"🔀 Dialihkan Emergency/Backup: {cnt_divert_pbu} Task | 🚧 Hard Access: {cnt_ha_pbu} Site ({len(ha_list_pbu)} Laporan)\n"
                    f"🚗 Total Jarak Aktual Tim: *{tot_km_act_pbu:,.1f} km* (Est. Jalan: {tot_km_act_pbu*1.35:,.0f} km)\n"
                    f"💰 Total Biaya Operasional: *Rp {tot_biaya_pbu:,.0f}*\n\n"
                    f"*📋 RINCIAN PERGERAKAN & ALASAN PER SITE:*\n"
                )
                for _, rm in df_pbu_routes.iterrows():
                    wa_mgmt += f"{rm['No']}. *{rm['SITE ID Target']} ({rm['Status']})* — Tim: {rm['Tier Tim #1']}\n"
                    wa_mgmt += f"   • Jarak Plan: {rm['HB->Target (km)']}km | Aktual: {rm['Total Jarak Aktual (km)']}km\n"
                    wa_mgmt += f"   • Analisa: {rm['Analisa Target / Divert']}\n"
                    if "🚧" in str(rm["Status Hard Access"]):
                        wa_mgmt += f"   • Akses: {rm['Status Hard Access']}\n"
                    if rm['Tier #1 Site/Info'] != "-":
                        wa_mgmt += f"   • T#1 ({rm['Tier #1 Tgl']}): {rm['Tier #1 Site/Info']} — {rm['Tier #1 Remark']}\n"
                    if rm['Tier #2 Site/Info'] != "-":
                        wa_mgmt += f"   • T#2 ({rm['Tier #2 Tgl']}): {rm['Tier #2 Site/Info']} — {rm['Tier #2 Remark']}\n"
                    if rm['Tier #3 Site/Info'] != "-":
                        wa_mgmt += f"   • T#3 ({rm['Tier #3 Tgl']}): {rm['Tier #3 Site/Info']} — {rm['Tier #3 Remark']}\n"
                    if rm['Remark Hasil Pekerjaan'] != "-":
                        wa_mgmt += f"   • Hasil: {rm['Remark Hasil Pekerjaan']}\n"

                cm_ex1, cm_ex2 = st.columns(2)
                with cm_ex1:
                    if st.button("🛠️ Siapkan Laporan Excel Manajemen (5 Sheet)", type="primary", use_container_width=True, key="btn_prep_pbu_xlsx"):
                        with st.spinner("Merakit Excel 5-Sheet Manajemen..."):
                            st.session_state["pbu_mgmt_xlsx_bytes"] = build_pbu_management_report_xlsx(df_pbu_routes, df_site_m_pbu, df_hb_m_pbu, df_hard_access=df_ha_pbu)
                    if st.session_state.get("pbu_mgmt_xlsx_bytes"):
                        st.download_button(
                            "📥 DOWNLOAD LAPORAN MANAJEMEN PBU (.xlsx)",
                            data=st.session_state["pbu_mgmt_xlsx_bytes"],
                            file_name=f"Laporan_Manajemen_Program_PBU_{datetime.date.today()}.xlsx",
                            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                            type="primary",
                            use_container_width=True,
                            key="dl_mgmt_pbu_xlsx"
                        )
                with cm_ex2:
                    st.link_button(
                        "📱 KIRIM LAPORAN DETAIL KE WHATSAPP MANAJEMEN",
                        f"https://wa.me/?text={urllib.parse.quote(wa_mgmt)}",
                        use_container_width=True
                    )

        # -----------------------------------------------------------------
        # SUB-MODUL 4: LIVE EDITOR TIMEPLAN, TIER TRACKER & BIAYA
        # -----------------------------------------------------------------
        elif sub_pbu_nav == "✏️ Live Editor Timeplan & Tier":
            if df_pbu_cloud.empty:
                st.warning("Belum ada data Timeplan Pangkalan Bun. Upload melalui panel di atas.")
            else:
                st.markdown("#### ✏️ Live Editor Timeplan, Tier Tim (#1-#3), Tracker Kunjungan & Biaya")
                st.info("💡 Koordinator/Admin dapat langsung memperbarui **Status Activity**, **Biaya**, **Tracker #1/#2/#3 Site ID**, dan **Remark** di tabel bawah ini. Setelah klik **Simpan ke Cloud**, jarak km & garis peta langsung terhitung ulang otomatis.")

                cfg_pbu = {}
                if stcol_pbu:
                    cfg_pbu[stcol_pbu] = st.column_config.SelectboxColumn("Status Activity", options=["Open", "In Progress", "Closed", "Pending"])
                ed_pbu = st.data_editor(df_pbu_cloud, use_container_width=True, num_rows="dynamic", column_config=cfg_pbu, key="ed_prog_pbu")

                if st.button("💾 SIMPAN CHECKLIST, TIER TRACKER & BIAYA PBU KE CLOUD (`Program_NOP_PBU`)", type="primary", use_container_width=True, key="btn_upd_pbu_prog"):
                    if gsheet_overwrite_ws("Program_NOP_PBU", ed_pbu):
                        st.session_state.pop("pbu_mgmt_xlsx_bytes", None)
                        st.success("✅ Data Timeplan & Tier Tracking Pangkalan Bun berhasil diperbarui ke Cloud!")
                        st.rerun()
                    else:
                        st.error("⚠️ Gagal menyimpan ke Google Sheets.")

        # -----------------------------------------------------------------
        # SUB-MODUL 5: HISTORY ONSITE & POWER SIDATA PER SITE
        # -----------------------------------------------------------------
        elif sub_pbu_nav == "⚡ History Onsite & Power SIDATA":
            if df_pbu_cloud.empty:
                st.warning("Belum ada daftar Site Program Pangkalan Bun.")
            else:
                with st.spinner("Menyinkronkan history onsite SIDATA..."):
                    onsite_idx_full_pbu = build_onsite_history_index(get_laporan_db())
                _render_site_history_inspector(df_pbu_cloud, scol_pbu, "pbu", onsite_idx_full_pbu)

    # =====================================================================
    # AREA 2: NOP PALANGKARAYA (TASK FORCE & AVAILABILITY)
    # =====================================================================
    else:
        COLS_PROG_PLK = ["Site ID", "Site Name", "Kabupaten", "Site Class", "RTS Name", "Leader Name", "RC Category", "RC Tier", "Remark", "RC Detail", "Action Cat", "Action Detail", "Support Needed", "Target Week (plan)", "Plan Visit Date", "Status", "Potensi kenaikan NOP Okt (poin)"]
        df_plk_cloud, ok_plk_c = gsheet_read_df("Program_NOP_PLK", COLS_PROG_PLK)
        local_plk_file = _find_first_existing(
            ["Plan_Task_Force_EOY26_Palangkaraya_rapi (1).xlsx", "Plan_Task_Force_EOY26_Palangkaraya_rapi.xlsx"],
            "*Task_Force*Palangkaraya*.xlsx"
        )
        if df_plk_cloud.empty and local_plk_file:
            try:
                df_plk_cloud = parse_excel_smart_header(local_plk_file, "Take_In")
            except Exception:
                pass

        with st.expander("📤 Upload & Simpan Excel Program NOP Palangkaraya ke Cloud (`Program_NOP_PLK`)", expanded=df_plk_cloud.empty):
            up_plk = st.file_uploader("Unggah File Excel Program Palangkaraya (.xlsx / .csv):", type=["xlsx", "xls", "csv"], key="up_prog_plk")
            mode_plk = st.radio("Mode Simpan:", ["Replace (Ganti List Program)", "Append (Tambah ke List Lama)"], horizontal=True, key="mode_prog_plk")
            if up_plk is not None:
                try:
                    if up_plk.name.lower().endswith(".csv"):
                        df_new_plk = pd.read_csv(up_plk)
                    else:
                        sheets_plk = get_excel_sheet_names_safe(up_plk)
                        def_sh = sheets_plk.index("Take_In") if "Take_In" in sheets_plk else 0
                        sh_plk = st.selectbox("Pilih Sheet sumber:", sheets_plk, index=def_sh, key="sh_prog_plk")
                        df_new_plk = parse_excel_smart_header(up_plk, sh_plk)
                    st.write(f"Preview ({len(df_new_plk)} site):")
                    st.dataframe(df_new_plk.head(10), use_container_width=True, hide_index=True)
                    if st.button("💾 SIMPAN PROGRAM NOP PALANGKARAYA KE CLOUD", type="primary", use_container_width=True, key="btn_save_prog_plk"):
                        df_save = pd.concat([df_plk_cloud, df_new_plk], ignore_index=True) if ("Append" in mode_plk and not df_plk_cloud.empty) else df_new_plk
                        if gsheet_overwrite_ws("Program_NOP_PLK", df_save):
                            st.success(f"✅ {len(df_save)} site Program NOP Palangkaraya tersimpan ke tab `Program_NOP_PLK`!")
                            st.rerun()
                        else:
                            st.error("⚠️ Gagal menyimpan ke Google Sheets.")
                except Exception as e:
                    st.error(f"Gagal membaca Excel Palangkaraya: {e}")

        if df_plk_cloud.empty:
            st.warning("Belum ada data Program NOP Palangkaraya. Silakan upload Excel di atas.")
        else:
            onsite_idx_plk = build_onsite_history_index(get_laporan_db())
            scol_plk = next((c for c in df_plk_cloud.columns if "site id" in c.lower() or c.lower() == "site_id"), df_plk_cloud.columns[0])
            stcol_plk = next((c for c in df_plk_cloud.columns if c.lower() == "status" or "status" in c.lower()), None)
            rccol_plk = next((c for c in df_plk_cloud.columns if "rc category" in c.lower() or "rc tier" in c.lower() or "rca" in c.lower()), None)
            avcol_plk = next((c for c in df_plk_cloud.columns if "potensi" in c.lower() or "avail" in c.lower() or "poin" in c.lower()), None)

            df_plk_enr = _enrich_with_sidata(df_plk_cloud, scol_plk, onsite_idx_plk)

            tot_plk = len(df_plk_enr)
            closed_mask = df_plk_enr[stcol_plk].astype(str).str.lower().str.contains("close|done|selesai") if stcol_plk else pd.Series([False]*tot_plk)
            cnt_closed_plk = int(closed_mask.sum())
            cnt_open_plk = tot_plk - cnt_closed_plk
            pct_plk = (cnt_closed_plk / tot_plk * 100) if tot_plk else 0.0
            cnt_onsite_plk = int((df_plk_enr["Onsite_SIDATA (x)"] > 0).sum())

            pot_avail_tot, pot_avail_done = 0.0, 0.0
            if avcol_plk:
                av_nums = pd.to_numeric(df_plk_enr[avcol_plk], errors="coerce").fillna(0.0)
                pot_avail_tot = float(av_nums.sum())
                pot_avail_done = float(av_nums[closed_mask].sum())

            base_avail_plk = st.number_input("Baseline Availability NOP Palangkaraya Saat Ini (%):", min_value=0.0, max_value=100.0, value=95.63, step=0.01, format="%.2f", key="base_av_plk")

            kpl1, kpl2, kpl3, kpl4, kpl5 = st.columns(5)
            kpl1.metric("Total Site Program", f"{tot_plk} Site")
            kpl2.metric("✅ Closed / Selesai", f"{cnt_closed_plk} ({pct_plk:.1f}%)")
            kpl3.metric("⏳ Open / Berjalan", f"{cnt_open_plk} Site")
            kpl4.metric("📡 Sudah Onsite SIDATA", f"{cnt_onsite_plk} Site")
            kpl5.metric("📈 Proyeksi Availability", f"{base_avail_plk + pot_avail_done:.2f}%", delta=f"+{pot_avail_done:.2f} / {pot_avail_tot:.2f} poin")

            cp_g1, cp_g2 = st.columns(2)
            with cp_g1:
                with st.container(border=True):
                    st.markdown("**🛑 Breakdown Problem Program Palangkaraya**")
                    if rccol_plk:
                        df_rc_p = df_plk_enr[rccol_plk].replace("", "Lainnya").value_counts().head(10).to_frame("Jumlah Site")
                        ui_bar_chart(df_rc_p, color="#F59E0B", horizontal=True)
                    else:
                        st.caption("Kolom RC Category tidak ditemukan.")
            with cp_g2:
                with st.container(border=True):
                    st.markdown("**📊 Status Pencapaian Target & Kunjungan Onsite**")
                    df_cmp_p = pd.DataFrame({
                        "Jumlah Site": [cnt_closed_plk, cnt_open_plk, cnt_onsite_plk, tot_plk - cnt_onsite_plk]
                    }, index=["Program Closed", "Program Open", "Sudah Record SIDATA", "Belum Record SIDATA"])
                    ui_bar_chart(df_cmp_p, color="#2563EB")

            st.markdown("#### ✅ Checklist Target Selesai, Status & Sinkronisasi History SIDATA (Palangkaraya)")
            flt_p1, flt_p2 = st.columns([1.5, 1.5])
            q_plk_p = flt_p1.text_input("🔍 Cari Site ID / Nama / RTS / Problem:", key="q_prog_plk")
            st_plk_f = flt_p2.selectbox("Filter Status:", ["Semua", "Open", "Closed", "Sudah Onsite SIDATA", "Belum Onsite SIDATA"], key="st_flt_plk")

            df_plk_show = df_plk_enr.copy()
            if st_plk_f == "Open" and stcol_plk:
                df_plk_show = df_plk_show[~df_plk_show[stcol_plk].astype(str).str.lower().str.contains("close|done|selesai")]
            elif st_plk_f == "Closed" and stcol_plk:
                df_plk_show = df_plk_show[df_plk_show[stcol_plk].astype(str).str.lower().str.contains("close|done|selesai")]
            elif st_plk_f == "Sudah Onsite SIDATA":
                df_plk_show = df_plk_show[df_plk_show["Onsite_SIDATA (x)"] > 0]
            elif st_plk_f == "Belum Onsite SIDATA":
                df_plk_show = df_plk_show[df_plk_show["Onsite_SIDATA (x)"] == 0]
            if q_plk_p.strip():
                qq = q_plk_p.strip().lower()
                df_plk_show = df_plk_show[df_plk_show.apply(lambda r: qq in " ".join(r.astype(str)).lower(), axis=1)]

            cfg_plk = {}
            if stcol_plk:
                cfg_plk[stcol_plk] = st.column_config.SelectboxColumn("Status", options=["Open", "On Progress", "Closed", "Pending"])
            ed_plk = st.data_editor(df_plk_show, use_container_width=True, num_rows="dynamic", column_config=cfg_plk, key="ed_prog_plk")

            cb_p1, cb_p2 = st.columns(2)
            with cb_p1:
                if st.button("💾 Simpan Checklist & Perubahan Program PLK ke Cloud", type="primary", use_container_width=True, key="btn_upd_plk_prog"):
                    if q_plk_p.strip() or st_plk_f != "Semua":
                        st.warning("⚠️ Set filter ke 'Semua' & kosongkan kotak cari terlebih dahulu sebelum menyimpan agar seluruh site tersimpan utuh.")
                    else:
                        cols_save_plk = [c for c in ed_plk.columns if c not in ["Onsite_SIDATA (x)", "Tgl_Onsite_Terakhir", "Status_Onsite_SIDATA", "Problem_Onsite_Terakhir", "Last_VAC", "Last_VDC", "Last_Load_DC"]]
                        if gsheet_overwrite_ws("Program_NOP_PLK", ed_plk[cols_save_plk]):
                            st.success("✅ Checklist Program NOP Palangkaraya berhasil disimpan ke Cloud!")
                            st.rerun()
            with cb_p2:
                if st.button("🛠️ Siapkan Download Excel PLK (.xlsx)", use_container_width=True, key="prep_dl_prog_plk"):
                    buf_plk_x = BytesIO()
                    with pd.ExcelWriter(buf_plk_x, engine="openpyxl") as xw:
                        df_plk_show.to_excel(xw, sheet_name="Program_NOP_PLK", index=False)
                    st.session_state["plk_prog_xlsx_bytes"] = buf_plk_x.getvalue()
                if st.session_state.get("plk_prog_xlsx_bytes"):
                    st.download_button("📥 Download Program PLK + History SIDATA (.xlsx)", data=st.session_state["plk_prog_xlsx_bytes"], file_name=f"Program_NOP_PLK_{datetime.date.today()}.xlsx", mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet", type="primary", use_container_width=True, key="dl_prog_plk")

            st.markdown("---")
            _render_site_history_inspector(df_plk_cloud, scol_plk, "plk", onsite_idx_plk)
# =========================================================================
# MENU 13: SiDAVID (SYSTEM DATA VERIFIKASI INTERNAL)
# =========================================================================
elif menu == "🛡️ SiDAVID (Verifikasi Internal)":
    st.markdown("<h1 style='text-align: center;'>🛡️ SiDAVID — System Data Verifikasi Internal</h1>", unsafe_allow_html=True)
    st.markdown("<p style='text-align: center; opacity: 0.85;'>Verifikasi Data Mateline (Sheet `Mateline`), Sertifikasi Tim (WAH, BA, BE), Mutasi Personel (Join/Keluar) & Validasi Data Internal Lainnya</p>", unsafe_allow_html=True)

    fme_plk_list = load_master_tim_names("PLK")
    fme_pbu_list = load_master_tim_names("PBU")
    semua_tim = sorted(list(set(fme_plk_list + fme_pbu_list)))

    COLS_MATELINE = ["NOP", "Nama_Teknisi", "ID_Mateline", "Status_Verifikasi", "Tanggal_Upload", "Catatan_Verifikasi"]
    COLS_SERTIF = ["Timestamp", "NOP", "Nama_Tim", "Jenis_Sertifikat", "Keterangan_Punya", "Tanggal_Training", "Tanggal_Release", "Tanggal_Expired", "Status_Masa_Berlaku", "URL_Bukti", "Catatan"]
    COLS_PERSONEL = ["Timestamp", "NOP", "Nama_Tim", "NIK_ID", "Jabatan_Role", "Status_Kerja", "Tanggal_Join", "Tanggal_Keluar", "Alasan_Keterangan"]
    COLS_LAINNYA = ["Timestamp", "NOP", "Kategori_Validasi", "Objek_Atau_Nama", "Keterangan_Punya", "Tanggal_Terbit", "Tanggal_Expired", "Status_Validasi", "URL_Bukti", "Catatan"]

    df_mat, ok_mat = gsheet_read_df("Mateline", COLS_MATELINE)
    df_srt, ok_srt = gsheet_read_df("SiDAVID_Sertifikat", COLS_SERTIF)
    df_prs, ok_prs = gsheet_read_df("SiDAVID_Personel", COLS_PERSONEL)
    df_lnn, ok_lnn = gsheet_read_df("SiDAVID_Validasi_Lain", COLS_LAINNYA)

    if not ok_mat:
        st.warning("⚠️ Koneksi Google Sheets belum aktif (`gcp_json`). Sinkronisasi cloud membutuhkan koneksi Spreadsheet aktif.")

    tot_mat = len(df_mat)
    tot_srt_punya = int((df_srt["Keterangan_Punya"].astype(str).str.strip().str.lower() == "punya").sum()) if not df_srt.empty else 0
    tot_prs_aktif = int((df_prs["Status_Kerja"].astype(str).str.strip().str.lower() == "aktif").sum()) if not df_prs.empty else 0
    tot_lnn = len(df_lnn)

    ux_hero(
        "INTERNAL COMPLIANCE & VERIFICATION",
        "SiDAVID Command Center",
        "Monitoring terpusat Verifikasi Mateline, Kepatuhan Sertifikat WAH/BA/BE, Status Tim Join/Keluar & Validasi Standby.",
        stats=[
            (f"{tot_mat} Baris", "Data Sheet Mateline"),
            (f"{tot_srt_punya} Sertifikat", "Sertifikat WAH/BA/BE"),
            (f"{tot_prs_aktif} Personel", "Tim Status Aktif"),
            (f"{tot_lnn} Item", "Validasi Lainnya")
        ]
    )

    tab_mat, tab_srt, tab_prs, tab_lnn = st.tabs([
        "📡 1. Verifikasi Mateline",
        "🎓 2. Sertifikasi (WAH, BA, BE)",
        "👥 3. Data Tim (Join & Keluar)",
        "🧩 4. Validasi Data Lainnya (Standby)"
    ])

    # =====================================================================
    # TAB 1: VERIFIKASI MATELINE (TARIK EXCEL ➔ SAVE KE SHEET 'Mateline')
    # =====================================================================
    with tab_mat:
        st.markdown("### 📡 Verifikasi Data Mateline (Import Excel ➔ Save Sheet `Mateline`)")
        st.info("💡 Unggah file Excel tarikan dari **Mateline** di bawah ini. Data otomatis dibaca, ditambahkan kolom verifikasi, tersimpan ke Google Sheets tab **`Mateline`**, dan dapat diunduh kembali sebagai Excel dengan nama sheet **`Mateline`**.")

        with st.container(border=True):
            c_m1, c_m2 = st.columns([2, 1])
            with c_m1:
                up_mateline = st.file_uploader("📂 Tarik / Upload File Excel Mateline (.xlsx / .xls / .csv):", type=["xlsx", "xls", "csv"], key="sd_up_mateline")
            with c_m2:
                mode_save_mat = st.radio("Mode Simpan ke Sheet `Mateline`:", ["Timpa / Replace Semua Data", "Tambahkan (Append) ke Data Lama"], key="sd_mode_mat")

            if up_mateline is not None:
                try:
                    if up_mateline.name.lower().endswith(".csv"):
                        df_up_m = pd.read_csv(up_mateline)
                    else:
                        sheets_m = get_excel_sheet_names_safe(up_mateline)
                        pilih_sh_m = st.selectbox("Pilih Sheet sumber dari file Excel Mateline:", sheets_m, key="sd_sh_mat")
                        df_up_m = read_excel_safe_raw(up_mateline, sheet_name=pilih_sh_m)

                    df_up_m = df_up_m.dropna(how="all").reset_index(drop=True)
                    if "Status_Verifikasi" not in df_up_m.columns:
                        df_up_m["Status_Verifikasi"] = "Belum Diverifikasi"
                    if "Catatan_Verifikasi" not in df_up_m.columns:
                        df_up_m["Catatan_Verifikasi"] = ""
                    df_up_m["Tanggal_Upload"] = datetime.datetime.now().strftime("%Y-%m-%d %H:%M")

                    st.markdown(f"**Preview Data Tarikan Mateline ({len(df_up_m)} baris):**")
                    st.dataframe(df_up_m.head(20), use_container_width=True, hide_index=True)

                    if st.button("💾 SIMPAN DATA TARIKAN KE SHEET 'Mateline'", type="primary", use_container_width=True, key="sd_btn_save_mat"):
                        with st.spinner("Menyimpan data ke Google Sheets tab 'Mateline'..."):
                            if "Append" in mode_save_mat and not df_mat.empty:
                                df_gabung = pd.concat([df_mat, df_up_m], ignore_index=True).fillna("")
                            else:
                                df_gabung = df_up_m.fillna("")
                            if gsheet_overwrite_ws("Mateline", df_gabung):
                                st.success(f"✅ Berhasil! {len(df_gabung)} baris data tersimpan permanen di sheet **`Mateline`**.")
                                st.rerun()
                            else:
                                st.error("⚠️ Gagal menyimpan ke Google Sheets. Pastikan kredensial `gcp_json` aktif.")
                except Exception as e:
                    st.error(f"Gagal membaca file Excel Mateline: {e}")

        st.markdown("---")
        st.markdown("#### ✏️ Live Verifikasi & Editor Sheet `Mateline`")
        if df_mat.empty:
            st.info("Belum ada data di sheet `Mateline`. Silakan unggah file Excel Mateline di atas.")
        else:
            c_fm1, c_fm2 = st.columns([2, 1])
            cari_mat = c_fm1.text_input("🔍 Cari data di Sheet Mateline (Nama / ID / Site):", key="sd_cari_mat")
            flt_ver_mat = c_fm2.selectbox("Filter Status Verifikasi:", ["Semua"] + sorted(df_mat["Status_Verifikasi"].astype(str).unique().tolist()) if "Status_Verifikasi" in df_mat.columns else ["Semua"], key="sd_flt_mat")

            df_mat_show = df_mat.copy()
            if flt_ver_mat != "Semua" and "Status_Verifikasi" in df_mat_show.columns:
                df_mat_show = df_mat_show[df_mat_show["Status_Verifikasi"].astype(str) == flt_ver_mat]
            if cari_mat.strip():
                q_m = cari_mat.strip().lower()
                df_mat_show = df_mat_show[df_mat_show.apply(lambda r: q_m in " ".join(r.astype(str)).lower(), axis=1)]

            col_cfg_mat = {}
            if "Status_Verifikasi" in df_mat_show.columns:
                col_cfg_mat["Status_Verifikasi"] = st.column_config.SelectboxColumn(
                    "Status Verifikasi",
                    options=["Belum Diverifikasi", "Valid / Sesuai", "Tidak Sesuai / Perlu Revisi", "Pending"]
                )

            edited_mat = st.data_editor(
                df_mat_show,
                use_container_width=True,
                num_rows="dynamic",
                column_config=col_cfg_mat,
                key="sd_ed_mateline"
            )

            c_bm1, c_bm2 = st.columns(2)
            with c_bm1:
                if st.button("💾 Update Hasil Verifikasi ke Sheet `Mateline`", type="primary", use_container_width=True, key="sd_upd_mat"):
                    if cari_mat.strip() or flt_ver_mat != "Semua":
                        st.warning("⚠️ Kosongkan pencarian & set filter ke 'Semua' terlebih dahulu sebelum menyimpan agar seluruh baris sheet `Mateline` tersimpan utuh.")
                    else:
                        with st.spinner("Menyimpan perubahan ke sheet 'Mateline'..."):
                            if gsheet_overwrite_ws("Mateline", edited_mat):
                                st.success("✅ Sheet `Mateline` berhasil diperbarui!")
                                st.rerun()
            with c_bm2:
                buf_mat = BytesIO()
                with pd.ExcelWriter(buf_mat, engine="openpyxl") as xw:
                    edited_mat.to_excel(xw, sheet_name="Mateline", index=False)
                st.download_button(
                    "📥 Download Excel (Sheet: Mateline)",
                    data=buf_mat.getvalue(),
                    file_name=f"Verifikasi_Mateline_{datetime.date.today()}.xlsx",
                    mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                    use_container_width=True,
                    key="sd_dl_mat"
                )

    # =====================================================================
    # TAB 2: SERTIFIKASI TIM (WAH, BA, BE)
    # =====================================================================
    with tab_srt:
        st.markdown("### 🎓 Monitoring & Verifikasi Sertifikat Tim (WAH, BA, BE)")

        with st.expander("➕ Input / Update Data Sertifikat Personel (WAH / BA / BE)", expanded=df_srt.empty):
            cs1, cs2, cs3 = st.columns(3)
            with cs1:
                srt_nop = st.selectbox("Area NOP:", ["Palangkaraya", "Pangkalan Bun"], key="sd_srt_nop")
                pilih_tim_srt = st.selectbox("Pilih Nama Tim (Master):", ["-- Ketik Manual --"] + semua_tim, key="sd_srt_tim_sel")
                if pilih_tim_srt == "-- Ketik Manual --":
                    srt_nama = st.text_input("Ketik Nama Tim *:", key="sd_srt_tim_txt").strip()
                else:
                    srt_nama = pilih_tim_srt
            with cs2:
                srt_jenis = st.selectbox("Jenis Sertifikat *:", ["WAH", "BA", "BE"], key="sd_srt_jenis")
                srt_punya = st.radio("Keterangan Kepemilikan *:", ["Punya", "Tidak Punya"], horizontal=True, key="sd_srt_punya")
            with cs3:
                if srt_punya == "Punya":
                    srt_tgl_tr = st.date_input("📅 Tanggal Training:", value=datetime.date.today(), key="sd_srt_tr").strftime("%Y-%m-%d")
                    srt_tgl_rel = st.date_input("📅 Tanggal Release Sertifikat:", value=datetime.date.today(), key="sd_srt_rel").strftime("%Y-%m-%d")
                    srt_tgl_exp = st.date_input("📅 Tanggal Expired Sertifikat:", value=datetime.date.today() + datetime.timedelta(days=365), key="sd_srt_exp").strftime("%Y-%m-%d")
                else:
                    srt_tgl_tr, srt_tgl_rel, srt_tgl_exp = "-", "-", "-"
                    st.info("Status: **Tidak Punya**. Kolom tanggal otomatis diisi `-`.")

            cs4, cs5 = st.columns([1.5, 1.5])
            with cs4:
                srt_cat = st.text_input("📝 Catatan / Nomor Sertifikat / Lembaga:", placeholder="Cth: No. Sertif / Rencana Batch Training...", key="sd_srt_cat")
            with cs5:
                srt_file = st.file_uploader("📎 Upload Scan / Foto Sertifikat (Opsional):", type=["png", "jpg", "jpeg", "pdf"], key="sd_srt_file")

            if st.button("💾 SIMPAN DATA SERTIFIKAT KE SiDAVID", type="primary", use_container_width=True, key="sd_btn_save_srt"):
                if not srt_nama:
                    st.error("⚠️ Nama Tim wajib diisi!")
                else:
                    with st.spinner("Menyimpan data sertifikat..."):
                        url_bukti_srt = ""
                        if srt_file is not None:
                            url_bukti_srt = upload_datalog(srt_file, "sidavid_sertifikat") if srt_file.name.lower().endswith(".pdf") else upload_image(srt_file, "sidavid_sertifikat")
                        st_berlaku = sidavid_eval_expired(srt_punya, srt_tgl_exp)
                        rec_srt = {
                            "Timestamp": datetime.datetime.now().strftime("%Y-%m-%d %H:%M:%S"),
                            "NOP": srt_nop,
                            "Nama_Tim": srt_nama,
                            "Jenis_Sertifikat": srt_jenis,
                            "Keterangan_Punya": srt_punya,
                            "Tanggal_Training": srt_tgl_tr,
                            "Tanggal_Release": srt_tgl_rel,
                            "Tanggal_Expired": srt_tgl_exp,
                            "Status_Masa_Berlaku": st_berlaku,
                            "URL_Bukti": url_bukti_srt or "",
                            "Catatan": srt_cat
                        }
                        if gsheet_append_row("SiDAVID_Sertifikat", COLS_SERTIF, rec_srt):
                            st.success(f"✅ Data sertifikat **{srt_jenis}** untuk **{srt_nama}** berhasil disimpan!")
                            st.rerun()
                        else:
                            st.error("⚠️ Gagal menyambung ke Google Sheets.")

        if df_srt.empty:
            st.info("Belum ada data sertifikat WAH / BA / BE yang tercatat.")
        else:
            df_srt_v = df_srt.copy()
            df_srt_v["Status_Masa_Berlaku"] = df_srt_v.apply(
                lambda r: sidavid_eval_expired(r.get("Keterangan_Punya"), r.get("Tanggal_Expired")), axis=1
            )

            m_s1, m_s2, m_s3, m_s4 = st.columns(4)
            cnt_wah = int(((df_srt_v["Jenis_Sertifikat"] == "WAH") & (df_srt_v["Keterangan_Punya"] == "Punya")).sum())
            cnt_ba = int(((df_srt_v["Jenis_Sertifikat"] == "BA") & (df_srt_v["Keterangan_Punya"] == "Punya")).sum())
            cnt_be = int(((df_srt_v["Jenis_Sertifikat"] == "BE") & (df_srt_v["Keterangan_Punya"] == "Punya")).sum())
            cnt_exp = int(df_srt_v["Status_Masa_Berlaku"].str.contains("Expired", na=False).sum())

            m_s1.metric("🧗 Punya WAH", f"{cnt_wah} Orang")
            m_s2.metric("🛡️ Punya BA", f"{cnt_ba} Orang")
            m_s3.metric("⚡ Punya BE", f"{cnt_be} Orang")
            m_s4.metric("🚨 Expired / Segera Exp", f"{cnt_exp} Sertifikat")

            # Matriks Ringkas 1 Baris per Personel (WAH | BA | BE)
            with st.expander("📊 Lihat Matriks Ringkasan 1 Baris per Personel (WAH | BA | BE)", expanded=True):
                mat_rows = []
                for (np_val, nm_val), grp in df_srt_v.groupby(["NOP", "Nama_Tim"]):
                    row_m = {"NOP": np_val, "Nama_Tim": nm_val, "WAH": "❌ Belum Terdata", "BA": "❌ Belum Terdata", "BE": "❌ Belum Terdata"}
                    for _, rg in grp.iterrows():
                        jns = str(rg.get("Jenis_Sertifikat", "")).upper()
                        if jns in ("WAH", "BA", "BE"):
                            if str(rg.get("Keterangan_Punya", "")).lower() == "punya":
                                row_m[jns] = f"{rg['Status_Masa_Berlaku']} (Exp: {rg.get('Tanggal_Expired', '-')})"
                            else:
                                row_m[jns] = "❌ Tidak Punya"
                    mat_rows.append(row_m)
                if mat_rows:
                    st.dataframe(pd.DataFrame(mat_rows), use_container_width=True, hide_index=True)

            cf_s1, cf_s2, cf_s3 = st.columns(3)
            flt_srt_nop = cf_s1.selectbox("Filter NOP:", ["Semua", "Palangkaraya", "Pangkalan Bun"], key="sd_fs_nop")
            flt_srt_jns = cf_s2.selectbox("Filter Jenis Sertifikat:", ["Semua", "WAH", "BA", "BE"], key="sd_fs_jns")
            flt_srt_pny = cf_s3.selectbox("Filter Kepemilikan:", ["Semua", "Punya", "Tidak Punya"], key="sd_fs_pny")

            df_srt_f = df_srt_v.copy()
            if flt_srt_nop != "Semua": df_srt_f = df_srt_f[df_srt_f["NOP"] == flt_srt_nop]
            if flt_srt_jns != "Semua": df_srt_f = df_srt_f[df_srt_f["Jenis_Sertifikat"] == flt_srt_jns]
            if flt_srt_pny != "Semua": df_srt_f = df_srt_f[df_srt_f["Keterangan_Punya"] == flt_srt_pny]

            st.markdown("#### 📋 Detail & Live Editor Data Sertifikat")
            edited_srt = st.data_editor(
                df_srt_f,
                use_container_width=True,
                num_rows="dynamic",
                column_config={
                    "Jenis_Sertifikat": st.column_config.SelectboxColumn("Jenis", options=["WAH", "BA", "BE"]),
                    "Keterangan_Punya": st.column_config.SelectboxColumn("Keterangan", options=["Punya", "Tidak Punya"]),
                    "URL_Bukti": st.column_config.LinkColumn("Bukti Sertifikat")
                },
                key="sd_ed_srt"
            )

            c_bs1, c_bs2 = st.columns(2)
            with c_bs1:
                if st.button("💾 Simpan Revisi Tabel Sertifikat", type="primary", use_container_width=True, key="sd_upd_srt"):
                    if flt_srt_nop != "Semua" or flt_srt_jns != "Semua" or flt_srt_pny != "Semua":
                        st.warning("⚠️ Set semua filter ke 'Semua' terlebih dahulu sebelum menyimpan revisi tabel.")
                    else:
                        if gsheet_overwrite_ws("SiDAVID_Sertifikat", edited_srt):
                            st.success("✅ Data Sertifikat berhasil diperbarui!")
                            st.rerun()
            with c_bs2:
                buf_s = BytesIO()
                with pd.ExcelWriter(buf_s, engine="openpyxl") as xw:
                    df_srt_f.to_excel(xw, sheet_name="Sertifikat_WAH_BA_BE", index=False)
                st.download_button(
                    "📥 Download Rekap Sertifikat (.xlsx)",
                    data=buf_s.getvalue(),
                    file_name=f"SiDAVID_Sertifikat_{datetime.date.today()}.xlsx",
                    mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                    use_container_width=True,
                    key="sd_dl_srt"
                )

    # =====================================================================
    # TAB 3: DATA TIM BARU JOIN & KELUAR (MUTASI PERSONEL)
    # =====================================================================
    with tab_prs:
        st.markdown("### 👥 Data Personel Tim (Tanggal Join & Tanggal Keluar)")

        with st.expander("➕ Input Tim Baru Join / Catat Tim Keluar", expanded=df_prs.empty):
            cp1, cp2, cp3 = st.columns(3)
            with cp1:
                prs_nop = st.selectbox("Area NOP:", ["Palangkaraya", "Pangkalan Bun"], key="sd_prs_nop")
                prs_nama = st.text_input("Nama Lengkap Personel *:", key="sd_prs_nama").strip()
                prs_nik = st.text_input("NIK / ID Karyawan:", key="sd_prs_nik").strip()
            with cp2:
                prs_role = st.selectbox("Jabatan / Role:", ["FME - TE", "FME - MBP", "FME - CME", "Leader", "Koordinator", "Admin Operational", "Admin NOP", "Helpdesk", "Logistik", "Lainnya"], key="sd_prs_role")
                prs_stat = st.radio("Status Kerja Saat Ini *:", ["Aktif", "Keluar / Resign", "Mutasi"], horizontal=True, key="sd_prs_stat")
            with cp3:
                prs_join = st.date_input("📅 Tanggal Join (Masuk):", value=datetime.date.today(), key="sd_prs_join").strftime("%Y-%m-%d")
                if prs_stat != "Aktif":
                    prs_out = st.date_input("📅 Tanggal Keluar / Mutasi:", value=datetime.date.today(), key="sd_prs_out").strftime("%Y-%m-%d")
                else:
                    prs_out = "-"
                    st.info("Status **Aktif** — Tanggal keluar otomatis `-`.")

            prs_ket = st.text_input("📝 Keterangan / Alasan Keluar / Cluster Penempatan:", key="sd_prs_ket")

            if st.button("💾 SIMPAN DATA PERSONEL", type="primary", use_container_width=True, key="sd_btn_save_prs"):
                if not prs_nama:
                    st.error("⚠️ Nama Personel wajib diisi!")
                else:
                    rec_prs = {
                        "Timestamp": datetime.datetime.now().strftime("%Y-%m-%d %H:%M:%S"),
                        "NOP": prs_nop,
                        "Nama_Tim": prs_nama,
                        "NIK_ID": prs_nik,
                        "Jabatan_Role": prs_role,
                        "Status_Kerja": prs_stat,
                        "Tanggal_Join": prs_join,
                        "Tanggal_Keluar": prs_out,
                        "Alasan_Keterangan": prs_ket
                    }
                    if gsheet_append_row("SiDAVID_Personel", COLS_PERSONEL, rec_prs):
                        st.success(f"✅ Data personel **{prs_nama}** ({prs_stat}) berhasil dicatat!")
                        st.rerun()
                    else:
                        st.error("⚠️ Gagal menyambung ke Google Sheets.")

        if df_prs.empty:
            st.info("Belum ada data histori Join / Keluar personel tim.")
        else:
            cnt_act = int((df_prs["Status_Kerja"] == "Aktif").sum())
            cnt_out = int((df_prs["Status_Kerja"] != "Aktif").sum())
            mp1, mp2, mp3 = st.columns(3)
            mp1.metric("👥 Total Data Tercatat", f"{len(df_prs)} Orang")
            mp2.metric("🟢 Personel Aktif", f"{cnt_act} Orang")
            mp3.metric("🔴 Personel Keluar / Mutasi", f"{cnt_out} Orang")

            st.markdown("#### 📋 Database Join & Keluar Tim (Klik 2x untuk Update Tanggal Keluar)")
            edited_prs = st.data_editor(
                df_prs,
                use_container_width=True,
                num_rows="dynamic",
                column_config={
                    "NOP": st.column_config.SelectboxColumn("NOP", options=["Palangkaraya", "Pangkalan Bun"]),
                    "Status_Kerja": st.column_config.SelectboxColumn("Status Kerja", options=["Aktif", "Keluar / Resign", "Mutasi"])
                },
                key="sd_ed_prs"
            )
            c_bp1, c_bp2 = st.columns(2)
            with c_bp1:
                if st.button("💾 Simpan Perubahan Data Personel", type="primary", use_container_width=True, key="sd_upd_prs"):
                    if gsheet_overwrite_ws("SiDAVID_Personel", edited_prs):
                        st.success("✅ Data Join/Keluar personel berhasil diperbarui!")
                        st.rerun()
            with c_bp2:
                buf_pr = BytesIO()
                with pd.ExcelWriter(buf_pr, engine="openpyxl") as xw:
                    edited_prs.to_excel(xw, sheet_name="Data_Tim_Join_Keluar", index=False)
                st.download_button(
                    "📥 Download Data Personel (.xlsx)",
                    data=buf_pr.getvalue(),
                    file_name=f"SiDAVID_Personel_{datetime.date.today()}.xlsx",
                    mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                    use_container_width=True,
                    key="sd_dl_prs"
                )

    # =====================================================================
    # TAB 4: VALIDASI DATA LAINNYA (MODUL STANDBY FLEKSIBEL)
    # =====================================================================
    with tab_lnn:
        st.markdown("### 🧩 Validasi Data Internal Lainnya (Modul Standby)")
        st.info("💡 Ruang validasi fleksibel yang siap pakai untuk kebutuhan audit internal lainnya (misal: **APD & Full Body Harness, SIM A/C, BPJS/KTP, MCU Kesehatan, Kalibrasi Alat Ukur, STNK Kendaraan Ops, SIK**, atau kategori baru sesuai kebutuhan NOP).")

        with st.expander("➕ Tambah Item Validasi Data Lainnya", expanded=df_lnn.empty):
            cl1, cl2, cl3 = st.columns(3)
            with cl1:
                lnn_nop = st.selectbox("Area NOP:", ["Palangkaraya", "Pangkalan Bun"], key="sd_lnn_nop")
                kat_opsi_std = ["APD & Full Body Harness", "SIM A / SIM C", "BPJS & KTP", "MCU / Kesehatan", "Kalibrasi Alat Ukur", "STNK / Kendaraan Ops", "Surat Izin Kerja / SIK", "+ Kategori Baru (Ketik Manual)"]
                lnn_kat_sel = st.selectbox("Kategori Validasi:", kat_opsi_std, key="sd_lnn_kat_sel")
                if lnn_kat_sel == "+ Kategori Baru (Ketik Manual)":
                    lnn_kat = st.text_input("Ketik Nama Kategori Validasi Baru *:", key="sd_lnn_kat_txt").strip()
                else:
                    lnn_kat = lnn_kat_sel
            with cl2:
                lnn_obj = st.text_input("Nama Personel / Aset / Objek *:", placeholder="Cth: Nama Teknisi atau No. Plat Mobil", key="sd_lnn_obj").strip()
                lnn_punya = st.selectbox("Keterangan Kepemilikan / Ada:", ["Punya / Ada", "Tidak Punya / Belum Ada"], key="sd_lnn_punya")
                lnn_stat = st.selectbox("Status Validasi:", ["Valid / Lengkap", "Pending / Proses", "Tidak Valid / Expired"], key="sd_lnn_stat")
            with cl3:
                lnn_terbit = st.date_input("📅 Tanggal Terbit / Cek:", value=datetime.date.today(), key="sd_lnn_terbit").strftime("%Y-%m-%d")
                pakai_exp = st.checkbox("Memiliki Tanggal Expired?", value=True, key="sd_lnn_chk_exp")
                lnn_exp = st.date_input("📅 Tanggal Expired:", value=datetime.date.today() + datetime.timedelta(days=365), key="sd_lnn_exp").strftime("%Y-%m-%d") if pakai_exp else "-"

            cl4, cl5 = st.columns(2)
            lnn_cat = cl4.text_input("📝 Catatan / Remark Validasi:", key="sd_lnn_cat")
            lnn_file = cl5.file_uploader("📎 Upload Bukti Dokumen / Foto (Opsional):", type=["png", "jpg", "jpeg", "pdf"], key="sd_lnn_file")

            if st.button("💾 SIMPAN VALIDASI DATA LAINNYA", type="primary", use_container_width=True, key="sd_btn_save_lnn"):
                if not lnn_kat or not lnn_obj:
                    st.error("⚠️ Kategori Validasi dan Nama Objek/Personel wajib diisi!")
                else:
                    with st.spinner("Menyimpan validasi internal..."):
                        url_lnn = ""
                        if lnn_file is not None:
                            url_lnn = upload_datalog(lnn_file, "sidavid_lainnya") if lnn_file.name.lower().endswith(".pdf") else upload_image(lnn_file, "sidavid_lainnya")
                        rec_lnn = {
                            "Timestamp": datetime.datetime.now().strftime("%Y-%m-%d %H:%M:%S"),
                            "NOP": lnn_nop,
                            "Kategori_Validasi": lnn_kat,
                            "Objek_Atau_Nama": lnn_obj,
                            "Keterangan_Punya": lnn_punya,
                            "Tanggal_Terbit": lnn_terbit,
                            "Tanggal_Expired": lnn_exp,
                            "Status_Validasi": lnn_stat,
                            "URL_Bukti": url_lnn or "",
                            "Catatan": lnn_cat
                        }
                        if gsheet_append_row("SiDAVID_Validasi_Lain", COLS_LAINNYA, rec_lnn):
                            st.success(f"✅ Data validasi **{lnn_kat}** ({lnn_obj}) berhasil disimpan!")
                            st.rerun()
                        else:
                            st.error("⚠️ Gagal menyambung ke Google Sheets.")

        if df_lnn.empty:
            st.info("Modul Standby Validasi Lainnya siap digunakan. Tambahkan item pertama melalui form di atas.")
        else:
            edited_lnn = st.data_editor(
                df_lnn,
                use_container_width=True,
                num_rows="dynamic",
                column_config={
                    "Status_Validasi": st.column_config.SelectboxColumn("Status Validasi", options=["Valid / Lengkap", "Pending / Proses", "Tidak Valid / Expired"]),
                    "URL_Bukti": st.column_config.LinkColumn("Bukti")
                },
                key="sd_ed_lnn"
            )
            c_bl1, c_bl2 = st.columns(2)
            with c_bl1:
                if st.button("💾 Simpan Perubahan Validasi Lainnya", type="primary", use_container_width=True, key="sd_upd_lnn"):
                    if gsheet_overwrite_ws("SiDAVID_Validasi_Lain", edited_lnn):
                        st.success("✅ Perubahan Validasi Lainnya berhasil disimpan!")
                        st.rerun()
            with c_bl2:
                buf_ln = BytesIO()
                with pd.ExcelWriter(buf_ln, engine="openpyxl") as xw:
                    edited_lnn.to_excel(xw, sheet_name="Validasi_Data_Lainnya", index=False)
                st.download_button(
                    "📥 Download Rekap Validasi Lainnya (.xlsx)",
                    data=buf_ln.getvalue(),
                    file_name=f"SiDAVID_Validasi_Lainnya_{datetime.date.today()}.xlsx",
                    mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                    use_container_width=True,
                    key="sd_dl_lnn"
                )

# =========================================================================
# MENU 14: FORM & PETA HARD AKSES (PANGKALAN BUN & PALANGKARAYA -> SHEET 'Site Hard Akses')
# =========================================================================
elif menu == "🚧 Form & Peta Hard Akses":
    st.markdown("<h1 style='text-align: center;'>🚧 Form & Peta Rute Hard Akses Site</h1>", unsafe_allow_html=True)
    st.markdown("<p style='text-align: center; opacity: 0.85;'>Pendataan Lokasi Hard Akses NOP Pangkalan Bun & Palangkaraya, Rute Wajib Sewa Moda (Trail / 4WD / Klotok), Upload Eviden Foto & Peta Rute Interaktif — Tersimpan di Sheet <b><code>Site Hard Akses</code></b></p>", unsafe_allow_html=True)

    # Muat koordinat site hanya saat diperlukan agar tetap ringan
    site_plk_list_ha, coords_plk_ha = load_site_master_plk()
    site_pbu_list_ha, coords_pbu_ha = load_site_master_pbu()
    df_sha_raw, ok_sha, ha_all_list, _ = load_site_hard_akses(
        nop_filter=None,
        site_coords_plk=coords_plk_ha,
        site_coords_pbu=coords_pbu_ha
    )

    is_admin_sha = (st.session_state.get("role") == "Admin") or bool(st.session_state.get("sha_admin_unlocked", False))

    cnt_sha_tot = len(ha_all_list)
    cnt_sha_plk = sum(1 for h in ha_all_list if "palangka" in h["nop"].lower())
    cnt_sha_pbu = sum(1 for h in ha_all_list if "pangkalan" in h["nop"].lower())
    cnt_sha_blk = sum(1 for h in ha_all_list if "blocked" in h["status"].lower() or "tidak tembus" in h["status"].lower())
    cnt_sha_pic = sum(len(h["photos"]) for h in ha_all_list)

    mh1, mh2, mh3, mh4, mh5 = st.columns(5)
    mh1.metric("🚧 Total Site Hard Akses", f"{cnt_sha_tot} Lokasi")
    mh2.metric("🔵 NOP Palangkaraya", f"{cnt_sha_plk} Lokasi")
    mh3.metric("🔴 NOP Pangkalan Bun", f"{cnt_sha_pbu} Lokasi")
    mh4.metric("🚫 Blocked / Tidak Tembus", f"{cnt_sha_blk} Lokasi")
    mh5.metric("📸 Total Eviden Foto", f"{cnt_sha_pic} Foto")

    sha_nav = st.radio(
        "Navigasi Hard Akses:",
        [
            "🗺️ Peta Rute & Titik Hard Akses",
            "📝 Form Input Hard Akses (PBU & PLK + Moda Sewa & Foto)",
            "📋 Galeri Eviden, Update Status & Export"
        ],
        horizontal=True,
        key="sha_nav_radio",
        label_visibility="collapsed"
    )
    st.markdown("<hr style='margin: 8px 0 18px 0;'>", unsafe_allow_html=True)

    # ---------------------------------------------------------------------
    # SUB-1: PETA RUTE HARD AKSES (PLK & PBU)
    # ---------------------------------------------------------------------
    if sha_nav == "🗺️ Peta Rute & Titik Hard Akses":
        st.markdown("#### 🗺️ Peta Interaktif Rute & Titik Kendala Hard Akses (Palangkaraya & Pangkalan Bun)")
        st.info("💡 Klik ikon **🚧** di peta untuk melihat **Foto Kondisi Jalan**, **Moda Sewa Wajib (Trail / 4WD / Klotok)**, estimasi biaya, dan rute detail. Garis beranimasi menghubungkan **🏠 Pos/Start ➔ 🚧 Titik Hard Akses ➔ 🗼 Site Tujuan**.")

        fm1, fm2, fm3, fm4 = st.columns([1.2, 1.2, 1.3, 1.1])
        flt_nop_m = fm1.selectbox("🏢 Filter NOP:", ["Semua NOP", "Pangkalan Bun", "Palangkaraya"], key="sha_map_nop")
        flt_st_m = fm2.selectbox("🚦 Filter Status:", ["Semua Status"] + HARD_ACCESS_STATUS, key="sha_map_st")
        flt_moda_m = fm3.selectbox("🚙 Filter Moda Wajib:", ["Semua Moda", "Motor Trail", "4WD", "Klotok", "Ferry", "Alat Berat", "Porter", "Corporate"], key="sha_map_moda")
        tema_sha_m = fm4.selectbox("🛰️ Base Layer:", ["Dark NOC Command Center", "Citra Satelit (Esri)", "Peta Jalan Terang (Street Map)"], key="sha_map_theme")

        ha_map_items = [
            h for h in ha_all_list
            if (flt_nop_m == "Semua NOP" or flt_nop_m.lower() in h["nop"].lower())
            and (flt_st_m == "Semua Status" or h["status"] == flt_st_m)
            and (flt_moda_m == "Semua Moda" or flt_moda_m.lower() in h["kendaraan"].lower())
        ]

        if not ha_map_items:
            st.warning("Belum ada titik Hard Akses untuk filter ini. Silakan input melalui tab **📝 Form Input Hard Akses**.")
        else:
            st.components.v1.html(
                render_hard_access_route_map_html(ha_map_items, map_theme=tema_sha_m, show_routes=True),
                height=605
            )

            # Ringkasan singkat di bawah peta
            df_m_prev = pd.DataFrame([{
                "NOP": h["nop"],
                "Site ID": h["site_id"],
                "Site Name": h["site_name"],
                "Lokasi Kendala": h["lokasi"],
                "Moda Wajib / Sewa": h["kendaraan"],
                "Est. Biaya Sewa (Rp)": h["biaya_sewa"],
                "Est. Waktu": h["estimasi"],
                "Status": h["status"],
                "Foto": f"{len(h['photos'])} Foto"
            } for h in ha_map_items])
            st.dataframe(df_m_prev, use_container_width=True, hide_index=True)

    # ---------------------------------------------------------------------
    # SUB-2: FORM INPUT HARD AKSES (PBU & PLK + MODA WAJIB + FOTO) -> SHEET 'Site Hard Akses'
    # ---------------------------------------------------------------------
    elif sha_nav == "📝 Form Input Hard Akses (PBU & PLK + Moda Sewa & Foto)":
        st.markdown("#### 📝 Form Input Data Lokasi & Rute Wajib Hard Akses ➔ Sheet `Site Hard Akses`")

        if not is_admin_sha:
            with st.container(border=True):
                st.warning("🔐 Pengisian / perubahan rute wajib Hard Akses diprioritaskan untuk Admin / Koordinator.")
                c_ul1, c_ul2 = st.columns([2, 1])
                pwd_sha = c_ul1.text_input("Masukkan Kata Sandi Admin:", type="password", key="sha_pwd_in", label_visibility="collapsed", placeholder="Kata sandi Admin...")
                if c_ul2.button("🔓 Buka Form Admin", type="primary", use_container_width=True, key="btn_sha_unlock"):
                    if pwd_sha == "KUT2027":
                        st.session_state["sha_admin_unlocked"] = True
                        st.rerun()
                    else:
                        st.error("❌ Sandi salah!")

        if is_admin_sha:
            with st.container(border=True):
                sha_nop = st.radio("🏢 Pilih Area NOP *:", ["🔴 Pangkalan Bun", "🔵 Palangkaraya"], horizontal=True, key="sha_in_nop")
                is_pbu_sel = "Pangkalan Bun" in sha_nop
                nop_clean = "Pangkalan Bun" if is_pbu_sel else "Palangkaraya"
                s_list_use = site_pbu_list_ha if is_pbu_sel else site_plk_list_ha
                c_dict_use = coords_pbu_ha if is_pbu_sel else coords_plk_ha

                c1, c2 = st.columns(2)
                with c1:
                    sid_pick = st.selectbox(f"📍 Pilih Site ID ({nop_clean} — {len(s_list_use)} Master Site) *:", ["-- Ketik Manual --"] + s_list_use, key=f"sha_sid_{nop_clean}")
                    if sid_pick == "-- Ketik Manual --":
                        sha_sid = st.text_input("Ketik Site ID *:", key="sha_sid_man").strip().upper()
                        sha_sid_norm = _norm_site_id(sha_sid)
                    else:
                        sha_sid = sid_pick
                        sha_sid_norm = _norm_site_id(sid_pick)

                    c_info = c_dict_use.get(sha_sid_norm)
                    def_s_lat = c_info[0] if c_info else (-2.35 if is_pbu_sel else -2.21)
                    def_s_lon = c_info[1] if c_info else (111.60 if is_pbu_sel else 113.92)
                    def_s_nm = c_info[2] if (c_info and len(c_info) >= 3) else (sid_pick.split(" - ", 1)[1] if " - " in sid_pick else sha_sid_norm)

                    sha_sname = st.text_input("Nama Site:", value=str(def_s_nm), key=f"sha_snm_{sha_sid_norm}")
                    sha_lokasi = st.text_input("🚧 Lokasi / Titik Kendala Hard Akses *:", placeholder="Cth: Km 18 Jl. Poros Kuala Kuayan / Penyeberangan Sungai Kahayan", key="sha_lok")

                with c2:
                    sha_kat = st.selectbox("🏷️ Kategori Kendala Akses *:", HARD_ACCESS_KATEGORI, key="sha_kat")
                    sha_lvl = st.selectbox("⚠️ Tingkat Kesulitan *:", HARD_ACCESS_LEVEL, key="sha_lvl")
                    sha_stat = st.selectbox("🚦 Status Akses Saat Ini *:", HARD_ACCESS_STATUS, key="sha_stat")

                st.markdown("---")
                st.markdown("##### 🚙 Rute Wajib & Kebutuhan Sewa Moda Transportasi (Diisi Admin)")
                cm1, cm2, cm3 = st.columns([2.2, 1.1, 1.1])
                sha_moda_sel = cm1.multiselect(
                    "Pilih Moda / Sewa Wajib pada Rute Ini * (Bisa pilih >1):",
                    HARD_ACCESS_MODA_WAJIB,
                    default=["🚙 Wajib 4WD / Double Cabin (4x4)"],
                    key="sha_moda_ms"
                )
                sha_moda_custom = cm1.text_input("➕ Tambah Moda / Catatan Sewa Lainnya (Opsional):", placeholder="Cth: Sewa Ketinting / Escort Loader Sawit", key="sha_moda_cust")
                sha_biaya_sewa = cm2.number_input("💰 Est. Biaya Sewa Moda (Rp):", min_value=0, value=0, step=50000, key="sha_biaya")
                sha_est_waktu = cm3.text_input("⏱️ Est. Waktu Tempuh:", placeholder="Cth: 5 Jam PP", key="sha_waktu")

                sha_rute_det = st.text_area(
                    "🛣️ Uraian Rute Wajib & Titik Estafet Moda:",
                    placeholder="Cth: Dari Pos Kuala Kuayan pakai Hilux 4x4 sampai Km 15, lanjut sewa Motor Trail 8 km ke site / naik Klotok 45 menit dari Dermaga...",
                    height=75,
                    key="sha_rute_det"
                )

                st.markdown("##### 🌐 Koordinat Peta Rute (Pos Start ➔ Titik Hard Akses ➔ Site Tujuan)")
                # Pilihan cepat ambil Koordinat Start dari Homebase Tim PBU/PLK
                df_hb_sel, _, _ = load_master_tim_df("PBU" if is_pbu_sel else "PLK")
                hb_pick_opts = ["-- Kosong / Input Manual --"]
                hb_coord_map = {}
                if not df_hb_sel.empty:
                    c_la_h = next((c for c in df_hb_sel.columns if "lat" in c.lower()), None)
                    c_lo_h = next((c for c in df_hb_sel.columns if "lon" in c.lower()), None)
                    if c_la_h and c_lo_h:
                        for _, rhb in df_hb_sel.iterrows():
                            la_h = pd.to_numeric(rhb.get(c_la_h), errors="coerce")
                            lo_h = pd.to_numeric(rhb.get(c_lo_h), errors="coerce")
                            nm_h = str(rhb.get("NAMA", "")).strip()
                            if nm_h and pd.notna(la_h) and pd.notna(lo_h):
                                lbl_h = f"HB {nm_h} ({la_h:.4f}, {lo_h:.4f})"
                                hb_pick_opts.append(lbl_h)
                                hb_coord_map[lbl_h] = (f"HB {nm_h}", float(la_h), float(lo_h))

                pick_hb_st = st.selectbox("🏠 Pilih Titik Start / Pos dari Homebase Tim (Opsional — untuk gambar garis rute di peta):", hb_pick_opts, key=f"sha_hb_pick_{nop_clean}")
                if pick_hb_st in hb_coord_map:
                    def_pos_nm, def_st_lat, def_st_lon = hb_coord_map[pick_hb_st]
                else:
                    def_pos_nm, def_st_lat, def_st_lon = "", 0.0, 0.0

                cg1, cg2, cg3 = st.columns(3)
                with cg1:
                    sha_pos_nm = st.text_input("Nama Pos / Titik Start:", value=def_pos_nm, placeholder="Cth: Pos Pangkalan Bun / Dermaga", key=f"sha_posnm_{pick_hb_st}")
                    sha_st_lat = st.number_input("Lat Start (0 = tanpa garis start):", value=float(def_st_lat), format="%.6f", key=f"sha_stlat_{pick_hb_st}")
                    sha_st_lon = st.number_input("Lon Start (0 = tanpa garis start):", value=float(def_st_lon), format="%.6f", key=f"sha_stlon_{pick_hb_st}")
                with cg2:
                    sha_k_lat = st.number_input("Lat Titik Hard Akses *:", value=float(def_s_lat), format="%.6f", key=f"sha_klat_{sha_sid_norm}")
                    sha_k_lon = st.number_input("Lon Titik Hard Akses *:", value=float(def_s_lon), format="%.6f", key=f"sha_klon_{sha_sid_norm}")
                with cg3:
                    sha_s_lat = st.number_input("Lat Site Tujuan (Auto Master):", value=float(def_s_lat), format="%.6f", key=f"sha_slat_{sha_sid_norm}")
                    sha_s_lon = st.number_input("Lon Site Tujuan (Auto Master):", value=float(def_s_lon), format="%.6f", key=f"sha_slon_{sha_sid_norm}")

                st.markdown("---")
                sha_pic = st.text_input("👨‍💼 Admin / Koordinator Pelapor *:", value=f"Admin {nop_clean}", key="sha_pic_in")
                sha_ket = st.text_area("📝 Keterangan Detail Kondisi Lapangan *:", placeholder="Jelaskan kondisi jalan, tinggi banjir/lumpur, nomor kontak penyedia klotok/trail, dll...", key="sha_ket_in")
                sha_pics = st.file_uploader("📸 Upload Foto-Foto Eviden Hard Akses (Bisa Pilih Banyak Foto):", type=["png", "jpg", "jpeg", "webp"], accept_multiple_files=True, key="sha_pics_up")

                if st.button("🚀 SIMPAN DATA HARD AKSES & FOTO KE SHEET `Site Hard Akses`", type="primary", use_container_width=True, key="btn_commit_sha"):
                    moda_final_list = list(sha_moda_sel) + ([sha_moda_custom.strip()] if sha_moda_custom.strip() else [])
                    if not sha_sid_norm or not sha_lokasi.strip() or not sha_ket.strip() or not moda_final_list:
                        st.error("⚠️ Site ID, Lokasi Kendala, minimal 1 Moda Wajib, dan Keterangan Detail wajib diisi!")
                    else:
                        with st.spinner("⏳ Mengunggah eviden foto ke Cloudinary & menyimpan ke Google Sheets tab `Site Hard Akses`..."):
                            urls_sha = upload_multiple_images(sha_pics, "site_hard_akses") if sha_pics else []
                            rec_sha = {
                                "Record_ID": f"SHA-{'PBU' if is_pbu_sel else 'PLK'}-{datetime.datetime.now().strftime('%Y%m%d%H%M%S')}",
                                "Timestamp": datetime.datetime.now().strftime("%Y-%m-%d %H:%M:%S"),
                                "NOP": nop_clean,
                                "Site_ID": sha_sid_norm,
                                "Site_Name": sha_sname.strip() or sha_sid_norm,
                                "Titik_Start_Pos": sha_pos_nm.strip() or "-",
                                "Lat_Start": f"{sha_st_lat:.6f}" if sha_st_lat != 0.0 else "",
                                "Lon_Start": f"{sha_st_lon:.6f}" if sha_st_lon != 0.0 else "",
                                "Lokasi_Kendala": sha_lokasi.strip(),
                                "Lat_Kendala": f"{sha_k_lat:.6f}",
                                "Lon_Kendala": f"{sha_k_lon:.6f}",
                                "Lat_Site": f"{sha_s_lat:.6f}",
                                "Lon_Site": f"{sha_s_lon:.6f}",
                                "Kategori_Hard_Access": sha_kat,
                                "Tingkat_Kesulitan": sha_lvl,
                                "Status_Akses": sha_stat,
                                "Rute_Wajib_Moda": ", ".join(moda_final_list),
                                "Estimasi_Biaya_Sewa": str(int(sha_biaya_sewa)),
                                "Estimasi_Waktu_Tempuh": sha_est_waktu.strip() or "-",
                                "Rekomendasi_Rute_Detail": sha_rute_det.strip() or "-",
                                "PIC_Input": sha_pic.strip() or "Admin",
                                "Keterangan_Detail": sha_ket.strip(),
                                "Foto_URLs": json.dumps(urls_sha, ensure_ascii=False)
                            }
                            if gsheet_append_row("Site Hard Akses", COLS_SITE_HARD_AKSES, rec_sha):
                                st.success(f"✅ Data Hard Akses **{sha_sid_norm} ({nop_clean})** beserta {len(urls_sha)} foto berhasil disimpan ke sheet **`Site Hard Akses`**!")
                                st.rerun()
                            else:
                                st.error("⚠️ Gagal menyimpan ke Google Sheets tab `Site Hard Akses`.")

    # ---------------------------------------------------------------------
    # SUB-3: GALERI EVIDEN, UPDATE STATUS/FOTO SUSULAN & EXPORT EXCEL
    # ---------------------------------------------------------------------
    else:
        st.markdown("#### 📋 Galeri Eviden Foto, Update Status & Database Sheet `Site Hard Akses`")
        if not ha_all_list:
            st.info("Belum ada data Hard Akses yang tercatat.")
        else:
            cg1, cg2, cg3 = st.columns([1.2, 1.5, 1.3])
            flt_g_nop = cg1.selectbox("Filter NOP:", ["Semua NOP", "Pangkalan Bun", "Palangkaraya"], key="sha_gal_nop")
            q_g_txt = cg2.text_input("🔍 Cari Site / Lokasi / Moda (Trail/4WD/Klotok):", key="sha_gal_q")
            flt_g_st = cg3.selectbox("Filter Status:", ["Semua Status"] + HARD_ACCESS_STATUS, key="sha_gal_st")

            ha_gal = [
                h for h in ha_all_list
                if (flt_g_nop == "Semua NOP" or flt_g_nop.lower() in h["nop"].lower())
                and (flt_g_st == "Semua Status" or h["status"] == flt_g_st)
                and (not q_g_txt.strip() or q_g_txt.strip().lower() in f"{h['site_id']} {h['site_name']} {h['lokasi']} {h['kendaraan']} {h['kategori']} {h['keterangan']}".lower())
            ]

            for idx_g, ha in enumerate(reversed(ha_gal)):
                st_l = ha["status"].lower()
                bc = "#10B981" if ("normal" in st_l) else ("#F59E0B" if ("terbatas" in st_l or "moda" in st_l) else "#EF4444")
                with st.container(border=True):
                    st.markdown(f"""
                        <div style="border-left:6px solid {bc}; padding-left:14px;">
                            <div style="display:flex; justify-content:space-between; flex-wrap:wrap; gap:8px; align-items:center;">
                                <span style="font-size:17px; font-weight:800;">🚧 [{ha['nop']}] {ha['site_id']} — {ha['site_name']}</span>
                                <span style="background:{bc}; color:#fff; padding:3px 10px; border-radius:99px; font-size:12px; font-weight:800;">{ha['status']}</span>
                            </div>
                            <div style="font-size:13px; margin-top:6px; line-height:1.65;">
                                📍 <b>Lokasi Kendala:</b> {ha['lokasi']} &nbsp;|&nbsp; 🏷️ <b>Kategori:</b> {ha['kategori']}<br>
                                🚙 <b>Moda Sewa Wajib:</b> <b>{ha['kendaraan']}</b> &nbsp;|&nbsp; 💰 <b>Est. Biaya:</b> Rp {ha['biaya_sewa']:,.0f} &nbsp;|&nbsp; ⏱️ <b>Waktu:</b> {ha['estimasi']}<br>
                                🛣️ <b>Rute Wajib:</b> {ha['rekomendasi_rute']}<br>
                                📝 <b>Detail:</b> {ha['keterangan']}<br>
                                <span style="opacity:0.75; font-size:12px;">👨‍💼 {ha['pic']} · 📅 {ha['timestamp']} · 🌐 {ha['lat']}, {ha['lon']}</span>
                            </div>
                        </div>
                    """, unsafe_allow_html=True)

                    cb1, cb2 = st.columns(2)
                    if ha["lat"] is not None and ha["lon"] is not None:
                        cb1.link_button("🧭 Buka Titik Hard Akses di Maps", f"https://www.google.com/maps?q={ha['lat']},{ha['lon']}", use_container_width=True)
                    wa_sha = (
                        f"*🚧 REPORT HARD AKSES — NOP {ha['nop'].upper()}*\n"
                        f"📍 *Site:* {ha['site_id']} ({ha['site_name']})\n"
                        f"🗺️ *Lokasi Kendala:* {ha['lokasi']}\n"
                        f"🏷️ *Kategori:* {ha['kategori']} | *Status:* {ha['status']}\n"
                        f"🚙 *Rute Wajib / Moda Sewa:* {ha['kendaraan']}\n"
                        f"💰 *Est. Biaya Sewa:* Rp {ha['biaya_sewa']:,.0f} | ⏱️ {ha['estimasi']}\n"
                        f"🛣️ *Rekomendasi Rute:* {ha['rekomendasi_rute']}\n"
                        f"📝 *Detail:* {ha['keterangan']}\n"
                        f"🌐 *Maps:* https://www.google.com/maps?q={ha['lat']},{ha['lon']}"
                    )
                    cb2.link_button("📱 Share Report ke WhatsApp", f"https://wa.me/?text={urllib.parse.quote(wa_sha)}", use_container_width=True)

                    if ha["photos"]:
                        tampilkan_grid_foto(ha["photos"], f"📸 Eviden Foto Hard Akses ({len(ha['photos'])} Foto)")

                    if is_admin_sha and ha.get("source_sheet") == "Site Hard Akses":
                        with st.expander(f"🛠️ Admin: Update Status / Moda Wajib / Tambah Foto ({ha['site_id']})", expanded=False):
                            cu1, cu2 = st.columns(2)
                            idx_st = HARD_ACCESS_STATUS.index(ha["status"]) if ha["status"] in HARD_ACCESS_STATUS else 0
                            u_st = cu1.selectbox("Status Akses Baru:", HARD_ACCESS_STATUS, index=idx_st, key=f"sha_ust_{ha['record_id']}_{idx_g}")
                            u_moda = cu2.text_input("Update Moda Wajib / Sewa:", value=ha["kendaraan"], key=f"sha_umoda_{ha['record_id']}_{idx_g}")
                            u_rute = st.text_input("Update Uraian Rute Wajib:", value=ha["rekomendasi_rute"], key=f"sha_urute_{ha['record_id']}_{idx_g}")
                            u_pics = st.file_uploader("➕ Tambah Foto Susulan:", type=["png", "jpg", "jpeg", "webp"], accept_multiple_files=True, key=f"sha_upic_{ha['record_id']}_{idx_g}")
                            if st.button("💾 Simpan Update ke `Site Hard Akses`", type="primary", key=f"btn_sha_upd_{ha['record_id']}_{idx_g}"):
                                with st.spinner("Memperbarui data & mengunggah foto..."):
                                    ex_u = upload_multiple_images(u_pics, "site_hard_akses") if u_pics else []
                                    merged_p = ha["photos"] + ex_u
                                    df_u = df_sha_raw.copy()
                                    mk = df_u["Record_ID"].astype(str) == str(ha["record_id"])
                                    if mk.any():
                                        df_u.loc[mk, "Status_Akses"] = u_st
                                        df_u.loc[mk, "Rute_Wajib_Moda"] = u_moda
                                        df_u.loc[mk, "Rekomendasi_Rute_Detail"] = u_rute
                                        df_u.loc[mk, "Foto_URLs"] = json.dumps(merged_p, ensure_ascii=False)
                                        if gsheet_overwrite_ws("Site Hard Akses", df_u):
                                            st.success("✅ Data Hard Akses berhasil diperbarui!")
                                            st.rerun()

            if is_admin_sha and not df_sha_raw.empty:
                with st.expander("🗂️ Live Editor Massal Sheet `Site Hard Akses` (Admin)", expanded=False):
                    ed_sha = st.data_editor(
                        df_sha_raw,
                        use_container_width=True,
                        num_rows="dynamic",
                        column_config={
                            "NOP": st.column_config.SelectboxColumn("NOP", options=["Pangkalan Bun", "Palangkaraya"]),
                            "Status_Akses": st.column_config.SelectboxColumn("Status Akses", options=HARD_ACCESS_STATUS),
                            "Tingkat_Kesulitan": st.column_config.SelectboxColumn("Tingkat Kesulitan", options=HARD_ACCESS_LEVEL)
                        },
                        key="ed_sha_table_all"
                    )
                    if st.button("💾 Simpan Perubahan Massal ke Sheet `Site Hard Akses`", type="primary", use_container_width=True, key="btn_save_ed_sha"):
                        if gsheet_overwrite_ws("Site Hard Akses", ed_sha):
                            st.success("✅ Sheet `Site Hard Akses` berhasil diperbarui!")
                            st.rerun()

            if st.button("🛠️ Siapkan Download Rekap `Site Hard Akses` (.xlsx)", use_container_width=True, key="btn_prep_sha_xl"):
                buf_sha = BytesIO()
                df_exp_sha = pd.DataFrame([{
                    "Record_ID": h["record_id"], "Timestamp": h["timestamp"], "NOP": h["nop"],
                    "Site_ID": h["site_id"], "Site_Name": h["site_name"], "Titik_Start": h["titik_start"],
                    "Lokasi_Kendala": h["lokasi"], "Kategori": h["kategori"], "Kesulitan": h["tingkat"],
                    "Status": h["status"], "Rute_Wajib_Moda": h["kendaraan"], "Est_Biaya_Sewa": h["biaya_sewa"],
                    "Est_Waktu": h["estimasi"], "Rekomendasi_Rute": h["rekomendasi_rute"],
                    "PIC": h["pic"], "Keterangan": h["keterangan"], "Jumlah_Foto": len(h["photos"]),
                    "Link_Foto": "\n".join(h["photos"])
                } for h in ha_gal])
                with pd.ExcelWriter(buf_sha, engine="openpyxl") as xw:
                    df_exp_sha.to_excel(xw, sheet_name="Site Hard Akses", index=False)
                st.session_state["sha_xlsx_bytes"] = buf_sha.getvalue()
            if st.session_state.get("sha_xlsx_bytes"):
                st.download_button(
                    "📥 Download Excel `Site Hard Akses` (.xlsx)",
                    data=st.session_state["sha_xlsx_bytes"],
                    file_name=f"Site_Hard_Akses_{datetime.date.today()}.xlsx",
                    mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                    type="primary",
                    use_container_width=True,
                    key="dl_sha_xl"
                )

# =========================================================================
# MENU 15: KELOLA MASTER TIM MANDIRI (ADMIN TAMBAH/HAPUS TIM & ROLE — PLK & PBU)
# =========================================================================
elif menu == "👥 Kelola Master Tim (PLK & PBU)":
    st.markdown("<h1 style='text-align: center;'>👥 Manajemen Master Tim Mandiri (PLK & PBU)</h1>", unsafe_allow_html=True)
    st.markdown("<p style='text-align: center; opacity: 0.85;'>Admin dapat menambah personel baru, mengubah Role, memperbarui Homebase (Lat/Lon), atau menghapus/menonaktifkan tim lama langsung ke Cloud — otomatis sinkron ke seluruh form SIDATA tanpa ubah file repositori!</p>", unsafe_allow_html=True)

    is_admin_mt = (st.session_state.get("role") == "Admin") or bool(st.session_state.get("mt_admin_unlocked", False))
    if not is_admin_mt:
        with st.container(border=True):
            st.markdown("<h3 style='text-align:center;'>🔐 Verifikasi Akses Admin Master Tim</h3>", unsafe_allow_html=True)
            c_mt_l1, c_mt_l2 = st.columns([2, 1])
            pwd_mt = c_mt_l1.text_input("Kata Sandi Admin:", type="password", key="mt_pwd_in", label_visibility="collapsed", placeholder="Masukkan kata sandi Admin...")
            if c_mt_l2.button("🔓 Masuk Mode Admin Tim", type="primary", use_container_width=True, key="btn_mt_unlock"):
                if pwd_mt == "KUT2027":
                    st.session_state["mt_admin_unlocked"] = True
                    st.rerun()
                else:
                    st.error("❌ Sandi salah!")
        st.stop()

    pilih_nop_mt = st.radio(
        "📍 Pilih Master Data Tim yang Akan Dikelola:",
        ["🔵 Master Tim NOP Palangkaraya (`Master_Tim_PLK`)", "🔴 Master Tim NOP Pangkalan Bun (`Master_HB_Tim_PBU`)"],
        horizontal=True,
        key="mt_nop_pick"
    )
    nop_k = "PLK" if "Palangkaraya" in pilih_nop_mt else "PBU"
    sh_target_mt = "Master_Tim_PLK" if nop_k == "PLK" else "Master_HB_Tim_PBU"
    cols_target_mt = COLS_MASTER_TIM_PLK if nop_k == "PLK" else COLS_MASTER_TIM_PBU

    df_mt, src_mt, ok_mt = load_master_tim_df(nop_k)

    tot_m = len(df_mt)
    cnt_akt = int((~df_mt["STATUS_AKTIF"].astype(str).str.lower().isin(["nonaktif", "resign", "keluar", "tidak aktif", "0", "false"])).sum()) if tot_m else 0
    cnt_non = tot_m - cnt_akt
    cnt_ldr = int(df_mt["ROLE"].astype(str).str.lower().str.contains("leader|koordinator").sum()) if tot_m else 0

    cm1, cm2, cm3, cm4 = st.columns(4)
    cm1.metric(f"Total Personel ({nop_k})", f"{tot_m} Orang")
    cm2.metric("🟢 Status Aktif (Muncul di Form)", f"{cnt_akt} Orang")
    cm3.metric("⚪ Nonaktif / Keluar", f"{cnt_non} Orang")
    cm4.metric("⭐ Leader / Koordinator", f"{cnt_ldr} Orang")
    st.caption(f"📡 Sumber aktif: `{src_mt}` — setiap perubahan yang disimpan langsung memperbarui dropdown seluruh form & peta SIDATA.")

    if "Excel Lokal" in src_mt and not df_mt.empty:
        if st.button(f"☁️ Migrasi & Simpan Data Awal ({len(df_mt)} Tim) ke Google Sheets `{sh_target_mt}` Sekarang", type="primary", use_container_width=True, key=f"btn_seed_mt_{nop_k}"):
            if gsheet_overwrite_ws(sh_target_mt, df_mt):
                st.success(f"✅ {len(df_mt)} personel berhasil dimigrasi ke Google Sheets `{sh_target_mt}`!")
                st.rerun()

    st.markdown("<hr style='margin: 12px 0 18px 0;'>", unsafe_allow_html=True)

    c_act_mt1, c_act_mt2 = st.columns(2)
    # Kiri: Tambah Tim Baru
    with c_act_mt1:
        with st.container(border=True):
            st.markdown(f"#### ➕ Tambah Personel Baru ({nop_k})")
            new_nm = st.text_input("Nama Lengkap Personel *:", placeholder="Cth: RIZKY PRATAMA", key=f"mt_add_nm_{nop_k}").strip().upper()
            ca1, ca2 = st.columns(2)
            new_role = ca1.selectbox("Role / Jabatan *:", MASTER_TIM_ROLES, key=f"mt_add_rl_{nop_k}")
            new_stat = ca2.selectbox("Status:", ["Aktif", "Nonaktif"], key=f"mt_add_st_{nop_k}")
            ca3, ca4 = st.columns(2)
            if nop_k == "PLK":
                new_cl_nop = ca3.selectbox("Cluster:", DAILY_CLUSTER_PLK_OPTS, key=f"mt_add_cl_{nop_k}")
            else:
                new_cl_nop = ca3.text_input("NOP:", value="NOP PANGKALAN BUN", key=f"mt_add_nop_{nop_k}")
            new_hp = ca4.text_input("No. HP / WhatsApp (Opsional):", placeholder="0812xxxx", key=f"mt_add_hp_{nop_k}")
            ca5, ca6 = st.columns(2)
            new_lat = ca5.text_input("Latitude Homebase (Opsional):", placeholder="-2.210000", key=f"mt_add_lat_{nop_k}")
            new_lon = ca6.text_input("Longitude Homebase (Opsional):", placeholder="113.920000", key=f"mt_add_lon_{nop_k}")

            if st.button(f"➕ TAMBAHKAN KE MASTER TIM {nop_k}", type="primary", use_container_width=True, key=f"btn_add_tim_{nop_k}"):
                if not new_nm:
                    st.error("⚠️ Nama Personel wajib diisi!")
                elif not df_mt.empty and new_nm in df_mt["NAMA"].astype(str).str.strip().str.upper().tolist():
                    st.warning(f"⚠️ Nama **{new_nm}** sudah ada di Master Tim {nop_k}.")
                else:
                    now_ts = datetime.datetime.now().strftime("%Y-%m-%d %H:%M")
                    if nop_k == "PLK":
                        new_row = {"NAMA": new_nm, "ROLE": new_role, "CLUSTER": new_cl_nop, "STATUS_AKTIF": new_stat, "NO_HP": new_hp.strip(), "LAT_HB": new_lat.strip(), "LON_HB": new_lon.strip(), "UPDATED_AT": now_ts}
                    else:
                        new_row = {"NAMA": new_nm, "NOP": new_cl_nop, "ROLE": new_role, "STATUS_AKTIF": new_stat, "NO_HP": new_hp.strip(), "Latitude Berangkat": new_lat.strip(), "Longtitude Berangkat": new_lon.strip(), "UPDATED_AT": now_ts}
                    df_merged = pd.concat([df_mt, pd.DataFrame([new_row])], ignore_index=True)
                    if gsheet_overwrite_ws(sh_target_mt, df_merged[cols_target_mt]):
                        st.success(f"✅ **{new_nm} ({new_role})** berhasil ditambahkan ke Master Tim `{sh_target_mt}` & langsung aktif di seluruh form!")
                        st.rerun()
                    else:
                        st.error("⚠️ Gagal menyimpan ke Google Sheets.")

    # Kanan: Hapus / Nonaktifkan Personel Lama
    with c_act_mt2:
        with st.container(border=True):
            st.markdown(f"#### 🗑️ Hapus / Ubah Status Personel ({nop_k})")
            if df_mt.empty:
                st.info("Belum ada data personel.")
            else:
                all_names_mt = df_mt["NAMA"].astype(str).tolist()
                pick_del_nm = st.selectbox("Pilih Nama Personel:", all_names_mt, key=f"mt_del_pick_{nop_k}")
                cur_row_mt = df_mt[df_mt["NAMA"].astype(str) == pick_del_nm].iloc[0]
                st.caption(f"Role saat ini: **{cur_row_mt.get('ROLE', '-')}** | Status: **{cur_row_mt.get('STATUS_AKTIF', 'Aktif')}**")

                upd_role_q = st.selectbox(
                    "Ganti Role Personel Ini:",
                    MASTER_TIM_ROLES,
                    index=MASTER_TIM_ROLES.index(cur_row_mt["ROLE"]) if cur_row_mt.get("ROLE") in MASTER_TIM_ROLES else 0,
                    key=f"mt_q_role_{nop_k}"
                )

                cd1, cd2, cd3 = st.columns(3)
                if cd1.button("💾 Update Role", use_container_width=True, key=f"btn_mt_url_{nop_k}"):
                    df_u = df_mt.copy()
                    df_u.loc[df_u["NAMA"].astype(str) == pick_del_nm, "ROLE"] = upd_role_q
                    df_u.loc[df_u["NAMA"].astype(str) == pick_del_nm, "UPDATED_AT"] = datetime.datetime.now().strftime("%Y-%m-%d %H:%M")
                    if gsheet_overwrite_ws(sh_target_mt, df_u[cols_target_mt]):
                        st.success(f"✅ Role **{pick_del_nm}** diubah menjadi **{upd_role_q}**!")
                        st.rerun()

                is_cur_act = str(cur_row_mt.get("STATUS_AKTIF", "Aktif")).strip().lower() not in ("nonaktif", "resign", "keluar", "tidak aktif")
                if cd2.button("⏸️ Nonaktifkan" if is_cur_act else "▶️ Aktifkan", use_container_width=True, key=f"btn_mt_tog_{nop_k}"):
                    df_u = df_mt.copy()
                    df_u.loc[df_u["NAMA"].astype(str) == pick_del_nm, "STATUS_AKTIF"] = "Nonaktif" if is_cur_act else "Aktif"
                    df_u.loc[df_u["NAMA"].astype(str) == pick_del_nm, "UPDATED_AT"] = datetime.datetime.now().strftime("%Y-%m-%d %H:%M")
                    if gsheet_overwrite_ws(sh_target_mt, df_u[cols_target_mt]):
                        st.success(f"✅ Status **{pick_del_nm}** diperbarui!")
                        st.rerun()

                if cd3.button("🗑️ Hapus Permanen", type="primary", use_container_width=True, key=f"btn_mt_del_{nop_k}"):
                    df_u = df_mt[df_mt["NAMA"].astype(str) != pick_del_nm].reset_index(drop=True)
                    if gsheet_overwrite_ws(sh_target_mt, df_u[cols_target_mt]):
                        st.success(f"🗑️ **{pick_del_nm}** berhasil dihapus permanen dari Master Tim `{sh_target_mt}`!")
                        st.rerun()

    # Live Editor Massal + Import/Export Excel
    st.markdown(f"#### ✏️ Live Editor Massal Master Tim {nop_k} (`{sh_target_mt}`)")
    st.caption("💡 Admin bisa langsung mengedit Nama, Role, Cluster, Status Aktif, No HP, atau koordinat Homebase banyak personel sekaligus di tabel ini, lalu klik Simpan.")
    cfg_mt = {
        "ROLE": st.column_config.SelectboxColumn("Role", options=MASTER_TIM_ROLES, default="FME / TE"),
        "STATUS_AKTIF": st.column_config.SelectboxColumn("Status Aktif", options=["Aktif", "Nonaktif"], default="Aktif")
    }
    if nop_k == "PLK":
        cfg_mt["CLUSTER"] = st.column_config.SelectboxColumn("Cluster", options=DAILY_CLUSTER_PLK_OPTS)
    ed_mt = st.data_editor(df_mt, use_container_width=True, num_rows="dynamic", column_config=cfg_mt, key=f"ed_mt_tbl_{nop_k}")

    cb_mt1, cb_mt2 = st.columns(2)
    with cb_mt1:
        if st.button(f"💾 SIMPAN PERUBAHAN MASSAL KE `{sh_target_mt}`", type="primary", use_container_width=True, key=f"btn_save_mt_all_{nop_k}"):
            df_clean_mt = ed_mt[ed_mt["NAMA"].astype(str).str.strip().str.lower().isin(["", "nan", "none"]) == False].copy()
            df_clean_mt["UPDATED_AT"] = datetime.datetime.now().strftime("%Y-%m-%d %H:%M")
            if gsheet_overwrite_ws(sh_target_mt, df_clean_mt[cols_target_mt]):
                st.success(f"✅ {len(df_clean_mt)} personel di `{sh_target_mt}` tersimpan & otomatis sinkron ke seluruh aplikasi!")
                st.rerun()
    with cb_mt2:
        buf_mt = BytesIO()
        with pd.ExcelWriter(buf_mt, engine="openpyxl") as xw:
            ed_mt.to_excel(xw, sheet_name=sh_target_mt, index=False)
        st.download_button(
            f"📥 Download Backup Master Tim {nop_k} (.xlsx)",
            data=buf_mt.getvalue(),
            file_name=f"{sh_target_mt}_{datetime.date.today()}.xlsx",
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
            use_container_width=True,
            key=f"dl_mt_xl_{nop_k}"
        )

    with st.expander(f"📤 Import Excel Master Tim {nop_k} (Replace / Append)", expanded=False):
        up_mt_xl = st.file_uploader(f"Upload Excel Daftar Tim {nop_k} (.xlsx / .csv):", type=["xlsx", "xls", "csv"], key=f"up_mt_xl_{nop_k}")
        mode_mt_up = st.radio("Mode Import:", ["Append (Tambah ke Tim Lama)", "Replace (Ganti Semua Tim)"], horizontal=True, key=f"mode_mt_up_{nop_k}")
        if up_mt_xl is not None:
            try:
                df_up_t = pd.read_csv(up_mt_xl) if up_mt_xl.name.lower().endswith(".csv") else read_excel_safe_raw(up_mt_xl, sheet_name=0, header=0)
                df_up_t = df_up_t.dropna(how="all").reset_index(drop=True)
                c_nm_u = next((c for c in df_up_t.columns if "nama" in str(c).lower()), df_up_t.columns[0])
                df_up_norm = pd.DataFrame()
                df_up_norm["NAMA"] = df_up_t[c_nm_u].astype(str).str.strip().str.upper()
                for col_t in cols_target_mt:
                    if col_t == "NAMA": continue
                    if col_t in df_up_t.columns:
                        df_up_norm[col_t] = df_up_t[col_t].fillna("").astype(str)
                    else:
                        df_up_norm[col_t] = "Aktif" if col_t == "STATUS_AKTIF" else ("FME / TE" if col_t == "ROLE" else ("NOP PANGKALAN BUN" if col_t == "NOP" else ""))
                df_up_norm = df_up_norm[~df_up_norm["NAMA"].isin(["", "NAN", "NONE"])].reset_index(drop=True)
                st.dataframe(df_up_norm.head(10), use_container_width=True, hide_index=True)
                if st.button(f"🚀 Simpan Hasil Import ({len(df_up_norm)} Tim) ke `{sh_target_mt}`", type="primary", use_container_width=True, key=f"btn_imp_mt_{nop_k}"):
                    if "Append" in mode_mt_up and not df_mt.empty:
                        df_comb = pd.concat([df_mt, df_up_norm], ignore_index=True).drop_duplicates(subset=["NAMA"], keep="last")
                    else:
                        df_comb = df_up_norm
                    if gsheet_overwrite_ws(sh_target_mt, df_comb[cols_target_mt]):
                        st.success(f"✅ {len(df_comb)} tim berhasil disimpan ke `{sh_target_mt}`!")
                        st.rerun()
            except Exception as e:
                st.error(f"Gagal memproses file Excel tim: {e}")

# FOOTER HAK CIPTA OKTA PRADIKA
# -------------------------------------------------------------------------
st.markdown("""
    <div class="footer-okta">
        System Application & Database Management<br>
        <span>Create Data By Okta Pradika</span> © 2026
    </div>
""", unsafe_allow_html=True)
