# 🚀 MERGE FULL — SiDATA + SiSAVA + SiHD (1 File Siap Copas)

## ⚠️ Penting Sebelum Mulai

Saya **TIDAK BISA** generate seluruh file `App (70).py` utuh dalam 1 response karena:
- File Anda ±5.000+ baris (batas max response saya 5.000 karakter)
- Kalau saya potong-potong, malah **lebih rawan error** daripada patch bertahap

## ✅ Solusi Aman: **1 File "Overlay"** (Sistem Drop-In)

Kita buat **1 file tambahan** (`sihd_addon.py`) yang di-**import** dari `App (70).py` — hanya **3 baris** yang perlu Anda tambahkan di file utama. Semua logika baru (SiSAVA, SiHD, WA, Dashboard) ada di file addon terpisah.

---

## 📄 FILE 1: `sihd_addon.py` (Buat File Baru)

Buat file baru di folder yang **sama** dengan `App (70).py`, beri nama **`sihd_addon.py`**, isi **lengkap**:

```python
# ============================================================================
# sihd_addon.py — PAKET ADDON SiHD + SiSAVA (Drop-In, Tidak Ubah App Utama)
# ============================================================================
import datetime as _dtm
import urllib.parse
import streamlit as st

# ---------------- KONSTANTA ----------------
SIHD_MENU = "🔄 SiHD — Shift Handover"
SIHD_STATUS = ["🟢 Standby", "🔵 On The Way", "🟡 On Site", "🟠 On Progress", "✅ Done", "🔴 Kendala", "⚫ Off"]
SIHD_SHIFTS = ["Shift 1 (Pagi)", "Shift 2 (Siang)", "Shift 3 (Malam)", "Lembur"]
SIHD_AKTIVITAS = [
    "Preventive Maintenance", "Troubleshooting BTS", "Perbaikan Genset", "Pengecekan Rectifier",
    "Cleaning Panel Surya", "Pengecekan Baterai", "Grounding Check", "Pembersihan Shelter",
    "Perbaikan Kabel Power", "Relokasi Perangkat", "Site Survey", "Monitoring Site",
    "Perbaikan AC", "Instalasi Perangkat", "Perbaikan Tower", "Lainnya"
]

def init_state():
    for _k, _v in [("sihd_log", []), ("sihd_current_shift", None), ("sihd_current_team", None)]:
        if _k not in st.session_state:
            st.session_state[_k] = _v

def wa_move(e):
    t = (f"📍 *LAPORAN PERGERAKAN TIM*\n━━━━━━━━━━━━━━━━━━━━\n"
         f"📅 {e.get('tanggal')} {e.get('jam')}\n⏰ {e.get('shift')}\n👥 {e.get('team')}\n"
         f"🏢 {e.get('nop')}\n🚦 {e.get('status')}\n━━━━━━━━━━━━━━━━━━━━\n"
         f"📍 Dari : {e.get('dari')}\n🎯 Ke   : {e.get('ke')}\n"
         f"🔧 Akt  : {e.get('aktivitas')}\n⚡ Prio : {e.get('prioritas')}\n")
    if e.get("ket"): t += f"📝 {e['ket']}\n"
    t += "━━━━━━━━━━━━━━━━━━━━\n_Live Update — SiHD_"
    return t

def wa_handover(ds, ks, tl, tb, tgl, ring, ken, sopen, sdone, reko):
    ts = tgl.strftime("%d %b %Y") if hasattr(tgl, "strftime") else str(tgl)
    t = (f"🔄 *SERAH TERIMA SHIFT*\n━━━━━━━━━━━━━━━━━━━━\n"
         f"📅 {ts}\n🕐 {ds} ➡️ {ks}\n👤 Tim Lama: {tl}\n👤 Tim Baru: {tb}\n"
         f"━━━━━━━━━━━━━━━━━━━━\n📋 Ringkasan:\n{ring}\n")
    if ken and ken != "-": t += f"━━━━━━━━━━━━━━━━━━━━\n⚠️ Kendala:\n{ken}\n"
    if sopen and sopen != "-": t += f"🔴 Site Belum Selesai:\n{sopen}\n"
    if sdone and sdone != "-": t += f"✅ Site Selesai:\n{sdone}\n"
    if reko and reko != "-": t += f"💡 Rekomendasi:\n{reko}\n"
    t += "━━━━━━━━━━━━━━━━━━━━\n🙏 _Mohon konfirmasi penerimaan shift._"
    return t

def render_sihd():
    init_state()
    st.markdown("<h1 style='text-align:center;'>🔄 SiHD — Shift Handover</h1>", unsafe_allow_html=True)
    st.markdown("<p style='text-align:center; opacity:.8;'>Pergerakan tim antar-site & serah terima shift dengan WA otomatis</p>", unsafe_allow_html=True)
    t1, t2, t3 = st.tabs(["📋 Input Pergerakan", "🔄 Serah Terima Shift", "📊 Log Aktivitas"])

    with t1:
        st.markdown("### 📍 Catat Pergerakan Tim")
        with st.container(border=True):
            c1, c2, c3 = st.columns(3)
            with c1:
                tgl = st.date_input("📅 Tanggal", value=_dtm.date.today(), key="sd_tgl")
                shf = st.selectbox("⏰ Shift", SIHD_SHIFTS, key="sd_shf")
            with c2:
                tim = st.text_input("👥 Nama Tim", placeholder="Cth: Tim Alpha - Okta", key="sd_tim")
                sta = st.selectbox("🚦 Status", SIHD_STATUS, key="sd_sta")
            with c3:
                jam = st.time_input("🕐 Jam", value=_dtm.datetime.now().time(), key="sd_jam")
                nop = st.selectbox("🏢 NOP", ["Palangkaraya", "Pangkalan Bun", "Tarakan", "Pontianak", "Lainnya"], key="sd_nop")
        with st.container(border=True):
            p1, p2 = st.columns(2)
            with p1:
                dari = st.text_input("📍 Site Asal", key="sd_dari")
                akt = st.selectbox("🔧 Aktivitas", SIHD_AKTIVITAS, key="sd_akt")
            with p2:
                ke = st.text_input("🎯 Site Tujuan", key="sd_ke")
                prio = st.selectbox("⚡ Prioritas", ["Normal", "High", "Critical"], key="sd_prio")
            ket = st.text_area("📝 Keterangan", key="sd_ket", height=80)
        if st.button("💾 SIMPAN PERGERAKAN", type="primary", use_container_width=True, key="sd_save"):
            if not tim or not dari or not ke:
                st.error("⚠️ Tim, Site Asal, Site Tujuan wajib diisi!")
            else:
           

*(Jawaban telah mencapai batas maksimum 5.000 karakter. Ketik **lanjutkan** untuk meneruskan pembahasan.)*
