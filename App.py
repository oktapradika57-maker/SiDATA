import datetime
import pandas as pd
import streamlit as st
import cloudinary
import cloudinary.uploader
import json
import os
import urllib.parse

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
        except Exception:
            return None
    return None

def upload_multiple_images(file_objs, folder_name="solar_bts_healthcheck"):
    urls = []
    if file_objs:
        for file in file_objs:
            url = upload_image(file, folder_name)
            if url: urls.append(url)
    return urls

# Fungsi Anti-Gagal untuk Menampilkan Foto
def tampilkan_foto(url_data, caption=""):
    if not url_data: return
    if isinstance(url_data, str) and url_data.startswith("http"):
        st.image(url_data, caption=caption, width=200)
    elif isinstance(url_data, list):
        valid_urls = [u for u in url_data if isinstance(u, str) and u.startswith("http")]
        if valid_urls:
            cols = st.columns(len(valid_urls) if len(valid_urls) < 4 else 4)
            for idx, u in enumerate(valid_urls):
                cols[idx%4].image(u, caption=caption, width=200)

# -------------------------------------------------------------------------
# 2. SETUP DATABASE
# -------------------------------------------------------------------------
st.set_page_config(page_title="Solar BTS Health Check", page_icon="⚡", layout="wide")
DB_FILE = "laporan_db.json"

def get_db():
    if os.path.exists(DB_FILE):
        try:
            with open(DB_FILE, "r") as f: return json.load(f)
        except: return []
    return []

def save_db(data):
    try:
        with open(DB_FILE, "w") as f: json.dump(data, f)
    except: pass

if 'laporan_db' not in st.session_state:
    st.session_state['laporan_db'] = get_db()

# -------------------------------------------------------------------------
# 3. NAVIGASI
# -------------------------------------------------------------------------
st.sidebar.title("Navigasi Aplikasi")
menu = st.sidebar.radio("Pilih Menu:", ["📝 Form Pengecekan", "📊 Hasil Pengecekan (Laporan)"])
st.sidebar.markdown("---")

# =========================================================================
# MENU 1: FORM PENGECEKAN
# =========================================================================
if menu == "📝 Form Pengecekan":
    st.title("⚡ Form Health Check Solar BTS")
    
    tabs = st.tabs(["1. Info Site", "2. Panel", "3. Panel DC", "4. SCC", "5. Baterai", "6. Grounding", "7. Submit"])

    with tabs[0]:
        c1, c2 = st.columns(2)
        with c1:
            site_name = st.text_input("Nama Site", placeholder="Contoh: BTS-PKY-001")
            nop_area = st.selectbox("NOP Area", ["Palangkaraya", "Pangkalan Bun", "Tarakan", "Pontianak", "Lainnya"])
        with c2:
            technician_name = st.text_input("Nama Teknisi")
            weather = st.selectbox("Cuaca", ["Cerah", "Berawan", "Hujan", "Badai"])
        st.divider()
        sc1, sc2 = st.columns(2)
        with sc1:
            site_condition = st.selectbox("Halaman & Pagar", ["Aman", "Rumput Liar", "Rusak"])
            tower_condition = st.selectbox("Fisik Tower", ["Aman", "Berkarat / Kendor"])
        with sc2:
            site_photos = st.file_uploader("Foto View Site (Bisa >1)", accept_multiple_files=True, key="spics")

    with tabs[1]:
        shading_status = st.selectbox("Status Shading?", ["Aman", "Sedikit Shading", "Shading Kritis"])
        shading_photos = st.file_uploader("Foto Bukti Shading (Bisa >1)", accept_multiple_files=True, key="shd")
        st.divider()
        
        konfigurasi_panel = st.radio("Metode Pengukuran Panel:", ["Individu (Satuan)", "Seri / String (Grup)"])
        panel_data = []

        if konfigurasi_panel == "Individu (Satuan)":
            num_panels = st.number_input("Jumlah Panel", min_value=1, value=24)
            for i in range(int(num_panels)):
                with st.expander(f"Panel #{i+1}", expanded=(i==0)):
                    c1, c2, c3 = st.columns([1, 1, 1])
                    with c1:
                        voc = st.number_input(f"Voc [V]", value=21.5, key=f"v_{i}")
                        isc = st.number_input(f"Isc [A]", value=5.2, key=f"i_{i}")
                        p_cond = st.selectbox(f"Fisik", ["Baik", "Kotor", "Retak", "Delaminasi"], key=f"c_{i}")
                    with c2: pb = st.file_uploader(f"📸 BEFORE", key=f"pb_{i}")
                    with c3: pa = st.file_uploader(f"✨ AFTER", key=f"pa_{i}")
                    panel_data.append({"tipe": "Individu", "id": f"Panel #{i+1}", "voc": voc, "isc": isc, "kondisi": p_cond, "foto_before_obj": pb, "foto_after_obj": pa})
        else: 
            num_strings = st.number_input("Jumlah String", min_value=1, value=4)
            for i in range(int(num_strings)):
                with st.expander(f"String #{i+1}", expanded=(i==0)):
                    c1, c2 = st.columns([1, 1.5])
                    with c1:
                        qty = st.number_input(f"Isi Panel", value=6, key=f"sq_{i}")
                        voc = st.number_input(f"Total Voc [V]", value=129.0, key=f"sv_{i}")
                        isc = st.number_input(f"Isc [A]", value=5.2, key=f"si_{i}")
                        p_cond = st.selectbox(f"Kondisi", ["Baik", "Kotor", "Retak"], key=f"sc_{i}")
                    with c2:
                        pb = st.file_uploader(f"📸 BEFORE (Bisa >1)", accept_multiple_files=True, key=f"spb_{i}")
                        pa = st.file_uploader(f"✨ AFTER (Bisa >1)", accept_multiple_files=True, key=f"spa_{i}")
                    panel_data.append({"tipe": "Seri", "id": f"String #{i+1}", "qty": qty, "voc": voc, "isc": isc, "kondisi": p_cond, "foto_before_objs": pb, "foto_after_objs": pa})

    with tabs[2]:
        c1, c2 = st.columns(2)
        with c1:
            jb_enclosure = st.selectbox("Box DC", ["Baik", "Bocor", "Sarang Serangga"])
            jb_breaker = st.selectbox("Breaker", ["Normal", "Trip/Putus"])
            jb_spd = st.selectbox("Arrester", ["Normal", "Rusak"])
            cabling = st.selectbox("Kabel", ["Rapi", "Chattering", "Terkelupas"])
        with c2: jb_photos = st.file_uploader("Foto Panel DC (Bisa >1)", accept_multiple_files=True, key="jb")

    with tabs[3]:
        c1, c2 = st.columns(2)
        with c1:
            scc_brand = st.text_input("Merek SCC", placeholder="Huawei")
            total_load_a = st.number_input("Total Load (A)", value=15.5)
            inv_status = st.selectbox("Inverter", ["Normal", "Error", "Tidak Pakai"])
            is_datalog_taken = st.checkbox("Datalog ditarik")
        with c2:
            scc_alarm = st.selectbox("Alarm SCC", ["Normal", "Ada Alarm"])
            scc_photos = st.file_uploader("Foto SCC (Bisa >1)", accept_multiple_files=True, key="scc")

    with tabs[4]:
        num_bat = st.number_input("Jumlah Baterai", min_value=1, value=4)
        bat_data = []
        for j in range(int(num_bat)):
            with st.expander(f"Baterai #{j+1}", expanded=(j==0)):
                c1, c2, c3 = st.columns([1, 1, 1.2])
                with c1:
                    bv = st.number_input(f"Voltase [V]", value=12.2, key=f"bv_{j}")
                    bt = st.number_input(f"Suhu [°C]", value=28.0, key=f"bt_{j}")
                with c2: bc = st.selectbox(f"Fisik", ["Normal", "Bengkak", "Korosi"], key=f"bc_{j}")
                with c3: bp = st.file_uploader(f"Foto Baterai", accept_multiple_files=True, key=f"bp_{j}")
                bat_data.append({"id": f"Baterai #{j+1}", "voltase": bv, "suhu": bt, "kondisi": bc, "foto_objs": bp})

    with tabs[5]:
        c1, c2 = st.columns(2)
        with c1:
            earth_resistance = st.number_input("Tahanan Grounding [Ohm]", value=2.1)
            grd_cable = st.selectbox("Kabel Grounding", ["Kuat", "Kendor", "Putus"])
        with c2: grd_photos = st.file_uploader("Foto Grounding (Bisa >1)", accept_multiple_files=True, key="grd")

    with tabs[6]:
        action_taken = st.text_area("Tindakan Action:")
        sparepart_needed = st.text_input("Daftar Sparepart:")
        final_status = st.radio("Status Akhir:", ["Normal", "Minor Issue", "Major/Critical"])

        if st.button("🚀 Upload & Submit", type="primary"):
            if not site_name: 
                st.error("⚠️ Mohon isi Nama Site terlebih dahulu di Tab 1!")
            else:
                with st.spinner("⏳ Menyimpan data & mengunggah foto... Mohon tunggu."):
                    p_res = []
                    for p in panel_data:
                        if p.get("tipe") == "Individu":
                            p_res.append({"Tipe": "Individu", "Panel": p["id"], "Voc": p["voc"], "Isc": p["isc"], "Kondisi": p["kondisi"], 
                                          "URL_Before": upload_image(p["foto_before_obj"]), "URL_After": upload_image(p["foto_after_obj"])})
                        else:
                            p_res.append({"Tipe": "Seri", "Panel": p["id"], "Qty": p["qty"], "Voc": p["voc"], "Isc": p["isc"], "Kondisi": p["kondisi"], 
                                          "URLs_Before": upload_multiple_images(p["foto_before_objs"]), "URLs_After": upload_multiple_images(p["foto_after_objs"])})
                    
                    b_res = []
                    for b in bat_data:
                        b_res.append({"Baterai": b["id"], "Voltase": b["voltase"], "Suhu": b["suhu"], "Kondisi": b["kondisi"], "URL_Fotos": upload_multiple_images(b["foto_objs"])})

                    st.session_state['laporan_db'].append({
                        "timestamp": datetime.datetime.now().strftime("%Y-%m-%d %H:%M:%S"),
                        "site_name": site_name, "nop": nop_area, "teknisi": technician_name,
                        "status": final_status, "action": action_taken, "sparepart": sparepart_needed,
                        "site_cond": site_condition, "tower_cond": tower_condition, "url_sites": upload_multiple_images(site_photos),
                        "shading_status": shading_status, "url_shadings": upload_multiple_images(shading_photos),
                        "jb_enclosure": jb_enclosure, "jb_breaker": jb_breaker, "jb_spd": jb_spd, "cabling": cabling, "url_jbs": upload_multiple_images(jb_photos),
                        "scc_brand": scc_brand, "total_load": total_load_a, "url_sccs": upload_multiple_images(scc_photos),
                        "earth_ohm": earth_resistance, "grd_cable": grd_cable, "url_grds": upload_multiple_images(grd_photos),
                        "panel_data": p_res, "battery_data": b_res,
                        "extras_fisik": [], "extras_panel": [], "extras_baterai": [], "extras_elektrikal": []
                    })
                    save_db(st.session_state['laporan_db'])            
                st.success("✅ Tersimpan! Silakan cek Menu Hasil Pengecekan.")

# =========================================================================
# MENU 2: HASIL PENGECEKAN (TAMPILAN MUNCUL & BISA NAMBAH FOTO BEBAS)
# =========================================================================
elif menu == "📊 Hasil Pengecekan (Laporan)":
    st.title("📊 Laporan Komprehensif")
    st.divider()

    db = st.session_state['laporan_db']

    if not db:
        st.info("⚠️ Data laporan saat ini kosong karena sistem baru di-restart. Silakan buat 1 laporan tes di menu Form Pengecekan.")
    else:
        for i in range(len(db) - 1, -1, -1):
            r = db[i]
            
            wa_text = f"""*REPORT SOLAR BTS* ⚡\n📍 *Site:* {r.get('site_name', '-')} ({r.get('nop', '-')})\n📅 *Tanggal:* {r.get('timestamp', '-')}\n👷 *Teknisi:* {r.get('teknisi', '-')}\n📊 *Status:* {r.get('status', '-')}\n\n*ACTION:*\n{r.get('action', '-')}\n\n*SPAREPART:*\n{r.get('sparepart', '-')}\n\n*Load:* {r.get('total_load', '-')} A | *Grounding:* {r.get('earth_ohm', '-')} Ohm"""
            wa_url = f"https://wa.me/?text={urllib.parse.quote(wa_text)}"

            with st.expander(f"📍 {r.get('site_name', 'Unknown')} | {r.get('timestamp', '')} | Status: {r.get('status', '')}"):
                
                st.link_button("📱 Generate Rangkuman ke WA", wa_url)
                
                # --- TAMPILAN UTAMA ---
                st.markdown(f"**Teknisi:** {r.get('teknisi', '-')} | **Load:** {r.get('total_load', '-')} A")
                st.markdown(f"**Action:** {r.get('action', '-')}")
                if r.get('sparepart'): st.warning(f"**Sparepart:** {r['sparepart']}")
                
                ltab1, ltab2, ltab3, ltab4 = st.tabs(["1. Fisik & Umum", "2. Panel Surya", "3. Baterai", "4. Elektrikal"])
                
                with ltab1:
                    st.write(f"- Kondisi: {r.get('site_cond', '-')} | Tower: {r.get('tower_cond', '-')} | Shading: {r.get('shading_status', '-')}")
                    tampilkan_foto(r.get('url_sites'), "View Site")
                    tampilkan_foto(r.get('url_shadings'), "Shading")
                    tampilkan_foto(r.get('extras_fisik'), "Susulan Fisik")

                with ltab2:
                    for p in r.get('panel_data', []):
                        st.write(f"**{p.get('Panel', '-')}** | Voc: {p.get('Voc','-')}V | Isc: {p.get('Isc','-')}A | Fisik: {p.get('Kondisi','-')}")
                        cb, ca = st.columns(2)
                        with cb: 
                            tampilkan_foto(p.get('URL_Before'), "Before")
                            tampilkan_foto(p.get('URLs_Before'), "Before")
                        with ca: 
                            tampilkan_foto(p.get('URL_After'), "After")
                            tampilkan_foto(p.get('URLs_After'), "After")
                        st.divider()
                    tampilkan_foto(r.get('extras_panel'), "Susulan Panel (Bebas)")

                with ltab3:
                    for b in r.get('battery_data', []):
                        st.write(f"**{b.get('Baterai', '-')}** | {b.get('Voltase','-')}V | Suhu: {b.get('Suhu','-')}°C | Fisik: {b.get('Kondisi','-')}")
                        tampilkan_foto(b.get('URL_Fotos'))
                        st.divider()
                    tampilkan_foto(r.get('extras_baterai'), "Susulan Baterai (Bebas)")

                with ltab4:
                    c1, c2 = st.columns(2)
                    with c1: st.write(f"**JB Box:** {r.get('jb_enclosure', '-')} | **Breaker:** {r.get('jb_breaker', '-')}")
                    with c2: st.write(f"**SCC:** {r.get('scc_brand', '-')} | **Grounding:** {r.get('earth_ohm', '-')} Ohm")
                    tampilkan_foto(r.get('url_jbs'), "Panel DC")
                    tampilkan_foto(r.get('url_sccs'), "SCC")
                    tampilkan_foto(r.get('url_grds'), "Grounding")
                    tampilkan_foto(r.get('extras_elektrikal'), "Susulan Elektrikal")

                st.divider()
                # --- FITUR EDIT & UPLOAD TAMBAHAN BEBAS ---
                st.markdown("### 🛠️ EDIT TEKS & TAMBAH FOTO BEBAS KAPAN SAJA")
                with st.container(border=True):
                    new_tek = st.text_input("Perbaiki Teknisi", r.get('teknisi',''), key=f"et_{i}")
                    new_act = st.text_area("Perbaiki Action", r.get('action',''), key=f"ea_{i}")
                    new_sp = st.text_input("Perbaiki Sparepart", r.get('sparepart',''), key=f"es_{i}")
                    
                    st.markdown("**Tambah Foto Susulan (Bebas Pilih Kategori):**")
                    up_f = st.file_uploader("Fisik / Shading", accept_multiple_files=True, key=f"uf_{i}")
                    up_p = st.file_uploader("Panel Surya", accept_multiple_files=True, key=f"up_{i}")
                    up_b = st.file_uploader("Baterai", accept_multiple_files=True, key=f"ub_{i}")
                    up_e = st.file_uploader("Elektrikal (SCC/JB/GND)", accept_multiple_files=True, key=f"ue_{i}")

                    if st.button("💾 Simpan Edit & Upload Foto Tambahan", key=f"btn_{i}", type="primary"):
                        with st.spinner("Mengunggah foto dan menyimpan data..."):
                            r['teknisi'], r['action'], r['sparepart'] = new_tek, new_act, new_sp
                            if up_f: r['extras_fisik'] = r.get('extras_fisik', []) + upload_multiple_images(up_f)
                            if up_p: r['extras_panel'] = r.get('extras_panel', []) + upload_multiple_images(up_p)
                            if up_b: r['extras_baterai'] = r.get('extras_baterai', []) + upload_multiple_images(up_b)
                            if up_e: r['extras_elektrikal'] = r.get('extras_elektrikal', []) + upload_multiple_images(up_e)
                            save_db(db)
                        st.success("✅ Tersimpan! Layar akan refresh.")
                        st.rerun()
