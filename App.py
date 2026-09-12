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
        except Exception as e:
            st.error(f"Gagal upload gambar: {e}")
            return None
    return None

def upload_multiple_images(file_objs, folder_name="solar_bts_healthcheck"):
    urls = []
    if file_objs:
        for file in file_objs:
            url = upload_image(file, folder_name)
            if url: urls.append(url)
    return urls

# FUNGSI EXTRAKTOR FOTO UNIVERSAL (Mencegah foto lama hilang)
def get_safe_urls(data_dict, keys):
    urls = []
    for k in keys:
        val = data_dict.get(k)
        if val:
            if isinstance(val, str) and val.startswith("http"):
                urls.append(val)
            elif isinstance(val, list):
                urls.extend([u for u in val if isinstance(u, str) and u.startswith("http")])
    return urls

# -------------------------------------------------------------------------
# 2. SETUP HYBRID DATABASE
# -------------------------------------------------------------------------
st.set_page_config(page_title="Solar BTS Health Check Pro", page_icon="⚡", layout="wide")

DB_FILE = "laporan_db.json"

def get_db():
    if os.path.exists(DB_FILE):
        try:
            with open(DB_FILE, "r") as f:
                return json.load(f)
        except:
            return []
    return []

def save_db(data):
    try:
        with open(DB_FILE, "w") as f:
            json.dump(data, f)
    except Exception as e:
        st.error(f"Gagal menyimpan database: {e}")

if 'laporan_db' not in st.session_state:
    st.session_state['laporan_db'] = get_db()

# -------------------------------------------------------------------------
# 3. NAVIGASI SIDEBAR
# -------------------------------------------------------------------------
st.sidebar.title("Navigasi Aplikasi")
menu = st.sidebar.radio("Pilih Menu:", ["📝 Form Pengecekan", "📊 Hasil Pengecekan (Laporan)"])
st.sidebar.markdown("---")

# =========================================================================
# MENU 1: FORM PENGECEKAN
# =========================================================================
if menu == "📝 Form Pengecekan":
    st.title("⚡ Form Health Check Solar BTS")
    st.info("💡 Pilih Konfigurasi Panel (Satuan vs Seri). Data otomatis tersimpan permanen.")
    
    tabs = st.tabs(["1. Info & View Site", "2. Solar Panel", "3. Panel DC & Kabel", "4. SCC & Inverter", "5. Baterai", "6. Grounding", "7. Action & Submit"])

    with tabs[0]:
        st.subheader("Informasi Umum & Kondisi Fisik Site")
        col1, col2 = st.columns(2)
        with col1:
            site_name = st.text_input("Nama / ID Site", placeholder="Contoh: BTS-PKY-001")
            nop_area = st.selectbox("NOP Area", ["Palangkaraya", "Pangkalan Bun", "Tarakan", "Pontianak", "Lainnya"])
            check_date = st.date_input("Tanggal Pengecekan", value=datetime.date.today())
        with col2:
            technician_name = st.text_input("Nama Teknisi", placeholder="Nama lengkap tim")
            weather = st.selectbox("Kondisi Cuaca", ["Cerah", "Berawan", "Hujan", "Badai/Ekstrem"])
        st.divider()
        site_col1, site_col2 = st.columns(2)
        with site_col1:
            site_condition = st.selectbox("Kondisi Halaman & Pagar", ["Bersih & Aman", "Banyak Rumput Liar / Belukar", "Pagar Rusak / Gembok Hilang"])
            tower_condition = st.selectbox("Kondisi Fisik Tower / Tiang", ["Aman Tidak Berkarat", "Berkarat / Baut Kendur"])
        with site_col2:
            site_photos = st.file_uploader("Upload Foto View Site (Bisa >1 Foto)", type=["jpg", "png", "jpeg"], accept_multiple_files=True, key="spics")

    with tabs[1]:
        shade_col1, shade_col2 = st.columns(2)
        with shade_col1:
            shading_status = st.selectbox("Status Shading?", ["Aman (Clear area)", "Sedikit Shading (Bayangan pohon)", "Shading Kritis (Tertutup rimbunan)"])
        with shade_col2:
            shading_photos = st.file_uploader("Upload Foto Bukti Shading (Bisa >1 Foto)", type=["jpg", "png", "jpeg"], accept_multiple_files=True, key="shd")
        st.divider()
        
        konfigurasi_panel = st.radio("Metode Pengukuran Panel:", ["Individu (Satuan)", "Seri / String (Grup/Seri)"])
        panel_data = []

        if konfigurasi_panel == "Individu (Satuan)":
            num_panels = st.number_input("Jumlah Total Panel", min_value=1, max_value=60, value=24)
            for i in range(int(num_panels)):
                with st.expander(f"Panel #{i+1}", expanded=(i == 0)):
                    p_col1, p_col2, p_col3 = st.columns([1, 1.2, 1.2])
                    with p_col1:
                        voc = st.number_input(f"Voc [V] - P{i+1}", value=21.5, step=0.1, key=f"v_{i}")
                        isc = st.number_input(f"Isc [A] - P{i+1}", value=5.2, step=0.1, key=f"i_{i}")
                        p_cond = st.selectbox(f"Fisik P{i+1}", ["Baik & Mulus", "Sangat Kotor", "Kaca Retak", "Delaminasi"], key=f"c_{i}")
                    with p_col2:
                        p_photo_before = st.file_uploader(f"📸 BEFORE P{i+1}", type=["jpg", "png", "jpeg"], key=f"pb_{i}")
                    with p_col3:
                        p_photo_after = st.file_uploader(f"✨ AFTER P{i+1}", type=["jpg", "png", "jpeg"], key=f"pa_{i}")
                    panel_data.append({"tipe": "Individu", "id": f"Panel #{i+1}", "voc": voc, "isc": isc, "kondisi": p_cond, "foto_before_obj": p_photo_before, "foto_after_obj": p_photo_after})
        else: 
            num_strings = st.number_input("Jumlah String / Seri", min_value=1, max_value=20, value=4)
            for i in range(int(num_strings)):
                with st.expander(f"String / Seri #{i+1}", expanded=(i == 0)):
                    p_col1, p_col2 = st.columns([1, 1.5])
                    with p_col1:
                        qty = st.number_input(f"Isi Panel Seri {i+1}", value=6, key=f"sq_{i}")
                        voc = st.number_input(f"Total Voc Seri {i+1} [V]", value=129.0, step=0.1, key=f"sv_{i}")
                        isc = st.number_input(f"Isc Seri {i+1} [A]", value=5.2, step=0.1, key=f"si_{i}")
                        p_cond = st.selectbox(f"Kondisi Mayoritas", ["Baik & Mulus", "Banyak Kotoran", "Ada Retak", "Delaminasi"], key=f"sc_{i}")
                    with p_col2:
                        p_photo_before = st.file_uploader(f"📸 BEFORE Seri {i+1} (Bisa >1)", type=["jpg", "png", "jpeg"], accept_multiple_files=True, key=f"spb_{i}")
                        p_photo_after = st.file_uploader(f"✨ AFTER Seri {i+1} (Bisa >1)", type=["jpg", "png", "jpeg"], accept_multiple_files=True, key=f"spa_{i}")
                    panel_data.append({"tipe": "Seri", "id": f"String #{i+1}", "qty": qty, "voc": voc, "isc": isc, "kondisi": p_cond, "foto_before_objs": p_photo_before, "foto_after_objs": p_photo_after})

    with tabs[2]:
        jb_col1, jb_col2 = st.columns(2)
        with jb_col1:
            jb_enclosure = st.selectbox("Kondisi Box DC", ["Bersih & Kedap Air", "Bocor / Kemasukan Air", "Sarang Serangga"])
            jb_breaker = st.selectbox("Status Breaker", ["Normal (On Semua)", "Trip / Putus / Gosong"])
            jb_spd = st.selectbox("Arrester / SPD", ["Normal (Hijau)", "Rusak (Merah/Hitam)"])
            cabling = st.selectbox("Kondisi Kabel & Kontaktor", ["Rapi & Kuat", "Chattering", "Terkelupas / Oksidasi"])
        with jb_col2:
            jb_photos = st.file_uploader("Upload Foto Panel DC (Bisa >1 Foto)", type=["jpg", "png", "jpeg"], accept_multiple_files=True, key="jb")

    with tabs[3]:
        scc_col1, scc_col2 = st.columns(2)
        with scc_col1:
            scc_brand = st.text_input("Merek SCC/Recty", placeholder="SmartGen, Huawei")
            system_out_v = st.number_input("Tegangan Output ke BTS [V]", value=48.0)
            total_load_a = st.number_input("Total Beban (Ampere)", value=15.5, step=0.1)
            inv_status = st.selectbox("Kondisi Inverter", ["Normal / 220V AC", "Error / Fault", "Tidak Pakai"])
            is_datalog_taken = st.checkbox("✅ Datalog berhasil ditarik")
        with scc_col2:
            scc_alarm = st.selectbox("Status Layar SCC", ["Normal (No Alarm)", "Ada Alarm Fault"])
            scc_photos = st.file_uploader("Upload Foto SCC/Inverter (Bisa >1 Foto)", type=["jpg", "png", "jpeg"], accept_multiple_files=True, key="scc")

    with tabs[4]:
        bat_total_col1, bat_total_col2 = st.columns(2)
        with bat_total_col1:
            num_batteries = st.number_input("Jumlah Blok Baterai (Maks 12)", min_value=1, max_value=12, value=4)
        with bat_total_col2:
            bat_total_v = st.number_input("Total Tegangan Bank Baterai [Volt]", value=52.8)
        
        bat_data = []
        for j in range(int(num_batteries)):
            with st.expander(f"Baterai #{j+1}", expanded=(j==0)):
                b_col1, b_col2, b_col3 = st.columns([1, 1, 1.2])
                with b_col1:
                    b_volt = st.number_input(f"Voltase Bat {j+1} [V]", value=12.2, step=0.1, key=f"bv_{j}")
                    b_temp = st.number_input(f"Suhu Bat {j+1} [°C]", value=28.0, step=0.1, key=f"bt_{j}")
                with b_col2:
                    b_cond = st.selectbox(f"Fisik Baterai {j+1}", ["Normal", "Bengkak", "Korosi"], key=f"bc_{j}")
                with b_col3:
                    b_photo = st.file_uploader(f"Upload Foto Baterai {j+1}", type=["jpg", "png", "jpeg"], accept_multiple_files=True, key=f"bp_{j}")
                bat_data.append({"id": f"Baterai #{j+1}", "voltase": b_volt, "suhu": b_temp, "kondisi": b_cond, "foto_objs": b_photo})

    with tabs[5]:
        grd_col1, grd_col2 = st.columns(2)
        with grd_col1:
            earth_resistance = st.number_input("Tahanan Grounding [Ohm]", value=2.1, step=0.1)
            grd_cable = st.selectbox("Kondisi Busbar & Kabel", ["Terhubung Kuat", "Kendor / Karat", "Putus / Dicuri"])
        with grd_col2:
            grd_photos = st.file_uploader("Upload Foto Grounding (Bisa >1 Foto)", type=["jpg", "png", "jpeg"], accept_multiple_files=True, key="grd")

    with tabs[6]:
        action_taken = st.text_area("🔧 Ketik Action / Eksekusi:", height=120)
        sparepart_needed = st.text_input("📦 Daftar Sparepart Dibutuhkan/Diganti:")
        final_status = st.radio("Status Akhir Site:", ["Normal (Aman)", "Minor Issue (Sudah Diperbaiki)", "Major / Critical (Butuh Part)"])

        if st.button("🚀 Upload & Submit Laporan", type="primary"):
            if not site_name or not technician_name:
                st.error("⚠️ Mohon isi Nama Site dan Nama Teknisi di Tab 1!")
            else:
                with st.spinner("⏳ Mengunggah data & foto... Mohon tunggu."):
                    url_sites = upload_multiple_images(site_photos, "solar_view")
                    url_shadings = upload_multiple_images(shading_photos, "solar_shade")
                    url_jbs = upload_multiple_images(jb_photos, "solar_jb")
                    url_sccs = upload_multiple_images(scc_photos, "solar_scc")
                    url_grds = upload_multiple_images(grd_photos, "solar_grd")

                    panel_results = []
                    for p in panel_data:
                        if p.get("tipe") == "Individu":
                            u_bef = upload_image(p["foto_before_obj"], "solar_panel") if p["foto_before_obj"] else None
                            u_aft = upload_image(p["foto_after_obj"], "solar_panel") if p["foto_after_obj"] else None
                            panel_results.append({"Tipe": "Individu", "Panel": p["id"], "Voc": p["voc"], "Isc": p["isc"], "Kondisi": p["kondisi"], "URL_Before": u_bef, "URL_After": u_aft})
                        else:
                            u_bef = upload_multiple_images(p["foto_before_objs"], "solar_panel")
                            u_aft = upload_multiple_images(p["foto_after_objs"], "solar_panel")
                            panel_results.append({"Tipe": "Seri", "Panel": p["id"], "Qty": p["qty"], "Voc": p["voc"], "Isc": p["isc"], "Kondisi": p["kondisi"], "URLs_Before": u_bef, "URLs_After": u_aft})
                    
                    bat_results = []
                    for b in bat_data:
                        u_bat = upload_multiple_images(b["foto_objs"], "solar_bat")
                        bat_results.append({"Baterai": b["id"], "Voltase": b["voltase"], "Suhu": b["suhu"], "Kondisi": b["kondisi"], "URL_Fotos": u_bat})

                    report_dict = {
                        "timestamp": datetime.datetime.now().strftime("%Y-%m-%d %H:%M:%S"),
                        "site_name": site_name, "nop": nop_area, "teknisi": technician_name,
                        "status": final_status, "action": action_taken, "sparepart": sparepart_needed,
                        "site_cond": site_condition, "tower_cond": tower_condition, "url_sites": url_sites,
                        "shading_status": shading_status, "url_shadings": url_shadings,
                        "jb_enclosure": jb_enclosure, "jb_breaker": jb_breaker, "jb_spd": jb_spd, "cabling": cabling, "url_jbs": url_jbs,
                        "scc_brand": scc_brand, "scc_alarm": scc_alarm, "total_load": total_load_a, "inv_status": inv_status,
                        "datalog_status": "Ditarik" if is_datalog_taken else "Tidak Ditarik", "url_sccs": url_sccs,
                        "earth_ohm": earth_resistance, "grd_cable": grd_cable, "url_grds": url_grds,
                        "panel_data": panel_results, "battery_data": bat_results,
                        "extras_fisik": [], "extras_panel": [], "extras_baterai": [], "extras_elektrikal": []
                    }
                    st.session_state['laporan_db'].append(report_dict) 
                    save_db(st.session_state['laporan_db'])            
                st.success(f"✅ Laporan Site {site_name} disimpan PERMANEN!")

# =========================================================================
# MENU 2: HASIL PENGECEKAN (DASHBOARD & FULL EDIT)
# =========================================================================
elif menu == "📊 Hasil Pengecekan (Laporan)":
    st.title("📊 Laporan Komprehensif & Edit Data")
    st.markdown("Rekam jejak parameter teknis, generate WhatsApp, dan perbaiki data yang salah.")
    st.divider()

    db = st.session_state['laporan_db']

    if len(db) == 0:
        st.info("Data laporan kosong. Silakan isi form 1 kali di menu sebelah kiri.")
    else:
        for i in range(len(db) - 1, -1, -1):
            r = db[i]
            
            wa_text = f"""*REPORT HEALTH CHECK SOLAR BTS* ⚡\n📍 *Site:* {r.get('site_name', '-')} ({r.get('nop', '-')})\n📅 *Tanggal:* {r.get('timestamp', '-')}\n👷 *Teknisi:* {r.get('teknisi', '-')}\n📊 *Status Akhir:* {r.get('status', '-')}\n\n*TINDAKAN (ACTION):*\n{r.get('action', '-')}\n\n*KEBUTUHAN SPAREPART:*\n{r.get('sparepart', '-')}\n\n*DATA TEKNIS UTAMA:*\n- Total Load: {r.get('total_load', '-')} A\n- Nilai Grounding: {r.get('earth_ohm', '-')} Ohm\n- SCC / Recty: {r.get('scc_brand', '-')}"""
            wa_url = f"https://wa.me/?text={urllib.parse.quote(wa_text)}"

            with st.expander(f"📍 {r.get('site_name', 'Unknown')} | {r.get('timestamp', '')} | Status: {r.get('status', 'Unknown')}"):
                
                st.link_button("📱 Generate Rangkuman ke WhatsApp", wa_url)
                
                edit_mode = st.toggle("✏️ Edit Seluruh Laporan & Tambah Foto", key=f"edit_toggle_{i}")
                
                up_fisik = up_panel = up_bat = up_elek = None

                if edit_mode:
                    st.info("Mode Edit: Perbaiki teks/angka atau tambahkan foto yang tertinggal di tiap Tab di bawah. Jangan lupa klik **Simpan Perubahan**!")
                    
                    r['teknisi'] = st.text_input("Teknisi", r.get('teknisi', ''), key=f"t_{i}")
                    r['action'] = st.text_area("Tindakan / Action", r.get('action', ''), key=f"a_{i}")
                    r['sparepart'] = st.text_input("Kebutuhan Sparepart", r.get('sparepart', ''), key=f"s_{i}")
                    
                    status_options = ["Normal (Aman)", "Minor Issue (Sudah Diperbaiki)", "Major / Critical (Butuh Part)"]
                    current_status = r.get('status', "Normal (Aman)")
                    safe_index = status_options.index(current_status) if current_status in status_options else 0
                    r['status'] = st.selectbox("Status Akhir", status_options, index=safe_index, key=f"st_{i}")
                else:
                    st.markdown(f"**Teknisi:** {r.get('teknisi', '-')} | **NOP:** {r.get('nop', '-')} | **Load Beban:** {r.get('total_load', '-')} A")
                    st.markdown(f"**Tindakan Eksekusi (Action):**\n> {r.get('action', '-')}")
                    if r.get('sparepart'):
                        st.warning(f"**Kebutuhan Sparepart:** {r['sparepart']}")
                
                ltab1, ltab2, ltab3, ltab4 = st.tabs(["Fisik & Umum", "Data Panel Surya", "Bank Baterai", "Elektrikal (SCC/JB)"])
                
                # --- TAB 1: FISIK ---
                with ltab1:
                    if edit_mode:
                        r['site_cond'] = st.text_input("Kondisi Site", r.get('site_cond', ''), key=f"sc_{i}")
                        r['tower_cond'] = st.text_input("Kondisi Tower", r.get('tower_cond', ''), key=f"tc_{i}")
                        r['shading_status'] = st.text_input("Status Shading", r.get('shading_status', ''), key=f"sh_{i}")
                        up_fisik = st.file_uploader("📸 Tambah Foto Susulan (Fisik/Shading)", accept_multiple_files=True, key=f"uf_{i}")
                    else:
                        st.write(f"- Kondisi Site: {r.get('site_cond', '-')} | Tower: {r.get('tower_cond', '-')}")
                        st.write(f"- Status Shading: {r.get('shading_status', '-')}")
                    
                    all_fisik = get_safe_urls(r, ['url_sites', 'url_shadings', 'extras_fisik'])
                    if all_fisik:
                        cols = st.columns(4)
                        for idx, url in enumerate(all_fisik): cols[idx%4].image(url, width=150)

                # --- TAB 2: PANEL SURYA (AMAN & BULLETPROOF) ---
                with ltab2:
                    if edit_mode:
                        up_panel = st.file_uploader("📸 Tambah Foto Susulan (Panel)", accept_multiple_files=True, key=f"up_{i}")
                    
                    for idx_p, p in enumerate(r.get('panel_data', [])):
                        if edit_mode:
                            st.markdown(f"**{p.get('Panel', '-')}**")
                            c1, c2, c3 = st.columns(3)
                            with c1: p['Voc'] = st.number_input(f"Voc", value=float(p.get('Voc', 0) or 0), key=f"pvoc_{i}_{idx_p}")
                            with c2: p['Isc'] = st.number_input(f"Isc", value=float(p.get('Isc', 0) or 0), key=f"pisc_{i}_{idx_p}")
                            with c3: p['Kondisi'] = st.text_input(f"Kondisi", value=p.get('Kondisi', ''), key=f"pcon_{i}_{idx_p}")
                        else:
                            tipe = p.get('Tipe')
                            # Render Teks Berdasarkan Tipe
                            if tipe == 'Individu' or tipe is None:
                                st.markdown(f"**{p.get('Panel', '-')}** - Voc: {p.get('Voc', '-')}V | Isc: {p.get('Isc', '-')}A | Fisik: {p.get('Kondisi', '-')}")
                            else:
                                st.markdown(f"**{p.get('Panel', '-')} ({p.get('Qty', '-')} Panel)** - Total Voc: {p.get('Voc', '-')}V | Isc: {p.get('Isc', '-')}A | Fisik: {p.get('Kondisi', '-')}")
                            
                            # Ekstraksi Foto Universal
                            bef_urls = get_safe_urls(p, ['URL_Before', 'URLs_Before', 'foto_before_obj'])
                            aft_urls = get_safe_urls(p, ['URL_After', 'URLs_After', 'foto_after_obj'])
                            
                            cb, ca = st.columns(2)
                            with cb:
                                for u in bef_urls: st.image(u, caption="Before", width=250)
                            with ca:
                                for u in aft_urls: st.image(u, caption="After", width=250)
                        st.divider()
                    
                    if r.get('extras_panel'):
                        st.markdown("**Foto Tambahan Panel:**")
                        cols = st.columns(4)
                        for idx, url in enumerate(r['extras_panel']): cols[idx%4].image(url, width=150)

                # --- TAB 3: BATERAI (AMAN & BULLETPROOF) ---
                with ltab3:
                    if edit_mode:
                        up_bat = st.file_uploader("📸 Tambah Foto Susulan (Baterai)", accept_multiple_files=True, key=f"ub_{i}")
                    
                    for idx_b, b in enumerate(r.get('battery_data', [])):
                        if edit_mode:
                            st.markdown(f"**{b.get('Baterai')}**")
                            c1, c2, c3 = st.columns(3)
                            with c1: b['Voltase'] = st.number_input(f"Voltase", value=float(b.get('Voltase', 0) or 0), key=f"bv_{i}_{idx_b}")
                            with c2: b['Suhu'] = st.number_input(f"Suhu °C", value=float(b.get('Suhu', 0) or 0), key=f"bs_{i}_{idx_b}")
                            with c3: b['Kondisi'] = st.text_input(f"Kondisi Fisik", value=b.get('Kondisi', ''), key=f"bcon_{i}_{idx_b}")
                        else:
                            st.markdown(f"**{b.get('Baterai')}** - {b.get('Voltase', '-')}V | Suhu: {b.get('Suhu', '-')}°C | Fisik: {b.get('Kondisi', '-')}")
                            
                            # Ekstraksi Foto Universal Baterai
                            bat_urls = get_safe_urls(b, ['URL_Fotos', 'URL_Foto', 'foto_objs'])
                            if bat_urls:
                                cols = st.columns(3)
                                for idx, u in enumerate(bat_urls): cols[idx%3].image(u, width=200)
                        st.divider()
                    
                    if r.get('extras_baterai'):
                        st.markdown("**Foto Tambahan Baterai:**")
                        cols = st.columns(4)
                        for idx, url in enumerate(r['extras_baterai']): cols[idx%4].image(url, width=150)

                # --- TAB 4: ELEKTRIKAL ---
                with ltab4:
                    if edit_mode:
                        c1, c2 = st.columns(2)
                        with c1:
                            r['total_load'] = st.number_input("Total Load (A)", value=float(r.get('total_load', 0) or 0), key=f"ld_{i}")
                            r['earth_ohm'] = st.number_input("Earth/Grounding (Ohm)", value=float(r.get('earth_ohm', 0) or 0), key=f"ea_{i}")
                            r['datalog_status'] = st.text_input("Status Datalog", r.get('datalog_status', ''), key=f"dt_{i}")
                        with c2:
                            r['jb_enclosure'] = st.text_input("Kondisi Box JB", r.get('jb_enclosure', ''), key=f"jbe_{i}")
                            r['scc_alarm'] = st.text_input("Alarm SCC", r.get('scc_alarm', ''), key=f"sca_{i}")
                        up_elek = st.file_uploader("📸 Tambah Foto Susulan (SCC/JB/Grounding)", accept_multiple_files=True, key=f"ue_{i}")
                    else:
                        c1, c2 = st.columns(2)
                        with c1:
                            st.write(f"**Panel DC / JB:** \n- Box: {r.get('jb_enclosure', '-')} \n- Breaker: {r.get('jb_breaker', '-')} \n- Kabel: {r.get('cabling', '-')}")
                        with c2:
                            st.write(f"**SCC / Elektrikal:** \n- SCC: {r.get('scc_brand', '-')} \n- Alarm: {r.get('scc_alarm', '-')} \n- Datalog: {r.get('datalog_status', '-')}")
                        st.write(f"**Grounding:** {r.get('earth_ohm', '-')} Ohm ({r.get('grd_cable', '-')})")
                    
                    all_elek = get_safe_urls(r, ['url_jbs', 'url_sccs', 'url_grds', 'extras_elektrikal'])
                    if all_elek:
                        cols = st.columns(4)
                        for idx, url in enumerate(all_elek): cols[idx%4].image(url, width=150)

                # --- TOMBOL SIMPAN ---
                if edit_mode:
                    st.divider()
                    if st.button("💾 SIMPAN SEMUA PERUBAHAN", key=f"save_{i}", type="primary"):
                        with st.spinner("Menyimpan teks & mengunggah foto tambahan ke Cloudinary..."):
                            if up_fisik: r['extras_fisik'] = (r.get('extras_fisik') or []) + upload_multiple_images(up_fisik, "solar_extra")
                            if up_panel: r['extras_panel'] = (r.get('extras_panel') or []) + upload_multiple_images(up_panel, "solar_extra")
                            if up_bat: r['extras_baterai'] = (r.get('extras_baterai') or []) + upload_multiple_images(up_bat, "solar_extra")
                            if up_elek: r['extras_elektrikal'] = (r.get('extras_elektrikal') or []) + upload_multiple_images(up_elek, "solar_extra")
                            save_db(st.session_state['laporan_db']) 
                        st.success("✅ Perubahan & Tambahan Foto berhasil disimpan!")
                        st.rerun()
