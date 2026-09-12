import datetime
import pandas as pd
import streamlit as st
import cloudinary
import cloudinary.uploader
import json
import os

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
    """Fungsi upload 1 gambar ke Cloudinary"""
    if file_obj is not None:
        try:
            response = cloudinary.uploader.upload(file_obj.getvalue(), folder=folder_name)
            return response.get('secure_url')
        except Exception as e:
            st.error(f"Gagal upload gambar: {e}")
            return None
    return None

def upload_multiple_images(file_objs, folder_name="solar_bts_healthcheck"):
    """Fungsi upload banyak gambar sekaligus, return list of URLs"""
    urls = []
    if file_objs:
        for file in file_objs:
            url = upload_image(file, folder_name)
            if url: urls.append(url)
    return urls

# -------------------------------------------------------------------------
# 2. SETUP HYBRID DATABASE (JSON + SESSION STATE)
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
    st.info("💡 Data tersimpan permanen. Support mode pengukuran Panel Satuan maupun String/Seri.")
    
    tabs = st.tabs([
        "1. Info Site", 
        "2. Solar Panel (Seri/Satuan)", 
        "3. Panel DC & Kabel", 
        "4. Inverter & SCC", 
        "5. Baterai", 
        "6. Grounding", 
        "7. Action & Submit"
    ])

    # --- TAB 1: INFORMASI SITE ---
    with tabs[0]:
        st.subheader("Informasi Umum & Kondisi Fisik Site")
        col1, col2 = st.columns(2)
        with col1:
            site_name = st.text_input("Nama / ID Site", placeholder="Contoh: BTS-PKY-001")
            nop_area = st.selectbox("NOP Area", ["Palangkaraya", "Pangkalan Bun", "Tarakan", "Pontianak", "Lainnya"])
            check_date = st.date_input("Tanggal Pengecekan", value=datetime.date.today())
        with col2:
            technician_name = st.text_input("Nama Teknisi", placeholder="Nama lengkap tim")
            weather = st.selectbox("Kondisi Cuaca saat Eksekusi", ["Cerah", "Berawan", "Hujan", "Badai/Ekstrem"])

        st.divider()
        st.markdown("**Dokumentasi Fisik Site**")
        site_col1, site_col2 = st.columns(2)
        with site_col1:
            site_condition = st.selectbox("Kondisi Halaman & Pagar", ["Bersih & Aman", "Banyak Rumput / Belukar", "Pagar Rusak"])
            tower_condition = st.selectbox("Kondisi Fisik Tower/Tiang", ["Aman (Tidak Berkarat)", "Berkarat / Baut Kendur"])
        with site_col2:
            site_photos = st.file_uploader("Upload Foto View Site (Bisa >1 Foto)", type=["jpg", "png", "jpeg"], accept_multiple_files=True, key="site_pics")

    # --- TAB 2: SOLAR PANEL (PILIHAN SERI / SATUAN) ---
    with tabs[1]:
        st.subheader("Pengecekan Modul Surya")
        
        st.markdown("**A. Lingkungan (Shading)**")
        shade_col1, shade_col2 = st.columns(2)
        with shade_col1:
            shading_status = st.selectbox("Status Shading?", ["Aman (Clear area)", "Sedikit Shading (Pohon)", "Shading Kritis"])
        with shade_col2:
            shading_photos = st.file_uploader("Foto Bukti Shading (Bisa >1)", type=["jpg", "png", "jpeg"], accept_multiple_files=True, key="shd")

        st.divider()
        st.markdown("**B. Pengukuran Kelistrikan & Foto Panel**")
        
        # FITUR BARU: PILIHAN TOPOLOGI PANEL
        panel_config = st.radio(
            "Pilih Metode Pengukuran Panel (Sesuai Konfigurasi Site):",
            ["Per String / Seri (Digabung, contoh: Huawei/Sungrow)", "Per Panel / Satuan (Individu, contoh: Morningstar)"]
        )
        
        panel_data = []

        if panel_config == "Per String / Seri (Digabung, contoh: Huawei/Sungrow)":
            num_strings = st.number_input("Jumlah String / Seri di Site", min_value=1, max_value=20, value=2)
            
            for i in range(int(num_strings)):
                with st.expander(f"⚙️ Pengukuran String #{i+1}", expanded=(i == 0)):
                    p_col1, p_col2 = st.columns(2)
                    with p_col1:
                        jml_panel_dlm_string = st.number_input(f"Jumlah Panel yang di-Seri pada String {i+1}", min_value=1, value=6, key=f"jml_{i}")
                        voc = st.number_input(f"Voc Total String {i+1} [V]", value=120.5, step=0.1, key=f"v_str_{i}")
                        isc = st.number_input(f"Isc Total String {i+1} [A]", value=8.5, step=0.1, key=f"i_str_{i}")
                        p_cond = st.selectbox(f"Kondisi Mayoritas String {i+1}", ["Baik", "Sangat Kotor", "Ada yang Retak"], key=f"c_str_{i}")
                    with p_col2:
                        st.info("Upload Foto Full View (Bisa seleksi beberapa foto sekaligus)")
                        p_photo_before = st.file_uploader(f"📸 Foto BEFORE (Kotor) - String {i+1}", type=["jpg", "png", "jpeg"], accept_multiple_files=True, key=f"pb_str_{i}")
                        p_photo_after = st.file_uploader(f"✨ Foto AFTER (Bersih) - String {i+1}", type=["jpg", "png", "jpeg"], accept_multiple_files=True, key=f"pa_str_{i}")

                    panel_data.append({
                        "id": f"String #{i+1}", "tipe": "String", "info": f"Isi {jml_panel_dlm_string} Panel",
                        "voc": voc, "isc": isc, "kondisi": p_cond, 
                        "foto_before_objs": p_photo_before, "foto_after_objs": p_photo_after
                    })

        else:
            # Mode Satuan
            num_panels = st.number_input("Jumlah Total Panel", min_value=1, max_value=60, value=12)
            for i in range(int(num_panels)):
                with st.expander(f"🔲 Pengukuran Panel Satuan #{i+1}", expanded=(i == 0)):
                    p_col1, p_col2 = st.columns(2)
                    with p_col1:
                        voc = st.number_input(f"Voc Panel {i+1} [V]", value=21.5, step=0.1, key=f"v_pan_{i}")
                        isc = st.number_input(f"Isc Panel {i+1} [A]", value=5.2, step=0.1, key=f"i_pan_{i}")
                        p_cond = st.selectbox(f"Kondisi Fisik Panel {i+1}", ["Baik", "Kotor", "Retak"], key=f"c_pan_{i}")
                    with p_col2:
                        p_photo_before = st.file_uploader(f"📸 Foto BEFORE - Panel {i+1}", type=["jpg", "png", "jpeg"], accept_multiple_files=True, key=f"pb_pan_{i}")
                        p_photo_after = st.file_uploader(f"✨ Foto AFTER - Panel {i+1}", type=["jpg", "png", "jpeg"], accept_multiple_files=True, key=f"pa_pan_{i}")

                    panel_data.append({
                        "id": f"Panel #{i+1}", "tipe": "Satuan", "info": "1 Panel",
                        "voc": voc, "isc": isc, "kondisi": p_cond, 
                        "foto_before_objs": p_photo_before, "foto_after_objs": p_photo_after
                    })

    # --- TAB 3: PANEL DC, JUNCTION BOX & KABEL ---
    with tabs[2]:
        st.subheader("Junction Box, Panel DC & Perkabelan")
        jb_col1, jb_col2 = st.columns(2)
        with jb_col1:
            jb_enclosure = st.selectbox("Kondisi Box Panel DC", ["Bersih & Kedap Air", "Bocor / Ada Serangga"])
            jb_breaker = st.selectbox("Status MCB / Fuse", ["Normal", "Trip / Putus / Gosong"])
            jb_spd = st.selectbox("Arrester / SPD", ["Normal (Hijau)", "Rusak (Merah/Hitam)"])
            cabling = st.selectbox("Kondisi Kabel & Kontaktor", ["Rapi & Kuat", "Oksidasi / Chattering"])
        with jb_col2:
            jb_photos = st.file_uploader("Foto Panel DC (Bisa >1 Foto)", type=["jpg", "png", "jpeg"], accept_multiple_files=True, key="jb")

    # --- TAB 4: SCC, RECTIFIER & INVERTER ---
    with tabs[3]:
        st.subheader("Pengontrol Daya, Rectifier & Inverter")
        scc_col1, scc_col2 = st.columns(2)
        with scc_col1:
            scc_brand = st.text_input("Merek SCC & Rectifier", placeholder="Contoh: Morningstar, SmartGen")
            system_out_v = st.number_input("Tegangan Output ke Beban [Volt]", value=48.0)
            total_load_a = st.number_input("Total Load (Ampere)", value=15.5, step=0.1)
            inv_status = st.selectbox("Kondisi Inverter", ["Normal / 220V", "Error", "Tidak Pakai Inverter"])
            is_datalog_taken = st.checkbox("✅ Datalog berhasil ditarik")
        with scc_col2:
            scc_alarm = st.selectbox("Status Layar / LED SCC", ["Normal", "Ada Alarm Fault"])
            scc_photos = st.file_uploader("Foto Layar SCC (Bisa >1)", type=["jpg", "png", "jpeg"], accept_multiple_files=True, key="scc")

    # --- TAB 5: BATERAI ---
    with tabs[4]:
        st.subheader("Pemeriksaan Baterai")
        bat_total_col1, bat_total_col2 = st.columns(2)
        with bat_total_col1:
            num_batteries = st.number_input("Jumlah Blok Baterai (Maks 12)", min_value=1, max_value=12, value=4)
        with bat_total_col2:
            bat_total_v = st.number_input("Total Tegangan Bank Baterai [V]", value=52.8)
        
        bat_data = []
        for j in range(int(num_batteries)):
            with st.expander(f"Unit Baterai #{j+1}", expanded=(j==0)):
                b_col1, b_col2, b_col3 = st.columns([1, 1, 1.2])
                with b_col1:
                    b_volt = st.number_input(f"Voltase {j+1} [V]", value=12.2, step=0.1, key=f"bv_{j}")
                    b_temp = st.number_input(f"Suhu {j+1} [°C]", value=28.0, step=0.1, key=f"bt_{j}")
                with b_col2:
                    b_cond = st.selectbox(f"Fisik Baterai {j+1}", ["Normal", "Bengkak", "Terminal Korosi"], key=f"bc_{j}")
                with b_col3:
                    b_photo = st.file_uploader(f"Foto Baterai {j+1}", type=["jpg", "png", "jpeg"], accept_multiple_files=True, key=f"bp_{j}")
                
                bat_data.append({"id": f"Baterai #{j+1}", "voltase": b_volt, "suhu": b_temp, "kondisi": b_cond, "foto_objs": b_photo})

    # --- TAB 6: GROUNDING ---
    with tabs[5]:
        st.subheader("Grounding & Proteksi Petir")
        grd_col1, grd_col2 = st.columns(2)
        with grd_col1:
            earth_resistance = st.number_input("Tahanan Grounding [Ohm]", value=2.1, step=0.1)
            grd_cable = st.selectbox("Kondisi Kabel BC", ["Kuat", "Kendor", "Putus"])
        with grd_col2:
            grd_photos = st.file_uploader("Foto Alat Ukur & Busbar", type=["jpg", "png", "jpeg"], accept_multiple_files=True, key="grd")

    # --- TAB 7: ACTION & SUBMIT ---
    with tabs[6]:
        st.subheader("Ringkasan & Submit")
        action_taken = st.text_area("🔧 Ketik Mandiri Action yang Dilakukan:", height=120)
        sparepart_needed = st.text_input("📦 Daftar Sparepart yang Dibutuhkan/Diganti:")
        final_status = st.radio("Status Akhir Site:", ["Normal (Aman)", "Minor Issue (Sudah Diperbaiki)", "Major (Butuh Part)"])

        if st.button("🚀 Upload & Submit Laporan", type="primary"):
            if not site_name or not technician_name:
                st.error("⚠️ Mohon isi Nama Site dan Nama Teknisi di Tab 1!")
            else:
                with st.spinner("⏳ Memproses multi-upload Cloudinary & DB... JANGAN DI-REFRESH."):
                    url_sites = upload_multiple_images(site_photos, "solar_view_site")
                    url_shadings = upload_multiple_images(shading_photos, "solar_shading")
                    url_jbs = upload_multiple_images(jb_photos, "solar_junction_box")
                    url_sccs = upload_multiple_images(scc_photos, "solar_scc_inverter")
                    url_grds = upload_multiple_images(grd_photos, "solar_grounding")

                    # Memproses Panel Array (Support List Gambar)
                    panel_results = []
                    for p in panel_data:
                        url_befs = upload_multiple_images(p["foto_before_objs"], "solar_panel_before")
                        url_afts = upload_multiple_images(p["foto_after_objs"], "solar_panel_after")
                        panel_results.append({
                            "Panel": p["id"], "Tipe": p["tipe"], "Info": p["info"],
                            "Voc": p["voc"], "Isc": p["isc"], "Kondisi": p["kondisi"], 
                            "URL_Befores": url_befs, "URL_Afters": url_afts
                        })
                    
                    bat_results = []
                    for b in bat_data:
                        url_bats = upload_multiple_images(b["foto_objs"], "solar_battery")
                        bat_results.append({
                            "Baterai": b["id"], "Voltase": b["voltase"], "Suhu": b["suhu"],
                            "Kondisi": b["kondisi"], "URL_Fotos": url_bats
                        })

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
                        "panel_config": panel_config, "panel_data": panel_results, "battery_data": bat_results
                    }
                    
                    st.session_state['laporan_db'].append(report_dict)
                    save_db(st.session_state['laporan_db'])

                st.success(f"✅ Laporan lengkap Site {site_name} berhasil disimpan PERMANEN!")

# =========================================================================
# MENU 2: HASIL PENGECEKAN
# =========================================================================
elif menu == "📊 Hasil Pengecekan (Laporan)":
    st.title("📊 Laporan Komprehensif Eksekusi Site")
    st.divider()

    db = st.session_state['laporan_db']

    if len(db) == 0:
        st.info("Data laporan masih kosong. Silakan isi form 1 kali.")
    else:
        for i in range(len(db) - 1, -1, -1):
            r = db[i]
            with st.expander(f"📍 {r['site_name']} | {r['timestamp']} | Status: {r.get('status', 'Unknown')}"):
                
                edit_mode = st.toggle("✏️ Edit Teks Laporan", key=f"edit_toggle_{i}")
                if edit_mode:
                    new_teknisi = st.text_input("Edit Teknisi:", value=r.get('teknisi', ''), key=f"tek_{i}")
                    new_action = st.text_area("Edit Action:", value=r.get('action', ''), key=f"act_{i}", height=100)
                    new_sparepart = st.text_input("Edit Sparepart:", value=r.get('sparepart', ''), key=f"sp_{i}")
                    if st.button("💾 Simpan Teks", key=f"save_{i}", type="primary"):
                        st.session_state['laporan_db'][i]['teknisi'] = new_teknisi
                        st.session_state['laporan_db'][i]['action'] = new_action
                        st.session_state['laporan_db'][i]['sparepart'] = new_sparepart
                        save_db(st.session_state['laporan_db']) 
                        st.success("Teks berhasil diupdate!")
                        st.rerun()
                    st.divider()

                st.markdown(f"**Teknisi:** {r.get('teknisi', '-')} | **Load:** {r.get('total_load', '-')} A")
                st.markdown(f"**Topologi Panel:** {r.get('panel_config', 'Satuan')}")
                st.markdown(f"**Action:**\n> {r.get('action', '-')}")
                if r.get('sparepart'):
                    st.warning(f"**Sparepart:** {r['sparepart']}")
                
                ltab1, ltab2, ltab3, ltab4 = st.tabs(["Fisik Site", "Data Solar Panel", "Baterai", "Elektrikal"])
                
                with ltab1:
                    st.write(f"- Kondisi Site: {r.get('site_cond', '-')} | Tower: {r.get('tower_cond', '-')}")
                    st.write(f"- Shading: {r.get('shading_status', '-')}")
                    
                    if r.get('url_sites'):
                        st.markdown("**Foto Site:**")
                        cols = st.columns(4)
                        for idx, url in enumerate(r['url_sites']): cols[idx%4].image(url, width=150)
                    if r.get('url_shadings'):
                        st.markdown("**Foto Shading:**")
                        cols = st.columns(4)
                        for idx, url in enumerate(r['url_shadings']): cols[idx%4].image(url, width=150)

                with ltab2:
                    for p in r.get('panel_data', []):
                        st.markdown(f"**{p.get('Panel', '-')}** ({p.get('Info', '')}) - Voc: {p.get('Voc', '-')}V | Isc: {p.get('Isc', '-')}A | Kondisi: {p.get('Kondisi', '-')}")
                        c_bef, c_aft = st.columns(2)
                        with c_bef:
                            st.caption("📷 BEFORE")
                            if p.get('URL_Befores'):
                                for u in p['URL_Befores']: st.image(u, width=200)
                            elif p.get('URL_Before'): # Fallback data lama
                                st.image(p['URL_Before'], width=200)
                        with c_aft:
                            st.caption("✨ AFTER")
                            if p.get('URL_Afters'):
                                for u in p['URL_Afters']: st.image(u, width=200)
                            elif p.get('URL_After'):
                                st.image(p['URL_After'], width=200)
                        st.divider()

                with ltab3:
                    for b in r.get('battery_data', []):
                        st.markdown(f"**{b.get('Baterai', '-')}** - {b.get('Voltase', '-')}V | Suhu: {b.get('Suhu', '-')}°C | Fisik: {b.get('Kondisi', '-')}")
                        if b.get('URL_Fotos'):
                            cols = st.columns(3)
                            for idx, url in enumerate(b['URL_Fotos']): cols[idx%3].image(url, width=200)
                        st.divider()

                with ltab4:
                    col_e1, col_e2 = st.columns(2)
                    with col_e1:
                        st.markdown("**Panel DC / JB**")
                        st.write(f"Box: {r.get('jb_enclosure', '-')} \nBreaker: {r.get('jb_breaker', '-')}")
                        if r.get('url_jbs'):
                            for url in r['url_jbs']: st.image(url, width=150)
                    with col_e2:
                        st.markdown("**SCC / Inverter**")
                        st.write(f"Merek: {r.get('scc_brand', '-')} \nDatalog: {r.get('datalog_status', '-')}")
                        if r.get('url_sccs'):
                            for url in r['url_sccs']: st.image(url, width=150)
                    st.divider()
                    st.write(f"**Grounding:** {r.get('earth_ohm', '-')} Ohm")
                    if r.get('url_grds'):
                        for url in r['url_grds']: st.image(url, width=150)
