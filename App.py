import datetime
import pandas as pd
import streamlit as st
import cloudinary
import cloudinary.uploader

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
# 2. SETUP HALAMAN & DATABASE
# -------------------------------------------------------------------------
st.set_page_config(page_title="Solar BTS Health Check Pro", page_icon="⚡", layout="wide")

if 'laporan_db' not in st.session_state:
    st.session_state['laporan_db'] = []

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
    st.title("⚡ Form Health Check Solar BTS (Detail & Ekstra)")
    st.markdown("Isi form dengan lengkap. Anda bisa memilih **banyak foto sekaligus** pada tombol upload (Tahan CTRL/Pilih jamak di HP).")
    
    # Menambah jumlah Tab agar lebih rapi
    tabs = st.tabs([
        "1. Info & View Site", 
        "2. Solar Panel", 
        "3. Panel DC & Kabel", 
        "4. Inverter & SCC", 
        "5. Baterai", 
        "6. Grounding", 
        "7. Action & Submit"
    ])

    # --- TAB 1: INFORMASI & VIEW SITE ---
    with tabs[0]:
        st.subheader("Informasi Umum & Kondisi Fisik Site")
        col1, col2 = st.columns(2)
        with col1:
            site_name = st.text_input("Nama / ID Site", placeholder="Contoh: BTS-PKY-001")
            nop_area = st.selectbox("NOP Area", ["Palangkaraya", "Pangkalan Bun", "Tarakan", "Pontianak", "Lainnya"])
            check_date = st.date_input("Tanggal Pengecekan", value=datetime.date.today())
        with col2:
            technician_name = st.text_input("Nama Teknisi", placeholder="Nama lengkap tim (Pisahkan koma jika >1)")
            weather = st.selectbox("Kondisi Cuaca saat Eksekusi", ["Cerah", "Berawan", "Hujan", "Badai/Ekstrem"])

        st.divider()
        st.markdown("**Dokumentasi Fisik Site (View Site)**")
        site_col1, site_col2 = st.columns(2)
        with site_col1:
            site_condition = st.selectbox("Kondisi Halaman & Pagar", ["Bersih & Aman", "Banyak Rumput Liar / Belukar", "Pagar Rusak / Gembok Hilang"])
            tower_condition = st.selectbox("Kondisi Fisik Tower / Tiang", ["Aman Tidak Berkarat", "Berkarat / Baut Kendur"])
        with site_col2:
            site_photos = st.file_uploader("Upload Foto View Site (Pagar, Halaman, Tower 4 Arah) - Bisa >1 Foto", type=["jpg", "png"], accept_multiple_files=True, key="site_pics")

    # --- TAB 2: SOLAR PANEL & SHADING ---
    with tabs[1]:
        st.subheader("Pengecekan Modul Surya & Shading")
        
        st.markdown("**A. Kondisi Lingkungan (Shading)**")
        shade_col1, shade_col2 = st.columns(2)
        with shade_col1:
            shading_status = st.selectbox("Status Shading?", ["Aman (Clear area)", "Sedikit Shading (Bayangan pohon pagi/sore)", "Shading Kritis (Tertutup rimbunan)"])
        with shade_col2:
            shading_photos = st.file_uploader("Upload Foto Bukti Shading (Bisa >1 Foto)", type=["jpg", "png"], accept_multiple_files=True, key="shd")

        st.divider()
        st.markdown("**B. Pengecekan Detail Per Panel (Fisik, Tegangan, Arus & Before-After)**")
        num_panels = st.number_input("Jumlah Panel (Umumnya 24-30)", min_value=1, max_value=60, value=24)

        panel_data = []
        for i in range(int(num_panels)):
            with st.expander(f"Panel / Modul #{i+1}", expanded=(i == 0)):
                p_col1, p_col2, p_col3 = st.columns([1, 1.2, 1.2])
                with p_col1:
                    voc = st.number_input(f"Voc [V]", min_value=0.0, value=21.5, step=0.1, key=f"v_{i}")
                    isc = st.number_input(f"Isc [A]", min_value=0.0, value=5.2, step=0.1, key=f"i_{i}")
                    p_cond = st.selectbox(f"Kondisi Fisik", ["Baik & Mulus", "Sangat Kotor/Berlumut", "Kaca Retak (Hotspot)", "Delaminasi (Terbakar)"], key=f"c_{i}")
                with p_col2:
                    p_photo_before = st.file_uploader(f"📸 Foto BEFORE - Panel {i+1}", type=["jpg", "png"], key=f"pb_{i}")
                with p_col3:
                    p_photo_after = st.file_uploader(f"✨ Foto AFTER - Panel {i+1}", type=["jpg", "png"], key=f"pa_{i}")

                panel_data.append({"id": f"Panel #{i+1}", "voc": voc, "isc": isc, "kondisi": p_cond, "foto_before_obj": p_photo_before, "foto_after_obj": p_photo_after})

    # --- TAB 3: PANEL DC, JUNCTION BOX & KABEL ---
    with tabs[2]:
        st.subheader("Junction Box, Panel DC & Perkabelan")
        jb_col1, jb_col2 = st.columns(2)
        with jb_col1:
            jb_enclosure = st.selectbox("Kondisi Box Panel DC", ["Bersih & Kedap Air (IP 65 Baik)", "Bocor / Kemasukan Air", "Ada Sarang Serangga/Semut"])
            jb_breaker = st.selectbox("Status MCB / Fuse / Breaker", ["Normal (On Semua)", "Ada yang Trip / Putus / Gosong"])
            jb_spd = st.selectbox("Arrester / Surge Protection (SPD)", ["Normal (Hijau)", "Rusak Tersambar Petir (Merah/Hitam)"])
            cabling = st.selectbox("Kondisi Kabel & Kontaktor (Jika ada)", ["Rapi & Kuat", "Kontak Lengket (Chattering)", "Kabel Terkelupas / Oksidasi"])
        with jb_col2:
            jb_photos = st.file_uploader("Upload Foto Detail Panel DC, Kabel, Kontaktor (Bisa >1 Foto)", type=["jpg", "png"], accept_multiple_files=True, key="jb")

    # --- TAB 4: SCC, RECTIFIER & INVERTER (EKSTRA) ---
    with tabs[3]:
        st.subheader("Pengontrol Daya, Rectifier & Inverter")
        scc_col1, scc_col2 = st.columns(2)
        with scc_col1:
            scc_brand = st.text_input("Merek/Tipe SCC & Rectifier", placeholder="Contoh: SmartGen, Eltek, Huawei")
            system_out_v = st.number_input("Tegangan Output ke Beban (BTS) [Volt]", value=48.0)
            total_load_a = st.number_input("Total Beban / Load (Ampere)", value=15.5, step=0.1)
            
            st.markdown("**Status Inverter (Jika Site Hybrid/Ada beban AC)**")
            inv_status = st.selectbox("Kondisi Inverter", ["Normal / Output 220V AC", "Error / Alarm Fault", "Tidak Pakai Inverter"])
            
            st.markdown("**Status Data**")
            is_datalog_taken = st.checkbox("✅ Datalog berhasil ditarik via Flashdisk/Kabel")
        with scc_col2:
            scc_alarm = st.selectbox("Status Layar / LED Indikator SCC", ["Normal (No Alarm)", "Ada Alarm Fault (Tulis di Action)"])
            scc_photos = st.file_uploader("Upload Foto Layar SCC, Recty, Inverter, Pembacaan Ampere (Bisa >1 Foto)", type=["jpg", "png"], accept_multiple_files=True, key="scc")

    # --- TAB 5: BATERAI (TAMBAHAN SUHU PER SEL) ---
    with tabs[4]:
        st.subheader("Pemeriksaan Baterai Detail")
        
        bat_total_col1, bat_total_col2 = st.columns(2)
        with bat_total_col1:
            num_batteries = st.number_input("Jumlah Blok Baterai (Maks 12)", min_value=1, max_value=12, value=4)
        with bat_total_col2:
            bat_total_v = st.number_input("Total Tegangan Bank Baterai [Volt]", value=52.8)
        
        bat_data = []
        for j in range(int(num_batteries)):
            with st.expander(f"Unit Baterai #{j+1}", expanded=(j==0)):
                b_col1, b_col2, b_col3 = st.columns([1, 1, 1.2])
                with b_col1:
                    b_volt = st.number_input(f"Voltase Bat {j+1} [V]", value=12.2, step=0.1, key=f"bv_{j}")
                    b_temp = st.number_input(f"Suhu Fisik Bat {j+1} [°C]", value=28.0, step=0.1, key=f"bt_{j}")
                with b_col2:
                    b_cond = st.selectbox(f"Fisik Baterai {j+1}", ["Normal (Tidak Menggelembung)", "Bengkak / Swelling", "Terminal Korosi Putih/Hijau"], key=f"bc_{j}")
                with b_col3:
                    b_photo = st.file_uploader(f"Upload Foto Detail Baterai {j+1}", type=["jpg", "png"], accept_multiple_files=True, key=f"bp_{j}")
                
                bat_data.append({"id": f"Baterai #{j+1}", "voltase": b_volt, "suhu": b_temp, "kondisi": b_cond, "foto_objs": b_photo})

    # --- TAB 6: GROUNDING ---
    with tabs[5]:
        st.subheader("Grounding & Proteksi Petir")
        grd_col1, grd_col2 = st.columns(2)
        with grd_col1:
            earth_resistance = st.number_input("Nilai Tahanan Grounding (Earth Tester) [Ohm]", value=2.1, step=0.1, help="Standar biasanya < 5 Ohm")
            grd_cable = st.selectbox("Kondisi Busbar & Kabel BC (Ground)", ["Terhubung Kuat & Utuh", "Baut Kendor / Karat", "Kabel Putus / Dicuri"])
        with grd_col2:
            grd_photos = st.file_uploader("Upload Foto Alat Ukur Grounding & Busbar (Bisa >1 Foto)", type=["jpg", "png"], accept_multiple_files=True, key="grd")

    # --- TAB 7: ACTION, SPAREPART & SUBMIT ---
    with tabs[6]:
        st.subheader("Ringkasan, Tindakan Eksekusi & Kebutuhan Material")
        
        action_taken = st.text_area("🔧 Ketik Mandiri Action / Eksekusi yang Dilakukan:", height=120, placeholder="Contoh:\n1. Pembersihan modul surya (24 unit)\n2. Pengencangan baut terminal Baterai 2\n3. Pembersihan kerak panel DC...")
        
        sparepart_needed = st.text_input("📦 Daftar Sparepart yang Dibutuhkan/Diganti (Jika Ada):", placeholder="Contoh: 1 Unit MCB DC 63A, 2 Meter Kabel Nanya...")
        
        final_status = st.radio("Status Akhir Site (Kesimpulan):", ["Normal (Aman, Siap Beroperasi Penuh)", "Minor Issue (Sudah Diperbaiki Sementara, Perlu Pantauan)", "Major / Critical (Berbahaya, Butuh Segera Penggantian Part)"])

        if st.button("🚀 Upload Foto & Submit Laporan", type="primary"):
            if not site_name or not technician_name:
                st.error("⚠️ Mohon isi Nama Site dan Nama Teknisi di Tab 1!")
            else:
                with st.spinner("⏳ Memproses banyak foto dan mengunggah ke sistem (Cloudinary)... Mohon bersabar."):
                    
                    # 1. Upload Foto View Site, Shading, JB, SCC, Grounding (Jamak)
                    url_sites = upload_multiple_images(site_photos, "solar_view_site")
                    url_shadings = upload_multiple_images(shading_photos, "solar_shading")
                    url_jbs = upload_multiple_images(jb_photos, "solar_junction_box")
                    url_sccs = upload_multiple_images(scc_photos, "solar_scc_inverter")
                    url_grds = upload_multiple_images(grd_photos, "solar_grounding")

                    # 2. Upload Foto Panel (Bef-Aft)
                    panel_results = []
                    for p in panel_data:
                        url_bef = upload_image(p["foto_before_obj"], "solar_panel_before") if p["foto_before_obj"] else None
                        url_aft = upload_image(p["foto_after_obj"], "solar_panel_after") if p["foto_after_obj"] else None
                        panel_results.append({
                            "Panel": p["id"], "Voc": p["voc"], "Isc": p["isc"], 
                            "Kondisi": p["kondisi"], "URL_Before": url_bef, "URL_After": url_aft
                        })
                    
                    # 3. Upload Foto Baterai (Bisa Jamak per Baterai)
                    bat_results = []
                    for b in bat_data:
                        url_bats = upload_multiple_images(b["foto_objs"], "solar_battery")
                        bat_results.append({
                            "Baterai": b["id"], "Voltase": b["voltase"], "Suhu": b["suhu"],
                            "Kondisi": b["kondisi"], "URL_Fotos": url_bats
                        })

                    # 4. Kumpulkan ke Dictionary Laporan
                    report_dict = {
                        "timestamp": datetime.datetime.now().strftime("%Y-%m-%d %H:%M:%S"),
                        "site_name": site_name, "nop": nop_area, "teknisi": technician_name,
                        "status": final_status, "action": action_taken, "sparepart": sparepart_needed,
                        
                        # Data Fisik Site & Shading
                        "site_cond": site_condition, "tower_cond": tower_condition, "url_sites": url_sites,
                        "shading_status": shading_status, "url_shadings": url_shadings,
                        
                        # Data DC, SCC, Load
                        "jb_enclosure": jb_enclosure, "jb_breaker": jb_breaker, "url_jbs": url_jbs,
                        "scc_brand": scc_brand, "scc_alarm": scc_alarm, "total_load": total_load_a, "inv_status": inv_status,
                        "datalog_status": "Ditarik" if is_datalog_taken else "Tidak Ditarik", "url_sccs": url_sccs,
                        
                        # Data Grounding
                        "earth_ohm": earth_resistance, "grd_cable": grd_cable, "url_grds": url_grds,
                        
                        # Data Arrays
                        "panel_data": panel_results, "battery_data": bat_results
                    }
                    
                    st.session_state['laporan_db'].append(report_dict)
                
                st.success(f"✅ Laporan lengkap Site {site_name} berhasil diunggah!")

# =========================================================================
# MENU 2: HASIL PENGECEKAN (DASHBOARD)
# =========================================================================
elif menu == "📊 Hasil Pengecekan (Laporan)":
    st.title("📊 Laporan Komprehensif Eksekusi Site")
    st.markdown("Rekam jejak parameter teknis, tindakan (*action*), dan bukti foto dokumentasi.")
    st.divider()

    db = st.session_state['laporan_db']

    if len(db) == 0:
        st.info("Data laporan masih kosong.")
    else:
        for index, r in enumerate(reversed(db)):
            with st.expander(f"📍 {r['site_name']} | {r['timestamp']} | Status: {r['status']}"):
                
                st.markdown(f"**Teknisi:** {r['teknisi']} | **NOP:** {r['nop']} | **Load Beban:** {r['total_load']} A")
                st.markdown(f"**Tindakan Eksekusi (Action):**\n> {r['action']}")
                if r['sparepart']:
                    st.warning(f"**Kebutuhan Sparepart:** {r['sparepart']}")
                
                # Membagi laporan dalam sub-tabs
                ltab1, ltab2, ltab3, ltab4 = st.tabs(["Fisik & Umum", "Data Panel (Bef-Aft)", "Baterai", "Elektrikal (SCC/JB/Gnd)"])
                
                # --- L-TAB 1: FISIK & SHADING ---
                with ltab1:
                    st.write(f"- Kondisi Site: {r['site_cond']} | Tower: {r['tower_cond']}")
                    st.write(f"- Status Shading: {r['shading_status']}")
                    
                    if r['url_sites']:
                        st.markdown("**Foto View Site:**")
                        cols = st.columns(len(r['url_sites']) if len(r['url_sites']) < 5 else 4)
                        for idx, url in enumerate(r['url_sites']): cols[idx%4].image(url, width=150)
                    if r['url_shadings']:
                        st.markdown("**Foto Shading:**")
                        cols = st.columns(len(r['url_shadings']) if len(r['url_shadings']) < 5 else 4)
                        for idx, url in enumerate(r['url_shadings']): cols[idx%4].image(url, width=150)

                # --- L-TAB 2: DATA PANEL ---
                with ltab2:
                    for p in r['panel_data']:
                        st.markdown(f"**{p['Panel']}** - Voc: {p['Voc']}V | Isc: {p['Isc']}A | Fisik: {p['Kondisi']}")
                        c_bef, c_aft = st.columns(2)
                        with c_bef:
                            if p['URL_Before']: st.image(p['URL_Before'], caption="Before", width=250)
                        with c_aft:
                            if p['URL_After']: st.image(p['URL_After'], caption="After", width=250)
                        st.divider()

                # --- L-TAB 3: BATERAI ---
                with ltab3:
                    for b in r['battery_data']:
                        st.markdown(f"**{b['Baterai']}** - {b['Voltase']}V | Suhu: {b['Suhu']}°C | Fisik: {b['Kondisi']}")
                        if b['URL_Fotos']:
                            cols = st.columns(len(b['URL_Fotos']) if len(b['URL_Fotos']) < 4 else 3)
                            for idx, url in enumerate(b['URL_Fotos']): cols[idx%3].image(url, width=200)
                        st.divider()

                # --- L-TAB 4: ELEKTRIKAL ---
                with ltab4:
                    col_e1, col_e2 = st.columns(2)
                    with col_e1:
                        st.markdown("**Panel DC / Junction Box**")
                        st.write(f"- Box: {r['jb_enclosure']} \n- Breaker: {r['jb_breaker']}")
                        if r['url_jbs']:
                            for url in r['url_jbs']: st.image(url, width=150)
                    with col_e2:
                        st.markdown("**SCC & Inverter**")
                        st.write(f"- Merek: {r['scc_brand']} \n- Status: {r['scc_alarm']} \n- Datalog: {r['datalog_status']}")
                        if r['url_sccs']:
                            for url in r['url_sccs']: st.image(url, width=150)
                    
                    st.divider()
                    st.markdown("**Grounding**")
                    st.write(f"- Nilai Tahanan: {r['earth_ohm']} Ohm \n- Kabel: {r['grd_cable']}")
                    if r['url_grds']:
                        cols = st.columns(3)
                        for idx, url in enumerate(r['url_grds']): cols[idx%3].image(url, width=150)
