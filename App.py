import datetime
import pandas as pd
import streamlit as st
import cloudinary
import cloudinary.uploader
import json
import os
import urllib.parse

# -------------------------------------------------------------------------
# 1. KONFIGURASI CLOUDINARY (AKUN ANDA)
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

# FUNGSI BARU: Render Foto Grid Super Rapi (Maksimal 3 Kolom per baris)
def tampilkan_grid_foto(url_data, caption=""):
    if not url_data: return
    
    urls = []
    if isinstance(url_data, str) and url_data.startswith("http"):
        urls = [url_data]
    elif isinstance(url_data, list):
        urls = [u for u in url_data if isinstance(u, str) and u.startswith("http")]
        
    if urls:
        if caption:
            st.markdown(f"*{caption}*")
        
        # Looping untuk membuat baris baru setiap 3 foto agar rapi dan tidak numpuk
        for i in range(0, len(urls), 3):
            cols = st.columns(3)
            for j in range(3):
                if i + j < len(urls):
                    cols[j].image(urls[i + j], use_container_width=True)

# -------------------------------------------------------------------------
# 2. SETUP DATABASE (HYBRID JSON BACKUP KE GOOGLE SHEETS TARGET)
# -------------------------------------------------------------------------
st.set_page_config(page_title="Preventive Maintenance BTS Pro", page_icon="⚡", layout="wide")
DB_FILE = "laporan_preventive_db.json"
GOOGLE_SHEET_URL_TARGET = "https://docs.google.com/spreadsheets/d/1HvgVicTWwO4RMQI6ZR3Mu3IgGicwjcLZl9mDN1auvJU/edit?usp=drivesdk"

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
# 3. NAVIGASI UTAMA
# -------------------------------------------------------------------------
st.sidebar.title("Navigasi Operasional")
st.sidebar.info(f"Target Sheet:\n`Report Preventive`\n🔗 [Buka Google Sheets]({GOOGLE_SHEET_URL_TARGET})")
menu = st.sidebar.radio("Pilih Menu:", ["📝 Form Preventive Check", "📊 Hasil Laporan & Dashboard"])
st.sidebar.markdown("---")

# =========================================================================
# MENU 1: FORM PENGECEKAN LENGKAP (SPS, PLN/RECTI, GENSET, SOLAR PANEL)
# =========================================================================
if menu == "📝 Form Preventive Check":
    st.title("⚡ Form Preventive Maintenance Komprehensif BTS")
    st.markdown("Isi checklist pemeliharaan berkala meliputi Solar Panel, PLN/Rectifier, Genset, dan Baterai.")
    
    tabs = st.tabs([
        "1. Info Site", 
        "2. Solar Panel (SPS)", 
        "3. PLN & Rectifier", 
        "4. Genset & BBM", 
        "5. Baterai & Grounding", 
        "6. Datalog & Submit"
    ])

    # --- TAB 1: INFO SITE ---
    with tabs[0]:
        st.subheader("Informasi Umum Site & Fisik")
        c1, c2 = st.columns(2)
        with c1:
            site_name = st.text_input("Nama / ID Site", placeholder="Contoh: BTS-PKY-001")
            nop_area = st.selectbox("NOP Area", ["Palangkaraya", "Pangkalan Bun", "Tarakan", "Pontianak", "Lainnya"])
            check_date = st.date_input("Tanggal Pengecekan", value=datetime.date.today())
        with c2:
            technician_name = st.text_input("Nama Tim Teknisi")
            weather = st.selectbox("Kondisi Cuaca", ["Cerah", "Berawan", "Hujan", "Badai"])
        st.divider()
        sc1, sc2 = st.columns(2)
        with sc1:
            site_condition = st.selectbox("Kondisi Halaman / Shelter", ["Bersih & Aman", "Banyak Rumput / Belukar", "Genangan Air"])
            tower_condition = st.selectbox("Fisik Tower & Pagar", ["Aman", "Berkarat / Kendor", "Pagar/Gembok Rusak"])
        with sc2:
            site_photos = st.file_uploader("Foto View Site (Pagar, Shelter, Tower - Bisa >1)", accept_multiple_files=True, key="spics")

    # --- TAB 2: SOLAR PANEL / SPS (SATUAN & SERI) ---
    with tabs[1]:
        st.subheader("Pemeriksaan Modul Surya (SPS) & Shading")
        shading_status = st.selectbox("Status Shading / Bayangan Pohon", ["Aman (Clear area)", "Sedikit Shading", "Shading Kritis"])
        shading_photos = st.file_uploader("Foto Bukti Shading (Bisa >1)", accept_multiple_files=True, key="shd")
        st.divider()
        
        konfigurasi_panel = st.radio("Metode Pengukuran SPS:", ["Individu (Satuan - ex: Morningstar)", "Seri / String (Grup - ex: Huawei/ZTE)"])
        panel_data = []

        if konfigurasi_panel == "Individu (Satuan)":
            num_panels = st.number_input("Jumlah Total Panel", min_value=1, value=24)
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
            num_strings = st.number_input("Jumlah String / Seri", min_value=1, value=4)
            for i in range(int(num_strings)):
                with st.expander(f"String Seri #{i+1}", expanded=(i==0)):
                    c1, c2 = st.columns([1, 1.5])
                    with c1:
                        qty = st.number_input(f"Isi Panel per Seri", value=6, key=f"sq_{i}")
                        voc = st.number_input(f"Total Voc Seri [V]", value=129.0, key=f"sv_{i}")
                        isc = st.number_input(f"Isc Seri [A]", value=5.2, key=f"si_{i}")
                        p_cond = st.selectbox(f"Kondisi", ["Baik", "Kotor", "Retak"], key=f"sc_{i}")
                    with c2:
                        pb = st.file_uploader(f"📸 BEFORE Seri (Bisa >1)", accept_multiple_files=True, key=f"spb_{i}")
                        pa = st.file_uploader(f"✨ AFTER Seri (Bisa >1)", accept_multiple_files=True, key=f"spa_{i}")
                    panel_data.append({"tipe": "Seri", "id": f"String #{i+1}", "qty": qty, "voc": voc, "isc": isc, "kondisi": p_cond, "foto_before_objs": pb, "foto_after_objs": pa})

    # --- TAB 3: STANDAR PENGECEKAN PLN & RECTIFIER / BTS ---
    with tabs[2]:
        st.subheader("Standar Pengecekan Grid PLN, Rectifier & Beban BTS")
        c1, c2 = st.columns(2)
        with c1:
            st.markdown("#### ⚡ Grid PLN")
            pln_status = st.selectbox("Status PLN", ["Normal (Hadir)", "Padam / Sering Drop", "Tidak Ada Jaringan PLN"])
            pln_r_s = st.number_input("Tegangan R-S [V]", value=380.0)
            pln_s_t = st.number_input("Tegangan S-T [V]", value=380.0)
            pln_r_n = st.number_input("Tegangan R-N [V]", value=220.0)
            
            st.markdown("#### 🔌 Panel DC / Junction Box")
            jb_enclosure = st.selectbox("Kondisi Box DC & Seal", ["Bersih & Kedap", "Bocor / Rusak", "Sarang Serangga"])
            jb_spd = st.selectbox("Status Arrester SPD DC", ["Normal (Hijau)", "Rusak (Merah/Hitam)"])
        
        with c2:
            st.markdown("#### 🔋 Rectifier & Beban BTS")
            rect_brand = st.text_input("Merek Rectifier", placeholder="Huawei / Eltek / ZTE")
            rect_out_v = st.number_input("Tegangan Output Rectifier [V]", value=53.5)
            total_load_a = st.number_input("Total Beban / Load BTS [A]", value=18.2)
            rect_alarm = st.selectbox("Status Alarm Rectifier", ["No Alarm (Normal)", "Ada Alarm Major/Minor"])
            rect_photos = st.file_uploader("Foto Tampilan Rectifier & MCB (Bisa >1)", accept_multiple_files=True, key="rect_pics")

    # --- TAB 4: STANDAR PENGECEKAN GENSET & KALKULATOR BBM ---
    with tabs[3]:
        st.subheader("Standar Pengecekan Genset & Manajemen BBM")
        
        g1, g2 = st.columns(2)
        with g1:
            genset_status = st.selectbox("Ketersediaan Genset di Site", ["Ada (Standby / Beroperasi)", "Tidak Ada Genset"])
            genset_brand = st.text_input("Merek / Kapasitas Genset", placeholder="Perkins / Cummins / SmartGen - 15kVA")
            genset_mode = st.radio("Mode Operasional Genset", ["Auto (AMF/ATS Berfungsi)", "Manual Only"])
            hour_meter = st.number_input("Hour Meter (Jam Operasi)", value=1250.5, step=0.1)
            
            st.markdown("#### 📊 Parameter Listrik Genset")
            gen_v = st.number_input("Tegangan Output Genset [V]", value=220.0)
            gen_hz = st.number_input("Frekuensi [Hz]", value=50.0)
            gen_temp = st.number_input("Suhu Mesin [°C]", value=75.0)
        
        with g2:
            st.markdown("#### ⛽ Leveling BBM & Kalkulator Tangki")
            tank_capacity = st.selectbox("Kapasitas Tangki Harian / Dasar", [200, 300, 500, 1000])
            current_fuel_pct = st.slider("Level BBM Saat Ini [%]", min_value=0, max_value=100, value=75)
            
            estimated_fuel_liter = (current_fuel_pct / 100.0) * tank_capacity
            estimated_backup_hours = estimated_fuel_liter / 3.5 
            
            st.info(f"📌 **Estimasi Volume BBM:** {estimated_fuel_liter:.1f} Liter dari total {tank_capacity} L.\n⏱️ **Estimasi Waktu Operasi:** ~{estimated_backup_hours:.1f} Jam.")
            
            genset_notes = st.selectbox("Kondisi Fisik & Aki Genset", ["Prima & Normal", "Aki Lemah / Drop", "Filter Perlu Dibersihkan", "Bocor / Rembesan Oli"])
            genset_photos = st.file_uploader("Foto Genset, Hour Meter & Level BBM (Bisa >1)", accept_multiple_files=True, key="genset_pics")

    # --- TAB 5: BATERAI & GROUNDING ---
    with tabs[4]:
        st.subheader("Pemeriksaan Bank Baterai & Grounding")
        b1, b2 = st.columns(2)
        with b1:
            num_bat = st.number_input("Jumlah Blok Baterai", min_value=1, max_value=16, value=4)
            bat_data = []
            for j in range(int(num_bat)):
                with st.expander(f"Baterai #{j+1}", expanded=(j==0)):
                    bc1, bc2 = st.columns(2)
                    with bc1:
                        bv = st.number_input(f"Voltase [V]", value=12.2, key=f"bv_{j}")
                        bt = st.number_input(f"Suhu [°C]", value=28.0, key=f"bt_{j}")
                    with bc2:
                        bc = st.selectbox(f"Kondisi", ["Normal", "Bengkak", "Korosi Terminal"], key=f"bc_{j}")
                        bp = st.file_uploader(f"Foto Bat #{j+1}", accept_multiple_files=True, key=f"bp_{j}")
                    bat_data.append({"id": f"Baterai #{j+1}", "voltase": bv, "suhu": bt, "kondisi": bc, "foto_objs": bp})
        
        with b2:
            st.markdown("#### 🌍 Sistem Grounding")
            earth_resistance = st.number_input("Tahanan Grounding [Ohm]", value=2.1)
            grd_cable = st.selectbox("Kondisi Kabel BC & Busbar", ["Kuat & Terhubung", "Kendor / Berkarat", "Putus"])
            grd_photos = st.file_uploader("Foto Hasil Earth Tester & Grounding (Bisa >1)", accept_multiple_files=True, key="grd")

    # --- TAB 6: DATALOG ALL FILE & SUBMIT ---
    with tabs[5]:
        st.subheader("Unggah Datalog All File & Finalisasi Laporan")
        
        st.markdown("📂 **Upload Datalog Universal (Semua Format File):**")
        datalog_files = st.file_uploader(
            "Pilih file datalog dari Rectifier, SCC, Inverter, atau Genset (Mendukung .csv, .xlsx, .txt, .dat, .zip)", 
            accept_multiple_files=True, 
            key="datalog_all"
        )
        
        st.divider()
        action_taken = st.text_area("🔧 Detail Tindakan (Action) di Lapangan:", height=100)
        sparepart_needed = st.text_input("📦 Kebutuhan Material / Sparepart (Jika Ada):")
        final_status = st.radio("Status Keseluruhan Site:", ["Normal (Aman / Siap Operasional)", "Minor Issue (Sudah Ditangani Sementara)", "Major / Critical (Perlu Eskalasi Cepat)"])

        if st.button("🚀 Submit Laporan & Sinkronkan ke Spreadsheet", type="primary"):
            if not site_name or not technician_name:
                st.error("⚠️ Mohon lengkapi Nama Site dan Nama Teknisi di Tab 1!")
            else:
                with st.spinner("⏳ Mengunggah foto ke Cloudinary & menyimpan data ke database..."):
                    url_sites = upload_multiple_images(site_photos, "preventive_view")
                    url_shadings = upload_multiple_images(shading_photos, "preventive_shade")
                    url_rects = upload_multiple_images(rect_photos, "preventive_rect")
                    url_gensets = upload_multiple_images(genset_photos, "preventive_genset")
                    url_grds = upload_multiple_images(grd_photos, "preventive_grd")
                    
                    datalog_urls = []
                    if datalog_files:
                        for df in datalog_files:
                            du = upload_image(df, "preventive_datalog")
                            if du: datalog_urls.append({"name": df.name, "url": du})

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

                    report_dict = {
                        "timestamp": datetime.datetime.now().strftime("%Y-%m-%d %H:%M:%S"),
                        "site_name": site_name, "nop": nop_area, "teknisi": technician_name,
                        "status": final_status, "action": action_taken, "sparepart": sparepart_needed,
                        
                        "site_cond": site_condition, "tower_cond": tower_condition, "url_sites": url_sites,
                        "shading_status": shading_status, "url_shadings": url_shadings,
                        "panel_data": p_res,
                        
                        "pln_status": pln_status, "rect_brand": rect_brand, "rect_out_v": rect_out_v, 
                        "total_load": total_load_a, "rect_alarm": rect_alarm, "url_rects": url_rects,
                        
                        "genset_status": genset_status, "genset_brand": genset_brand, "hour_meter": hour_meter,
                        "tank_capacity": tank_capacity, "fuel_pct": current_fuel_pct, "fuel_liter": estimated_fuel_liter,
                        "url_gensets": url_gensets,
                        
                        "battery_data": b_res, "earth_ohm": earth_resistance, "url_grds": url_grds,
                        
                        "datalog_files": datalog_urls,
                        
                        "extras_fisik": [], "extras_panel": [], "extras_baterai": [], "extras_elektrikal": []
                    }
                    
                    st.session_state['laporan_db'].append(report_dict) 
                    save_db(st.session_state['laporan_db'])            
                st.success(f"✅ Laporan Site **{site_name}** berhasil disimpan!")

# =========================================================================
# MENU 2: HASIL LAPORAN, EDIT & GENERATE WHATSAPP
# =========================================================================
elif menu == "📊 Hasil Laporan & Dashboard":
    st.title("📊 Dashboard Laporan Preventive Maintenance")
    st.markdown(f"Kelola data, edit teks laporan, tambah foto susulan bebas, dan kirim rekap ke WhatsApp.")
    st.divider()

    db = st.session_state['laporan_db']

    if not db:
        st.info("⚠️ Belum ada data laporan tersimpan. Silakan isi form di menu sebelah kiri.")
    else:
        for i in range(len(db) - 1, -1, -1):
            r = db[i]
            
            wa_text = f"""*REPORT PREVENTIVE MAINTENANCE BTS* ⚡\n📍 *Site:* {r.get('site_name', '-')} ({r.get('nop', '-')})\n📅 *Tanggal:* {r.get('timestamp', '-')}\n👷 *Teknisi:* {r.get('teknisi', '-')}\n📊 *Status:* {r.get('status', '-')}\n\n*POWER & LOAD:*\n- PLN: {r.get('pln_status', '-')}\n- Rectifier: {r.get('rect_brand', '-')} ({r.get('rect_out_v', '-')}V)\n- Load BTS: {r.get('total_load', '-')} A\n- Genset BBM: {r.get('fuel_pct', '-')}% (~{r.get('fuel_liter', '-')} Liter)\n\n*TINDAKAN (ACTION):*\n{r.get('action', '-')}\n\n*SPAREPART:*\n{r.get('sparepart', '-')}\n\n🔗 *Google Sheets Target:* Report Preventive"""
            wa_url = f"https://wa.me/?text={urllib.parse.quote(wa_text)}"

            with st.expander(f"📍 {r.get('site_name', 'Unknown')} | {r.get('timestamp', '')} | Status: {r.get('status', '')}"):
                
                col_btn1, col_btn2 = st.columns(2)
                with col_btn1:
                    st.link_button("📱 Kirim Rangkuman ke WhatsApp", wa_url)
                with col_btn2:
                    st.link_button("📈 Buka Spreadsheet Target", GOOGLE_SHEET_URL_TARGET)
                
                st.markdown(f"**Teknisi:** {r.get('teknisi', '-')} | **NOP:** {r.get('nop', '-')} | **Load:** {r.get('total_load', '-')} A")
                st.markdown(f"**Action:** {r.get('action', '-')}")
                if r.get('sparepart'): st.warning(f"**Sparepart:** {r['sparepart']}")
                
                ltab1, ltab2, ltab3, ltab4, ltab5 = st.tabs(["1. Fisik & View", "2. SPS Panel", "3. PLN, Recti & Genset", "4. Baterai & Grounding", "5. Datalog & File"])
                
                with ltab1:
                    st.write(f"- Kondisi Site: {r.get('site_cond', '-')} | Tower: {r.get('tower_cond', '-')} | Shading: {r.get('shading_status', '-')}")
                    tampilkan_grid_foto(r.get('url_sites'), "📸 Foto View Site")
                    tampilkan_grid_foto(r.get('url_shadings'), "📸 Foto Shading")
                    tampilkan_grid_foto(r.get('extras_fisik'), "📸 Foto Susulan Fisik (Bebas)")

                with ltab2:
                    for p in r.get('panel_data', []):
                        st.write(f"**{p.get('Panel', '-')}** | Voc: {p.get('Voc','-')}V | Isc: {p.get('Isc','-')}A | Fisik: {p.get('Kondisi','-')}")
                        cb, ca = st.columns(2)
                        with cb: 
                            tampilkan_grid_foto(p.get('URL_Before') or p.get('URLs_Before'), "📸 BEFORE")
                        with ca: 
                            tampilkan_grid_foto(p.get('URL_After') or p.get('URLs_After'), "✨ AFTER")
                        st.divider()
                    tampilkan_grid_foto(r.get('extras_panel'), "📸 Foto Susulan Panel (Bebas)")

                with ltab3:
                    c_p1, c_p2 = st.columns(2)
                    with c_p1:
                        st.markdown("**Grid PLN & Rectifier:**")
                        st.write(f"- PLN: {r.get('pln_status', '-')}")
                        st.write(f"- Rectifier: {r.get('rect_brand', '-')} ({r.get('rect_out_v', '-')}V)")
                        st.write(f"- Alarm Recti: {r.get('rect_alarm', '-')}")
                        tampilkan_grid_foto(r.get('url_rects'), "📸 Foto Rectifier")
                    with c_p2:
                        st.markdown("**Genset & Manajemen BBM:**")
                        st.write(f"- Status: {r.get('genset_status', '-')}")
                        st.write(f"- Tipe: {r.get('genset_brand', '-')}")
                        st.write(f"- Level BBM: {r.get('fuel_pct', '-')}% (~{r.get('fuel_liter', 0):.1f} L)")
                        st.write(f"- Hour Meter: {r.get('hour_meter', '-')} Jam")
                        tampilkan_grid_foto(r.get('url_gensets'), "📸 Foto Genset")
                    
                    st.divider()
                    tampilkan_grid_foto(r.get('extras_elektrikal'), "📸 Foto Susulan Elektrikal/Mesin (Bebas)")

                with ltab4:
                    for b in r.get('battery_data', []):
                        st.write(f"**{b.get('Baterai', '-')}** | {b.get('Voltase','-')}V | Suhu: {b.get('Suhu','-')}°C | Fisik: {b.get('Kondisi','-')}")
                        tampilkan_grid_foto(b.get('URL_Fotos'))
                        st.divider()
                    st.write(f"**Grounding Resistance:** {r.get('earth_ohm', '-')} Ohm")
                    tampilkan_grid_foto(r.get('url_grds'), "📸 Foto Grounding")
                    st.divider()
                    tampilkan_grid_foto(r.get('extras_baterai'), "📸 Foto Susulan Baterai/Grounding (Bebas)")

                with ltab5:
                    st.markdown("📂 **Datalog Universal yang Diunggah:**")
                    dls = r.get('datalog_files', [])
                    if dls:
                        for dl in dls:
                            st.markdown(f"- [{dl['name']}]({dl['url']})")
                    else:
                        st.info("Tidak ada file datalog diunggah pada site ini.")

                st.divider()
                st.markdown("### 🛠️ EDIT TEKS & TAMBAH FOTO BEBAS KAPAN SAJA")
                with st.container(border=True):
                    new_tek = st.text_input("Perbaiki Teknisi", r.get('teknisi',''), key=f"et_{i}")
                    new_act = st.text_area("Perbaiki Action", r.get('action',''), key=f"ea_{i}")
                    new_sp = st.text_input("Perbaiki Sparepart", r.get('sparepart',''), key=f"es_{i}")
                    
                    st.markdown("**Tambah Foto Susulan (Bebas Upload Foto Apapun Sesuai Kategori):**")
                    up_f = st.file_uploader("📸 Fisik / Shading", accept_multiple_files=True, key=f"uf_{i}")
                    up_p = st.file_uploader("📸 Panel Surya", accept_multiple_files=True, key=f"up_{i}")
                    up_b = st.file_uploader("📸 Baterai / Grounding", accept_multiple_files=True, key=f"ub_{i}")
                    up_e = st.file_uploader("📸 Elektrikal / Mesin", accept_multiple_files=True, key=f"ue_{i}")

                    if st.button("💾 Simpan Edit & Upload Susulan", key=f"btn_{i}", type="primary"):
                        with st.spinner("Menyimpan teks & foto..."):
                            r['teknisi'], r['action'], r['sparepart'] = new_tek, new_act, new_sp
                            if up_f: r['extras_fisik'] = r.get('extras_fisik', []) + upload_multiple_images(up_f)
                            if up_p: r['extras_panel'] = r.get('extras_panel', []) + upload_multiple_images(up_p)
                            if up_b: r['extras_baterai'] = r.get('extras_baterai', []) + upload_multiple_images(up_b)
                            if up_e: r['extras_elektrikal'] = r.get('extras_elektrikal', []) + upload_multiple_images(up_e)
                            save_db(db)
                        st.success("✅ Tersimpan! Layar akan refresh.")
                        st.rerun()
