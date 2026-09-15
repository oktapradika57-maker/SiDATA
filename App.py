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
        except Exception: return None
    return None

def upload_multiple_images(file_objs, folder_name="solar_bts_healthcheck"):
    urls = []
    if file_objs:
        for file in file_objs:
            url = upload_image(file, folder_name)
            if url: urls.append(url)
    return urls

def tampilkan_grid_foto(url_data, caption=""):
    if not url_data: return
    urls = []
    if isinstance(url_data, str) and url_data.startswith("http"): urls = [url_data]
    elif isinstance(url_data, list): urls = [u for u in url_data if isinstance(u, str) and u.startswith("http")]
        
    if urls:
        if caption: st.markdown(f"*{caption}*")
        
        # Trik HTML ditulis 1 baris lurus agar Streamlit tidak menciptakan spasi kosong
        img_html = '<div style="display: grid; grid-template-columns: repeat(auto-fit, minmax(150px, 1fr)); gap: 10px; margin-bottom: 15px;">'
        for u in urls:
            img_html += f'<a href="{u}" target="_blank"><img src="{u}" style="width: 100%; height: 150px; object-fit: cover; border-radius: 8px; box-shadow: 0px 4px 6px rgba(0,0,0,0.1);"></a>'
        img_html += '</div>'
        
        st.markdown(img_html, unsafe_allow_html=True)
# -------------------------------------------------------------------------
# 2. SETUP DATABASE: SINKRONISASI KE GOOGLE SHEETS
# -------------------------------------------------------------------------
st.set_page_config(page_title="Preventive Maintenance BTS", page_icon="⚡", layout="wide")

SHEET_ID = "1HvgVicTWwO4RMQI6ZR3Mu3IgGicwjcLZl9mDN1auvJU"
SHEET_NAME = "Report Preventive"

def connect_gsheets():
    try:
        # SUDAH DISESUAIKAN DENGAN NAMA VARIABEL SECRETS ANDA
        creds_json = st.secrets["gcp_json"]
        creds_dict = json.loads(creds_json)
        scopes = ["https://www.googleapis.com/auth/spreadsheets", "https://www.googleapis.com/auth/drive"]
        creds = Credentials.from_service_account_info(creds_dict, scopes=scopes)
        client = gspread.authorize(creds)
        return client.open_by_key(SHEET_ID).worksheet(SHEET_NAME)
    except Exception as e:
        st.error(f"⚠️ Koneksi Spreadsheet Gagal! Detail Error: {e}")
        return None

@st.cache_data(ttl=600) 
def fetch_data_from_gsheets():
    sheet = connect_gsheets()
    if sheet:
        data = sheet.get_all_values()
        if len(data) > 1: 
            db = []
            for row in data[1:]:
                if len(row) >= 6: 
                    try:
                        db.append(json.loads(row[5]))
                    except: pass
            return db
    return []

if 'laporan_db' not in st.session_state:
    st.session_state['laporan_db'] = fetch_data_from_gsheets()

# -------------------------------------------------------------------------
# 3. NAVIGASI UTAMA
# -------------------------------------------------------------------------
st.sidebar.title("Navigasi Operasional")
st.sidebar.info(f"Database Taut:\n`Report Preventive`\n🔗 [Buka Google Sheets](https://docs.google.com/spreadsheets/d/{SHEET_ID}/edit)")
menu = st.sidebar.radio("Pilih Menu:", ["📝 Form Preventive Check", "📊 Hasil Laporan & Dashboard"])
st.sidebar.markdown("---")

# =========================================================================
# MENU 1: FORM PENGECEKAN LENGKAP
# =========================================================================
if menu == "📝 Form Preventive Check":
    st.title("⚡ Form Preventive Maintenance Komprehensif")
    st.info("Data dan Foto Anda akan LANGSUNG di-tembak masuk ke Google Sheets saat Submit.")
    
    tabs = st.tabs(["1. Info Site", "2. SPS Panel", "3. PLN & Rectifier", "4. Genset & BBM", "5. Baterai & Grounding", "6. Datalog & Submit"])

    # TAB 1: INFO SITE
    with tabs[0]:
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
            site_condition = st.selectbox("Kondisi Halaman", ["Bersih & Aman", "Banyak Rumput", "Genangan Air"])
            tower_condition = st.selectbox("Fisik Tower", ["Aman", "Berkarat", "Rusak"])
        with sc2:
            site_photos = st.file_uploader("Foto View Site (Bisa >1)", accept_multiple_files=True, key="spics")

    # TAB 2: SPS
    with tabs[1]:
        shading_status = st.selectbox("Status Shading?", ["Aman", "Sedikit Shading", "Kritis"])
        shading_photos = st.file_uploader("Foto Shading (Bisa >1)", accept_multiple_files=True, key="shd")
        st.divider()
        
        konfigurasi_panel = st.radio("Metode Pengukuran SPS:", ["Individu (Satuan)", "Seri / String (Grup)"])
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
                with st.expander(f"String Seri #{i+1}", expanded=(i==0)):
                    c1, c2 = st.columns([1, 1.5])
                    with c1:
                        qty = st.number_input(f"Isi Panel per Seri", value=6, key=f"sq_{i}")
                        voc = st.number_input(f"Total Voc [V]", value=129.0, key=f"sv_{i}")
                        isc = st.number_input(f"Isc [A]", value=5.2, key=f"si_{i}")
                        p_cond = st.selectbox(f"Kondisi", ["Baik", "Kotor", "Retak"], key=f"sc_{i}")
                    with c2:
                        pb = st.file_uploader(f"📸 BEFORE (Bisa >1)", accept_multiple_files=True, key=f"spb_{i}")
                        pa = st.file_uploader(f"✨ AFTER (Bisa >1)", accept_multiple_files=True, key=f"spa_{i}")
                    panel_data.append({"tipe": "Seri", "id": f"String #{i+1}", "qty": qty, "voc": voc, "isc": isc, "kondisi": p_cond, "foto_before_objs": pb, "foto_after_objs": pa})

    # TAB 3: PLN & RECTI
    with tabs[2]:
        c1, c2 = st.columns(2)
        with c1:
            pln_status = st.selectbox("Status PLN", ["Normal", "Padam", "Tidak Ada PLN"])
            jb_enclosure = st.selectbox("Box DC & Seal", ["Bersih", "Bocor", "Sarang Serangga"])
            jb_spd = st.selectbox("Arrester SPD", ["Normal", "Rusak"])
        with c2:
            rect_brand = st.text_input("Merek Rectifier", placeholder="Huawei / ZTE")
            rect_out_v = st.number_input("Tegangan Output Rectifier [V]", value=53.5)
            total_load_a = st.number_input("Total Beban BTS [A]", value=18.2)
            rect_alarm = st.selectbox("Alarm Rectifier", ["No Alarm", "Ada Alarm"])
            rect_photos = st.file_uploader("Foto Rectifier (Bisa >1)", accept_multiple_files=True, key="rect_pics")

    # TAB 4: GENSET & BBM
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
            genset_photos = st.file_uploader("Foto Genset & BBM (Bisa >1)", accept_multiple_files=True, key="gen_pics")

    # TAB 5: BATERAI & GROUNDING
    with tabs[4]:
        b1, b2 = st.columns(2)
        with b1:
            num_bat = st.number_input("Jumlah Baterai", min_value=1, value=4)
            bat_data = []
            for j in range(int(num_bat)):
                with st.expander(f"Baterai #{j+1}", expanded=(j==0)):
                    bc1, bc2 = st.columns(2)
                    with bc1:
                        bv = st.number_input(f"Voltase [V]", value=12.2, key=f"bv_{j}")
                        bt = st.number_input(f"Suhu [°C]", value=28.0, key=f"bt_{j}")
                    with bc2:
                        bc = st.selectbox(f"Kondisi", ["Normal", "Bengkak", "Korosi"], key=f"bc_{j}")
                        bp = st.file_uploader(f"Foto Bat #{j+1}", accept_multiple_files=True, key=f"bp_{j}")
                    bat_data.append({"id": f"Baterai #{j+1}", "voltase": bv, "suhu": bt, "kondisi": bc, "foto_objs": bp})
        with b2:
            earth_resistance = st.number_input("Tahanan Grounding [Ohm]", value=2.1)
            grd_cable = st.selectbox("Kabel Grounding", ["Kuat", "Kendor", "Putus"])
            grd_photos = st.file_uploader("Foto Grounding (Bisa >1)", accept_multiple_files=True, key="grd")

    # TAB 6: SUBMIT
    with tabs[5]:
        st.markdown("📂 **Upload Datalog (Semua Format):**")
        datalog_files = st.file_uploader("Upload .csv, .xlsx, .txt, dll", accept_multiple_files=True, key="datalog_all")
        st.divider()
        action_taken = st.text_area("🔧 Action di Lapangan:")
        sparepart_needed = st.text_input("📦 Sparepart Diganti:")
        final_status = st.radio("Status Site:", ["Normal", "Minor Issue", "Major/Critical"])

        if st.button("🚀 Upload & Sinkronkan ke Spreadsheet", type="primary"):
            if not site_name:
                st.error("⚠️ Mohon isi Nama Site!")
            else:
                with st.spinner("⏳ Mengupload ke Cloudinary dan mengirim data ke Spreadsheet... Mohon tunggu."):
                    url_sites = upload_multiple_images(site_photos, "prev_view")
                    url_shadings = upload_multiple_images(shading_photos, "prev_shade")
                    url_rects = upload_multiple_images(rect_photos, "prev_rect")
                    url_gensets = upload_multiple_images(genset_photos, "prev_genset")
                    url_grds = upload_multiple_images(grd_photos, "prev_grd")
                    
                    datalog_urls = []
                    if datalog_files:
                        for df in datalog_files:
                            du = upload_image(df, "prev_datalog")
                            if du: datalog_urls.append({"name": df.name, "url": du})

                    p_res = []
                    for p in panel_data:
                        if p.get("tipe") == "Individu":
                            p_res.append({"Tipe": "Individu", "Panel": p["id"], "Voc": p["voc"], "Isc": p["isc"], "Kondisi": p["kondisi"], "URL_Before": upload_image(p["foto_before_obj"]), "URL_After": upload_image(p["foto_after_obj"])})
                        else:
                            p_res.append({"Tipe": "Seri", "Panel": p["id"], "Qty": p["qty"], "Voc": p["voc"], "Isc": p["isc"], "Kondisi": p["kondisi"], "URLs_Before": upload_multiple_images(p["foto_before_objs"]), "URLs_After": upload_multiple_images(p["foto_after_objs"])})
                    
                    b_res = []
                    for b in bat_data: b_res.append({"Baterai": b["id"], "Voltase": b["voltase"], "Suhu": b["suhu"], "Kondisi": b["kondisi"], "URL_Fotos": upload_multiple_images(b["foto_objs"])})

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
                        "tank_capacity": tank_capacity, "fuel_pct": current_fuel_pct, "fuel_liter": est_fuel,
                        "url_gensets": url_gensets,
                        "battery_data": b_res, "earth_ohm": earth_resistance, "url_grds": url_grds,
                        "datalog_files": datalog_urls,
                        "extras_fisik": [], "extras_panel": [], "extras_baterai": [], "extras_elektrikal": []
                    }
                    
                    # PROSES PENYIMPANAN LANGSUNG KE SPREADSHEET
                    sheet = connect_gsheets()
                    if sheet:
                        row_data = [
                            report_dict.get('timestamp', ''), report_dict.get('site_name', ''),
                            report_dict.get('nop', ''), report_dict.get('teknisi', ''),
                            report_dict.get('status', ''), json.dumps(report_dict)
                        ]
                        sheet.append_row(row_data)
                        st.cache_data.clear() # Reset cache laporan
                        st.session_state['laporan_db'].append(report_dict)
                        st.success("✅ BERHASIL! Data telah masuk ke Google Sheets permanen.")
                    else:
                        st.error("⚠️ Gagal menyambung ke Spreadsheet, mohon periksa Setting Secret API Anda.")

# =========================================================================
# MENU 2: HASIL LAPORAN (DITARIK DARI SPREADSHEET)
# =========================================================================
elif menu == "📊 Hasil Laporan & Dashboard":
    st.title("📊 Dashboard Laporan Terpusat (Spreadsheet)")
    st.divider()

    db = st.session_state['laporan_db']

    if not db:
        st.info("⚠️ Belum ada data di Spreadsheet / Koneksi Gagal.")
    else:
        for i in range(len(db) - 1, -1, -1):
            r = db[i]
            wa_text = f"""*REPORT PREVENTIVE* ⚡\n📍 *Site:* {r.get('site_name', '-')} ({r.get('nop', '-')})\n📅 *Tanggal:* {r.get('timestamp', '-')}\n👷 *Teknisi:* {r.get('teknisi', '-')}\n📊 *Status:* {r.get('status', '-')}\n\n*POWER & LOAD:*\n- PLN: {r.get('pln_status', '-')}\n- Rectifier: {r.get('rect_brand', '-')} ({r.get('rect_out_v', '-')}V)\n- Load BTS: {r.get('total_load', '-')} A\n\n*TINDAKAN:*\n{r.get('action', '-')}\n\n*SPAREPART:*\n{r.get('sparepart', '-')}"""
            wa_url = f"https://wa.me/?text={urllib.parse.quote(wa_text)}"

            with st.expander(f"📍 {r.get('site_name', 'Unknown')} | {r.get('timestamp', '')} | Status: {r.get('status', '')}"):
                
                c_btn1, c_btn2 = st.columns(2)
                with c_btn1: st.link_button("📱 Kirim Rangkuman ke WhatsApp", wa_url)
                with c_btn2: st.link_button("📈 Buka Spreadsheet Target", f"https://docs.google.com/spreadsheets/d/{SHEET_ID}/edit")
                
                st.markdown(f"**Teknisi:** {r.get('teknisi', '-')} | **Load:** {r.get('total_load', '-')} A")
                st.markdown(f"**Action:** {r.get('action', '-')}")
                if r.get('sparepart'): st.warning(f"**Sparepart:** {r['sparepart']}")
                
                ltab1, ltab2, ltab3, ltab4, ltab5 = st.tabs(["1. Fisik", "2. Panel SPS", "3. Recti & Genset", "4. Baterai & Gnd", "5. File"])
                
                with ltab1:
                    st.write(f"- Kondisi Site: {r.get('site_cond', '-')} | Tower: {r.get('tower_cond', '-')} | Shading: {r.get('shading_status', '-')}")
                    tampilkan_grid_foto(r.get('url_sites'), "📸 View Site")
                    tampilkan_grid_foto(r.get('url_shadings'), "📸 Shading")
                    tampilkan_grid_foto(r.get('extras_fisik'), "📸 Tambahan")

                with ltab2:
                    for p in r.get('panel_data', []):
                        st.markdown(f"**{p.get('Panel', '-')}** | Voc: {p.get('Voc','-')}V | Isc: {p.get('Isc','-')}A")
                        cb, ca = st.columns(2)
                        with cb: tampilkan_grid_foto(p.get('URL_Before') or p.get('URLs_Before'), "📸 BEFORE")
                        with ca: tampilkan_grid_foto(p.get('URL_After') or p.get('URLs_After'), "✨ AFTER")
                        st.divider()
                    tampilkan_grid_foto(r.get('extras_panel'), "📸 Tambahan Panel")

                with ltab3:
                    c_p1, c_p2 = st.columns(2)
                    with c_p1:
                        st.write(f"**PLN:** {r.get('pln_status', '-')} | **Rectifier:** {r.get('rect_brand', '-')} ({r.get('rect_out_v', '-')}V)")
                        tampilkan_grid_foto(r.get('url_rects'), "📸 Rectifier")
                    with c_p2:
                        st.write(f"**Genset:** {r.get('genset_status', '-')} | **BBM:** {r.get('fuel_pct', '-')}%")
                        tampilkan_grid_foto(r.get('url_gensets'), "📸 Genset")
                    st.divider()
                    tampilkan_grid_foto(r.get('extras_elektrikal'), "📸 Tambahan Mesin/Elektrikal")

                with ltab4:
                    for b in r.get('battery_data', []):
                        st.write(f"**{b.get('Baterai', '-')}** | {b.get('Voltase','-')}V | Suhu: {b.get('Suhu','-')}°C")
                        tampilkan_grid_foto(b.get('URL_Fotos'))
                        st.divider()
                    st.write(f"**Grounding:** {r.get('earth_ohm', '-')} Ohm")
                    tampilkan_grid_foto(r.get('url_grds'), "📸 Grounding")
                    tampilkan_grid_foto(r.get('extras_baterai'), "📸 Tambahan Baterai")

                with ltab5:
                    dls = r.get('datalog_files', [])
                    for dl in dls: st.markdown(f"- [{dl['name']}]({dl['url']})")

                st.divider()
                st.markdown("### 🛠️ EDIT TEKS & TAMBAH FOTO BEBAS")
                with st.container(border=True):
                    new_tek = st.text_input("Edit Teknisi", r.get('teknisi',''), key=f"et_{i}")
                    new_act = st.text_area("Edit Action", r.get('action',''), key=f"ea_{i}")
                    new_sp = st.text_input("Edit Sparepart", r.get('sparepart',''), key=f"es_{i}")
                    
                    st.markdown("**Tambah Foto Susulan Bebas:**")
                    up_f = st.file_uploader("Fisik / Shading", accept_multiple_files=True, key=f"uf_{i}")
                    up_p = st.file_uploader("Panel Surya", accept_multiple_files=True, key=f"up_{i}")
                    up_b = st.file_uploader("Baterai / Grounding", accept_multiple_files=True, key=f"ub_{i}")
                    up_e = st.file_uploader("Elektrikal / Mesin", accept_multiple_files=True, key=f"ue_{i}")

                    if st.button("💾 Simpan Edit ke Spreadsheet", key=f"btn_{i}", type="primary"):
                        with st.spinner("Sinkronisasi Update ke Spreadsheet..."):
                            r['teknisi'], r['action'], r['sparepart'] = new_tek, new_act, new_sp
                            if up_f: r['extras_fisik'] = r.get('extras_fisik', []) + upload_multiple_images(up_f)
                            if up_p: r['extras_panel'] = r.get('extras_panel', []) + upload_multiple_images(up_p)
                            if up_b: r['extras_baterai'] = r.get('extras_baterai', []) + upload_multiple_images(up_b)
                            if up_e: r['extras_elektrikal'] = r.get('extras_elektrikal', []) + upload_multiple_images(up_e)
                            
                            # Update ke Google Sheets
                            sheet = connect_gsheets()
                            if sheet:
                                row_num = i + 2 # Google Sheet index mulai dari 2 (1 itu Header)
                                sheet.update_cell(row_num, 4, new_tek)
                                sheet.update_cell(row_num, 5, r.get('status', ''))
                                sheet.update_cell(row_num, 6, json.dumps(r))
                                st.cache_data.clear() # Refresh Cache
                        st.success("✅ Perubahan Teks & Foto tersimpan permanen di Google Sheets!")
                        st.rerun()
