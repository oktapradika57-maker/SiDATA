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

# Fungsi untuk upload gambar ke Cloudinary
def upload_image(file_obj, folder_name="solar_bts_healthcheck"):
    if file_obj is not None:
        try:
            # Membaca file Streamlit menjadi bytes dan mengunggahnya
            response = cloudinary.uploader.upload(file_obj.getvalue(), folder=folder_name)
            return response.get('secure_url')
        except Exception as e:
            st.error(f"Gagal upload gambar ke Cloudinary: {e}")
            return None
    return None

# -------------------------------------------------------------------------
# 2. SETUP HALAMAN & DATABASE SEMENTARA (SESSION STATE)
# -------------------------------------------------------------------------
st.set_page_config(page_title="Solar BTS Health Check", page_icon="⚡", layout="wide")

# Database sementara di RAM (hilang jika aplikasi direstart)
if 'laporan_db' not in st.session_state:
    st.session_state['laporan_db'] = []

# -------------------------------------------------------------------------
# 3. NAVIGASI SIDEBAR
# -------------------------------------------------------------------------
st.sidebar.title("Navigasi Aplikasi")
menu = st.sidebar.radio("Pilih Menu:", ["📝 Form Pengecekan", "📊 Hasil Pengecekan (Laporan)"])
st.sidebar.markdown("---")
st.sidebar.info("Gunakan menu ini untuk berpindah antara pengisian form lapangan dan pengecekan hasil eksekusi.")

# =========================================================================
# MENU 1: FORM PENGECEKAN (INPUT DATA LAPANGAN)
# =========================================================================
if menu == "📝 Form Pengecekan":
    st.title("⚡ Form Health Check Solar BTS")
    st.markdown("Isi form di bawah ini. Foto akan otomatis terunggah ke **Cloudinary** saat tombol Submit ditekan.")
    
    tabs = st.tabs([
        "1. Info Site", 
        "2. Solar Panel (Bef-Aft)", 
        "3. Panel DC & Kabel", 
        "4. SCC & Recty", 
        "5. Baterai", 
        "6. Grounding", 
        "7. Action & Submit"
    ])

    # --- TAB 1: INFORMASI SITE ---
    with tabs[0]:
        st.subheader("Informasi Umum Site")
        col1, col2 = st.columns(2)
        with col1:
            site_name = st.text_input("Nama / ID Site", placeholder="Contoh: BTS-PKY-001")
            nop_area = st.selectbox("NOP Area", ["Palangkaraya", "Pangkalan Bun", "Tarakan", "Pontianak", "Lainnya"])
            check_date = st.date_input("Tanggal Pengecekan", value=datetime.date.today())
        with col2:
            technician_name = st.text_input("Nama Teknisi", placeholder="Nama lengkap tim")
            weather = st.selectbox("Kondisi Cuaca", ["Cerah", "Berawan", "Hujan / Mendung"])

    # --- TAB 2: SOLAR PANEL (BEFORE - AFTER) ---
    with tabs[1]:
        st.subheader("Pengecekan Modul Surya & Shading")
        
        st.markdown("**A. Kondisi Lingkungan (Shadowing)**")
        shade_col1, shade_col2 = st.columns(2)
        with shade_col1:
            shading_status = st.selectbox("Apakah terdapat efek bayangan (shading)?", ["Aman (Clear area)", "Ada Sedikit Shading (Bisa dipangkas)", "Shading Kritis (Menutup panel total)"])
        with shade_col2:
            shading_photos = st.file_uploader("Upload Foto Area Shading (Bisa >1)", type=["jpg", "jpeg", "png"], accept_multiple_files=True, key="shd")

        st.divider()
        st.markdown("**B. Pengecekan Fisik & Tegangan Per Panel**")
        num_panels = st.number_input("Jumlah Panel (Umumnya 24-30)", min_value=1, max_value=60, value=24)

        panel_data = []
        for i in range(int(num_panels)):
            with st.expander(f"Panel #{i+1}", expanded=(i == 0)):
                p_col1, p_col2, p_col3 = st.columns([1, 1.2, 1.2])
                with p_col1:
                    voc = st.number_input(f"Voc Panel {i+1} [V]", min_value=0.0, value=21.5, step=0.1, key=f"v_{i}")
                    isc = st.number_input(f"Isc Panel {i+1} [A]", min_value=0.0, value=5.2, step=0.1, key=f"i_{i}")
                    p_cond = st.selectbox(f"Kondisi {i+1}", ["Baik", "Kotor/Soiling", "Retak", "Delaminasi"], key=f"c_{i}")
                with p_col2:
                    p_photo_before = st.file_uploader(f"📸 Foto BEFORE (Kotor) - Panel {i+1}", type=["jpg", "jpeg", "png"], key=f"pb_{i}")
                with p_col3:
                    p_photo_after = st.file_uploader(f"✨ Foto AFTER (Bersih) - Panel {i+1}", type=["jpg", "jpeg", "png"], key=f"pa_{i}")

                panel_data.append({
                    "id": f"Panel #{i+1}", "voc": voc, "isc": isc, "kondisi": p_cond,
                    "foto_before_obj": p_photo_before, "foto_after_obj": p_photo_after
                })

    # --- TAB 3: PANEL DC & KABEL ---
    with tabs[2]:
        st.subheader("Panel DC, Perkabelan & Kontaktor")
        jb_col1, jb_col2 = st.columns(2)
        with jb_col1:
            jb_enclosure = st.selectbox("Kondisi Box DC & Seal Karet", ["Baik & Kedap Air", "Rusak / Seal Lepas"])
            jb_spd = st.selectbox("Status Arrester / SPD", ["Normal (Hijau)", "Rusak (Merah/Hitam)"])
            cabling = st.selectbox("Kondisi Perkabelan & Kontaktor", ["Rapi & Aman", "Oksidasi / Chattering", "Gosong/Berantakan"])
        with jb_col2:
            jb_photos = st.file_uploader("Upload Foto Panel DC & Kabel (Bisa >1)", type=["jpg", "jpeg", "png"], accept_multiple_files=True, key="jb")

    # --- TAB 4: SCC, RECTIFIER & DATALOG ---
    with tabs[3]:
        st.subheader("SCC / Rectifier & Datalog")
        scc_col1, scc_col2 = st.columns(2)
        with scc_col1:
            scc_brand = st.text_input("Merek SCC / Rectifier", placeholder="Contoh: SmartGen, Eltek")
            system_out_v = st.number_input("Tegangan Output ke BTS [Volt]", value=48.0)
            is_datalog_taken = st.checkbox("✅ Datalog berhasil ditarik / di-download oleh tim")
            datalog_file = st.file_uploader("Upload File Datalog (CSV/Excel) - Jika ada", type=["csv", "xlsx", "xls", "txt"], key="dl")
        with scc_col2:
            scc_photos = st.file_uploader("Upload Foto SCC / Recty (Bisa >1)", type=["jpg", "jpeg", "png"], accept_multiple_files=True, key="scc")

    # --- TAB 5: BATERAI ---
    with tabs[4]:
        st.subheader("Pemeriksaan Baterai Per Unit")
        num_batteries = st.number_input("Jumlah Blok Baterai (Maks 11)", min_value=1, max_value=11, value=4)
        bat_total_v = st.number_input("Total Tegangan Keseluruhan Bank Baterai [V]", value=52.8)
        
        bat_data = []
        for j in range(int(num_batteries)):
            with st.expander(f"Unit Baterai #{j+1}", expanded=(j==0)):
                b_col1, b_col2 = st.columns(2)
                with b_col1:
                    b_volt = st.number_input(f"Tegangan Baterai {j+1} [V]", value=12.2, step=0.1, key=f"bv_{j}")
                    b_cond = st.selectbox(f"Kondisi Fisik Baterai {j+1}", ["Normal / Mulus", "Bengkak", "Bocor / Korosi"], key=f"bc_{j}")
                with b_col2:
                    b_photo = st.file_uploader(f"Upload Foto Baterai {j+1}", type=["jpg", "jpeg", "png"], key=f"bp_{j}")
                
                bat_data.append({
                    "id": f"Baterai #{j+1}", "voltase": b_volt, "kondisi": b_cond, "foto_obj": b_photo
                })

    # --- TAB 6: GROUNDING ---
    with tabs[5]:
        st.subheader("Grounding & Lightning Protection")
        grd_col1, grd_col2 = st.columns(2)
        with grd_col1:
            earth_resistance = st.number_input("Tahanan Pentanahan (Earth Tester) [Ohm]", value=2.1, step=0.1)
        with grd_col2:
            grd_photos = st.file_uploader("Upload Foto Grounding", type=["jpg", "jpeg", "png"], accept_multiple_files=True, key="grd")

    # --- TAB 7: ACTION & SUBMIT ---
    with tabs[6]:
        st.subheader("Ringkasan & Action Eksekusi")
        action_taken = st.text_area("Ketik Mandiri Action/Eksekusi yang Telah Dilakukan di Site:", height=100, placeholder="Misal: Membersihkan karat panel DC, mengencangkan kabel, membersihkan modul panel kotor...")
        final_status = st.radio("Status Akhir Site:", ["Normal (Siap Optimal)", "Minor Issue (Sudah Ditangani)", "Major / Critical (Butuh Penggantian Part)"])

        if st.button("🚀 Upload & Submit Laporan", type="primary"):
            if not site_name or not technician_name:
                st.error("⚠️ Mohon isi Nama Site dan Nama Teknisi di Tab 1 terlebih dahulu!")
            else:
                with st.spinner("⏳ Sedang memproses dan mengunggah foto ke Cloudinary... Mohon tunggu jangan tutup halaman."):
                    
                    # 1. Upload Foto Panel (Before - After)
                    panel_results = []
                    for p in panel_data:
                        url_bef = upload_image(p["foto_before_obj"], "solar_panel_before") if p["foto_before_obj"] else None
                        url_aft = upload_image(p["foto_after_obj"], "solar_panel_after") if p["foto_after_obj"] else None
                        panel_results.append({
                            "Panel": p["id"], "Voc": p["voc"], "Isc": p["isc"], 
                            "Kondisi": p["kondisi"], "URL_Before": url_bef, "URL_After": url_aft
                        })
                    
                    # 2. Upload Foto Baterai
                    bat_results = []
                    for b in bat_data:
                        url_bat = upload_image(b["foto_obj"], "solar_battery") if b["foto_obj"] else None
                        bat_results.append({
                            "Baterai": b["id"], "Voltase": b["voltase"], 
                            "Kondisi": b["kondisi"], "URL_Foto": url_bat
                        })

                    # 3. Kumpulkan semua data ke dalam Dictionary
                    report_dict = {
                        "timestamp": datetime.datetime.now().strftime("%Y-%m-%d %H:%M:%S"),
                        "site_name": site_name,
                        "nop": nop_area,
                        "teknisi": technician_name,
                        "status": final_status,
                        "action": action_taken,
                        "shading_status": shading_status,
                        "scc_brand": scc_brand,
                        "datalog_status": "Sudah Ditarik" if is_datalog_taken else "Belum Ditarik",
                        "panel_data": panel_results,
                        "battery_data": bat_results
                    }
                    
                    # 4. Simpan ke Database Session
                    st.session_state['laporan_db'].append(report_dict)
                
                st.success(f"✅ Laporan Site {site_name} berhasil diunggah! Silakan cek di menu 'Hasil Pengecekan'.")

# =========================================================================
# MENU 2: HASIL PENGECEKAN (DASHBOARD LAPORAN)
# =========================================================================
elif menu == "📊 Hasil Pengecekan (Laporan)":
    st.title("📊 Hasil Pengecekan & Eksekusi Site")
    st.markdown("Berikut adalah daftar site yang telah disubmit dan foto yang berhasil tersimpan di Cloudinary.")
    st.divider()

    db = st.session_state['laporan_db']

    if len(db) == 0:
        st.info("Belum ada data laporan yang masuk. Silakan isi form di menu 'Form Pengecekan' terlebih dahulu.")
    else:
        # Tampilkan laporan dari yang terbaru (reverse)
        for index, report in enumerate(reversed(db)):
            with st.expander(f"📍 {report['site_name']} - {report['nop']} | {report['timestamp']} | Status: {report['status']}"):
                
                # Header Info
                st.markdown(f"**Tim Eksekutor:** {report['teknisi']} | **SCC/Recty:** {report['scc_brand']} | **Status Datalog:** {report['datalog_status']}")
                st.markdown(f"**Status Lingkungan (Shading):** {report['shading_status']}")
                
                st.markdown("**Tindakan/Action yang dilakukan:**")
                st.info(report['action'] if report['action'] else "Tidak ada catatan action.")
                
                # Tab rincian data
                tab_p, tab_b = st.tabs(["Data Panel Surya (Before-After)", "Data Baterai"])
                
                # --- SUB TAB: DATA PANEL ---
                with tab_p:
                    for p in report['panel_data']:
                        st.markdown(f"#### {p['Panel']}")
                        st.markdown(f"Voc: **{p['Voc']} V** | Isc: **{p['Isc']} A** | Kondisi: **{p['Kondisi']}**")
                        
                        # Layout Foto
                        col_b, col_a = st.columns(2)
                        with col_b:
                            st.caption("📷 BEFORE (Kotor)")
                            if p['URL_Before']:
                                st.image(p['URL_Before'], use_container_width=True)
                            else:
                                st.write("*(Tidak ada foto before)*")
                        with col_a:
                            st.caption("✨ AFTER (Bersih)")
                            if p['URL_After']:
                                st.image(p['URL_After'], use_container_width=True)
                            else:
                                st.write("*(Tidak ada foto after)*")
                        st.divider()

                # --- SUB TAB: DATA BATERAI ---
                with tab_b:
                    for b in report['battery_data']:
                        st.markdown(f"#### {b['Baterai']}")
                        st.markdown(f"Voltase: **{b['Voltase']} V** | Fisik: **{b['Kondisi']}**")
                        if b['URL_Foto']:
                            st.image(b['URL_Foto'], width=300)
                        else:
                            st.write("*(Tidak ada foto baterai)*")
                        st.divider()
