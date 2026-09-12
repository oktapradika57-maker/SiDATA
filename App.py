import datetime
from PIL import Image
import pandas as pd
import streamlit as st

# Konfigurasi Halaman
st.set_page_config(
    page_title="Solar BTS Health Check System",
    page_icon="⚡",
    layout="wide",
)

st.title("⚡ Preventive Maintenance & Health Check Solar BTS")
st.markdown("Formulir pemeriksaan berkala sistem PLTS dan Catu Daya BTS Telekomunikasi.")
st.markdown("---")

# Navigasi Form
tabs = st.tabs([
    "1. Info Site", 
    "2. Solar Panel", 
    "3. Panel DC & Kabel", 
    "4. SCC & Rectifier", 
    "5. Baterai", 
    "6. Grounding", 
    "7. Action & Submit"
])

# -------------------------------------------------------------------------
# TAB 1: INFORMASI SITE
# -------------------------------------------------------------------------
with tabs[0]:
    st.subheader("Informasi Umum Site")
    col1, col2 = st.columns(2)

    with col1:
        site_name = st.text_input("Nama / ID Site", placeholder="Contoh: BTS-PKY-001")
        nop_area = st.selectbox("Network Operation Point (NOP)", ["Palangkaraya", "Pangkalan Bun", "Tarakan", "Pontianak", "Lainnya"])
        check_date = st.date_input("Tanggal Pengecekan", value=datetime.date.today())

    with col2:
        technician_name = st.text_input("Nama Teknisi", placeholder="Nama lengkap teknisi")
        weather = st.selectbox("Kondisi Cuaca", ["Cerah", "Berawan", "Hujan / Mendung"])

# -------------------------------------------------------------------------
# TAB 2: SOLAR PANEL & SHADOWING
# -------------------------------------------------------------------------
with tabs[1]:
    st.subheader("Pemeriksaan Modul Panel Surya & Lingkungan")
    
    # Cek Shadowing (Shading)
    st.markdown("**A. Kondisi Lingkungan (Shadowing/Shading)**")
    shade_col1, shade_col2 = st.columns(2)
    with shade_col1:
        shading_status = st.selectbox("Apakah terdapat efek bayangan (pohon/bangunan) pada area panel?", 
                                      ["Aman (Clear area)", "Ada Sedikit Shading (Bisa dipangkas)", "Shading Kritis (Menutup panel total)"])
    with shade_col2:
        shading_photos = st.file_uploader("Upload Foto Area (Bisa pilih >1 foto)", type=["jpg", "jpeg", "png"], accept_multiple_files=True, key="shading_photos")
    
    st.divider()

    # Cek Per Panel
    st.markdown("**B. Pengecekan Fisik & Tegangan Per Panel**")
    num_panels = st.number_input("Jumlah Panel yang Diperiksa (Umumnya 24-30)", min_value=1, max_value=60, value=24, step=1)

    panel_data = []
    for i in range(int(num_panels)):
        with st.expander(f"Panel #{i+1}", expanded=(i == 0)):
            p_col1, p_col2 = st.columns(2)

            with p_col1:
                voc = st.number_input(f"Voc - Panel {i+1} [Volt]", min_value=0.0, value=21.5, step=0.1, key=f"voc_{i}")
                isc = st.number_input(f"Isc - Panel {i+1} [Ampere]", min_value=0.0, value=5.2, step=0.1, key=f"isc_{i}")
                p_condition = st.selectbox(f"Kondisi Fisik - Panel {i+1}", ["Baik", "Retak / Pecah", "Kotor (Soiling)", "Delaminasi"], key=f"cond_{i}")

            with p_col2:
                p_photo = st.file_uploader(f"Upload Foto Panel {i+1}", type=["jpg", "jpeg", "png"], key=f"photo_panel_{i}")
                if p_photo is not None:
                    st.image(p_photo, caption=f"Panel {i+1}", width=200)

            panel_data.append({"Panel_ID": f"Panel #{i+1}", "Voc": voc, "Isc": isc, "Kondisi": p_condition})

# -------------------------------------------------------------------------
# TAB 3: JUNCTION BOX, KABEL & KONTAKTOR
# -------------------------------------------------------------------------
with tabs[2]:
    st.subheader("Panel DC, Perkabelan & Kontaktor")
    
    jb_col1, jb_col2 = st.columns(2)
    with jb_col1:
        jb_enclosure = st.selectbox("Kondisi Box DC & Seal Karet", ["Baik & Kedap Air", "Rusak / Seal Lepas"])
        jb_fuse = st.selectbox("Status Breaker / MCB / Fuse", ["Normal / Aman", "Trip / Putus"])
        jb_spd = st.selectbox("Status Arrester / SPD", ["Normal (Hijau)", "Rusak (Merah/Hitam)"])
        
        st.markdown("**Perkabelan & Kontaktor**")
        cabling = st.selectbox("Kondisi Perkabelan Utama", ["Rapi & Aman", "Terkelupas / Berantakan", "Oksidasi / Gosong"])
        contactor = st.selectbox("Fungsi & Kondisi Kontaktor", ["Normal / Tidak Berisik", "Dengung / Chattering", "Kontak Gosong / Lengket", "Tidak Ada Kontaktor"])

    with jb_col2:
        jb_photos = st.file_uploader("Upload Foto Panel DC & Kabel (Bisa pilih >1 foto)", type=["jpg", "jpeg", "png"], accept_multiple_files=True, key="jb_photos")
        if jb_photos:
            for pic in jb_photos:
                st.image(pic, width=150)

# -------------------------------------------------------------------------
# TAB 4: SCC & RECTIFIER + DATALOG
# -------------------------------------------------------------------------
with tabs[3]:
    st.subheader("SCC / Rectifier & Pengambilan Data")
    
    scc_col1, scc_col2 = st.columns(2)
    with scc_col1:
        scc_brand = st.text_input("Merek / Tipe SCC atau Rectifier", placeholder="Contoh: SmartGen, Eltek, Huawei, dll.")
        system_out_v = st.number_input("Tegangan Output ke BTS [Volt]", min_value=0.0, value=48.0, step=0.1)
        scc_alarm = st.selectbox("Status Indikator & Alarm", ["Normal (No Alarm)", "Ada Alarm Fault / Error"])
        alarm_code = st.text_input("Kode Alarm (Jika ada)", placeholder="Contoh: Overvoltage, Err-03")
        
        st.markdown("**Pengecekan Datalog**")
        is_datalog_taken = st.checkbox("✅ Datalog berhasil ditarik / di-download oleh tim")
        datalog_file = st.file_uploader("Upload File Datalog (CSV/TXT/Excel)", type=["csv", "txt", "xlsx", "xls"], key="datalog_file")

    with scc_col2:
        scc_photos = st.file_uploader("Upload Foto SCC / Display Recty (Bisa >1 foto)", type=["jpg", "jpeg", "png"], accept_multiple_files=True, key="scc_photos")
        if scc_photos:
            for pic in scc_photos:
                st.image(pic, width=150)

# -------------------------------------------------------------------------
# TAB 5: BANK BATERAI (Maks 11 Unit)
# -------------------------------------------------------------------------
with tabs[4]:
    st.subheader("Pemeriksaan Baterai Per Unit")
    st.markdown("Masukkan jumlah baterai yang dicek (Maksimal 11 unit/foto).")

    num_batteries = st.number_input("Jumlah Unit/Blok Baterai yang dicek", min_value=1, max_value=11, value=4, step=1)
    
    bat_total_v = st.number_input("Total Tegangan Keseluruhan Bank Baterai [Volt]", min_value=0.0, value=52.8, step=0.1)
    bat_temp = st.number_input("Suhu Ruang Baterai [°C]", min_value=15.0, value=28.5)

    bat_data = []
    for j in range(int(num_batteries)):
        with st.expander(f"Unit Baterai #{j+1}", expanded=(j == 0)):
            b_col1, b_col2 = st.columns(2)
            
            with b_col1:
                b_volt = st.number_input(f"Tegangan Baterai {j+1} [V]", min_value=0.0, value=12.2, step=0.1, key=f"b_volt_{j}")
                b_cond = st.selectbox(f"Kondisi Fisik Baterai {j+1}", ["Normal / Mulus", "Bengkak (Swelling)", "Bocor", "Terminal Korosi"], key=f"b_cond_{j}")
            
            with b_col2:
                b_photo = st.file_uploader(f"Upload Foto Baterai {j+1}", type=["jpg", "jpeg", "png"], key=f"b_photo_{j}")
                if b_photo:
                    st.image(b_photo, width=150)
            
            bat_data.append({"Baterai_ID": f"Unit #{j+1}", "Voltase": b_volt, "Kondisi": b_cond})

# -------------------------------------------------------------------------
# TAB 6: GROUNDING
# -------------------------------------------------------------------------
with tabs[5]:
    st.subheader("Grounding & Lightning Protection")
    grd_col1, grd_col2 = st.columns(2)
    with grd_col1:
        earth_resistance = st.number_input("Tahanan Pentanahan (Earth Tester) [Ohm]", min_value=0.0, value=2.1, step=0.1)
        grd_cable = st.selectbox("Kondisi Kabel & Busbar Grounding", ["Terhubung Kuat", "Kendor / Berkarat", "Putus"])
    with grd_col2:
        grd_photos = st.file_uploader("Upload Foto Grounding", type=["jpg", "jpeg", "png"], accept_multiple_files=True, key="grd_photos")

# -------------------------------------------------------------------------
# TAB 7: ACTION & SUBMIT
# -------------------------------------------------------------------------
with tabs[6]:
    st.subheader("Ringkasan, Action & Submit")

    st.markdown("**Action / Eksekusi yang Telah Dilakukan di Site:**")
    action_taken = st.text_area(
        "Ketik mandiri detail tindakan (Misal: Pengencangan terminal SCC, pembersihan karat panel DC, tarik datalog).", 
        height=100
    )
    
    sparepart_needed = st.text_input("Daftar Sparepart yang Dibutuhkan / Diganti (Jika ada)")
    
    final_status = st.radio("Status Akhir Site:", [
        "Normal (Siap Beroperasi Optimal)",
        "Minor Issue (Sudah Ditangani / Butuh Pantauan)",
        "Major / Critical (Butuh Eskalasi / Penggantian Part)"
    ])

    if not site_name or not technician_name:
        st.warning("⚠️ Mohon lengkapi **Nama Site** dan **Nama Teknisi** di Tab 1.")
    else:
        if st.button("🚀 Submit Laporan Health Check", type="primary"):
            st.balloons()
            st.success(f"✅ Laporan Site **{site_name}** berhasil disubmit!")
            
            # Preview Tabel
            st.markdown("#### Preview Data Panel")
            st.dataframe(pd.DataFrame(panel_data), use_container_width=True)
            
            st.markdown("#### Preview Data Baterai")
            st.dataframe(pd.DataFrame(bat_data), use_container_width=True)
            
            st.info(f"**Action Terekam:** {action_taken if action_taken else 'Tidak ada catatan action.'}")
            st.info(f"**Status Datalog:** {'Sudah ditarik' if is_datalog_taken else 'Belum ditarik'}")
