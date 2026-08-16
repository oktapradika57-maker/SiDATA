import streamlit as st
import pandas as pd
import numpy as np
import io

# Konfigurasi Halaman
st.set_page_config(page_title="Analisa Pergerakan Tim & Backup MBP", layout="wide")
st.title("⚡ Analisa Pergerakan PIC, Status Backup & Jarak MBP (Kalteng)")
st.write("Unggah file Excel tiket Anda untuk memantau aktivitas PIC, kalkulasi jarak tempuh ke site, serta posibilitas backup site.")

# Fungsi Kalkulasi Jarak Haversine (km)
def calculate_haversine(lat1, lon1, lat2, lon2):
    try:
        R = 6371.0 # Radius bumi dalam km
        lat1, lon1, lat2, lon2 = map(np.radians, [float(lat1), float(lon1), float(lat2), float(lon2)])
        dlat = lat2 - lat1
        dlon = lon2 - lon1
        a = np.sin(dlat / 2.0)**2 + np.cos(lat1) * np.cos(lat2) * np.sin(dlon / 2.0)**2
        c = 2 * np.arcsin(np.sqrt(a))
        return round(R * c, 2)
    except (ValueError, TypeError):
        return np.nan

# Mapping Koordinat Default Posko Utama Kalteng (Fallback jika koordinat PIC tidak diisi di Excel)
DEFAULT_CITY_COORDS = {
    "Palangka Raya": (-2.2088, 113.9161),
    "Sampit": (-2.5333, 112.9500),
    "Pangkalan Bun": (-2.6833, 111.6167),
    "Muara Teweh": (-0.9542, 114.8964),
    "Buntok": (-1.7333, 114.8333),
    "Kuala Kapuas": (-3.0083, 114.3833),
    "Puruk Cahu": (-0.6167, 114.5833),
    "Tamiang Layang": (-2.1433, 115.1611)
}

# Sidebar Pengaturan Posko
st.sidebar.header("📍 Pengaturan Posko Tim / Base")
use_default_coords = st.sidebar.checkbox("Gunakan Koordinat Default Kota jika Koordinat PIC Kosong", value=True)

# Fitur Upload File
uploaded_file = st.file_uploader("Upload File Excel (Ticket MBP)", type=["xlsx", "xls"])

if uploaded_file:
    with st.spinner('Memproses data, menghitung jarak tempuh, dan menganalisis pergerakan PIC...'):
        df = pd.read_excel(uploaded_file, sheet_name=0)
        
        # Preprocessing Kolom Tanggal & Waktu
        df['Created At'] = pd.to_datetime(df['Created At'], errors='coerce')
        df['Cleared Time'] = pd.to_datetime(df['Cleared Time'], errors='coerce')
        df['Take Over Date'] = pd.to_datetime(df['Take Over Date'], errors='coerce')
        
        df['Date'] = df['Created At'].dt.date
        df['Take Over Date Only'] = df['Take Over Date'].dt.date
        
        # Kalkulasi Downtime
        df['Downtime'] = df['Cleared Time'] - df['Created At']
        df['Downtime (Jam)'] = (df['Downtime'].dt.total_seconds() / 3600).round(2)
        
        # Deteksi Status Backup
        df['Is_Backup'] = df['RH Start'].apply(lambda x: True if pd.notna(x) and str(x).strip() != '' and float(x) > 0 else False)

        # Kalkulasi Jarak Tempuh
        def get_distance(row):
            # Prioritas 1: Ambil koordinat PIC & Site dari row Excel jika ada
            pic_lat = row.get('PIC Lat') or row.get('Base Lat')
            pic_lon = row.get('PIC Long') or row.get('Base Long')
            site_lat = row.get('Site Lat') or row.get('Latitude')
            site_lon = row.get('Site Long') or row.get('Longitude')

            # Fallback jika koordinat PIC/Base kosong tetapi opsi default aktif
            if (pd.isna(pic_lat) or pd.isna(pic_lon)) and use_default_coords:
                city_name = row.get('City')
                if city_name in DEFAULT_CITY_COORDS:
                    pic_lat, pic_lon = DEFAULT_CITY_COORDS[city_name]

            if pd.notna(pic_lat) and pd.notna(pic_lon) and pd.notna(site_lat) and pd.notna(site_lon):
                return calculate_haversine(pic_lat, pic_lon, site_lat, site_lon)
            return np.nan

        df['Jarak Tempuh (km)'] = df.apply(get_distance, axis=1)

        # Penilaian Posibilitas Backup
        def assess_possibility(rc):
            if pd.isna(rc):
                return "Cek Manual (Tidak ada RC)"
            rc_str = str(rc).lower()
            if 'pln off' in rc_str or 'baterai' in rc_str or 'sewa daya' in rc_str or 'solar cell' in rc_str:
                return "Tinggi (Issue Power/PLN)"
            elif 'rectifier' in rc_str or 'ups' in rc_str:
                return "Rendah (Butuh Perbaikan Rectifier)"
            elif 'osp' in rc_str or 'transport' in rc_str or 'cme' in rc_str or 'telkom' in rc_str or 'isp' in rc_str:
                return "Tidak Bisa (Issue Transmisi/Kabel/Hardware)"
            else:
                return f"Lainnya ({rc})"

        # 1. TABEL ANALISA PERGERAKAN TIM (PIC)
        pic_analysis = []
        if 'PIC Take Over Ticket' in df.columns:
            grouped_pic = df[df['PIC Take Over Ticket'].notna()].groupby(['PIC Take Over Ticket', 'Take Over Date Only', 'City'])
            for (pic, date, city), group in grouped_pic:
                total_handled = group['Site Id'].nunique()
                backed_up_group = group[group['Is_Backup']]
                no_backup_group = group[~group['Is_Backup']]
                
                sites_backup = ", ".join(backed_up_group['Site Id'].astype(str).unique()) if len(backed_up_group) > 0 else "-"
                sites_no_backup = ", ".join(no_backup_group['Site Id'].astype(str).unique()) if len(no_backup_group) > 0 else "-"
                
                # Jarak Tempuh Per PIC / Tanggal
                tot_dist = group['Jarak Tempuh (km)'].sum()
                avg_dist = group['Jarak Tempuh (km)'].mean()
                dist_info = f"{tot_dist:.1f} km (Rata-rata: {avg_dist:.1f} km/site)" if pd.notna(tot_dist) and tot_dist > 0 else "Data Koordinat Tidak Lengkap"

                # Ringkasan RC / Alasan
                if len(no_backup_group) == 0:
                    rc_info = "Semua Sukses Backup"
                    posibility = "-"
                else:
                    rcs = no_backup_group['INAP RC 1'].dropna().value_counts()
                    rc_info = ", ".join([f"{k} ({v})" for k, v in rcs.items()]) if not rcs.empty else "Auto Resolved / No RC"
                    
                    possibilities = [assess_possibility(rc) for rc in no_backup_group['INAP RC 1']]
                    pos_series = pd.Series(possibilities).value_counts()
                    posibility = ", ".join([f"{k} ({v} site)" for k, v in pos_series.items()])
                
                pic_analysis.append({
                    'PIC': pic,
                    'Tanggal Take Over': date,
                    'Kota (City)': city,
                    'Total Site Down': total_handled,
                    'Estimasi Jarak Tempuh': dist_info,
                    'Site Sukses Backup': sites_backup,
                    'Site Tidak Di-backup': sites_no_backup,
                    'Alasan (INAP RC 1)': rc_info,
                    'Posibilitas Backup': posibility,
                    'Remark Lapangan': ""
                })
        df_pic_report = pd.DataFrame(pic_analysis)

        # 2. TABEL RINGKASAN PER CITY
        city_analysis = []
        for city, group in df.groupby('City'):
            total_tiket = len(group)
            unique_sites = group['Site Id'].nunique()
            mbp_group = group[group['Is_Backup']]
            total_backup = len(mbp_group)
            site_backup_list = ", ".join(mbp_group['Site Id'].astype(str).dropna().unique()) if total_backup > 0 else "-"
            avg_city_dist = group['Jarak Tempuh (km)'].mean()
            
            city_analysis.append({
                'City': city,
                'Total Tiket Down': total_tiket,
                'Total Site Down (Unique)': unique_sites,
                'Total MBP Backup': total_backup,
                'Rata-rata Jarak ke Site (km)': round(avg_city_dist, 2) if pd.notna(avg_city_dist) else "-",
                'Site yang Di-Backup': site_backup_list,
                'Remark Area': ""
            })
        df_city_report = pd.DataFrame(city_analysis)

        # 3. TABEL DETAIL DATA & ENVA TIME
        cols_detail = ['Date', 'City', 'Site Id', 'Site Name', 'PIC Take Over Ticket', 'Jarak Tempuh (km)', 'Created At', 'Cleared Time', 'Downtime (Jam)', 'RH Start', 'INAP RC 1']
        cols_available = [col for col in cols_detail if col in df.columns]
        df_detail = df[cols_available].copy()
        df_detail['Remark (Alasan Tidak Bisa Backup)'] = ""

        # STRUKTUR TAMPILAN STREAMLIT (TABS)
        tab1, tab2, tab3, tab4, tab5 = st.tabs([
            "🏃‍♂️ Pergerakan Tim (PIC)", 
            "🏙️ Analisa per City", 
            "📋 Detail & Enva Time", 
            "📈 Pivot Interaktif", 
            "💾 Download Excel"
        ])
        
        with tab1:
            st.subheader("Aktivitas, Pergerakan & Mobilisasi Jarak PIC")
            st.write("Menampilkan pergerakan harian PIC, estimasi akumulasi jarak perjalanan, site backup, serta evaluasi kendala.")
            if not df_pic_report.empty:
                st.dataframe(df_pic_report, use_container_width=True)
            else:
                st.warning("Kolom 'PIC Take Over Ticket' tidak ditemukan pada data.")
            
        with tab2:
            st.subheader("Rekapitulasi Total Down Site, Jarak Rata-rata & Status Backup per City")
            st.dataframe(df_city_report, use_container_width=True)
            
        with tab3:
            st.subheader("Detail Tiket, Kalkulasi Jarak & Enva Time (Downtime)")
            st.dataframe(df_detail, use_container_width=True)

        with tab4:
            st.subheader("Custom Pivot Table")
            col1, col2, col3 = st.columns(3)
            with col1:
                pivot_index = st.selectbox("Baris (Index):", options=df.columns, index=df.columns.get_loc('City') if 'City' in df.columns else 0)
            with col2:
                pivot_columns = st.selectbox("Kolom (Opsional):", options=['None'] + list(df.columns), index=0)
            with col3:
                pivot_values = st.selectbox("Values (Dihitung):", options=df.columns, index=df.columns.get_loc('Jarak Tempuh (km)') if 'Jarak Tempuh (km)' in df.columns else 0)
                pivot_agg = st.selectbox("Metode Agregasi:", options=['sum', 'mean', 'count', 'nunique'])
            
            try:
                if pivot_columns == 'None':
                    pivot_df = pd.pivot_table(df, index=pivot_index, values=pivot_values, aggfunc=pivot_agg)
                else:
                    pivot_df = pd.pivot_table(df, index=pivot_index, columns=pivot_columns, values=pivot_values, aggfunc=pivot_agg)
                st.dataframe(pivot_df, use_container_width=True)
            except Exception as e:
                st.warning(f"Gagal membuat pivot table dengan kombinasi tersebut: {e}")

        with tab5:
            st.subheader("Unduh Laporan Lengkap ke Excel")
            st.write("Hasil rekapitulasi pergerakan PIC, analisa jarak tempuh, rekap kota, dan detail waktu siap diunduh.")
            
            buffer = io.BytesIO()
            with pd.ExcelWriter(buffer, engine='xlsxwriter') as writer:
                if not df_pic_report.empty:
                    df_pic_report.to_excel(writer, sheet_name='Pergerakan_Tim_PIC', index=False)
                df_city_report.to_excel(writer, sheet_name='Analisa_Per_City', index=False)
                df_detail.to_excel(writer, sheet_name='Detail_Enva_Time_Jarak', index=False)
            
            st.download_button(
                label="📥 Download Analisa_Lengkap_MBP.xlsx",
                data=buffer.getvalue(),
                file_name="Analisa_Pergerakan_Jarak_Dan_Backup_MBP.xlsx",
                mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
            )
