import streamlit as st
import pandas as pd
import io

# Konfigurasi Halaman
st.set_page_config(page_title="Analisa Pergerakan Tim & Backup MBP", layout="wide")
st.title("⚡ Analisa Pergerakan PIC & Status Backup MBP (Kalteng)")
st.write("Unggah file Excel tiket Anda untuk memantau aktivitas PIC per tanggal, beban site down per City, serta posibilitas backup site.")

# Fitur Upload File
uploaded_file = st.file_uploader("Upload File Excel (Ticket MBP)", type=["xlsx", "xls"])

if uploaded_file:
    with st.spinner('Memproses data dan menganalisis pergerakan PIC...'):
        # Membaca sheet pertama
        df = pd.read_excel(uploaded_file, sheet_name=0)
        
        # Preprocessing Kolom Tanggal & Waktu
        df['Created At'] = pd.to_datetime(df['Created At'], errors='coerce')
        df['Cleared Time'] = pd.to_datetime(df['Cleared Time'], errors='coerce')
        df['Take Over Date'] = pd.to_datetime(df['Take Over Date'], errors='coerce')
        
        df['Date'] = df['Created At'].dt.date
        df['Take Over Date Only'] = df['Take Over Date'].dt.date
        
        # Kalkulasi Downtime (Enva Time)
        df['Downtime'] = df['Cleared Time'] - df['Created At']
        df['Downtime (Jam)'] = df['Downtime'].dt.total_seconds() / 3600
        
        # Deteksi Status Backup (RH Start valid / > 0)
        df['Is_Backup'] = df['RH Start'].apply(lambda x: True if pd.notna(x) and x > 0 else False)

        # Fungsi Penilaian Posibilitas Backup untuk Site yang Tidak Di-backup
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
                
                sites_backup = ", ".join(backed_up_group['Site Id'].unique()) if len(backed_up_group) > 0 else "-"
                sites_no_backup = ", ".join(no_backup_group['Site Id'].unique()) if len(no_backup_group) > 0 else "-"
                
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
                    'Site Sukses Backup': sites_backup,
                    'Site Tidak Di-backup': sites_no_backup,
                    'Alasan (INAP RC 1)': rc_info,
                    'Posibilitas Backup': posibility,
                    'Remark Lapangan': "" # Kolom kosong untuk diisi user
                })
        df_pic_report = pd.DataFrame(pic_analysis)

        # 2. TABEL RINGKASAN PER CITY
        city_analysis = []
        for city, group in df.groupby('City'):
            total_tiket = len(group)
            unique_sites = group['Site Id'].nunique()
            mbp_group = group[group['Is_Backup']]
            total_backup = len(mbp_group)
            site_backup_list = ", ".join(mbp_group['Site Id'].dropna().unique()) if total_backup > 0 else "-"
            
            city_analysis.append({
                'City': city,
                'Total Tiket Down': total_tiket,
                'Total Site Down (Unique)': unique_sites,
                'Total MBP Backup': total_backup,
                'Site yang Di-Backup': site_backup_list,
                'Remark Area': ""
            })
        df_city_report = pd.DataFrame(city_analysis)

        # 3. TABEL DETAIL DATA & ENVA TIME
        cols_detail = ['Date', 'City', 'Site Id', 'Site Name', 'PIC Take Over Ticket', 'Created At', 'Cleared Time', 'Downtime (Jam)', 'RH Start', 'INAP RC 1']
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
            st.subheader("Aktivitas dan Pergerakan PIC Berdasarkan Tanggal & Kota")
            st.write("Menampilkan tanggal PIC mengambil alih tiket, site yang berhasil di-backup, site yang gagal, serta evaluasi posibilitas backup.")
            if not df_pic_report.empty:
                st.dataframe(df_pic_report, use_container_width=True)
            else:
                st.warning("Kolom 'PIC Take Over Ticket' tidak ditemukan pada data.")
            
        with tab2:
            st.subheader("Rekapitulasi Total Down Site & Status Backup per City")
            st.dataframe(df_city_report, use_container_width=True)
            
        with tab3:
            st.subheader("Detail Tiket & Kalkulasi Enva Time (Downtime)")
            st.dataframe(df_detail, use_container_width=True)

        with tab4:
            st.subheader("Custom Pivot Table")
            col1, col2, col3 = st.columns(3)
            with col1:
                pivot_index = st.selectbox("Baris (Index):", options=df.columns, index=df.columns.get_loc('City') if 'City' in df.columns else 0)
            with col2:
                pivot_columns = st.selectbox("Kolom (Opsional):", options=['None'] + list(df.columns), index=0)
            with col3:
                pivot_values = st.selectbox("Values (Dihitung):", options=df.columns, index=df.columns.get_loc('Site Id') if 'Site Id' in df.columns else 0)
                pivot_agg = st.selectbox("Metode Agregasi:", options=['count', 'sum', 'mean', 'nunique'])
            
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
            st.write("Hasil rekapitulasi pergerakan PIC, rekap kota, dan detail waktu siap diunduh dalam format Excel (.xlsx).")
            
            buffer = io.BytesIO()
            with pd.ExcelWriter(buffer, engine='xlsxwriter') as writer:
                if not df_pic_report.empty:
                    df_pic_report.to_excel(writer, sheet_name='Pergerakan_Tim_PIC', index=False)
                df_city_report.to_excel(writer, sheet_name='Analisa_Per_City', index=False)
                df_detail.to_excel(writer, sheet_name='Detail_Enva_Time', index=False)
            
            st.download_button(
                label="📥 Download Analisa_Lengkap_MBP.xlsx",
                data=buffer.getvalue(),
                file_name="Analisa_Pergerakan_Dan_Backup_MBP.xlsx",
                mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
            )
