import streamlit as st
import pandas as pd
import io

# Konfigurasi Halaman
st.set_page_config(page_title="Analisa Pemadaman MBP", layout="wide")
st.title("⚡ Analisa Tiket Pemadaman & Pergerakan MBP")
st.write("Unggah file Excel raw data Anda, dan sistem akan mengkalkulasi Downtime (Enva Time) serta merangkum pergerakan MBP.")

# Fitur Upload File
uploaded_file = st.file_uploader("Upload File Excel (Ticket MBP)", type=["xlsx", "xls"])

if uploaded_file:
    with st.spinner('Membaca dan memproses data...'):
        # Membaca sheet pertama
        df = pd.read_excel(uploaded_file, sheet_name=0)
        
        # 1. PREPROCESSING & KALKULASI DOWNTIME
        df['Created At'] = pd.to_datetime(df['Created At'], errors='coerce')
        df['Cleared Time'] = pd.to_datetime(df['Cleared Time'], errors='coerce')
        
        # Menghitung selisih waktu (Downtime)
        df['Downtime'] = df['Cleared Time'] - df['Created At']
        df['Downtime (Jam)'] = df['Downtime'].dt.total_seconds() / 3600
        df['Date'] = df['Created At'].dt.date
        
        # Filter khusus yang di-backup MBP (ada RH Start)
        df_mbp = df[df['RH Start'].notna() & (df['RH Start'] != 0)].copy()
        
        # 2. PROSES DATA: KATEGORI CITY
        city_analysis = []
        for city, group in df.groupby('City'):
            total_tiket = len(group)
            unique_sites = group['Site Id'].nunique()
            
            # Filter MBP untuk kota ini
            mbp_group = group[group['RH Start'].notna() & (group['RH Start'] != 0)]
            total_backup = len(mbp_group)
            site_backup_list = ", ".join(mbp_group['Site Id'].dropna().unique())
            
            city_analysis.append({
                'City': city,
                'Total Tiket Down': total_tiket,
                'Total Site Down (Unique)': unique_sites,
                'Total MBP Backup': total_backup,
                'Site yang Di-Backup MBP': site_backup_list,
                'Remark (Alasan Tidak Bisa Backup)': ""
            })
        df_city_report = pd.DataFrame(city_analysis)

        # 3. PROSES DATA: PERGERAKAN HARIAN
        date_analysis = []
        for date, group in df_mbp.groupby('Date'):
            date_analysis.append({
                'Tanggal': date,
                'Total Backup': len(group),
                'Site yang Ter-Backup': ", ".join(group['Site Id'].dropna().unique()),
                'Remark (Alasan Tidak Bisa Backup)': ""
            })
        df_date_report = pd.DataFrame(date_analysis)
        
        # 4. PROSES DATA: DETAIL
        cols_detail = ['Date', 'City', 'Site Id', 'Site Name', 'Created At', 'Cleared Time', 'Downtime', 'Downtime (Jam)', 'RH Start', 'RH Stop']
        # Filter hanya kolom yang tersedia di raw data untuk mencegah error
        cols_available = [col for col in cols_detail if col in df.columns]
        df_detail = df[cols_available].copy()
        df_detail['Remark (Alasan Tidak Bisa Backup)'] = ""

        # TAMPILAN TAB DI STREAMLIT
        tab1, tab2, tab3, tab4 = st.tabs(["📊 Analisa per City", "📅 Pergerakan Harian", "📈 Pivot Table Interaktif", "💾 Download Hasil"])
        
        with tab1:
            st.subheader("Total Down Site & Status Backup (Kategori City)")
            st.dataframe(df_city_report, use_container_width=True)
            
        with tab2:
            st.subheader("Pergerakan Backup MBP per Tanggal")
            st.dataframe(df_date_report, use_container_width=True)
            
        with tab3:
            st.subheader("Custom Pivot Table")
            st.write("Silakan pilih parameter untuk membangun Pivot Table Anda sendiri.")
            col1, col2, col3 = st.columns(3)
            with col1:
                pivot_index = st.selectbox("Pilih Baris (Index):", options=df.columns, index=df.columns.get_loc('City') if 'City' in df.columns else 0)
            with col2:
                pivot_columns = st.selectbox("Pilih Kolom (Opsional):", options=['None'] + list(df.columns), index=0)
            with col3:
                pivot_values = st.selectbox("Pilih Data yang Dihitung (Values):", options=df.columns, index=df.columns.get_loc('Site Id') if 'Site Id' in df.columns else 0)
                pivot_agg = st.selectbox("Metode Kalkulasi:", options=['count', 'sum', 'mean', 'nunique'])
            
            try:
                if pivot_columns == 'None':
                    pivot_df = pd.pivot_table(df, index=pivot_index, values=pivot_values, aggfunc=pivot_agg)
                else:
                    pivot_df = pd.pivot_table(df, index=pivot_index, columns=pivot_columns, values=pivot_values, aggfunc=pivot_agg)
                st.dataframe(pivot_df, use_container_width=True)
            except Exception as e:
                st.warning(f"Tidak dapat membuat pivot dengan kombinasi tersebut. Silakan pilih parameter lain. Error: {e}")

        with tab4:
            st.subheader("Download Hasil Analisa")
            st.write("Semua format tabel di atas, lengkap dengan kolom remark kosong siap ditarik ke dalam 1 file Excel.")
            
            # Membuat file Excel di dalam memory (Buffer)
            buffer = io.BytesIO()
            with pd.ExcelWriter(buffer, engine='xlsxwriter') as writer:
                df_city_report.to_excel(writer, sheet_name='Analisa_Per_City', index=False)
                df_date_report.to_excel(writer, sheet_name='Pergerakan_Harian', index=False)
                df_detail.to_excel(writer, sheet_name='Detail_Data_&_Enva_Time', index=False)
            
            st.download_button(
                label="📥 Download Analisa_Pemadaman_MBP.xlsx",
                data=buffer.getvalue(),
                file_name="Analisa_Pemadaman_MBP.xlsx",
                mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
            )
