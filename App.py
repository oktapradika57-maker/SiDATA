import streamlit as st
import pandas as pd
import plotly.express as px
from wordcloud import WordCloud
import matplotlib.pyplot as plt
import re
from collections import Counter

# Konfigurasi Halaman Streamlit
st.set_page_config(
    page_title="Dashboard Laporan & Analisis Teks",
    page_icon="📊",
    layout="wide"
)

st.title("📊 Dashboard Laporan & Analisis Teks (MBP1)")
st.markdown("Aplikasi ini membaca data laporan Excel dan melakukan analisis data visual serta analisis teks otomatis.")

# Sidebar untuk Pengaturan File
st.sidebar.header("📁 Pengaturan Data")
uploaded_file = st.sidebar.file_uploader("Unggah File Excel (mbp1.xlsx)", type=["xlsx", "xls"])

# Fungsi Pembaca Data
@st.cache_data
def load_data(file):
    if file is not None:
        return pd.read_excel(file, sheet_name=None)
    try:
        return pd.read_excel("mbp1.xlsx", sheet_name=None)
    except FileNotFoundError:
        return None

dict_df = load_data(uploaded_file)

if dict_df is None:
    st.info("💡 Silakan unggah file `mbp1.xlsx` melalui sidebar atau letakkan file tersebut di direktori utama repositori GitHub Anda.")
    st.stop()

# Pilih Sheet
sheet_names = list(dict_df.keys())
selected_sheet = st.sidebar.selectbox("Pilih Sheet Laporan", sheet_names)
df = dict_df[selected_sheet].copy()

# Tab Layout
tab1, tab2, tab3 = st.tabs(["📋 Ringkasan Data", "📈 Analisis Laporan", "🔤 Analisis Teks"])

# --- TAB 1: RINGKASAN DATA ---
with tab1:
    st.subheader(f"Data Sheet: {selected_sheet}")
    
    col1, col2, col3 = st.columns(3)
    col1.metric("Total Baris", df.shape[0])
    col2.metric("Total Kolom", df.shape[1])
    col3.metric("Kolom Teks Terdeteksi", len(df.select_dtypes(include=['object', 'string']).columns))
    
    st.dataframe(df, use_container_width=True)

# --- TAB 2: ANALISIS LAPORAN ---
with tab2:
    st.subheader("Visualisasi Laporan Data")
    numeric_cols = df.select_dtypes(include=['float64', 'int64']).columns.tolist()
    categorical_cols = df.select_dtypes(include=['object', 'category']).columns.tolist()

    if numeric_cols:
        col_x = st.selectbox("Pilih Kolom Sumbu X (Kategori/Tanggal)", options=df.columns, index=0)
        col_y = st.selectbox("Pilih Kolom Sumbu Y (Nilai)", options=numeric_cols, index=0)
        chart_type = st.radio("Tipe Grafik", ["Bar Chart", "Line Chart", "Box Plot"], horizontal=True)

        if chart_type == "Bar Chart":
            fig = px.bar(df, x=col_x, y=col_y, title=f"Grafik Batang {col_y} berdasarkan {col_x}")
        elif chart_type == "Line Chart":
            fig = px.line(df, x=col_x, y=col_y, title=f"Grafik Garis {col_y} berdasarkan {col_x}")
        else:
            fig = px.box(df, x=col_x, y=col_y, title=f"Box Plot {col_y} berdasarkan {col_x}")

        st.plotly_chart(fig, use_container_width=True)
    else:
        st.warning("Tidak ditemukan kolom numerik pada sheet ini untuk membuat grafik kuantitatif.")

# --- TAB 3: ANALISIS TEKS ---
with tab3:
    st.subheader("Analisis Teks & Komentar")
    text_cols = df.select_dtypes(include=['object', 'string']).columns.tolist()

    if text_cols:
        selected_text_col = st.selectbox("Pilih Kolom Teks untuk Menganalisis", text_cols)
        
        # Bersihkan Teks
        text_data = df[selected_text_col].dropna().astype(str)
        combined_text = " ".join(text_data)
        
        # Pembersihan kata sederhana (dapat disesuaikan)
        clean_text = re.sub(r'[^\w\s]', '', combined_text.lower())
        words = [w for w in clean_text.split() if len(w) > 2]

        col_text1, col_text2 = st.columns(2)

        with col_text1:
            st.write("**Kata Paling Sering Muncul (Top 10)**")
            word_counts = Counter(words).most_common(10)
            df_words = pd.DataFrame(word_counts, columns=['Kata', 'Frekuensi'])
            
            fig_words = px.bar(df_words, x='Frekuensi', y='Kata', orientation='h', title="Top 10 Kata Terbanyak")
            fig_words.update_layout(yaxis={'categoryorder': 'total ascending'})
            st.plotly_chart(fig_words, use_container_width=True)

        with col_text2:
            st.write("**Word Cloud**")
            if words:
                wordcloud = WordCloud(width=600, height=400, background_color='white').generate(" ".join(words))
                fig_wc, ax = plt.subplots(figsize=(8, 5))
                ax.imshow(wordcloud, interpolation='bilinear')
                ax.axis('off')
                st.pyplot(fig_wc)
            else:
                st.info("Teks terlalu singkat untuk membuat Word Cloud.")

        st.markdown("---")
        st.write("**Distribusi Panjang Teks (Jumlah Karakter)**")
        df['Panjang_Teks'] = text_data.apply(len)
        fig_len = px.histogram(df, x='Panjang_Teks', nbins=20, title="Distribusi Panjang Karakter Teks")
        st.plotly_chart(fig_len, use_container_width=True)

    else:
        st.warning("Tidak ditemukan kolom berbasis teks pada sheet ini.")
