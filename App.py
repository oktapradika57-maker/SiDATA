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
st.markdown(
    "Formulir pemeriksaan berkala sistem PLTS dan Catu Daya BTS Telekomunikasi."
)
st.markdown("---")

# Menggunakan Tabs untuk Navigasi Bagian Form
tabs = st.tabs(
    [
        "1. Info Site",
        "2. Solar Panel",
        "3. Junction Box",
        "4. SCC & Rectifier",
        "5. Baterai",
        "6. Grounding",
        "7. Ringkasan & Submit",
    ]
)

# -------------------------------------------------------------------------
# TAB 1: INFORMASI SITE
# -------------------------------------------------------------------------
with tabs[0]:
  st.subheader("Informasi Umum Site")
  col1, col2 = st.columns(2)

  with col1:
    site_name = st.text_input(
        "Nama / ID Site", placeholder="Contoh: BTS-PKY-001"
    )
    nop_area = st.selectbox(
        "Network Operation Point (NOP)",
        ["Palangkaraya", "Pangkalan Bun", "Tarakan", "Pontianak", "Lainnya"],
    )
    check_date = st.date_input(
        "Tanggal Pengecekan", value=datetime.date.today()
    )

  with col2:
    technician_name = st.text_input(
        "Nama Teknisi", placeholder="Nama lengkap teknisi"
    )
    weather = st.selectbox(
        "Kondisi Cuaca", ["Cerah", "Berawan", "Hujan / Mendung"]
    )

  st.info(
      "Pastikan semua parameter terisi dengan benar sebelum melanjutkan ke"
      " tab berikutnya."
  )

# -------------------------------------------------------------------------
# TAB 2: SOLAR PANEL (PER SATUAN / PANEL)
# -------------------------------------------------------------------------
with tabs[1]:
  st.subheader("Pemeriksaan Modul Panel Surya (PV Array)")
  st.markdown(
      "Masukkan data dan unggah foto dokumentasi untuk masing-masing panel atau"
      " string."
  )

  num_panels = st.number_input(
      "Jumlah Panel / String yang Diperiksa",
      min_value=1,
      max_value=20,
      value=4,
      step=1,
  )

  panel_data = []
  for i in range(int(num_panels)):
    with st.expander(f"Panel / String #{i+1}", expanded=(i == 0)):
      p_col1, p_col2 = st.columns(2)

      with p_col1:
        voc = st.number_input(
            f"Tegangan Open Circuit (Voc) - Panel {i+1} [Volt]",
            min_value=0.0,
            value=21.5,
            step=0.1,
            key=f"voc_{i}",
        )
        isc = st.number_input(
            f"Arus Short Circuit (Isc) - Panel {i+1} [Ampere]",
            min_value=0.0,
            value=5.2,
            step=0.1,
            key=f"isc_{i}",
        )
        p_condition = st.selectbox(
            f"Kondisi Fisik - Panel {i+1}",
            ["Baik", "Retak / Pecah", "Kotor / Soiling", "Delaminasi"],
            key=f"cond_{i}",
        )

      with p_col2:
        p_photo = st.file_uploader(
            f"Upload Foto Fisik - Panel {i+1}",
            type=["jpg", "jpeg", "png"],
            key=f"photo_panel_{i}",
        )
        if p_photo is not None:
          st.image(
              p_photo,
              caption=f"Preview Panel {i+1}",
              use_container_width=True,
          )

      panel_data.append({
          "Panel_ID": f"Panel #{i+1}",
          "Voc": voc,
          "Isc": isc,
          "Condition": p_condition,
          "Has_Photo": True if p_photo is not None else False,
      })

# -------------------------------------------------------------------------
# TAB 3: JUNCTION BOX / COMBINER BOX
# -------------------------------------------------------------------------
with tabs[2]:
  st.subheader("Junction Box / Combiner Box (DC Box)")

  jb_col1, jb_col2 = st.columns(2)
  with jb_col1:
    jb_enclosure = st.selectbox(
        "Kondisi Enklosur & Seal Karet",
        ["Baik & Kedap Air", "Rusak / Seal Lepas", "Berkarat / Ada Serangga"],
    )
    jb_fuse = st.selectbox(
        "Status Fuse & MCB DC",
        ["Normal / Aman", "Trip / Perlu Reset", "Putus / Terbakar"],
    )
    jb_spd = st.selectbox(
        "Indikator Surge Protection Device (SPD)",
        ["Normal (Hijau)", "Rusak / Triggered (Merah/Hitam)"],
    )
    jb_terminal = st.selectbox(
        "Kondisi Terminal & Busbar",
        ["Kencang & Bersih", "Kendur", "Oksidasi / Gosong"],
    )

  with jb_col2:
    jb_photo = st.file_uploader(
        "Upload Foto Dalam Junction Box",
        type=["jpg", "jpeg", "png"],
        key="jb_photo",
    )
    if jb_photo is not None:
      st.image(
          jb_photo, caption="Preview Junction Box", use_container_width=True
      )

# -------------------------------------------------------------------------
# TAB 4: SOLAR CHARGE CONTROLLER (SCC) / RECTIFIER
# -------------------------------------------------------------------------
with tabs[3]:
  st.subheader("Solar Charge Controller (SCC) / Rectifier")

  scc_col1, scc_col2 = st.columns(2)
  with scc_col1:
    scc_alarm = st.selectbox(
        "Status Indikator & Alarm",
        ["Normal (No Alarm)", "Ada Alarm Fault / Error"],
    )
    alarm_code = st.text_input(
        "Kode Alarm (Jika ada)", placeholder="Contoh: Err-03 / Overvoltage"
    )
    float_v = st.number_input(
        "Pengukuran Float Voltage [V]", min_value=0.0, value=54.2, step=0.1
    )
    system_out_v = st.number_input(
        "Tegangan Output DC ke BTS [V]", min_value=0.0, value=48.0, step=0.1
    )
    scc_fan = st.selectbox(
        "Kondisi Kipas Pendingin (Fan)",
        ["Berputar Normal", "Berisik / Macet", "Mati Total"],
    )

  with scc_col2:
    scc_photo = st.file_uploader(
        "Upload Foto Display SCC / Rectifier",
        type=["jpg", "jpeg", "png"],
        key="scc_photo",
    )
    if scc_photo is not None:
      st.image(
          scc_photo, caption="Preview SCC / Rectifier", use_container_width=True
      )

# -------------------------------------------------------------------------
# TAB 5: BANK BATERAI (BATTERY BANK)
# -------------------------------------------------------------------------
with tabs[4]:
  st.subheader("Pemeriksaan Bank Baterai")

  bat_col1, bat_col2 = st.columns(2)
  with bat_col1:
    bat_physical = st.selectbox(
        "Kondisi Fisik Casing Baterai",
        ["Normal / Mulus", "Bengkak (Swelling)", "Retak / Bocor"],
    )
    bat_temp = st.number_input(
        "Suhu Ruang Baterai [°C]", min_value=15.0, max_value=60.0, value=28.5
    )
    bat_term = st.selectbox(
        "Kondisi Terminal Baterai",
        [
            "Bersih & Kencang",
            "Ada Kerak Oksidasi (Putih/Hijau)",
            "Kendur / Panas",
        ],
    )
    bat_total_v = st.number_input(
        "Total Tegangan Bank Baterai [V]", min_value=0.0, value=52.8, step=0.1
    )

  with bat_col2:
    bat_photo = st.file_uploader(
        "Upload Foto Bank Baterai", type=["jpg", "jpeg", "png"], key="bat_photo"
    )
    if bat_photo is not None:
      st.image(
          bat_photo, caption="Preview Bank Baterai", use_container_width=True
      )

# -------------------------------------------------------------------------
# TAB 6: GROUNDING & LIGHTNING PROTECTION
# -------------------------------------------------------------------------
with tabs[5]:
  st.subheader("Sistem Pembumian & Penangkal Petir")

  grd_col1, grd_col2 = st.columns(2)
  with grd_col1:
    earth_resistance = st.number_input(
        "Nilai Tahanan Pentanahan (Earth Tester) [Ohm]",
        min_value=0.0,
        value=2.1,
        step=0.1,
    )
    grd_cable = st.selectbox(
        "Kondisi Kabel & Sambungan Grounding",
        ["Terhubung Kuat & Utuh", "Kendor", "Putus / Hilang (Pencurian)"],
    )

  with grd_col2:
    grd_photo = st.file_uploader(
        "Upload Foto Pengukuran Grounding / Batang Ground",
        type=["jpg", "jpeg", "png"],
        key="grd_photo",
    )
    if grd_photo is not None:
      st.image(
          grd_photo, caption="Preview Grounding", use_container_width=True
      )

# -------------------------------------------------------------------------
# TAB 7: RINGKASAN & SUBMIT
# -------------------------------------------------------------------------
with tabs[6]:
  st.subheader("Ringkasan Hasil Health Check")

  if not site_name or not technician_name:
    st.warning(
        "⚠️ Mohon lengkapi **Nama Site** dan **Nama Teknisi** di Tab 1 (Info"
        " Site) sebelum melakukan submit."
    )
  else:
    st.success(
        "✅ Data siap disubmit. Berikut ringkasan pemeriksaan untuk site:"
        f" **{site_name}**"
    )

    # Tampilkan Ringkasan Data Panel
    st.markdown("### Ringkasan Panel Surya")
    df_panels = pd.DataFrame(panel_data)
    st.dataframe(df_panels, use_container_width=True)

    # Catatan Tambahan
    notes = st.text_area(
        "Catatan Tambahan / Temuan Penting di Lapangan",
        placeholder="Tuliskan temuan atau tindakan perbaikan yang sudah dilakukan...",
    )
    sparepart_needed = st.text_input(
        "Daftar Sparepart yang Perlu Penggantian (Jika ada)"
    )

    final_status = st.radio(
        "Status Keseluruhan Site:",
        [
            "Normal / Healthy (Siap Beroperasi Optimal)",
            "Minor Issue (Sudah Ditangani di Tempat)",
            "Major Issue / Critical (Perlu Eskalasi & Penggantian Sparepart)",
        ],
    )

    if st.button("🚀 Kirim Laporan Health Check", type="primary"):
      st.balloons()
      st.success(
          f"Laporan untuk site **{site_name}** (NOP: {nop_area}) berhasil"
          " disimpan dan dikirim ke sistem!"
      )
      # Di sini Anda bisa menambahkan logika tambahan seperti export ke CSV/Excel atau kirim ke Google Sheets API / Database.
