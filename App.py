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
from fpdf import FPDF
import requests
from io import BytesIO
from PIL import Image
import base64
import tempfile 

# Mengimpor modul PPTX
try:
    from pptx import Presentation
    from pptx.util import Inches, Pt
    from pptx.dml.color import RGBColor
    HAS_PPTX = True
except ImportError:
    HAS_PPTX = False

# -------------------------------------------------------------------------
# SETUP HALAMAN & FUNGSI LOGO PRESISI (.WEBP)
# -------------------------------------------------------------------------
st.set_page_config(page_title="RPO - PLK & PBU", page_icon="⚡", layout="wide")

LOGO_UTAMA = "koperasi-jasa-konstruksi-tower-event-organizer-network-monitoring-telekomunikasi-kisel-group-logo-kut (2).webp"

def get_base64_of_bin_file(bin_file):
    try:
        with open(bin_file, 'rb') as f:
            data = f.read()
        return base64.b64encode(data).decode()
    except:
        return None

def render_header_logo(is_login=False):
    logo_candidates = [
        (LOGO_UTAMA, "image/webp"),
        ("logo.png", "image/png"),
        ("logo.webp", "image/webp")
    ]
    for filename, mime in logo_candidates:
        if os.path.exists(filename):
            bg_base64 = get_base64_of_bin_file(filename)
            if bg_base64:
                max_h = "135px" if is_login else "100px"
                pad = "16px 36px" if is_login else "12px 28px"
                st.markdown(
                    f"""
                    <div style="display: flex; justify-content: center; align-items: center; margin-bottom: 20px;">
                        <div style="background: linear-gradient(145deg, #ffffff, #f3f6f9); padding: {pad}; border-radius: 20px; box-shadow: 0 10px 28px rgba(0,0,0,0.22); border: 2px solid var(--primary-color); display: inline-flex; align-items: center; justify-content: center;">
                            <img src="data:{mime};base64,{bg_base64}" style="max-height: {max_h}; width: auto; max-width: 100%; object-fit: contain; display: block;">
                        </div>
                    </div>
                    """, 
                    unsafe_allow_html=True
                )
                return

# -------------------------------------------------------------------------
# FUNGSI PEMBUAT DOKUMEN (PDF & PPTX)
# -------------------------------------------------------------------------
def clean_text(text):
    if not text: return "-"
    return str(text).encode('latin-1', 'ignore').decode('latin-1')

def build_pptx(db_list):
    prs = Presentation()
    slide_layout = prs.slide_layouts[0]
    slide = prs.slides.add_slide(slide_layout)
    title = slide.shapes.title
    subtitle = slide.placeholders[1]
    title.text = "Laporan Lengkap Preventive Maintenance"
    subtitle.text = f"Total Site Terinspeksi: {len(db_list)}\nGenerated on: {datetime.date.today()}"
    
    for r in db_list:
        slide_layout = prs.slide_layouts[5]
        slide = prs.slides.add_slide(slide_layout)
        title = slide.shapes.title
        title.text = f"Site: {r.get('site_name', '-')} | Status: {r.get('status', '-')}"
        
        table_shape = slide.shapes.add_table(6, 2, Inches(0.5), Inches(1.5), Inches(4.5), Inches(3.0))
        table = table_shape.table
        
        table_data = [
            ("Tanggal", r.get('timestamp', '-')),
            ("Kategori / NOP", f"{r.get('kategori', 'SPS')} / {r.get('nop', '-')}"),
            ("Teknisi", r.get('teknisi', '-')),
            ("Tegangan / Load", f"{r.get('rect_out_v', '-')} V / {r.get('total_load', '-')} A"),
            ("Action Lapangan", r.get('action', '-')),
            ("Sparepart Diganti", r.get('sparepart', '-'))
        ]
        
        for row_idx, (k, v) in enumerate(table_data):
            table.cell(row_idx, 0).text = k
            table.cell(row_idx, 1).text = str(v)
            for cell in [table.cell(row_idx, 0), table.cell(row_idx, 1)]:
                for paragraph in cell.text_frame.paragraphs: paragraph.font.size = Pt(13)
        
        site_urls = r.get('url_sites', [])
        if site_urls and len(site_urls) > 0:
            try:
                img_url = site_urls[0].replace("upload/v", "upload/c_limit,w_500,q_80,f_jpg/v")
                resp = requests.get(img_url, timeout=5)
                if resp.status_code == 200:
                    image_stream = BytesIO(resp.content)
                    slide.shapes.add_picture(image_stream, Inches(5.3), Inches(1.5), width=Inches(4.2))
            except: pass
                
    out = BytesIO()
    prs.save(out)
    return out.getvalue()

def build_pdf(r):
    pdf = FPDF()
    pdf.add_page()
    pdf.set_auto_page_break(auto=True, margin=15)
    pdf.set_font("helvetica", "B", 14)
    pdf.cell(0, 10, clean_text("BERITA ACARA PREVENTIVE MAINTENANCE"), ln=True, align="C")
    pdf.set_font("helvetica", "B", 12)
    pdf.cell(0, 6, clean_text("SITE TELEKOMUNIKASI"), ln=True, align="C")
    pdf.ln(5)
    
    pdf.set_font("helvetica", "", 10)
    intro = f"Berdasarkan hasil inspeksi dan pengerjaan lapangan pada tanggal {r.get('timestamp', '-')}, dengan ini diterangkan bahwa tim teknisi telah melaksanakan kegiatan Preventive Maintenance (Pemeliharaan Berkala) pada:"
    pdf.multi_cell(0, 6, clean_text(intro))
    pdf.ln(2)
    
    pdf.set_font("helvetica", "B", 10)
    pdf.cell(40, 6, clean_text("Nama Site / ID"), 0, 0)
    pdf.cell(0, 6, clean_text(f": {r.get('site_name', '-')}"), 0, 1)
    pdf.cell(40, 6, clean_text("Kategori Site"), 0, 0)
    pdf.cell(0, 6, clean_text(f": {r.get('kategori', 'SPS')}"), 0, 1)
    pdf.cell(40, 6, clean_text("Regional / NOP"), 0, 0)
    pdf.cell(0, 6, clean_text(f": {r.get('nop', '-')}"), 0, 1)
    pdf.cell(40, 6, clean_text("Pelaksana (Tim)"), 0, 0)
    pdf.cell(0, 6, clean_text(f": {r.get('teknisi', '-')}"), 0, 1)
    pdf.cell(40, 6, clean_text("Status Akhir"), 0, 0)
    pdf.cell(0, 6, clean_text(f": {r.get('status', '-')}"), 0, 1)
    pdf.ln(5)
    
    pdf.set_font("helvetica", "B", 11)
    pdf.cell(0, 8, clean_text("1. RINCIAN KEGIATAN & TINDAKAN (ACTION)"), ln=True)
    pdf.set_font("helvetica", "", 10)
    action_text = r.get('action', '-')
    if not action_text.strip(): action_text = "Tidak ada tindakan/catatan khusus yang dilaporkan oleh tim."
    pdf.multi_cell(0, 6, clean_text("Adapun rincian tindakan dan pekerjaan yang telah diselesaikan oleh tim di lokasi adalah sebagai berikut:\n" + action_text))
    pdf.ln(2)
    
    pdf.set_font("helvetica", "B", 10)
    pdf.cell(0, 6, clean_text("Penggunaan / Penggantian Sparepart:"), ln=True)
    pdf.set_font("helvetica", "", 10)
    pdf.multi_cell(0, 6, clean_text(r.get('sparepart', '-')))
    pdf.ln(5)
    
    pdf.set_font("helvetica", "B", 11)
    pdf.cell(0, 8, clean_text("2. HASIL PENGECEKAN PARAMETER SITE"), ln=True)
    pdf.set_font("helvetica", "", 10)
    
    pdf.cell(0, 6, clean_text(f"[v] Fisik & Lingkungan : {r.get('site_cond', '-')} | Fisik Tower: {r.get('tower_cond', '-')}"), ln=True)
    pdf.cell(0, 6, clean_text(f"[v] Modul Surya (SPS)  : Kondisi Shading {r.get('shading_status', '-')}"), ln=True)
    pdf.cell(0, 6, clean_text(f"[v] Jaringan Grid PLN  : {r.get('pln_status', '-')}"), ln=True)
    pdf.cell(0, 6, clean_text(f"[v] Rectifier & Beban  : {r.get('rect_brand', '-')} | Out: {r.get('rect_out_v', '-')}V | Load BTS: {r.get('total_load', '-')}A"), ln=True)
    pdf.cell(0, 6, clean_text(f"[v] Genset & Level BBM : {r.get('genset_status', '-')} | Ketersediaan BBM: {r.get('fuel_pct', '-')}%"), ln=True)
    pdf.cell(0, 6, clean_text(f"[v] Sistem Grounding   : Terukur {r.get('earth_ohm', '-')} Ohm"), ln=True)
    pdf.ln(5)
    pdf.multi_cell(0, 6, clean_text("Demikian Berita Acara ini dibuat sebenar-benarnya sesuai dengan kondisi aktual di lapangan untuk dapat dipergunakan sebagaimana mestinya."))
    pdf.ln(10)

    pdf.set_font("helvetica", "", 10)
    pdf.cell(90, 6, "", 0, 0) 
    pdf.cell(90, 6, clean_text("Data dibuat oleh,"), 0, 1, "C")
    try:
        TTD_OKTA_B64 = "iVBORw0KGgoAAAANSUhEUgAAAF0AAABkCAYAAADkBDymAAAVAklEQVR42u2de3hU1bXA197nfeY9yUwykwSIgjxVCvEBioBWW21v7y3fbfp9fWlvq2ktjwq+COpkeAiIQSmUNtQr91L97m28t7VetajViCLPAIogJDzCIyGZZDKZ53mfs+8fWDTykEBmQjDrr+Sbsx/rN/uss9baa58BGJABybYQQlBjYyPXl3Ogv2rQq5a/4dFl7SYA+L8B6DmSzqTDTzNs0cBKz6FYNO8nDO8YgJ5DUTXWpRlaxwD0nD5JdYeuWfG+nAL+qjHHAD5ZM2MD0HMoiqrSGNF9utK/cuaFE9hCOzLTA9BzKF0JnDGNWGLAvOQuGqX8ecLQSWOHkwHoOZLR5S9RXV1SZPLPS1ID5iVHcsPQsS5EGa7rMNIHoOfqtqbUfFk1EoT08Ty+StBlxfDSLCdfArHCV0cEG1cgpa3jfT2Pr5R5yaR1B0Ja52ULvbp2k0BLjmBHpJn2uX2Gk021/PSnU5W+VFYU2aJYQn/vsoP+dE1dPhj4NiZhuhle7xTstg5gkFPG3LQV/7Hxo1n33PxmXymbSMmg6+yJywr60t+/d5tpklKaJltn3jvp489/VlNfvx7tN+9ZXrPhjtkVk/sEvM9DD3G5bMZrl8ODdG1dE7+k5u0fYLBEKfLO83Mqpnz8xWsqysr0X/7ohj+YFi75zbotpX0QjeJE0srwtBzr9yu9unqT0PrJsbs5Vtg5596J277s+jQ2/momje8CIc8BQjnzmMur9opOkfGOnjJK69cuY3X1JkFnpJ9gIBvOBzgAQLhiahRjOrN01VuBXCrqBT3f0LVM+RjUf6GHQiEs0co0iqY+ePRXU/b1pK2imdtVzN6YS0Ulic83CdXnpuWioLuDd34HLOr4Q/dP3dPTtiOCkw8zNFVaU1PP5EpRkzIKDR239lvoS1a9PdowddfjD0y+IJ+3vByZmKDmuJUekrNolIIShjb6J/TnXt7nME24xUkJL13MwLpuHkgpcHXONEVcPgLS3C+hx1o6/tlht22qqCiTLmZgN8prFFimBBDKiaK6roo2wRvtd9CX/e69URhR7Myf3fjRxQ48ffqYNKYoecmfDrpy4KMjhCyX4JCT/Qp6bW0tpRlkKmLQy701uKZbHVRnR9ZL3Gb+5gBrWgx348ThmX4F/VisaIrICofn3Dux19wuQ4b2eJouzLaSyeMtHkIs9fsTkdxvoIdCdbxmqKPjniNv9+bgqUymNSVFPVk3L0yenyAjSeDSkPOCjr3odgB6a7i8vFejOZVv7xIdgi/bSjKsWEwA2qG/QK+prXexLB5aOf2Wbb09uNmUTusKrRJCsurCGHrGTxN8vN9Ab4ukb81k0u8CQK/fnTU19xleN11Q9dLerEamDEUXCyIV6RfQq2s3CciyfNcMSu/OSryCEJElrR1a03xv9blw+d/H1dS81c0NTcmGSEPfbV7U1RF69bo3/HV1dTTAl6R2lVbtRgozTeXl5WbWViHHsxRNOwCgxz704tXbrkjEUzeohs4Egp4Dj9x74+ajUft3j+1vX/yZE0Bwo1Tv9BV7cpbsWlW7x64mkoUMxQYSaSX4/vY3hXwvn5LlwKsAYJwTuolhmKA4XszmBFsjnc2YGD3O6y9ds3GSqqs3O/Jcf+HSWns6nv5GeOUmz7FO3ram5p/kNWtOXnfI3iZghRFHTyjJStEoIQRVLX/D4xRtJYxgC6qq7k+0ddF2u9hOc3CCp+116UNvRe9/OGz9o81ZlV2w6oPBFLbUhx4am9WAotBX4JJkNQAAR8+3zaJn68bIklxGgoHqJ8rHaJ8Gb7U7m13rKMIeRp/bHOEV1Z/QcLK38uihmnrRgcximyAGVVX1LVi+gaEoHmiGaRVF4bBsSBvnzZx0zrK9s0InsjJeBWNHtm/Fg0ePNFkGnHeZW+gPe7x6put7BAnV4U+BAwCUl5eb02Z+tM0uUt1MIQf6YIZGF5RzCYVqWb4gUOxyOItUEzyKrLnlVNJUBTrDYO4oEH3fiJLJ0fJy1CPzS5/llsHLVn9QdN2oW18JZxm6Ly+PsQn2IACc1xeMlfg9ugW1C2ffmPy83QZ4l93RyWFJ7e6lpFKyHyz9S++i0No6HiKQz9j0Qsy6g8jIOOMpw0KGbloWaeE5stvATMf8X9560Xf+GaE/9vR7g2mMuqZORUa2V7pdEIxkRqLO59rHqt+/1rQseeHsW/Z2+yBv+w9sgnskncR+jxd3PzlHGQFVRtu7eWXVmwSJwYWcQAVZxvLTNJ/fFY1raUOLMjpETBvs4gVPbOmvsmNazwgdcfS1mmXuhhzIto/2tLA0J57HIwsh8v43nLz4/BdSFLQoUqMFG8IOkTgLXTG2m3dEM1eWDkENi36zcRLL0AWI4uyqHDe1DIpbFNPGEWo749Vjj/9osgQ5kjPbdEVxsiNcB3MxgUJ3Ifj8jiFfdnx50bNbrpY0kniwoqybfbYF8wcbun7cUFQ/EJxKSal05VMbx4qcUWJ3efO3HRTGe8zYNoOBDKXBFsaQOyvnTO3TxNdp0J9cuTUvnU7Cwu+U5eSbzy8OII5CX5r00ghM0S3uz90j2nrxcMq8g6ZRcWfaMZoTKIdJ8mnMJPfrpr67K8UolKXc+vTD178Al5CcBj2dSQwjmG7J1QSKvB4UiXWcM8Xw5MpP8mKpmCc/j+B5S96eiDi20Od05UumoRAgE2gav0KbMAybzN/27T/8yd9fTK0HKDfvDx0YCiZJIQRAyCUMHbMwyCLWB7maQHukLdGZiHbzONa90Wprb4l6O2JyvoNWhyYNMpm1uVLEio83iHXEzRs7r7mis3n3bpliwcZGk1odTfLLJY0+PHLk4MAdawXfwz+FNsXQgkCg5VICfhr0UIhgmtlWMLzAlrNj3MdbG3WbI+Bd+vzeSbFoxC+w4qATx5oZlmVkrxNFDZ07oacztMCoL/js+YbNLvpbo4klWxuZcArMUgvphcE8atLhFjzIZTMYliXFuqa6a2uPJd5pzAw1LOUAwCUMnfVs8SmaZZZ/LujozXB56Zodzs647td0+xWmga+xELnKwIwfY3psOtP2nsflSKQTdL2sk7QzaLgddnAeOW6IpsWzImuUmFh3ZFKJMgKEyyTjpmoyXwsUeCPxeGKMqvM+t8vItwi+WlOt6JH2zrRlif9qF+ndT1S/Txw2NqLrXFtegRGpKC9LXDLQUzoMAUyOXGyny1Yf9DerZh6SUyWmZhVYIDvnLn2f4njGALAUYtIJwc41GCjzJkjpuFfE9y2YcdO8M/X18JLNt9sEtaFy+sSNAADzV24LcsjcFH5owtZ5y7dPHjek9Jn/Xv/3m2lMvnWsfcgfWX2f5BxZ8vLD946K/fjR3WOLAq6XGK09SlH2YkSTUa3t5nVzl6wXeZstaiHq46pfTTiEclhTeRp0m0gNUTXzwwvp6PGn3p0Ql5hxiu72HYkmKYvC7RyQVoFHDTTrOz7YJkUrKspOC/dnrHidw8R31iAsIyujWZlbd+p/k7kVANbVvELE4wd3cVOnIuOBp7YP15KY3HWHR33zZYp58X/3ZEJ1hD6wfo/bHTAOPlp+W+LzuZ1QTb3IU+YoNS1fP2/p2998ctXmw6pNeSeco0ML3aDzNBMUWeOtC+lIUs0TNqcY8bKkPTx93Hln9Ly8nzIk44w7R6Fn6txdEks9M+/kZnhoVcQeicV8o8f43t27Y9MVBKHDAAByRhdMk4veVARkq8sxpugG2Wzf8rHDMi31ke9dkXr0C/2GT9bs1ANA/apVe+xdZuxmnGBmL165edPcGRPezTb0U5sY99XUMF1JzUw2T4hfSEfVj992dMns6w+Hp4/pUQo1GAiAYZ15ocuqMIIzoeGUpxNrLjE0ZfuMacWdQOExFrEaAADyvbYRALjj6RdPUEnJ5N+bf6uBdaaYoqw4Qsg61/jTp49JPz7rlvWlXvsKAmTwE8s2/GzFite5nEDPi17jzcjxZDh87kn2tgSCAIhiz7jSLR2utSjq8Kn0hMWWEdN4FQDAJrBDebv9GEIIMlIGOW0I2f18nqIbMiEEMrpeyLHseW/R/eQnYzOVMyb+J8cyB2OG45e1tYTKOnQkKgGWouKQY6Hbg6aqGkn4wuZ0KESwpOgerS126NN8OaVZ1o8dbl9nqK6O1nVC3TC4sdOyLEQooTCZ0Bra26JBnfAKAIChQzFH0z3ejJ43a+IGhucO7W7Z8K2sQ+doR4Fos+W81u/OO0GLdegdoSrovto9e4ttblZfufIuFQBgX8ewWwryNN7lbjrWsdFZlFGUWHl5uYkQIhTjKLY73IalJorsnH4EAEDTSRGxLqwCQG9/7TUw2CuXLfvQlt2VDkwe4NxDRwgRRWZOiwuSmWQpZVENJ1c5YRVJnmQpaP3oKwbpeiY2jKFREwDAyxs7HHJapzoTWlOBM2UfN4z8DQAAEyNgs9MXVBodDoctm41qVml1WFahW2ZSsHupPsm+WXCGjRcCX+NotP/kKt9xO0aoIZ7Eke/fMTTB8O4riUr2AQC88eqbNMXoJXZR0xHhbOmWxpaaV1pEVuC9Vw3nLjiyjkTopmPNFpdV6ABgZzuPdvUFdBO6P7PWrq3jKYouDNrHHZ2x4nVOTmqTvIMcb9mduPRPH1ssZnHQO4hvBQAQ8vPsqmZZhmWkgAY9HC7XGhpJniIb0YrvFF1wppRhVZvdrlFZhd7cynQFAndeElWtx+XCwQjj1ooKpDt13zTE4y1zysfEEqnmRMOGgwWmTsz5vziZepbihfkIbIKNxZRd5BkAAMNID+Zo86LK6EQBDy8uciayBv311xs5T7490NMN1t5b6t2HTXQZo02D2bP8j0cCkqwOY+Pyq5Mnh+iGxsMHol2ZUaZpNfwjc4gYzq3pZsrGIE5TuBgAgKRYJSzLX/BRl6XP7XNoulZY7GYOZA36y1v3UwKXEFGOTkWcxvwLkYGkZIZihtrbFon8wulxvxIOTzXS6SCy2ZwspvE4QRRObSVyjBlEFNG60i2morV3nHQXzVKKg2MXOp9MqvMuAPggG4m/U9DpDgUfOnR4P7kEEs+raiN2VsAUGPo1CLPt4eljPgQA2LGjwhg8eNTtYKGhozxjDn6Wfsg4ZFVt5lmHGzCjAgAQgxSxGF8Q9IXVu4qkjDWUvX7Shqz66cmkQgIBf6CvVrphGFZV1ckC1aNNLQGaGIxO1Bv9JO/zm9CEwsIJRdPerKj47HVQol0chilaT6YcZmub2EYIwQSII8+e3+PaxbVrm3jJNO73upwvhbNYCUEDAFw5aRQxDSqvr1a612Pz/u0AsACgdqXgWoEXvs0S9ftz5gzq5sKueuLq0BfbRhMu0Ayrs8QvFzU1w5bK51vzTELkqpl5qfCsns1jX+TEdI7m36mcMbYx61nGIIC+o0NrA0IQ5Di3DADg9QqFaktcAABVVamvmyb39NrFX2s6n7bpDEYURU4QUzeBYkytSxqEMYn2NEc+Z+Gm6bzAts6fNf7tbOuLAQAqKsp0KW0mQlXvUn2x0i0LW/8yxZ2ZuWj/bUDR3rWLR9Seb9uupEaIZbUkU1qMRipuOppyIkSdtz0PrW3i5yyse8ghsl3zZ5W9mAt9T/nppSWmH7yy2BfQTc3Qnvn3Q9diojzidtK7etI2zyUOIUQ8QVGqNChIgk4HN5Lh+fOy56EVHxUr0cgihuJ2Vc3ODfBT5gUAQDch6fKUeuEC6sQv+kFKMZ5Ip3QfC2SzoSZ65DUwlGVLJdBx7yBawQAjLROCNgpt+bJ2v16w6etd8czXaYaqWV45sTGX+p6CzjIetS0iCbkGvuL1Rm7HRvM6iRcqTT15q8/t+6Qn7Qs9nWtLA+xRjrG1dyUy0xDh/BwvntW8zFmwYzBirH8zTCvq8VPzwxVlUq51pj97ICXaaAp8ALAvlxOINEqP2TnQedPgNGxA+IHSeI++tMqr//SPv3/+2EdbYxJd+dA3g8nlD3a/bvqCLaU0oG9rpuqnafw/z867+FPfFw3dkuUOmnWMyOXgM6s+mKWlpZ28zf1JV5f0uI215l5Mf0533laVyOpzf9lZ9cDSnc12kTUSnYZpYTVIY8Lpiv7a6oU3b+7rGPAUdEbk2zRFuSVXAz+yuH6OqiuJpx+76S+LV+/2YB1LT8+dsPVi+uyIxooZGv15qKf4d/s7Do1KqJZfVtTIkIDvlXmzruqAS0ROQU+4+Q5np2kL1RE6nOW69AcX7XoYsKU/+8TNzwEAzL3/mi4A+OvF9qsoRpHgZNunTy9IA8A2uETllMu4pqJMZ2ichG07/Vk1KU9snqXpWmTpI+Of6e2+RR5G8jzfAZe4dKt7UQz1UFw1xwBAr5+5DK1t4qW2zseA4L3LKsf9VzaU0XTijSZj+y916N0O78pqpp7BaGxvD7LgtwdKEifalzEc2pot4IQQihBWcHOZ9n4F/ffhb7QLLEtCq3YM7a0B5j6166aurug8ClDNk7PHZ+334qrWNHhYmpirw1PT/cq8AADwHFOnSuYPAeCiDtaF6gjdvmHr3apuDmMxnrek8vqsvs25LUJ7dEuKQT+Q094NEJ49vt5CWAut/HjKhXb6wJL6q2Mbtz8lsoxWXTl+7pLKG7L++mxNTRc4RF7rl9ABAOS08VtZ0qeFVn5yVU86W/i7hqJfL9jyINGtH3o94urqyvF/zFUZsmRCkQrmif4A/axbRQtqGq9MxtO/Bl1Zu+yxiTvP6QbO3zwMW8xdAGQII7CvLnvk2rdzrcgPHtozB7PG3hcWjV3fb6EDADz65NY8STPuZ0XekjXmDSkTP4oVyfQV+7yaTgoFxj48k0kMUlRLp2n6zZVVZdtRH2yCAAD86OE9T4GJ171QPWrPpQ79nG+f+NQWL3ikeue1GMh1qbh+ta4bBMUMLwKjg3Zw+1wFrr8+WzEiCgCwKtx3imCauJEgRGFAciNrmwj/w0d3r7n77jq+P8z3svj1lw9fPsLrKkkNGTJFG4CeKyUsI99hx1yuDzRkxab3F2ltlUUamW39Zb6XBXSnACUYs1a/uTMvB+jN7Uk6mdaPD0DPoXhc3lK7ncoMQM+hMAzO70xIxwag5zKspiAv3+vWB6DnSAghKJ1KS4lUuqO/zPly8F6w180XYdkhD6z0HMmaNYCTCT1y/XBrAHquZH96M02ILt1zzxB1AHquAqPkG2rQK6/GGBMYkAE5m/w/QHuDWhcmpxEAAAAASUVORK5CYII="
        fd, ttd_path = tempfile.mkstemp(suffix=".png")
        os.close(fd)
        with open(ttd_path, "wb") as f:
            f.write(base64.b64decode(TTD_OKTA_B64))
        pdf.image(ttd_path, x=135, y=pdf.get_y() + 2, w=30)
        os.remove(ttd_path)
    except Exception: pass
        
    pdf.ln(22)
    pdf.set_font("helvetica", "B", 10)
    pdf.cell(90, 6, "", 0, 0) 
    pdf.cell(90, 6, clean_text("(          Okta Pradika          )"), 0, 1, "C")
    pdf.set_font("helvetica", "", 10)
    pdf.cell(90, 6, "", 0, 0) 
    pdf.cell(90, 6, clean_text("Koordinator KUT Palangkaraya"), 0, 1, "C")
    
    pdf.add_page()
    pdf.set_font("helvetica", "B", 12)
    pdf.cell(0, 10, clean_text("LAMPIRAN DOKUMENTASI FOTO (FULL)"), ln=True, align="C")
    pdf.ln(5)
    
    def draw_photo_grid(url_list, title):
        urls = [u for u in url_list if isinstance(u, str) and u.startswith("http")]
        if not urls: return
        pdf.set_font("helvetica", "B", 11)
        pdf.cell(0, 8, clean_text(title), ln=True)
        max_img_w = 160 
        for u in urls:
            opt_url = u
            if "upload/v" in opt_url: opt_url = opt_url.replace("upload/v", "upload/c_limit,w_800,f_jpg/v")
            try:
                response = requests.get(opt_url, timeout=12)
                if response.status_code == 200:
                    img = Image.open(BytesIO(response.content))
                    if img.mode in ('RGBA', 'P', 'LA'): img = img.convert('RGB')
                    fd, temp_path = tempfile.mkstemp(suffix=".jpg")
                    os.close(fd)
                    img.save(temp_path, format="JPEG", quality=85)
                    w_orig, h_orig = img.size
                    calc_h = (max_img_w / w_orig) * h_orig
                    img_w_adj = max_img_w
                    if calc_h > 240: 
                        calc_h = 240
                        img_w_adj = (calc_h / h_orig) * w_orig
                    if pdf.get_y() + calc_h > 275: pdf.add_page()
                    pdf.image(temp_path, x=(210 - img_w_adj)/2, y=pdf.get_y(), w=img_w_adj)
                    pdf.set_y(pdf.get_y() + calc_h + 10)
                    if os.path.exists(temp_path): os.remove(temp_path)
            except Exception: pass
        pdf.ln(5)
        
    fisik_urls = (r.get('url_sites') or []) + (r.get('url_shadings') or []) + (r.get('extras_fisik') or [])
    draw_photo_grid(fisik_urls, "A. Dokumentasi Fisik & Lingkungan Site")
    panel_urls = []
    for p in r.get('panel_data', []):
        if isinstance(p.get('URL_Before'), list): panel_urls.extend(p.get('URL_Before'))
        elif p.get('URL_Before'): panel_urls.append(p.get('URL_Before'))
        panel_urls.extend(p.get('URLs_Before') or [])
        if isinstance(p.get('URL_After'), list): panel_urls.extend(p.get('URL_After'))
        elif p.get('URL_After'): panel_urls.append(p.get('URL_After'))
        panel_urls.extend(p.get('URLs_After') or [])
    panel_urls.extend(r.get('extras_panel') or [])
    draw_photo_grid(panel_urls, "B. Dokumentasi Modul Surya (SPS)")
    elek_urls = (r.get('url_rects') or []) + (r.get('url_gensets') or []) + (r.get('extras_elektrikal') or [])
    draw_photo_grid(elek_urls, "C. Dokumentasi Rectifier & Mesin Genset")
    bat_urls = []
    for b in r.get('battery_data', []): bat_urls.extend(b.get('URL_Fotos') or ( [b.get('URL_Foto')] if b.get('URL_Foto') else [] ))
    bat_urls.extend(r.get('url_grds') or [])
    bat_urls.extend(r.get('extras_baterai') or [])
    draw_photo_grid(bat_urls, "D. Dokumentasi Bank Baterai & Grounding")

    try: return bytes(pdf.output())
    except Exception:
        out = pdf.output(dest='S')
        return out.encode('latin-1', 'ignore') if isinstance(out, str) else out

# -------------------------------------------------------------------------
# CUSTOM CSS (Global & Desain Dashboard Profesional)
# -------------------------------------------------------------------------
st.markdown("""
    <style>
        [data-testid="stSidebar"] { display: none !important; }
        [data-testid="collapsedControl"] { display: none !important; }
        
        h1, h2, h3 { color: var(--primary-color) !important; font-family: 'Segoe UI', sans-serif; }
        div[data-testid="stExpander"] details {
            border: 1px solid rgba(128, 128, 128, 0.3);
            border-left: 5px solid var(--primary-color);
            border-radius: 12px;
            background-color: var(--secondary-background-color);
            margin-bottom: 12px;
            box-shadow: 0 4px 10px rgba(0,0,0,0.06);
            transition: all 0.25s ease;
        }
        div[data-testid="stExpander"] details:hover {
            border-color: var(--primary-color);
            box-shadow: 0 6px 16px rgba(0,0,0,0.14);
        }
        div[data-testid="stExpander"] summary {
            font-size: 16px !important;
            font-weight: 700 !important;
            color: var(--text-color) !important;
            padding: 12px 16px;
        }
        
        .stButton>button { 
            border-radius: 10px; 
            font-weight: 800 !important; 
            transition: all 0.3s; 
            border: 1px solid var(--primary-color); 
        }
        .stButton>button:hover { 
            transform: translateY(-2px); 
            box-shadow: 0 5px 15px rgba(0,0,0,0.2); 
            filter: brightness(1.1);
        }
        
        .footer-okta { text-align: center; padding: 25px; margin-top: 50px; color: var(--text-color); font-size: 15px; border-top: 1px solid var(--primary-color); background-color: var(--secondary-background-color); border-radius: 10px; opacity: 0.85; }
        .footer-okta span { color: var(--primary-color); font-weight: 800; letter-spacing: 1px; font-size: 16px; }
        
        .login-box { border: 2px solid var(--primary-color); padding: 36px 28px; border-radius: 18px; background-color: var(--secondary-background-color); text-align: center; height: 100%; box-shadow: 0 10px 25px rgba(0,0,0,0.08); transition: all 0.3s ease; }
        .login-box:hover { transform: translateY(-4px); box-shadow: 0 15px 32px rgba(0,0,0,0.14); }
        .login-title { color: var(--text-color); font-size: 24px; font-weight: 800; margin-bottom: 12px; }
        .login-desc { color: var(--text-color); font-size: 15px; margin-bottom: 24px; line-height: 1.6; opacity: 0.9; }
        
        .area-card {
            border: 1px solid rgba(128,128,128,0.25);
            border-top: 6px solid var(--primary-color);
            padding: 32px 26px;
            border-radius: 18px;
            background-color: var(--secondary-background-color);
            text-align: center;
            box-shadow: 0 10px 24px rgba(0,0,0,0.08);
            margin-bottom: 14px;
            transition: all 0.3s ease;
            height: 100%;
        }
        .area-card:hover {
            transform: translateY(-4px);
            box-shadow: 0 14px 30px rgba(0,0,0,0.14);
        }
        .area-badge {
            display: inline-block;
            padding: 6px 16px;
            border-radius: 30px;
            font-weight: 800;
            font-size: 13px;
            letter-spacing: 0.8px;
            margin-bottom: 14px;
            background-color: rgba(0, 123, 255, 0.15);
            color: #007BFF;
            border: 1px solid #007BFF;
        }
    </style>
""", unsafe_allow_html=True)

# -------------------------------------------------------------------------
# SISTEM LOGIN & ROLE AKSES 
# -------------------------------------------------------------------------
if 'role' not in st.session_state: st.session_state['role'] = None
if 'dashboard_area' not in st.session_state: st.session_state['dashboard_area'] = None

if st.session_state['role'] is None:
    
    st.markdown("""
        <style>
        div.stButton > button {
            min-height: 75px !important;
            font-size: 20px !important;
            font-weight: 900 !important;
            text-transform: uppercase;
            border-radius: 12px !important;
            letter-spacing: 1px;
        }
        button[kind="primary"] {
            background-color: #FF3333 !important;
            color: #FFFFFF !important;
            border: 2px solid #CC0000 !important;
            box-shadow: 0 6px 12px rgba(255, 51, 51, 0.4);
        }
        button[kind="secondary"] {
            background-color: #007BFF !important;
            color: #FFFFFF !important;
            border: 2px solid #0056b3 !important;
            box-shadow: 0 6px 12px rgba(0, 123, 255, 0.4);
        }
        div.stButton > button:hover {
            transform: scale(1.03) !important;
            filter: brightness(1.2) !important;
        }
        </style>
    """, unsafe_allow_html=True)
    
    st.markdown("<br>", unsafe_allow_html=True)
    render_header_logo(is_login=True) 
    st.markdown("<h1 style='text-align: center; font-size: 40px; margin-top: 5px;'>RPO NOP PLK & PBU Dashboard</h1>", unsafe_allow_html=True)
    st.markdown("<p style='text-align: center; font-size: 18px; margin-bottom: 45px; letter-spacing: 1px; opacity: 0.85;'>Report Online,Dasboard Minitoring dan evidance taksforce</p>", unsafe_allow_html=True)
    col_v, col_l, col_a = st.columns([3, 3, 3])
    
    with col_v:
        st.markdown("""
            <div class='login-box' style='padding-bottom: 20px;'>
                <div style='font-size: 55px; margin-bottom: 15px;'>👷‍♂️</div>
                <div class='login-title'>Mode Lapangan</div>
                <div class='login-desc'>Pilih area operasional Anda agar langsung diarahkan ke form pendataan yang tepat.</div>
            </div>
        """, unsafe_allow_html=True)
        st.write("")
        if st.button("🔴 PANGKALAN BUN", type="primary", use_container_width=True):
            st.session_state['role'] = 'Lapangan'
            st.session_state['area_lapangan'] = 'Pangkalan Bun'
            st.session_state['dashboard_area'] = None
            st.session_state['active_menu'] = None
            st.rerun()
        if st.button("🔵 PALANGKARAYA", type="secondary", use_container_width=True):
            st.session_state['role'] = 'Lapangan'
            st.session_state['area_lapangan'] = 'Palangkaraya'
            st.session_state['dashboard_area'] = None
            st.session_state['active_menu'] = None
            st.rerun()

    with col_l:
        st.markdown("""
            <div class='login-box' style='padding-bottom: 20px; border-color: #FF9800;'>
                <div style='font-size: 55px; margin-bottom: 15px;'>📦</div>
                <div class='login-title'>Gudang / Logistik</div>
                <div class='login-desc'>Akses eksklusif pengelolaan stok, in/out barang, sparepart, dan pelaporan SPMS.</div>
            </div>
        """, unsafe_allow_html=True)
        st.write("")
        st.write("")
        st.write("")
        if st.button("📦 MASUK LOGISTIK", type="secondary", use_container_width=True):
            st.session_state['role'] = 'Logistik'
            st.session_state['area_lapangan'] = 'Logistik'
            st.session_state['dashboard_area'] = None
            st.session_state['active_menu'] = None
            st.rerun()
            
    with col_a:
        st.markdown("""
            <div class='login-box'>
                <div style='font-size: 55px; margin-bottom: 15px;'>🔐</div>
                <div class='login-title'>Mode Admin</div>
                <div class='login-desc'>Akses khusus operasional untuk revisi data lapangan, dan analisa grafik Power.</div>
            </div>
        """, unsafe_allow_html=True)
        st.write("")
        admin_pass = st.text_input("Kata Sandi Admin:", type="password", placeholder="Masukkan Sandi...", key="pwd_login")
        if st.button("🔓 LOGIN ADMIN", type="primary", use_container_width=True):
            if admin_pass == "KUT2027":
                st.session_state['role'] = 'Admin'
                st.session_state['dashboard_area'] = None
                st.session_state['active_menu'] = None
                st.rerun()
            else:
                st.error("❌ Kata sandi salah! Akses ditolak.")
                
    st.markdown("<br><br>", unsafe_allow_html=True)
    st.markdown("<div style='text-align: center; font-size: 13px; opacity: 0.7;'>Secured System Application & Database Management<br><b>Created By Okta Pradika © 2026</b></div>", unsafe_allow_html=True)
    st.stop()


# -------------------------------------------------------------------------
# FUNGSI PEMBACA MASTER DATA EXCEL (NAMA FME & SITE ID)
# -------------------------------------------------------------------------
@st.cache_data(ttl=3600)
def load_master_lists():
    fme_plk, fme_pbu, site_plk, site_pbu = [], [], [], []
    try:
        df1 = pd.read_excel("NIK NOP PLK.xlsx")
        if 'NAMA' in df1.columns: fme_plk = df1['NAMA'].dropna().astype(str).tolist()
    except: pass
    
    try:
        df2 = pd.read_excel("NAMA NOP PANGKALANBUN.xlsx")
        if 'NAMA' in df2.columns: fme_pbu = df2['NAMA'].dropna().astype(str).tolist()
    except: pass
    
    try:
        df3 = pd.read_excel("Site ID Sidata Palangkaraya.xlsx")
        if 'Site ID' in df3.columns: site_plk = df3['Site ID'].dropna().astype(str).tolist()
    except: pass
    
    try:
        df4 = pd.read_excel("Site ID Sidata Pangkalanbun.xlsx")
        if 'Site ID' in df4.columns: site_pbu = df4['Site ID'].dropna().astype(str).tolist()
    except: pass
    
    return sorted(fme_plk), sorted(fme_pbu), sorted(site_plk), sorted(site_pbu)


# -------------------------------------------------------------------------
# KONFIGURASI UPLOAD CLOUDINARY
# -------------------------------------------------------------------------
cloudinary.config( cloud_name = "fxm61tjv", api_key = "624877324969231", api_secret = "LIFO6pfEg9fOM3nbsY8FBbVTpSI", secure = True )

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

def upload_datalog(file_obj, folder_name="prev_datalog"):
    if file_obj is not None:
        try:
            response = cloudinary.uploader.upload(file_obj.getvalue(), folder=folder_name, resource_type="raw", public_id=file_obj.name)
            return response.get('secure_url')
        except Exception as e: return None
    return None

def tampilkan_grid_foto(url_data, caption=""):
    if not url_data: return
    urls = []
    if isinstance(url_data, str) and url_data.startswith("http"): urls = [url_data]
    elif isinstance(url_data, list): urls = [u for u in url_data if isinstance(u, str) and u.startswith("http")]
        
    if urls:
        if caption: st.markdown(f"*{caption}*")
        img_html = '<div style="display: grid; grid-template-columns: repeat(auto-fit, minmax(280px, 1fr)); gap: 12px; margin-bottom: 15px;">'
        for u in urls:
            img_html += f'<div style="background-color: var(--secondary-background-color); border: 1px solid var(--primary-color); border-radius: 8px; padding: 6px; text-align: center;"><a href="{u}" target="_blank"><img src="{u}" style="max-width: 100%; height: auto; max-height: 400px; object-fit: contain; border-radius: 6px;"></a></div>'
        img_html += '</div>'
        st.markdown(img_html, unsafe_allow_html=True)


# -------------------------------------------------------------------------
# SETUP DATABASE GOOGLE SHEETS
# -------------------------------------------------------------------------
SHEET_ID = "1HvgVicTWwO4RMQI6ZR3Mu3IgGicwjcLZl9mDN1auvJU"

def connect_gsheets(sheet_name="Report Preventive"):
    try:
        creds_json = st.secrets["gcp_json"]
        creds_dict = json.loads(creds_json)
        scopes = ["https://www.googleapis.com/auth/spreadsheets", "https://www.googleapis.com/auth/drive"]
        creds = Credentials.from_service_account_info(creds_dict, scopes=scopes)
        client = gspread.authorize(creds)
        return client.open_by_key(SHEET_ID).worksheet(sheet_name)
    except Exception as e: return None

@st.cache_data(ttl=600) 
def fetch_data_from_gsheets():
    all_data = []
    
    sheet_prev = connect_gsheets("Report Preventive")
    if sheet_prev:
        data = sheet_prev.get_all_values()
        for idx, row in enumerate(data[1:]):
            if len(row) >= 6: 
                try: 
                    item = json.loads(row[5])
                    if item.get('kategori') == 'SiUPDATE': item['kategori'] = 'SPS'
                    item['_sheet'] = "Report Preventive"
                    item['_row'] = idx + 2
                    all_data.append(item)
                except: pass
                
    sheet_pbu = connect_gsheets("Report PBU")
    if sheet_pbu:
        data_pbu = sheet_pbu.get_all_values()
        for idx, row in enumerate(data_pbu[1:]):
            if len(row) >= 6: 
                try: 
                    item = json.loads(row[5])
                    item['_sheet'] = "Report PBU"
                    item['_row'] = idx + 2
                    all_data.append(item)
                except: pass

    # --- TAMBAHAN UNTUK SHEET PROJECT ---
    sheet_proj = connect_gsheets("Project")
    if sheet_proj:
        data_proj = sheet_proj.get_all_values()
        for idx, row in enumerate(data_proj[1:]):
            if len(row) >= 6: 
                try: 
                    item = json.loads(row[5])
                    item['_sheet'] = "Project"
                    item['_row'] = idx + 2
                    all_data.append(item)
                except: pass
                
    return all_data

if 'laporan_db' not in st.session_state: 
    st.session_state['laporan_db'] = fetch_data_from_gsheets()

# -------------------------------------------------------------------------
# WA GENERATOR PANGKALAN BUN (Otomatis Teks)
# -------------------------------------------------------------------------
def generate_wa_pbu(r):
    text = f"*SITE VISIT REPORT*\n\n"
    text += f"> INFORMASI SITE\n"
    text += f"- FME NAME : {r.get('pb_te', '-')}\n"
    text += f"- Site ID : {r.get('pb_site_id', '-')}\n"
    text += f"- Site Name : {r.get('site_name', '-')}\n"
    text += f"- Ticket# : {r.get('pb_ticket', '-')}\n"
    text += f"- Tanggal On Site : {r.get('timestamp', '-').split(' ')[0]}\n\n"
    text += f"> Catuan Power Utama : {r.get('pb_power', '-')}\n\n"
    
    rects = r.get('pb_rectifiers', [])
    text += f"> DATA RECTIFIER\n"
    text += f"- Jumlah Rectifier: {len(rects)}\n\n"
    
    for idx, rect in enumerate(rects):
        text += f"Rectifier #{idx+1}\n"
        text += f"- Tipe Rectifier : {rect.get('tipe', '-')}\n"
        text += f"- Type Controler : {rect.get('controler', '-')}\n"
        text += f"- Type Module : {rect.get('module', '-')}\n"
        text += f"- Jumlah Battery : {rect.get('jml_batt', '-')}\n"
        text += f"- Jenis Batt OK: {rect.get('batt_ok', '-')}\n"
        text += f"- Jenis Batt NOK: {rect.get('batt_nok', '-')}\n"
        text += f"- Jumlah Module : (OK : {rect.get('mod_ok', '-')}; NOK : {rect.get('mod_nok', '-')})\n"
        text += f"- BCL : {rect.get('bcl', '-')}\n"
        text += f"- Load Current: {rect.get('load', '-')}\n"
        text += f"- DC Voltage:  {rect.get('dc_v', '-')}\n"
        text += f"- LVD Prio:  {rect.get('lvd_p', '-')}\n"
        text += f"- LVD Non Prio : {rect.get('lvd_np', '-')}\n"
        text += f"- EQP LVD Prio: {rect.get('eqp_p', '-')}\n"
        text += f"- EQP LVD Non Prio: {rect.get('eqp_np', '-')}\n"
        text += f"- Battery Backup Time : {rect.get('backup', '-')}\n\n"

    en = r.get('pb_enva', {})
    text += f"> * ENVA Setelah Simulasi (Bukan Cabut Kabel)\n"
    text += f"- Main Fail : {en.get('main_fail', '-')}\n"
    text += f"- Rectifier Fail : {en.get('rect_fail', '-')}\n"
    text += f"- Low Batt : {en.get('low_batt', '-')}\n"
    text += f"- High Temp : {en.get('high_temp', '-')}\n"
    text += f"- Genset Run : {en.get('gen_run', '-')}\n"
    text += f"- Genset Fail : {en.get('gen_fail', '-')}\n\n"

    kel = r.get('pb_kelistrikan', {})
    text += f"> DATA KELISTRIKAN\n"
    text += f"- Daya KWH: {kel.get('daya', '-')}\n"
    text += f"- Jumlah Phasa: {kel.get('phasa', '-')}\n"
    text += f"- Tegangan Pengukuran R-S-T terhadap N: {kel.get('v_rst_n', '-')}\n"
    text += f"- Tegangan Pengukuran Phasa RS-RT-ST: {kel.get('v_rs_rt_st', '-')}\n"
    text += f"- Pengukuran Arus R-S-T: {kel.get('arus_rst', '-')}\n"
    text += f"- Tegangan Pengukuran G-N: {kel.get('v_gn', '-')}\n\n"

    acp = r.get('pb_acpdb', {})
    text += f"> DATA ACPDB\n"
    text += f"- Ada ACPDB: {acp.get('ada', '-')}\n"
    text += f"- Tegangan Pengukuran R-S-T terhadap N: {acp.get('v_rst_n', '-')}\n"
    text += f"- Tegangan Pengukuran Phasa RS-RT-ST: {acp.get('v_rs_rt_st', '-')}\n"
    text += f"- Pengukuran Arus R-S-T: {acp.get('arus_rst', '-')}\n"
    text += f"- Tegangan Pengukuran G-N: {acp.get('v_gn', '-')}\n\n"

    gen = r.get('pb_genset', {})
    text += f"> DATA GENSET\n"
    text += f"- Ada Genset: {gen.get('ada', '-')}\n"
    text += f"- Capacity Genset : {gen.get('cap', '-')}\n"
    text += f"- Genset Status : {gen.get('status', '-')}\n"
    text += f"- Config Running : {gen.get('config', '-')}\n"
    text += f"- Kapasitas Tangki : {gen.get('tangki', '-')}\n"
    text += f"- Running Hour : {gen.get('rh', '-')}\n"
    text += f"- Sisa BBM : {gen.get('bbm', '-')}\n\n"

    text += f"> Activity Report\n"
    text += f"1. Power & CME Activity\n- " + ("\n- ".join(r.get('pb_act_power', [])) if isinstance(r.get('pb_act_power'), list) else str(r.get('pb_act_power','-'))) + "\n\n"
    text += f"2. Transport Activity\n- " + ("\n- ".join(r.get('pb_act_trans', [])) if isinstance(r.get('pb_act_trans'), list) else str(r.get('pb_act_trans','-'))) + "\n\n"
    text += f"3. OSA Impact Service Activity\n- " + ("\n- ".join(r.get('pb_act_osa', [])) if isinstance(r.get('pb_act_osa'), list) else str(r.get('pb_act_osa','-'))) + "\n\n"

    text += f"> Temuan/ Catatan :\n"
    temuan_lines = str(r.get('action', '-')).split('\n')
    for line in temuan_lines:
        if line.strip(): text += f"{line}\n"
        
    text += f"\n\n🔗 Foto & Dokumentasi lengkap tersimpan di Dashboard SPS Portal."
    return text


# -------------------------------------------------------------------------
# NAVIGASI UTAMA (MENU KLIK TANPA SIDEBAR & TOMBOL KEMBALI)
# -------------------------------------------------------------------------
if 'active_menu' not in st.session_state:
    st.session_state['active_menu'] = None

# --- KODE WATERMARK GLOBAL (MUNCUL DI MENU UTAMA & DALAM FORM) ---
bg_base64 = get_base64_of_bin_file(LOGO_UTAMA)
if bg_base64:
    st.markdown(
        f"""
        <style>
        .watermark-kut {{
            position: fixed;
            top: 50%;
            left: 50%;
            transform: translate(-50%, -50%);
            opacity: 0.05;
            z-index: 0;
            pointer-events: none;
        }}
        .watermark-kut img {{
            width: 750px;
            max-width: 90vw;
            filter: grayscale(100%);
        }}
        </style>
        <div class="watermark-kut">
            <img src="data:image/webp;base64,{bg_base64}">
        </div>
        """, unsafe_allow_html=True
    )
# -----------------------------------------------------------------

if st.session_state['active_menu'] is None:
    render_header_logo()
    role_badge = "🟢 MODE ADMIN" if st.session_state['role'] == 'Admin' else f"👷‍♂️ TIM LAPANGAN ({st.session_state.get('area_lapangan', 'Umum').upper()})"
    st.markdown("<h1 style='text-align: center; margin-bottom: 5px;'>⚡ Menu Utama Operasional</h1>", unsafe_allow_html=True)
    st.markdown(f"<p style='text-align: center; opacity: 0.85; font-size: 16px; margin-bottom: 25px;'>Status Akses: <b>{role_badge}</b> — Silakan klik menu di bawah ini untuk masuk ke fitur</p>", unsafe_allow_html=True)

    col_m1, col_m2 = st.columns(2)
    with col_m1:
        if st.session_state['role'] == 'Admin':
            st.markdown("""
                <div class='area-card' style='border-top-color: #007BFF;'>
                    <div style='font-size: 46px; margin-bottom: 8px;'>📝</div>
                    <div class='login-title'>Form Preventive Check</div>
                    <div class='login-desc'>Input laporan Preventive Maintenance Site SPS, Site Reguler Palangkaraya, dan Site Visit Pangkalan Bun.</div>
                </div>
            """, unsafe_allow_html=True)
            if st.button("📝 BUKA FORM PREVENTIVE CHECK", type="primary", use_container_width=True, key="m_form_adm"):
                st.session_state['active_menu'] = "📝 Form Preventive Check"
                st.rerun()
        else:
            st.markdown("""
                <div class='area-card' style='border-top-color: #007BFF;'>
                    <div style='font-size: 46px; margin-bottom: 8px;'>📸</div>
                    <div class='login-title'>Form Pelaporan (Lapangan)</div>
                    <div class='login-desc'>Input data pengerjaan lapangan, pengukuran parameter kelistrikan AC/DC, dan upload foto eviden wajib.</div>
                </div>
            """, unsafe_allow_html=True)
            if st.button("📸 BUKA FORM PELAPORAN LAPANGAN", type="primary", use_container_width=True, key="m_form_lap"):
                st.session_state['active_menu'] = "📸 Form Pelaporan (Lapangan)"
                st.rerun()

        st.markdown("<br>", unsafe_allow_html=True)
        st.markdown("""
            <div class='area-card' style='border-top-color: #00E676;'>
                <div style='font-size: 46px; margin-bottom: 8px;'>📈</div>
                <div class='login-title'>Monitoring Improvement</div>
                <div class='login-desc'>Pantau grafik Kurva S (Plan vs Actual), progress tim lapangan (Done/On Progress), dan Live Editor Tracker.</div>
            </div>
        """, unsafe_allow_html=True)
        if st.button("📈 BUKA MONITORING IMPROVEMENT", type="secondary", use_container_width=True, key="m_mon"):
            st.session_state['active_menu'] = "📈 Monitoring Improvement"
            st.rerun()

    with col_m2:
        st.markdown("""
            <div class='area-card' style='border-top-color: #FF3333;'>
                <div style='font-size: 46px; margin-bottom: 8px;'>📊</div>
                <div class='login-title'>Hasil Laporan & Dashboard</div>
                <div class='login-desc'>Lihat rekapitulasi laporan NOP Palangkaraya & Pangkalan Bun, unduh Excel/PPTX/PDF, dan bagikan ke WhatsApp.</div>
            </div>
        """, unsafe_allow_html=True)
        if st.button("📊 BUKA HASIL LAPORAN & DASHBOARD", type="primary", use_container_width=True, key="m_dash"):
            st.session_state['active_menu'] = "📊 Hasil Laporan & Dashboard"
            st.rerun()

        st.markdown("<br>", unsafe_allow_html=True)
        st.markdown("""
            <div class='area-card' style='border-top-color: #FFC107;'>
                <div style='font-size: 46px; margin-bottom: 8px;'>⚡</div>
                <div class='login-title'>Dapot Power Hasil PMS</div>
                <div class='login-desc'>Executive Dashboard untuk compare bulanan tegangan & beban AC/DC pada site kunjungan rutin maupun 1x kunjungan.</div>
            </div>
        """, unsafe_allow_html=True)
        if st.button("⚡ BUKA DAPOT POWER HASIL PMS", type="secondary", use_container_width=True, key="m_dapot"):
            st.session_state['active_menu'] = "⚡ Dapot Power Hasil PMS"
            st.rerun()

    st.markdown("<hr style='margin: 30px 0;'>", unsafe_allow_html=True)
    c_out1, c_out2, c_out3, c_out4 = st.columns(4)
    with c_out1:
        st.markdown("""
            <div class='area-card' style='border-top-color: #9C27B0;'>
                <div style='font-size: 42px; margin-bottom: 8px;'>📅</div>
                <div class='login-title' style='font-size: 20px;'>Daily Tracker PLK</div>
                <div class='login-desc' style='font-size: 13px;'>To Do List harian tim Onsite, update status pengerjaan, dan pencatatan FME.</div>
            </div>
        """, unsafe_allow_html=True)
        if st.button("📅 BUKA DAILY TRACKER", type="primary", use_container_width=True, key="m_tracker"):
            st.session_state['active_menu'] = "📅 Daily Tracker Palangkaraya"
            st.rerun()
            
    with c_out2:
        st.markdown("""
            <div class='area-card' style='border-top-color: #FFD700; background: linear-gradient(145deg, #0f172a, #1e293b);'>
                <div style='font-size: 42px; margin-bottom: 8px;'>💎</div>
                <div class='login-title' style='font-size: 20px; color: #FFD700 !important;'>Executive Project</div>
                <div class='login-desc' style='font-size: 13px; color: #cbd5e1 !important;'>Dashboard Khusus PO & SVA: Interactive Timeline, Margin & Live Update.</div>
            </div>
        """, unsafe_allow_html=True)
        if st.button("💎 BUKA EXECUTIVE PROJECT", type="primary", use_container_width=True, key="m_proj_dash"):
            st.session_state['active_menu'] = "💎 Executive Project Dashboard"
            st.rerun()

    with c_out3:
        if st.session_state['role'] in ['Admin', 'Logistik']:
            st.markdown("""
                <div class='area-card' style='border-top-color: #FF9800;'>
                    <div style='font-size: 42px; margin-bottom: 8px;'>📦</div>
                    <div class='login-title' style='font-size: 20px;'>Logistik & SPMS</div>
                    <div class='login-desc' style='font-size: 13px;'>Manajemen Inbound/Outbound stok material, oli, filter genset, dan gudang.</div>
                </div>
            """, unsafe_allow_html=True)
            if st.button("📦 BUKA LOGISTIK", type="secondary", use_container_width=True, key="m_log"):
                st.session_state['active_menu'] = "📦 Logistik & SPMS"
                st.rerun()
                
    with c_out4:
        st.markdown("""
            <div class='area-card' style='border-top-color: #9E9E9E;'>
                <div style='font-size: 42px; margin-bottom: 8px;'>🚪</div>
                <div class='login-title' style='font-size: 20px;'>Log Out</div>
                <div class='login-desc' style='font-size: 13px;'>Keluar dari sesi akun Anda dengan aman untuk melindungi data operasional.</div>
            </div>
        """, unsafe_allow_html=True)
        if st.button("🚪 KELUAR AKUN", type="primary", use_container_width=True, key="btn_logout_main"):
            st.session_state['role'] = None
            st.session_state['dashboard_area'] = None
            st.session_state['active_menu'] = None
            st.session_state['proj_dash_unlocked'] = False
            st.session_state['proj_unlocked'] = False
            st.rerun()
    st.stop()
else:
    menu = st.session_state['active_menu']
    c_top1, c_top2, c_top3 = st.columns([2.2, 5.6, 2.2])
    with c_top1:
        if st.button("⬅️ KEMBALI KE MENU UTAMA", type="secondary", use_container_width=True, key="btn_back_main"):
            st.session_state['active_menu'] = None
            st.session_state['dashboard_area'] = None
            st.rerun()
    with c_top2:
        st.markdown(
            f"<div style='text-align: center; padding: 10px 16px; border-radius: 10px; background-color: var(--secondary-background-color); border: 1px solid var(--primary-color); font-weight: 800; font-size: 16px;'>📍 Menu Aktif: {menu}</div>",
            unsafe_allow_html=True
        )
    with c_top3:
        if st.button("🚪 Keluar Akun (Log Out)", use_container_width=True, key="btn_logout_top"):
            st.session_state['role'] = None
            st.session_state['dashboard_area'] = None
            st.session_state['active_menu'] = None
            st.rerun()
    st.markdown("<hr style='margin: 15px 0 25px 0;'>", unsafe_allow_html=True)


# =========================================================================
# MENU 1: FORM PENGISIAN LENGKAP
# =========================================================================
if menu in ["📝 Form Preventive Check", "📸 Form Pelaporan (Lapangan)"]:
    
    # Ambil Data Master untuk Dropdown
    fme_plk_list, fme_pbu_list, site_plk_list, site_pbu_list = load_master_lists()
    
    if menu == "📸 Form Pelaporan (Lapangan)":
        if st.session_state.get('area_lapangan') == 'Pangkalan Bun':
            st.markdown("<h1 style='text-align: center;'>📸 Form Pelaporan (Pangkalan Bun)</h1>", unsafe_allow_html=True)
            st.info("Sistem telah mengunci akses Anda ke Form Pangkalan Bun berdasarkan pilihan di layar login.")
            standar_form = "Pangkalan Bun"
            nop_area_default = "Pangkalan Bun"
        else:
            st.markdown("<h1 style='text-align: center;'>📸 Form Pelaporan (Standar/Palangkaraya)</h1>", unsafe_allow_html=True)
            st.info("Sistem telah mengunci akses Anda ke Form Standar berdasarkan pilihan di layar login.")
            standar_form = "Standar"
            nop_area_default = "Palangkaraya"
    else:
        st.markdown("<h1 style='text-align: center;'>⚡ Form Preventive Maintenance (Admin)</h1>", unsafe_allow_html=True)
        pilih_form = st.radio("Pilih Format Laporan:", ["Standar (Palangkaraya/Reguler)", "Pangkalan Bun"], horizontal=True)
        standar_form = "Pangkalan Bun" if "Pangkalan Bun" in pilih_form else "Standar"
        nop_area_default = "Palangkaraya"

    # -------------------------------------------------------------------------
    # BLOK A: FORM STANDAR PALANGKARAYA (DIPISAH SPS VS REGULER TAPI 1 DATABASE)
    # -------------------------------------------------------------------------
    if standar_form == "Standar":
        
        st.info("💡 Data dan Lampiran (Foto & Datalog) Anda akan dienkripsi dan dikirim langsung ke Cloud Server.")
        site_category = st.radio("Kategori Site (Wajib Pilih):", ["SPS", "Site Reguler", "Project"], horizontal=True)
        st.markdown("<hr style='margin:10px 0;'>", unsafe_allow_html=True)
        
        if site_category == "SPS":
            tabs = st.tabs(["📌 1. Info Site", "☀️ 2. SPS Panel", "🔌 3. PLN & Recti", "⛽ 4. Genset & BBM", "🔋 5. Baterai & Gnd", "📤 6. Upload & Submit"])

            with tabs[0]:
                st.markdown("#### Kategori Kunjungan: SPS")
                c1, c2 = st.columns(2)
                with c1:
                    if site_plk_list:
                        site_name = st.selectbox("Pilih Site ID", [""] + site_plk_list)
                    else:
                        site_name = st.text_input("Nama / ID Site", placeholder="Contoh: BTS-PKY-001")
                        
                    opsi_nop = ["Palangkaraya", "Pangkalan Bun", "Tarakan", "Pontianak", "Lainnya"]
                    idx_default = opsi_nop.index(nop_area_default) if nop_area_default in opsi_nop else 0
                    nop_area = st.selectbox("NOP Area", opsi_nop, index=idx_default)
                    check_date = st.date_input("Tanggal Pengecekan", value=datetime.date.today())
                with c2:
                    if fme_plk_list:
                        technician_name = st.multiselect("Pilih Nama Pelaksana / FME (Bisa >1):", fme_plk_list)
                    else:
                        tek_str = st.text_input("Nama Pelaksana / Teknisi")
                        technician_name = [tek_str] if tek_str else []
                        
                    weather = st.selectbox("Kondisi Cuaca", ["Cerah", "Berawan", "Hujan", "Kabut Asap menutup cahaya matahari", "Lembab tidak ada panas", "Badai"])
                st.divider()
                sc1, sc2 = st.columns(2)
                with sc1:
                    site_condition = st.selectbox("Kondisi Halaman", ["Bersih & Aman", "Banyak Rumput", "Genangan Air"])
                    tower_condition = st.selectbox("Fisik Tower", ["Aman", "Berkarat", "Rusak"])
                with sc2:
                    site_photos = st.file_uploader("📸 Foto View Site (Bisa >1)", accept_multiple_files=True, key="spics")

            with tabs[1]:
                st.markdown("### 📝 Kelengkapan Data SPS & SCC")
                c_pnl1, c_pnl2, c_pnl3 = st.columns(3)
                with c_pnl1: status_panel = st.selectbox("Status Panel", ["Normal", "Kotor", "Retak / Pecah", "Terbakar", "Kritis / Kosong"])
                with c_pnl2: jml_panel_rusak = st.number_input("Jumlah Panel Rusak", min_value=0, value=0)
                with c_pnl3: merk_panel = st.text_input("Type Merk Panel", placeholder="Contoh: Canadian Solar 550Wp")

                c_scc1, c_scc2, c_scc3 = st.columns(3)
                with c_scc1: status_scc = st.selectbox("Status SCC", ["Normal", "Alarm", "Rusak", "Tidak Ada"])
                with c_scc2: jml_scc_nok = st.number_input("Jumlah SCC NOK", min_value=0, value=0)
                with c_scc3: merk_scc = st.text_input("Type/Merk SCC", placeholder="Contoh: Huawei / Shoto")
                st.divider()

                shading_status = st.selectbox("Status Shading?", ["Aman", "Sedikit Shading", "Kritis"])
                shading_photos = st.file_uploader("📸 Foto Shading (Bisa >1)", accept_multiple_files=True, key="shd")
                st.divider()
                
                konfigurasi_panel = st.radio("Metode Pengukuran SPS:", ["Individu (Satuan)", "Seri / String (Grup)"])
                panel_data = []

                if konfigurasi_panel == "Individu (Satuan)":
                    num_panels = st.number_input("Jumlah Panel", min_value=1, value=24)
                    for i in range(int(num_panels)):
                        with st.expander(f"⚙️ Panel #{i+1}", expanded=(i==0)):
                            c1, c2, c3 = st.columns([1, 1, 1])
                            with c1:
                                voc = st.number_input(f"Voc [V]", value=21.5, key=f"v_{i}")
                                isc = st.number_input(f"Isc [A]", value=5.2, key=f"i_{i}")
                                p_cond = st.selectbox(f"Fisik", ["Baik", "Kotor", "Retak", "Delaminasi"], key=f"c_{i}")
                            with c2: pb = st.file_uploader(f"📸 KONDISI PANEL", key=f"pb_{i}")
                            with c3: pa = st.file_uploader(f"✨ PENGUKURAN", key=f"pa_{i}")
                            panel_data.append({"tipe": "Individu", "id": f"Panel #{i+1}", "voc": voc, "isc": isc, "kondisi": p_cond, "foto_before_obj": pb, "foto_after_obj": pa})
                else: 
                    num_strings = st.number_input("Jumlah String", min_value=1, value=4)
                    for i in range(int(num_strings)):
                        with st.expander(f"🔌 String Seri #{i+1}", expanded=(i==0)):
                            c1, c2 = st.columns([1, 1.5])
                            with c1:
                                qty = st.number_input(f"Isi Panel per Seri", value=6, key=f"sq_{i}")
                                voc = st.number_input(f"Total Voc [V]", value=129.0, key=f"sv_{i}")
                                isc = st.number_input(f"Isc [A]", value=5.2, key=f"si_{i}")
                                p_cond = st.selectbox(f"Kondisi", ["Baik", "Kotor", "Retak"], key=f"sc_{i}")
                            with c2:
                                pb = st.file_uploader(f"📸 KONDISI PANEL (Bisa >1)", accept_multiple_files=True, key=f"spb_{i}")
                                pa = st.file_uploader(f"✨ PENGUKURAN (Bisa >1)", accept_multiple_files=True, key=f"spa_{i}")
                            panel_data.append({"tipe": "Seri", "id": f"String #{i+1}", "qty": qty, "voc": voc, "isc": isc, "kondisi": p_cond, "foto_before_objs": pb, "foto_after_objs": pa})

            with tabs[2]:
                st.markdown("### 🔌 Pengecekan Jaringan PLN & PSB")
                c_psb1, c_psb2 = st.columns(2)
                with c_psb1: possibility_psb = st.selectbox("Possibility PSB (PLN)", ["Bisa Dilakukan", "Sulit / Terlalu Jauh", "Sudah Tersambung", "Tidak Memungkinkan"])
                with c_psb2: jarak_pln = st.text_input("Estimasi Jarak tower ke jaringan PLN", placeholder="Contoh: 150 Meter / 2 KM")
                st.divider()

                c1, c2 = st.columns(2)
                with c1:
                    pln_status = st.selectbox("Status PLN", ["Normal", "Padam", "Power Perusahaan", "Tidak Ada PLN"])
                    jb_enclosure = st.selectbox("Box DC & Seal", ["Bersih", "Bocor", "Sarang Serangga"])
                    jb_spd = st.selectbox("Arrester SPD", ["Normal", "Rusak"])
                with c2:
                    rect_brand = st.text_input("Merek Rectifier", placeholder="Huawei / ZTE")
                    rect_out_v = st.number_input("Tegangan Output Rectifier [V]", value=53.5)
                    total_load_a = st.number_input("Total Beban BTS [A]", value=18.2)
                    rect_alarm = st.selectbox("Alarm Rectifier", ["No Alarm", "Ada Alarm"])
                    rect_photos = st.file_uploader("📸 Foto Rectifier (Bisa >1)", accept_multiple_files=True, key="rect_pics")

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
                    genset_photos = st.file_uploader("📸 Foto Genset & BBM (Bisa >1)", accept_multiple_files=True, key="gen_pics")

            with tabs[4]:
                st.markdown("### 📝 Kelengkapan Data Baterai")
                c_batm1, c_batm2, c_batm3 = st.columns(3)
                with c_batm1: status_battery = st.selectbox("Status Total Baterai", ["Normal", "Degradasi (Drop)", "Rusak / Menggelembung", "Kritis / Hilang"])
                with c_batm2: jml_bat_rusak = st.number_input("Jumlah Baterai Rusak", min_value=0, value=0)
                with c_batm3: tipe_battery = st.text_input("Type / Merk Baterai", placeholder="Contoh: VRLA Shoto 100Ah")
                st.divider()

                b1, b2 = st.columns(2)
                with b1:
                    num_bat = st.number_input("Jumlah Baterai (Pengukuran)", min_value=1, value=4)
                    bat_data = []
                    for j in range(int(num_bat)):
                        with st.expander(f"🔋 Baterai #{j+1}", expanded=(j==0)):
                            bc1, bc2 = st.columns(2)
                            with bc1:
                                bv = st.number_input(f"Voltase [V]", value=12.2, key=f"bv_{j}")
                                bt = st.number_input(f"Load Charging [A]", value=28.0, key=f"bt_{j}")
                            with bc2:
                                bc = st.selectbox(f"Kondisi", ["Normal", "Bengkak", "Rusak tidak bisa charging", "Over Charging", "Korosi"], key=f"bc_{j}")
                                bp = st.file_uploader(f"Foto Bat #{j+1}", accept_multiple_files=True, key=f"bp_{j}")
                            bat_data.append({"id": f"Baterai #{j+1}", "voltase": bv, "suhu": bt, "kondisi": bc, "foto_objs": bp})
                with b2:
                    earth_resistance = st.number_input("Tahanan Grounding [Ohm]", value=2.1)
                    grd_cable = st.selectbox("Kabel Grounding", ["Kuat", "Kendor", "Putus"])
                    grd_photos = st.file_uploader("📸 Foto Grounding (Bisa >1)", accept_multiple_files=True, key="grd")

            with tabs[5]:
                st.markdown("📂 **Upload Datalog (Bisa .csv, .xlsx, .zip, dll):**")
                datalog_files = st.file_uploader("Semua file tersimpan utuh", accept_multiple_files=True, key="datalog_all")
                st.divider()
                
                action_taken = st.text_area("🔧 Rincian Pekerjaan & Action di Lapangan:")
                sparepart_needed = st.text_input("📦 Penggantian Sparepart:")
                remark = st.text_area("📝 Remark / Catatan Khusus Laporan:")
                final_status = st.radio("Status Akhir Site:", ["Normal", "Minor Issue", "Major/Critical", "Update Progress"])

                if st.button("🚀 UPLOAD & SINKRONKAN DATA SPS", type="primary", use_container_width=True):
                    if not site_name:
                        st.error("⚠️ Mohon pilih Nama Site!")
                    elif not technician_name:
                        st.error("⚠️ Mohon pilih Pelaksana / FME!")
                    else:
                        with st.spinner("⏳ Sedang memproses dan mengamankan data ke Cloud..."):
                            url_sites = upload_multiple_images(site_photos, "prev_view")
                            url_shadings = upload_multiple_images(shading_photos, "prev_shade")
                            url_rects = upload_multiple_images(rect_photos, "prev_rect")
                            url_gensets = upload_multiple_images(genset_photos, "prev_genset")
                            url_grds = upload_multiple_images(grd_photos, "prev_grd")
                            
                            datalog_urls = []
                            if datalog_files:
                                for df in datalog_files:
                                    du = upload_datalog(df, "prev_datalog")
                                    if du: datalog_urls.append({"name": df.name, "url": du})

                            p_res = []
                            for p in panel_data:
                                if p.get("tipe") == "Individu":
                                    p_res.append({"Tipe": "Individu", "Panel": p["id"], "Voc": p["voc"], "Isc": p["isc"], "Kondisi": p["kondisi"], "URL_Before": upload_image(p["foto_before_obj"]), "URL_After": upload_image(p["foto_after_obj"])})
                                else:
                                    p_res.append({"Tipe": "Seri", "Panel": p["id"], "Qty": p["qty"], "Voc": p["voc"], "Isc": p["isc"], "Kondisi": p["kondisi"], "URLs_Before": upload_multiple_images(p["foto_before_objs"]), "URLs_After": upload_multiple_images(p["foto_after_objs"])})
                            
                            b_res = []
                            for b in bat_data: b_res.append({"Baterai": b["id"], "Voltase": b["voltase"], "Suhu": b["suhu"], "Kondisi": b["kondisi"], "URL_Fotos": upload_multiple_images(b["foto_objs"])})

                            tek_str_final = ", ".join(technician_name)
                            report_dict = {
                                "timestamp": datetime.datetime.now().strftime("%Y-%m-%d %H:%M:%S"),
                                "site_name": site_name, 
                                "kategori": "SPS", 
                                "nop": nop_area, 
                                "teknisi": tek_str_final,
                                "status": final_status, 
                                "action": action_taken, 
                                "sparepart": sparepart_needed,
                                
                                "status_panel": status_panel, "jml_panel_rusak": jml_panel_rusak, "merk_panel": merk_panel,
                                "status_scc": status_scc, "jml_scc_nok": jml_scc_nok, "merk_scc": merk_scc,
                                "status_battery": status_battery, "jml_bat_rusak": jml_bat_rusak, "tipe_battery": tipe_battery,
                                "possibility_psb": possibility_psb, "jarak_pln": jarak_pln, "remark": remark,
                                
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
                            
                            sheet = connect_gsheets("Report Preventive")
                            if sheet:
                                row_data = [
                                    report_dict.get('timestamp', ''), report_dict.get('site_name', ''),
                                    report_dict.get('nop', ''), report_dict.get('teknisi', ''),
                                    report_dict.get('status', ''), json.dumps(report_dict)
                                ]
                                sheet.insert_row(row_data, index=2)
                                st.cache_data.clear()
                                st.session_state['laporan_db'].append(report_dict)
                                st.success("✅ BERHASIL! Data SPS telah diamankan ke Baris Teratas Spreadsheet.")
                            else:
                                st.error("⚠️ Gagal menyambung ke Server Database.")

        elif site_category == "Site Reguler":
            # -------------------------------------------------------------------------
            # BLOK A.2: FORM KHUSUS SITE REGULER BARU (PALANGKARAYA - WAJIB LENGKAP 100%)
            # -------------------------------------------------------------------------
            tabs_reg = st.tabs(["📌 1. Info & Problem", "⚡ 2. Action & Parameter", "📤 3. Submit"])
            cond_data = {}
            cond_files = {}
            
            with tabs_reg[0]:
                st.markdown("#### Kategori Kunjungan: Site Reguler (Maintenance/Troubleshooting)")
                c1, c2 = st.columns(2)
                with c1:
                    if site_plk_list:
                        site_name_reg = st.selectbox("Pilih Site ID (Wajib):", [""] + site_plk_list, key="reg_site")
                    else:
                        site_name_reg = st.text_input("Nama / ID Site (Wajib):", key="reg_site")
                        
                    opsi_nop = ["Palangkaraya", "Pangkalan Bun", "Tarakan", "Pontianak", "Lainnya"]
                    idx_default = opsi_nop.index(nop_area_default) if nop_area_default in opsi_nop else 0
                    nop_area_reg = st.selectbox("NOP Area:", opsi_nop, index=idx_default, key="reg_nop")
                with c2:
                    if fme_plk_list:
                        teknisi_reg = st.multiselect("Pilih Nama FME (Wajib, Bisa >1):", fme_plk_list, key="reg_fme")
                    else:
                        t_str = st.text_input("Nama FME (Wajib):", key="reg_fme")
                        teknisi_reg = [t_str] if t_str else []
                    
                    role_fme_reg = st.selectbox("Role FME:", ["TE", "MBP", "CME"], key="reg_role")
                    check_date_reg = st.date_input("Tanggal Kunjungan:", value=datetime.date.today(), key="reg_date")
                
                st.divider()
                st.markdown("### 🚨 Problem Alarm")
                prob_alarm = st.selectbox(
                    "Pilih Problem Utama:", 
                    ["Down power", "Fliker power", "Fliker Site Down", "Transmisi Problem", "cell down"]
                )
                foto_prob = st.file_uploader("📸 Foto Problem Alarm (WAJIB DILAMPIRKAN):", accept_multiple_files=True, key="f_prob")

            with tabs_reg[1]:
                st.markdown("### 🔧 Action & Perbaikan")
                
                # --- AUTO-BACA & SIMPAN ACTION BARU DARI DATABASE LAMA ---
                base_act_opts = [
                    "Replace RRU", "Replace Optik", "Replace SFP", "Replace Module APR", 
                    "Teg PLN Drop pelaporan PLN", "Reset Hardware", "MCB Trip", "Add Baterai",
                    "Cleaning Fan BTS", "Cleaning FAN Recti", "Cleaning FAN Perangkat TRM", "Cleaning FAN Shelter",
                    "Troubleshoot Genset / ATS"
                ]
                existing_actions = set(base_act_opts)
                if 'laporan_db' in st.session_state and st.session_state['laporan_db']:
                    for r_db in st.session_state['laporan_db']:
                        if r_db.get('kategori') == 'Site Reguler' and r_db.get('action_reguler'):
                            acts = r_db.get('action_reguler')
                            if isinstance(acts, list):
                                for a in acts: existing_actions.add(a)
                            elif isinstance(acts, str):
                                existing_actions.add(acts)
                
                act_opts = sorted(list(existing_actions))
                act_sel = st.multiselect("Pilih Action Dikerjakan (WAJIB PILIH, Bisa Multiple):", act_opts)
                
                # Input Manual (Otomatis digabung ke daftar terpilih & history)
                custom_act = st.text_input("➕ Tambah Action Manual (Ketik di sini jika tidak ada di pilihan atas):", placeholder="Contoh: Ganti Filter Oli Mesin...")
                if custom_act:
                    if custom_act not in act_sel:
                        act_sel.append(custom_act)
                        
                foto_act = st.file_uploader("📸 Foto Action / Perbaikan Diluar Pengukuran (WAJIB DILAMPIRKAN):", accept_multiple_files=True, key="f_act")

                # Pengukuran Power Wajib Muncul Untuk Semua Action
                show_cond_power = True 
                show_cond_genset = any("Genset" in item or "ATS" in item for item in act_sel)
                
                if show_cond_power or show_cond_genset:
                    st.divider()
                    st.warning("⚠️ **WAJIB DIISI 100%:** Seluruh kolom nilai ukur dan foto eviden di bawah ini wajib dilengkapi satu per satu tanpa terkecuali!")
                
                if show_cond_power:
                    st.markdown("#### ⚡ A. Pengukuran AC/DC & Eviden Foto")
                    cc1, cc2, cc3 = st.columns(3)
                    with cc1:
                        cond_data['v_rn'] = st.text_input("Teg AC R-N [V] *", placeholder="Contoh: 220")
                        cond_files['f_v_rn'] = st.file_uploader("📸 Foto Eviden Teg AC R-N *", key="f_v_rn")
                        cond_data['i_r'] = st.text_input("Load AC R [A] *", placeholder="Contoh: 12.5")
                        cond_files['f_i_r'] = st.file_uploader("📸 Foto Eviden Load AC R *", key="f_i_r")
                    with cc2:
                        cond_data['v_sn'] = st.text_input("Teg AC S-N [V] *", placeholder="Contoh: 221")
                        cond_files['f_v_sn'] = st.file_uploader("📸 Foto Eviden Teg AC S-N *", key="f_v_sn")
                        cond_data['i_s'] = st.text_input("Load AC S [A] *", placeholder="Contoh: 11.8")
                        cond_files['f_i_s'] = st.file_uploader("📸 Foto Eviden Load AC S *", key="f_i_s")
                    with cc3:
                        cond_data['v_tn'] = st.text_input("Teg AC T-N [V] *", placeholder="Contoh: 219")
                        cond_files['f_v_tn'] = st.file_uploader("📸 Foto Eviden Teg AC T-N *", key="f_v_tn")
                        cond_data['i_t'] = st.text_input("Load AC T [A] *", placeholder="Contoh: 10.4")
                        cond_files['f_i_t'] = st.file_uploader("📸 Foto Eviden Load AC T *", key="f_i_t")

                    st.markdown("**Parameter DC & Modul**")
                    cc7, cc8 = st.columns(2)
                    with cc7:
                        cond_data['load_dc'] = st.text_input("Load DC Total [A] *", placeholder="Contoh: 45.2")
                        cond_files['f_ldc'] = st.file_uploader("📸 Foto Eviden Load DC *", key="f_ldc")
                        cond_data['set_bcl'] = st.text_input("Settingan BCL [%] *", placeholder="Contoh: 100%")
                        cond_files['f_bcl'] = st.file_uploader("📸 Foto Eviden Settingan BCL *", key="f_bcl")
                    with cc8:
                        cond_data['set_cap'] = st.text_input("Settingan Capacity [Ah] *", placeholder="Contoh: 400Ah")
                        cond_files['f_cap'] = st.file_uploader("📸 Foto Eviden Settingan Cap *", key="f_cap")
                        cond_data['mod_rusak'] = st.text_input("Module Kontroler Rusak * (Isi 'Aman/0' jika normal)")
                        cond_files['f_mod'] = st.file_uploader("📸 Foto Eviden Module *", key="f_mod")

                if show_cond_genset:
                    st.markdown("#### ⚙️ B. Pengecekan Mesin Genset / ATS & Eviden Foto")
                    cg1, cg2, cg3 = st.columns(3)
                    with cg1:
                        cond_data['gen_v'] = st.text_input("Tegangan Out Genset [V] *", placeholder="Cth: 220")
                        cond_files['f_gen_v'] = st.file_uploader("📸 Foto Tegangan Genset *", key="f_gen_v")
                        cond_data['gen_i'] = st.text_input("Load Arus Genset [A] *", placeholder="Cth: 25.4")
                        cond_files['f_gen_i'] = st.file_uploader("📸 Foto Arus Genset *", key="f_gen_i")
                    with cg2:
                        cond_data['gen_ats'] = st.selectbox("Status Panel ATS *", ["Auto Normal", "Manual", "Rusak / Bypass"])
                        cond_files['f_gen_ats'] = st.file_uploader("📸 Foto Panel ATS *", key="f_gen_ats")
                        cond_data['gen_issue'] = st.selectbox("Issue Mesin Genset *", ["Normal / Aman", "Hunting / Suara Kasar", "Overheat", "Bocor Oli / BBM", "Mati Total / Tidak Bisa Start"])
                        cond_files['f_gen_issue'] = st.file_uploader("📸 Foto Fisik Mesin *", key="f_gen_issue")
                    with cg3:
                        cond_data['gen_hm'] = st.text_input("Hour Meter (HM) *", placeholder="Cth: 1250.5")
                        cond_files['f_gen_hm'] = st.file_uploader("📸 Foto Hour Meter *", key="f_gen_hm")
                        cond_data['gen_bbm'] = st.text_input("Estimasi Sisa BBM [Liter] *", placeholder="Cth: 150")
                        cond_files['f_gen_bbm'] = st.file_uploader("📸 Foto Tangki/Sisa BBM *", key="f_gen_bbm")

                st.divider()
                cond_data['note_action'] = st.text_area("📝 NOTE / REMARK Action Lapangan (WAJIB DIISI) *:", placeholder="Jelaskan secara spesifik hasil pengerjaan, part yang diganti, atau kendala di lapangan...")

            with tabs_reg[2]:
                st.markdown("### 📤 Konfirmasi & Submit Laporan")
                final_status_reg = st.radio("Status Akhir Site:", ["Normal", "Minor Issue", "Major/Critical", "Update Progress"], key="stat_reg")
                
                if st.button("🚀 UPLOAD & SINKRONKAN DATA REGULER", type="primary", use_container_width=True):
                    # Validasi Ketat 100% Tanpa Kurang 1 Apapun
                    item_kurang = []
                    if not str(site_name_reg).strip(): item_kurang.append("Site ID")
                    if not teknisi_reg: item_kurang.append("Nama FME")
                    if not foto_prob: item_kurang.append("Foto Problem Alarm")
                    if not act_sel: item_kurang.append("Pilihan Action Dikerjakan")
                    if not foto_act: item_kurang.append("Foto Action / Perbaikan Utama")
                    
                    wajib_teks = []
                    wajib_foto = []
                    
                    if show_cond_power:
                        wajib_teks.extend([
                            ('v_rn', 'Isian Teg AC R-N'), ('v_sn', 'Isian Teg AC S-N'), ('v_tn', 'Isian Teg AC T-N'),
                            ('i_r', 'Isian Load AC R'), ('i_s', 'Isian Load AC S'), ('i_t', 'Isian Load AC T'),
                            ('load_dc', 'Isian Load DC Total'), ('set_bcl', 'Isian Settingan BCL'),
                            ('set_cap', 'Isian Capacity'), ('mod_rusak', 'Isian Module Kontroler')
                        ])
                        wajib_foto.extend([
                            ('f_v_rn', 'Foto Teg AC R-N'), ('f_v_sn', 'Foto Teg AC S-N'), ('f_v_tn', 'Foto Teg AC T-N'),
                            ('f_i_r', 'Foto Load AC R'), ('f_i_s', 'Foto Load AC S'), ('f_i_t', 'Foto Load AC T'),
                            ('f_ldc', 'Foto Load DC'), ('f_bcl', 'Foto BCL'), ('f_cap', 'Foto Capacity'), ('f_mod', 'Foto Module')
                        ])
                        
                    if show_cond_genset:
                        wajib_teks.extend([
                            ('gen_v', 'Isian Tegangan Output Genset'), ('gen_i', 'Isian Arus Load Genset'),
                            ('gen_hm', 'Isian Hour Meter (HM)'), ('gen_bbm', 'Isian Sisa BBM (Liter)')
                        ])
                        wajib_foto.extend([
                            ('f_gen_v', 'Foto Eviden Tegangan Genset'), ('f_gen_i', 'Foto Eviden Arus Genset'),
                            ('f_gen_ats', 'Foto Eviden Panel ATS'), ('f_gen_issue', 'Foto Eviden Fisik Mesin'),
                            ('f_gen_hm', 'Foto Eviden Hour Meter'), ('f_gen_bbm', 'Foto Eviden Tangki BBM')
                        ])
                        
                    wajib_teks.append(('note_action', 'Isian Note / Remark Action'))
                    
                    for k_teks, label_teks in wajib_teks:
                        if not str(cond_data.get(k_teks, '')).strip(): item_kurang.append(label_teks)
                    for k_foto, label_foto in wajib_foto:
                        if cond_files.get(k_foto) is None: item_kurang.append(label_foto)

                    if item_kurang:
                        st.error(f"⚠️ **LAPORAN DITOLAK!** Tim wajib melengkapi seluruh data dan foto sesuai Action yang dipilih. Data yang masih kosong:\n- " + "\n- ".join(item_kurang))
                    else:
                        with st.spinner("⏳ Sedang memproses dan mengamankan seluruh foto eviden ke Cloud..."):
                            url_probs = upload_multiple_images(foto_prob, "reg_prob")
                            url_acts = upload_multiple_images(foto_act, "reg_act")
                            
                            if show_cond_power:
                                for k_f in ['v_rn', 'v_sn', 'v_tn', 'i_r', 'i_s', 'i_t', 'ldc', 'bcl', 'cap', 'mod']:
                                    cond_data[f'url_{k_f}'] = upload_image(cond_files.get(f'f_{k_f}'), "reg_param")
                                cond_data['volt_ac_rst'] = f"{cond_data['v_rn']} / {cond_data['v_sn']} / {cond_data['v_tn']}"
                                cond_data['load_ac_rst'] = f"{cond_data['i_r']} / {cond_data['i_s']} / {cond_data['i_t']}"
                                
                            if show_cond_genset:
                                for k_g in ['gen_v', 'gen_i', 'gen_ats', 'gen_issue', 'gen_hm', 'gen_bbm']:
                                    cond_data[f'url_{k_g}'] = upload_image(cond_files.get(f'f_{k_g}'), "reg_param")
                            
                            tek_str_reg = ", ".join(teknisi_reg)
                            report_dict = {
                                "timestamp": datetime.datetime.now().strftime("%Y-%m-%d %H:%M:%S"),
                                "site_name": site_name_reg,
                                "kategori": "Site Reguler",
                                "nop": nop_area_reg,
                                "teknisi": tek_str_reg,
                                "role_fme": role_fme_reg,
                                "status": final_status_reg,
                                "action": f"Problem: {prob_alarm} | Action: {', '.join(act_sel)} | Note: {cond_data['note_action']}",
                                "sparepart": "-",
                                "problem_alarm": prob_alarm,
                                "action_reguler": act_sel,
                                "url_problem": url_probs,
                                "url_action": url_acts,
                                "cond_data": cond_data,
                                "pln_status": "-", "rect_brand": "-", "rect_out_v": "-", 
                                "total_load": cond_data.get('load_dc', '-'),
                                "url_sites": url_probs + url_acts, 
                                "extras_fisik": [], "extras_panel": [], "extras_baterai": [], "extras_elektrikal": []
                            }
                            
                            sheet = connect_gsheets("Report Preventive")
                            if sheet:
                                row_data = [
                                    report_dict.get('timestamp', ''), report_dict.get('site_name', ''),
                                    report_dict.get('nop', ''), report_dict.get('teknisi', ''),
                                    report_dict.get('status', ''), json.dumps(report_dict)
                                ]
                                sheet.insert_row(row_data, index=2)
                                st.cache_data.clear()
                                st.session_state['laporan_db'].append(report_dict)
                                st.success("✅ BERHASIL! Data Laporan Site Reguler telah diamankan ke Baris Teratas Spreadsheet.")
                            else:
                                st.error("⚠️ Gagal menyambung ke Server Database.")

        elif site_category == "Project":
            # -------------------------------------------------------------------------
            # BLOK A.3: FORM KHUSUS PROJECT (PO & SVA) DENGAN PASSWORD
            # -------------------------------------------------------------------------
            if 'proj_unlocked' not in st.session_state:
                st.session_state['proj_unlocked'] = False
                
            if not st.session_state['proj_unlocked'] and st.session_state['role'] != 'Admin':
                with st.container(border=True):
                    st.markdown("<h3 style='text-align: center; color: #9C27B0;'>🔒 Akses Form Project Terkunci</h3>", unsafe_allow_html=True)
                    st.markdown("<p style='text-align: center;'>Form ini khusus untuk tim pengerjaan Project (PO / SVA).</p>", unsafe_allow_html=True)
                    proj_pass = st.text_input("Masukkan Kata Sandi Khusus Project:", type="password", key="pwd_proj")
                    if st.button("🔓 BUKA FORM PROJECT", type="primary", use_container_width=True):
                        import base64
                        if base64.b64encode(proj_pass.encode()).decode() == "WXVpMjAyNSM=":
                            st.session_state['proj_unlocked'] = True
                            st.rerun()
                        else:
                            st.error("❌ Kata sandi salah! Akses ditolak.")
            else:
                tabs_proj = st.tabs(["📌 1. Info & Multi-Site PO/SVA", "📸 2. Dokumentasi (Opsional Jika Belum Mulai)", "📤 3. Submit Project"])
                proj_data = {}
                proj_files = {}

                with tabs_proj[0]:
                    st.markdown("#### Kategori Kunjungan: Project Lapangan")
                    jenis_project = st.selectbox("Pilih Jenis Project:", ["PO", "SVA"])
                    proj_data['jenis_project'] = jenis_project

                    if fme_plk_list:
                        teknisi_proj = st.multiselect("Nama FME / Tim (Wajib):", fme_plk_list, key="proj_fme")
                    else:
                        t_str = st.text_input("Nama FME / Tim (Wajib):", key="proj_fme")
                        teknisi_proj = [t_str] if t_str else []

                    opsi_nop = ["Palangkaraya", "Pangkalan Bun", "Tarakan", "Pontianak", "Lainnya"]
                    nop_area_proj = st.selectbox("NOP Area:", opsi_nop, index=0, key="proj_nop")
                    st.divider()

                    if jenis_project == "PO":
                        st.markdown("### 📋 Detail Kontrak Purchase Order (PO)")
                        c_po1, c_po2 = st.columns(2)
                        with c_po1:
                            proj_data['no_po'] = st.text_input("Nomor PO *")
                            proj_data['mitra'] = st.text_input("Mitra Pemberi Kerja *", placeholder="Cth: Telkomsel / Huawei / Kisel...")
                        with c_po2:
                            nilai_po_str = st.text_input("Total Nilai PO (Rp) *", value="0", help="Ketik angka saja, misal: 50000000")
                            proj_data['nilai_po'] = nilai_po_str
                            tipe_pekerjaan = st.selectbox("Program Utama PO *", ["Relokasi NE", "Instalasi Imacro", "Perbaikan / Modernisasi", "Lainnya"])
                            proj_data['tipe_pekerjaan'] = tipe_pekerjaan

                        st.markdown("### 🗺️ Tabel Multi-Site, Rincian Pekerjaan, Timeline & Alokasi Biaya")
                        st.info("💡 **Klik baris di bawah (atau klik tanda `+` di baris paling bawah tabel)** untuk menambahkan banyak rute site sekaligus, termasuk **Dana Tersupport (Cair)** dan **Alasan jika Mandek**!")
                        
                        df_po = pd.DataFrame([
                            {
                                "Site Asal": "",
                                "Site Tujuan": "",
                                "Pekerjaan Spesifik": "",
                                "Budget Per Site": 0,
                                "Dana Tersupport": 0,
                                "Start Date": datetime.date.today().strftime("%Y-%m-%d"),
                                "End Date": datetime.date.today().strftime("%Y-%m-%d"),
                                "Progress (%)": 0,
                                "Status": "Not Started",
                                "Kategori Kendala": "Lancar / Normal",
                                "Remark": ""
                            }
                        ])
                        
                        edited_po = st.data_editor(
                            df_po, 
                            num_rows="dynamic", 
                            use_container_width=True,
                            column_config={
                                "Site Asal": st.column_config.TextColumn("Site Asal (Dari)"),
                                "Site Tujuan": st.column_config.TextColumn("Site Tujuan (Ke)"),
                                "Pekerjaan Spesifik": st.column_config.TextColumn("Pekerjaan di Site"),
                                "Budget Per Site": st.column_config.NumberColumn("Total Budget Site (Rp)", min_value=0, default=0, step=100000),
                                "Dana Tersupport": st.column_config.NumberColumn("Dana Tersupport / Cair (Rp)", min_value=0, default=0, step=100000),
                                "Start Date": st.column_config.TextColumn("Tgl Mulai (YYYY-MM-DD)"),
                                "End Date": st.column_config.TextColumn("Target Selesai (YYYY-MM-DD)"),
                                "Progress (%)": st.column_config.NumberColumn("Progress (%)", min_value=0, max_value=100, default=0),
                                "Status": st.column_config.SelectboxColumn("Status Site", options=["Not Started", "On Progress", "Done", "Pending Issue"]),
                                "Kategori Kendala": st.column_config.SelectboxColumn("Kenapa Mandek / Kendala", options=[
                                    "Lancar / Normal",
                                    "Mandek - Pending Keuangan (Dana Belum Cair)",
                                    "Mandek - Pending Material / Sparepart",
                                    "Mandek - Izin Site / Community Issue",
                                    "Mandek - Kendala Teknis / Cuaca"
                                ]),
                                "Remark": st.column_config.TextColumn("Detail Kendala / Catatan")
                            },
                            key="tabel_po_input"
                        )
                        
                        try:
                            val_po = float(str(nilai_po_str).replace('.', '').replace(',', ''))
                        except:
                            val_po = 0.0
                            
                        df_clean_po = edited_po.copy()
                        if not df_clean_po.empty:
                            df_clean_po["Budget Per Site"] = pd.to_numeric(df_clean_po["Budget Per Site"], errors='coerce').fillna(0.0)
                            df_clean_po["Dana Tersupport"] = pd.to_numeric(df_clean_po["Dana Tersupport"], errors='coerce').fillna(0.0)
                            df_clean_po["Progress (%)"] = pd.to_numeric(df_clean_po["Progress (%)"], errors='coerce').fillna(0.0)
                            
                        total_cost = df_clean_po["Budget Per Site"].sum() if not df_clean_po.empty else 0.0
                        total_sup = df_clean_po["Dana Tersupport"].sum() if not df_clean_po.empty else 0.0
                        pending_fin = max(total_cost - total_sup, 0.0)
                        margin = val_po - total_cost
                        avg_progress = df_clean_po["Progress (%)"].mean() if not df_clean_po.empty else 0.0
                        
                        cm1, cm2, cm3, cm4, cm5 = st.columns(5)
                        cm1.metric("💰 Total Nilai PO", f"Rp {val_po:,.0f}")
                        cm2.metric("💸 Total Budget Site", f"Rp {total_cost:,.0f}")
                        cm3.metric("✅ Dana Tersupport", f"Rp {total_sup:,.0f}")
                        cm4.metric("🚨 Pending Keuangan", f"Rp {pending_fin:,.0f}")
                        cm5.metric("📈 Estimasi Margin", f"Rp {margin:,.0f}", delta=f"{avg_progress:.0f}% Progress")
                        
                        proj_data['tabel_pekerjaan'] = df_clean_po.astype(str).to_dict('records')
                        proj_data['total_cost'] = float(total_cost)
                        proj_data['total_support'] = float(total_sup)
                        proj_data['pending_finance'] = float(pending_fin)
                        proj_data['margin'] = float(margin)
                        proj_data['avg_progress'] = float(avg_progress)

                        if tipe_pekerjaan == "Instalasi Imacro":
                            st.markdown("**📌 Checklist Milestone Imacro:**")
                            c_chk1, c_chk2 = st.columns(2)
                            with c_chk1:
                                p1 = st.checkbox("Dismantle")
                                p2 = st.checkbox("Instalasi")
                                p3 = st.checkbox("Penarikan Optik")
                            with c_chk2:
                                p4 = st.checkbox("Material Terkirim/Lengkap")
                                p5 = st.checkbox("PSB PLN")
                            proj_data['progress_imacro'] = [k for k, v in zip(["Dismantle", "Instalasi", "Penarikan Optik", "Material", "PSB PLN"], [p1, p2, p3, p4, p5]) if v]

                    elif jenis_project == "SVA":
                        st.markdown("### 📋 Detail Site Verification / Audit (SVA)")
                        c_sva1, c_sva2 = st.columns(2)
                        with c_sva1:
                            proj_data['site_sva'] = st.selectbox("Pilih Site ID *", [""] + site_plk_list, key="sva_site") if site_plk_list else st.text_input("Site ID *")
                            proj_data['no_tiket'] = st.text_input("Nomor Tiket SVA *")
                        with c_sva2:
                            tgl_sva_val = st.date_input("Tanggal Pengerjaan SVA *", value=datetime.date.today())
                            proj_data['tgl_sva'] = tgl_sva_val.strftime("%Y-%m-%d")
                            proj_data['status_sva'] = st.selectbox("Status SVA", ["Not Started", "On Progress", "Done", "Pending Issue"])

                        proj_data['temuan_sva'] = st.text_area("Temuan / Hasil Audit Lapangan:", placeholder="Catat detail temuan SVA di sini...")

                with tabs_proj[1]:
                    st.markdown("### 📸 Lampiran Dokumentasi Project")
                    st.info("💡 ** Catatan Penting:** Jika pekerjaan belum dimulai (*Not Started*) atau baru perencanaan PO, foto di bawah ini **BOLEH DIKOSONGKAN** dan dapat di-update menyusul di *Executive Project Dashboard*.")

                    if jenis_project == "PO":
                        foto_awal = st.file_uploader("📸 Foto Kondisi Awal (Before)", accept_multiple_files=True, key="p_awal")
                        foto_proses = st.file_uploader("📸 Foto Proses Pengerjaan (On Progress)", accept_multiple_files=True, key="p_pros")
                        foto_akhir = st.file_uploader("📸 Foto Hasil Akhir (After) / BAST", accept_multiple_files=True, key="p_akhir")
                        proj_files['foto_awal'] = foto_awal
                        proj_files['foto_proses'] = foto_proses
                        proj_files['foto_akhir'] = foto_akhir
                    else:
                        foto_sva_1 = st.file_uploader("📸 Foto View Site & Perangkat", accept_multiple_files=True, key="sva_1")
                        foto_sva_2 = st.file_uploader("📸 Foto Temuan SVA / Defect", accept_multiple_files=True, key="sva_2")
                        proj_files['foto_sva_1'] = foto_sva_1
                        proj_files['foto_sva_2'] = foto_sva_2

                with tabs_proj[2]:
                    st.markdown("### 📤 Konfirmasi & Submit Laporan Project")
                    proj_data['note_akhir'] = st.text_area("📝 Catatan Akhir Project:", placeholder="Kendala atau catatan penting lainnya...")
                    final_status_proj = st.radio("Status Keseluruhan Project:", ["Not Started", "On Progress", "Done", "Pending"], key="stat_proj")

                    if st.button("🚀 SIMPAN & SINKRONKAN PROJECT KE CLOUD", type="primary", use_container_width=True):
                        item_kurang = []
                        if not teknisi_proj: item_kurang.append("Nama FME/Tim")

                        if jenis_project == "PO":
                            if not proj_data.get('no_po'): item_kurang.append("Nomor PO")
                            if not proj_data.get('mitra'): item_kurang.append("Mitra Pemberi Kerja")
                            if final_status_proj == "Done" and not proj_files.get('foto_akhir') and not proj_files.get('foto_proses'):
                                item_kurang.append("Minimal Foto Proses / Akhir (Wajib karena status dipilih Done)")
                        else:
                            if not proj_data.get('site_sva'): item_kurang.append("Site ID SVA")
                            if not proj_data.get('no_tiket'): item_kurang.append("Nomor Tiket SVA")

                        if item_kurang:
                            st.error(f"⚠️ **LAPORAN DITOLAK!** Mohon lengkapi isian berikut:\n- " + "\n- ".join(item_kurang))
                        else:
                            with st.spinner("⏳ Sedang memproses dan mengamankan data Project ke Cloud..."):
                                url_awal, url_proses, url_akhir = [], [], []
                                url_sva1, url_sva2 = [], []

                                if jenis_project == "PO":
                                    url_awal = upload_multiple_images(proj_files.get('foto_awal'), "proj_po")
                                    url_proses = upload_multiple_images(proj_files.get('foto_proses'), "proj_po")
                                    url_akhir = upload_multiple_images(proj_files.get('foto_akhir'), "proj_po")
                                else:
                                    url_sva1 = upload_multiple_images(proj_files.get('foto_sva_1'), "proj_sva")
                                    url_sva2 = upload_multiple_images(proj_files.get('foto_sva_2'), "proj_sva")

                                proj_data['url_awal'] = url_awal
                                proj_data['url_proses'] = url_proses
                                proj_data['url_akhir'] = url_akhir
                                proj_data['url_sva1'] = url_sva1
                                proj_data['url_sva2'] = url_sva2

                                tek_str_proj = ", ".join(teknisi_proj)
                                nama_site_db = proj_data.get('site_sva') if jenis_project == "SVA" else f"PO: {proj_data.get('no_po')}"

                                report_dict = {
                                    "timestamp": datetime.datetime.now().strftime("%Y-%m-%d %H:%M:%S"),
                                    "site_name": nama_site_db,
                                    "kategori": "Project",
                                    "nop": nop_area_proj,
                                    "teknisi": tek_str_proj,
                                    "status": final_status_proj,
                                    "action": f"Project {jenis_project}",
                                    "sparepart": "-",
                                    "proj_data": proj_data,
                                    "url_sites": url_awal + url_sva1,
                                    "url_action": url_proses + url_akhir + url_sva2,
                                    "pln_status": "-", "rect_brand": "-", "rect_out_v": "-",
                                    "total_load": "-", "cond_data": {}
                                }

                                sheet = connect_gsheets("Project")
                                if sheet:
                                    row_data = [
                                        report_dict.get('timestamp', ''), report_dict.get('site_name', ''),
                                        report_dict.get('nop', ''), report_dict.get('teknisi', ''),
                                        report_dict.get('status', ''), json.dumps(report_dict, default=str)
                                    ]
                                    try:
                                        sheet.insert_row(row_data, index=2)
                                    except Exception:
                                        sheet.insert_row(["Timestamp", "Site Name", "NOP", "Teknisi", "Status", "Data JSON"], index=1)
                                        sheet.insert_row(row_data, index=2)
                                        
                                    st.cache_data.clear()
                                    st.session_state['laporan_db'] = fetch_data_from_gsheets()
                                    st.success("✅ BERHASIL! Data Project telah tersimpan dan siap dipantau di menu 💎 Executive Project.")
                                else:
                                    st.error("⚠️ Gagal menyambung ke Sheet 'Project'. Pastikan sheet bernama 'Project' sudah dibuat di Spreadsheet Utama Anda.")

    # -------------------------------------------------------------------------
    # BLOK B: FORM PANGKALAN BUN (LENGKAP DENGAN UPLOAD FOTO & AUTO WA)
    # -------------------------------------------------------------------------
    elif standar_form == "Pangkalan Bun":
        
        st.info("💡 **Format Khusus Site Visit Report (Pangkalan Bun).** Silakan isi parameter di bawah ini dan lampirkan foto dokumentasi pada setiap bagian. Seluruh foto akan tersimpan di Cloud & muncul di Dashboard.")
        tabs_pb = st.tabs(["📌 Info Site & Power", "🔌 Rectifier & ENVA", "⚡ Kelistrikan & ACPDB", "⛽ Genset & Log", "📤 Activity & Submit"])
        
        pb_payload = {}
        
        with tabs_pb[0]:
            st.markdown("### 📌 INFORMASI SITE")
            c1, c2 = st.columns(2)
            with c1:
                if fme_pbu_list:
                    pb_te = st.multiselect("FME NAME (Bisa >1):", fme_pbu_list, key="pb_te")
                else:
                    pt_str = st.text_input("FME NAME", key="pb_te")
                    pb_te = [pt_str] if pt_str else []
                    
                if site_pbu_list:
                    pb_sid = st.selectbox("Pilih Site ID", [""] + site_pbu_list, key="pb_sid")
                else:
                    pb_sid = st.text_input("Site ID", key="pb_sid")
                    
            with c2:
                pb_tick = st.text_input("Ticket#")
                pb_tgl = st.text_input("Tanggal On Site", value=datetime.date.today().strftime("%d %B %Y"))
                pb_catuan = st.selectbox("Catuan Power Utama", ["GENSET", "PLN", "SPS", "LAINNYA"])
                
            pb_foto_site = st.file_uploader("📸 Foto Site / Kondisi Umum (Bisa >1)", accept_multiple_files=True, key="pb_foto_site")
            
            tek_pb_str = ", ".join(pb_te) if pb_te else "-"
            pb_payload.update({"te_name": tek_pb_str, "site_id": pb_sid, "site_name": pb_sid, "ticket": pb_tick, "tgl_onsite": pb_tgl, "power_utama": pb_catuan})
            
        with tabs_pb[1]:
            st.markdown("### 🔌 DATA RECTIFIER")
            pb_jml_rect = st.number_input("Jumlah Rectifier", min_value=1, value=2)
            
            pb_recti_list = []
            for i in range(int(pb_jml_rect)):
                with st.expander(f"⚙️ Rectifier #{i+1}", expanded=(i==0)):
                    c1, c2 = st.columns(2)
                    with c1:
                        tr = st.text_input("Tipe Rectifier (ZTE, dll)", key=f"tr_{i}")
                        tc = st.text_input("Type Controler (Csu 501B, dll)", key=f"tc_{i}")
                        tm = st.text_input("Type Module (ZXD2000, dll)", key=f"tm_{i}")
                        jb = st.text_input("Jumlah Battery (Bank)", key=f"jb_{i}")
                        bok = st.text_input("Jenis Batt OK", key=f"bok_{i}")
                        bnok = st.text_input("Jenis Batt NOK", key=f"bnok_{i}")
                        jmo = st.text_input("Jumlah Module OK", key=f"jmo_{i}")
                        jmn = st.text_input("Jumlah Module NOK", key=f"jmn_{i}")
                    with c2:
                        bcl = st.text_input("BCL (%)", key=f"bcl_{i}")
                        lc = st.text_input("Load Current (A)", key=f"lc_{i}")
                        dcv = st.text_input("DC Voltage (V)", key=f"dcv_{i}")
                        lp = st.text_input("LVD Prio", key=f"lp_{i}")
                        lnp = st.text_input("LVD Non Prio", key=f"lnp_{i}")
                        elp = st.text_input("EQP LVD Prio", key=f"elp_{i}")
                        elnp = st.text_input("EQP LVD Non Prio", key=f"elnp_{i}")
                        bbt = st.text_input("Battery Backup Time", key=f"bbt_{i}")
                    
                    pb_recti_list.append({
                        "tipe": tr, "controler": tc, "module": tm, "jml_batt": jb, "batt_ok": bok, "batt_nok": bnok,
                        "mod_ok": jmo, "mod_nok": jmn, "bcl": bcl, "load": lc, "dc_v": dcv, "lvd_p": lp, "lvd_np": lnp, "eqp_p": elp, "eqp_np": elnp, "backup": bbt
                    })
            
            st.divider()
            st.markdown("### 🚨 ENVA Setelah Simulasi (Bukan Cabut Kabel)")
            opt_enva = ["OK", "NOK", "-"]
            c1, c2 = st.columns(2)
            with c1:
                em = st.selectbox("Main Fail", opt_enva, key="em")
                er = st.selectbox("Rectifier Fail", opt_enva, key="er")
                el = st.selectbox("Low Batt", opt_enva, key="el")
            with c2:
                eh = st.selectbox("High Temp", opt_enva, key="eh")
                egr = st.selectbox("Genset Run", opt_enva, key="egr")
                egf = st.selectbox("Genset Fail", opt_enva, key="egf")
                
            pb_foto_recti = st.file_uploader("📸 Foto Rectifier, Baterai & ENVA (Bisa >1)", accept_multiple_files=True, key="pb_foto_recti")
            pb_payload["enva"] = {"main_fail": em, "rect_fail": er, "low_batt": el, "high_temp": eh, "gen_run": egr, "gen_fail": egf}

        with tabs_pb[2]:
            st.markdown("### ⚡ DATA KELISTRIKAN")
            c1, c2 = st.columns(2)
            with c1:
                kd = st.text_input("Daya KWH", key="kd")
                kj = st.text_input("Jumlah Phasa", key="kj")
                kv1 = st.text_input("Tegangan R-S-T thd N", key="kv1")
            with c2:
                kv2 = st.text_input("Tegangan Phasa RS-RT-ST", key="kv2")
                ka = st.text_input("Pengukuran Arus R-S-T", key="ka")
                kgn = st.text_input("Tegangan Pengukuran G-N", key="kgn")
            pb_payload["kelistrikan"] = {"daya": kd, "phasa": kj, "v_rst_n": kv1, "v_rs_rt_st": kv2, "arus_rst": ka, "v_gn": kgn}
            
            st.divider()
            st.markdown("### 🎛 DATA ACPDB")
            c3, c4 = st.columns(2)
            with c3:
                aa = st.selectbox("Ada ACPDB", ["Ya", "Tidak", "-"], key="aa")
                av1 = st.text_input("Tegangan R-S-T thd N (ACPDB)", key="av1")
                av2 = st.text_input("Tegangan Phasa RS-RT-ST (ACPDB)", key="av2")
            with c4:
                aar = st.text_input("Pengukuran Arus R-S-T (ACPDB)", key="aar")
                agn = st.text_input("Tegangan Pengukuran G-N (ACPDB)", key="agn")
            
            pb_foto_listrik = st.file_uploader("📸 Foto Kelistrikan & ACPDB (Bisa >1)", accept_multiple_files=True, key="pb_foto_listrik")
            pb_payload["acpdb"] = {"ada": aa, "v_rst_n": av1, "v_rs_rt_st": av2, "arus_rst": aar, "v_gn": agn}
            
        with tabs_pb[3]:
            st.markdown("### ⛽ DATA GENSET")
            c1, c2 = st.columns(2)
            with c1:
                ga = st.selectbox("Ada Genset", ["Ya", "Tidak"], key="ga")
                gc = st.text_input("Capacity Genset (kVA)", key="gc")
                gs = st.text_input("Genset Status (OK/NOK)", key="gs")
                gco = st.text_input("Config Running", key="gco")
            with c2:
                gt = st.text_input("Kapasitas Tangki", key="gt")
                grh = st.text_input("Running Hour", key="grh")
                gsb = st.text_input("Sisa BBM", key="gsb")
                
            pb_foto_genset = st.file_uploader("📸 Foto Genset & Tangki BBM (Bisa >1)", accept_multiple_files=True, key="pb_foto_genset")
            pb_payload["genset"] = {"ada": ga, "cap": gc, "status": gs, "config": gco, "tangki": gt, "rh": grh, "bbm": gsb}

        with tabs_pb[4]:
            st.markdown("### 📝 ACTIVITY REPORT & CATATAN")
            pb_act1 = st.multiselect("1. Power & CME Activity", ["PMS/ PMG", "Pembersihan/ PM SPS", "Regional Programm", "NOP Programm", "Site Down Troubleshoot", "PLN Undervolt", "Load Balancing", "Rectifier Troubleshoot", "Battery Troubleshoot/ Claim Warranty", "Pergantian Rectifier Controller (SC/ CSU)", "Mapping Rectifier Controller (SC/ CSU)", "Grounding"])
            pb_act2 = st.multiselect("2. Transport Activity", ["Packet Loss", "NMS Unremote", "On Site VLAN Balancing"])
            pb_act3 = st.multiselect("3. OSA Impact Service Activity", ["Validasi ENVA", "BBU Min Voltage", "RRU Min Voltage", "Cell Down", "Disable PA", "Int. Fault", "Cell Flicker", "BBU Temp", "RRU Temp", "Zero Traffic", "Zero Payload", "VSWR", "Optical port link fault", "GNSS/GPS", "BFAN/FCE"])
            pb_temuan = st.text_area("Temuan/ Catatan (Gunakan Enter untuk list ke bawah):", placeholder="1. \n2. \n3.")
            
            pb_foto_act = st.file_uploader("📸 Foto Temuan / Action Lapangan (Bisa >1)", accept_multiple_files=True, key="pb_foto_act")
            pb_status_akhir = st.radio("Status Akhir Site:", ["Normal", "Minor Issue", "Major/Critical"])
            
            if st.button("🚀 SUBMIT SITE VISIT & BUKA WA", type="primary", use_container_width=True):
                if not pb_sid: st.error("⚠️ Mohon pilih Site ID!")
                elif not pb_te: st.error("⚠️ Mohon pilih FME NAME!")
                else:
                    with st.spinner("⏳ Sedang mengunggah foto ke Cloud & menyimpan data Laporan..."):
                        
                        url_pb_site = upload_multiple_images(pb_foto_site, "pb_site")
                        url_pb_recti = upload_multiple_images(pb_foto_recti, "pb_recti")
                        url_pb_listrik = upload_multiple_images(pb_foto_listrik, "pb_listrik")
                        url_pb_genset = upload_multiple_images(pb_foto_genset, "pb_genset")
                        url_pb_act = upload_multiple_images(pb_foto_act, "pb_action")
                        
                        report_dict = {
                            "timestamp": datetime.datetime.now().strftime("%Y-%m-%d %H:%M:%S"),
                            "site_name": pb_sid,
                            "kategori": "Pangkalan Bun", 
                            "nop": "Pangkalan Bun", 
                            "teknisi": tek_pb_str,
                            "status": pb_status_akhir, 
                            "action": pb_temuan,
                            
                            "pb_te": tek_pb_str, "pb_site_id": pb_sid, "pb_ticket": pb_tick, "pb_power": pb_catuan,
                            "pb_rectifiers": pb_recti_list,
                            "pb_enva": pb_payload["enva"],
                            "pb_kelistrikan": pb_payload["kelistrikan"],
                            "pb_acpdb": pb_payload["acpdb"],
                            "pb_genset": pb_payload["genset"],
                            "pb_act_power": pb_act1, "pb_act_trans": pb_act2, "pb_act_osa": pb_act3,
                            
                            "pb_url_site": url_pb_site,
                            "pb_url_recti": url_pb_recti,
                            "pb_url_listrik": url_pb_listrik,
                            "pb_url_genset": url_pb_genset,
                            "pb_url_act": url_pb_act,
                            
                            "panel_data": [], "battery_data": [], "url_sites": [], "total_load": "-", "rect_out_v": "-", "sparepart": "-"
                        }
                        
                        sheet_pbu = connect_gsheets("Report PBU")
                        if sheet_pbu:
                            row_data = [
                                report_dict['timestamp'], report_dict['site_name'], report_dict['nop'], 
                                report_dict['teknisi'], report_dict['status'], json.dumps(report_dict)
                            ]
                            sheet_pbu.insert_row(row_data, index=2)
                            st.cache_data.clear()
                            st.session_state['laporan_db'] = fetch_data_from_gsheets()
                            
                            wa_text = generate_wa_pbu(report_dict)
                            wa_url = f"https://wa.me/?text={urllib.parse.quote(wa_text)}"
                            
                            st.success("✅ BERHASIL! Data Visit & Dokumentasi Foto Pangkalan Bun telah tersimpan di Baris Teratas.")
                            st.markdown("### 📱 Langkah Terakhir: Bagikan Laporan")
                            st.markdown("Silakan tekan tombol di bawah ini untuk langsung membagikan rangkuman laporan Anda ke Grup WhatsApp.")
                            st.link_button("🚀 BUKA WHATSAPP & KIRIM LAPORAN", wa_url, type="primary")
                            
                        else:
                            st.error("⚠️ Gagal menyambung ke Server Database.")


# =========================================================================
# MENU 2: HASIL LAPORAN & DASHBOARD (DENGAN PILIHAN AREA ELEGAN)
# =========================================================================
elif menu == "📊 Hasil Laporan & Dashboard":
    
    st.markdown("<h1 style='text-align: center;'>📊 Executive Dashboard & Analytics Report</h1>", unsafe_allow_html=True)
    st.markdown("<p style='text-align: center; opacity: 0.8; margin-bottom: 25px;'>Sistem Monitoring Terpadu Hasil Preventive Maintenance & Site Visit</p>", unsafe_allow_html=True)
    
    db = st.session_state['laporan_db']

    if not db:
        st.warning("⚠ Belum ada data di Server / Koneksi Sedang Proses.")
    else:
        idx_plk = [idx for idx, r in enumerate(db) if r.get('kategori') not in ['Pangkalan Bun', 'Project']]
        idx_pbu = [idx for idx, r in enumerate(db) if r.get('kategori') == 'Pangkalan Bun']

        # ---------------------------------------------------------------------
        # LAYAR PILIHAN AREA DASHBOARD (AGAR TIDAK NUMPUK)
        # ---------------------------------------------------------------------
        if st.session_state.get('dashboard_area') is None:
            col_m1, col_m2, col_m3 = st.columns(3)
            with col_m1: st.metric("Total Seluruh Laporan", f"{len(db)} Site")
            with col_m2: st.metric("Laporan NOP Palangkaraya", f"{len(idx_plk)} Site")
            with col_m3: st.metric("Laporan NOP Pangkalan Bun", f"{len(idx_pbu)} Site")
            
            st.markdown("<hr style='margin: 22px 0;'>", unsafe_allow_html=True)
            st.markdown("<h3 style='text-align: center; margin-bottom: 25px;'>📍 Pilih Area Operasional untuk Membuka Daftar Site</h3>", unsafe_allow_html=True)

            col_d1, col_d_space, col_d2 = st.columns([4.5, 0.4, 4.5])

            with col_d1:
                st.markdown(f"""
                    <div class='area-card' style='border-top-color: #007BFF;'>
                        <div style='font-size: 50px; margin-bottom: 10px;'>🗼</div>
                        <div class='area-badge'>AREA OPERASIONAL 1</div>
                        <div class='login-title'>NOP PALANGKARAYA</div>
                        <div class='login-desc'>Daftar Site Preventive Maintenance (SPS & Reguler), Laporan Project, Dokumentasi Foto Modul Surya, Baterai, Datalog & Berita Acara PDF.</div>
                        <div style='font-size: 20px; font-weight: 800; color: #007BFF; margin-bottom: 8px;'>📊 {len(idx_plk)} Laporan Tersedia</div>
                    </div>
                """, unsafe_allow_html=True)
                if st.button("🔵 BUKA DASHBOARD PALANGKARAYA", type="secondary", use_container_width=True, key="btn_open_plk"):
                    st.session_state['dashboard_area'] = "Palangkaraya"
                    st.rerun()

            with col_d2:
                st.markdown(f"""
                    <div class='area-card' style='border-top-color: #FF3333;'>
                        <div style='font-size: 50px; margin-bottom: 10px;'>📡</div>
                        <div class='area-badge' style='color: #FF3333; border-color: #FF3333; background-color: rgba(255,51,51,0.12);'>AREA OPERASIONAL 2</div>
                        <div class='login-title'>NOP PANGKALAN BUN</div>
                        <div class='login-desc'>Daftar Site Visit Report FME, Parameter Rectifier, Simulasi ENVA, Kelistrikan, ACPDB, Genset & Dokumentasi Foto Lapangan.</div>
                        <div style='font-size: 20px; font-weight: 800; color: #FF3333; margin-bottom: 8px;'>📊 {len(idx_pbu)} Laporan Tersedia</div>
                    </div>
                """, unsafe_allow_html=True)
                if st.button("🔴 BUKA DASHBOARD PANGKALAN BUN", type="primary", use_container_width=True, key="btn_open_pbu"):
                    st.session_state['dashboard_area'] = "Pangkalan Bun"
                    st.rerun()

            st.stop()

        # ---------------------------------------------------------------------
        # JIKA AREA SUDAH DIPILIH -> TAMPILKAN NAVIGASI KEMBALI & DAFTAR SITE
        # ---------------------------------------------------------------------
        aktif_area = st.session_state.get('dashboard_area')
        
        c_nav1, c_nav2 = st.columns([1.5, 3.5])
        with c_nav1:
            if st.button("⬅️ Ganti Area Operasional Dashboard", use_container_width=True):
                st.session_state['dashboard_area'] = None
                st.rerun()
        with c_nav2:
            warna_area = "#007BFF" if aktif_area == "Palangkaraya" else "#FF3333"
            st.markdown(
                f"<div style='padding: 8px 18px; border-radius: 10px; border-left: 6px solid {warna_area}; background-color: var(--secondary-background-color); font-size: 18px; font-weight: 800;'>📍 Menampilkan Data NOP {aktif_area.upper()}</div>",
                unsafe_allow_html=True
            )

        st.markdown("<hr style='margin: 18px 0;'>", unsafe_allow_html=True)

        # Filter berdasarkan Area yang dipilih + Pencarian Cepat
        if aktif_area == "Palangkaraya":
            c_flt1, c_flt2 = st.columns([2, 2])
            with c_flt1:
                filter_kat = st.radio("Sub-Kategori Site Palangkaraya:", ["Semua (Palangkaraya)", "SPS", "Site Reguler"], horizontal=True)
            with c_flt2:
                cari_site = st.text_input("🔍 Cari Nama Site / Teknisi:", placeholder="Ketik ID atau Nama Site...")
                
            filtered_indices = []
            for idx in idx_plk:
                r = db[idx]
                kat = r.get('kategori', 'SPS')
                match_kat = (filter_kat == "Semua (Palangkaraya)") or (kat == filter_kat)
                match_cari = True
                if cari_site.strip():
                    q = cari_site.strip().lower()
                    match_cari = (q in str(r.get('site_name', '')).lower()) or (q in str(r.get('teknisi', '')).lower())
                if match_kat and match_cari:
                    filtered_indices.append(idx)
        else:
            filter_kat = "Pangkalan Bun"
            cari_site = st.text_input("🔍 Cari Site ID / Nama Site / FME Name (Pangkalan Bun):", placeholder="Ketik Site ID atau Nama FME...")
            filtered_indices = []
            for idx in idx_pbu:
                r = db[idx]
                match_cari = True
                if cari_site.strip():
                    q = cari_site.strip().lower()
                    match_cari = (q in str(r.get('site_name', '')).lower()) or (q in str(r.get('pb_te', '')).lower()) or (q in str(r.get('pb_ticket', '')).lower())
                if match_cari:
                    filtered_indices.append(idx)

        # Metrik Ringkasan Area Aktif
        export_db = [db[idx] for idx in filtered_indices]
        jml_normal = sum(1 for x in export_db if x.get('status') in ['Normal', 'Done'])
        jml_issue = len(export_db) - jml_normal

        cm1, cm2, cm3 = st.columns(3)
        cm1.metric(f"Total Laporan ({aktif_area})", f"{len(export_db)} Berkas")
        cm2.metric("🟢 Status Normal / Done", f"{jml_normal} Berkas")
        cm3.metric("🟡/🔴 Issue / Progress", f"{jml_issue} Berkas")

        if len(filtered_indices) == 0:
            st.info(f"Tidak ada data laporan yang cocok untuk filter saat ini di area {aktif_area}.")
            st.stop()

        st.markdown("<br>", unsafe_allow_html=True)

        # ---------------------------------------------------------------------
        # TABEL EXCEL FLAT & POWERPOINT (SESUAI AREA AKTIF)
        # ---------------------------------------------------------------------
        with st.container(border=True):
            st.markdown(f"#### 📥 Ekspor Rekapitulasi Laporan — {aktif_area}")
            summary_list = []
            
            for r in export_db:
                if r.get('kategori') == 'Pangkalan Bun':
                    kel = r.get('pb_kelistrikan', {})
                    acp = r.get('pb_acpdb', {})
                    enva = r.get('pb_enva', {})
                    gen = r.get('pb_genset', {})

                    row_data = {
                        "Timestamp": r.get('timestamp', '-'),
                        "Nama Site": r.get('site_name', '-'),
                        "Kategori": "Pangkalan Bun",
                        "FME NAME": r.get('pb_te', '-'),
                        "Ticket#": r.get('pb_ticket', '-'),
                        "Catuan Power": r.get('pb_power', '-'),
                        "Status Akhir": r.get('status', '-'),
                        
                        "Main Fail (ENVA)": enva.get('main_fail', '-'),
                        "Rect Fail (ENVA)": enva.get('rect_fail', '-'),
                        "Low Batt (ENVA)": enva.get('low_batt', '-'),
                        "High Temp (ENVA)": enva.get('high_temp', '-'),
                        "Genset Run (ENVA)": enva.get('gen_run', '-'),
                        "Genset Fail (ENVA)": enva.get('gen_fail', '-'),
                        
                        "KWH Daya": kel.get('daya', '-'),
                        "KWH Phasa": kel.get('phasa', '-'),
                        "Tegangan RST-N": kel.get('v_rst_n', '-'),
                        "Tegangan RS-RT-ST": kel.get('v_rs_rt_st', '-'),
                        "Arus RST": kel.get('arus_rst', '-'),
                        "Tegangan G-N": kel.get('v_gn', '-'),
                        
                        "ACPDB Ada": acp.get('ada', '-'),
                        "ACPDB RST-N": acp.get('v_rst_n', '-'),
                        "ACPDB RS-RT-ST": acp.get('v_rs_rt_st', '-'),
                        "ACPDB Arus": acp.get('arus_rst', '-'),
                        "ACPDB G-N": acp.get('v_gn', '-'),
                        
                        "Genset Ada": gen.get('ada', '-'),
                        "Genset Cap": gen.get('cap', '-'),
                        "Genset Status": gen.get('status', '-'),
                        "Genset Config": gen.get('config', '-'),
                        "Genset Tangki": gen.get('tangki', '-'),
                        "Genset RH": gen.get('rh', '-'),
                        "Genset BBM": gen.get('bbm', '-'),
                        
                        "Power & CME Act": ", ".join(r.get('pb_act_power', [])) if isinstance(r.get('pb_act_power'), list) else str(r.get('pb_act_power','-')),
                        "Transport Act": ", ".join(r.get('pb_act_trans', [])) if isinstance(r.get('pb_act_trans'), list) else str(r.get('pb_act_trans','-')),
                        "OSA Act": ", ".join(r.get('pb_act_osa', [])) if isinstance(r.get('pb_act_osa'), list) else str(r.get('pb_act_osa','-')),
                        "Catatan/Temuan": r.get('action', '-')
                    }
                    
                    rects = r.get('pb_rectifiers', [])
                    for i in range(4):
                        prefix = f"Rect{i+1} "
                        if i < len(rects):
                            rect = rects[i]
                            row_data[prefix + "Tipe"] = rect.get('tipe', '-')
                            row_data[prefix + "Module"] = rect.get('module', '-')
                            row_data[prefix + "Load(A)"] = rect.get('load', '-')
                            row_data[prefix + "DC(V)"] = rect.get('dc_v', '-')
                            row_data[prefix + "Batt(OK/NOK)"] = f"{rect.get('batt_ok', '-')}/{rect.get('batt_nok', '-')}"
                        else:
                            row_data[prefix + "Tipe"] = "-"
                            row_data[prefix + "Module"] = "-"
                            row_data[prefix + "Load(A)"] = "-"
                            row_data[prefix + "DC(V)"] = "-"
                            row_data[prefix + "Batt(OK/NOK)"] = "-"

                    summary_list.append(row_data)

                else:
                    panel_issues = [p.get('Panel', 'Panel') for p in r.get('panel_data', []) if p.get('kondisi') and p.get('kondisi') != "Baik"]
                    bat_issues = [b.get('Baterai', 'Baterai') for b in r.get('battery_data', []) if b.get('Kondisi') and b.get('Kondisi') != "Normal"]

                    cond_d = r.get('cond_data', {})
                    pd_d = r.get('proj_data', {})
                    summary_list.append({
                        "Timestamp": r.get('timestamp', '-'),
                        "Nama Site": r.get('site_name', '-'),
                        "Kategori": r.get('kategori', 'SPS'),
                        "NOP": r.get('nop', '-'),
                        "Teknisi / FME": r.get('teknisi', '-'),
                        "Role FME (Reguler)": r.get('role_fme', '-'),
                        "Status Akhir": r.get('status', '-'),
                        "Kondisi Site": r.get('site_cond', '-'),
                        "Fisik Tower": r.get('tower_cond', '-'),
                        "Shading Panel": r.get('shading_status', '-'),
                        "Status Panel": r.get('status_panel', '-'),
                        "Jumlah Panel Rusak": r.get('jml_panel_rusak', 0),
                        "Type Merk Panel": r.get('merk_panel', '-'),
                        "Status Battery": r.get('status_battery', '-'),
                        "Jumlah battery Rusak": r.get('jml_bat_rusak', 0),
                        "Type Battery": r.get('tipe_battery', '-'),
                        "status SCC": r.get('status_scc', '-'),
                        "Jumlah SCC NOK": r.get('jml_scc_nok', 0),
                        "Type/Merk SCC": r.get('merk_scc', '-'),
                        "Possibility PSB": r.get('possibility_psb', '-'),
                        "Estimasi Jarak tower ke jaringan PLN": r.get('jarak_pln', '-'),
                        "SPS Rusak (Detail)": ", ".join(panel_issues) if panel_issues else "Aman (Baik)",
                        "Load Beban (A)": r.get('total_load', '-'),
                        "PLN": r.get('pln_status', '-'),
                        "Rectifier (Merek)": r.get('rect_brand', '-'),
                        "Voltase Recti (V)": r.get('rect_out_v', '-'),
                        "Genset": r.get('genset_status', '-'),
                        "Level BBM (%)": r.get('fuel_pct', '-'),
                        "Baterai Rusak (Detail)": ", ".join(bat_issues) if bat_issues else "Aman (Normal)",
                        "Grounding (Ohm)": r.get('earth_ohm', '-'),
                        "Action / Pekerjaan": r.get('action', '-'),
                        "Sparepart Diganti": r.get('sparepart', '-'),
                        "Remark": r.get('remark', '-'),
                        
                        "Reguler - Problem Alarm": r.get('problem_alarm', '-'),
                        "Reguler - Teg AC (RN/SN/TN)": f"{cond_d.get('v_rn','-')} / {cond_d.get('v_sn','-')} / {cond_d.get('v_tn','-')}",
                        "Reguler - Arus AC (R/S/T)": f"{cond_d.get('i_r','-')} / {cond_d.get('i_s','-')} / {cond_d.get('i_t','-')}",
                        "Reguler - Load DC": cond_d.get('load_dc', '-'),
                        "Reguler - Set BCL": cond_d.get('set_bcl', '-'),
                        "Reguler - Set Capacity": cond_d.get('set_cap', '-'),
                        "Reguler - Module Rusak": cond_d.get('mod_rusak', '-'),
                        "Reguler - Note/Remark": cond_d.get('note_action', '-'),
                        
                        "Project - Jenis": pd_d.get('jenis_project', '-'),
                        "Project - No PO/Tiket": str(pd_d.get('no_po', '')) + str(pd_d.get('no_tiket', '')),
                        "Project - Mitra": pd_d.get('mitra', '-'),
                        "Project - Pekerjaan": pd_d.get('tipe_pekerjaan', '-'),
                        "Project - Temuan SVA": pd_d.get('temuan_sva', '-'),
                        "Project - Catatan Akhir": pd_d.get('note_akhir', '-')
                    })
                
            df_export = pd.DataFrame(summary_list)
            
            col_dl_ex, col_dl_ppt = st.columns(2)
            with col_dl_ex:
                try:
                    excel_buffer = BytesIO()
                    with pd.ExcelWriter(excel_buffer, engine='xlsxwriter') as writer:
                        df_export.to_excel(writer, index=False, sheet_name='Database PM')
                        workbook = writer.book
                        worksheet = writer.sheets['Database PM']
                        
                        header_format = workbook.add_format({'bold': True, 'font_color': 'white', 'bg_color': '#112240', 'border': 1, 'align': 'center', 'valign': 'vcenter'})
                        for col_num, value in enumerate(df_export.columns.values):
                            worksheet.write(0, col_num, value, header_format)

                        cell_center = workbook.add_format({'align': 'center', 'valign': 'vcenter', 'border': 1})
                        cell_left = workbook.add_format({'align': 'left', 'valign': 'vcenter', 'border': 1, 'text_wrap': True})

                        for i, col_name in enumerate(df_export.columns):
                            if col_name in ["Action / Pekerjaan", "Remark", "Possibility PSB", "Estimasi Jarak tower ke jaringan PLN", "Type Merk Panel", "Type Battery", "SPS Rusak (Detail)", "Baterai Rusak (Detail)", "Sparepart Diganti", "Catatan/Temuan", "Reguler - Problem Alarm", "Reguler - Note/Remark", "Project - Temuan SVA", "Project - Catatan Akhir"]:
                                worksheet.set_column(i, i, 35, cell_left)
                            else:
                                worksheet.set_column(i, i, 18, cell_center)

                        worksheet.freeze_panes(1, 0)
                        worksheet.autofilter(0, 0, len(df_export), len(df_export.columns) - 1)
                        
                    file_data = excel_buffer.getvalue()
                    st.download_button(label=f"📊 Download Excel ({aktif_area})", data=file_data, file_name=f"Database_PM_{aktif_area}_{datetime.date.today()}.xlsx", mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet", type="primary", use_container_width=True)
                except Exception:
                    csv_data = df_export.to_csv(index=False, sep=",").encode('utf-8')
                    st.download_button(label="📥 Download Laporan (CSV)", data=csv_data, file_name=f"Database_PM_{datetime.date.today()}.csv", mime="text/csv", type="secondary")

            with col_dl_ppt:
                if aktif_area == "Palangkaraya":
                    if HAS_PPTX:
                        with st.spinner("Menyiapkan PPTX..."):
                            pptx_data = build_pptx([x for x in export_db if x.get('kategori') != 'Pangkalan Bun' and x.get('kategori') != 'Site Reguler' and x.get('kategori') != 'Project'])
                            st.download_button(label="📽 Download Presentasi Report (PPTX)", data=pptx_data, file_name=f"Report_PM_KUT_{datetime.date.today()}.pptx", mime="application/vnd.openxmlformats-officedocument.presentationml.presentation", type="primary", use_container_width=True)
                    else:
                        st.error("Library `python-pptx` belum ter-install di server.")
                else:
                    st.info("💡 Rekap lengkap parameter FME Pangkalan Bun dapat diunduh melalui tombol Excel di samping.")

        st.markdown("### 📋 Daftar Laporan Site Terinspeksi")

        # ---------------------------------------------------------------------
        # TAMPILAN KARTU UI MODERN & URUTAN TERBARU DI ATAS
        # ---------------------------------------------------------------------
        for i in filtered_indices:
            r = db[i]
            site_id = r.get('site_name', 'Unknown')
            kategori_site = r.get('kategori', 'SPS')
            stat = r.get('status', '')
            
            # Pewarnaan Label Status UI
            if stat == "Normal" or stat == "Done":
                warna_status = "#00E676" # Hijau
                icon = "🟢"
            elif stat in ["Minor Issue", "Update Progress", "On Progress"]:
                warna_status = "#FFC107" # Kuning
                icon = "🟡"
            else:
                warna_status = "#FF5252" # Merah
                icon = "🔴"
            
            # CEK KATEGORI DATA
            is_new_reguler = kategori_site == "Site Reguler" and ("problem_alarm" in r or "cond_data" in r)
            is_project = kategori_site == "Project"
            
            # FORMAT WHATSAPP DINAMIS
            if kategori_site == "Pangkalan Bun":
                wa_text = generate_wa_pbu(r)
            elif is_project:
                pd_data = r.get('proj_data', {})
                wa_text = f"*REPORT PROJECT ({pd_data.get('jenis_project', '')})* ⚡\n📍 *Project:* {site_id} ({r.get('nop', '-')})\n📅 *Tanggal:* {r.get('timestamp', '-')}\n👷 *Tim/FME:* {r.get('teknisi', '-')}\n📊 *Status Laporan:* {r.get('status', '-')}\n"
                
                if pd_data.get('jenis_project') == "PO":
                    wa_text += f"\n*📋 DETAIL PO:*\n- No PO: {pd_data.get('no_po', '-')}\n- Mitra: {pd_data.get('mitra', '-')}\n- Pekerjaan Utama: {pd_data.get('tipe_pekerjaan', '-')}\n- Progress Keseluruhan: {pd_data.get('avg_progress', 0):.1f}%\n"
                    
                    tabel_p = pd_data.get('tabel_pekerjaan', [])
                    if tabel_p:
                        wa_text += f"\n*🗺️ RINCIAN MULTI-SITE ({len(tabel_p)} Rute/Item):*\n"
                        for idx, tp in enumerate(tabel_p):
                            asal = tp.get('Site Asal', '')
                            tuj = tp.get('Site Tujuan', '')
                            pek = tp.get('Pekerjaan Spesifik', '')
                            prog = tp.get('Progress (%)', 0)
                            stat_tp = tp.get('Status', '')
                            rute = f"{asal} ➡️ {tuj}" if asal and tuj else (asal or tuj or "Site Project")
                            wa_text += f"  {idx+1}. {rute} | {pek} ({prog}% - {stat_tp})\n"
                            
                    if pd_data.get('tipe_pekerjaan') == "Instalasi Imacro":
                        wa_text += f"\n*🛠️ Progress Imacro:* {', '.join(pd_data.get('progress_imacro', []))}\n"
                elif pd_data.get('jenis_project') == "SVA":
                    wa_text += f"\n*📋 DETAIL SVA:*\n- No Tiket: {pd_data.get('no_tiket', '-')}\n- Temuan: {pd_data.get('temuan_sva', '-')}\n"
                
                if pd_data.get('note_akhir'):
                    wa_text += f"\n*📝 NOTE:* {pd_data.get('note_akhir')}\n"
            elif is_new_reguler:
                wa_text = f"*BERITA ACARA SITE REGULER* ⚡\n📍 *Site:* {site_id} ({r.get('nop', '-')})\n📅 *Tanggal:* {r.get('timestamp', '-')}\n👷 *FME:* {r.get('teknisi', '-')} ({r.get('role_fme', '-')})\n📊 *Status:* {r.get('status', '-')}\n\n*🚨 PROBLEM ALARM:*\n{r.get('problem_alarm', '-')}\n\n*🔧 ACTION / PERBAIKAN:*\n{', '.join(r.get('action_reguler', [])) if isinstance(r.get('action_reguler'), list) else r.get('action_reguler', '-')}\n"
                cd = r.get('cond_data', {})
                if cd:
                    if cd.get('v_rn'):
                        wa_text += f"\n*⚡ PARAMETER AC/DC:*\n- Teg AC (R/S/T): {cd.get('v_rn','-')}/{cd.get('v_sn','-')}/{cd.get('v_tn','-')} V\n- Arus AC (R/S/T): {cd.get('i_r','-')}/{cd.get('i_s','-')}/{cd.get('i_t','-')} A\n- Load DC: {cd.get('load_dc', '-')} A\n- Set BCL: {cd.get('set_bcl', '-')} | Cap: {cd.get('set_cap', '-')}\n- Modul Rusak: {cd.get('mod_rusak', '-')}\n"
                    if cd.get('gen_v'):
                        wa_text += f"\n*⚙️ PARAMETER GENSET:*\n- Teg Output: {cd.get('gen_v','-')} V | Load: {cd.get('gen_i','-')} A\n- Status ATS: {cd.get('gen_ats','-')}\n- Issue Mesin: {cd.get('gen_issue','-')}\n- Hour Meter: {cd.get('gen_hm','-')} HM\n- Sisa BBM: {cd.get('gen_bbm','-')} Liter\n"
                    
                    if cd.get('note_action'):
                        wa_text += f"\n*📝 NOTE/REMARK:*\n{cd.get('note_action')}\n"
            else:
                wa_text = f"*BERITA ACARA PREVENTIVE MAINTENANCE* ⚡\n📍 *Site:* {site_id} ({kategori_site} - {r.get('nop', '-')})\n📅 *Tanggal:* {r.get('timestamp', '-')}\n👷 *Pelaksana:* {r.get('teknisi', '-')}\n📊 *Status:* {r.get('status', '-')}\n\n*RINCIAN TINDAKAN:*\n{r.get('action', '-')}\n\n*POWER & LOAD:*\n- PLN: {r.get('pln_status', '-')}\n- Rectifier: {r.get('rect_brand', '-')} ({r.get('rect_out_v', '-')}V)\n- Load BTS: {r.get('total_load', '-')} A\n\n*SPAREPART:*\n{r.get('sparepart', '-')}"
            
            wa_url = f"https://wa.me/?text={urllib.parse.quote(wa_text)}"

            # BUNGKUSAN CARD UTAMA
            with st.container(border=False):
                fme_tampil = r.get('teknisi', '-') if kategori_site != "Pangkalan Bun" else r.get('pb_te', '-')
                
                # Desain HTML Card Premium
                st.markdown(f"""
                    <div style="padding: 18px 20px; border-radius: 12px; border-left: 8px solid {warna_status}; background-color: var(--secondary-background-color); margin-bottom: 12px; box-shadow: 0 6px 16px rgba(0,0,0,0.08); border-top: 1px solid rgba(128,128,128,0.1); border-right: 1px solid rgba(128,128,128,0.1); border-bottom: 1px solid rgba(128,128,128,0.1);">
                        <div style="display: flex; justify-content: space-between; align-items: center; margin-bottom: 12px;">
                            <h3 style="margin: 0; padding: 0; font-size: 22px; font-weight: 800;">{icon} {site_id}</h3>
                            <span style="font-size: 13px; font-weight: 700; color: #ffffff; background-color: {warna_status}; padding: 4px 12px; border-radius: 20px; box-shadow: 0 2px 4px rgba(0,0,0,0.2);">📅 {r.get('timestamp', '')}</span>
                        </div>
                        <div style="font-size: 15px; margin-bottom: 8px; display: flex; align-items: center; gap: 8px; flex-wrap: wrap;">
                            <span style="background-color: rgba(128,128,128,0.1); padding: 4px 10px; border-radius: 6px;"><b>📝 Kategori:</b> {kategori_site} ({r.get('nop', '-')})</span>
                            <span style="background-color: rgba(128,128,128,0.1); padding: 4px 10px; border-radius: 6px;"><b>👷 Tim:</b> {fme_tampil}</span>
                        </div>
                        <div style="font-size: 15px; border-top: 1px dashed rgba(128,128,128,0.3); padding-top: 8px; margin-top: 8px;">
                            <b>📊 Status Akhir:</b> <span style="color: {warna_status}; font-weight: 900; font-size: 16px;">{stat}</span>
                        </div>
                    </div>
                """, unsafe_allow_html=True)
                
                # Tombol Aksi Cepat di luar Expander
                c_btn1, c_btn2 = st.columns(2)
                with c_btn1:
                    st.link_button("📱 Share Rangkuman Laporan (WA)", wa_url, use_container_width=True)
                with c_btn2:
                    if kategori_site != "Pangkalan Bun" and not is_new_reguler and not is_project:
                        if st.checkbox("📄 Buat PDF Resmi", key=f"prep_pdf_{i}"):
                            with st.spinner("⏳ Rendering Dokumen PDF Resolusi Tinggi..."):
                                try:
                                    pdf_bytes = build_pdf(r)
                                    st.download_button(label="📥 Download PDF", data=pdf_bytes, file_name=f"Berita_Acara_{site_id}.pdf", mime="application/pdf", key=f"dl_pdf_{i}", type="primary", use_container_width=True)
                                except Exception as e:
                                    st.error(f"Gagal memproses PDF: {e}")
                    else:
                        st.button("📄 Format Tanpa PDF", disabled=True, key=f"dl_pdf_dis_{i}", use_container_width=True)

                # Rincian Data
                with st.expander("📂 BUKA DETAIL & DOKUMENTASI LAPORAN"):
                    
                    if kategori_site == "Pangkalan Bun":
                        st.markdown("### 📋 SITE VISIT REPORT (Pangkalan Bun)")
                        st.write(f"**FME NAME:** {r.get('pb_te', '-')} | **Ticket:** {r.get('pb_ticket', '-')}")
                        st.write(f"**Catuan Power:** {r.get('pb_power', '-')} | **Status Akhir:** {stat}")
                        
                        tpb1, tpb2, tpb3, tpb4 = st.tabs(["📌 Info & Recti", "⚡ Kelistrikan & ACPDB", "⛽ Genset & ENVA", "📝 Temuan & Activity"])
                        with tpb1:
                            st.markdown("**Data Rectifier:**")
                            for rect in r.get('pb_rectifiers', []): st.write(f"🔹 **{rect.get('tipe')}** | BCL: {rect.get('bcl')}% | Beban: {rect.get('load')}A | Voltase: {rect.get('dc_v')}V")
                            st.markdown("<br>", unsafe_allow_html=True)
                            tampilkan_grid_foto(r.get('pb_url_site', []), "📸 Foto Site / Kondisi Umum")
                            tampilkan_grid_foto(r.get('pb_url_recti', []), "📸 Foto Rectifier & ENVA")
                        with tpb2:
                            kel = r.get('pb_kelistrikan', {})
                            acp = r.get('pb_acpdb', {})
                            st.write(f"**Kelistrikan:** Daya KWH: {kel.get('daya', '-')} | Phasa: {kel.get('phasa', '-')} | Teg. G-N: {kel.get('v_gn', '-')}")
                            st.write(f"**ACPDB:** Tersedia: {acp.get('ada', '-')} | Teg. RST thd N: {acp.get('v_rst_n', '-')}")
                            st.markdown("<br>", unsafe_allow_html=True)
                            tampilkan_grid_foto(r.get('pb_url_listrik', []), "📸 Foto Kelistrikan & ACPDB")
                        with tpb3:
                            gen = r.get('pb_genset', {})
                            en = r.get('pb_enva', {})
                            st.write(f"**Genset:** Ada: {gen.get('ada', '-')} | Kapasitas: {gen.get('cap', '-')} | Tangki: {gen.get('tangki', '-')} | Sisa: {gen.get('bbm', '-')}")
                            st.write(f"**ENVA:** Main Fail: {en.get('main_fail', '-')} | Rectifier Fail: {en.get('rect_fail', '-')} | High Temp: {en.get('high_temp', '-')}")
                            st.markdown("<br>", unsafe_allow_html=True)
                            tampilkan_grid_foto(r.get('pb_url_genset', []), "📸 Foto Genset / Tangki")
                        with tpb4:
                            st.markdown("**Catatan/Temuan Tim:**")
                            st.info(r.get('action', '-'))
                            st.markdown("**Activity Report:**")
                            if r.get('pb_act_power'): st.write(f"- Power & CME: {', '.join(r.get('pb_act_power'))}")
                            if r.get('pb_act_trans'): st.write(f"- Transport: {', '.join(r.get('pb_act_trans'))}")
                            if r.get('pb_act_osa'): st.write(f"- OSA Impact: {', '.join(r.get('pb_act_osa'))}")
                            st.markdown("<br>", unsafe_allow_html=True)
                            tampilkan_grid_foto(r.get('pb_url_act', []), "📸 Foto Temuan / Action Lapangan")

                    elif is_project:
                        st.markdown("### 📋 REPORT PROJECT")
                        pd_data = r.get('proj_data', {})
                        st.markdown(f"**👨‍🔧 Tim:** {r.get('teknisi', '-')} &nbsp;&nbsp;|&nbsp;&nbsp; **🏷️ Kategori:** Project {pd_data.get('jenis_project', '-')}")
                        
                        if pd_data.get('jenis_project') == 'PO':
                            st.info(f"**📋 Detail PO:** No PO: {pd_data.get('no_po','-')} | Mitra: {pd_data.get('mitra','-')} | Pekerjaan Utama: {pd_data.get('tipe_pekerjaan','-')}")
                            
                            val_po = float(str(pd_data.get('nilai_po', '0')).replace('.', '').replace(',', '')) if pd_data.get('nilai_po') else 0.0
                            t_cost = pd_data.get('total_cost', 0.0)
                            t_margin = pd_data.get('margin', 0.0)
                            avg_prog = pd_data.get('avg_progress', 0.0)
                            
                            c_mar1, c_mar2, c_mar3, c_mar4 = st.columns(4)
                            c_mar1.metric("💰 Nilai PO Total", f"Rp {val_po:,.0f}")
                            c_mar2.metric("💸 Total Cost Pelaksanaan", f"Rp {t_cost:,.0f}")
                            c_mar3.metric("📈 Profit Margin", f"Rp {t_margin:,.0f}")
                            c_mar4.metric("📊 Rata-Rata Progress", f"{avg_prog:.1f}%")
                            
                            tabel_p = pd_data.get('tabel_pekerjaan', [])
                            if tabel_p:
                                st.markdown("#### 🗺️ Rincian Timeline, Progress & Cost Per-Site")
                                df_show = pd.DataFrame(tabel_p)
                                # Gunakan column_config agar progress bar muncul cantik!
                                st.dataframe(
                                    df_show, 
                                    use_container_width=True,
                                    hide_index=True,
                                    column_config={
                                        "Budget Per Site": st.column_config.NumberColumn("Budget / Cost", format="Rp %d"),
                                        "Progress (%)": st.column_config.ProgressColumn("Progress Aktual", min_value=0, max_value=100, format="%d%%"),
                                        "Start Date": st.column_config.DateColumn("Tgl Mulai"),
                                        "End Date": st.column_config.DateColumn("Target Selesai"),
                                    }
                                )

                            if pd_data.get('tipe_pekerjaan') == 'Instalasi Imacro':
                                st.warning(f"**🛠️ Progress Checklists:** {', '.join(pd_data.get('progress_imacro', []))}")
                                
                        elif pd_data.get('jenis_project') == 'SVA':
                            st.info(f"**📋 Detail SVA:** No Tiket: {pd_data.get('no_tiket','-')} | Tanggal: {pd_data.get('tgl_sva','-')} | Temuan: {pd_data.get('temuan_sva','-')}")
                            
                        st.markdown("<br>", unsafe_allow_html=True)
                        if pd_data.get('jenis_project') == 'PO':
                            t1, t2, t3 = st.tabs(["📸 Foto Before", "📸 Foto On Progress", "📸 Foto After / BAST"])
                            with t1: tampilkan_grid_foto(pd_data.get('url_awal', []), "Kondisi Awal")
                            with t2: tampilkan_grid_foto(pd_data.get('url_proses', []), "Proses Pengerjaan")
                            with t3: tampilkan_grid_foto(pd_data.get('url_akhir', []), "Hasil Akhir (BAST)")
                        else:
                            t1, t2 = st.tabs(["📸 View Site & Perangkat", "📸 Temuan / Defect"])
                            with t1: tampilkan_grid_foto(pd_data.get('url_sva1', []), "Foto View Site & Perangkat")
                            with t2: tampilkan_grid_foto(pd_data.get('url_sva2', []), "Foto Temuan Defect")
                    elif is_new_reguler:
                        st.markdown("### 📋 SITE REGULER REPORT")
                        st.markdown(f"**👨‍🔧 FME:** {r.get('teknisi', '-')} ({r.get('role_fme', '-')}) &nbsp;&nbsp;|&nbsp;&nbsp; **🚨 Problem:** {r.get('problem_alarm', '-')}")
                        st.markdown(f"**🔧 Action Dikerjakan:** {', '.join(r.get('action_reguler', [])) if isinstance(r.get('action_reguler'), list) else r.get('action_reguler', '-')}")
                        
                        cd = r.get('cond_data', {})
                        if cd:
                            teks_info = ""
                            if cd.get('v_rn'):
                                v_rst = cd.get('volt_ac_rst') or f"{cd.get('v_rn','-')}/{cd.get('v_sn','-')}/{cd.get('v_tn','-')}"
                                i_rst = cd.get('load_ac_rst') or f"{cd.get('i_r','-')}/{cd.get('i_s','-')}/{cd.get('i_t','-')}"
                                teks_info += f"**⚡ Power AC/DC:** Teg AC ({v_rst} V) | Arus AC ({i_rst} A) | Load DC ({cd.get('load_dc','-')} A) | BCL ({cd.get('set_bcl','-')}) | Cap ({cd.get('set_cap','-')})\n\n"
                            if cd.get('gen_v'):
                                teks_info += f"**⚙️ Genset/ATS:** Teg Out ({cd.get('gen_v','-')} V) | Arus ({cd.get('gen_i','-')} A) | ATS ({cd.get('gen_ats','-')}) | Issue ({cd.get('gen_issue','-')}) | HM ({cd.get('gen_hm','-')}) | BBM ({cd.get('gen_bbm','-')} L)\n\n"
                                
                            if teks_info: st.info(teks_info)
                            if cd.get('note_action'):
                                st.warning(f"**📝 NOTE/REMARK Lapangan:** {cd.get('note_action')}")
                            
                        st.markdown("<br>", unsafe_allow_html=True)
                        t1, t2, t3 = st.tabs(["📸 Problem Alarm", "📸 Action Perbaikan", "📸 Eviden Parameter"])
                        with t1: tampilkan_grid_foto(r.get('url_problem', []), "Dokumentasi Problem Alarm")
                        with t2: tampilkan_grid_foto(r.get('url_action', []), "Dokumentasi Perbaikan / Action")
                        with t3: 
                            if cd:
                                eviden_list = [
                                    (cd.get('url_v_rn'), f"Teg AC R-N: {cd.get('v_rn', '-')} V"),
                                    (cd.get('url_v_sn'), f"Teg AC S-N: {cd.get('v_sn', '-')} V"),
                                    (cd.get('url_v_tn'), f"Teg AC T-N: {cd.get('v_tn', '-')} V"),
                                    (cd.get('url_i_r'), f"Load AC R: {cd.get('i_r', '-')} A"),
                                    (cd.get('url_i_s'), f"Load AC S: {cd.get('i_s', '-')} A"),
                                    (cd.get('url_i_t'), f"Load AC T: {cd.get('i_t', '-')} A"),
                                    (cd.get('url_ldc'), f"Load DC Total: {cd.get('load_dc', '-')} A"),
                                    (cd.get('url_bcl'), f"Set BCL: {cd.get('set_bcl', '-')}"),
                                    (cd.get('url_cap'), f"Set Capacity: {cd.get('set_cap', '-')}"),
                                    (cd.get('url_mod'), f"Module Kontroler: {cd.get('mod_rusak', '-')}"),
                                    (cd.get('url_gen_v'), f"Tegangan Out Genset: {cd.get('gen_v', '-')} V"),
                                    (cd.get('url_gen_i'), f"Load Arus Genset: {cd.get('gen_i', '-')} A"),
                                    (cd.get('url_gen_ats'), f"Panel ATS: {cd.get('gen_ats', '-')}"),
                                    (cd.get('url_gen_issue'), f"Kondisi Mesin: {cd.get('gen_issue', '-')}"),
                                    (cd.get('url_gen_hm'), f"Hour Meter: {cd.get('gen_hm', '-')} HM"),
                                    (cd.get('url_gen_bbm'), f"Sisa BBM: {cd.get('gen_bbm', '-')} Liter")
                                ]
                                ada_foto = False
                                for u_ev, cap_ev in eviden_list:
                                    if u_ev:
                                        ada_foto = True
                                        tampilkan_grid_foto([u_ev], f"📌 {cap_ev}")
                                if not ada_foto:
                                    st.write("Belum ada foto parameter pengukuran pada laporan ini.")
                            else:
                                st.write("Tidak ada parameter pengukuran yang dilaporkan.")

                    else:
                        st.warning("⏱️ Data di bawah ini direkam menggunakan Form Versi Lama (sebelum ada pemisahan layout).")
                        st.markdown(f"**👨‍🔧 Pelaksana (Teknisi):** {r.get('teknisi', '-')} &nbsp;&nbsp;|&nbsp;&nbsp; **⚡ Beban Load:** {r.get('total_load', '-')} A")
                        st.markdown(f"**🔧 Action Dikerjakan:** {r.get('action', '-')}")
                        if r.get('sparepart'): st.warning(f"**📦 Sparepart Diganti:** {r['sparepart']}")
                        if r.get('remark'): st.info(f"**📝 Remark:** {r['remark']}")
                        
                        st.markdown("<br>", unsafe_allow_html=True)
                        
                        ltab1, ltab2, ltab3, ltab4, ltab5, ltab6 = st.tabs(["🏗️ 1. Fisik", "☀️ 2. Panel SPS", "🔌 3. Recti & Genset", "🔋 4. Baterai & Gnd", "📂 5. Datalog", "📈 6. Analisa Power BBU"])
                        
                        with ltab1:
                            st.write(f"- Kondisi Site: {r.get('site_cond', '-')} | Tower: {r.get('tower_cond', '-')} | Shading: {r.get('shading_status', '-')}")
                            tampilkan_grid_foto(r.get('url_sites'), "📸 View Site")
                            tampilkan_grid_foto(r.get('url_shadings'), "📸 Shading")
                            tampilkan_grid_foto(r.get('extras_fisik'), "📸 Tambahan")

                        with ltab2:
                            st.write(f"**Status Panel:** {r.get('status_panel', '-')} ({r.get('jml_panel_rusak', 0)} Rusak) | **Merk Panel:** {r.get('merk_panel', '-')}")
                            st.write(f"**Status SCC:** {r.get('status_scc', '-')} ({r.get('jml_scc_nok', 0)} Rusak) | **Merk SCC:** {r.get('merk_scc', '-')}")
                            for p in r.get('panel_data', []):
                                st.markdown(f"**{p.get('Panel', '-')}** | Voc: {p.get('Voc','-')}V | Isc: {p.get('Isc','-')}A")
                                cb, ca = st.columns(2)
                                with cb: tampilkan_grid_foto(p.get('URL_Before') or p.get('URLs_Before'), "📸 KONDISI PANEL")
                                with ca: tampilkan_grid_foto(p.get('URL_After') or p.get('URLs_After'), "✨ PENGUKURAN")
                                st.divider()
                            tampilkan_grid_foto(r.get('extras_panel'), "📸 Tambahan Panel")

                        with ltab3:
                            c_p1, c_p2 = st.columns(2)
                            with c_p1:
                                st.write(f"**PLN:** {r.get('pln_status', '-')} | **Rectifier:** {r.get('rect_brand', '-')} ({r.get('rect_out_v', '-')}V)")
                                st.write(f"**Possibility PSB:** {r.get('possibility_psb', '-')} | **Jarak PLN:** {r.get('jarak_pln', '-')}")
                                tampilkan_grid_foto(r.get('url_rects'), "📸 Rectifier")
                            with c_p2:
                                st.write(f"**Genset:** {r.get('genset_status', '-')} | **BBM:** {r.get('fuel_pct', '-')}%")
                                tampilkan_grid_foto(r.get('url_gensets'), "📸 Genset")
                            st.divider()
                            tampilkan_grid_foto(r.get('extras_elektrikal'), "📸 Tambahan Mesin/Elektrikal")

                        with ltab4:
                            st.markdown("### 🔋 Bank Baterai")
                            st.write(f"**Status Baterai Keseluruhan:** {r.get('status_battery', '-')} ({r.get('jml_bat_rusak', 0)} Rusak) | **Merk Baterai:** {r.get('tipe_battery', '-')}")
                            for b in r.get('battery_data', []):
                                with st.container(border=True):
                                    bc1, bc2 = st.columns([2, 1])
                                    with bc1:
                                        st.markdown(f"**{b.get('Baterai', 'Baterai')}**")
                                        st.write(f"🔹 **Voltase:** {b.get('Voltase', '-') } V  |  🌡 **Load Charging:** {b.get('Suhu', '-') } A")
                                        st.write(f"🔍 **Kondisi Fisik:** {b.get('Kondisi', '-')}")
                                    bat_urls = b.get('URL_Fotos') or b.get('URL_Foto')
                                    if bat_urls:
                                        st.markdown("")
                                        tampilkan_grid_foto(bat_urls, "📸 Dokumentasi Baterai")
                                st.markdown("")
                            st.markdown("---")
                            st.markdown("### 🌍 Sistem Grounding")
                            st.write(f"⚡ **Tahanan Grounding:** {r.get('earth_ohm', '-')} Ohm")
                            tampilkan_grid_foto(r.get('url_grds'), "📸 Foto Grounding")
                            tampilkan_grid_foto(r.get('extras_baterai'), "📸 Tambahan Baterai & Grounding")

                        with ltab5:
                            st.markdown("📂 **Datalog Universal yang Tersimpan:**")
                            dls = r.get('datalog_files', [])
                            if dls:
                                for dl in dls: st.markdown(f"- 🔗 [{dl['name']}]({dl['url']})")
                            else: st.info("Tidak ada file datalog diunggah pada site ini.")

                        with ltab6:
                            st.markdown("### ⚡ Grafik Analisa Power & Voltage")
                            st.info(f"Sistem sedang melacak file data histori **`{site_id}.xlsx`** di repositori...")
                            
                            excel_filename = f"{site_id}.xlsx"
                            if os.path.exists(excel_filename):
                                try:
                                    df_power = pd.read_excel(excel_filename, engine='openpyxl')
                                    if 'Begin Time' in df_power.columns and 'MinVoltageOfBBU(V)' in df_power.columns:
                                        df_power['Waktu'] = pd.to_datetime(df_power['Begin Time']).dt.strftime('%H:%M')
                                        df_power.set_index('Waktu', inplace=True)
                                        
                                        chart_cols = ['MaxVoltageOfBBU(V)', 'MinVoltageOfBBU(V)', 'AvgVoltageOfBBU(V)']
                                        chart_data = df_power[chart_cols].copy()
                                        chart_data['Batas Hold (55V)'] = 55.0
                                        
                                        st.line_chart(chart_data, color=["#64FFDA", "#FF5252", "#FFC107", "#FFFFFF"])
                                        
                                        min_voltage = df_power['MinVoltageOfBBU(V)'].min()
                                        if min_voltage < 55.0:
                                            st.error(f"🚨 **ANALISA DROP VOLTAGE:** Ditemukan tegangan Drop di bawah batas 55V! Tegangan terendah terekam di angka **{min_voltage} V**.")
                                        else:
                                            st.success(f"✅ **ANALISA STABIL:** Kondisi Power aman. Tegangan berhasil di-hold (tidak jatuh di bawah batas 55V). Tegangan terendah terekam: {min_voltage} V.")
                                    else:
                                        st.warning(f"File '{excel_filename}' ditemukan, tapi struktur kolomnya tidak sesuai format U2000/U2020. Pastikan ada kolom 'Begin Time' dan 'MinVoltageOfBBU(V)'.")
                                except Exception as e:
                                    st.error(f"Gagal membaca file {excel_filename}. Pesan Error: {e}")
                            else:
                                st.warning(f"File log power **`{excel_filename}`** belum di-upload ke repositori/folder lokal server. Silakan upload file excel dengan nama site tersebut untuk melihat grafik otomatis.")

                    # Tombol Revisi Admin ditaruh paling bawah expander agar rapi
                    if st.session_state['role'] == 'Admin':
                        st.markdown("<br>", unsafe_allow_html=True)
                        with st.container(border=True):
                            
                            # JIKA INI PROJECT PO, TAMPILKAN EDITOR TABEL & UPLOAD FOTO SUSULAN
                            if is_project and r.get('proj_data', {}).get('jenis_project') == 'PO':
                                st.markdown("<h4 style='color: var(--primary-color);'>🛠 UPDATE PROGRESS PROJECT & FOTO SUSULAN</h4>", unsafe_allow_html=True)
                                st.info("💡 Update progress di tabel ini. Saat disimpan, Margin & Progress % akan dikalkulasi ulang otomatis!")
                                
                                df_edit_po = pd.DataFrame(r.get('proj_data', {}).get('tabel_pekerjaan', []))
                                if df_edit_po.empty:
                                    df_edit_po = pd.DataFrame(columns=["Site Asal", "Site Tujuan", "Pekerjaan Spesifik", "Budget Per Site", "Start Date", "End Date", "Progress (%)", "Status", "Remark"])
                                
                                updated_po_table = st.data_editor(
                                    df_edit_po, 
                                    num_rows="dynamic", 
                                    use_container_width=True,
                                    column_config={
                                        "Site Asal": st.column_config.TextColumn("Site Asal (Dari)"),
                                        "Site Tujuan": st.column_config.TextColumn("Site Tujuan (Ke)"),
                                        "Pekerjaan Spesifik": st.column_config.TextColumn("Pekerjaan"),
                                        "Budget Per Site": st.column_config.NumberColumn("Budget / Cost (Rp)", min_value=0, default=0, step=100000),
                                        "Start Date": st.column_config.DateColumn("Mulai", format="YYYY-MM-DD"),
                                        "End Date": st.column_config.DateColumn("Target Selesai", format="YYYY-MM-DD"),
                                        "Progress (%)": st.column_config.NumberColumn("Progress (%)", min_value=0, max_value=100, default=0),
                                        "Status": st.column_config.SelectboxColumn("Status", options=["Not Started", "On Progress", "Done", "Pending Issue"])
                                    },
                                    key=f"edit_po_tbl_{i}"
                                )
                                
                                new_status_proj = st.selectbox("Update Status Keseluruhan Project:", ["Not Started", "On Progress", "Done", "Pending"], index=["Not Started", "On Progress", "Done", "Pending"].index(r.get('status', 'Not Started')) if r.get('status', 'Not Started') in ["Not Started", "On Progress", "Done", "Pending"] else 0, key=f"estat_proj_{i}")
                                
                                st.markdown("**Tambah Lampiran Foto Susulan (Otomatis ditambahkan ke laporan):**")
                                up_p_pros = st.file_uploader("📸 Tambah Foto On Progress", accept_multiple_files=True, key=f"up_p_pros_{i}")
                                up_p_akh = st.file_uploader("📸 Tambah Foto Hasil Akhir/BAST", accept_multiple_files=True, key=f"up_p_akh_{i}")
                                
                                if st.button("💾 Simpan Update Project", key=f"bproj_{i}", type="primary"):
                                    with st.spinner("Menyimpan Update Project..."):
                                        r['proj_data']['tabel_pekerjaan'] = updated_po_table.to_dict('records')
                                        
                                        # Hitung ulang Margin & Progress
                                        try: val_po = float(str(r['proj_data'].get('nilai_po', '0')).replace('.', '').replace(',', ''))
                                        except: val_po = 0.0
                                        t_cost = updated_po_table["Budget Per Site"].sum() if not updated_po_table.empty else 0.0
                                        r['proj_data']['total_cost'] = float(t_cost)
                                        r['proj_data']['margin'] = float(val_po - t_cost)
                                        r['proj_data']['avg_progress'] = float(updated_po_table["Progress (%)"].mean() if not updated_po_table.empty else 0.0)
                                        
                                        r['status'] = new_status_proj
                                        
                                        # Menambah foto ke cloud
                                        if up_p_pros:
                                            r['proj_data']['url_proses'] = r['proj_data'].get('url_proses', []) + upload_multiple_images(up_p_pros, "proj_po")
                                        if up_p_akh:
                                            r['proj_data']['url_akhir'] = r['proj_data'].get('url_akhir', []) + upload_multiple_images(up_p_akh, "proj_po")
                                            
                                        # Simpan ke Sheet Project
                                        sheet = connect_gsheets(r.get('_sheet', 'Project'))
                                        if sheet:
                                            sheet.update_cell(r.get('_row', i + 2), 5, new_status_proj)
                                            sheet.update_cell(r.get('_row', i + 2), 6, json.dumps(r))
                                            st.cache_data.clear()
                                    st.success("✅ UPDATE PROJECT BERHASIL TERSIMPAN!")
                                    st.rerun()

                            else:
                                # JIKA BUKAN PROJECT PO (SPS & REGULER BISA REVISI NAMA SITE)
                                st.markdown("<h4 style='color: var(--primary-color);'>🛠 REVISI NAMA SITE / FME</h4>", unsafe_allow_html=True)
                                ce1, ce2 = st.columns(2)
                                with ce1: new_sreg = st.text_input("Edit Nama Site", r.get('site_name',''), key=f"esreg_{i}")
                                with ce2: new_freg = st.text_input("Edit Nama FME / Tim", r.get('teknisi',''), key=f"efreg_{i}")
                                
                                if st.button("💾 Simpan Perubahan", key=f"breg_{i}"):
                                    with st.spinner("Menyimpan..."):
                                        r['site_name'], r['teknisi'] = new_sreg, new_freg
                                        sheet = connect_gsheets(r.get('_sheet', 'Report Preventive'))
                                        if sheet:
                                            sheet.update_cell(r.get('_row', i + 2), 2, new_sreg)
                                            sheet.update_cell(r.get('_row', i + 2), 4, new_freg)
                                            sheet.update_cell(r.get('_row', i + 2), 6, json.dumps(r))
                                            st.cache_data.clear() 
                                    st.success("✅ REVISI BERHASIL!")
                                    st.rerun()

# =========================================================================
# MENU 3: EXCEL LIVE-EDITOR (MONITORING IMPROVEMENT)
# =========================================================================
elif menu == "📈 Monitoring Improvement":
    st.markdown("<h1 style='text-align: center;'>📈 Master Tracker Improvement</h1>", unsafe_allow_html=True)
    
    file_master = "Monitoring_Availability_Improvement_Visit_SPS_NOP_PLK.xlsx"
    
    if not os.path.exists(file_master):
        st.error(f"❌ File '{file_master}' tidak ditemukan di sistem/server. Pastikan Anda telah meletakkan file tersebut satu folder dengan aplikasi.")
    else:
        tab_dashboard, tab_tim, tab_editor = st.tabs(["📊 Kurva S", "👥 Progress Tim (Done/Berjalan)", "📝 Live Editor Data"])
        
        with tab_dashboard:
            st.markdown("### 📈 S-Curve Target vs Aktual")
            st.info("💡 Data grafik ini dibaca secara Real-Time dari **Sheet 'Kurva S'** di file Excel Master Anda.")
            
            try:
                df_kurva = pd.read_excel(file_master, sheet_name='Kurva S', header=None)
                total_site = df_kurva.iloc[6, 13] if pd.notna(df_kurva.iloc[6, 13]) else 0
                minggu_berj = df_kurva.iloc[7, 13] if pd.notna(df_kurva.iloc[7, 13]) else 0
                site_plan = df_kurva.iloc[8, 13] if pd.notna(df_kurva.iloc[8, 13]) else 0
                
                m1, m2, m3 = st.columns(3)
                m1.metric(label="Total Site Target", value=f"{int(total_site)} Site")
                m2.metric(label="Site dengan Plan", value=f"{int(site_plan)} Site")
                m3.metric(label="Minggu Berjalan", value=f"Week {int(minggu_berj)}")
                
                st.markdown("---")
                
                df_chart = df_kurva.iloc[6:58, :11].copy()
                df_chart.columns = ['Week', 'Mulai', 'Label', 'Plan_Minggu', 'Plan_Kumulatif', 'Actual_Minggu', 'Actual_Kumulatif', 'Pct_Plan', 'Pct_Actual', 'Gap', 'Achievement']
                df_chart = df_chart.dropna(subset=['Week', 'Plan_Kumulatif'])
                
                df_chart['Week'] = "W" + df_chart['Week'].astype(str)
                df_chart.set_index('Week', inplace=True)
                
                col_chart, col_table = st.columns([2, 1])
                with col_chart:
                    st.markdown("**Perbandingan Kumulatif Plan vs Actual**")
                    st.line_chart(df_chart[['Plan_Kumulatif', 'Actual_Kumulatif']], color=["#FF5252", "#64FFDA"])
                    
                with col_table:
                    st.markdown("**Tabel Gap Mingguan**")
                    st.dataframe(df_chart[['Plan_Kumulatif', 'Actual_Kumulatif', 'Gap']], use_container_width=True)
            except Exception as e:
                st.error(f"Gagal memvisualisasikan Grafik S-Curve: {e}")

        with tab_tim:
            st.markdown("### 👥 Dashboard Eksekusi Lapangan")
            st.info("💡 Sistem akan membaca Sheet **Tracker Improvement** untuk menghitung berapa site yang sudah Selesai (Done) dan masih dalam pengerjaan (On Progress).")
            
            try:
                df_track = pd.read_excel(file_master, sheet_name='Tracker Improvement', header=None)
                headers_track = df_track.iloc[3].fillna("").astype(str).tolist()
                df_t = df_track.iloc[4:].copy()
                df_t.columns = headers_track
                
                status_cols = [c for c in headers_track if 'status' in c.lower() or 'progress' in c.lower() or 'keterangan' in c.lower() or 'aktual' in c.lower()]
                default_idx = headers_track.index(status_cols[0]) if status_cols else 0
                
                st.markdown("**Pilih kolom pada Excel yang berisi status pengerjaan tim:**")
                sel_col = st.selectbox("Pilih Kolom Status:", headers_track, index=default_idx)
                
                if sel_col:
                    df_t[sel_col] = df_t[sel_col].astype(str).str.strip().str.upper()
                    df_valid = df_t[~df_t[sel_col].isin(['NAN', 'NAT', 'NONE', ''])]
                    
                    counts = df_valid[sel_col].value_counts().reset_index()
                    counts.columns = ['Status', 'Jumlah']
                    
                    if not counts.empty:
                        done_kws = ['DONE', 'SELESAI', 'CLOSE', 'OK', 'COMPLETED']
                        prog_kws = ['PROGRESS', 'BERJALAN', 'ON GOING', 'OPEN', 'PENDING', 'ON PROGRESS']
                        
                        jml_done = counts[counts['Status'].apply(lambda x: any(k in x for k in done_kws))]['Jumlah'].sum()
                        jml_prog = counts[counts['Status'].apply(lambda x: any(k in x for k in prog_kws))]['Jumlah'].sum()
                        jml_lain = counts['Jumlah'].sum() - (jml_done + jml_prog)
                        
                        st.markdown("#### 🎯 Hasil Kalkulasi Progress")
                        c_t1, c_t2, c_t3 = st.columns(3)
                        c_t1.metric("✅ TOTAL SELESAI (Done)", f"{jml_done} Site")
                        c_t2.metric("⏳ SEDANG BERJALAN (On Progress)", f"{jml_prog} Site")
                        c_t3.metric("📝 Lainnya / Belum Mulai", f"{jml_lain} Site")
                        
                        st.markdown("#### 📊 Grafik Breakdown Status")
                        st.bar_chart(counts.set_index('Status'), color="#64FFDA")
                        
                        with st.expander("Tampilkan Rekap Tabel Mentah Status", expanded=False):
                            st.dataframe(counts, use_container_width=True)
                    else:
                        st.warning(f"Data pada kolom '{sel_col}' kosong atau seluruhnya terbaca sebagai cell kosong.")
            except Exception as e:
                st.error(f"Gagal memproses data tracker tim: {e}")

        with tab_editor:
            st.markdown("### 📝 Tabel Master Tracker & Viewer")
            st.write("Klik ganda (*double click*) pada sel tabel untuk merubah isi data secara instan.")
            
            sheet_options = ["Tracker Improvement", "Jadwal Visit SPS", "Dashboard", "Kurva S"]
            sheet_choice = st.selectbox("Pilih Sheet Excel yang Ingin Diedit / Dilihat:", sheet_options)
            
            try:
                df_raw = pd.read_excel(file_master, sheet_name=sheet_choice, header=None)
                
                if sheet_choice == "Dashboard":
                    st.info("💡 Karena sheet Dashboard bawaan Excel biasanya berisi desain sel yang di-merge dan banyak bagan, tampilannya di web mungkin terlihat sebagai sekumpulan teks kasar. Gunakan **Tab 👥 Progress Tim** di atas untuk tampilan Dashboard yang jauh lebih rapi.")
                    st.dataframe(df_raw, use_container_width=True, height=500)
                
                else:
                    headers = df_raw.iloc[3].fillna("").astype(str).tolist()
                    df_data = df_raw.iloc[4:].copy()
                    df_data.columns = headers
                    df_data = df_data.reset_index(drop=True)
                    
                    edited_df = st.data_editor(
                        df_data,
                        use_container_width=True,
                        num_rows="dynamic",
                        key=f"editor_{sheet_choice}"
                    )
                    
                    if st.button("💾 Simpan Perubahan & Download Excel Utuh", type="primary"):
                        state_key = f"editor_{sheet_choice}"
                        changes = st.session_state[state_key]
                        
                        if changes.get("edited_rows") or changes.get("added_rows"):
                            with st.spinner("Menyuntikkan data baru ke Master Excel (Menjaga Format)..."):
                                import openpyxl
                                wb = openpyxl.load_workbook(file_master)
                                ws = wb[sheet_choice]
                                
                                for row_idx_str, col_changes in changes.get("edited_rows", {}).items():
                                    row_idx = int(row_idx_str)
                                    excel_row = row_idx + 5 
                                    
                                    for col_name, new_val in col_changes.items():
                                        if col_name in headers:
                                            col_idx = headers.index(col_name) + 1
                                            ws.cell(row=excel_row, column=col_idx).value = new_val
                                            
                                for added_row in changes.get("added_rows", []):
                                    excel_row = ws.max_row + 1
                                    for col_name, new_val in added_row.items():
                                        if col_name in headers:
                                            col_idx = headers.index(col_name) + 1
                                            ws.cell(row=excel_row, column=col_idx).value = new_val
                                
                                out_buffer = BytesIO()
                                wb.save(out_buffer)
                                
                                st.success("✅ Berhasil! Perubahan telah disuntikkan ke file asli.")
                                st.download_button(
                                    label=f"📥 Download File Master Terupdate",
                                    data=out_buffer.getvalue(),
                                    file_name=f"Updated_{file_master}",
                                    mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                                    type="primary"
                                )
                        else:
                            st.warning("Belum ada sel yang diubah. Klik ganda pada tabel untuk mengubah isinya.")
                            
            except Exception as e:
                st.error(f"Terjadi kendala saat membaca sheet {sheet_choice}: {e}")

# =========================================================================
# MENU 4: DAPOT POWER HASIL PMS (MENU TERPISAH KHUSUS EXECUTIVE DASHBOARD)
# =========================================================================
elif menu == "⚡ Dapot Power Hasil PMS":
    st.markdown("<h1 style='text-align: center;'>⚡ Dapot Power Hasil PMS (Palangkaraya)</h1>", unsafe_allow_html=True)
    st.markdown("<p style='text-align: center; opacity: 0.8;'>Power Infrastructure Performance Dashboard — Analisa Perbandingan Bulanan Tegangan & Arus AC/DC</p>", unsafe_allow_html=True)

    dapot_file = "Dapot_Power_Executive_Dashboard.xlsx"

    if not os.path.exists(dapot_file):
        st.error(f"❌ File **`{dapot_file}`** belum ditemukan di folder server/repositori. Pastikan file tersebut sudah di-upload ke GitHub satu folder dengan `App.py`.")
    else:
        try:
            df_kpi = pd.read_excel(dapot_file, sheet_name='DB_KPI')
            df_trend = pd.read_excel(dapot_file, sheet_name='DB_TREND')

            df_kpi['SITE_ID'] = df_kpi['SITE_ID'].astype(str).str.strip().str.upper()
            df_trend['SITE_ID'] = df_trend['SITE_ID'].astype(str).str.strip().str.upper()
            df_trend['BULAN'] = df_trend['BULAN'].astype(str).str.strip()

            visit_counts = df_trend['SITE_ID'].value_counts()
            sites_rutin = visit_counts[visit_counts > 1].index.tolist()
            sites_sekali = visit_counts[visit_counts == 1].index.tolist()
            all_visited_sites = visit_counts.index.tolist()

            m1, m2, m3, m4 = st.columns(4)
            m1.metric("🏢 Total Site Terdata (KPI)", f"{len(df_kpi)} Site")
            m2.metric("🔄 Site Dikunjungi Rutin (>1x)", f"{len(sites_rutin)} Site")
            m3.metric("📍 Site 1x Kunjungan", f"{len(sites_sekali)} Site")
            m4.metric("📅 Total Record PMS", f"{len(df_trend)} Kunjungan")

            st.markdown("---")

            tab_comp, tab_rekap, tab_master = st.tabs([
                "📈 Compare Bulanan AC & DC (Per Site)", 
                "📋 Daftar Site Rutin vs Pernah Dikunjungi", 
                "🗂️ Master Data Konsolidasi & Live Editor"
            ])

            with tab_comp:
                st.markdown("### 🔍 Filter & Pilih Site untuk Compare Bulanan")
                
                c_f1, c_f2 = st.columns([1.5, 2])
                with c_f1:
                    kat_kunjungan = st.radio(
                        "Filter Kategori Kunjungan Site:",
                        [
                            "🔄 Dikunjungi Rutin (>1 Bulan / Bisa Compare)", 
                            "📍 Pernah Dikunjungi (1x Kunjungan)", 
                            "🌐 Semua Site Terdata"
                        ]
                    )
                
                if "Rutin" in kat_kunjungan:
                    pilihan_sites = sites_rutin
                elif "1x" in kat_kunjungan:
                    pilihan_sites = sites_sekali
                else:
                    pilihan_sites = all_visited_sites

                with c_f2:
                    format_opsi = {s: f"{s} ({visit_counts.get(s, 0)}x Kunjungan)" for s in pilihan_sites}
                    selected_site = st.selectbox(
                        "Pilih Site ID untuk Melihat Grafik & KPI:",
                        options=pilihan_sites,
                        format_func=lambda x: format_opsi.get(x, x)
                    )

                if selected_site:
                    kpi_row = df_kpi[df_kpi['SITE_ID'] == selected_site]
                    st.markdown(f"#### 📌 Status KPI Terakhir — Site **{selected_site}**")
                    
                    k1, k2, k3, k4, k5, k6 = st.columns(6)
                    if not kpi_row.empty:
                        kr = kpi_row.iloc[0]
                        k1.metric("Last Inspection", str(kr.get('TGL_UPDATE', '-')))
                        k2.metric("Status PLN", str(kr.get('STATUS_PLN', '-')).upper())
                        k3.metric("Total Baterai", f"{kr.get('TOTAL_BATERAI', 0)} Bank")
                        k4.metric("Total Modul", f"{kr.get('TOTAL_MODUL', 0)} Pcs")
                        k5.metric("Latest Voltage AC", f"{kr.get('TEGANGAN_AC', 0)} V")
                        k6.metric("Latest Load DC", f"{kr.get('LOAD_DC', 0)} A")
                    else:
                        st.info("Data ringkasan di sheet DB_KPI belum tersedia untuk Site ini.")

                    st.markdown("---")

                    df_site_trend = df_trend[df_trend['SITE_ID'] == selected_site].sort_values(by='BULAN').copy()
                    
                    if not df_site_trend.empty:
                        for col_num in ['VAC', 'VDC', 'IAC', 'LDC']:
                            df_site_trend[col_num] = pd.to_numeric(df_site_trend[col_num], errors='coerce').fillna(0.0)

                        df_site_trend['Delta_VAC'] = df_site_trend['VAC'].diff().fillna(0.0).round(2)
                        df_site_trend['Delta_VDC'] = df_site_trend['VDC'].diff().fillna(0.0).round(2)
                        df_site_trend['Delta_IAC'] = df_site_trend['IAC'].diff().fillna(0.0).round(2)
                        df_site_trend['Delta_LDC'] = df_site_trend['LDC'].diff().fillna(0.0).round(2)

                        df_plot = df_site_trend.set_index('BULAN')

                        st.markdown(f"#### 📊 Grafik Perbandingan Bulanan AC & DC — **{selected_site}** ({len(df_site_trend)} Bulan Kunjungan)")
                        
                        col_g1, col_g2 = st.columns(2)
                        with col_g1:
                            with st.container(border=True):
                                st.markdown("**⚡ Perbandingan Tegangan: AC (VAC) vs DC (VDC) per Bulan**")
                                st.line_chart(df_plot[['VAC', 'VDC']], color=["#FF5252", "#00E676"])
                                st.caption("Merah = Tegangan AC (Volt) | Hijau = Tegangan DC Rectifier (Volt)")
                                
                        with col_g2:
                            with st.container(border=True):
                                st.markdown("**🔌 Perbandingan Arus/Beban: Arus AC (IAC) vs Load DC (LDC) per Bulan**")
                                st.line_chart(df_plot[['IAC', 'LDC']], color=["#FFC107", "#2979FF"])
                                st.caption("Kuning = Arus AC (Ampere) | Biru = Beban Load DC (Ampere)")

                        col_g3, col_g4 = st.columns(2)
                        with col_g3:
                            with st.container(border=True):
                                st.markdown("**🔋 Detail Trend Tegangan DC (VDC) Antar Bulan**")
                                st.bar_chart(df_plot[['VDC']], color="#00E676")
                        with col_g4:
                            with st.container(border=True):
                                st.markdown("**⚡ Detail Trend Tegangan AC (VAC) Antar Bulan**")
                                st.bar_chart(df_plot[['VAC']], color="#FF5252")

                        st.markdown("#### 📋 Tabel Histori & Selisih (Compare) Antar Bulan")
                        st.dataframe(
                            df_site_trend[['BULAN', 'VAC', 'Delta_VAC', 'VDC', 'Delta_VDC', 'IAC', 'Delta_IAC', 'LDC', 'Delta_LDC']],
                            use_container_width=True,
                            hide_index=True
                        )
                    else:
                        st.warning("Data trend bulanan tidak ditemukan untuk site ini.")

            with tab_rekap:
                st.markdown("### 📋 Rekapitulasi Frekuensi Kunjungan Site PMS")
                
                rekap_df = df_trend.groupby('SITE_ID').agg(
                    Jumlah_Kunjungan=('BULAN', 'count'),
                    Daftar_Bulan_Kunjungan=('BULAN', lambda x: ", ".join(sorted(x.astype(str).unique()))),
                    Rata_VAC=('VAC', 'mean'),
                    Rata_VDC=('VDC', 'mean'),
                    Rata_IAC=('IAC', 'mean'),
                    Rata_LDC=('LDC', 'mean')
                ).reset_index()

                rekap_df['Kategori_Kunjungan'] = rekap_df['Jumlah_Kunjungan'].apply(
                    lambda x: "Dikunjungi Rutin (>1x)" if x > 1 else "Pernah Dikunjungi (1x)"
                )
                rekap_df = rekap_df.sort_values(by='Jumlah_Kunjungan', ascending=False).reset_index(drop=True)
                rekap_df[['Rata_VAC', 'Rata_VDC', 'Rata_IAC', 'Rata_LDC']] = rekap_df[['Rata_VAC', 'Rata_VDC', 'Rata_IAC', 'Rata_LDC']].round(2)

                filter_tabel = st.radio(
                    "Tampilkan Tabel:", 
                    ["Semua Site", "Hanya Site Dikunjungi Rutin (>1x)", "Hanya Site Pernah Dikunjungi (1x)"], 
                    horizontal=True
                )
                
                if "Rutin" in filter_tabel:
                    tampil_rekap = rekap_df[rekap_df['Jumlah_Kunjungan'] > 1]
                elif "1x" in filter_tabel:
                    tampil_rekap = rekap_df[rekap_df['Jumlah_Kunjungan'] == 1]
                else:
                    tampil_rekap = rekap_df

                st.dataframe(tampil_rekap, use_container_width=True, height=450)

            with tab_master:
                st.markdown("### 🗂️ Explorer & Live Editor Database Dapot Power")
                pilih_sheet_dapot = st.selectbox(
                    "Pilih Sheet dari File Dapot_Power_Executive_Dashboard.xlsx:", 
                    ["DB_TREND", "DB_KPI", "Master_Data_Konsolidasi"]
                )
                
                df_dapot_sheet = pd.read_excel(dapot_file, sheet_name=pilih_sheet_dapot)
                edited_dapot = st.data_editor(
                    df_dapot_sheet, 
                    use_container_width=True, 
                    num_rows="dynamic", 
                    key=f"ed_dapot_{pilih_sheet_dapot}"
                )

                if st.button("💾 Simpan Perubahan & Download File Dapot Power", type="primary"):
                    state_dapot = st.session_state[f"ed_dapot_{pilih_sheet_dapot}"]
                    if state_dapot.get("edited_rows") or state_dapot.get("added_rows"):
                        with st.spinner("Menyuntikkan pembaruan ke file Dapot_Power_Executive_Dashboard.xlsx..."):
                            import openpyxl
                            wb_d = openpyxl.load_workbook(dapot_file)
                            ws_d = wb_d[pilih_sheet_dapot]
                            headers_d = list(df_dapot_sheet.columns)

                            for r_idx_str, c_changes in state_dapot.get("edited_rows", {}).items():
                                xl_row = int(r_idx_str) + 2
                                for c_name, n_val in c_changes.items():
                                    if c_name in headers_d:
                                        ws_d.cell(row=xl_row, column=headers_d.index(c_name) + 1).value = n_val

                            for add_r in state_dapot.get("added_rows", []):
                                xl_row = ws_d.max_row + 1
                                for c_name, n_val in add_r.items():
                                    if c_name in headers_d:
                                        ws_d.cell(row=xl_row, column=headers_d.index(c_name) + 1).value = n_val

                            buf_d = BytesIO()
                            wb_d.save(buf_d)
                            st.success("✅ Perubahan berhasil disimpan ke dalam struktur asli Excel!")
                            st.download_button(
                                label="📥 Download Dapot_Power_Executive_Dashboard.xlsx Terupdate",
                                data=buf_d.getvalue(),
                                file_name=f"Updated_{dapot_file}",
                                mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                                type="primary"
                            )
                    else:
                        st.warning("Belum ada sel yang diubah pada tabel di atas.")

        except Exception as e:
            st.error(f"Terjadi kesalahan saat memuat Dapot Power Executive Dashboard: {e}")

# =========================================================================
# MENU 5: DAILY TRACKER (TO DO LIST NOP PALANGKARAYA)
# =========================================================================
elif menu == "📅 Daily Tracker Palangkaraya":
    st.markdown("<h1 style='text-align: center;'>📅 Daily Tracker & To Do List (Palangkaraya)</h1>", unsafe_allow_html=True)
    st.markdown("<p style='text-align: center; opacity: 0.8;'>Update rencana kerja harian, realisasi site, dan pencocokan otomatis dengan Master Data Sidata.</p>", unsafe_allow_html=True)
    
    file_todo = "To Do List NOP Palangkaraya.xlsx"
    
    if not os.path.exists(file_todo):
        st.error(f"❌ File '{file_todo}' tidak ditemukan di server. Pastikan file tersebut sudah di-upload satu folder dengan aplikasi `App.py`.")
    else:
        fme_plk_list, _, site_plk_list, _ = load_master_lists()
        
        tab_harian, tab_database = st.tabs(["📝 Form Update Tim Onsite", "🗂️ Live Editor Master To-Do List"])
        
        with tab_harian:
            st.info("💡 **Khusus Tim Onsite / FME:** Data Site ID dan Nama FME di bawah ini sudah otomatis membaca dari Repositori Sidata. Silakan pilih dan laporkan progress harian.")
            
            c_td1, c_td2 = st.columns(2)
            with c_td1:
                pilih_site_td = st.selectbox("📍 Pilih Site ID Target (Auto-sync):", [""] + site_plk_list, key="td_site")
                pilih_fme_td = st.selectbox("👷 Nama FME / Tim Onsite:", [""] + fme_plk_list, key="td_fme")
            with c_td2:
                tanggal_td = st.date_input("📅 Tanggal Pengerjaan:", value=datetime.date.today())
                status_td = st.selectbox("📊 Update Status Harian:", ["On Progress", "Done", "Pending", "Cancel / Reschedule"])
            
            catatan_td = st.text_area("📝 Keterangan / Remarks Lapangan:", placeholder="Ketik keterangan mengenai pengerjaan hari ini atau alasan jika pending...")
            
            if st.button("💾 Simpan Update ke To Do List", type="primary", use_container_width=True):
                if not pilih_site_td or not pilih_fme_td:
                    st.error("⚠️ Site ID dan Nama FME wajib dipilih!")
                else:
                    with st.spinner("Menyimpan ke file Excel To Do List..."):
                        try:
                            import openpyxl
                            wb_todo = openpyxl.load_workbook(file_todo)
                            
                            sheet_target = "to do list" if "to do list" in wb_todo.sheetnames else wb_todo.sheetnames[0]
                            ws_todo = wb_todo[sheet_target]
                            
                            next_row = ws_todo.max_row + 1
                            
                            ws_todo.cell(row=next_row, column=11).value = tanggal_td.strftime("%Y-%m-%d") # Start Date
                            ws_todo.cell(row=next_row, column=3).value = pilih_site_td # Site ID
                            ws_todo.cell(row=next_row, column=4).value = pilih_fme_td # FME Name
                            ws_todo.cell(row=next_row, column=13).value = status_td # Status Task
                            ws_todo.cell(row=next_row, column=14).value = catatan_td # Remark
                            
                            wb_todo.save(file_todo)
                            st.success(f"✅ Laporan Diterima! Data kunjungan Site {pilih_site_td} oleh {pilih_fme_td} telah ter-update di file Tracker.")
                        except Exception as e:
                            st.error(f"Terjadi kesalahan saat menyimpan ke Excel: {e}")

        with tab_database:
            st.markdown("### 🗂️ Tabel Master To Do List (Bisa di-edit langsung oleh Admin)")
            st.write("Klik dua kali pada tabel di bawah (khususnya bagian kolom Status) untuk mengubah data secara cepat.")
            
            try:
                sheet_tujuan = "to do list"
                try:
                    df_aktif = pd.read_excel(file_todo, sheet_name=sheet_tujuan, header=2)
                    if df_aktif.empty or 'Site ID' not in df_aktif.columns:
                        df_aktif = pd.read_excel(file_todo, sheet_name=sheet_tujuan)
                except:
                    df_aktif = pd.read_excel(file_todo, sheet_name=0)
                
                edited_todo = st.data_editor(df_aktif, use_container_width=True, num_rows="dynamic", key="ed_todo_master")
                
                if st.button("💾 Simpan Perubahan Excel Permanen", type="primary"):
                    state_todo = st.session_state["ed_todo_master"]
                    if state_todo.get("edited_rows") or state_todo.get("added_rows"):
                        with st.spinner("Menyuntikkan pembaruan formasi langsung ke file Excel..."):
                            import openpyxl
                            wb_t = openpyxl.load_workbook(file_todo)
                            sheet_target = "to do list" if "to do list" in wb_t.sheetnames else wb_t.sheetnames[0]
                            ws_t = wb_t[sheet_target]
                            headers_t = list(df_aktif.columns)

                            offset_row = 3
                            
                            for r_idx_str, c_changes in state_todo.get("edited_rows", {}).items():
                                xl_row = int(r_idx_str) + 1 + offset_row
                                for c_name, n_val in c_changes.items():
                                    if c_name in headers_t:
                                        col_index = headers_t.index(c_name) + 1
                                        ws_t.cell(row=xl_row, column=col_index).value = n_val

                            for add_r in state_todo.get("added_rows", []):
                                xl_row = ws_t.max_row + 1
                                for c_name, n_val in add_r.items():
                                    if c_name in headers_t:
                                        col_index = headers_t.index(c_name) + 1
                                        ws_t.cell(row=xl_row, column=col_index).value = n_val

                            buf_t = BytesIO()
                            wb_t.save(buf_t)
                            st.success("✅ Perubahan berhasil disimpan permanen ke Master To Do List!")
                            st.download_button(
                                label="📥 Download File Tracker Terupdate",
                                data=buf_t.getvalue(),
                                file_name=f"Updated_To_Do_List_PLK.xlsx",
                                mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                                type="primary"
                            )
                    else:
                        st.warning("Belum ada data pada tabel yang Bapak ubah.")
                        
            except Exception as e:
                st.error(f"Gagal memproses Live Editor To Do List: {e}.")

# =========================================================================
# MENU 6: LOGISTIK & SPMS (MANAJEMEN STOK GUDANG)
# =========================================================================
elif menu == "📦 Logistik & SPMS":
    st.markdown("<h1 style='text-align: center;'>📦 Manajemen Logistik & SPMS</h1>", unsafe_allow_html=True)
    st.markdown("<p style='text-align: center; opacity: 0.8;'>Sistem Pencatatan Barang Masuk (Inbound), Barang Keluar (Outbound), dan Auto-Kalkulasi Stok.</p>", unsafe_allow_html=True)
    
    sheet_logistik = connect_gsheets("Data Logistik")
    if sheet_logistik is None:
        st.error("⚠️ **PENTING:** Silakan buat Sheet baru bernama **'Data Logistik'** (tanpa tanda kutip) di dalam file Google Spreadsheet Database Utama Anda.")
        st.stop()
        
    data_log = sheet_logistik.get_all_values()
    headers_standar = ["Tanggal", "Tipe Transaksi", "Kategori", "Merek / Nama Barang", "Untuk Genset/Mesin", "Jumlah", "PIC / Pengambil", "Site ID", "Keterangan"]
    
    # AUTO-FIX: Cek jika sheet kosong atau baris 1 berisi sel kosong/spasi gaib
    if not data_log or len(data_log) == 0 or not any(str(c).strip() for c in data_log[0]):
        sheet_logistik.clear() # Bersihkan sisa format/spasi di Sheet
        sheet_logistik.insert_row(headers_standar, index=1)
        data_log = [headers_standar]
        
    header_db = data_log[0]
    data_rows = data_log[1:]
    
    # Ratakan panjang baris dengan panjang header agar tidak Error di Pandas
    data_rows_padded = [row + [''] * (len(header_db) - len(row)) for row in data_rows]
    
    df_log = pd.DataFrame(data_rows_padded, columns=header_db)
    
    # Proteksi Darurat jika nama kolom masih belum terbaca sempurna
    if 'Tipe Transaksi' not in df_log.columns: df_log['Tipe Transaksi'] = ''
    if 'Merek / Nama Barang' not in df_log.columns: df_log['Merek / Nama Barang'] = ''
    if 'Jumlah' not in df_log.columns: df_log['Jumlah'] = '0'
    if 'Kategori' not in df_log.columns: df_log['Kategori'] = ''
    
    tab_in, tab_out, tab_dash = st.tabs(["📥 Inbound (Barang Masuk)", "📤 Outbound (Ambil Barang)", "📊 Dashboard Stok & Ekspor"])
    
    kategori_opsi = ["Filter Oli", "Filter Solar", "Filter Udara", "Oli Mesin", "Water Separator", "Sparepart Genset", "Modul Kontroler", "Material CME", "Lainnya"]
    
    with tab_in:
        st.markdown("### 📥 Form Barang Masuk Gudang (Inbound)")
        with st.container(border=True):
            c1, c2 = st.columns(2)
            with c1:
                tgl_in = st.date_input("📅 Tanggal Masuk", key="in_tgl")
                kat_in = st.selectbox("📂 Kategori Barang", kategori_opsi, key="in_kat")
                nama_in = st.text_input("🏷️ Merek / Nama Barang Lengkap (Wajib)", placeholder="Cth: Filter Oli Donaldson / Oli Meditran SX", key="in_nama")
            with c2:
                genset_in = st.text_input("⚙️ Diperuntukkan Untuk Tipe Genset", placeholder="Cth: Perkins 20kVA / Kubota / Umum", key="in_gen")
                qty_in = st.number_input("📦 Jumlah Masuk (Pcs/Liter)", min_value=1, value=1, key="in_qty")
                pic_in = st.text_input("👨‍💼 PIC / Penerima Gudang", placeholder="Nama penerima barang", key="in_pic")
            ket_in = st.text_input("📝 Keterangan (Opsional)", placeholder="Cth: DO dari pusat / Pembelian lokal", key="in_ket")
            
            if st.button("💾 SIMPAN DATA MASUK (INBOUND)", type="primary", use_container_width=True, key="btn_in"):
                if not nama_in: st.error("⚠️ Merek/Nama Barang Wajib diisi!")
                else:
                    with st.spinner("Menyimpan stok ke Database..."):
                        row_in = [tgl_in.strftime("%Y-%m-%d"), "INBOUND", kat_in, nama_in, genset_in, qty_in, pic_in, "-", ket_in]
                        sheet_logistik.insert_row(row_in, index=2)
                    st.success(f"✅ Berhasil! {qty_in} qty '{nama_in}' telah ditambahkan ke dalam stok gudang.")
                    st.rerun()
                    
    with tab_out:
        st.markdown("### 📤 Form Pengambilan Barang (Outbound)")
        with st.container(border=True):
            st.info("💡 **Catatan:** Pilih Merek Barang yang sudah terdaftar di stok, atau ketik manual jika tidak ada.")
            c3, c4 = st.columns(2)
            with c3:
                tgl_out = st.date_input("📅 Tanggal Pengambilan", key="out_tgl")
                kat_out = st.selectbox("📂 Kategori Barang Keluar", kategori_opsi, key="out_kat")
                
                stok_tersedia = df_log[df_log['Tipe Transaksi'] == 'INBOUND']['Merek / Nama Barang'].unique().tolist() if not df_log.empty else []
                if stok_tersedia:
                    nama_out = st.selectbox("🏷️ Pilih Merek / Nama Barang dari Stok", ["- Ketik Manual / Barang Baru -"] + stok_tersedia, key="out_nama_sel")
                    if nama_out == "- Ketik Manual / Barang Baru -":
                        nama_out = st.text_input("Ketik Merek / Nama Barang:", key="out_nama_txt")
                else:
                    nama_out = st.text_input("🏷️ Merek / Nama Barang Keluar", key="out_nama_txt")
                    
            with c4:
                qty_out = st.number_input("📦 Jumlah Diambil (Pcs/Liter)", min_value=1, value=1, key="out_qty")
                fme_plk_list, _, site_plk_list, _ = load_master_lists()
                pic_out = st.selectbox("👷 Nama Tim / FME Pengambil", [""] + fme_plk_list, key="out_pic") if fme_plk_list else st.text_input("👷 FME / Tim Pengambil")
                site_out = st.selectbox("📍 Digunakan Untuk Site Target", [""] + site_plk_list, key="out_site") if site_plk_list else st.text_input("📍 Digunakan Untuk Site")
                
            ket_out = st.text_input("📝 Keterangan Pengerjaan", placeholder="Cth: PM Genset / Perbaikan Breakdown...", key="out_ket")
            
            if st.button("🚀 SIMPAN PENGAMBILAN & POTONG STOK", type="primary", use_container_width=True, key="btn_out"):
                if not nama_out or not pic_out or not site_out: 
                    st.error("⚠️ Nama Barang, Nama FME, dan Site Target wajib diisi agar bisa dilacak!")
                else:
                    with st.spinner("Memotong stok dan merekam data..."):
                        row_out = [tgl_out.strftime("%Y-%m-%d"), "OUTBOUND", kat_out, nama_out, "-", qty_out, pic_out, site_out, ket_out]
                        sheet_logistik.insert_row(row_out, index=2)
                    st.success(f"✅ Pengambilan tercatat! {qty_out} qty '{nama_out}' telah dikeluarkan untuk Site {site_out}.")
                    st.rerun()
                    
    with tab_dash:
        st.markdown("### 📊 Kalkulasi Stok Otomatis (In vs Out)")
        if df_log.empty:
            st.info("Belum ada data transaksi Logistik di dalam Database.")
        else:
            try:
                df_calc = df_log.copy()
                df_calc['Jumlah'] = pd.to_numeric(df_calc['Jumlah'], errors='coerce').fillna(0)
                
                # Kalkulasi Inbound
                df_in = df_calc[df_calc['Tipe Transaksi'] == 'INBOUND'].groupby(['Kategori', 'Merek / Nama Barang', 'Untuk Genset/Mesin'])['Jumlah'].sum().reset_index()
                df_in.rename(columns={'Jumlah': 'Total Masuk'}, inplace=True)
                
                # Kalkulasi Outbound
                df_out = df_calc[df_calc['Tipe Transaksi'] == 'OUTBOUND'].groupby(['Kategori', 'Merek / Nama Barang'])['Jumlah'].sum().reset_index()
                df_out.rename(columns={'Jumlah': 'Total Keluar'}, inplace=True)
                
                # Gabungkan & Hitung Sisa
                df_stok = pd.merge(df_in, df_out, on=['Kategori', 'Merek / Nama Barang'], how='left').fillna(0)
                df_stok['Sisa Stok Tersedia'] = df_stok['Total Masuk'] - df_stok['Total Keluar']
                
                # Highlight Stok Habis
                def color_stock(val):
                    color = '#FF5252' if val <= 0 else '#00E676'
                    return f'color: {color}; font-weight: bold;'
                
                st.dataframe(df_stok.style.map(color_stock, subset=['Sisa Stok Tersedia']), use_container_width=True)
                
                st.markdown("### 📥 Ekspor & Rekap Laporan Logistik")
                c_dl1, c_dl2 = st.columns(2)
                with c_dl1:
                    csv_riwayat = df_log.to_csv(index=False).encode('utf-8')
                    st.download_button("📥 Download Riwayat Keluar-Masuk (CSV)", data=csv_riwayat, file_name=f"Riwayat_In_Out_Logistik_{datetime.date.today()}.csv", mime="text/csv", use_container_width=True)
                with c_dl2:
                    csv_stok = df_stok.to_csv(index=False).encode('utf-8')
                    st.download_button("📊 Download Ringkasan Stok (CSV)", data=csv_stok, file_name=f"Sisa_Stok_Logistik_{datetime.date.today()}.csv", mime="text/csv", type="primary", use_container_width=True)
                    
            except Exception as e:
                st.error(f"Terjadi kesalahan saat mengalkulasi stok: {e}")
# =========================================================================
# MENU 7: EXECUTIVE PROJECT DASHBOARD (TERPISAH & TERPROTEKSI SANDI)
# =========================================================================
elif menu == "💎 Executive Project Dashboard":
    if 'proj_dash_unlocked' not in st.session_state:
        st.session_state['proj_dash_unlocked'] = False

    if not st.session_state['proj_dash_unlocked']:
        st.markdown("<br>", unsafe_allow_html=True)
        c_lk1, c_lk2, c_lk3 = st.columns([2.5, 5, 2.5])
        with c_lk2:
            st.markdown("""
                <div style="background: linear-gradient(145deg, #0f172a, #1e293b); padding: 40px 30px; border-radius: 20px; border: 2px solid #FFD700; box-shadow: 0 15px 35px rgba(0,0,0,0.4); text-align: center; color: white;">
                    <div style="font-size: 60px; margin-bottom: 10px;">🔐</div>
                    <h2 style="color: #FFD700 !important; margin-bottom: 8px; font-weight: 900; letter-spacing: 1px;">EXECUTIVE PROJECT VAULT</h2>
                    <p style="color: #cbd5e1; font-size: 14px; margin-bottom: 25px;">Area Rahasia Finansial PO, SVA, Profit Margin & Interactive Timeline.<br>Masukkan kata sandi otorisasi untuk membuka dashboard.</p>
                </div>
            """, unsafe_allow_html=True)
            st.write("")
            pwd_vault = st.text_input("Kata Sandi Otorisasi Project:", type="password", placeholder="Masukkan kata sandi rahasia...", key="input_pwd_vault")
            if st.button("🔓 BUKA EXECUTIVE DASHBOARD", type="primary", use_container_width=True):
                import base64
                if base64.b64encode(pwd_vault.encode()).decode() == "WXVpMjAyNSM=":
                    st.session_state['proj_dash_unlocked'] = True
                    st.rerun()
                else:
                    st.error("❌ Kata sandi otorisasi salah! Akses ke Executive Project ditolak.")
        st.stop()

    # DESAIN LUXURY INTERNATIONAL PMO DASHBOARD
    st.markdown("""
        <style>
        .lux-header {
            background: linear-gradient(135deg, #090d16 0%, #172554 50%, #0f172a 100%);
            padding: 30px 35px;
            border-radius: 20px;
            border: 1px solid rgba(255, 215, 0, 0.35);
            box-shadow: 0 12px 30px rgba(0,0,0,0.25);
            color: white;
            margin-bottom: 24px;
        }
        .lux-kpi {
            background: linear-gradient(145deg, #111827, #1f2937);
            border: 1px solid rgba(255,255,255,0.1);
            border-left: 5px solid #FFD700;
            border-radius: 16px;
            padding: 18px 20px;
            color: white;
            box-shadow: 0 8px 20px rgba(0,0,0,0.18);
            margin-bottom: 15px;
        }
        .lux-kpi-title { font-size: 11px; text-transform: uppercase; letter-spacing: 1.2px; color: #9ca3af; font-weight: 700; margin-bottom: 6px; }
        .lux-kpi-val { font-size: 21px; font-weight: 900; color: #ffffff; letter-spacing: 0.5px; }
        .lux-kpi-sub { font-size: 11px; color: #10b981; font-weight: 700; margin-top: 4px; }
        .tl-card {
            background: var(--secondary-background-color);
            border: 1px solid rgba(128,128,128,0.25);
            border-radius: 14px;
            padding: 16px 20px;
            margin-bottom: 12px;
            box-shadow: 0 4px 12px rgba(0,0,0,0.06);
        }
        .badge-done { background: rgba(16, 185, 129, 0.15); color: #10b981; border: 1px solid #10b981; padding: 3px 10px; border-radius: 20px; font-size: 12px; font-weight: 800; }
        .badge-prog { background: rgba(245, 158, 11, 0.15); color: #f59e0b; border: 1px solid #f59e0b; padding: 3px 10px; border-radius: 20px; font-size: 12px; font-weight: 800; }
        .badge-wait { background: rgba(107, 114, 128, 0.15); color: #9ca3af; border: 1px solid #9ca3af; padding: 3px 10px; border-radius: 20px; font-size: 12px; font-weight: 800; }
        .badge-issue { background: rgba(239, 68, 68, 0.15); color: #ef4444; border: 1px solid #ef4444; padding: 3px 10px; border-radius: 20px; font-size: 12px; font-weight: 800; }
        </style>
    """, unsafe_allow_html=True)

    c_hd1, c_hd2 = st.columns([8, 2])
    with c_hd1:
        st.markdown("""
            <div class="lux-header">
                <div style="font-size: 12px; color: #FFD700; font-weight: 800; letter-spacing: 2px; text-transform: uppercase;">💎 Enterprise Portfolio & Financial Intelligence</div>
                <h1 style="color: #ffffff !important; font-size: 32px; font-weight: 900; margin: 5px 0 8px 0;">Executive Project & Timeline Command Center</h1>
                <div style="color: #cbd5e1; font-size: 14px;">Monitoring Real-Time Purchase Order (PO), Grafik Budget vs Dana Tersupport, Pending Keuangan & Analisa Project Mandek.</div>
            </div>
        """, unsafe_allow_html=True)
    with c_hd2:
        st.write("")
        if st.button("🔒 Kunci Vault", use_container_width=True):
            st.session_state['proj_dash_unlocked'] = False
            st.rerun()
        if st.button("🔄 Sync Cloud", type="primary", use_container_width=True):
            st.cache_data.clear()
            st.session_state['laporan_db'] = fetch_data_from_gsheets()
            st.rerun()

    db_all = st.session_state.get('laporan_db', [])
    proj_indices = [idx for idx, r in enumerate(db_all) if r.get('kategori') == 'Project']

    if not proj_indices:
        st.info("💡 Belum ada data Project (PO / SVA) yang di-submit. Silakan input melalui menu Form Pelaporan -> Kategori: Project.")
    else:
        total_po_val = 0.0
        total_cost_val = 0.0
        total_support_val = 0.0
        total_routes = 0
        routes_done = 0
        routes_mandek = 0
        all_timeline_rows = []

        for idx in proj_indices:
            r_p = db_all[idx]
            pd_p = r_p.get('proj_data', {})
            if pd_p.get('jenis_project') == 'PO':
                try:
                    v_po = float(str(pd_p.get('nilai_po', '0')).replace('.', '').replace(',', ''))
                except:
                    v_po = 0.0
                total_po_val += v_po
                
                tbl_p = pd_p.get('tabel_pekerjaan', [])
                if not tbl_p:
                    tbl_p = [{
                        "Site Asal": r_p.get('site_name', 'PO Utama'),
                        "Site Tujuan": "-",
                        "Pekerjaan Spesifik": pd_p.get('tipe_pekerjaan', 'Project PO'),
                        "Budget Per Site": v_po,
                        "Dana Tersupport": 0,
                        "Start Date": r_p.get('timestamp', '')[:10],
                        "End Date": r_p.get('timestamp', '')[:10],
                        "Progress (%)": pd_p.get('avg_progress', 0),
                        "Status": r_p.get('status', 'Not Started'),
                        "Kategori Kendala": "Lancar / Normal",
                        "Remark": pd_p.get('note_akhir', '-')
                    }]

                for r_i, row_t in enumerate(tbl_p):
                    asal_s = str(row_t.get('Site Asal', '')).strip()
                    tuj_s = str(row_t.get('Site Tujuan', '')).strip()
                    pek_s = str(row_t.get('Pekerjaan Spesifik', '')).strip()
                    try:
                        c_site = float(str(row_t.get('Budget Per Site', '0')).replace('.', '').replace(',', ''))
                    except:
                        c_site = 0.0
                    try:
                        sup_site = float(str(row_t.get('Dana Tersupport', '0')).replace('.', '').replace(',', ''))
                    except:
                        sup_site = 0.0

                    # Fallback jika budget per site masih 0 pada baris tunggal agar grafik langsung tampil proporsional
                    if c_site == 0 and len(tbl_p) == 1 and v_po > 0:
                        c_site = v_po

                    if not asal_s or asal_s == 'nan':
                        asal_s = f"Site-{r_i+1}"
                    if not pek_s or pek_s == 'nan':
                        pek_s = str(pd_p.get('tipe_pekerjaan', 'Pekerjaan PO'))

                    total_routes += 1
                    total_cost_val += c_site
                    total_support_val += sup_site
                    pending_fin_site = max(c_site - sup_site, 0.0)
                    
                    st_site = str(row_t.get('Status', 'Not Started') or 'Not Started')
                    kendala_site = str(row_t.get('Kategori Kendala', 'Lancar / Normal') or 'Lancar / Normal')
                    if kendala_site == 'nan':
                        kendala_site = 'Lancar / Normal'
                        
                    try:
                        pr_site = float(row_t.get('Progress (%)', 0) or 0)
                    except:
                        pr_site = 0.0

                    # Auto-deteksi jika dana belum tersupport 100% dan status belum Done
                    if st_site != 'Done' and pr_site < 100 and pending_fin_site > 0 and kendala_site == 'Lancar / Normal':
                        kendala_site = "Mandek - Pending Keuangan (Dana Belum Cair)"
                        
                    if st_site == 'Done' or pr_site >= 100:
                        routes_done += 1
                    elif st_site == 'Pending Issue' or "Mandek" in kendala_site or pending_fin_site > 0:
                        routes_mandek += 1
                        
                    rute_label = f"{asal_s} ➔ {tuj_s}" if tuj_s and tuj_s not in ['-', 'nan', ''] else asal_s
                    all_timeline_rows.append({
                        "No PO": str(pd_p.get('no_po', '-')),
                        "Mitra": str(pd_p.get('mitra', '-')),
                        "Program": str(pd_p.get('tipe_pekerjaan', '-')),
                        "Site Asal": asal_s,
                        "Site Tujuan": tuj_s if tuj_s not in ['nan', ''] else '-',
                        "Rute Site": f"[{pd_p.get('no_po', '-')}] {rute_label}",
                        "Pekerjaan Spesifik": pek_s,
                        "Total Budget (Rp)": c_site,
                        "Dana Tersupport (Rp)": sup_site,
                        "Pending Keuangan (Rp)": pending_fin_site,
                        "Start Date": str(row_t.get('Start Date', '') or r_p.get('timestamp', '')[:10]),
                        "End Date": str(row_t.get('End Date', '') or r_p.get('timestamp', '')[:10]),
                        "Progress (%)": pr_site,
                        "Status": st_site,
                        "Kategori Kendala": kendala_site,
                        "Remark": str(row_t.get('Remark', '') or pd_p.get('note_akhir', '-') or '-'),
                        "PIC Tim": r_p.get('teknisi', '-')
                    })

        total_margin_val = total_po_val - total_cost_val
        total_pending_fin = max(total_cost_val - total_support_val, 0.0)
        support_ratio = (total_support_val / total_cost_val * 100) if total_cost_val > 0 else 0.0
        margin_pct = (total_margin_val / total_po_val * 100) if total_po_val > 0 else 0.0
        route_completion_pct = (routes_done / total_routes * 100) if total_routes > 0 else 0.0

        # 6 LUXURY KPI CARDS
        k1, k2, k3 = st.columns(3)
        with k1:
            st.markdown(f"""
                <div class="lux-kpi" style="border-left-color: #3b82f6;">
                    <div class="lux-kpi-title">💰 Total Nilai Kontrak PO</div>
                    <div class="lux-kpi-val">Rp {total_po_val:,.0f}</div>
                    <div class="lux-kpi-sub" style="color: #60a5fa;">{len(proj_indices)} Kontrak Project Terdaftar</div>
                </div>
            """, unsafe_allow_html=True)
        with k2:
            st.markdown(f"""
                <div class="lux-kpi" style="border-left-color: #f59e0b;">
                    <div class="lux-kpi-title">💸 Total Budget / RAB Rute</div>
                    <div class="lux-kpi-val">Rp {total_cost_val:,.0f}</div>
                    <div class="lux-kpi-sub" style="color: #fbbf24;">Alokasi untuk {total_routes} Rute Site</div>
                </div>
            """, unsafe_allow_html=True)
        with k3:
            warna_m = "#10b981" if total_margin_val >= 0 else "#ef4444"
            st.markdown(f"""
                <div class="lux-kpi" style="border-left-color: {warna_m};">
                    <div class="lux-kpi-title">📈 Estimasi Net Margin PO</div>
                    <div class="lux-kpi-val" style="color: {warna_m};">Rp {total_margin_val:,.0f}</div>
                    <div class="lux-kpi-sub" style="color: {warna_m};">Profitability: {margin_pct:.1f}% dari Nilai PO</div>
                </div>
            """, unsafe_allow_html=True)

        k4, k5, k6 = st.columns(3)
        with k4:
            st.markdown(f"""
                <div class="lux-kpi" style="border-left-color: #10b981;">
                    <div class="lux-kpi-title">✅ Dana Sudah Tersupport (Cair)</div>
                    <div class="lux-kpi-val" style="color: #10b981;">Rp {total_support_val:,.0f}</div>
                    <div class="lux-kpi-sub" style="color: #34d399;">{support_ratio:.1f}% dari Total Budget Rute</div>
                </div>
            """, unsafe_allow_html=True)
        with k5:
            st.markdown(f"""
                <div class="lux-kpi" style="border-left-color: #ef4444;">
                    <div class="lux-kpi-title">🚨 Pending Keuangan (Belum Cair)</div>
                    <div class="lux-kpi-val" style="color: #ef4444;">Rp {total_pending_fin:,.0f}</div>
                    <div class="lux-kpi-sub" style="color: #f87171;">Dana tertahan / belum disupport</div>
                </div>
            """, unsafe_allow_html=True)
        with k6:
            st.markdown(f"""
                <div class="lux-kpi" style="border-left-color: #a855f7;">
                    <div class="lux-kpi-title">🎯 Clearance & Bottleneck</div>
                    <div class="lux-kpi-val">{routes_done}/{total_routes} Clear | <span style="color:#ef4444;">{routes_mandek} Tertahan</span></div>
                    <div class="lux-kpi-sub" style="color: #c084fc;">{route_completion_pct:.1f}% Rute Site Selesai</div>
                </div>
            """, unsafe_allow_html=True)

        tab_tl, tab_po_list, tab_export = st.tabs([
            "📊 Grafik Budget vs Support, Timeline & Analisa Mandek",
            "📂 Klik Nomor PO / SVA (Update Dana Support, Timeline & Foto)",
            "📥 Executive Financial Export"
        ])

        with tab_tl:
            if not all_timeline_rows:
                st.info("Belum ada rincian PO yang dimasukkan.")
            else:
                df_tl = pd.DataFrame(all_timeline_rows)

                # =====================================================================
                # 1. GRAFIK KOMPARASI TOTAL BUDGET VS DANA TERSUPPORT VS PENDING KEUANGAN
                # =====================================================================
                st.markdown("### 📊 1. Grafik Komparasi Budget vs Dana Tersupport & Pending Keuangan")
                cg_1, cg_2 = st.columns([1.4, 1])
                with cg_1:
                    with st.container(border=True):
                        st.markdown("**💰 Komparasi Total Budget vs Sudah Tersupport vs Pending Keuangan (Per Rute Site)**")
                        df_chart_fin = df_tl.set_index("Rute Site")[["Total Budget (Rp)", "Dana Tersupport (Rp)", "Pending Keuangan (Rp)"]]
                        st.bar_chart(df_chart_fin, color=["#3b82f6", "#10b981", "#ef4444"])
                        st.caption("Biru = Total Budget Site | Hijau = Dana Sudah Tersupport | Merah = Pending Keuangan (Belum Cair)")

                with cg_2:
                    with st.container(border=True):
                        st.markdown("**🛑 Analisa Bottleneck: Kenapa Project Mandek / Tertahan?**")
                        df_kendala = df_tl["Kategori Kendala"].value_counts().reset_index()
                        df_kendala.columns = ["Alasan / Kendala", "Jumlah Site"]
                        st.bar_chart(df_kendala.set_index("Alasan / Kendala"), color="#f59e0b")
                        
                        df_stuck = df_tl[(df_tl["Pending Keuangan (Rp)"] > 0) | (df_tl["Status"] != "Done")]
                        if not df_stuck.empty:
                            st.markdown("**🚨 Daftar Rute Site Tertahan / Butuh Support Keuangan:**")
                            for _, r_st in df_stuck.head(5).iterrows():
                                st.markdown(
                                    f"<div style='font-size:12px; padding:6px 10px; margin-bottom:4px; border-left:4px solid #ef4444; background:rgba(239,68,68,0.08); border-radius:6px;'>"
                                    f"<b>{r_st['Rute Site']}</b> — Pending Keuangan: <b style='color:#ef4444;'>Rp {r_st['Pending Keuangan (Rp)']:,.0f}</b><br>"
                                    f"📌 <i>Kendala: {r_st['Kategori Kendala']} | Remark: {r_st['Remark']}</i></div>",
                                    unsafe_allow_html=True
                                )

                st.markdown("<hr style='margin: 20px 0;'>", unsafe_allow_html=True)

                # =====================================================================
                # 2. VISUAL TIMELINE GANTT & ABSENSI PENGERJAAN SITE
                # =====================================================================
                st.markdown("### 🗓️ 2. Visual Timeline Chart & Absensi Progress Per-Site")
                cf1, cf2, cf3 = st.columns(3)
                with cf1:
                    flt_po = st.selectbox("🔍 Filter Nomor PO:", ["Semua PO"] + sorted(df_tl["No PO"].unique().tolist()))
                with cf2:
                    flt_stat = st.selectbox("📊 Filter Status Site:", ["Semua Status", "Done", "On Progress", "Not Started", "Pending Issue"])
                with cf3:
                    flt_cari = st.text_input("🔎 Cari Site Asal / Tujuan / Kendala:", placeholder="Ketik Site ID atau Kendala...")

                df_tl_view = df_tl.copy()
                if flt_po != "Semua PO":
                    df_tl_view = df_tl_view[df_tl_view["No PO"] == flt_po]
                if flt_stat != "Semua Status":
                    df_tl_view = df_tl_view[df_tl_view["Status"] == flt_stat]
                if flt_cari.strip():
                    q_tl = flt_cari.strip().lower()
                    df_tl_view = df_tl_view[
                        df_tl_view["Rute Site"].str.lower().str.contains(q_tl) |
                        df_tl_view["Pekerjaan Spesifik"].str.lower().str.contains(q_tl) |
                        df_tl_view["Kategori Kendala"].str.lower().str.contains(q_tl)
                    ]

                with st.container(border=True):
                    st.markdown("**📈 Grafik Perbandingan Progress Fisik (%) vs Persentase Support Keuangan (%)**")
                    df_prog_cmp = df_tl_view.copy()
                    df_prog_cmp["Support Keuangan (%)"] = df_prog_cmp.apply(
                        lambda x: min((x["Dana Tersupport (Rp)"] / x["Total Budget (Rp)"] * 100), 100.0) if x["Total Budget (Rp)"] > 0 else 0.0,
                        axis=1
                    ).round(1)
                    st.bar_chart(df_prog_cmp.set_index("Rute Site")[["Progress (%)", "Support Keuangan (%)"]], color=["#a855f7", "#10b981"])
                    st.caption("Ungu = Progress Fisik Lapangan (%) | Hijau = Persentase Dana Sudah Disupport Keuangan (%)")

                for _, row_item in df_tl_view.iterrows():
                    prog_val = float(row_item["Progress (%)"])
                    b_tot = float(row_item["Total Budget (Rp)"])
                    b_sup = float(row_item["Dana Tersupport (Rp)"])
                    b_pend = float(row_item["Pending Keuangan (Rp)"])
                    sup_pct = min((b_sup / b_tot * 100), 100.0) if b_tot > 0 else 0.0

                    stat_val = row_item["Status"]
                    if stat_val == "Done" or prog_val >= 100:
                        b_cls = "badge-done"
                        bar_col = "#10b981"
                    elif stat_val == "On Progress":
                        b_cls = "badge-prog"
                        bar_col = "#f59e0b"
                    elif stat_val == "Pending Issue" or "Mandek" in str(row_item["Kategori Kendala"]):
                        b_cls = "badge-issue"
                        bar_col = "#ef4444"
                    else:
                        b_cls = "badge-wait"
                        bar_col = "#6b7280"

                    warna_kendala = "#ef4444" if ("Mandek" in str(row_item["Kategori Kendala"]) or b_pend > 0) and stat_val != "Done" else "#10b981"

                    st.markdown(f"""
                        <div class="tl-card" style="border-left: 6px solid {bar_col};">
                            <div style="display:flex; justify-content:space-between; align-items:center; flex-wrap:wrap; gap:8px; margin-bottom:8px;">
                                <div>
                                    <span style="font-size:12px; font-weight:800; color:#3b82f6; background:rgba(59,130,246,0.12); padding:3px 8px; border-radius:6px; margin-right:8px;">PO: {row_item['No PO']}</span>
                                    <span style="font-size:17px; font-weight:900;">📍 {row_item['Site Asal']} ➡️ {row_item['Site Tujuan']}</span>
                                </div>
                                <div style="display:flex; gap:8px; align-items:center;">
                                    <span style="font-size:12px; font-weight:800; color:{warna_kendala}; background:rgba(239,68,68,0.1); padding:3px 10px; border-radius:20px; border:1px solid {warna_kendala};">🚦 {row_item['Kategori Kendala']}</span>
                                    <span class="{b_cls}">{stat_val} ({prog_val:.0f}%)</span>
                                </div>
                            </div>
                            <div style="font-size:13px; opacity:0.92; margin-bottom:10px; display:flex; gap:16px; flex-wrap:wrap;">
                                <span>🛠️ <b>Pekerjaan:</b> {row_item['Pekerjaan Spesifik']}</span>
                                <span>💰 <b>Total Budget:</b> Rp {b_tot:,.0f}</span>
                                <span style="color:#10b981; font-weight:700;">✅ <b>Tersupport:</b> Rp {b_sup:,.0f} ({sup_pct:.0f}%)</span>
                                <span style="color:#ef4444; font-weight:700;">🚨 <b>Pending Keuangan:</b> Rp {b_pend:,.0f}</span>
                                <span>🗓️ <b>Timeline:</b> {row_item['Start Date']} ➔ {row_item['End Date']}</span>
                            </div>
                            <div style="font-size:11px; font-weight:700; margin-bottom:2px;">🛠️ Progress Fisik Lapangan ({prog_val:.0f}%)</div>
                            <div style="width:100%; background:rgba(128,128,128,0.2); height:8px; border-radius:10px; overflow:hidden; margin-bottom:6px;">
                                <div style="width:{min(max(prog_val, 0), 100)}%; background:{bar_col}; height:100%; border-radius:10px;"></div>
                            </div>
                            <div style="font-size:11px; font-weight:700; margin-bottom:2px;">💳 Pencairan / Support Keuangan ({sup_pct:.0f}%)</div>
                            <div style="width:100%; background:rgba(128,128,128,0.2); height:8px; border-radius:10px; overflow:hidden; margin-bottom:6px;">
                                <div style="width:{min(max(sup_pct, 0), 100)}%; background:#10b981; height:100%; border-radius:10px;"></div>
                            </div>
                            <div style="font-size:12px; opacity:0.8;">📝 <b>Keterangan / Kenapa Mandek:</b> {row_item['Remark']} | 👷 <b>PIC:</b> {row_item['PIC Tim']}</div>
                        </div>
                    """, unsafe_allow_html=True)

        with tab_po_list:
            st.markdown("### 📂 Pilih Nomor PO / SVA untuk Update Dana Support, Timeline & Foto")
            opsi_proj_select = {
                idx: f"[{db_all[idx].get('proj_data',{}).get('jenis_project','PO')}] {db_all[idx].get('site_name','Project')} — Mitra: {db_all[idx].get('proj_data',{}).get('mitra','-')} ({db_all[idx].get('status','-')})"
                for idx in proj_indices
            }
            sel_proj_idx = st.selectbox("📌 Klik / Pilih Nomor PO atau Tiket SVA:", options=proj_indices, format_func=lambda x: opsi_proj_select[x])

            if sel_proj_idx is not None:
                r = db_all[sel_proj_idx]
                pd_data = r.get('proj_data', {})

                st.markdown("<hr style='margin: 15px 0;'>", unsafe_allow_html=True)
                if pd_data.get('jenis_project') == 'PO':
                    val_po = float(str(pd_data.get('nilai_po', '0')).replace('.', '').replace(',', '')) if pd_data.get('nilai_po') else 0.0
                    t_cost = float(pd_data.get('total_cost', 0.0) or 0.0)
                    t_sup = float(pd_data.get('total_support', 0.0) or 0.0)
                    t_margin = float(pd_data.get('margin', 0.0) or 0.0)

                    c_m1, c_m2, c_m3, c_m4, c_m5 = st.columns(5)
                    c_m1.metric("💰 Nilai PO", f"Rp {val_po:,.0f}")
                    c_m2.metric("💸 Total Budget Site", f"Rp {t_cost:,.0f}")
                    c_m3.metric("✅ Dana Tersupport", f"Rp {t_sup:,.0f}")
                    c_m4.metric("🚨 Pending Keuangan", f"Rp {max(t_cost - t_sup, 0.0):,.0f}")
                    c_m5.metric("📈 Net Margin PO", f"Rp {t_margin:,.0f}")

                    st.markdown("#### 🛠️ Live Editor: Update Rute Site, Dana Tersupport, Alasan Mandek & Progress (%)")
                    st.info("💡 Klik dua kali pada sel tabel di bawah ini untuk mengisi **Dana Tersupport (Rp)** dari keuangan, memilih **Kategori Kendala (Kenapa Mandek)**, atau mengupdate tanggal & progress.")

                    df_edit_po = pd.DataFrame(pd_data.get('tabel_pekerjaan', []))
                    cols_standar = ["Site Asal", "Site Tujuan", "Pekerjaan Spesifik", "Budget Per Site", "Dana Tersupport", "Start Date", "End Date", "Progress (%)", "Status", "Kategori Kendala", "Remark"]
                    if df_edit_po.empty:
                        df_edit_po = pd.DataFrame(columns=cols_standar)
                    else:
                        for c_st in cols_standar:
                            if c_st not in df_edit_po.columns:
                                if c_st in ["Budget Per Site", "Dana Tersupport", "Progress (%)"]:
                                    df_edit_po[c_st] = 0
                                elif c_st == "Kategori Kendala":
                                    df_edit_po[c_st] = "Lancar / Normal"
                                else:
                                    df_edit_po[c_st] = ""
                        df_edit_po = df_edit_po[cols_standar]
                        df_edit_po["Budget Per Site"] = pd.to_numeric(df_edit_po["Budget Per Site"], errors='coerce').fillna(0)
                        df_edit_po["Dana Tersupport"] = pd.to_numeric(df_edit_po["Dana Tersupport"], errors='coerce').fillna(0)
                        df_edit_po["Progress (%)"] = pd.to_numeric(df_edit_po["Progress (%)"], errors='coerce').fillna(0)

                    updated_po_table = st.data_editor(
                        df_edit_po,
                        num_rows="dynamic",
                        use_container_width=True,
                        column_config={
                            "Site Asal": st.column_config.TextColumn("Site Asal (Dari)"),
                            "Site Tujuan": st.column_config.TextColumn("Site Tujuan (Ke)"),
                            "Pekerjaan Spesifik": st.column_config.TextColumn("Pekerjaan"),
                            "Budget Per Site": st.column_config.NumberColumn("Total Budget (Rp)", min_value=0, default=0, step=100000),
                            "Dana Tersupport": st.column_config.NumberColumn("Dana Tersupport (Rp)", min_value=0, default=0, step=100000),
                            "Start Date": st.column_config.TextColumn("Mulai (YYYY-MM-DD)"),
                            "End Date": st.column_config.TextColumn("Selesai (YYYY-MM-DD)"),
                            "Progress (%)": st.column_config.NumberColumn("Progress (%)", min_value=0, max_value=100, default=0),
                            "Status": st.column_config.SelectboxColumn("Status", options=["Not Started", "On Progress", "Done", "Pending Issue"]),
                            "Kategori Kendala": st.column_config.SelectboxColumn("Kenapa Mandek / Kendala", options=[
                                "Lancar / Normal",
                                "Mandek - Pending Keuangan (Dana Belum Cair)",
                                "Mandek - Pending Material / Sparepart",
                                "Mandek - Izin Site / Community Issue",
                                "Mandek - Kendala Teknis / Cuaca"
                            ]),
                            "Remark": st.column_config.TextColumn("Detail Kendala / Catatan")
                        },
                        key=f"lux_edit_po_{sel_proj_idx}"
                    )

                    c_up1, c_up2 = st.columns(2)
                    with c_up1:
                        stat_opts_p = ["Not Started", "On Progress", "Done", "Pending"]
                        cur_st_p = r.get('status', 'On Progress')
                        new_status_proj = st.selectbox("Status Keseluruhan PO:", stat_opts_p, index=stat_opts_p.index(cur_st_p) if cur_st_p in stat_opts_p else 1, key=f"lux_stat_{sel_proj_idx}")
                        new_nilai_po = st.text_input("Revisi Nilai PO Total (Rp):", value=str(pd_data.get('nilai_po', '0')), key=f"lux_valpo_{sel_proj_idx}")
                    with c_up2:
                        new_note_proj = st.text_area("Update Catatan / Remark PO:", value=str(pd_data.get('note_akhir', '')), key=f"lux_note_{sel_proj_idx}")

                    st.markdown("#### 📸 Upload Dokumentasi Foto Susulan (Saat Pekerjaan Dimulai / Selesai)")
                    cu_f1, cu_f2, cu_f3 = st.columns(3)
                    with cu_f1:
                        up_p_awal = st.file_uploader("➕ Tambah Foto Before", accept_multiple_files=True, key=f"lux_f_awal_{sel_proj_idx}")
                    with cu_f2:
                        up_p_pros = st.file_uploader("➕ Tambah Foto On Progress", accept_multiple_files=True, key=f"lux_f_pros_{sel_proj_idx}")
                    with cu_f3:
                        up_p_akh = st.file_uploader("➕ Tambah Foto After / BAST", accept_multiple_files=True, key=f"lux_f_akh_{sel_proj_idx}")

                    if st.button("💾 SIMPAN UPDATE PO & KALKULASI ULANG MARGIN", type="primary", use_container_width=True, key=f"btn_save_lux_{sel_proj_idx}"):
                        with st.spinner("⏳ Menyimpan perubahan tabel, dana support & mengunggah foto susulan ke Cloud..."):
                            r['proj_data']['tabel_pekerjaan'] = updated_po_table.astype(str).to_dict('records')
                            r['proj_data']['nilai_po'] = new_nilai_po
                            r['proj_data']['note_akhir'] = new_note_proj
                            try:
                                v_po_new = float(str(new_nilai_po).replace('.', '').replace(',', ''))
                            except:
                                v_po_new = 0.0
                            t_cost_new = pd.to_numeric(updated_po_table["Budget Per Site"], errors='coerce').fillna(0).sum() if not updated_po_table.empty else 0.0
                            t_sup_new = pd.to_numeric(updated_po_table["Dana Tersupport"], errors='coerce').fillna(0).sum() if not updated_po_table.empty else 0.0
                            avg_p_new = pd.to_numeric(updated_po_table["Progress (%)"], errors='coerce').fillna(0).mean() if not updated_po_table.empty else 0.0

                            r['proj_data']['total_cost'] = float(t_cost_new)
                            r['proj_data']['total_support'] = float(t_sup_new)
                            r['proj_data']['pending_finance'] = float(max(t_cost_new - t_sup_new, 0.0))
                            r['proj_data']['margin'] = float(v_po_new - t_cost_new)
                            r['proj_data']['avg_progress'] = float(avg_p_new)
                            r['status'] = new_status_proj

                            if up_p_awal:
                                r['proj_data']['url_awal'] = r['proj_data'].get('url_awal', []) + upload_multiple_images(up_p_awal, "proj_po")
                            if up_p_pros:
                                r['proj_data']['url_proses'] = r['proj_data'].get('url_proses', []) + upload_multiple_images(up_p_pros, "proj_po")
                            if up_p_akh:
                                r['proj_data']['url_akhir'] = r['proj_data'].get('url_akhir', []) + upload_multiple_images(up_p_akh, "proj_po")

                            sheet = connect_gsheets(r.get('_sheet', 'Project'))
                            if sheet:
                                sheet.update_cell(r.get('_row', sel_proj_idx + 2), 5, new_status_proj)
                                sheet.update_cell(r.get('_row', sel_proj_idx + 2), 6, json.dumps(r, default=str))
                                st.cache_data.clear()
                        st.success("✅ BERHASIL! Data Rute, Dana Tersupport, Timeline, Margin & Foto PO telah diperbarui.")
                        st.rerun()

                    st.markdown("#### 🖼️ Galeri Dokumentasi PO Saat Ini")
                    gt1, gt2, gt3 = st.tabs(["📸 Foto Kondisi Awal (Before)", "📸 Foto Pengerjaan (On Progress)", "📸 Foto Selesai (After / BAST)"])
                    with gt1: tampilkan_grid_foto(pd_data.get('url_awal', []), "Dokumentasi Kondisi Awal")
                    with gt2: tampilkan_grid_foto(pd_data.get('url_proses', []), "Dokumentasi On Progress")
                    with gt3: tampilkan_grid_foto(pd_data.get('url_akhir', []), "Dokumentasi Akhir / BAST")

                else:
                    st.info(f"**📋 Detail SVA:** Site: {pd_data.get('site_sva','-')} | No Tiket: {pd_data.get('no_tiket','-')} | Tanggal: {pd_data.get('tgl_sva','-')} | Temuan: {pd_data.get('temuan_sva','-')}")
                    gt_s1, gt_s2 = st.tabs(["📸 View Site & Perangkat", "📸 Temuan / Defect SVA"])
                    with gt_s1: tampilkan_grid_foto(pd_data.get('url_sva1', []), "Foto View Site & Perangkat")
                    with gt_s2: tampilkan_grid_foto(pd_data.get('url_sva2', []), "Foto Temuan Defect")

        with tab_export:
            st.markdown("### 📥 Unduh Rekapitulasi Finansial & Timeline Project (Excel)")
            if all_timeline_rows:
                df_exp_proj = pd.DataFrame(all_timeline_rows)
                st.dataframe(df_exp_proj, use_container_width=True)
                try:
                    buf_p = BytesIO()
                    with pd.ExcelWriter(buf_p, engine='xlsxwriter') as writer:
                        df_exp_proj.to_excel(writer, index=False, sheet_name='Executive_Project')
                    st.download_button("💎 Download Executive Project Report (.xlsx)", data=buf_p.getvalue(), file_name=f"Executive_Project_Report_{datetime.date.today()}.xlsx", mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet", type="primary", use_container_width=True)
                except Exception:
                    csv_p = df_exp_proj.to_csv(index=False).encode('utf-8')
                    st.download_button("💎 Download Executive Project Report (.csv)", data=csv_p, file_name=f"Executive_Project_Report_{datetime.date.today()}.csv", mime="text/csv", type="primary", use_container_width=True)
            else:
                st.info("Belum ada rincian tabel rute PO untuk diekspor.")# -------------------------------------------------------------------------
# FOOTER HAK CIPTA OKTA PRADIKA
# -------------------------------------------------------------------------
st.markdown("""
    <div class="footer-okta">
        System Application & Database Management<br>
        <span>Create Data By Okta Pradika</span> © 2026
    </div>
""", unsafe_allow_html=True)
