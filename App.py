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
# FUNGSI PEMBUAT DOKUMEN (PDF & PPTX) - GLOBAL SCOPE
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
    st.markdown("<h1 style='text-align: center; font-size: 40px; margin-top: 5px;'>RPO NOP PLK & PBU</h1>", unsafe_allow_html=True)
    st.markdown("<p style='text-align: center; font-size: 18px; margin-bottom: 45px; letter-spacing: 1px; opacity: 0.85;'>Report Online & Dasboard Monitoring Activity Taks Force</p>", unsafe_allow_html=True)
    col_v, col_space, col_a = st.columns([4, 0.6, 4])
    
    with col_v:
        st.markdown("""
            <div class='login-box' style='padding-bottom: 20px;'>
                <div style='font-size: 55px; margin-bottom: 15px;'>👷‍♂️</div>
                <div class='login-title'>Mode Tim Lapangan</div>
                <div class='login-desc'>Pilih area operasional Anda agar langsung diarahkan ke form pendataan yang tepat.</div>
            </div>
        """, unsafe_allow_html=True)
        st.write("")
        c_btn1, c_btn2 = st.columns(2)
        with c_btn1:
            if st.button("🔴 MASUK PANGKALAN BUN", type="primary", use_container_width=True):
                st.session_state['role'] = 'Lapangan'
                st.session_state['area_lapangan'] = 'Pangkalan Bun'
                st.session_state['dashboard_area'] = None
                st.rerun()
        with c_btn2:
            if st.button("🔵 MASUK PALANGKARAYA", type="secondary", use_container_width=True):
                st.session_state['role'] = 'Lapangan'
                st.session_state['area_lapangan'] = 'Palangkaraya'
                st.session_state['dashboard_area'] = None
                st.rerun()
            
    with col_a:
        st.markdown("""
            <div class='login-box'>
                <div style='font-size: 55px; margin-bottom: 15px;'>🔐</div>
                <div class='login-title'>Mode Admin</div>
                <div class='login-desc'>Akses khusus operasional untuk revisi data lapangan, perbaikan dashboard, input form PM utuh, dan analisa grafik Power.</div>
            </div>
        """, unsafe_allow_html=True)
        st.write("")
        admin_pass = st.text_input("Kata Sandi Admin:", type="password", placeholder="Masukkan Sandi...", key="pwd_login")
        if st.button("🔓 LOGIN ADMIN", type="primary", use_container_width=True):
            if admin_pass == "KUT2027":
                st.session_state['role'] = 'Admin'
                st.session_state['dashboard_area'] = None
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
st.markdown("""
    <style>
        [data-testid="stSidebar"] { display: none !important; }
        [data-testid="collapsedControl"] { display: none !important; }
    </style>
""", unsafe_allow_html=True)

if 'active_menu' not in st.session_state:
    st.session_state['active_menu'] = None

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
    c_out1, c_out2, c_out3 = st.columns([3, 4, 3])
    with c_out2:
        if st.button("🚪 KELUAR AKUN (LOG OUT)", use_container_width=True, key="btn_logout_main"):
            st.session_state['role'] = None
            st.session_state['dashboard_area'] = None
            st.session_state['active_menu'] = None
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
    
    render_header_logo()
    
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
        site_category = st.radio("Kategori Site (Wajib Pilih):", ["SPS", "Site Reguler"], horizontal=True)
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

        else:
            # -------------------------------------------------------------------------
            # BLOK A.2: FORM KHUSUS SITE REGULER BARU (PALANGKARAYA)
            # -------------------------------------------------------------------------
            tabs_reg = st.tabs(["📌 1. Info & Problem", "⚡ 2. Action & Parameter", "📤 3. Submit"])
            
                            else:
                                st.error("⚠️ Gagal menyambung ke Server Database.")

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
    
    render_header_logo()
    st.markdown("<h1 style='text-align: center;'>📊 Executive Dashboard & Analytics Report</h1>", unsafe_allow_html=True)
    st.markdown("<p style='text-align: center; opacity: 0.8; margin-bottom: 25px;'>Sistem Monitoring Terpadu Hasil Preventive Maintenance & Site Visit</p>", unsafe_allow_html=True)
    
    db = st.session_state['laporan_db']

    if not db:
        st.warning("⚠ Belum ada data di Server / Koneksi Sedang Proses.")
    else:
        idx_plk = [idx for idx, r in enumerate(db) if r.get('kategori') != 'Pangkalan Bun']
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
                        <div class='login-desc'>Daftar Site Preventive Maintenance (SPS & Site Reguler), Dokumentasi Foto Modul Surya, Baterai, Datalog & Berita Acara PDF.</div>
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
            if st.button("⬅️ Ganti Area Operasional", use_container_width=True):
                st.session_state['dashboard_area'] = None
                st.rerun()
        with c_nav2:
            warna_area = "#007BFF" if aktif_area == "Palangkaraya" else "#FF3333"
            st.markdown(
                f"<div style='padding: 8px 18px; border-radius: 10px; border-left: 6px solid {warna_area}; background-color: var(--secondary-background-color); font-size: 18px; font-weight: 800;'>📍 Area Aktif: NOP {aktif_area.upper()}</div>",
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
        jml_normal = sum(1 for x in export_db if x.get('status') == 'Normal')
        jml_issue = len(export_db) - jml_normal

        cm1, cm2, cm3 = st.columns(3)
        cm1.metric(f"Total Site ({aktif_area})", f"{len(export_db)} Site")
        cm2.metric("🟢 Status Normal", f"{jml_normal} Site")
        cm3.metric("🟡/🔴 Memiliki Issue / Progress", f"{jml_issue} Site")

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
                        "Reguler - Note/Remark": cond_d.get('note_action', '-')
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
                            if col_name in ["Action / Pekerjaan", "Remark", "Possibility PSB", "Estimasi Jarak tower ke jaringan PLN", "Type Merk Panel", "Type Battery", "SPS Rusak (Detail)", "Baterai Rusak (Detail)", "Sparepart Diganti", "Catatan/Temuan", "Reguler - Problem Alarm", "Reguler - Note/Remark"]:
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
                            pptx_data = build_pptx([x for x in export_db if x.get('kategori') != 'Pangkalan Bun' and x.get('kategori') != 'Site Reguler'])
                            st.download_button(label="📽 Download Presentasi Report (PPTX)", data=pptx_data, file_name=f"Report_PM_KUT_{datetime.date.today()}.pptx", mime="application/vnd.openxmlformats-officedocument.presentationml.presentation", type="primary", use_container_width=True)
                    else:
                        st.error("Library `python-pptx` belum ter-install di server.")
                else:
                    st.info("💡 Rekap lengkap parameter FME Pangkalan Bun dapat diunduh melalui tombol Excel di samping.")

        st.markdown("### 📋 Daftar Laporan Site Terinspeksi")

        # ---------------------------------------------------------------------
        # TAMPILAN KARTU EXPANDER (KOMPATIBEL 100% DENGAN DATA LAMA)
        # ---------------------------------------------------------------------
        for i in reversed(filtered_indices):
            r = db[i]
            site_id = r.get('site_name', 'Unknown')
            kategori_site = r.get('kategori', 'SPS')
            stat = r.get('status', '')
            icon = "🟢" if stat == "Normal" else "🟡" if stat in ["Minor Issue", "Update Progress"] else "🔴"
            
            # CEK APAKAH INI DATA REGULER BARU ATAU LAMA
            is_new_reguler = kategori_site == "Site Reguler" and ("problem_alarm" in r or "cond_data" in r)
            
            # FORMAT WHATSAPP DINAMIS
            if kategori_site == "Pangkalan Bun":
                wa_text = generate_wa_pbu(r)
            elif is_new_reguler:
                wa_text = f"*BERITA ACARA SITE REGULER* ⚡\n📍 *Site:* {site_id} ({r.get('nop', '-')})\n📅 *Tanggal:* {r.get('timestamp', '-')}\n👷 *FME:* {r.get('teknisi', '-')} ({r.get('role_fme', '-')})\n📊 *Status:* {r.get('status', '-')}\n\n*🚨 PROBLEM ALARM:*\n{r.get('problem_alarm', '-')}\n\n*🔧 ACTION / PERBAIKAN:*\n{', '.join(r.get('action_reguler', [])) if isinstance(r.get('action_reguler'), list) else r.get('action_reguler', '-')}\n"
                cd = r.get('cond_data', {})
                if cd:
                    wa_text += f"\n*⚡ PARAMETER KELISTRIKAN:*\n- Load AC: {cd.get('load_ac_rst', '-')}\n- Volt AC: {cd.get('volt_ac_rst', '-')}\n- Load DC: {cd.get('load_dc', '-')}\n- Set BCL: {cd.get('set_bcl', '-')}\n- Set Cap: {cd.get('set_cap', '-')}\n- Modul Rusak: {cd.get('mod_rusak', '-')}\n"
                    if cd.get('note_action'):
                        wa_text += f"\n*📝 NOTE/REMARK:*\n{cd.get('note_action')}\n"
            else:
                wa_text = f"*BERITA ACARA PREVENTIVE MAINTENANCE* ⚡\n📍 *Site:* {site_id} ({kategori_site} - {r.get('nop', '-')})\n📅 *Tanggal:* {r.get('timestamp', '-')}\n👷 *Pelaksana:* {r.get('teknisi', '-')}\n📊 *Status:* {r.get('status', '-')}\n\n*RINCIAN TINDAKAN:*\n{r.get('action', '-')}\n\n*POWER & LOAD:*\n- PLN: {r.get('pln_status', '-')}\n- Rectifier: {r.get('rect_brand', '-')} ({r.get('rect_out_v', '-')}V)\n- Load BTS: {r.get('total_load', '-')} A\n\n*SPAREPART:*\n{r.get('sparepart', '-')}"
            
            wa_url = f"https://wa.me/?text={urllib.parse.quote(wa_text)}"

            with st.expander(f"{icon}  |  {site_id}  |  {kategori_site}  |  {r.get('timestamp', '')}  |  Status: {stat}"):
                
                if kategori_site != "Pangkalan Bun" and not is_new_reguler:
                    if st.checkbox("📄 Buat Berita Acara (PDF Resmi)", key=f"prep_pdf_{i}"):
                        with st.spinner("⏳ Rendering Dokumen PDF Resolusi Tinggi..."):
                            try:
                                pdf_bytes = build_pdf(r)
                                st.download_button(label="📥 Download PDF Berita Acara", data=pdf_bytes, file_name=f"Berita_Acara_{site_id}.pdf", mime="application/pdf", key=f"dl_pdf_{i}", type="primary")
                            except Exception as e:
                                st.error(f"Terjadi kesalahan saat menyusun PDF: {e}")
                
                st.markdown("<hr style='border: 1px solid var(--primary-color); margin: 12px 0;'>", unsafe_allow_html=True)
                st.link_button("📱 Share Rangkuman Laporan ke WhatsApp", wa_url, use_container_width=True)
                
                # =========================================================================
                # TAMPILAN PANGKALAN BUN
                # =========================================================================
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

                # =========================================================================
                # TAMPILAN SITE REGULER (VERSI BARU)
                # =========================================================================
                elif is_new_reguler:
                wa_text = f"*BERITA ACARA SITE REGULER* ⚡\n📍 *Site:* {site_id} ({r.get('nop', '-')})\n📅 *Tanggal:* {r.get('timestamp', '-')}\n👷 *FME:* {r.get('teknisi', '-')} ({r.get('role_fme', '-')})\n📊 *Status:* {r.get('status', '-')}\n\n*🚨 PROBLEM ALARM:*\n{r.get('problem_alarm', '-')}\n\n*🔧 ACTION / PERBAIKAN:*\n{', '.join(r.get('action_reguler', [])) if isinstance(r.get('action_reguler'), list) else r.get('action_reguler', '-')}\n"
                cd = r.get('cond_data', {})
                if cd:
                    wa_text += f"\n*⚡ PARAMETER KELISTRIKAN:*\n- Teg AC (RN/SN/TN): {cd.get('v_rn','-')} / {cd.get('v_sn','-')} / {cd.get('v_tn','-')} V\n- Arus AC (R/S/T): {cd.get('i_r','-')} / {cd.get('i_s','-')} / {cd.get('i_t','-')} A\n- Load DC: {cd.get('load_dc', '-')} A\n- Set BCL: {cd.get('set_bcl', '-')}\n- Set Cap: {cd.get('set_cap', '-')}\n- Modul Rusak: {cd.get('mod_rusak', '-')}\n"
                    if cd.get('note_action'):
                        wa_text += f"\n*📝 NOTE/REMARK:*\n{cd.get('note_action')}\n"
            else:
                wa_text = f"*BERITA ACARA PREVENTIVE MAINTENANCE* ⚡\n📍 *Site:* {site_id} ({kategori_site} - {r.get('nop', '-')})\n📅 *Tanggal:* {r.get('timestamp', '-')}\n👷 *Pelaksana:* {r.get('teknisi', '-')}\n📊 *Status:* {r.get('status', '-')}\n\n*RINCIAN TINDAKAN:*\n{r.get('action', '-')}\n\n*POWER & LOAD:*\n- PLN: {r.get('pln_status', '-')}\n- Rectifier: {r.get('rect_brand', '-')} ({r.get('rect_out_v', '-')}V)\n- Load BTS: {r.get('total_load', '-')} A\n\n*SPAREPART:*\n{r.get('sparepart', '-')}"
            
            wa_url = f"https://wa.me/?text={urllib.parse.quote(wa_text)}"

            with st.expander(f"{icon}  |  {site_id}  |  {kategori_site}  |  {r.get('timestamp', '')}  |  Status: {stat}"):
                
                if kategori_site != "Pangkalan Bun" and not is_new_reguler:
                    if st.checkbox("📄 Buat Berita Acara (PDF Resmi)", key=f"prep_pdf_{i}"):
                        with st.spinner("⏳ Rendering Dokumen PDF Resolusi Tinggi..."):
                            try:
                                pdf_bytes = build_pdf(r)
                                st.download_button(label="📥 Download PDF Berita Acara", data=pdf_bytes, file_name=f"Berita_Acara_{site_id}.pdf", mime="application/pdf", key=f"dl_pdf_{i}", type="primary")
                            except Exception as e:
                                st.error(f"Terjadi kesalahan saat menyusun PDF: {e}")
                
                st.markdown("<hr style='border: 1px solid var(--primary-color); margin: 12px 0;'>", unsafe_allow_html=True)
                st.link_button("📱 Share Rangkuman Laporan ke WhatsApp", wa_url, use_container_width=True)
                
                # =========================================================================
                # TAMPILAN PANGKALAN BUN
                # =========================================================================
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

                # =========================================================================
                # TAMPILAN SITE REGULER (VERSI BARU)
                # =========================================================================
                elif is_new_reguler:
                    st.markdown("### 📋 SITE REGULER REPORT")
                    st.markdown(f"**👨‍🔧 FME:** {r.get('teknisi', '-')} ({r.get('role_fme', '-')}) &nbsp;&nbsp;|&nbsp;&nbsp; **🚨 Problem:** {r.get('problem_alarm', '-')}")
                    st.markdown(f"**🔧 Action Dikerjakan:** {', '.join(r.get('action_reguler', [])) if isinstance(r.get('action_reguler'), list) else r.get('action_reguler', '-')}")
                    
                    cond_data = r.get('cond_data', {})
                    if cond_data:
                        v_rst = cond_data.get('volt_ac_rst') or f"{cond_data.get('v_rn','-')}/{cond_data.get('v_sn','-')}/{cond_data.get('v_tn','-')}"
                        i_rst = cond_data.get('load_ac_rst') or f"{cond_data.get('i_r','-')}/{cond_data.get('i_s','-')}/{cond_data.get('i_t','-')}"
                        st.info(f"**⚡ Parameter Kelistrikan:** Teg AC R/S/T ({v_rst} V) | Load AC R/S/T ({i_rst} A) | Load DC ({cond_data.get('load_dc','-')} A) | BCL ({cond_data.get('set_bcl','-')}) | Capacity ({cond_data.get('set_cap','-')}) | Modul Rusak ({cond_data.get('mod_rusak','-')})")
                        if cond_data.get('note_action'):
                            st.warning(f"**📝 NOTE/REMARK Lapangan:** {cond_data.get('note_action')}")
                        
                    st.markdown("<br>", unsafe_allow_html=True)
                    t1, t2, t3 = st.tabs(["📸 Foto Problem Alarm", "📸 Foto Action / Perbaikan", "📸 Eviden Pengukuran Power"])
                    with t1: tampilkan_grid_foto(r.get('url_problem', []), "Dokumentasi Problem Alarm")
                    with t2: tampilkan_grid_foto(r.get('url_action', []), "Dokumentasi Perbaikan / Action")
                    with t3:
                        if cond_data:
                            eviden_list = [
                                (cond_data.get('url_v_rn'), f"Teg AC R-N: {cond_data.get('v_rn', '-')} V"),
                                (cond_data.get('url_v_sn'), f"Teg AC S-N: {cond_data.get('v_sn', '-')} V"),
                                (cond_data.get('url_v_tn'), f"Teg AC T-N: {cond_data.get('v_tn', '-')} V"),
                                (cond_data.get('url_i_r'), f"Load AC R: {cond_data.get('i_r', '-')} A"),
                                (cond_data.get('url_i_s'), f"Load AC S: {cond_data.get('i_s', '-')} A"),
                                (cond_data.get('url_i_t'), f"Load AC T: {cond_data.get('i_t', '-')} A"),
                                (cond_data.get('url_ldc'), f"Load DC Total: {cond_data.get('load_dc', '-')} A"),
                                (cond_data.get('url_bcl'), f"Settingan BCL: {cond_data.get('set_bcl', '-')}"),
                                (cond_data.get('url_cap'), f"Settingan Capacity: {cond_data.get('set_cap', '-')}"),
                                (cond_data.get('url_mod'), f"Module Kontroler: {cond_data.get('mod_rusak', '-')}")
                            ]
                            ada_foto_param = False
                            for u_ev, cap_ev in eviden_list:
                                if u_ev:
                                    ada_foto_param = True
                                    tampilkan_grid_foto([u_ev], f"📌 {cap_ev}")
                            if not ada_foto_param:
                                st.info("Belum ada foto eviden per-parameter pada laporan ini.")
                    
                    if st.session_state['role'] == 'Admin':
                        with st.container(border=True):
                            st.markdown("<h4 style='color: var(--primary-color);'>🛠 REVISI DATA & TAMBAH LAMPIRAN</h4>", unsafe_allow_html=True)
                            c_edit1, c_edit2 = st.columns(2)
                            with c_edit1:
                                new_site = st.text_input("Edit Nama Site / ID", r.get('site_name',''), key=f"esite_{i}")
                                kat_options = ["SPS", "Site Reguler"]
                                curr_kat = r.get('kategori', 'SPS')
                                if curr_kat not in kat_options: curr_kat = "SPS"
                                new_kat = st.selectbox("Edit Kategori Site", kat_options, index=kat_options.index(curr_kat), key=f"ekat_{i}")
                                
                                nop_options = ["Palangkaraya", "Pangkalan Bun", "Tarakan", "Pontianak", "Lainnya"]
                                curr_nop = r.get('nop', 'Palangkaraya')
                                if curr_nop not in nop_options: curr_nop = "Lainnya"
                                new_nop = st.selectbox("Edit NOP Area", nop_options, index=nop_options.index(curr_nop), key=f"enop_{i}")
                                new_tek = st.text_input("Edit Teknisi", r.get('teknisi',''), key=f"et_{i}")
                            with c_edit2:
                                status_options = ["Normal", "Minor Issue", "Major/Critical"]
                                curr_stat = r.get('status', 'Normal')
                                if curr_stat not in status_options: curr_stat = "Normal"
                                new_status = st.selectbox("Edit Status Akhir", status_options, index=status_options.index(curr_stat), key=f"estat_{i}")
                                new_sp = st.text_input("Edit Sparepart", r.get('sparepart',''), key=f"es_{i}")
                            new_act = st.text_area("Edit Action / Tindakan", r.get('action',''), key=f"ea_{i}")
                            st.markdown("**Tambah Lampiran Susulan (Otomatis Masuk Cloud):**")
                            up_f = st.file_uploader("📸 Fisik / Shading", accept_multiple_files=True, key=f"uf_{i}")
                            up_p = st.file_uploader("📸 Panel Surya", accept_multiple_files=True, key=f"up_{i}")
                            up_b = st.file_uploader("📸 Baterai / Grounding", accept_multiple_files=True, key=f"ub_{i}")
                            up_e = st.file_uploader("📸 Elektrikal / Mesin", accept_multiple_files=True, key=f"ue_{i}")
                            up_dl = st.file_uploader("📂 Datalog (Zip, Csv, xlsx)", accept_multiple_files=True, key=f"udl_{i}")

                            if st.button("💾 Simpan Perubahan ke Server", key=f"btn_{i}"):
                                with st.spinner("Mengirim Revisi ke Database Utama..."):
                                    r['site_name'], r['kategori'], r['nop'], r['status'], r['teknisi'], r['action'], r['sparepart'] = new_site, new_kat, new_nop, new_status, new_tek, new_act, new_sp
                                    if up_f: r['extras_fisik'] = r.get('extras_fisik', []) + upload_multiple_images(up_f)
                                    if up_p: r['extras_panel'] = r.get('extras_panel', []) + upload_multiple_images(up_p)
                                    if up_b: r['extras_baterai'] = r.get('extras_baterai', []) + upload_multiple_images(up_b)
                                    if up_e: r['extras_elektrikal'] = r.get('extras_elektrikal', []) + upload_multiple_images(up_e)
                                    if up_dl:
                                        new_datalog = []
                                        for df in up_dl:
                                            du = upload_datalog(df, "prev_datalog")
                                            if du: new_datalog.append({"name": df.name, "url": du})
                                        r['datalog_files'] = r.get('datalog_files', []) + new_datalog
                                    
                                    sheet = connect_gsheets(r.get('_sheet', 'Report Preventive'))
                                    if sheet:
                                        row_num = r.get('_row', i + 2)
                                        sheet.update_cell(row_num, 2, new_site)
                                        sheet.update_cell(row_num, 3, new_nop)
                                        sheet.update_cell(row_num, 4, new_tek)
                                        sheet.update_cell(row_num, 5, new_status)
                                        sheet.update_cell(row_num, 6, json.dumps(r))
                                        st.cache_data.clear() 
                                st.success("✅ REVISI BERHASIL! Data & Foto tersimpan permanen.")
                                st.rerun()

# =========================================================================
# MENU 3: EXCEL LIVE-EDITOR (MONITORING IMPROVEMENT)
# =========================================================================
elif menu == "📈 Monitoring Improvement":
    render_header_logo()
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
    render_header_logo()
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

# -------------------------------------------------------------------------
# FOOTER HAK CIPTA OKTA PRADIKA
# -------------------------------------------------------------------------
st.markdown("""
    <div class="footer-okta">
        🚀 System Application & Database Management<br>
        <span>Create Data By Okta Pradika</span> © 2026
    </div>
""", unsafe_allow_html=True)
