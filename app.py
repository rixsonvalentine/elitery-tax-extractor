import streamlit as st
import pdfplumber
import re
import pandas as pd
from io import BytesIO

st.set_page_config(page_title="Elitery Tax Extractor v11", page_icon="🧾")
st.title("📊 Elitery Tax Extractor (Ultra-Logic v11)")

def clean_strict(text):
    if not text: return ""
    return re.sub(r'\s+', ' ', text).strip()

def extract_surgical_v11(pdf_file):
    try:
        with pdfplumber.open(pdf_file) as pdf:
            full_text = ""
            for p in pdf.pages:
                t = p.extract_text()
                if t: full_text += t + "\n"
            
            if not full_text.strip():
                return {"NOMOR_BPU": "ERROR: Scan/Kosong", "NAMA_FILE": pdf_file.name}

            clean_full = clean_strict(full_text)
            res = {}

            # 1. NOMOR & MASA PAJAK 
            # Mencari pola alfanumerik 9-10 digit di dekat Masa Pajak MM-YYYY
            masa_match = re.search(r'(\d{2}-20\d{2})', clean_full)
            res['MASA_PAJAK'] = masa_match.group(1) if masa_match else "N/A"
            
            # Cari Nomor BPU (biasanya 9 karakter sebelum atau sesudah Masa Pajak)
            nomor_pattern = re.findall(r'\b[A-Z0-9]{9}\b', clean_full)
            res['NOMOR_BPU'] = next((n for n in nomor_pattern if n != "INDONESIA"), "N/A")

            res['STATUS'] = "NORMAL" if "NORMAL" in clean_full.upper() else "PEMBETULAN" [cite: 8]

            # 2. BAGIAN A (Penerima - Elitery) 
            # Menangani teks yang menempel seperti ":0313..."
            res['A.1_NPWP_NIK_PENERIMA'] = re.search(r'A\.1.*?[:\s]*(\d{15,16})', clean_full).group(1) if re.search(r'A\.1.*?[:\s]*(\d{15,16})', clean_full) else "N/A" [cite: 10]
            res["A.2_NAMA_PENERIMA"] = re.search(r'A\.2\s+NAMA\s*[:\s]*(.*?)\s+A\.3', clean_full, re.I).group(1).strip() if re.search(r'A\.2\s+NAMA\s*[:\s]*(.*?)\s+A\.3', clean_full, re.I) else "N/A" [cite: 10]
            res["A.3_NITKU_PENERIMA"] = re.search(r'A\.3.*?NITKU\)\s*[:\s]*(\d+)', clean_full).group(1) if re.search(r'A\.3.*?NITKU\)\s*[:\s]*(\d+)', clean_full) else "N/A" [cite: 10]

            # 3. BAGIAN B (Transaksi)
            res["B.2_JENIS_PPH"] = "Pasal 23" if "Pasal 23" in clean_full else "N/A" [cite: 15]
            
            # Kode Objek Pajak (B.3) 
            b3_match = re.search(r'(\d{2}-\d{3}-\d{2})', clean_full)
            res['B.3_KODE_OBJEK_PAJAK'] = b3_match.group(1) if b3_match else "N/A" [cite: 20]

            # DPP dan PPh (B.5 & B.7) [cite: 24, 27]
            # Mencari angka berformat 1.000.000
            money_patterns = re.findall(r'(\d{1,3}(?:\.\d{3})+)', clean_full)
            if len(money_patterns) >= 2:
                # Biasanya DPP adalah angka ribuan besar pertama di area B
                # Dan PPh adalah angka ribuan setelah Tarif
                res["B.5_DPP"] = money_patterns[-2] [cite: 24]
                res["B.7_PPH_DIPOTONG"] = money_patterns[-1] [cite: 27]
            else:
                res["B.5_DPP"] = "0"
                res["B.7_PPH_DIPOTONG"] = "0"

            # 4. DOKUMEN DASAR
            res["B.9_JENIS_DOKUMEN"] = re.search(r'Jenis Dokumen\s*[:\s]*(.*?)\s+Tanggal', clean_full, re.I).group(1).strip() if re.search(r'Jenis Dokumen\s*[:\s]*(.*?)\s+Tanggal', clean_full, re.I) else "N/A"
            doc_num = re.search(r'Nomor Dokumen\s*[:\s]*(\d+)', clean_full, re.I)
            res["B.10_NOMOR_DOKUMEN"] = "'" + doc_num.group(1) if doc_num else "N/A"
            res["B.11_TANGGAL_DOKUMEN"] = re.search(r'Tanggal\s*[:\s]*(\d{2}\s\w+\s\d{4})', clean_full, re.I).group(1) if re.search(r'Tanggal\s*[:\s]*(\d{2}\s\w+\s\d{4})', clean_full, re.I) else "N/A"

            # 5. BAGIAN C (Pemotong) 
            c_block = full_text.split('C. IDENTITAS')[-1]
            c_clean = clean_strict(c_block)
            
            res['C.1_NPWP_PEMOTONG'] = re.search(r'C\.1.*?[:\s]*(\d{15,16})', c_clean).group(1) if re.search(r'C\.1.*?[:\s]*(\d{15,16})', c_clean) else "N/A" [cite: 41]
            res['C.2_NITKU_PEMOTONG'] = re.search(r'C\.2.*?NITKU\)\s*[:\s]*(\d+)', c_clean).group(1) if re.search(r'C\.2.*?NITKU\)\s*[:\s]*(\d+)', c_clean) else "N/A" [cite: 41]
            
            # Nama Pemotong (C.3)
            c3_match = re.search(r'C\.3\s+NAMA\s+PEMOTONG.*?PPh\s*[:\s]*(.*?)\s*C\.4', c_clean, re.I)
            res["C.3_NAMA_PEMOTONG"] = c3_match.group(1).strip() if c3_match else "N/A" [cite: 41]
            
            res["C.4_TANGGAL_BPU"] = re.search(r'C\.4\s+TANGGAL\s*[:\s]*(\d{2}\s\w+\s\d{4})', c_clean).group(1) if re.search(r'C\.4\s+TANGGAL\s*[:\s]*(\d{2}\s\w+\s\d{4})', c_clean) else "N/A" [cite: 41]
            res["C.5_NAMA_PENANDATANGAN"] = re.search(r'C\.5\s+NAMA\s+PENANDATANGAN\s*[:\s]*(.*?)\s*(?:C\.6|Pernyataan|$)', c_clean).group(1).strip() if re.search(r'C\.5\s+NAMA\s+PENANDATANGAN\s*[:\s]*(.*?)\s*(?:C\.6|Pernyataan|$)', c_clean) else "N/A" [cite: 41]

            res["NAMA_FILE"] = pdf_file.name
            return res
    except Exception as e:
        return {"NOMOR_BPU": f"ERROR: {str(e)}", "NAMA_FILE": pdf_file.name}

# UI Streamlit
uploaded_files = st.file_uploader("Upload PDF Bukti Potong", type="pdf", accept_multiple_files=True)

if uploaded_files:
    if st.button("🚀 Jalankan Ekstraksi Versi 11"):
        all_results = [extract_surgical_v11(f) for f in uploaded_files]
        df = pd.DataFrame(all_results)
        st.dataframe(df)
        
        output = BytesIO()
        with pd.ExcelWriter(output, engine='xlsxwriter') as writer:
            df.to_excel(writer, index=False)
        st.download_button(label="📥 Download Excel", data=output.getvalue(), file_name="Rekap_Pajak_Final_v11.xlsx")
