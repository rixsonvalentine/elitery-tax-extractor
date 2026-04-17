import streamlit as st
import pdfplumber
import re
import pandas as pd
from io import BytesIO

st.set_page_config(page_title="Elitery Tax Extractor v12", page_icon="🏦")
st.title("📊 Elitery Tax Extractor (Surgical Precision v12)")

def extract_surgical_v12(pdf_file):
    try:
        with pdfplumber.open(pdf_file) as pdf:
            full_text = ""
            for p in pdf.pages:
                t = p.extract_text()
                if t: full_text += t + "\n"
            
            if not full_text.strip():
                return {"NOMOR_BPU": "ERROR: Scan/Kosong", "NAMA_FILE": pdf_file.name}

            # Normalisasi teks: hapus spasi berlebih tapi pertahankan baris
            clean_lines = [re.sub(r' +', ' ', line).strip() for line in full_text.split('\n')]
            clean_full = " ".join(clean_lines)
            res = {}

            # 1. NOMOR & MASA PAJAK 
            # Mencari 9 digit alfanumerik yang diikuti oleh masa pajak MM-YYYY
            res['MASA_PAJAK'] = re.search(r'(\d{2}-202\d)', clean_full).group(1) if re.search(r'(\d{2}-202\d)', clean_full) else "N/A"
            
            # Cari Nomor BPU (9 karakter alfanumerik yang bukan kata kunci)
            potential_nums = re.findall(r'\b[A-Z0-9]{9}\b', clean_full)
            res['NOMOR_BPU'] = next((n for n in potential_nums if any(char.isdigit() for char in n) and n != "INDONESIA"), "N/A")

            # STATUS [cite: 8]
            res['STATUS'] = "PEMBETULAN" if "PEMBETULAN" in clean_full.upper() else "NORMAL"

            # 2. BAGIAN A (Penerima - Elitery) [cite: 10]
            # Ekstraksi NPWP (A.1): Mencari 15-16 digit setelah "A.1"
            a_block = clean_full.split('A. IDENTITAS')[1].split('B. PEMOTONGAN')[0] if 'A. IDENTITAS' in clean_full else clean_full
            res['A.1_NPWP_NIK_PENERIMA'] = re.search(r'(\d{15,16})', a_block).group(1) if re.search(r'(\d{15,16})', a_block) else "N/A"
            res["A.2_NAMA_PENERIMA"] = re.search(r'A\.2\s+NAMA\s*[:\s]*(.*?)\s+A\.3', clean_full, re.I).group(1).strip() if re.search(r'A\.2\s+NAMA\s*[:\s]*(.*?)\s+A\.3', clean_full, re.I) else "N/A"
            res["A.3_NITKU_PENERIMA"] = re.search(r'NITKU\)\s*[:\s]*(\d+)', a_block).group(1) if re.search(r'NITKU\)\s*[:\s]*(\d+)', a_block) else "N/A"

            # 3. BAGIAN B (Transaksi) [cite: 15, 20]
            res["B.2_JENIS_PPH"] = "Pasal 23" if "Pasal 23" in clean_full else "N/A"
            res['B.3_KODE_OBJEK_PAJAK'] = re.search(r'(\d{2}-\d{3}-\d{2})', clean_full).group(1) if re.search(r'(\d{2}-\d{3}-\d{2})', clean_full) else "N/A"

            # DPP (B.5) & PPh (B.7) 
            # Mencari angka berformat 1.000.000
            money_vals = re.findall(r'(\d{1,3}(?:\.\d{3})+)', clean_full)
            if len(money_vals) >= 2:
                # Berdasarkan struktur: DPP (B.5) muncul sebelum PPh (B.7)
                res["B.5_DPP"] = money_vals[-2]
                res["B.7_PPH_DIPOTONG"] = money_vals[-1]
            else:
                res["B.5_DPP"] = "0"
                res["B.7_PPH_DIPOTONG"] = "0"
            
            # Deskripsi Jasa (B.4) [cite: 23]
            b4_match = re.search(r'B\.4\s+(.*?)\s+(?:\d{1,3}(?:\.\d{3})+)', clean_full, re.S)
            res['B.4_OBJEK_PAJAK'] = b4_match.group(1).strip() if b4_match else "N/A"

            # 4. DOKUMEN DASAR [cite: 28, 30, 36]
            res["B.9_JENIS_DOKUMEN"] = re.search(r'Jenis Dokumen\s*[:\s]*(.*?)\s+Tanggal', clean_full, re.I).group(1).strip() if re.search(r'Jenis Dokumen\s*[:\s]*(.*?)\s+Tanggal', clean_full, re.I) else "N/A"
            res["B.10_NOMOR_DOKUMEN"] = "'" + re.search(r'Nomor Dokumen\s*[:\s]*(\d+)', clean_full, re.I).group(1) if re.search(r'Nomor Dokumen\s*[:\s]*(\d+)', clean_full, re.I) else "N/A"
            res["B.11_TANGGAL_DOKUMEN"] = re.search(r'Tanggal\s*[:\s]*(\d{2}\s\w+\s\d{4})', clean_full, re.I).group(1) if re.search(r'Tanggal\s*[:\s]*(\d{2}\s\w+\s\d{4})', clean_full, re.I) else "N/A"

            # 5. BAGIAN C (Pemotong) 
            c_block = full_text.split('C. IDENTITAS')[1] if 'C. IDENTITAS' in full_text else ""
            c_clean = clean_strict(c_block)
            
            res['C.1_NPWP_PEMOTONG'] = re.search(r'(\d{15,16})', c_clean).group(1) if re.search(r'(\d{15,16})', c_clean) else "N/A"
            res['C.2_NITKU_PEMOTONG'] = re.search(r'NITKU\)\s*[:\s]*(\d+)', c_clean).group(1) if re.search(r'NITKU\)\s*[:\s]*(\d+)', c_clean) else "N/A"
            
            # Nama Pemotong (C.3)
            # Mencari teks setelah label Nama Pemotong dan sebelum Tanggal (C.4)
            c3_match = re.search(r'PPh\s*[:\s]*(.*?)\s*C\.4', c_clean, re.I)
            res["C.3_NAMA_PEMOTONG"] = c3_match.group(1).strip() if c3_match else "N/A"
            
            res["C.4_TANGGAL_BPU"] = re.search(r'C\.4\s+TANGGAL\s*[:\s]*(\d{2}\s\w+\s\d{4})', c_clean).group(1) if re.search(r'C\.4\s+TANGGAL\s*[:\s]*(\d{2}\s\w+\s\d{4})', c_clean) else "N/A"
            res["C.5_NAMA_PENANDATANGAN"] = re.search(r'C\.5\s+NAMA\s+PENANDATANGAN\s*[:\s]*(.*?)\s*(?:C\.6|$)', c_clean).group(1).strip() if re.search(r'C\.5\s+NAMA\s+PENANDATANGAN\s*[:\s]*(.*?)\s*(?:C\.6|$)', c_clean) else "N/A"

            res["NAMA_FILE"] = pdf_file.name
            return res
    except Exception as e:
        return {"NOMOR_BPU": f"ERROR: {str(e)}", "NAMA_FILE": pdf_file.name}

# UI Streamlit
uploaded_files = st.file_uploader("Upload PDF Bukti Potong", type="pdf", accept_multiple_files=True)

if uploaded_files:
    if st.button("🚀 Jalankan Ekstraksi Versi 12"):
        all_results = [extract_surgical_v12(f) for f in uploaded_files]
        df = pd.DataFrame(all_results)
        st.dataframe(df)
        
        output = BytesIO()
        with pd.ExcelWriter(output, engine='xlsxwriter') as writer:
            df.to_excel(writer, index=False)
        st.download_button(label="📥 Download Excel", data=output.getvalue(), file_name="Rekap_Pajak_Final_v12.xlsx")
