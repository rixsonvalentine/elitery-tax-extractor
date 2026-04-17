import streamlit as st
import pdfplumber
import re
import pandas as pd
from io import BytesIO

st.set_page_config(page_title="Elitery Tax Extractor v9", page_icon="🧾")
st.title("📊 Elitery Full Tax Extractor (Ultra-Precise v9)")

def extract_surgical_v9(pdf_file):
    try:
        with pdfplumber.open(pdf_file) as pdf:
            # Mengambil teks dari semua halaman
            raw_text = ""
            for p in pdf.pages:
                t = p.extract_text()
                if t: raw_text += t + "\n"
            
            # Normalisasi teks dasar
            text_clean = re.sub(r' +', ' ', raw_text)
            
            if not text_clean.strip():
                return {"NOMOR_BPU": "ERROR: Scan/Kosong", "NAMA_FILE": pdf_file.name}

            res = {}

            # 1. NOMOR & MASA PAJAK (Mencari berdasarkan pola Tanggal MM-YYYY)
            masa_match = re.search(r'(\d{2}-202\d)', text_clean)
            if masa_match:
                res['MASA_PAJAK'] = masa_match.group(1)
                # Ambil kata alfanumerik tepat sebelum tanggal masa pajak
                pre_text = text_clean[:masa_match.start()].strip()
                words = pre_text.split()
                res['NOMOR_BPU'] = words[-1] if words else "N/A"
            else:
                res['MASA_PAJAK'] = "N/A"
                res['NOMOR_BPU'] = "N/A"

            res['STATUS'] = "NORMAL" if "NORMAL" in text_clean.upper() else "PEMBETULAN"

            # 2. BAGIAN A (Penerima - Elitery)
            # Mencari 16 digit NPWP pertama setelah label A.1
            a1_block = text_clean.split('A.1')[1] if 'A.1' in text_clean else ""
            res['A.1_NPWP_NIK_PENERIMA'] = re.search(r'(\d{15,16})', a1_block).group(1) if re.search(r'(\d{15,16})', a1_block) else "N/A"
            res["A.2_NAMA_PENERIMA"] = re.search(r'A\.2\s+NAMA\s*:\s*(.*?)\n', text_clean).group(1).strip() if re.search(r'A\.2\s+NAMA\s*:\s*(.*?)\n', text_clean) else "N/A"
            res["A.3_NITKU_PENERIMA"] = re.search(r'A\.3.*?NITKU\)\s*:\s*(\d+)', text_clean, re.S).group(1) if re.search(r'A\.3.*?NITKU\)\s*:\s*(\d+)', text_clean, re.S) else "N/A"

            # 3. BAGIAN B (Transaksi)
            res["B.2_JENIS_PPH"] = re.search(r'B\.2\s+Jenis\s+PPh\s*:\s*(.*?)\n', text_clean, re.I).group(1).strip() if re.search(r'B\.2\s+Jenis\s+PPh\s*:\s*(.*?)\n', text_clean, re.I) else "N/A"
            
            # Kode Objek Pajak (Pola XX-XXX-XX)
            b3_match = re.search(r'(\d{2}-\d{3}-\d{2})', text_clean)
            res['B.3_KODE_OBJEK_PAJAK'] = b3_match.group(1) if b3_match else "N/A"

            # DPP dan PPh (Mencari angka ribuan dengan titik)
            amounts = re.findall(r'(\d{1,3}(?:\.\d{3})+)', text_clean)
            res["B.5_DPP"] = amounts[-2] if len(amounts) >= 2 else "0"
            res["B.7_PPH_DIPOTONG"] = amounts[-1] if len(amounts) >= 1 else "0"
            
            # Tarif (B.6)
            tarif_match = re.search(r'B\.6\s*(\d+)', text_clean)
            res['B.6_TARIF'] = tarif_match.group(1) if tarif_match else "N/A"
            
            # Deskripsi Jasa (B.4)
            b4_match = re.search(r'B\.4\s+(.*?)\s+\d{1,3}(?:\.\d{3})+', text_clean, re.S)
            res['B.4_OBJEK_PAJAK'] = b4_match.group(1).replace('\n', ' ').strip() if b4_match else "N/A"

            # Dokumen Dasar (B.9 - B.11)
            res["B.9_JENIS_DOKUMEN"] = re.search(r'Jenis\s+Dokumen\s*:\s*(.*?)\s+Tanggal', text_clean).group(1).strip() if re.search(r'Jenis\s+Dokumen\s*:\s*(.*?)\s+Tanggal', text_clean) else "N/A"
            # Paksa menjadi string agar tidak scientific di Excel
            res["B.10_NOMOR_DOKUMEN"] = "'" + re.search(r'Nomor\s+Dokumen\s*:\s*(\d+)', text_clean).group(1) if re.search(r'Nomor\s+Dokumen\s*:\s*(\d+)', text_clean) else "N/A"
            res["B.11_TANGGAL_DOKUMEN"] = re.search(r'Tanggal\s*:\s*(\d{2}\s\w+\s\d{4})', text_clean).group(1) if re.search(r'Tanggal\s*:\s*(\d{2}\s\w+\s\d{4})', text_clean) else "N/A"

            # 4. BAGIAN C (Pemotong)
            c_block = text_clean.split('C. IDENTITAS')[-1]
            res['C.1_NPWP_PEMOTONG'] = re.search(r'C\.1.*?(\d{15,16})', c_block, re.S).group(1) if re.search(r'C\.1.*?(\d{15,16})', c_block, re.S) else "N/A"
            res['C.2_NITKU_PEMOTONG'] = re.search(r'C\.2.*?NITKU\)\s*.*?:\s*(\d+)', c_block, re.S).group(1) if re.search(r'C\.2.*?NITKU\)\s*.*?:\s*(\d+)', c_block, re.S) else "N/A"
            res["C.3_NAMA_PEMOTONG"] = re.search(r'PPh\s*:\s*(.*?)\n', c_block).group(1).strip() if re.search(r'PPh\s*:\s*(.*?)\n', c_block) else "N/A"
            res["C.4_TANGGAL_BPU"] = re.search(r'C\.4\s+TANGGAL\s*:\s*(.*?)\s*C\.5', c_block, re.S).group(1).strip() if re.search(r'C\.4\s+TANGGAL\s*:\s*(.*?)\s*C\.5', c_block, re.S) else "N/A"
            res["C.5_NAMA_PENANDATANGAN"] = re.search(r'C\.5\s+NAMA\s+PENANDATANGAN\s*:\s*(.*?)\s*(?:C\.6|Pernyataan|$)', c_block, re.S).group(1).strip() if re.search(r'C\.5\s+NAMA\s+PENANDATANGAN\s*:\s*(.*?)\s*(?:C\.6|Pernyataan|$)', c_block, re.S) else "N/A"

            res["NAMA_FILE"] = pdf_file.name
            return res
    except Exception as e:
        return {"NOMOR_BPU": f"ERROR: {str(e)}", "NAMA_FILE": pdf_file.name}

# UI Streamlit
uploaded_files = st.file_uploader("Upload PDF Bukti Potong", type="pdf", accept_multiple_files=True)

if uploaded_files:
    if st.button("🚀 Jalankan Ekstraksi Versi 9"):
        all_results = [extract_surgical_v9(f) for f in uploaded_files]
        df = pd.DataFrame(all_results)
        st.dataframe(df)
        
        output = BytesIO()
        with pd.ExcelWriter(output, engine='xlsxwriter') as writer:
            df.to_excel(writer, index=False)
        st.download_button(label="📥 Download Excel", data=output.getvalue(), file_name="Rekap_Pajak_Fixed_v9.xlsx")
