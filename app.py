import streamlit as st
import pdfplumber
import re
import pandas as pd
from io import BytesIO

st.set_page_config(page_title="Elitery Tax Extractor v10", page_icon="📑")
st.title("📊 Elitery Tax Extractor (Master-Logic v10)")

def clean_text_v10(text):
    if not text: return ""
    return re.sub(r'\s+', ' ', text).strip()

def extract_surgical_v10(pdf_file):
    try:
        with pdfplumber.open(pdf_file) as pdf:
            full_text = ""
            for p in pdf.pages:
                t = p.extract_text()
                if t: full_text += t + "\n"
            
            if not full_text.strip():
                return {"NOMOR_BPU": "ERROR: Scan/Kosong", "NAMA_FILE": pdf_file.name}

            # Normalisasi teks untuk pencarian baris
            clean_full = clean_text_v10(full_text)
            res = {}

            # 1. NOMOR & MASA PAJAK
            # Mencari pola: [Kode Alfanumerik] [MM-YYYY] [Status]
            h_match = re.search(r'([A-Z0-9]{8,15})\s+(\d{2}-20\d{2})\s+.*?(NORMAL|PEMBETULAN)', clean_full)
            if h_match:
                res['NOMOR_BPU'] = h_match.group(1)
                res['MASA_PAJAK'] = h_match.group(2)
                res['STATUS'] = h_match.group(3)
            else:
                res['NOMOR_BPU'] = re.search(r'\b([A-Z0-9]{9})\b', clean_full).group(1) if re.search(r'\b([A-Z0-9]{9})\b', clean_full) else "N/A"
                res['MASA_PAJAK'] = re.search(r'(\d{2}-20\d{2})', clean_full).group(1) if re.search(r'(\d{2}-20\d{2})', clean_full) else "N/A"
                res['STATUS'] = "NORMAL" if "NORMAL" in clean_full else "PEMBETULAN"

            # 2. BAGIAN A (Penerima - Elitery)
            res['A.1_NPWP_NIK_PENERIMA'] = re.search(r'A\.1.*?:\s*(\d+)', clean_full).group(1) if re.search(r'A\.1.*?:\s*(\d+)', clean_full) else "N/A"
            res["A.2_NAMA_PENERIMA"] = re.search(r'A\.2\s+NAMA\s*:\s*(.*?)\s+A\.3', clean_full).group(1).strip() if re.search(r'A\.2\s+NAMA\s*:\s*(.*?)\s+A\.3', clean_full) else "N/A"
            res["A.3_NITKU_PENERIMA"] = re.search(r'A\.3.*?NITKU\)\s*:\s*(\d+)', clean_full).group(1) if re.search(r'A\.3.*?NITKU\)\s*:\s*(\d+)', clean_full) else "N/A"

            # 3. BAGIAN B (Transaksi)
            res["B.2_JENIS_PPH"] = re.search(r'B\.2\s+Jenis\s+PPh\s*:\s*(.*?)\s+KODE', clean_full, re.I).group(1).strip() if re.search(r'B\.2\s+Jenis\s+PPh\s*:\s*(.*?)\s+KODE', clean_full, re.I) else "N/A"
            
            # Mencari Tabel B.3 - B.7 (Kode: XX-XXX-XX)
            table_row = re.search(r'(\d{2}-\d{3}-\d{2})\s+(.*?)\s+(\d{1,3}(?:\.\d{3})+)\s+(\d+)\s+(\d{1,3}(?:\.\d{3})+)', clean_full)
            if table_row:
                res["B.3_KODE_OBJEK_PAJAK"] = table_row.group(1)
                res["B.4_OBJEK_PAJAK"] = table_row.group(2).strip()
                res["B.5_DPP"] = table_row.group(3)
                res["B.6_TARIF"] = table_row.group(4)
                res["B.7_PPH_DIPOTONG"] = table_row.group(5)
            else:
                # Fallback jika angka menempel
                res["B.3_KODE_OBJEK_PAJAK"] = re.search(r'(\d{2}-\d{3}-\d{2})', clean_full).group(1) if re.search(r'(\d{2}-\d{3}-\d{2})', clean_full) else "N/A"
                amounts = re.findall(r'(\d{1,3}(?:\.\d{3})+)', clean_full)
                res["B.5_DPP"] = amounts[-2] if len(amounts) >= 2 else "0"
                res["B.7_PPH_DIPOTONG"] = amounts[-1] if len(amounts) >= 1 else "0"
                res["B.4_OBJEK_PAJAK"] = "N/A"
                res["B.6_TARIF"] = "N/A"

            # Dokumen Dasar
            res["B.9_JENIS_DOKUMEN"] = re.search(r'Jenis\s+Dokumen\s*:\s*(.*?)\s+Tanggal', clean_full).group(1).strip() if re.search(r'Jenis\s+Dokumen\s*:\s*(.*?)\s+Tanggal', clean_full) else "N/A"
            # Pastikan Nomor Dokumen ditarik dengan benar (Abaikan scientific notation)
            doc_num = re.search(r'Nomor\s+Dokumen\s*:\s*(\d+)', clean_full)
            res["B.10_NOMOR_DOKUMEN"] = "'" + doc_num.group(1) if doc_num else "N/A"
            res["B.11_TANGGAL_DOKUMEN"] = re.search(r'Tanggal\s*:\s*(\d{2}\s\w+\s\d{4})', clean_full).group(1) if re.search(r'Tanggal\s*:\s*(\d{2}\s\w+\s\d{4})', clean_full) else "N/A"

            # 4. BAGIAN C (Pemotong)
            c_block = full_text.split('C. IDENTITAS')[-1]
            c_clean = clean_text_v10(c_block)
            
            res['C.1_NPWP_PEMOTONG'] = re.search(r'C\.1.*?:\s*(\d+)', c_clean).group(1) if re.search(r'C\.1.*?:\s*(\d+)', c_clean) else "N/A"
            res['C.2_NITKU_PEMOTONG'] = re.search(r'C\.2.*?NITKU\)\s*/\s*SUBUNIT.*?:\s*(\d+)', c_clean).group(1) if re.search(r'C\.2.*?NITKU\)\s*/\s*SUBUNIT.*?:\s*(\d+)', c_clean) else "N/A"
            
            # Nama Pemotong (C.3)
            c3_match = re.search(r'C\.3\s+NAMA\s+PEMOTONG.*?\s*PPh\s*:\s*(.*?)\s*C\.4', c_clean, re.I)
            res["C.3_NAMA_PEMOTONG"] = c3_match.group(1).strip() if c3_match else "N/A"
            
            res["C.4_TANGGAL_BPU"] = re.search(r'C\.4\s+TANGGAL\s*:\s*(.*?)\s*C\.5', c_clean).group(1).strip() if re.search(r'C\.4\s+TANGGAL\s*:\s*(.*?)\s*C\.5', c_clean) else "N/A"
            res["C.5_NAMA_PENANDATANGAN"] = re.search(r'C\.5\s+NAMA\s+PENANDATANGAN\s*:\s*(.*?)\s*(?:C\.6|Pernyataan|$)', c_clean).group(1).strip() if re.search(r'C\.5\s+NAMA\s+PENANDATANGAN\s*:\s*(.*?)\s*(?:C\.6|Pernyataan|$)', c_clean) else "N/A"

            res["NAMA_FILE"] = pdf_file.name
            return res
    except Exception as e:
        return {"NOMOR_BPU": f"ERROR: {str(e)}", "NAMA_FILE": pdf_file.name}

# UI Streamlit
uploaded_files = st.file_uploader("Upload PDF Bukti Potong", type="pdf", accept_multiple_files=True)

if uploaded_files:
    if st.button("🚀 Jalankan Ekstraksi Versi 10"):
        all_results = [extract_surgical_v10(f) for f in uploaded_files]
        df = pd.DataFrame(all_results)
        st.dataframe(df)
        
        output = BytesIO()
        with pd.ExcelWriter(output, engine='xlsxwriter') as writer:
            df.to_excel(writer, index=False)
        st.download_button(label="📥 Download Excel", data=output.getvalue(), file_name="Rekap_Pajak_Final_v10.xlsx")
