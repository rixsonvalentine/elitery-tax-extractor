import streamlit as st
import pdfplumber
import re
import pandas as pd
from io import BytesIO

st.set_page_config(page_title="Elitery Full Tax Extractor v8", page_icon="📝")
st.title("📊 Elitery Full Tax Extractor (Ultra-Robust v8)")

def extract_surgical_v8(pdf_file):
    try:
        with pdfplumber.open(pdf_file) as pdf:
            # Ambil teks dan bersihkan spasi berlebih
            raw_text = "\n".join([p.extract_text() for p in pdf.pages if p.extract_text()])
            text_clean = re.sub(r' +', ' ', raw_text) # Normalisasi spasi
            
            if not text_clean.strip():
                return {"NOMOR_BPU": "ERROR: Scan/Kosong", "NAMA_FILE": pdf_file.name}

            res = {}

            # 1. NOMOR, MASA PAJAK, & STATUS
            # Mencari pola: [Kode] [MM-YYYY] [Status]
            header_regex = r'([A-Z0-9]{5,})\s+(\d{2}-\d{4})\s+(?:TIDAK\s+)?(?:FINAL\s+)?(NORMAL|PEMBETULAN)'
            h_match = re.search(header_regex, text_clean)
            if h_match:
                res['NOMOR_BPU'] = h_match.group(1)
                res['MASA_PAJAK'] = h_match.group(2)
                res['STATUS'] = h_match.group(3)
            else:
                res['NOMOR_BPU'] = "N/A"
                res['MASA_PAJAK'] = "N/A"
                res['STATUS'] = "NORMAL" if "NORMAL" in text_clean.upper() else "PEMBETULAN"

            # 2. BAGIAN A (Penerima)
            res["A.1_NPWP_NIK_PENERIMA"] = re.search(r'A\.1\s+NPWP\s*/\s*NIK\s*:\s*(\d+)', text_clean, re.I).group(1) if re.search(r'A\.1\s+NPWP\s*/\s*NIK\s*:\s*(\d+)', text_clean, re.I) else "N/A"
            res["A.2_NAMA_PENERIMA"] = re.search(r'A\.2\s+NAMA\s*:\s*(.*?)\s+A\.3', text_clean, re.S).group(1).strip() if re.search(r'A\.2\s+NAMA\s*:\s*(.*?)\s+A\.3', text_clean, re.S) else "N/A"
            res["A.3_NITKU_PENERIMA"] = re.search(r'A\.3.*?NITKU\)\s*:\s*(\d+)', text_clean, re.S).group(1) if re.search(r'A\.3.*?NITKU\)\s*:\s*(\d+)', text_clean, re.S) else "N/A"

            # 3. BAGIAN B (Transaksi)
            res["B.2_JENIS_PPH"] = re.search(r'B\.2\s+Jenis\s+PPh\s*:\s*(.*?)\n', text_clean, re.I).group(1).strip() if re.search(r'B\.2\s+Jenis\s+PPh\s*:\s*(.*?)\n', text_clean, re.I) else "N/A"
            
            # Mencari baris tabel B.3 - B.7
            table_row = re.search(r'(\d{2}-\d{3}-\d{2})\s+(.*?)(?:\.|\s)(\d{1,3}(?:\.\d{3})*)\s+(\d+)\s+(\d{1,3}(?:\.\d{3})*)', text_clean)
            if table_row:
                res["B.3_KODE_OBJEK_PAJAK"] = table_row.group(1)
                res["B.4_OBJEK_PAJAK"] = table_row.group(2).strip()
                res["B.5_DPP"] = table_row.group(3)
                res["B.6_TARIF"] = table_row.group(4)
                res["B.7_PPH_DIPOTONG"] = table_row.group(5)
            else:
                res["B.3_KODE_OBJEK_PAJAK"] = re.search(r'(\d{2}-\d{3}-\d{2})', text_clean).group(1) if re.search(r'(\d{2}-\d{3}-\d{2})', text_clean) else "N/A"
                # Fallback jika baris berantakan, ambil angka ribuan terakhir
                amounts = re.findall(r'(\d{1,3}(?:\.\d{3})+)', text_clean)
                res["B.5_DPP"] = amounts[-2] if len(amounts) >= 2 else "0"
                res["B.7_PPH_DIPOTONG"] = amounts[-1] if len(amounts) >= 1 else "0"
                res["B.4_OBJEK_PAJAK"] = "N/A"
                res["B.6_TARIF"] = "N/A"

            # Dokumen Dasar
            res["B.9_JENIS_DOKUMEN"] = re.search(r'Jenis\s+Dokumen\s*:\s*(.*?)\s+Tanggal', text_clean).group(1).strip() if re.search(r'Jenis\s+Dokumen\s*:\s*(.*?)\s+Tanggal', text_clean) else "N/A"
            res["B.10_NOMOR_DOKUMEN"] = re.search(r'Nomor\s+Dokumen\s*:\s*(\d+)', text_clean).group(1) if re.search(r'Nomor\s+Dokumen\s*:\s*(\d+)', text_clean) else "N/A"
            res["B.11_TANGGAL_DOKUMEN"] = re.search(r'Tanggal\s*:\s*(\d{2}\s\w+\s\d{4})', text_clean).group(1) if re.search(r'Tanggal\s*:\s*(\d{2}\s\w+\s\d{4})', text_clean) else "N/A"

            # 4. BAGIAN C (Pemotong)
            c_block = text_clean.split('C. IDENTITAS')[-1]
            res["C.1_NPWP_PEMOTONG"] = re.search(r'C\.1\s+NPWP\s*/\s*NIK\s*:\s*(\d+)', c_block).group(1) if re.search(r'C\.1\s+NPWP\s*/\s*NIK\s*:\s*(\d+)', c_block) else "N/A"
            res["C.2_NITKU_PEMOTONG"] = re.search(r'NITKU\s*\)\s*/\s*SUBUNIT\s*ORGANISASI\s*:\s*(\d+)', c_block).group(1) if re.search(r'NITKU\s*\)\s*/\s*SUBUNIT\s*ORGANISASI\s*:\s*(\d+)', c_block) else "N/A"
            res["C.3_NAMA_PEMOTONG"] = re.search(r'PPh\s*:\s*(.*?)\n', c_block).group(1).strip() if re.search(r'PPh\s*:\s*(.*?)\n', c_block) else "N/A"
            res["C.4_TANGGAL_BPU"] = re.search(r'C\.4\s+TANGGAL\s*:\s*(.*?)\s*C\.5', c_block, re.S).group(1).strip() if re.search(r'C\.4\s+TANGGAL\s*:\s*(.*?)\s*C\.5', c_block, re.S) else "N/A"
            res["C.5_NAMA_PENANDATANGAN"] = re.search(r'C\.5\s+NAMA\s+PENANDATANGAN\s*:\s*(.*?)\s*(?:C\.6|Pernyataan|$)', c_block, re.S).group(1).strip() if re.search(r'C\.5\s+NAMA\s+PENANDATANGAN\s*:\s*(.*?)\s*(?:C\.6|Pernyataan|$)', c_block, re.S) else "N/A"

            res["NAMA_FILE"] = pdf_file.name
            return res
    except Exception as e:
        return {"NOMOR_BPU": f"ERROR: {str(e)}", "NAMA_FILE": pdf_file.name}

# Antarmuka Streamlit
uploaded_files = st.file_uploader("Upload PDF Bukti Potong", type="pdf", accept_multiple_files=True)

if uploaded_files:
    if st.button("🚀 Jalankan Ekstraksi Ultra-Robust"):
        all_results = [extract_surgical_v8(f) for f in uploaded_files]
        df = pd.DataFrame(all_results)
        st.dataframe(df)
        
        output = BytesIO()
        with pd.ExcelWriter(output, engine='xlsxwriter') as writer:
            df.to_excel(writer, index=False)
        st.download_button(label="📥 Download Excel", data=output.getvalue(), file_name="Rekap_Pajak_Full_v8.xlsx")
