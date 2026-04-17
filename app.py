import streamlit as st
import pdfplumber
import re
import pandas as pd
from io import BytesIO

st.set_page_config(page_title="Elitery Full Tax Extractor", page_icon="📝")
st.title("📊 Elitery Full Tax Extractor (A.1 - C.5)")

def extract_full_data(pdf_file):
    try:
        with pdfplumber.open(pdf_file) as pdf:
            full_text = "\n".join([p.extract_text() for p in pdf.pages if p.extract_text()])
            
            # Header Info
            no_masa = re.search(r'NOMOR\s+MASA PAJAK\s*\n\s*([A-Z0-9]+)\s+(\d{2}-\d{4})', full_text)
            
            res = {
                "NOMOR_BPU": no_masa.group(1) if no_masa else "N/A",
                "MASA_PAJAK": no_masa.group(2) if no_masa else "N/A",
                "STATUS": "NORMAL" if "NORMAL" in full_text.upper() else "PEMBETULAN"
            }

            # BAGIAN A
            res["A.1_NPWP_NIK_PENERIMA"] = re.search(r'A\.1\s+NPWP/NIK\s+:\s+(\d+)', full_text).group(1) if re.search(r'A\.1\s+NPWP/NIK\s+:\s+(\d+)', full_text) else "N/A"
            res["A.2_NAMA_PENERIMA"] = re.search(r'A\.2\s+NAMA\s+:\s+(.*?)\n', full_text).group(1).strip() if re.search(r'A\.2\s+NAMA\s+:\s+(.*?)\n', full_text) else "N/A"
            res["A.3_NITKU_PENERIMA"] = re.search(r'A\.3\s+NOMOR IDENTITAS.*?NITKU\)\s*:\s*(\d+)', full_text).group(1) if re.search(r'A\.3\s+NOMOR IDENTITAS.*?NITKU\)\s*:\s*(\d+)', full_text) else "N/A"

            # BAGIAN B
            res["B.1_JENIS_FASILITAS"] = re.search(r'B\.1\s+Jenis Fasilitas\s*:\s*(.*?)\n', full_text).group(1).strip() if re.search(r'B\.1\s+Jenis Fasilitas\s*:\s*(.*?)\n', full_text) else "N/A"
            res["B.2_JENIS_PPH"] = re.search(r'B\.2\s+Jenis PPh\s+(.*?)\n', full_text).group(1).strip() if re.search(r'B\.2\s+Jenis PPh\s+(.*?)\n', full_text) else "N/A"
            res["B.3_KODE_OBJEK_PAJAK"] = re.search(r'B\.3\s+(\d{2}-\d{3}-\d{2})', full_text).group(1) if re.search(r'B\.3\s+(\d{2}-\d{3}-\d{2})', full_text) else "N/A"
            
            # B.4 Objek Pajak Deskripsi
            b4_match = re.search(r'B\.4\s+(.*?)\n\s*(\d{1,3}(?:\.\d{3})+)', full_text, re.S)
            res["B.4_OBJEK_PAJAK"] = b4_match.group(1).replace('\n', ' ').strip() if b4_match else "N/A"

            # Keuangan B.5 - B.7
            amounts = re.findall(r'(\d{1,3}(?:\.\d{3})+)', full_text)
            res["B.5_DPP"] = amounts[-2] if len(amounts) >= 2 else "0"
            res["B.7_PPH_DIPOTONG"] = amounts[-1] if len(amounts) >= 1 else "0"
            
            tarif = re.search(r'TARIF\s*\(%\)\s*B\.6\s*(\d+)', full_text)
            res["B.6_TARIF"] = tarif.group(1) if tarif else "N/A"

            # Dokumen Dasar B.9 - B.11
            res["B.9_JENIS_DOKUMEN"] = re.search(r'Jenis Dokumen\s*:\s*(.*?)\n', full_text).group(1).strip() if re.search(r'Jenis Dokumen\s*:\s*(.*?)\n', full_text) else "N/A"
            res["B.10_NOMOR_DOKUMEN"] = re.search(r'Nomor Dokumen\s*:\s*(\d+)', full_text).group(1) if re.search(r'Nomor Dokumen\s*:\s*(\d+)', full_text) else "N/A"
            res["B.11_TANGGAL_DOKUMEN"] = re.search(r'Tanggal\s*:\s*(\d{2}\s\w+\s\d{4})', full_text).group(1) if re.search(r'Tanggal\s*:\s*(\d{2}\s\w+\s\d{4})', full_text) else "N/A"

            # BAGIAN C
            c_block = full_text.split('C. IDENTITAS')[-1]
            res["C.1_NPWP_PEMOTONG"] = re.search(r'C\.1\s+NPWP/NIK\s*:\s*(\d+)', c_block).group(1) if re.search(r'C\.1\s+NPWP/NIK\s*:\s*(\d+)', c_block) else "N/A"
            res["C.2_NITKU_PEMOTONG"] = re.search(r'C\.2.*?NITKU\)\s*:\s*(\d+)', c_block).group(1) if re.search(r'C\.2.*?NITKU\)\s*:\s*(\d+)', c_block) else "N/A"
            res["C.3_NAMA_PEMOTONG"] = re.search(r'C\.3\s+NAMA PEMOTONG.*?\n(.*?)\n', c_block).group(1).strip() if re.search(r'C\.3\s+NAMA PEMOTONG.*?\n(.*?)\n', c_block) else "N/A"
            res["C.4_TANGGAL_BPU"] = re.search(r'C\.4\s+TANGGAL\s*:\s*(\d{2}\s\w+\s\d{4})', c_block).group(1) if re.search(r'C\.4\s+TANGGAL\s*:\s*(\d{2}\s\w+\s\d{4})', c_block) else "N/A"
            res["C.5_NAMA_PENANDATANGAN"] = re.search(r'C\.5\s+NAMA PENANDATANGAN\s*:\s*(.*?)\n', c_block).group(1).strip() if re.search(r'C\.5\s+NAMA PENANDATANGAN\s*:\s*(.*?)\n', c_block) else "N/A"

            res["NAMA_FILE"] = pdf_file.name
            return res
    except Exception as e:
        return {"ERROR": str(e), "NAMA_FILE": pdf_file.name}

uploaded_files = st.file_uploader("Upload PDF Bukti Potong", type="pdf", accept_multiple_files=True)

if uploaded_files:
    if st.button("🚀 Jalankan Ekstraksi Penuh"):
        all_results = [extract_full_data(f) for f in uploaded_files]
        df = pd.DataFrame(all_results)
        st.dataframe(df)
        
        output = BytesIO()
        with pd.ExcelWriter(output, engine='xlsxwriter') as writer:
            df.to_excel(writer, index=False)
        
        st.download_button(label="📥 Download Excel", data=output.getvalue(), file_name="Rekap_Full_Pajak.xlsx")
