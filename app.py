import streamlit as st
import pdfplumber
import re
import pandas as pd
from io import BytesIO

st.set_page_config(page_title="Elitery Tax Extractor v16", page_icon="🛡️")
st.title("🛡️ Elitery Tax Extractor (The Guardian v16)")

def get_nearest_digits(text, keyword, length=10):
    """Mencari deretan angka terdekat setelah kata kunci tertentu."""
    part = text.split(keyword)[-1] if keyword in text else ""
    match = re.search(r'(\d{' + str(length) + r',})', part)
    return match.group(1) if match else "N/A"

def extract_surgical_v16(pdf_file):
    try:
        with pdfplumber.open(pdf_file) as pdf:
            full_text = "\n".join([p.extract_text() for p in pdf.pages if p.extract_text()])
            if not full_text.strip():
                return {"NOMOR_BPU": "ERROR: Scan/Kosong", "NAMA_FILE": pdf_file.name}

            clean_text = re.sub(r'\s+', ' ', full_text)
            res = {}

            # --- 1. NOMOR BPU & MASA PAJAK ---
            # Mencari Masa Pajak MM-YYYY [cite: 4]
            masa = re.search(r'(\d{2}-202\d)', clean_text)
            res['MASA_PAJAK'] = masa.group(1) if masa else "N/A"
            
            # Mencari Nomor BPU (Cari kata 'NOMOR' atau 'BPPU' lalu ambil kode 9 digit terdekat) [cite: 3, 4]
            nomor_search = re.search(r'(?:NOMOR|BPPU).*?([A-Z0-9]{8,12})', clean_text, re.I)
            res['NOMOR_BPU'] = nomor_search.group(1) if (nomor_search and nomor_search.group(1) != "INDONESIA") else "N/A"
            res['STATUS'] = "PEMBETULAN" if "PEMBETULAN" in clean_text.upper() else "NORMAL"

            # --- 2. BAGIAN A (Penerima) ---
            # A.1 NPWP: Cari angka 15-16 digit setelah label A.1 [cite: 10]
            res['A.1_NPWP_NIK_PENERIMA'] = "'" + get_nearest_digits(clean_text, "A.1", 15)
            
            # A.2 NAMA [cite: 10]
            a2 = re.search(r'A\.2\s+NAMA\s*[:\s]*(.*?)\s+A\.3', clean_text, re.I | re.S)
            res["A.2_NAMA_PENERIMA"] = a2.group(1).strip() if a2 else "N/A"
            
            # A.3 NITKU: Cari angka panjang setelah label NITKU [cite: 10]
            res["A.3_NITKU_PENERIMA"] = "'" + get_nearest_digits(clean_text, "NITKU", 16)

            # --- 3. BAGIAN B (Transaksi) ---
            res["B.2_JENIS_PPH"] = "Pasal 23" if "Pasal 23" in clean_text else "N/A"
            b3 = re.search(r'(\d{2}-\d{3}-\d{2})', clean_text) # [cite: 20]
            res['B.3_KODE_OBJEK_PAJAK'] = b3.group(1) if b3 else "N/A"

            # B.5 DPP & B.7 PPh (Mencari angka ribuan dengan titik) [cite: 24, 27]
            money_vals = re.findall(r'(\d{1,3}(?:\.\d{3})+)', clean_text)
            if len(money_vals) >= 2:
                res["B.5_DPP"] = money_vals[-2]
                res["B.7_PPH_DIPOTONG"] = money_vals[-1]
            else:
                res["B.5_DPP"] = "0"
                res["B.7_PPH_DIPOTONG"] = "0"

            # --- 4. DOKUMEN DASAR ---
            res["B.9_JENIS_DOKUMEN"] = re.search(r'Jenis Dokumen\s*[:\s]*(.*?)\s+Tanggal', clean_text, re.I).group(1).strip() if re.search(r'Jenis Dokumen\s*[:\s]*(.*?)\s+Tanggal', clean_text, re.I) else "N/A"
            doc_num = re.search(r'Nomor Dokumen\s*[:\s]*(\d+)', clean_text, re.I)
            res["B.10_NOMOR_DOKUMEN"] = "'" + doc_num.group(1) if doc_num else "N/A"
            res["B.11_TANGGAL_DOKUMEN"] = re.search(r'Tanggal\s*[:\s]*(\d{2}\s\w+\s\d{4})', clean_text, re.I).group(1) if re.search(r'Tanggal\s*[:\s]*(\d{2}\s\w+\s\d{4})', clean_text, re.I) else "N/A"

            # --- 5. BAGIAN C (Pemotong) ---
            c_part = full_text.split('C. IDENTITAS')[-1] if 'C. IDENTITAS' in full_text else ""
            c_clean = re.sub(r'\s+', ' ', c_part)
            
            # C.1 NPWP PEMOTONG 
            res['C.1_NPWP_PEMOTONG'] = "'" + get_nearest_digits(c_clean, "C.1", 15)
            
            # C.2 NITKU PEMOTONG 
            res['C.2_NITKU_PEMOTONG'] = "'" + get_nearest_digits(c_clean, "NITKU", 16)
            
            # C.3 NAMA PEMOTONG 
            c3_raw = re.search(r'PPh\s*[:\s]*(.*?)\s*C\.4', c_clean, re.I | re.S)
            if c3_raw:
                nama_raw = c3_raw.group(1).strip()
                # Hapus angka NITKU yang mungkin menempel di depan nama
                res["C.3_NAMA_PEMOTONG"] = re.sub(r'^[\d-]+\s*-\s*', '', nama_raw)
            else:
                res["C.3_NAMA_PEMOTONG"] = "N/A"
            
            res["C.4_TANGGAL_BPU"] = re.search(r'C\.4\s+TANGGAL\s*[:\s]*(\d{2}\s\w+\s\d{4})', c_clean).group(1) if re.search(r'C\.4\s+TANGGAL\s*[:\s]*(\d{2}\s\w+\s\d{4})', c_clean) else "N/A"
            res["C.5_NAMA_PENANDATANGAN"] = re.search(r'C\.5\s+NAMA\s+PENANDATANGAN\s*[:\s]*(.*?)\s*(?:C\.6|$)', c_clean, re.S).group(1).strip() if re.search(r'C\.5\s+NAMA\s+PENANDATANGAN\s*[:\s]*(.*?)\s*(?:C\.6|$)', c_clean, re.S) else "N/A"

            res["NAMA_FILE"] = pdf_file.name
            return res
    except Exception as e:
        return {"NOMOR_BPU": f"ERROR: {str(e)}", "NAMA_FILE": pdf_file.name}

# --- APLIKASI STREAMLIT ---
if 'df_result' not in st.session_state:
    st.session_state.df_result = None

uploaded_files = st.file_uploader("Upload PDF Bukti Potong", type="pdf", accept_multiple_files=True)

if uploaded_files:
    if st.button("🚀 Jalankan Ekstraksi Versi 16"):
        with st.spinner('Memproses data lengkap A.1 - C.5...'):
            all_results = [extract_surgical_v16(f) for f in uploaded_files]
            st.session_state.df_result = pd.DataFrame(all_results)
            st.success("Selesai!")

if st.session_state.df_result is not None:
    st.dataframe(st.session_state.df_result)
    output = BytesIO()
    with pd.ExcelWriter(output, engine='xlsxwriter') as writer:
        st.session_state.df_result.to_excel(writer, index=False)
    st.download_button(label="📥 Download Hasil Rekap (Excel)", data=output.getvalue(), file_name="Rekap_Pajak_Elitery_v16.xlsx")
