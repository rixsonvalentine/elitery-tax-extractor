import streamlit as st
import pdfplumber
import re
import pandas as pd
from io import BytesIO

st.set_page_config(page_title="Elitery Tax Extractor v17", page_icon="🎯")
st.title("🎯 Elitery Tax Extractor (Ultra-Precision v17)")

def clean_name_field(text):
    """Membersihkan nama dari NITKU atau label sisa."""
    if not text or text == "N/A": return "N/A"
    # Hapus pola NITKU (22 digit) dan tanda hubung di depannya
    text = re.sub(r'^\d{22}\s*-\s*', '', text)
    # Hapus label umum
    labels = ["NAMA PEMOTONG DAN/ATAU PEMUNGUT", "PPh", "DAN/ATAU PEMUNGUT", "NAMA :", ":"]
    for l in labels:
        text = re.sub(rf'(?i){l}', '', text)
    return text.strip()

def extract_surgical_v17(pdf_file):
    try:
        with pdfplumber.open(pdf_file) as pdf:
            full_text = "\n".join([p.extract_text() for p in pdf.pages if p.extract_text()])
            if not full_text.strip():
                return {"NOMOR_BPU": "ERROR: Scan/Kosong", "NAMA_FILE": pdf_file.name}

            # Bersihkan teks dasar
            text_clean = re.sub(r' +', ' ', full_text)
            res = {}

            # --- 1. NOMOR BPU & MASA PAJAK ---
            masa = re.search(r'(\d{2}-202\d)', text_clean)
            res['MASA_PAJAK'] = masa.group(1) if masa else "N/A"
            
            # Cari Nomor BPU: Kode 9 digit alfanumerik (Bukan UNIFIKASI/INDONESIA)
            nomor = "N/A"
            potential_nums = re.findall(r'\b[A-Z0-9]{9}\b', text_clean[:1000])
            for n in potential_nums:
                if any(c.isdigit() for c in n) and n not in ["UNIFIKASI", "INDONESIA", "PENERIMA"]:
                    nomor = n
                    break
            res['NOMOR_BPU'] = nomor
            res['STATUS'] = "PEMBETULAN" if "PEMBETULAN" in text_clean.upper() else "NORMAL"

            # --- 2. BAGIAN A (Penerima) ---
            # NPWP (15-16 digit)
            a1 = re.search(r'A\.1.*?(\d{15,16})', text_clean, re.S)
            res['A.1_NPWP_NIK_PENERIMA'] = "'" + a1.group(1) if a1 else "N/A"
            
            # Nama Penerima
            a2 = re.search(r'A\.2\s+NAMA\s*[:\s]*(.*?)(?:\s+A\.3|NPWP)', text_clean, re.I | re.S)
            res["A.2_NAMA_PENERIMA"] = a2.group(1).strip() if a2 else "N/A"
            
            # NITKU Penerima (22 digit)
            a3 = re.search(r'A\.3.*?(\d{22})', text_clean, re.S)
            res["A.3_NITKU_PENERIMA"] = "'" + a3.group(1) if a3 else "N/A"

            # --- 3. BAGIAN B (Transaksi) ---
            res["B.2_JENIS_PPH"] = "Pasal 23" if "Pasal 23" in text_clean else "N/A"
            b3 = re.search(r'(\d{2}-\d{3}-\d{2})', text_clean)
            res['B.3_KODE_OBJEK_PAJAK'] = b3.group(1) if b3 else "N/A"

            # DPP & PPh
            money_vals = re.findall(r'(\d{1,3}(?:\.\d{3})+)', text_clean)
            if len(money_vals) >= 2:
                res["B.5_DPP"] = money_vals[-2]
                res["B.7_PPH_DIPOTONG"] = money_vals[-1]
            else:
                res["B.5_DPP"] = "0"
                res["B.7_PPH_DIPOTONG"] = "0"

            # --- 4. DOKUMEN DASAR ---
            res["B.9_JENIS_DOKUMEN"] = re.search(r'Jenis Dokumen\s*[:\s]*(.*?)\s+Tanggal', text_clean, re.I).group(1).strip() if re.search(r'Jenis Dokumen\s*[:\s]*(.*?)\s+Tanggal', text_clean, re.I) else "N/A"
            doc_num = re.search(r'Nomor Dokumen\s*[:\s]*(\d+)', text_clean, re.I)
            res["B.10_NOMOR_DOKUMEN"] = "'" + doc_num.group(1) if doc_num else "N/A"
            res["B.11_TANGGAL_DOKUMEN"] = re.search(r'Tanggal\s*[:\s]*(\d{2}\s\w+\s\d{4})', text_clean, re.I).group(1) if re.search(r'Tanggal\s*[:\s]*(\d{2}\s\w+\s\d{4})', text_clean, re.I) else "N/A"

            # --- 5. BAGIAN C (Pemotong) ---
            c_block = text_clean.split('C. IDENTITAS')[-1] if 'C. IDENTITAS' in text_clean else text_clean
            
            # NPWP Pemotong (15-16 digit)
            c1 = re.search(r'C\.1.*?(\d{15,16})', c_block, re.S)
            res['C.1_NPWP_PEMOTONG'] = "'" + c1.group(1) if c1 else "N/A"
            
            # NITKU Pemotong (22 digit)
            c2 = re.search(r'(\d{22})', c_block)
            res['C.2_NITKU_PEMOTONG'] = "'" + c2.group(1) if c2 else "N/A"
            
            # Nama Pemotong (C.3)
            # Mengambil nama yang bersih dari NITKU dan label
            c3_match = re.search(r'C\.3\s+NAMA PEMOTONG.*?\n(.*?)\s+C\.4', c_block, re.S | re.I)
            if c3_match:
                res["C.3_NAMA_PEMOTONG"] = clean_name_field(c3_match.group(1))
            else:
                fallback_c3 = re.search(r'PPh\s*[:\s]*(.*?)\s*C\.4', c_block, re.I | re.S)
                res["C.3_NAMA_PEMOTONG"] = clean_name_field(fallback_c3.group(1)) if fallback_c3 else "N/A"
            
            res["C.4_TANGGAL_BPU"] = re.search(r'C\.4\s+TANGGAL\s*[:\s]*(\d{2}\s\w+\s\d{4})', c_block).group(1) if re.search(r'C\.4\s+TANGGAL\s*[:\s]*(\d{2}\s\w+\s\d{4})', c_block) else "N/A"
            res["C.5_NAMA_PENANDATANGAN"] = re.search(r'C\.5\s+NAMA\s+PENANDATANGAN\s*[:\s]*(.*?)(?:\s*C\.6|$)', c_block, re.S).group(1).strip() if re.search(r'C\.5\s+NAMA\s+PENANDATANGAN\s*[:\s]*(.*?)(?:\s*C\.6|$)', c_block, re.S) else "N/A"

            res["NAMA_FILE"] = pdf_file.name
            return res
    except Exception as e:
        return {"NOMOR_BPU": f"ERROR: {str(e)}", "NAMA_FILE": pdf_file.name}

# --- APLIKASI STREAMLIT ---
if 'df_result' not in st.session_state:
    st.session_state.df_result = None

uploaded_files = st.file_uploader("Upload PDF Bukti Potong", type="pdf", accept_multiple_files=True)

if uploaded_files:
    if st.button("🚀 Jalankan Ekstraksi Versi 17"):
        with st.spinner('Memproses data...'):
            all_results = [extract_surgical_v17(f) for f in uploaded_files]
            st.session_state.df_result = pd.DataFrame(all_results)
            st.success("Ekstraksi Berhasil!")

if st.session_state.df_result is not None:
    st.dataframe(st.session_state.df_result)
    output = BytesIO()
    with pd.ExcelWriter(output, engine='xlsxwriter') as writer:
        st.session_state.df_result.to_excel(writer, index=False)
    st.download_button(label="📥 Download Hasil Rekap", data=output.getvalue(), file_name="Rekap_Pajak_Elitery_v17.xlsx")
