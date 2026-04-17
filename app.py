import streamlit as st
import pdfplumber
import re
import pandas as pd
from io import BytesIO

st.set_page_config(page_title="Elitery Tax Extractor v15", page_icon="💎")
st.title("💎 Elitery Tax Extractor (Final Polish v15)")

def clean_id_prefix(text):
    """Menghapus awalan angka ID, tanda hubung, dan label PPh dari teks."""
    if not text: return "N/A"
    # Hapus "PPh :" atau "PPh:"
    text = re.sub(r'(?i)PPh\s*[:\s]*', '', text)
    # Hapus deretan angka panjang di awal (ID/NITKU) yang diikuti tanda hubung
    text = re.search(r'(?:[\d-]+\s*-\s*)?(.*)', text).group(1)
    return text.strip()

def extract_surgical_v15(pdf_file):
    try:
        with pdfplumber.open(pdf_file) as pdf:
            full_text = ""
            for p in pdf.pages:
                t = p.extract_text()
                if t: full_text += t + "\n"
            
            if not full_text.strip():
                return {"NOMOR_BPU": "ERROR: Scan/Kosong", "NAMA_FILE": pdf_file.name}

            # Normalisasi teks dasar
            clean_text = re.sub(r' +', ' ', full_text)
            res = {}

            # --- 1. NOMOR BPU & MASA PAJAK ---
            masa_match = re.search(r'(\d{2}-202\d)', clean_text)
            res['MASA_PAJAK'] = masa_match.group(1) if masa_match else "N/A"
            
            # Cari Nomor BPU (9 Karakter Alfanumerik di header)
            # Biasanya muncul setelah kata 'BPPU' atau sebelum Masa Pajak
            header_area = clean_text[:500]
            nomor_match = re.search(r'\b([A-Z0-9]{9})\b', header_area)
            res['NOMOR_BPU'] = nomor_match.group(1) if (nomor_match and nomor_match.group(1) != "INDONESIA") else "N/A"
            res['STATUS'] = "PEMBETULAN" if "PEMBETULAN" in clean_text.upper() else "NORMAL"

            # --- 2. BAGIAN A (Penerima) ---
            # A.1 NPWP
            a1 = re.search(r'A\.1.*?[:\s]*(\d{15,16})', clean_text)
            res['A.1_NPWP_NIK_PENERIMA'] = "'" + a1.group(1) if a1 else "N/A"
            
            # A.2 NAMA
            a2 = re.search(r'A\.2\s+NAMA\s*[:\s]*(.*?)\s+A\.3', clean_text, re.I | re.S)
            res["A.2_NAMA_PENERIMA"] = a2.group(1).strip() if a2 else "N/A"
            
            # A.3 NITKU (Ambil digitnya saja sebelum tanda hubung)
            a3 = re.search(r'NITKU\)\s*[:\s]*(\d+)', clean_text)
            res["A.3_NITKU_PENERIMA"] = "'" + a3.group(1) if a3 else "N/A"

            # --- 3. BAGIAN B (Transaksi) ---
            res["B.2_JENIS_PPH"] = "Pasal 23" if "Pasal 23" in clean_text else "N/A"
            b3 = re.search(r'(\d{2}-\d{3}-\d{2})', clean_text)
            res['B.3_KODE_OBJEK_PAJAK'] = b3.group(1) if b3 else "N/A"

            # B.5 DPP & B.7 PPh
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
            c_clean = re.sub(r' +', ' ', c_part)
            
            # C.1 NPWP PEMOTONG (Ambil 15-16 digit angka)
            c1 = re.search(r'C\.1.*?[:\s]*(\d{15,16})', c_clean)
            res['C.1_NPWP_PEMOTONG'] = "'" + c1.group(1) if c1 else "N/A"
            
            # C.2 NITKU PEMOTONG
            c2 = re.search(r'NITKU\)\s*[:\s]*(\d+)', c_clean)
            res['C.2_NITKU_PEMOTONG'] = "'" + c2.group(1) if c2 else "N/A"
            
            # C.3 NAMA PEMOTONG (Pembersihan Lanjutan)
            c3_raw = re.search(r'PPh\s*[:\s]*(.*?)\s*C\.4', c_clean, re.I | re.S)
            res["C.3_NAMA_PEMOTONG"] = clean_id_prefix(c3_raw.group(1)) if c3_raw else "N/A"
            
            # C.4 TANGGAL BPU
            c4 = re.search(r'C\.4\s+TANGGAL\s*[:\s]*(\d{2}\s\w+\s\d{4})', c_clean)
            res["C.4_TANGGAL_BPU"] = c4.group(1) if c4 else "N/A"
            
            # C.5 NAMA PENANDATANGAN
            c5 = re.search(r'C\.5\s+NAMA\s+PENANDATANGAN\s*[:\s]*(.*?)\s*(?:C\.6|$)', c_clean, re.S)
            res["C.5_NAMA_PENANDATANGAN"] = c5.group(1).strip() if c5 else "N/A"

            res["NAMA_FILE"] = pdf_file.name
            return res
    except Exception as e:
        return {"NOMOR_BPU": f"ERROR: {str(e)}", "NAMA_FILE": pdf_file.name}

# --- APLIKASI STREAMLIT ---
if 'df_result' not in st.session_state:
    st.session_state.df_result = None

uploaded_files = st.file_uploader("Upload PDF Bukti Potong", type="pdf", accept_multiple_files=True)

if uploaded_files:
    if st.button("🚀 Jalankan Ekstraksi Versi 15"):
        with st.spinner('Mengekstrak data lengkap A.1 - C.5...'):
            all_results = [extract_surgical_v15(f) for f in uploaded_files]
            st.session_state.df_result = pd.DataFrame(all_results)
            st.success("Ekstraksi selesai!")

if st.session_state.df_result is not None:
    st.dataframe(st.session_state.df_result)
    output = BytesIO()
    with pd.ExcelWriter(output, engine='xlsxwriter') as writer:
        st.session_state.df_result.to_excel(writer, index=False)
    st.download_button(label="📥 Download Hasil Rekap (Excel)", data=output.getvalue(), file_name="Rekap_Pajak_Elitery_v15.xlsx")
