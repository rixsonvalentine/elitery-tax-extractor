import streamlit as st
import pdfplumber
import re
import pandas as pd
from io import BytesIO

st.set_page_config(page_title="Elitery Tax Extractor v13", page_icon="🛡️")
st.title("🛡️ Elitery Tax Extractor (Final Stable v13)")

# --- FUNGSI EKSTRAKSI DATA ---
def extract_surgical_v13(pdf_file):
    try:
        with pdfplumber.open(pdf_file) as pdf:
            full_text = ""
            for p in pdf.pages:
                t = p.extract_text()
                if t: full_text += t + "\n"
            
            if not full_text.strip():
                return {"NOMOR_BPU": "ERROR: Scan/Kosong", "NAMA_FILE": pdf_file.name}

            # Normalisasi teks: hapus spasi berlebih
            clean_text = re.sub(r' +', ' ', full_text)
            res = {}

            # 1. NOMOR & MASA PAJAK [cite: 4]
            # Mencari 9 digit alfanumerik setelah kata 'PAJAK' atau di area header
            masa_match = re.search(r'(\d{2}-202\d)', clean_text)
            res['MASA_PAJAK'] = masa_match.group(1) if masa_match else "N/A"
            
            nomor_search = re.search(r'PAJAK.*?\b([A-Z0-9]{9})\b', clean_text, re.S)
            res['NOMOR_BPU'] = nomor_search.group(1) if nomor_search else "N/A"
            res['STATUS'] = "PEMBETULAN" if "PEMBETULAN" in clean_text.upper() else "NORMAL"

            # 2. BAGIAN A (Penerima - Elitery) 
            # A.1 NPWP: Mengambil digit setelah titik dua (:)
            res['A.1_NPWP_NIK_PENERIMA'] = re.search(r'A\.1.*?[:\s]*(\d{15,16})', clean_text).group(1) if re.search(r'A\.1.*?[:\s]*(\d{15,16})', clean_text) else "N/A"
            res["A.2_NAMA_PENERIMA"] = re.search(r'A\.2\s+NAMA\s*[:\s]*(.*?)\s+A\.3', clean_text, re.I).group(1).strip() if re.search(r'A\.2\s+NAMA\s*[:\s]*(.*?)\s+A\.3', clean_text, re.I) else "N/A"
            res["A.3_NITKU_PENERIMA"] = re.search(r'NITKU\)\s*[:\s]*(\d{10,})', clean_text).group(1) if re.search(r'NITKU\)\s*[:\s]*(\d{10,})', clean_text) else "N/A"

            # 3. BAGIAN B (Transaksi) [cite: 15, 20, 24]
            res["B.2_JENIS_PPH"] = "Pasal 23" if "Pasal 23" in clean_text else "N/A"
            res['B.3_KODE_OBJEK_PAJAK'] = re.search(r'(\d{2}-\d{3}-\d{2})', clean_text).group(1) if re.search(r'(\d{2}-\d{3}-\d{2})', clean_text) else "N/A"

            # B.5 DPP & B.7 PPh (Menangani angka yang menempel pada teks)
            money_vals = re.findall(r'(\d{1,3}(?:\.\d{3})+)', clean_text)
            if len(money_vals) >= 2:
                res["B.5_DPP"] = money_vals[-2]
                res["B.7_PPH_DIPOTONG"] = money_vals[-1]
            else:
                res["B.5_DPP"] = "0"
                res["B.7_PPH_DIPOTONG"] = "0"

            # 4. DOKUMEN DASAR [cite: 28, 30, 36]
            res["B.9_JENIS_DOKUMEN"] = re.search(r'Jenis Dokumen\s*[:\s]*(.*?)\s+Tanggal', clean_text, re.I).group(1).strip() if re.search(r'Jenis Dokumen\s*[:\s]*(.*?)\s+Tanggal', clean_text, re.I) else "N/A"
            res["B.10_NOMOR_DOKUMEN"] = "'" + re.search(r'Nomor Dokumen\s*[:\s]*(\d+)', clean_text, re.I).group(1) if re.search(r'Nomor Dokumen\s*[:\s]*(\d+)', clean_text, re.I) else "N/A"
            res["B.11_TANGGAL_DOKUMEN"] = re.search(r'Tanggal\s*[:\s]*(\d{2}\s\w+\s\d{4})', clean_text, re.I).group(1) if re.search(r'Tanggal\s*[:\s]*(\d{2}\s\w+\s\d{4})', clean_text, re.I) else "N/A"

            # 5. BAGIAN C (Pemotong) 
            c_block = full_text.split('C. IDENTITAS')[-1] if 'C. IDENTITAS' in full_text else ""
            res['C.1_NPWP_PEMOTONG'] = re.search(r'C\.1.*?[:\s]*(\d{15,16})', c_block).group(1) if re.search(r'C\.1.*?[:\s]*(\d{15,16})', c_block) else "N/A"
            res['C.2_NITKU_PEMOTONG'] = re.search(r'NITKU\)\s*[:\s]*(\d{10,})', c_block).group(1) if re.search(r'NITKU\)\s*[:\s]*(\d{10,})', c_block) else "N/A"
            
            # C.3 Nama Pemotong: Cari teks setelah 'PPh:'
            c3_match = re.search(r'PPh\s*[:\s]*(.*?)\s*C\.4', c_block, re.I | re.S)
            res["C.3_NAMA_PEMOTONG"] = c3_match.group(1).strip() if c3_match else "N/A"
            
            res["C.4_TANGGAL_BPU"] = re.search(r'C\.4\s+TANGGAL\s*[:\s]*(\d{2}\s\w+\s\d{4})', c_block).group(1) if re.search(r'C\.4\s+TANGGAL\s*[:\s]*(\d{2}\s\w+\s\d{4})', c_block) else "N/A"
            res["C.5_NAMA_PENANDATANGAN"] = re.search(r'C\.5\s+NAMA\s+PENANDATANGAN\s*[:\s]*(.*?)\s*(?:C\.6|$)', c_block, re.S).group(1).strip() if re.search(r'C\.5\s+NAMA\s+PENANDATANGAN\s*[:\s]*(.*?)\s*(?:C\.6|$)', c_block, re.S) else "N/A"

            res["NAMA_FILE"] = pdf_file.name
            return res
    except Exception as e:
        return {"NOMOR_BPU": f"ERROR: {str(e)}", "NAMA_FILE": pdf_file.name}

# --- LOGIKA APLIKASI (STREAMLIT STATE) ---
if 'df_result' not in st.session_state:
    st.session_state.df_result = None

uploaded_files = st.file_uploader("Upload PDF Bukti Potong", type="pdf", accept_multiple_files=True)

if uploaded_files:
    if st.button("🚀 Jalankan Ekstraksi Versi 13"):
        with st.spinner('Sedang memproses...'):
            all_results = [extract_surgical_v13(f) for f in uploaded_files]
            st.session_state.df_result = pd.DataFrame(all_results)
            st.success(f"Berhasil memproses {len(all_results)} file!")

if st.session_state.df_result is not None:
    st.dataframe(st.session_state.df_result)
    
    # Tombol Download menggunakan data dari session state
    output = BytesIO()
    with pd.ExcelWriter(output, engine='xlsxwriter') as writer:
        st.session_state.df_result.to_excel(writer, index=False)
    
    st.download_button(
        label="📥 Download Rekap Excel",
        data=output.getvalue(),
        file_name="Rekap_Pajak_Elitery_v13.xlsx",
        mime="application/vnd.ms-excel"
    )
