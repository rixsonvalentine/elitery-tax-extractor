import streamlit as st
import pdfplumber
import re
import pandas as pd
from io import BytesIO

st.set_page_config(page_title="Elitery Tax Extractor v14", page_icon="🚀")
st.title("🚀 Elitery Tax Extractor (Hyper-Precision v14)")

def extract_surgical_v14(pdf_file):
    try:
        with pdfplumber.open(pdf_file) as pdf:
            full_text = ""
            for p in pdf.pages:
                t = p.extract_text()
                if t: full_text += t + "\n"
            
            if not full_text.strip():
                return {"NOMOR_BPU": "ERROR: Scan/Kosong", "NAMA_FILE": pdf_file.name}

            # Normalisasi teks: Hapus spasi ganda dan rapikan baris
            clean_text = re.sub(r' +', ' ', full_text)
            res = {}

            # --- 1. NOMOR BPU & MASA PAJAK ---
            masa_match = re.search(r'(\d{2}-202\d)', clean_text)
            res['MASA_PAJAK'] = masa_match.group(1) if masa_match else "N/A"
            
            # Mencari kode 9 digit (huruf/angka) yang berdekatan dengan masa pajak
            # Pola: Mencari teks alfanumerik 9 digit sebelum pola tanggal
            nomor_search = re.search(r'([A-Z0-9]{9})\s+(\d{2}-202\d)', clean_text)
            if nomor_search:
                res['NOMOR_BPU'] = nomor_search.group(1)
            else:
                # Fallback: Cari pola 9 digit alfanumerik di area atas
                fallback_nomor = re.findall(r'\b[A-Z0-9]{9}\b', clean_text[:500])
                res['NOMOR_BPU'] = next((n for n in fallback_nomor if any(c.isdigit() for c in n) and n != "INDONESIA"), "N/A")

            res['STATUS'] = "PEMBETULAN" if "PEMBETULAN" in clean_text.upper() else "NORMAL"

            # --- 2. BAGIAN A (Penerima) ---
            # A.1 NPWP (Mencari 15-16 digit angka setelah label A.1)
            a1 = re.search(r'A\.1.*?[:\s]*(\d{15,16})', clean_text)
            res['A.1_NPWP_NIK_PENERIMA'] = "'" + a1.group(1) if a1 else "N/A"
            
            # A.2 NAMA
            a2 = re.search(r'A\.2\s+NAMA\s*[:\s]*(.*?)\s+A\.3', clean_text, re.I | re.S)
            res["A.2_NAMA_PENERIMA"] = a2.group(1).strip() if a2 else "N/A"
            
            # A.3 NITKU (Mencari digit panjang setelah label NITKU)
            a3 = re.search(r'NITKU\)\s*[:\s]*(\d{16,})', clean_text)
            res["A.3_NITKU_PENERIMA"] = "'" + a3.group(1) if a3 else "N/A"

            # --- 3. BAGIAN B (Transaksi) ---
            res["B.2_JENIS_PPH"] = "Pasal 23" if "Pasal 23" in clean_text else "N/A"
            
            # B.3 KODE OBJEK PAJAK
            b3 = re.search(r'(\d{2}-\d{3}-\d{2})', clean_text)
            res['B.3_KODE_OBJEK_PAJAK'] = b3.group(1) if b3 else "N/A"

            # B.5 DPP & B.7 PPh (Mencari angka format Indonesia 1.234.567)
            # Mengambil angka ribuan yang muncul di baris yang sama dengan kode objek pajak atau di bawahnya
            money_vals = re.findall(r'(\d{1,3}(?:\.\d{3})+)', clean_text)
            if len(money_vals) >= 2:
                res["B.5_DPP"] = money_vals[-2]
                res["B.7_PPH_DIPOTONG"] = money_vals[-1]
            else:
                res["B.5_DPP"] = "0"
                res["B.7_PPH_DIPOTONG"] = "0"

            # --- 4. DOKUMEN DASAR (B.9 - B.11) ---
            res["B.9_JENIS_DOKUMEN"] = re.search(r'Jenis Dokumen\s*[:\s]*(.*?)\s+Tanggal', clean_text, re.I).group(1).strip() if re.search(r'Jenis Dokumen\s*[:\s]*(.*?)\s+Tanggal', clean_text, re.I) else "N/A"
            doc_num = re.search(r'Nomor Dokumen\s*[:\s]*(\d+)', clean_text, re.I)
            res["B.10_NOMOR_DOKUMEN"] = "'" + doc_num.group(1) if doc_num else "N/A"
            res["B.11_TANGGAL_DOKUMEN"] = re.search(r'Tanggal\s*[:\s]*(\d{2}\s\w+\s\d{4})', clean_text, re.I).group(1) if re.search(r'Tanggal\s*[:\s]*(\d{2}\s\w+\s\d{4})', clean_text, re.I) else "N/A"

            # --- 5. BAGIAN C (Pemotong) ---
            c_block = full_text.split('C. IDENTITAS')[-1] if 'C. IDENTITAS' in full_text else ""
            c_clean = re.sub(r' +', ' ', c_block)
            
            # C.1 NPWP PEMOTONG
            c1 = re.search(r'C\.1.*?[:\s]*(\d{15,16})', c_clean)
            res['C.1_NPWP_PEMOTONG'] = "'" + c1.group(1) if c1 else "N/A"
            
            # C.2 NITKU PEMOTONG
            c2 = re.search(r'NITKU\)\s*[:\s]*(\d{16,})', c_clean)
            res['C.2_NITKU_PEMOTONG'] = "'" + c2.group(1) if c2 else "N/A"
            
            # C.3 NAMA PEMOTONG (Dibersihkan dari label PPh)
            c3 = re.search(r'PPh\s*[:\s]*(.*?)\s*C\.4', c_clean, re.I | re.S)
            if c3:
                nama_p = c3.group(1).strip().replace('\n', ' ')
                res["C.3_NAMA_PEMOTONG"] = re.sub(r'^[:\s]+', '', nama_p)
            else:
                res["C.3_NAMA_PEMOTONG"] = "N/A"
            
            # C.4 TANGGAL BPU
            c4 = re.search(r'C\.4\s+TANGGAL\s*[:\s]*(\d{2}\s\w+\s\d{4})', c_clean)
            res["C.4_TANGGAL_BPU"] = c4.group(1) if c4 else "N/A"
            
            # C.5 NAMA PENANDATANGAN
            c5 = re.search(r'C\.5\s+NAMA\s+PENANDATANGAN\s*[:\s]*(.*?)\s*(?:C\.6|$)', c_clean, re.S)
            res["C.5_NAMA_PENANDATANGAN"] = c5.group(1).strip().replace('\n', ' ') if c5 else "N/A"

            res["NAMA_FILE"] = pdf_file.name
            return res
    except Exception as e:
        return {"NOMOR_BPU": f"ERROR: {str(e)}", "NAMA_FILE": pdf_file.name}

# --- APLIKASI STREAMLIT ---
if 'df_result' not in st.session_state:
    st.session_state.df_result = None

uploaded_files = st.file_uploader("Upload PDF Bukti Potong", type="pdf", accept_multiple_files=True)

if uploaded_files:
    if st.button("🚀 Jalankan Ekstraksi Versi 14"):
        with st.spinner('Memproses dokumen...'):
            all_results = [extract_surgical_v14(f) for f in uploaded_files]
            st.session_state.df_result = pd.DataFrame(all_results)
            st.success(f"Berhasil! {len(all_results)} file telah direkap.")

if st.session_state.df_result is not None:
    st.dataframe(st.session_state.df_result)
    
    output = BytesIO()
    with pd.ExcelWriter(output, engine='xlsxwriter') as writer:
        st.session_state.df_result.to_excel(writer, index=False)
    
    st.download_button(
        label="📥 Download Hasil Rekap (Excel)",
        data=output.getvalue(),
        file_name="Rekap_Pajak_Elitery_v14.xlsx",
        mime="application/vnd.ms-excel"
    )
