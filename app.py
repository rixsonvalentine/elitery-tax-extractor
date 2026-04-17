import streamlit as st
import pdfplumber
import re
import pandas as pd
from io import BytesIO

st.set_page_config(page_title="Elitery Tax Extractor v18", page_icon="🎯")
st.title("🎯 Elitery Tax Extractor (The Finisher v18)")

def clean_final_name(text):
    """Membersihkan teks dari label-label standar untuk mendapatkan nama asli."""
    if not text or len(text.strip()) < 2: return "N/A"
    # Hapus label boilerplate
    blacklist = [
        r"NAMA PEMOTONG DAN/ATAU PEMUNGUT", 
        r"DAN/ATAU PEMUNGUT PPh", 
        r"PPh", r"NAMA", r":", r"^-", r"^\d{22}"
    ]
    cleaned = text
    for pattern in blacklist:
        cleaned = re.sub(rf'(?i){pattern}', '', cleaned).strip()
    
    # Jika masih ada tanda hubung di depan, hapus
    cleaned = re.sub(r'^[-\s]+', '', cleaned)
    return cleaned.strip() if cleaned else "N/A"

def extract_surgical_v18(pdf_file):
    try:
        with pdfplumber.open(pdf_file) as pdf:
            full_text = "\n".join([p.extract_text() for p in pdf.pages if p.extract_text()])
            if not full_text.strip():
                return {"NOMOR_BPU": "ERROR: Scan/Kosong", "NAMA_FILE": pdf_file.name}

            # Normalisasi teks (tetap pertahankan newline untuk pemisahan bagian)
            text_norm = re.sub(r' +', ' ', full_text)
            res = {}

            # --- 1. NOMOR BPU & MASA PAJAK ---
            masa = re.search(r'(\d{2}-202\d)', text_norm)
            res['MASA_PAJAK'] = masa.group(1) if masa else "N/A"
            
            # Cari Nomor BPU (Kode alfanumerik 9 digit di area header)
            potential_nums = re.findall(r'\b[A-Z0-9]{9}\b', text_norm[:600])
            res['NOMOR_BPU'] = next((n for n in potential_nums if any(c.isdigit() for c in n) and n not in ["UNIFIKASI", "INDONESIA"]), "N/A")
            res['STATUS'] = "PEMBETULAN" if "PEMBETULAN" in text_norm.upper() else "NORMAL"

            # --- 2. BAGIAN A (Penerima) ---
            # NPWP (A.1)
            a1 = re.search(r'A\.1.*?(\d{15,16})', text_norm, re.S)
            res['A.1_NPWP_NIK_PENERIMA'] = "'" + a1.group(1) if a1 else "N/A"
            
            # Nama (A.2)
            a2 = re.search(r'A\.2\s+NAMA\s*[:\s]*(.*?)(?:\n|A\.3|NPWP)', text_norm, re.I)
            res["A.2_NAMA_PENERIMA"] = a2.group(1).strip() if a2 else "N/A"
            
            # NITKU (A.3) - 22 digit
            a3 = re.search(r'A\.3.*?(\d{22})', text_norm, re.S)
            res["A.3_NITKU_PENERIMA"] = "'" + a3.group(1) if a3 else "N/A"

            # --- 3. BAGIAN B (Transaksi) ---
            res["B.2_JENIS_PPH"] = "Pasal 23" if "Pasal 23" in text_norm else "N/A"
            b3 = re.search(r'(\d{2}-\d{3}-\d{2})', text_norm)
            res['B.3_KODE_OBJEK_PAJAK'] = b3.group(1) if b3 else "N/A"

            # DPP (B.5) & PPh (B.7)
            money_vals = re.findall(r'(\d{1,3}(?:\.\d{3})+)', text_norm)
            if len(money_vals) >= 2:
                res["B.5_DPP"] = money_vals[-2]
                res["B.7_PPH_DIPOTONG"] = money_vals[-1]
            else:
                res["B.5_DPP"] = "0"
                res["B.7_PPH_DIPOTONG"] = "0"

            # --- 4. DOKUMEN DASAR (B.9 - B.11) ---
            # Menarik Nomor Dokumen (B.10) lebih fleksibel (bisa angka/huruf panjang)
            b9_match = re.search(r'Jenis Dokumen\s*[:\s]*(.*?)\s+Tanggal', text_norm, re.I)
            res["B.9_JENIS_DOKUMEN"] = b9_match.group(1).strip() if b9_match else "N/A"
            
            doc_num = re.search(r'Nomor Dokumen\s*[:\s]*([A-Z0-9]+)', text_norm, re.I)
            res["B.10_NOMOR_DOKUMEN"] = "'" + doc_num.group(1) if doc_num else "N/A"
            
            tgl_dok = re.search(r'Tanggal\s*[:\s]*(\d{2}\s\w+\s\d{4})', text_norm, re.I)
            res["B.11_TANGGAL_DOKUMEN"] = tgl_dok.group(1) if tgl_dok else "N/A"

            # --- 5. BAGIAN C (Pemotong) ---
            c_block = text_norm.split('C. IDENTITAS')[-1] if 'C. IDENTITAS' in text_norm else text_norm
            
            # C.1 NPWP
            c1 = re.search(r'C\.1.*?(\d{15,16})', c_block, re.S)
            res['C.1_NPWP_PEMOTONG'] = "'" + c1.group(1) if c1 else "N/A"
            
            # C.2 NITKU (22 digit)
            c2 = re.search(r'(\d{22})', c_block)
            res['C.2_NITKU_PEMOTONG'] = "'" + c2.group(1) if c2 else "N/A"
            
            # C.3 NAMA PEMOTONG (Ambil seluruh teks antara C.3 dan C.4 lalu bersihkan)
            c3_match = re.search(r'C\.3(.*?)C\.4', c_block, re.S | re.I)
            if c3_match:
                res["C.3_NAMA_PEMOTONG"] = clean_final_name(c3_match.group(1))
            else:
                res["C.3_NAMA_PEMOTONG"] = "N/A"
            
            res["C.4_TANGGAL_BPU"] = re.search(r'C\.4\s+TANGGAL\s*[:\s]*(\d{2}\s\w+\s\d{4})', c_block).group(1) if re.search(r'C\.4\s+TANGGAL\s*[:\s]*(\d{2}\s\w+\s\d{4})', c_block) else "N/A"
            res["C.5_NAMA_PENANDATANGAN"] = re.search(r'C\.5\s+NAMA\s+PENANDATANGAN\s*[:\s]*(.*?)(?:\n|C\.6|$)', c_block, re.I).group(1).strip() if re.search(r'C\.5\s+NAMA\s+PENANDATANGAN\s*[:\s]*(.*?)(?:\n|C\.6|$)', c_block, re.I) else "N/A"

            res["NAMA_FILE"] = pdf_file.name
            return res
    except Exception as e:
        return {"NOMOR_BPU": f"ERROR: {str(e)}", "NAMA_FILE": pdf_file.name}

# --- STREAMLIT UI ---
if 'df_result' not in st.session_state:
    st.session_state.df_result = None

uploaded_files = st.file_uploader("Upload PDF Bukti Potong", type="pdf", accept_multiple_files=True)

if uploaded_files:
    if st.button("🚀 Jalankan Ekstraksi Final v18"):
        with st.spinner('Memproses data...'):
            all_results = [extract_surgical_v18(f) for f in uploaded_files]
            st.session_state.df_result = pd.DataFrame(all_results)
            st.success("Ekstraksi Selesai!")

if st.session_state.df_result is not None:
    st.dataframe(st.session_state.df_result)
    output = BytesIO()
    with pd.ExcelWriter(output, engine='xlsxwriter') as writer:
        st.session_state.df_result.to_excel(writer, index=False)
    st.download_button(label="📥 Download Hasil Rekap", data=output.getvalue(), file_name="Rekap_Pajak_Elitery_v18.xlsx")
