import streamlit as st
import pandas as pd
import re
from io import StringIO, BytesIO

# Konfigurasi halaman
st.set_page_config(page_title="Laporan Keterlambatan", layout="wide")

st.title("Konversi Laporan Keterlambatan - TXT ke Excel")
st.caption("Halaman ini digunakan untuk memproses laporan keterlambatan karyawan.")

# Fungsi parsing txt
def parse_txt_to_dataframe(txt):
    known_departments = sorted([
        "DIE MAKING", "INJECT MOULDING", "JOIN MANUAL", "QC OUTGOING",
        "MASTER CARTON", "P. CONTROL", "PURCHASING", "PRODUCTION",
        "LAMINATING", "SECURITY", "STAMPING", "WELDING", "PAINTING",
        "ASSEMBLY", "QUALITY CONTROL", "MAINTENANCE", "LOGISTIC",
        "ENGINEERING", "PPIC", "HRD", "SORTIR 1", "WP ESATEC", "FOLD & GLUE",
        "FIN & ACC", "UV HOOK","VACUUM FORMING", "QC PROCESS", "PON HANDLE",
        "P. CONTROL", "HIGH FREQ", "RIGID BOX","SHEETER BOARD", "MOLD SETTER",
        "SHEETER PVC"
    ], key=len, reverse=True)

    cleaned_txt = txt.replace('\r', '').replace('\f', '')
    lines = cleaned_txt.strip().splitlines()

    data = []
    for line in lines:
        if not line.strip().startswith('³') or '³' not in line[1:]:
            continue
        cleaned_check = line.strip().strip('³')
        if ('NO' in cleaned_check and 'NIK' in cleaned_check and 'NAMA' in cleaned_check and 'BAGIAN' in cleaned_check and 'STATUS' in cleaned_check):
            continue
        parts = [p.strip() for p in cleaned_check.split('³')]
        if len(parts) < 4:
            continue
        first_col_parts = parts[0].split()
        if len(first_col_parts) < 2:
            continue
        no = first_col_parts[0]
        nik = first_col_parts[1]
        name_and_dept_str = " ".join(first_col_parts[2:]) if len(first_col_parts) > 2 else ""
        nama, bagian = "-", "-"
        found_dept = False
        if name_and_dept_str:
            for dept in known_departments:
                if name_and_dept_str.endswith(dept):
                    nama = name_and_dept_str[:-len(dept)].strip()
                    bagian = dept
                    found_dept = True
                    break
        if not found_dept and name_and_dept_str:
            words = name_and_dept_str.split()
            if len(words) > 1:
                bagian = words[-1]
                nama = " ".join(words[:-1])
            else:
                nama = name_and_dept_str
        try:
            no_int = int(no) if no.isdigit() else no
            row = [no_int, nik, nama, bagian] + parts[1:]
            data.append(row)
        except:
            row = [no, nik, nama, bagian] + parts[1:]
            data.append(row)

    if data:
        max_cols = max(len(row) for row in data)
        for i, row in enumerate(data):
            while len(row) < max_cols:
                row.append("")
        max_cols = min(max_cols, 16)
        for i, row in enumerate(data):
            data[i] = row[:16]
        base_columns = ['NO', 'NIK', 'NAMA', 'BAGIAN', 'STATUS', "T1", "T2", "T3", "CUTI", "ALPA",
                        "SAKIT", "IJIN TDK MASUK", "LIBUR RESMI", "DGN SURAT", "1/2 HARI", "TDK ABSEN"]
        columns = base_columns[:max_cols]
        return pd.DataFrame(data, columns=columns)
    else:
        return pd.DataFrame()

# Upload file
uploaded_file = st.file_uploader("Upload file .txt", type=["txt"])
 
if uploaded_file:
    stringio = StringIO(uploaded_file.getvalue().decode("latin-1"))
    raw_text = stringio.read()
    try:
        df = parse_txt_to_dataframe(raw_text)
        st.subheader("Data Tersusun:")
        st.dataframe(df)
        towrite = BytesIO()
        df.to_excel(towrite, index=False, sheet_name="Laporan")
        towrite.seek(0)
        st.download_button(
            label="📥 Unduh sebagai Excel",
            data=towrite,
            file_name="Laporan_Keterlambatan.xlsx",
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
        )
    except Exception as e:
        st.error(f"Terjadi kesalahan saat memproses file: {e}")
else:
    st.info("Silakan unggah file .txt terlebih dahulu.")
