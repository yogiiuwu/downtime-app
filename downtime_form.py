
import streamlit as st
from datetime import datetime
from openpyxl import load_workbook
from openpyxl.comments import Comment
from openpyxl.utils import get_column_letter
from openpyxl.cell.cell import MergedCell
from openpyxl.styles import PatternFill
from copy import copy
from PIL import Image
import tempfile
import re
import gspread
import json
import ast
from oauth2client.service_account import ServiceAccountCredentials
import hashlib
import pandas as pd
import shutil
import sqlite3

# === INISIALISASI DATABASE SQLITE ===
def init_db():
    conn = sqlite3.connect('downtime.db')
    c = conn.cursor()
    c.execute('''
    CREATE TABLE IF NOT EXISTS downtime (
        id INTEGER PRIMARY KEY AUTOINCREMENT,
        tanggal_input TEXT,
        line_produksi TEXT,
        nama_produk TEXT,
        kode_produk TEXT,
        lot TEXT,
        tanggal_produksi TEXT,
        jenis TEXT,
        jam TEXT,
        durasi INTEGER,
        komentar TEXT
    )
    ''')
    conn.commit()
    conn.close()

# Panggil inisialisasi sekali saat aplikasi dijalankan
init_db()

# LOGIN SECTION
def hash_password(password):
    return hashlib.sha256(password.encode()).hexdigest()

users = {
    "admin": hash_password("admin"),
    "yogi": hash_password("8081"),
    "arfian": hash_password("2178"),
    "cakrahayu": hash_password("cakrahayu2003"),
    "herawati": hash_password("herawati"),
    "rokhim": hash_password("2090"),
    "sheva": hash_password("2175")
}

def check_login(username, password):
    return username in users and users[username] == hash_password(password)

if "logged_in" not in st.session_state:
    st.session_state.logged_in = False
# Tambahkan halaman reset password
def reset_password():
    st.title("🔑 Reset Password")

    username = st.text_input("Masukkan Username")

    if st.button("Cek Username"):
        if username in users:
            st.session_state.reset_user = username
            st.session_state.step_reset = "new_password"
            st.rerun()
        else:
            st.error("Username tidak ditemukan!")

# Form input password baru
def input_password_baru():
    st.title("🔑 Buat Password Baru")

    new_password = st.text_input("Masukkan Password Baru", type="password")
    confirm_password = st.text_input("Ulangi Password Baru", type="password")

    if st.button("Simpan Password Baru"):
        if new_password != confirm_password:
            st.error("Password tidak sama, ulangi lagi.")
        elif new_password == "":
            st.error("Password tidak boleh kosong.")
        else:
            # Simpan password baru (di-hash)
            users[st.session_state.reset_user] = hash_password(new_password)
            st.success("Password berhasil diubah, silakan login.")
            st.session_state.reset_user = None
            st.session_state.step_reset = None
            st.rerun()

if "step_reset" not in st.session_state:
    st.session_state.step_reset = None

if st.session_state.step_reset == "reset":
    reset_password()
    st.stop()

if st.session_state.step_reset == "new_password":
    input_password_baru()
    st.stop()

if not st.session_state.logged_in:
    st.title("🔐 Login")
    username = st.text_input("Username")
    password = st.text_input("Password", type="password")
    col_login, col_lupa = st.columns(2)
    with col_login:
        if st.button("Login"):
            if check_login(username, password):
                st.session_state.logged_in = True
                st.session_state.username = username
                st.success("Login berhasil!")
                st.rerun()
            else:
                st.error("Username atau password salah")

    with col_lupa:
        if st.button("Lupa Password?"):
            st.session_state.step_reset = "reset"
            st.rerun()

    st.stop()

# SIMPAN KE SQLite
def simpan_downtime_ke_sqlite(metadata, entry):
    conn = sqlite3.connect('downtime.db')
    c = conn.cursor()
    c.execute('''
        INSERT INTO downtime (
            tanggal_input, line_produksi, nama_produk, kode_produk,
            lot, tanggal_produksi, jenis, jam, durasi, komentar
        ) VALUES (?, ?, ?, ?, ?, ?, ?, ?, ?, ?)
    ''', (
        str(datetime.now()),
        metadata["line_produksi"],
        metadata["nama_produk"],
        metadata["kode_produk"],
        metadata["lot"],
        str(metadata["tanggal_produksi"]),
        entry["jenis"],
        entry["jam"],
        entry["durasi"],
        entry["komentar"]
    ))
    conn.commit()
    conn.close()

# GOOGLE SHEETS FINAL SUPER STABIL FIX
def get_google_sheet(sheet_name):
    scope = ["https://spreadsheets.google.com/feeds", "https://www.googleapis.com/auth/drive"]
    creds_dict = json.loads(json.dumps(dict(st.secrets["gcp_service_account"])))  # FIXED FINAL VERSION
    creds = ServiceAccountCredentials.from_json_keyfile_dict(creds_dict, scope)
    client = gspread.authorize(creds)
    spreadsheet = client.open(sheet_name)
    return spreadsheet

# LOGIN SECTION
def hash_password(password):
    return hashlib.sha256(password.encode()).hexdigest()

users = {
    "admin": hash_password("admin"),
    "yogi": hash_password("8081"),
    "arfian": hash_password("2178"),
    "cakrahayu": hash_password("cakrahayu2003"),
    "herawati": hash_password("herawati"),
    "rokhim": hash_password("2090"),
    "sheva": hash_password("2175")
}

def check_login(username, password):
    return username in users and users[username] == hash_password(password)

if "logged_in" not in st.session_state:
    st.session_state.logged_in = False

if not st.session_state.logged_in:
    st.title("🔐 Login")
    username = st.text_input("Username")
    password = st.text_input("Password", type="password")
    if st.button("Login"):
        if check_login(username, password):
            st.session_state.logged_in = True
            st.success("Login berhasil!")
            st.rerun()
        else:
            st.error("Username atau password salah")
    st.stop()


def normalize(text):
    return re.sub(r"\s+", " ", str(text).strip()).lower()
def find_or_create_lot_block(ws, lot, template_start_row=8, template_height=18):
    max_row = ws.max_row
    lot_column = "F"

    # === CARI APAKAH LOT SUDAH ADA ===
    for row in range(template_start_row, max_row + 1, template_height + 2):
        lot_cell = ws[f"{lot_column}{row + 2}"]
        if lot_cell.value and str(lot_cell.value).strip().lower() == lot.lower():
            return row

    # === CARI BLOK KOSONG ===
    for row in range(template_start_row, max_row + 1, template_height + 2):
        if not (ws.cell(row=row, column=3).value or
                ws.cell(row=row + 1, column=3).value or
                ws.cell(row=row + 2, column=3).value):
            new_start = row
            break
    else:
        # === SALIN TEMPLATE BARU ===
        new_start = max_row + 2
        for i in range(template_height):
            src_row = template_start_row + i
            dst_row = new_start + i
            for col in range(1, ws.max_column + 1):
                src_cell = ws.cell(row=src_row, column=col)
                dst_cell = ws.cell(row=dst_row, column=col)
                if not isinstance(src_cell, MergedCell):
                    dst_cell.value = src_cell.value
                if src_cell.has_style:
                    dst_cell._style = copy(src_cell._style)
                if src_cell.fill != PatternFill():
                    dst_cell.fill = copy(src_cell.fill)
                if src_cell.comment:
                    dst_cell.comment = Comment(src_cell.comment.text, "User")

        # SALIN MERGE TEMPLATE
        offset = new_start - template_start_row
        for rng in list(ws.merged_cells.ranges):
            if rng.min_row >= template_start_row and rng.max_row <= template_start_row + template_height:
                ws.merge_cells(start_row=rng.min_row + offset,
                               end_row=rng.max_row + offset,
                               start_column=rng.min_col,
                               end_column=rng.max_col)

    # === NOMOR KOLOM A UNTUK DOWNTIME ===
    for i in range(11):
        cell = ws.cell(row=new_start + 5 + i, column=1)
        if not isinstance(cell, MergedCell):
            cell.value = i + 1

    # === FORMULA TOTAL ===
    ws[f"AC{new_start + 16}"] = f"=SUM(AC{new_start + 5}:AC{new_start + 15})"
    ws[f"AD{new_start + 16}"] = f"=SUM(AD{new_start + 5}:AD{new_start + 15})"

    # === HITUNG BLOK DENGAN RANGE MERGE DI KOLOM A ===
    blok_nomor = 1
    for rng in ws.merged_cells.ranges:
        if rng.min_col == 1 and rng.max_col == 1:
            if (rng.max_row - rng.min_row + 1) == template_height:
                if rng.max_row < new_start:
                    blok_nomor += 1

    # === MERGE & ISI NOMOR BLOK ===
    ws.merge_cells(start_row=new_start, end_row=new_start + template_height - 1,
                   start_column=1, end_column=1)
    ws.cell(row=new_start, column=1).value = blok_nomor

    return new_start

# Simpan ke Excel 
def simpan_downtime_ke_excel(template_path, metadata, entry):
    wb = load_workbook(template_path)
    sheet_name = metadata["line_produksi"]
    if sheet_name not in wb.sheetnames:
        st.error(f"❌ Sheet '{sheet_name}' tidak ditemukan.")
        return

    ws = wb[sheet_name]
    blok_awal = find_or_create_lot_block(ws, metadata["lot"])

    ws[f"C{blok_awal}"] = metadata["nama_produk"]
    ws[f"C{blok_awal + 1}"] = metadata["kode_produk"]
    ws[f"C{blok_awal + 2}"] = metadata["lot"]
    ws[f"C{blok_awal + 3}"] = str(metadata["tanggal_produksi"])

    jenis_downtime = normalize(entry["jenis"])
    jam_index = int(str(entry["jam"]).split(":")[0]) + 5
    durasi = float(entry["durasi"])
    komentar = entry.get("komentar", "")
    found = False

    for i in range(5, 16):
        row = blok_awal + i
        cell_value = ws.cell(row=row, column=4).value
        if normalize(cell_value) == jenis_downtime:
            cell = ws.cell(row=row, column=jam_index)
            cell.value = durasi  # Set, not add
            if komentar:
                cell.comment = Comment(komentar, "User")
            total_menit = sum(float(ws.cell(row=row, column=col).value or 0) for col in range(5, 29))
            ws.cell(row=row, column=29).value = total_menit
            ws.cell(row=row, column=30).value = round(total_menit / 60, 2)
            found = True
            break

    if not found:
        st.error(f"❌ Jenis downtime '{entry['jenis']}' tidak ditemukan pada LOT ini.")
       
    wb.save(template_path)





# ======================= STREAMLIT APP ========================

st.set_page_config(page_title="DOWNTIME SOFTBAG II", layout="wide")

# Tambahkan tombol logout
col_logout, col_title = st.columns([1, 9])  # membuat 2 kolom: tombol & judul

with col_logout:
    if st.button("🔒 Logout"):
        st.session_state.logged_in = False
        st.session_state.username = None
        st.rerun()

with col_title:
    st.title(f"Form Input Downtime Packing (User: {st.session_state.username})")

if "excel_path" not in st.session_state:
    tmp = tempfile.NamedTemporaryFile(delete=False, suffix=".xlsx")
    with open("template_downtime_multi.xlsx", "rb") as fsrc:
        tmp.write(fsrc.read())
    tmp.close()
    st.session_state.excel_path = tmp.name

if "updated_excel" not in st.session_state:
    with open(st.session_state.excel_path, "rb") as f:
        st.session_state.updated_excel = f.read()

if "history_downtime" not in st.session_state:
    st.session_state.history_downtime = []

logo = Image.open("otsuka_logo.png")
col1, col2 = st.columns([1, 10])
with col1:
    st.image(logo, width=60)
with col2:
    st.markdown("## Form Input Downtime Packing")

line_options = [
    "DT ALT A", "DT ALT B", "DT Autocase A", "DT Autocase B",
    "DT Carton Erector", "DT Carton Sealing A", "DT Carton Sealing B"
]
line_produksi = st.selectbox("Pilih Jenis Mesin", line_options)

downtime_mapping = {
    "DT ALT A": {
        "Utility Downtime": ["Pressure Air Drop", "Listrik Padam"],
        "Proses Downtime": ["Setting Mesin", "Posisi Produk Abnormal"],
        "Mesin ALT Downtime": [
            "Parts Electrical / Control Error", "Parts Mechanical Error",
            "Parts Pneumatic Error", "Robot Spider", "No Operator"
        ]
    },
    "DT ALT B": {
        "Utility Downtime": ["Pressure Air Drop", "Listrik Padam"],
        "Proses Downtime": ["Setting Mesin", "Posisi Produk Abnormal"],
        "Mesin ALT Downtime": [
            "Parts Electrical / Control Error", "Parts Mechanical Error",
            "Parts Pneumatic Error", "Robot Spider", "No Operator"
        ]
    },
    "DT Autocase A": {
        "Utility Downtime": ["Pressure Air Drop", "Listrik Padam"],
        "Proses Downtime": ["Inspeksi Proses", "Material Habis", "Ganti Ribbon/Label"],
        "Mesin ALT Downtime": [
            "Parts Electrical / Control Error", "Parts Mechanical Error",
            "Parts Pneumatic Error", "Conveyor Transfer Trouble", "Motor Trouble", "No Operator"
        ]
    },
    "DT Autocase B": {
        "Utility Downtime": ["Pressure Air Drop", "Listrik Padam"],
        "Proses Downtime": ["Inspeksi Proses", "Material Habis", "Ganti Ribbon/Label"],
        "Mesin ALT Downtime": [
            "Parts Electrical / Control Error", "Parts Mechanical Error",
            "Parts Pneumatic Error", "Conveyor Transfer Trouble", "Motor Trouble", "No Operator"
        ]
    },
    "DT Carton Erector": {
        "Utility Downtime": ["Pressure Air Drop", "Listrik Padam", "Nitrogen Supply"],
        "Proses Downtime": ["Menunggu Produk Sterilisasi", "Inspeksi Proses", "Material Habis", "Penggantian Material"],
        "Mesin Weight Checker Downtime": [
            "Parts Electrical / Control Error", "Parts Mechanical Error",
            "Parts Pneumatic Error", "Conveyor Error", "Motor Error", "No Operator"
        ]
    },
    "DT Carton Sealing A": {
        "Utility Downtime": ["Pressure Air Drop", "Listrik Padam"],
        "Proses Downtime": ["Menunggu Produk Sterilisasi", "Inspeksi Proses", "Material Habis"],
        "Mesin Weight Checker Downtime": [
            "Parts Electrical / Control Error", "Parts Mechanical Error",
            "Parts Pneumatic Error", "Conveyor Error", "Motor Error", "No Operator"
        ]
    },
    "DT Carton Sealing B": {
        "Utility Downtime": ["Pressure Air Drop", "Listrik Padam"],
        "Proses Downtime": ["Menunggu Produk Sterilisasi", "Inspeksi Proses", "Material Habis"],
        "Mesin Weight Checker Downtime": [
            "Parts Electrical / Control Error", "Parts Mechanical Error",
            "Parts Pneumatic Error", "Conveyor Error", "Motor Error", "No Operator"
        ]
    }
}

# Fungsi untuk simpan ke Google Sheet
def simpan_downtime_ke_sheet(sheet, metadata, entry):
    sheet.append_row([
        str(datetime.now()),  
        # Timestamp
        metadata["line_produksi"],
        metadata["nama_produk"],
        metadata["kode_produk"],
        metadata["lot"],
        str(metadata["tanggal_produksi"]),
        entry["jenis"],
        entry["jam"],
        entry["durasi"],
        entry["komentar"]
    ])

# ====== FORM STREAMLIT ======
st.subheader("📦 Data Produk")
col1, col2 = st.columns(2)
with col1:
    nama_produk = st.text_input("Nama Produk")
    kode_produk = st.text_input("Kode Produk")
with col2:
    lot = st.text_input("Kode LOT")
    tgl_produksi = st.date_input("Tanggal Produksi", value=datetime.today())

st.subheader("⏱️ Input Downtime")
opsi = downtime_mapping.get(line_produksi, {})
col1, col2 = st.columns(2)
with col1:
    kategori = st.selectbox("Kategori Downtime", list(opsi.keys()))
with col2:
    jenis = st.selectbox("Jenis Downtime", opsi.get(kategori, []))
col3, col4 = st.columns(2)
with col3:
    jam = st.selectbox("Jam Terjadi", [f"{j:02d}:00" for j in range(24)])
with col4:
    durasi = st.number_input("Durasi (menit)", min_value=1, max_value=60)
komentar = st.text_input("Komentar")

col_tombol1, col_tombol2 = st.columns(2)
with col_tombol1:
    tambah = st.button("➕ Tambahkan Downtime")

if tambah:
    if not nama_produk or not kode_produk or not lot:
        st.warning("⚠️ Harap isi semua data produk.")
    else:
        metadata = {
            "nama_produk": nama_produk,
            "kode_produk": kode_produk,
            "lot": lot,
            "tanggal_produksi": tgl_produksi,
            "line_produksi": line_produksi
        }
        entry = {
            "jenis": jenis,
            "jam": jam,
            "durasi": durasi,
            "komentar": komentar
        }

        # Simpan ke SQLite (selalu dilakukan)
        simpan_downtime_ke_sqlite(metadata, entry)

        # Simpan ke Excel lokal
        simpan_downtime_ke_excel(st.session_state.excel_path, metadata, entry)

        # Simpan ke Google Sheet (jika internet ada)
        try:
            gsheet = get_google_sheet("DATABASE")
            simpan_downtime_ke_sheet(gsheet.sheet1, metadata, entry)
            st.success("✅ Data berhasil disimpan ke Google Sheet!")
        except Exception as e:
            st.warning(f"Gagal simpan ke Google Sheet: {e}")

        # Update file Excel dan log
        with open(st.session_state.excel_path, "rb") as f:
            st.session_state.updated_excel = f.read()

        st.session_state.history_downtime.append(
            f"✅ {metadata['nama_produk']} (LOT: {metadata['lot']}) - {entry['durasi']} menit ditambahkan.")

with col_tombol2:
    with open(st.session_state.excel_path, "rb") as f:
        excel_bytes = f.read()
    filename = f"DT_{nama_produk.replace(' ', '_').upper()}_{lot.upper()}.xlsx"
    if st.download_button("📥 Download Excel", data=excel_bytes, file_name=filename,
                          mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"):
        st.session_state.history_downtime.clear()
        tmp = tempfile.NamedTemporaryFile(delete=False, suffix=".xlsx")
        with open("template_downtime_multi.xlsx", "rb") as fsrc:
            tmp.write(fsrc.read())
        tmp.close()
        st.session_state.excel_path = tmp.name

if st.session_state.history_downtime:
    st.subheader("📋 Riwayat Downtime")
    for msg in st.session_state.history_downtime:
        st.success(msg)
        # ================= TAMBAHAN MENU DOWNLOAD =================
st.subheader("📥 Download Data Downtime Per Bulan")

try:
    gsheet = get_google_sheet("DATABASE")
    rows = gsheet.sheet1.get_all_records()
    df = pd.DataFrame(rows)
    
    if df.empty:
        st.warning("❌ Data Google Sheet masih kosong!")
    else:
        # Konversi tanggal produksi ke datetime
        df["Tanggal Produksi"] = pd.to_datetime(df["Tanggal Produksi"], errors='coerce')
        df = df.dropna(subset=["Tanggal Produksi"])  # Drop jika ada error parsing tanggal
        df["Bulan-Tahun"] = df["Tanggal Produksi"].dt.strftime('%Y-%m')

        bulan_options = sorted(df["Bulan-Tahun"].unique())
        selected_bulan = st.selectbox("Pilih Bulan Produksi:", bulan_options)

        if st.button("📥 Download Excel Bulanan"):
            filtered_df = df[df["Bulan-Tahun"] == selected_bulan]
            
            # Buat file excel baru dari template
            tmp_file = tempfile.NamedTemporaryFile(delete=False, suffix=".xlsx")
            shutil.copy("template_downtime_multi.xlsx", tmp_file.name)
            tmp_file.close()

            # Proses isi data per baris
            for idx, row in filtered_df.iterrows():
                metadata = {
                    "nama_produk": row["Nama Produk"],
                    "kode_produk": row["Kode Produk"],
                    "lot": row["Lot"],
                    "tanggal_produksi": row["Tanggal Produksi"].date(),
                    "line_produksi": row["Line Produksi"]
                }
                entry = {
                    "jenis": row["Jenis"],
                    "jam": row["Jam"],
                    "durasi": row["Durasi"],
                    "komentar": row["Komentar"]
                }
                simpan_downtime_ke_excel(tmp_file.name, metadata, entry)
            
            with open(tmp_file.name, "rb") as f:
                excel_bytes = f.read()
            filename = f"Downtime_{selected_bulan}.xlsx"
            st.download_button("⬇️ Klik untuk Download", data=excel_bytes, file_name=filename,
                               mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet")
except Exception as e:
    st.error(f"❌ Gagal membaca Google Sheet: {e}")

if st.session_state.get("username") == "admin":
    if st.button("🔄 Reset Downtime Data (Admin Only)"):
        gsheet.sheet1.clear()
        st.success("✅ Semua data downtime berhasil direset oleh Admin!")
