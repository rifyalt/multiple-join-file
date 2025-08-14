import streamlit as st
import pandas as pd
import numpy as np
#import plotly.express as px
#import openpyxl
import os
from io import BytesIO
from datetime import datetime

# ====== CONFIG PAGE ======
st.set_page_config(
    page_title="Aplikasi Join Data | naratix",
    page_icon="🚀",
    layout="wide",
    initial_sidebar_state="expanded"
)

# --- main stub agar panggilan terjaga ---
def main():
    # Streamlit mengeksekusi script top-to-bottom, jadi fungsi ini bisa kosong.
    # Kalau nanti mau, pindahkan seluruh UI ke dalam fungsi ini.
    pass

# Placeholder agar .empty() tidak error bila dipanggil belakangan
import streamlit as st
progress_bar = st.empty()
status_text = st.empty()

# ====== STYLE MINIMAL ======
st.markdown("""
<style>
[data-testid="stSidebar"] {
    background-color: #61C9A8;
}
[data-testid="stSidebar"] h2 {
    color: #1e3a8a;
    font-weight: 700;
    font-size: 1.3rem;
}
.metric-card {
    background-color: #FFE19C;
    border-radius: 8px;
    padding: 1rem;
    border: 1px solid #FFE19C;
    box-shadow: 0 2px 8px rgba(0,0,0,0.04);
    text-align: center;
}
.metric-card:hover {
    transform: translateY(-2px);
    box-shadow: 0 4px 12px rgba(0,0,0,0.08);
}
</style>
""", unsafe_allow_html=True)

# ====== RUNNING TEXT ======
st.markdown("""
<style>
.running-banner {
    background: linear-gradient(90deg, #E5989B, #E5989B);
    color: white;
    font-size: 1.1rem;
    font-weight: bold;
    padding: 0.6rem 0;
    overflow: hidden;
    position: relative;
    border-radius: 8px;
    margin-bottom: 1rem;
    box-shadow: 0 3px 8px rgba(0,0,0,0.15);
}
.running-banner span {
    display: inline-block;
    padding-left: 100%;
    text-shadow: 1px 1px 2px rgba(0,0,0,0.3);
    animation: smooth-marquee 18s linear infinite;
    white-space: nowrap;
}
@keyframes smooth-marquee {
    0%   { transform: translateX(0); }
    100% { transform: translateX(-100%); }
}
</style>
<div class="running-banner">
    <span>🚀 Selamat datang di Naratix Data Integration Platform — Integrasi, Analisis, dan Visualisasi Data Lebih Cepat & Mudah! 📊</span>
</div>
""", unsafe_allow_html=True)


# ====== HEADER CORPORATE MODERN ======
st.markdown("""
<style>
.custom-header {
    background: #61C9A8;
    padding: 1.5rem 2rem;
    border-radius: 12px;
    box-shadow: 0 4px 15px rgba(0,0,0,0.1);
    color: white;
    margin-bottom: 1.5rem;
}
.custom-header h1 {
    font-size: 1.8rem;
    margin: 0;
    font-weight: 800;
}
.custom-header p {
    font-size: 1rem;
    opacity: 0.9;
    margin-top: 0.3rem;
}
</style>
<div class="custom-header">
    <h1>🚀 Naratix Data Integration Platform</h1>
    <p>Solusi Enterprise untuk Integrasi, Analisis, & Visualisasi Data</p>
</div>
""", unsafe_allow_html=True)

# ====== MAPPING KOLOM ======
rename_mapping = {
    "Invoice Date": "Tanggal Invoice",
    "Invoice Number": "Nomor Invoice Supplier",
    "Amount Payable": "Tagihan",
    "Agency Reference No": "Refrensi",
    "Web Inv No.": "Inv Web",
    "Booking ID": "Booking ID",
    "Description": "Deskripsi",
    "Billing Period": "Periode Tagihan",
    "Due Date": "Due Date",
    "Remark": "Keterangan",
    "Guest Name": "Nama Tamu",
    "Web User": "Nama Akun",
    "Booking Date": "Tanggal Reservasi",
    "Check In Date": "Check-In Date",
    "Check Out Date": "Check-Out Date",
    "Booking Status": "Status Reservasi",
    "Hotel Name": "Nama Hotel",
    "Room Name": "Tipe Kamar",
    "No Of Room": "Jumlah Kamar",
    "Hotel Address": "Alamat Hotel",
    "Manual Remark": "Keterangan Tambahan",
    "Agent Code": "Kode Agen",
    "Address": "Alamat",
    "Agent Name": "Nama Agen",
    "Credit Days": "Tagihan Berjalan",
    "Check In": "Check-In Date",
    "Check Out": "Check-Out Date",
    "OS Ref": "Refrensi",
    "Currency": "Mata Uang",
    "Outstanding Amount": "Tagihan",
    "Late Days": "Keterlambatan",
    "Aging": "Aging",
    "Note": "Catatan",
    "Platform": "Kode Agen",
    "Booking number": "Booking ID",
    "Source Refrence ID": "Refrensi",
    "Customer reference": "Refrensi Tamu",
    "Customer code": "Kode Tamu",
    "Customer Name": "Nama Agen",
    "Booking date": "Tanggal Reservasi",
    "Invoice date": "Tanggal Invoice",
    "Arrival date": "Check-In Date",
    "Departure date": "Check-Out Date",
    "Service Type": "Nama Hotel",
    "Number of passengers": "Jumlah Tamu",
    "Destination": "Kota Tujuan",
    "Hotel Country": "Negara",
    "Lead guest": "Nama Tamu",
    "Gross": "Tagihan",
    "Commission (amount)": "Komisi",
    "Net amount": "NTA",
    "Voucher Date": "Tanggal Invoice",
    "Debit( IDR )": "Debit",
    "Credit( IDR )": "Kredit",
    "OutStanding( IDR )": "Tagihan",
    "InvoiceDueDate": "Due Date",
    "Pax Name": "Nama Tamu",
    "Entity Name": "Nama Hotel",
    "No Of Nights": "Total Malam",
    "Agency LPO Number": "Refrensi",
    "lastCancellationDate": "Pembatalan",
    "Remarks": "Keterangan Tambahan",
    "Consultant Name": "PIC",
    "Agent Local Currency": "Mata Uang",
    "TBOH Confirmation No": "Booking ID",
    "PLACE OF SUPPLY": "Negara",
    "VAT AMOUNT (AED)": "PPN",
    "SUBUSER NAME": "Nama User",
    "BOOKED BY": "Pemesan",
    "PRODUCT TYPE": "Tipe Produk",
    "BOOKING MODE": "Nama Akun",
    "Pic Hotel":"PIC",
    "PIC Request":"PIC Tim Lain",
    "No BC":"No BC",
    "Rsv No":"Booking ID",
    "Issued Date":"Tanggal Invoice"
}
# normalize mapping (lowercase keys)
rename_mapping_norm = {k.lower().strip(): v for k, v in rename_mapping.items()}

# ====== HELPERS ======
def coalesce_duplicate_columns(df: pd.DataFrame) -> pd.DataFrame:
    df = df.copy()
    df.columns = [str(c) for c in df.columns]
    cols_series = pd.Series(df.columns)
    dup_cols = cols_series[cols_series.duplicated()].unique()
    if len(dup_cols) == 0:
        return df
    for col in dup_cols:
        same = [c for c in df.columns if c == col]
        combined = df[same].apply(
            lambda row: next((v for v in row if pd.notna(v) and str(v).strip() != ""), pd.NA),
            axis=1
        )
        df = df.drop(columns=same)
        df[col] = combined
    unique_cols = []
    for c in cols_series:
        if c not in unique_cols:
            unique_cols.append(c)
    final_cols = [c for c in unique_cols if c in df.columns]
    return df[final_cols]

def normalize_and_rename(df: pd.DataFrame) -> pd.DataFrame:
    df = df.copy()
    df.columns = [str(c).strip() for c in df.columns]
    new_cols = []
    for c in df.columns:
        key = c.lower().strip()
        new_cols.append(rename_mapping_norm.get(key, c))
    df.columns = new_cols
    df = coalesce_duplicate_columns(df)
    return df

# ===== Fungsi Bersihkan Data =====
def bersihkan_data(df):
    # Trim semua kolom
    df = df.applymap(lambda x: x.strip() if isinstance(x, str) else x)
    # Isi Invoice No kosong
    if "Invoice No" in df.columns:
        df["Invoice No"] = df["Invoice No"].fillna("blank")
        df["Invoice No"] = df["Invoice No"].replace("", "blank")
    return df

# ====== SIDEBAR CORPORATE ======
with st.sidebar:
    st.markdown("## 📂 Data Management")
    
    # File Upload Section
    st.markdown("### 📊 File Supplier")
    st.markdown("*Upload file data supplier untuk diproses*")
    file_utama_list = st.file_uploader(
        "Pilih File Excel Supplier", 
        type=["xlsx", "xls"], 
        accept_multiple_files=True,
        key="supplier_files",
        help="Support multiple files - Format: .xlsx, .xls"
    )
    
    st.markdown("### 🗃️ File Data Master")
    st.markdown("*Upload file data master untuk referensi*")
    file_data = st.file_uploader(
        "Pilih File Excel Data Master", 
        type=["xlsx", "xls"],
        accept_multiple_files=True,
        key="data_file",
        help="File referensi utama - Format: .xlsx, .xls"
    )

# ====== VALIDASI FILE ======
if not file_utama_list or not file_data:
    st.markdown("""
    <h3 style="margin-bottom:0.5rem;">📋 Spesifikasi Teknis</h3>
    <ul style="margin-top:0; padding-left:1.2rem; line-height:1.6;">
        <li><strong>Format Supported:</strong> Excel (.xlsx, .xls)</li>
        <li><strong>Required Columns:</strong> 'Booking ID' dan 'Invoice No' pada data master</li>
        <li><strong>Auto Mapping:</strong> 80+ kolom mapping otomatis</li>
        <li><strong>Data Cleansing:</strong> Automatic duplicate handling & normalization</li>
    </ul>

    <div style="
        background: linear-gradient(135deg, #eff6ff 0%, #dbeafe 100%);
        padding: 1rem;
        border-radius: 6px;
        margin-top: 1rem;
        border-left: 4px solid #3b82f6;
        font-size: 0.95rem;
        line-height: 1.5;
    ">
        <strong>💡 Pro Tip:</strong> Pastikan file data master mengandung kolom 
        <code>Booking ID</code> dan <code>Invoice No</code> untuk hasil mapping yang optimal.
    </div>
    """, unsafe_allow_html=True)
    st.stop()

# Progress bar dengan styling corporate
progress_container = st.container()
with progress_container:
    col1, col2 = st.columns([3, 1])
    with col1:
        progress_bar = st.progress(0)
    with col2:
        status_text = st.empty()

# ====== PROSES FILE UTAMA ======
status_text.markdown("**📊 Processing...**")
progress_bar.progress(25)

df_utama_list = []
main_cols_sets = []
file_status = []

for i, f in enumerate(file_utama_list):
    try:
        df = pd.read_excel(f)
        file_status.append(f"✅ {f.name}")
        #st.success(f"✅ File berhasil diproses: **{f.name}** ({len(df):,} records)")
    except Exception as e:
        file_status.append(f"❌ {f.name}")
        st.error(f"❌ Error memproses file **{f.name}**: {e}")
        st.stop()
    df = normalize_and_rename(df)
    df_utama_list.append(df)
    main_cols_sets.append(set(df.columns))

if len(main_cols_sets) > 0:
    common_main_cols = set.intersection(*main_cols_sets)
else:
    common_main_cols = set()

df_utama_all = pd.concat(df_utama_list, ignore_index=True, sort=False)
df_utama_all = coalesce_duplicate_columns(df_utama_all)
df_utama_all = df_utama_all.loc[:, ~df_utama_all.columns.duplicated()]

progress_bar.progress(50)
status_text.markdown("**🔄 Merging...**")

if "Booking ID" in df_utama_all.columns:
    agg_funcs = {}
    for col in df_utama_all.columns:
        if col == "Booking ID":
            continue
        low = col.lower().strip()
        if low in ["check-in date", "check in date", "arrival date"]:
            agg_funcs[col] = lambda x: pd.to_datetime(x, errors="coerce").min()
        elif low in ["check-out date", "check out date", "departure date"]:
            agg_funcs[col] = lambda x: pd.to_datetime(x, errors="coerce").max()
        else:
            agg_funcs[col] = lambda x: ', '.join(sorted(set([str(v).strip() for v in x if pd.notna(v) and str(v).strip()!=''])))

    try:
        df_utama_all = df_utama_all.groupby("Booking ID", as_index=False).agg(agg_funcs)
    except Exception:
        df_utama_all = df_utama_all.groupby("Booking ID", as_index=False).agg(
            lambda x: ', '.join(sorted(set([str(v).strip() for v in x if pd.notna(v) and str(v).strip()!=''])))
        )
else:
    st.error("❌ Kolom 'Booking ID' tidak ditemukan di file supplier.")
    st.stop()

progress_bar.progress(75)
#status_text.markdown("**🗂️ Processing Master Data...**")

# ====== PROSES FILE DATA ======
progress_bar.progress(75)

# Proses multiple file data master
df_data_list = []
data_cols_sets = []
for f in file_data:
    try:
        df = pd.read_excel(f)
        # Normalisasi kolom
        df = normalize_and_rename(df)
        df_data_list.append(df)
        data_cols_sets.append(set(df.columns))
    except Exception as e:
        st.error(f"❌ Error memproses file data master **{f.name}**: {e}")
        st.stop()

if len(df_data_list) > 0:
    df_data = pd.concat(df_data_list, ignore_index=True, sort=False)
else:
    st.error("❌ Tidak ada file data master yang valid.")
    st.stop()

df_data.columns = [str(c).strip() for c in df_data.columns]

# Mapping Source Rescode → Booking ID jika belum ada
if "Source Rescode" in df_data.columns and "Booking ID" not in df_data.columns:
    df_data = df_data.rename(columns={"Source Rescode": "Booking ID"})

lower_map = {c.lower().strip(): c for c in df_data.columns}
if "source rescode" in lower_map and "Booking ID" not in df_data.columns:
    orig = lower_map["source rescode"]
    df_data = df_data.rename(columns={orig: "Booking ID"})

df_data = df_data.loc[:, ~df_data.columns.duplicated()]

# Validasi kolom wajib
if "Booking ID" not in df_data.columns or "Invoice No" not in df_data.columns:
    st.error("❌ File Data Master harus memiliki kolom 'Booking ID' dan 'Invoice No'.")
    st.stop()

# Gabungkan invoice untuk setiap Booking ID
df_data_invoice = (
    df_data[["Booking ID", "Invoice No"]]
    .dropna(subset=["Booking ID"])
    .astype({"Booking ID": str, "Invoice No": str})
    .groupby("Booking ID", as_index=False)["Invoice No"]
    .agg(lambda vals: ', '.join(sorted(set([v.strip() for v in vals if v]))))
)

# ====== MERGE DATA ======
df_merge = pd.merge(df_utama_all, df_data_invoice, on="Booking ID", how="left")

if len(common_main_cols) > 0:
    common_main_cols = [c for c in common_main_cols if c in df_merge.columns]
    final_cols = ["Booking ID"] + [c for c in sorted(common_main_cols) if c != "Booking ID"]
    if "Invoice No" in df_merge.columns:
        final_cols.append("Invoice No")
    if len(final_cols) <= 1:
        df_final = df_merge.copy()
    else:
        final_cols = [c for c in final_cols if c in df_merge.columns]
        df_final = df_merge[final_cols].copy()
else:
    df_final = df_merge.copy()
    if "Invoice No" not in df_final.columns and "Invoice No" in df_merge.columns:
        df_final["Invoice No"] = df_merge["Invoice No"]

progress_bar.progress(100)
status_text.markdown("**✅ Complete!**")


# ====== FILTER CONTROL PANEL ======
with st.sidebar:
    st.markdown("---")
    st.markdown("## 🎛️ Control Panel")
    
    # Global Search
    st.markdown("### Global Search")
    search_query = st.text_input(
        "Search across all columns", 
        value="", 
        placeholder="Enter search keywords...",
        help="Pencarian akan dilakukan di semua kolom data"
    )

    # Advanced Filters
    st.markdown("### Advanced Filters")
    
    # Date Range Filter
    use_date_filter = False
    if "Check-In Date" in df_final.columns:
        try:
            df_final["Check-In Date"] = pd.to_datetime(df_final["Check-In Date"], errors="coerce")
            min_ci = df_final["Check-In Date"].min().date()
            max_ci = df_final["Check-In Date"].max().date()
            use_date_filter = True
        except:
            pass

    if "Check-Out Date" in df_final.columns:
        try:
            df_final["Check-Out Date"] = pd.to_datetime(df_final["Check-Out Date"], errors="coerce")
            min_co = df_final["Check-Out Date"].min().date()
            max_co = df_final["Check-Out Date"].max().date()
            use_date_filter = True
        except:
            pass

    if use_date_filter:
        col1, col2 = st.columns(2)
        with col1:
            start_date = st.date_input("Check-In ≥", value=min_ci, help="Tanggal check-in minimum")
        with col2:
            end_date = st.date_input("Check-Out ≤", value=max_co, help="Tanggal check-out maksimum")
    else:
        start_date = None
        end_date = None

    # Agent Filter
    nama_agen_opts = sorted(df_final["Nama Agen"].dropna().unique()) if "Nama Agen" in df_final.columns else []
    if nama_agen_opts:
        nama_agen = st.multiselect(
            "🏢 Filter by Agent", 
            options=nama_agen_opts,
            help="Pilih agen tertentu untuk analisis"
        )
    else:
        nama_agen = []

    # Hotel Filter
    hotel_opts = sorted(df_final["Nama Hotel"].dropna().unique()) if "Nama Hotel" in df_final.columns else []
    if hotel_opts and len(hotel_opts) <= 100:  # Limit options untuk performance
        selected_hotels = st.multiselect(
            "🏨 Filter by Hotel",
            options=hotel_opts[:50],  # Limit tampilan
            help="Pilih hotel tertentu (max 50 teratas)"
        )
    else:
        selected_hotels = []

    # Quick Actions
    st.markdown("### ⚡ Quick Actions")
    if st.button("🔄 Reset All Filters", help="Reset semua filter ke pengaturan default"):
        st.rerun()
    
    # Export Options
    st.markdown("### Export Options")
    export_format = st.selectbox(
        "Choose Export Format",
        ["Excel (.xlsx)", "CSV (.csv)"],
        help="Pilih format file untuk export"
    )

# ====== APPLY FILTERS ======
df_filtered = df_final.copy()

# Apply search filter
if search_query:
    mask = df_filtered.astype(str).apply(
        lambda row: row.str.contains(search_query, case=False, na=False)
    ).any(axis=1)
    df_filtered = df_filtered[mask]

# Apply date filters
if start_date and "Check-In Date" in df_filtered.columns:
    df_filtered = df_filtered[df_filtered["Check-In Date"] >= pd.to_datetime(start_date)]

if end_date and "Check-Out Date" in df_filtered.columns:
    df_filtered = df_filtered[df_filtered["Check-Out Date"] <= pd.to_datetime(end_date)]

# Apply agent filter
if nama_agen and "Nama Agen" in df_filtered.columns:
    df_filtered = df_filtered[df_filtered["Nama Agen"].isin(nama_agen)]

# Apply hotel filter
if selected_hotels and "Nama Hotel" in df_filtered.columns:
    df_filtered = df_filtered[df_filtered["Nama Hotel"].isin(selected_hotels)]

# ====== HASIL ANALISIS ======
st.markdown("---")

# Filter Status
if len(df_filtered) != len(df_final):
    filter_ratio = (len(df_filtered) / len(df_final)) * 100
    st.markdown(f"""
    <div class="status-info">
        🔍 <strong>Active Filters:</strong> Showing {len(df_filtered):,} of {len(df_final):,} records 
        ({filter_ratio:.1f}% of total data)
    </div>
    """, unsafe_allow_html=True)

# Data Overview
col1, col2 = st.columns([3, 1])
with col1:
    st.markdown("### 📋 Data Integration Results")

# Display results
if not df_filtered.empty:
    st.dataframe(
        df_filtered, 
        use_container_width=True,
        height=450
    )
else:
    st.info("No data available to display.")


# ====== DOWNLOAD SECTION ======
def to_excel_bytes(df: pd.DataFrame) -> bytes:
    out = BytesIO()
    with pd.ExcelWriter(out, engine="xlsxwriter") as writer:
        df.to_excel(writer, index=False, sheet_name="Integrated_Data")
        summary_data = {
            "Metric": ["Total Records", "Unique Booking IDs", "Files Processed", "Processing Date"],
            "Value": [len(df), df["Booking ID"].nunique(), len(file_utama_list), datetime.now().strftime("%Y-%m-%d %H:%M")]
        }
        pd.DataFrame(summary_data).to_excel(writer, index=False, sheet_name="Summary")
    return out.getvalue()

def to_csv_bytes(df: pd.DataFrame) -> bytes:
    return df.to_csv(index=False).encode('utf-8')

if not df_filtered.empty:
    # Hitung tambahan informasi
    size_bytes = df_filtered.memory_usage(deep=True).sum()
    size_str = f"{size_bytes/1024/1024:.2f} MB" if size_bytes > 1024*1024 else f"{size_bytes/1024:.2f} KB"
    nama_tamu_count = df_filtered["Nama Tamu"].nunique() if "Nama Tamu" in df_filtered.columns else 0
    nama_hotel_count = df_filtered["Nama Hotel"].nunique() if "Nama Hotel" in df_filtered.columns else 0
    invoice_no_count = df_filtered["Invoice No"].nunique() if "Invoice No" in df_filtered.columns else 0
    invoice_no_blank = df_filtered["Invoice No"].isna().sum() if "Invoice No" in df_filtered.columns else 0

# ====== DOWNLOAD SECTION ======
def to_excel_bytes(df: pd.DataFrame) -> bytes:
    out = BytesIO()
    with pd.ExcelWriter(out, engine="xlsxwriter") as writer:
        df.to_excel(writer, index=False, sheet_name="Integrated_Data")
        summary_data = {
            "Metric": ["Total Records", "Unique Booking IDs", "Files Processed", "Processing Date"],
            "Value": [len(df), df["Booking ID"].nunique(), len(file_utama_list), datetime.now().strftime("%Y-%m-%d %H:%M")]
        }
        pd.DataFrame(summary_data).to_excel(writer, index=False, sheet_name="Summary")
    return out.getvalue()

def to_csv_bytes(df: pd.DataFrame) -> bytes:
    return df.to_csv(index=False).encode('utf-8')

if not df_filtered.empty:
    # Hitung metrik
    size_bytes = df_filtered.memory_usage(deep=True).sum()
    size_str = f"{size_bytes/1024/1024:.2f} MB" if size_bytes > 1024*1024 else f"{size_bytes/1024:.2f} KB"
    nama_tamu_count = df_filtered["Nama Tamu"].nunique() if "Nama Tamu" in df_filtered.columns else 0
    nama_hotel_count = df_filtered["Nama Hotel"].nunique() if "Nama Hotel" in df_filtered.columns else 0
    invoice_no_count = df_filtered["Invoice No"].nunique() if "Invoice No" in df_filtered.columns else 0
    invoice_no_blank = df_filtered["Invoice No"].isna().sum() if "Invoice No" in df_filtered.columns else 0

    # === Tampilan UI Export Data Overview ===
    export_html = f"""
    <div style="
        background:linear-gradient(135deg, #f8fafc 0%, #eef2ff 100%);
        border:1px solid #e2e8f0;
        padding:1.5rem;
        border-radius:14px;
        margin-bottom:1.5rem;
        box-shadow:0 4px 10px rgba(0,0,0,0.05);
    ">
        <h3 style="margin:0 0 1.2rem 0; color:#1e3a8a; text-align:center;">Export Data Overview</h3>
        <div style="display:grid; grid-template-columns:repeat(auto-fit, minmax(160px, 1fr)); gap:0.8rem;">
            <div style="background:white; border-radius:10px; padding:0.8rem; border:1px solid #e5e7eb; text-align:center;">
                <div style="font-size:1.5rem;">📊</div>
                <div style="font-size:0.85rem; color:#6b7280;">Total Records</div>
                <div style="font-size:1.2rem; font-weight:700;">{len(df_filtered):,}</div>
            </div>
            <div style="background:white; border-radius:10px; padding:0.8rem; border:1px solid #e5e7eb; text-align:center;">
                <div style="font-size:1.5rem;">🔖</div>
                <div style="font-size:0.85rem; color:#6b7280;">Unique Booking IDs</div>
                <div style="font-size:1.2rem; font-weight:700;">{df_filtered['Booking ID'].nunique():,}</div>
            </div>
            <div style="background:white; border-radius:10px; padding:0.8rem; border:1px solid #e5e7eb; text-align:center;">
                <div style="font-size:1.5rem;">💾</div>
                <div style="font-size:0.85rem; color:#6b7280;">Data Size</div>
                <div style="font-size:1.1rem; font-weight:600;">{size_str}</div>
            </div>
            <div style="background:white; border-radius:10px; padding:0.8rem; border:1px solid #e5e7eb; text-align:center;">
                <div style="font-size:1.5rem;">🧍</div>
                <div style="font-size:0.85rem; color:#6b7280;">Unique Guests</div>
                <div style="font-size:1.2rem; font-weight:700;">{nama_tamu_count:,}</div>
            </div>
            <div style="background:white; border-radius:10px; padding:0.8rem; border:1px solid #e5e7eb; text-align:center;">
                <div style="font-size:1.5rem;">🏨</div>
                <div style="font-size:0.85rem; color:#6b7280;">Unique Hotels</div>
                <div style="font-size:1.2rem; font-weight:700;">{nama_hotel_count:,}</div>
            </div>
            <div style="background:white; border-radius:10px; padding:0.8rem; border:1px solid #e5e7eb; text-align:center;">
                <div style="font-size:1.5rem;">🧾</div>
                <div style="font-size:0.85rem; color:#6b7280;">Unique Invoice No</div>
                <div style="font-size:1.2rem; font-weight:700;">{invoice_no_count:,}</div>
            </div>
            <div style="background:white; border-radius:10px; padding:0.8rem; border:1px solid #e5e7eb; text-align:center;">
                <div style="font-size:1.5rem;">⚠️</div>
                <div style="font-size:0.85rem; color:#6b7280;">Invoice No (Blank)</div>
                <div style="font-size:1.2rem; font-weight:700; color:#dc2626;">{invoice_no_blank:,}</div>
            </div>
        </div>
    </div>
    """
    st.markdown(export_html, unsafe_allow_html=True)

    # === Tombol download kecil di kiri ===
    col1, col2 = st.columns([3, 1])
    with col1:
        timestamp = datetime.now().strftime("%Y%m%d_%H%M%S")
        if export_format == "Excel (.xlsx)":
            file_data = to_excel_bytes(df_filtered)
            filename = f"narasight_data_{timestamp}.xlsx"
            mime_type = "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
        else:
            file_data = to_csv_bytes(df_filtered)
            filename = f"narasight_data_{timestamp}.csv"
            mime_type = "text/csv"

        st.markdown("""
            <style>
            div[data-testid="stDownloadButton"] button {
                padding: 0.3rem 0.6rem;
                font-size: 0.85rem !important;
                border-radius: 6px;
                background-color: #2563eb;
                color: white;
            }
            </style>
        """, unsafe_allow_html=True)

        st.download_button(
            label=f"📥 {export_format}",
            data=file_data,
            file_name=filename,
            mime=mime_type,
            use_container_width=False
        )
else:
    st.warning("⚠️ Tidak ada data untuk diexport. Sesuaikan filter atau reset pengaturan.", icon="⚠️")


    # Footer
st.markdown("""
<hr style="margin-top: 3rem; margin-bottom: 1rem; border: none; border-top: 1px solid #ccc;" />
<div style='text-align: center; font-size: 0.85rem; color: gray;'>
    📊 Aplikasi Data Relasional narasight | Dibuat dengan ❤️ oleh <a href='https://www.linkedin.com/in/rifyalt/'>Rifyal Tumber</a><br>
    © 2025 - Versi 1.0 | Hubungi +62 878 8103 3781 jika ada kendala teknis
</div>
""", unsafe_allow_html=True)

# --- panggil hanya saat dieksekusi langsung oleh streamlit ---
if __name__ == "__main__":
    main()