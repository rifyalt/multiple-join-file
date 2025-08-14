# app.py
import streamlit as st
import pandas as pd
from io import BytesIO
from functools import reduce
from st_aggrid import AgGrid, GridOptionsBuilder, DataReturnMode, GridUpdateMode
from rapidfuzz import fuzz
import plotly.express as px

# Optional imports (aktifkan jika tersedia di environment)
try:
    from openai import OpenAI
except Exception:
    OpenAI = None

# Prophet (aktifkan jika tersedia di environment)
try:
    from prophet import Prophet
except Exception:
    Prophet = None

# ===== USER CREDENTIALS =====
USER_CREDENTIALS = {
    "admin": "admin123",
    "user1": "perta1",
}

# ===== SESSION LOGIN CHECK =====
if "logged_in" not in st.session_state:
    st.session_state.logged_in = False
if "username" not in st.session_state:
    st.session_state.username = ""

def login_page():
    st.title("🔐 Login Aplikasi narasight")
    username = st.text_input("👤 Username")
    password = st.text_input("🔑 Password", type="password")

    if st.button("🔓 Login"):
        if username in USER_CREDENTIALS and USER_CREDENTIALS[username] == password:
            st.session_state.logged_in = True
            st.session_state.username = username
            st.success(f"✅ Selamat datang, {username}!")
            st.experimental_rerun()
        else:
            st.error("❌ Username atau password salah.")

# ===== JALANKAN LOGIN JIKA BELUM MASUK =====
if not st.session_state.logged_in:
    login_page()
    st.stop()

# ===== PAGE CONFIG =====
st.set_page_config(page_title="Join Data App", layout="wide", page_icon="📊")
st.markdown("""
    <style>
        html, body, [class*="css"]  {
            font-family: 'Segoe UI', sans-serif;
        }
        .stApp {
            background-color: #f0f8ff;
        }
        .block-container {
            padding: 2rem 2rem;
        }
        .stSidebar {
            background-color: #dbeafe;
        }
        .stButton>button, .stDownloadButton>button {
            background-color: #3b82f6;
            color: white;
            border-radius: 0.5rem;
            padding: 0.5rem 1rem;
        }
        h1, h2, h3, h4, h5, h6 {
            color: #1e3a8a;
        }

        .marquee-container {
            width: 100%;
            overflow: hidden;
            white-space: nowrap;
            box-sizing: border-box;
            background: #fef9c3;
            padding: 8px 0;
            border: 1px solid #fde68a;
            border-radius: 6px;
            box-shadow: 0 1px 2px rgba(0,0,0,0.1);
            margin-top: 1rem;
            margin-bottom: 1.5rem;
        }
        .marquee-text {
            display: inline-block;
            padding-left: 100%;
            animation: marquee 55s linear infinite; 
            font-weight: bold;
            color: #92400e;
            font-size: 1rem;
        }
        @keyframes marquee {
            0%   { transform: translateX(-100%); }
            100% { transform: translateX(100%); }
        }

        .narasi-ai {
            background-color: white;
            padding: 1rem;
            border-radius: 8px;
            box-shadow: 0 1px 3px rgba(0,0,0,0.1);
            border: 1px solid #e5e7eb;
            margin-top: 1rem;
            color: #111827;
            font-size: 1rem;
        }
    </style>

    <div class="marquee-container">
        <div class="marquee-text">
            📢 Info Penting: File utama wajib diupload! Pastikan minimal input 2 file agar proses join data berhasil.
        </div>
    </div>
""", unsafe_allow_html=True)

# ===== SIDEBAR =====
with st.sidebar:
    st.markdown("## 📊 Join Multiple Excel")
    st.markdown("*Tentukan Parameter Key (mis. Travel Request Number, Booking ID, Invoice No atau Company Code)*")
    st.markdown("---")

# ===== FILE UPLOAD =====
# file_uploader dengan accept_multiple_files=True akan mengembalikan list atau None
data_utama_files = st.file_uploader(
    "🗂️ File Utama (Wajib, Multi Upload)",
    type=["xlsx"],
    key="mandatory",
    accept_multiple_files=True
)

col1, col2 = st.columns(2)
additional_files = []

with col1:
    file_data1 = st.file_uploader(
        "📄 File Data 1 (Optional single)",
        type=["xlsx"],
        key="data1"
    )
    if file_data1:
        additional_files.append(file_data1)

    multi_files_2 = st.file_uploader(
        "📂 File Data 2 (Multi Upload, Optional)",
        type=["xlsx"],
        accept_multiple_files=True,
        key="data2"
    )
    if multi_files_2:
        additional_files.extend(multi_files_2)

with col2:
    multi_files_3 = st.file_uploader(
        "📂 File Data 3 (Multi Upload, Optional)",
        type=["xlsx"],
        accept_multiple_files=True,
        key="data3"
    )
    if multi_files_3:
        additional_files.extend(multi_files_3)

    multi_files_4 = st.file_uploader(
        "📂 File Data 4 (Multi Upload, Optional)",
        type=["xlsx"],
        accept_multiple_files=True,
        key="data4"
    )
    if multi_files_4:
        additional_files.extend(multi_files_4)

# ===== CLEANING FUNCTION =====
def clean_and_cast_columns(df: pd.DataFrame) -> pd.DataFrame:
    df = df.copy()
    for col in df.columns:
        # Trim strings
        if df[col].dtype == object:
            df[col] = df[col].astype(str).str.strip()
            # coba cast ke numeric jika bisa
            try:
                df[col] = pd.to_numeric(df[col], errors='raise')
            except Exception:
                df[col] = df[col].astype(str)
        elif pd.api.types.is_numeric_dtype(df[col]):
            df[col] = pd.to_numeric(df[col], errors='coerce')
        # else biarkan
    return df

# ===== JOIN PROSES =====
if data_utama_files:
    try:
        # kumpulkan semua dataframe valid dari file_utama
        df_list = []
        for f in data_utama_files:
            try:
                df = pd.read_excel(f)
                if "Travel Request Number" not in df.columns:
                    st.warning(f"File utama '{getattr(f, 'name', str(f))}' tidak memiliki kolom 'Travel Request Number'. Dilewati.")
                    continue
                df["Travel Request Number"] = df["Travel Request Number"].astype(str).str.strip()
                df = df.apply(lambda x: x.str.strip() if x.dtype == "object" else x)
                df_list.append(df)
            except Exception as e:
                st.warning(f"Gagal baca file utama '{getattr(f, 'name', str(f))}': {e}")

        # proses file tambahan jika ada
        processed_files = 0
        for f in additional_files:
            if not f:
                continue
            try:
                df = pd.read_excel(f)
                if "Travel Request Number" not in df.columns:
                    st.warning(f"File '{getattr(f, 'name', str(f))}' tidak memiliki kolom 'Travel Request Number'. Dilewati.")
                    continue
                df["Travel Request Number"] = df["Travel Request Number"].astype(str).str.strip()
                df = df.apply(lambda x: x.str.strip() if x.dtype == "object" else x)
                df_list.append(df)
                processed_files += 1
            except Exception as e:
                st.warning(f"Gagal baca file tambahan '{getattr(f, 'name', str(f))}': {e}")

        if len(df_list) == 0:
            st.error("Tidak ada file utama yang valid (dengan kolom 'Travel Request Number') untuk diproses.")
            st.stop()

        # fungsi merge dengan prioritas: kolom non-TRN dari df1 dipertahankan, df2 melengkapi
        def merge_with_priority(df1, df2):
            df_merged = pd.merge(df1, df2, on="Travel Request Number", how="outer", suffixes=("", "_dup"))
            for col in df2.columns:
                if col == "Travel Request Number":
                    continue
                dup_col = f"{col}_dup"
                if dup_col in df_merged.columns:
                    # combine_first: ambil nilai df1 dulu, lalu df2
                    df_merged[col] = df_merged[col].combine_first(df_merged[dup_col])
                    df_merged.drop(columns=[dup_col], inplace=True, errors='ignore')
            return df_merged

        join_result = reduce(merge_with_priority, df_list)

        # Hapus kolom kosong total
        join_result.dropna(axis=1, how='all', inplace=True)
        # Hapus baris yang semua nilainya kosong
        join_result.dropna(how='all', inplace=True)

        # Ganti NaT pada kolom datetime jadi string kosong (agar aman tampil)
        for col in join_result.select_dtypes(include=["datetime", "datetimetz"]).columns:
            join_result[col] = join_result[col].fillna(pd.NaT).astype(str).replace("NaT", "")

        # ===== FITUR SEARCH DI SIDEBAR =====
        st.sidebar.markdown("## 🕵️ Pencarian Data")
        search_keyword = st.sidebar.text_input("🔍 Cari Kata Kunci (semua kolom & semua tipe data)")

        if search_keyword:
            keyword = search_keyword.strip().lower()
            mask = join_result.apply(lambda row: keyword in ' '.join(row.map(lambda v: '' if pd.isna(v) else str(v)).map(str).str.lower()), axis=1)
            join_result = join_result[mask]

        # ===== FILTER DATA USER-FRIENDLY =====
        st.sidebar.markdown("## 🗂️ Filter Tanggal")

        if 'Check-In Date' in join_result.columns and 'Check-Out Date' in join_result.columns:
            # Pastikan kolom tanggal dalam format datetime
            join_result['Check-In Date'] = pd.to_datetime(join_result['Check-In Date'], errors='coerce')
            join_result['Check-Out Date'] = pd.to_datetime(join_result['Check-Out Date'], errors='coerce')

            # Tentukan rentang minimum dan maksimum (beri nilai default jika NaT)
            min_checkin = join_result['Check-In Date'].min()
            max_checkout = join_result['Check-Out Date'].max()

            # Jika kolom kosong total, berikan peringatan
            if pd.isna(min_checkin) or pd.isna(max_checkout):
                st.sidebar.warning("Kolom 'Check-In Date' atau 'Check-Out Date' mengandung sedikit/tipe data tidak valid. Filter tanggal dinonaktifkan.")
            else:
                st.sidebar.markdown("Pilih rentang tanggal untuk menampilkan data:")
                selected_date_range = st.sidebar.date_input(
                    label="📅 Periode Tanggal (Check-In s.d Check-Out)",
                    value=(min_checkin.date(), max_checkout.date()),
                    min_value=min_checkin.date(),
                    max_value=max_checkout.date()
                )

                if isinstance(selected_date_range, tuple) and len(selected_date_range) == 2:
                    start_date, end_date = selected_date_range
                    filtered = join_result[
                        (join_result['Check-In Date'] <= pd.to_datetime(end_date)) &
                        (join_result['Check-Out Date'] >= pd.to_datetime(start_date))
                    ]
                    join_result = filtered

                    st.sidebar.success(
                        f"📆 Menampilkan data dari {start_date.strftime('%d %b %Y')} - {end_date.strftime('%d %b %Y')}\n"
                        f"📄 Jumlah data: {len(join_result)}"
                    )
        else:
            st.sidebar.warning("Kolom 'Check-In Date' dan/atau 'Check-Out Date' tidak ditemukan.")

        # Company Code filter
        if 'Company Code' in join_result.columns:
            labels = ['Semua'] + sorted(join_result['Company Code'].dropna().astype(str).unique().tolist())
            selected_code = st.sidebar.selectbox("🎫 Company Code", labels)
            if selected_code != "Semua":
                join_result = join_result[join_result['Company Code'].astype(str) == selected_code]

        # ===== ANALISIS SIMILARITAS NAMA HOTEL =====
        def analyze_hotel_similarity(df, threshold=90):
            if 'Hotel Name' not in df.columns:
                st.info("Kolom 'Hotel Name' tidak ditemukan untuk analisis kemiripan.")
                return

            hotel_names = df['Hotel Name'].dropna().unique().tolist()
            similar_pairs = []
            for i, name in enumerate(hotel_names):
                for other in hotel_names[i+1:]:
                    try:
                        score = fuzz.token_sort_ratio(name, other)
                        if score >= threshold:
                            similar_pairs.append({
                                'Hotel 1': name,
                                'Hotel 2': other,
                                'Similarity %': int(score)
                            })
                    except Exception:
                        continue

            if similar_pairs:
                st.markdown("## 🔍 Analisis Kemiripan Nama Hotel")
                st.dataframe(pd.DataFrame(similar_pairs).sort_values('Similarity %', ascending=False))
            else:
                st.info(f"Tidak ada nama hotel dengan kemiripan di atas {threshold}%.")

        analyze_hotel_similarity(join_result, threshold=85)

        # ===== PREVIEW DATA INTERAKTIF =====
        st.markdown("## 👀 Preview Data")
        gb = GridOptionsBuilder.from_dataframe(join_result)
        gb.configure_pagination(paginationAutoPageSize=False, paginationPageSize=20)
        gb.configure_default_column(filterable=True, sortable=True, resizable=True)
        gridOptions = gb.build()
        AgGrid(
            join_result,
            gridOptions=gridOptions,
            data_return_mode=DataReturnMode.FILTERED_AND_SORTED,
            update_mode=GridUpdateMode.SELECTION_CHANGED,
            fit_columns_on_grid_load=False,
            height=400,
            enable_enterprise_modules=False,
            allow_unsafe_jscode=True,
            custom_js="""
            function(e) {
                let api = e.api;
                api.sizeColumnsToFit();
                setTimeout(function() {
                    const allColumnIds = [];
                    api.getColumnDefs().forEach(function(colDef) {
                        allColumnIds.push(colDef.field);
                    });
                    api.autoSizeColumns(allColumnIds, false);
                }, 100);
            }
            """
        )

        # ===== INSIGHT =====
        st.markdown("---")
        st.markdown("## 🗂️ File Info")

        data_size_bytes = join_result.memory_usage(deep=True).sum()
        data_size_mb = data_size_bytes / (1024 ** 2)

        # Voucher counts (treat NaN explicitly)
        voucher_counts = {}
        if 'Voucher Hotel' in join_result.columns:
            join_result['Voucher Hotel'] = join_result['Voucher Hotel'].fillna('nan').astype(str).str.strip()
            voucher_counts = join_result['Voucher Hotel'].value_counts(dropna=False).to_dict()

        voucher_yes = voucher_counts.get('Yes', 0)
        voucher_no = voucher_counts.get('No', 0)
        voucher_nan = voucher_counts.get('nan', 0)

        col1i, col2i, col3i = st.columns(3)

        with col1i:
            st.metric("📁 Ukuran Join Data", f"{data_size_mb:.2f} MB")
            st.metric("🧾 Total Baris", f"{len(join_result)}")
            st.metric("📊 Total Kolom", f"{join_result.shape[1]}")
            st.metric("🎟️ Voucher 'nan' (blank)", f"{voucher_nan}")

        with col2i:
            st.metric("🏢 Perusahaan Unik", f"{join_result['Company Code'].nunique() if 'Company Code' in join_result.columns else 0}")
            st.metric("🧑‍💼 Employee Number", f"{join_result['Employee Number'].nunique() if 'Employee Number' in join_result.columns else 0}")
            st.metric("🎟️ Voucher 'Yes'", f"{voucher_yes}")

        with col3i:
            st.metric("🏨 Hotel Unik", f"{join_result['Hotel Name'].nunique() if 'Hotel Name' in join_result.columns else 0}")
            if 'Number of Rooms Night' in join_result.columns:
                total_night = pd.to_numeric(join_result['Number of Rooms Night'], errors='coerce').sum()
                st.metric("🛏️ Total Room Night", f"{total_night:,.0f}")
                st.metric("🎟️ Voucher 'No'", f"{voucher_no}")

        # ===== ANALISA TOP 10 (PLOTLY + TABEL) =====
        st.markdown("## 📊 Summarized")

        def plot_top(df, col, title):
            if col not in df.columns:
                st.info(f"Kolom '{col}' tidak ditemukan.")
                return
            top = df[col].fillna('nan').astype(str).value_counts().head(10).reset_index()
            top.columns = [col, 'Jumlah']
            colL, colR = st.columns([3, 1])
            with colL:
                fig = px.bar(
                    top.sort_values('Jumlah', ascending=True),
                    x='Jumlah',
                    y=col,
                    orientation='h',
                    title=title,
                    text='Jumlah'
                )
                fig.update_layout(yaxis=dict(categoryorder='total ascending'))
                st.plotly_chart(fig, use_container_width=True)
            with colR:
                st.markdown(f"#### 📋 Top 10: {col}")
                st.dataframe(top.sort_values("Jumlah", ascending=False), use_container_width=True)

        # cek dan plot beberapa kolom yang umum
        if 'City' in join_result.columns:
            plot_top(join_result, 'City', '🏙️ Top 10 City')
        if 'Employee Name' in join_result.columns:
            plot_top(join_result, 'Employee Name', '👤 Top 10 Employee Name')
        if 'Traveling Purpose' in join_result.columns:
            plot_top(join_result, 'Traveling Purpose', '👤 Top 10 Traveling Purpose')

        possible_dirs = [c for c in join_result.columns if 'direktorat' in c.lower() or 'directorate' in c.lower()]
        if possible_dirs:
            plot_top(join_result, possible_dirs[0], f"🏢 Top 10 {possible_dirs[0]}")

        if 'Nama Fungsi' in join_result.columns:
            plot_top(join_result, 'Nama Fungsi', '🧩 Top 10 Nama Fungsi')

        # ===== FUNGSI ANALISA ROOM NIGHT =====
        def show_room_night_analysis(df):
            if 'Hotel Name' in df.columns and 'Number of Rooms Night' in df.columns:
                df['Number of Rooms Night'] = pd.to_numeric(df['Number of Rooms Night'], errors='coerce').fillna(0)
                top_hotel_rooms = (
                    df.groupby('Hotel Name')['Number of Rooms Night']
                    .sum()
                    .sort_values(ascending=False)
                    .head(10)
                    .reset_index()
                )
                colA, colB = st.columns([3, 1])
                with colA:
                    fig_top_hotel = px.bar(
                        top_hotel_rooms.sort_values('Number of Rooms Night', ascending=True),
                        x='Number of Rooms Night',
                        y='Hotel Name',
                        orientation='h',
                        title='🏨 Top 10 Hotel berdasarkan Jumlah Room Night',
                        text='Number of Rooms Night'
                    )
                    fig_top_hotel.update_layout(yaxis=dict(categoryorder='total ascending'))
                    st.plotly_chart(fig_top_hotel, use_container_width=True)
                with colB:
                    st.markdown("#### 📋 Tabel Top 10 Hotel")
                    st.dataframe(top_hotel_rooms, use_container_width=True)
            else:
                st.info("Data tidak memiliki kolom 'Hotel Name' dan/atau 'Number of Rooms Night' untuk analisa room night.")

            if 'Check-In Date' in df.columns and 'Number of Rooms Night' in df.columns:
                df['Check-In Date'] = pd.to_datetime(df['Check-In Date'], errors='coerce')
                df['Number of Rooms Night'] = pd.to_numeric(df['Number of Rooms Night'], errors='coerce').fillna(0)

                df_ts = (
                    df.groupby('Check-In Date')['Number of Rooms Night']
                    .sum()
                    .reset_index()
                    .sort_values('Check-In Date')
                )

                if not df_ts.empty:
                    fig_ts = px.line(
                        df_ts,
                        x='Check-In Date',
                        y='Number of Rooms Night',
                        markers=True,
                        title='📅 Tren Room Night per Tanggal Check-In'
                    )
                    st.plotly_chart(fig_ts, use_container_width=True)

        def show_voucher_amount_analysis(df):
            if 'Check-In Date' in df.columns and 'Voucher Hotel Amount' in df.columns:
                df['Check-In Date'] = pd.to_datetime(df['Check-In Date'], errors='coerce')
                df['Voucher Hotel Amount'] = (
                    df['Voucher Hotel Amount']
                    .astype(str)
                    .replace('[^\d.,-]', '', regex=True)
                    .str.replace(',', '', regex=False)
                )
                df['Voucher Hotel Amount'] = pd.to_numeric(df['Voucher Hotel Amount'], errors='coerce').fillna(0)

                df_voucher_ts = (
                    df.groupby('Check-In Date')['Voucher Hotel Amount']
                    .sum()
                    .reset_index()
                    .sort_values('Check-In Date')
                )

                if not df_voucher_ts.empty:
                    fig_voucher_ts = px.line(
                        df_voucher_ts,
                        x='Check-In Date',
                        y='Voucher Hotel Amount',
                        markers=True,
                        title='💵 Tren Voucher Hotel Amount per Tanggal Check-In'
                    )
                    st.plotly_chart(fig_voucher_ts, use_container_width=True)
            else:
                st.info("Data tidak memiliki kolom 'Check-In Date' dan/atau 'Voucher Hotel Amount' untuk analisa.")

        def show_forecasting_travel_request(df):
            st.markdown("### Prediksi Jumlah Perjalanan (Travel Request)")
            st.markdown("<small style='color:gray'>Menggunakan Prophet (jika tersedia) untuk prediksi bulanan.</small>", unsafe_allow_html=True)

            if 'Check-In Date' not in df.columns:
                st.warning("Kolom 'Check-In Date' tidak ditemukan.")
                return

            if Prophet is None:
                st.info("Library Prophet tidak ditemukan di environment. Install prophet jika ingin fitur forecasting.")
                return

            if st.button("Analyze Forecast"):
                df_local = df.copy()
                df_local['Check-In Date'] = pd.to_datetime(df_local['Check-In Date'], errors='coerce')
                df_local['month'] = df_local['Check-In Date'].dt.to_period('M').dt.to_timestamp()
                df_monthly = df_local.groupby('month').agg({'Travel Request Number': 'count'}).reset_index()
                df_monthly.columns = ['ds', 'y']
                df_monthly = df_monthly[df_monthly['y'] > 0]
                if df_monthly.empty or len(df_monthly) < 4:
                    st.warning("Data tidak cukup untuk membuat prediksi (minimal 4 bulan dengan data).")
                    return

                import plotly.graph_objs as go
                with st.spinner("🔄 Memprediksi tren perjalanan..."):
                    model = Prophet()
                    model.fit(df_monthly)
                    future = model.make_future_dataframe(periods=6, freq='M')
                    forecast = model.predict(future)

                    fig = go.Figure()
                    fig.add_trace(go.Scatter(x=df_monthly['ds'], y=df_monthly['y'], mode='lines+markers', name='Aktual'))
                    fig.add_trace(go.Scatter(x=forecast['ds'], y=forecast['yhat'], mode='lines', name='Prediksi'))

                    fig.update_layout(
                        title='📈 Prediksi Jumlah Perjalanan (Travel Request) per Bulan',
                        xaxis_title='Bulan',
                        yaxis_title='Jumlah Travel Request',
                        legend=dict(orientation="h", yanchor="bottom", y=1.02, xanchor="right", x=1)
                    )
                    st.plotly_chart(fig, use_container_width=True)

        # panggil analisa
        show_room_night_analysis(join_result)
        show_voucher_amount_analysis(join_result)
        show_forecasting_travel_request(join_result)

        # ===== OPEN AI NARRATIVE (opsional) =====
        if OpenAI is None:
            st.info("Library OpenAI tidak ditemukan di environment. Fitur narasi tidak tersedia.")
        else:
            # Pastikan key ada di st.secrets
            openai_available = "OPENAI_API_KEY" in st.secrets and st.secrets["OPENAI_API_KEY"]
            if not openai_available:
                st.info("OpenAI API Key tidak ditemukan di st.secrets. Tambahkan 'OPENAI_API_KEY' jika ingin menggunakan fitur narasi.")
            else:
                client = OpenAI(api_key=st.secrets["OPENAI_API_KEY"])

                def generate_narrative(df, topic="summary"):
                    try:
                        if df.empty:
                            return "Data tidak tersedia untuk dianalisis."
                        sample_text = df.head(20).to_string(index=False)
                        prompt = f"""
Berikan analisa naratif singkat dan mudah dipahami dari data berikut terkait topik: {topic}.
Gunakan gaya bahasa profesional, dan tonjolkan insight menarik jika ada.

Data:
{sample_text}
"""
                        # Berbagai versi SDK OpenAI berbeda; coba-catch agar tidak crash bila API mismatch
                        try:
                            response = client.chat.completions.create(
                                model="gpt-4",
                                messages=[
                                    {"role": "system", "content": "Kamu adalah asisten data yang profesional dan mudah dimengerti."},
                                    {"role": "user", "content": prompt}
                                ],
                                temperature=0.4,
                                max_tokens=500
                            )
                            # response structure mungkin berbeda; tangani fleksibel
                            if hasattr(response, "choices"):
                                return response.choices[0].message.content
                            elif isinstance(response, dict):
                                # fallback parsing
                                choices = response.get("choices", [])
                                if choices:
                                    return choices[0].get("message", {}).get("content", "") or choices[0].get("text", "")
                                return str(response)
                            else:
                                return str(response)
                        except Exception as e:
                            return f"Gagal memanggil API OpenAI: {e}"
                    except Exception as e:
                        return f"Gagal menghasilkan narasi: {e}"

                with st.expander("🦖 Aku Bantu Kasih Narasi, MAU? (Klik untuk lihat narasi)"):
                    topic_desc = st.text_input("Jelaskan topik narasi (misal: ringkasan tren hotel, analisa room night, dsb)", value="ringkasan tren")
                    if st.button("Berikan Narasi"):
                        with st.spinner("AI Sedang membuat narasi..."):
                            narrative = generate_narrative(join_result, topic_desc)
                            st.markdown("**Berikut hasilnya:**")
                            st.markdown(f"""<div class="narasi-ai">{narrative}</div>""", unsafe_allow_html=True)

        # ===== DOWNLOAD BUTTON =====
        if st.checkbox("✅ Aktifkan Download"):
            output = BytesIO()
            with pd.ExcelWriter(output, engine='xlsxwriter') as writer:
                join_result.to_excel(writer, index=False, sheet_name='Joined Data')
            output.seek(0)
            st.download_button(
                "⬇️ Download Excel",
                data=output,
                file_name="joined_data.xlsx",
                mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
            )

    except Exception as e:
        st.error(f"❌ Terjadi kesalahan saat join/filter: {e}")
else:
    st.info("⬆️ Silakan upload minimal 1 file utama terlebih dahulu.")      

# ===== FOOTER =====
st.markdown("""
<hr style="margin-top: 3rem; margin-bottom: 1rem; border: none; border-top: 1px solid #ccc;" />
<div style='text-align: center; font-size: 0.85rem; color: gray;'>
    📊 Aplikasi Data Relasional narasight | Dibuat dengan ❤️ oleh <a href='https://www.linkedin.com/in/rifyalt/'>Rifyal Tumber</a><br>
    © 2025 - Versi 1.0 | Hubungi +62 878 8103 3781 jika ada kendala teknis
</div>
""", unsafe_allow_html=True)
