import streamlit as st
import pandas as pd
import plotly.express as px
import plotly.graph_objects as go
import numpy as np
import calendar
import textwrap

# 1. KONFIGURASI HALAMAN
st.set_page_config(page_title="IT Support Dashboard", layout="wide")

st.markdown("""
    <style>
    .main { background-color: #F8F9FA; }
    .stMetric { 
        background-color: #ffffff; 
        padding: 15px; 
        border-radius: 10px; 
        box-shadow: 2px 2px 5px rgba(0,0,0,0.05); 
        border: 1px solid #efefef;
    }
    </style>
    """, unsafe_allow_html=True)

st.title("🖥️ Dashboard Informasi IT Support")

# 2. FUNGSI PEMBERSIHAN DATA
def clean_data(df):
    df['Tanggal'] = pd.to_datetime(df['Tanggal'], errors='coerce')
    # Ekstraksi Tahun & Bulan
    df['Tahun'] = df['Tanggal'].dt.year
    df['Bulan'] = df['Tanggal'].dt.month
    df.columns = df.columns.str.strip()
    
    # Cleaning Permasalahan
    df = df.dropna(subset=['Permasalahan'])
    df = df[~df['Permasalahan'].astype(str).str.strip().isin(['', '-', '.', 'nan', 'NaN'])]
    
    def refine_txt(t):
        t = str(t).lower().strip()
        if 'lan' in t or 'utp' in t or 'kabel' in t: return 'Troubleshoot Jaringan LAN'
        if 'internet' in t or 'wifi' in t or 'konek' in t: return 'Troubleshoot Jaringan Internet'
        if 'ups' in t or 'listrik' in t: return 'Maintenance UPS'
        if 'cctv' in t: return 'Troubleshoot CCTV'
        return t.title()
    
    df['Problem_Clean'] = df['Permasalahan'].apply(refine_txt)
    
    # Cleaning Lokasi
    df = df.dropna(subset=['Lokasi'])
    df['Loc_Clean'] = df['Lokasi'].astype(str).str.strip().str.upper()
    
    # Cleaning Waktu & Tanggal
    def get_hour(x):
        try:
            if pd.isna(x) or str(x) == '00:00': return None
            return int(str(x).split(':')[0])
        except: return None
    
    df['Hour'] = df['Jam Mulai'].apply(get_hour)
    df['Tanggal'] = pd.to_datetime(df['Tanggal'], errors='coerce')
    
    month_mapping = {
        'January': 'Januari', 'February': 'Februari', 'March': 'Maret',
        'April': 'April', 'May': 'Mei', 'June': 'Juni',
        'July': 'Juli', 'August': 'Agustus', 'September': 'September',
        'October': 'Oktober', 'November': 'November', 'December': 'Desember'
    }
    
    df['Bulan_Nama'] = df['Tanggal'].dt.month_name().map(month_mapping)
    return df.dropna(subset=['Tanggal', 'Problem_Clean'])

# 3. SIDEBAR UPLOAD & FILTER
st.sidebar.header("Data & Filter")
uploaded_file = st.sidebar.file_uploader("Upload file Excel atau CSV", type=['csv', 'xlsx'])

if uploaded_file is not None:
    try:
        if uploaded_file.name.endswith('.csv'):
            data_raw = pd.read_csv(uploaded_file)
        else:
            data_raw = pd.read_excel(uploaded_file)
        
        df_full = clean_data(data_raw)
        # ===== FILTER TAHUN =====
        tahun_list = sorted(df_full['Tahun'].dropna().unique())
        selected_year = st.sidebar.selectbox("📆 Pilih Tahun", tahun_list)

        
        list_bulan = ["Pilih Bulan...", "Semua Bulan", "Januari", "Februari", "Maret", "April", "Mei", "Juni", 
                      "Juli", "Agustus", "September", "Oktober", "November", "Desember"]
        
        selected_month = st.sidebar.selectbox("📅 Pilih Periode Laporan:", list_bulan)

        # --- LOGIKA PENAMPILAN DATA ---
        df_year = df_full[df_full['Tahun'] == selected_year]
        if selected_month == "Pilih Bulan...":
            st.info("Silakan pilih periode bulan pada sidebar untuk menampilkan analisis data.")
                                  
        else:
            if selected_month == "Semua Bulan":
                df_filtered = df_year
                st.sidebar.success(f"Menampilkan Rangkuman Semua Bulan - {selected_year}")
            else:
                df_filtered = df_year[df_year['Bulan_Nama'] == selected_month]
                
            if df_filtered.empty:
                st.warning(f"⚠️ Tidak ada data ditemukan untuk periode {selected_month}.")
            else:
                # --- KPI METRICS ---
                st.markdown(f"### 📌 Ringkasan pada Bulan: {selected_month}")
                kpi1, kpi2, kpi3 = st.columns(3)
                
                peak_h = df_filtered['Hour'].mode()[0] if not df_filtered['Hour'].mode().empty else 0
                
                kpi1.metric("Total Gangguan", f"{len(df_filtered)} Kasus")
                kpi2.metric("Lokasi Terpadat", df_filtered['Loc_Clean'].mode()[0])
                kpi3.metric("Jam Paling Rawan", f"{int(peak_h):02d}:00 WIB")
                
                st.markdown("---")

                
                # VISUALISASI 2: Distribusi Lokasi
                st.subheader("Distribusi Gangguan IT Berdasarkan Lokasi")
                top_l = df_filtered['Loc_Clean'].value_counts().head(3).reset_index()

                fig_l = px.pie(top_l, 
                               names='Loc_Clean', 
                               values='count', 
                               hole=0.4,
                               color_discrete_sequence=px.colors.qualitative.Safe,
                               labels={'Loc_Clean': 'Nama Lokasi', 'count': 'Jumlah Kasus'})
                fig_l.update_traces(
    textinfo='percent',
    textfont_size=13,
    textposition='inside'
)

                fig_l.update_traces(hovertemplate="<b>%{label}</b><br>Jumlah: %{value} Kasus<br>Persentase: %{percent}")
                fig_l.update_layout(
                    margin=dict(l=20, r=20, t=30, b=20),
                    height=450,
                    legend_title_text='Daftar Lokasi'
                )
                st.plotly_chart(fig_l, use_container_width=True)
                # --- TAMBAHAN: TABEL DETAIL GANGGUAN PADA TOP LOKASI ---
                st.markdown("#### Detail Laporan pada Top Lokasi")
                
                # 1. Mengambil list nama lokasi yang masuk di visual Pie Chart (Top 3)
                top_locations = top_l['Loc_Clean'].tolist()
                
                # 2. Filter data utama hanya untuk lokasi-lokasi tersebut
                df_top_detail = df_filtered[df_filtered['Loc_Clean'].isin(top_locations)].copy()
                
                # 3. Sort (Urutkan) berdasarkan Lokasi agar muncul berturut-turut
                df_top_detail = df_top_detail.sort_values(by=['Loc_Clean', 'Tanggal'])
                
                # 4. Merapikan format Tanggal (Hari-Bulan-Tahun)
                df_top_detail['Tgl_Indo'] = df_top_detail['Tanggal'].dt.strftime('%d-%m-%Y')
                
                # 5. Pilih kolom: Lokasi, Permasalahan, Tanggal, Jam Mulai
                tabel_final = df_top_detail[['Loc_Clean', 'Permasalahan', 'Tgl_Indo', 'Jam Mulai']]
                tabel_final.columns = ['LOKASI', 'PERMASALAHAN', 'TANGGAL', 'JAM MULAI']
                
                # 6. Tampilkan Tabel
                st.dataframe(tabel_final, use_container_width=True, hide_index=True)
                
                st.markdown("---")

                # --- VISUALISASI BARIS 3 ---
                st.subheader("Pola Waktu Terjadinya Gangguan IT")
                h_dist = df_filtered['Hour'].dropna().value_counts().reindex(range(24), fill_value=0).reset_index()
                fig_h = px.line(h_dist, x='Hour', y='count', markers=True, 
                                 labels={'count':'Jumlah', 'Hour':'Jam (WIB)'})
                fig_h.update_traces(line_color='#E74C3C', line_width=3)
                fig_h.update_layout(xaxis=dict(tickmode='linear', showline=True, linecolor='black'),
                                    yaxis=dict(showline=True, linecolor='black'))
                st.plotly_chart(fig_h, use_container_width=True)

                # VISUALISASI 4: Kalender
                if selected_month != "Semua Bulan":
                    st.subheader(f"Intensitas Gangguan IT Harian pada Bulan: {selected_month}")
                    
                    month_idx = list_bulan.index(selected_month) - 1 
                    cal = calendar.monthcalendar(selected_year, month_idx)
                    daily = df_filtered.groupby(df_filtered['Tanggal'].dt.day).size()
                    
                    z, day_numbers, hover_texts = [], [], []
                    
                    for week in cal:
                        z_week, day_week, hover_week = [], [], []
                        for day in week:
                            if day == 0:
                                z_week.append(np.nan) 
                                day_week.append("")
                                hover_week.append("")
                            else:
                                count = daily.get(day, 0)
                                z_week.append(count)
                                day_week.append(str(day))
                                hover_week.append(f"Tanggal: {day} {selected_month}<br>Total Kasus: {count}")
                        z.append(z_week)
                        day_numbers.append(day_week)
                        hover_texts.append(hover_week)

                    fig_cal = go.Figure(data=go.Heatmap(
                        z=z,
                        x=['Senin', 'Selasa', 'Rabu', 'Kamis', 'Jumat', 'Sabtu', 'Minggu'],
                        y=[f"Minggu {i+1}" for i in range(len(z))][::-1],
                        colorscale='YlOrRd',
                        showscale=False,
                        xgap=3, 
                        ygap=3,
                        customdata=hover_texts,
                        hovertemplate="%{customdata}<extra></extra>"
                    ))

                    fig_cal.add_trace(go.Scatter(
                        x=['Senin', 'Selasa', 'Rabu', 'Kamis', 'Jumat', 'Sabtu', 'Minggu'] * len(z),
                        y=np.repeat([f"Minggu {i+1}" for i in range(len(z))][::-1], 7),
                        mode='text',
                        text=np.array(day_numbers).flatten(),
                        textfont=dict(size=16, color="black", family="Arial Black"),
                        hoverinfo='skip'
                    ))

                    fig_cal.update_layout(
                        height=450,
                        margin=dict(l=10, r=10, t=30, b=10),
                        xaxis=dict(side='top', fixedrange=True),
                        yaxis=dict(showticklabels=False, fixedrange=True, autorange="reversed"),
                        plot_bgcolor='rgba(0,0,0,0)',
                        paper_bgcolor='rgba(0,0,0,0)',
                        showlegend=False
                    )
                    st.plotly_chart(fig_cal, use_container_width=True) 
                    
                # --- 4. ANALISIS INSIGHT (DI AKHIR DASHBOARD) ---
                st.markdown("---")
                st.subheader("Analisis Insight")

                # Menyiapkan variabel dasar
                total_skrg = len(df_filtered)
                top_prob = df_filtered['Problem_Clean'].mode()[0] if not df_filtered.empty else "-"
                top_loc = df_filtered['Loc_Clean'].mode()[0] if not df_filtered.empty else "-"
                peak_time = f"{int(df_filtered['Hour'].mode()[0]):02d}:00" if not df_filtered['Hour'].mode().empty else "-"

                # Logic Insight
                if selected_month == "Semua Bulan":
                    # 1. Insight untuk Semua Bulan
                    monthly_counts = df_full.groupby('Bulan_Nama').size()
                    # Mengurutkan berdasarkan urutan bulan yang benar
                    ordered_months = [m for m in list_bulan if m in monthly_counts.index]
                    monthly_counts = monthly_counts.reindex(ordered_months)
                    
                    peak_month_name = monthly_counts.idxmax()
                    peak_month_val = monthly_counts.max()
                    
                    st.markdown(f"""
                    <div style="background-color: #ffffff; padding: 20px; border-radius: 10px; border-left: 5px solid #1f4e78; box-shadow: 2px 2px 5px rgba(0,0,0,0.05);">
                        <h4>Rangkuman Performa Tahunan</h4>
                        <ul>
                            <li>Sepanjang periode laporan, <b>Bulan {peak_month_name}</b> merupakan periode dengan tingkat gangguan tertinggi, yaitu sebanyak <b>{peak_month_val} kasus</b>.</li>
                            <li>Secara akumulatif, lokasi yang paling sering membutuhkan penanganan teknisi adalah <b>{top_loc}</b>.</li>
                            <li>Jam operasional yang paling krusial dengan frekuensi troubleshoot tertinggi terjadi pada pukul <b>{peak_time} WIB</b>.</li>
                        </ul>
                    </div>
                    """, unsafe_allow_html=True)

                else:
                    # 2. Insight untuk Bulan Spesifik (Januari - Desember) + Perbandingan MoM
                    # Mencari data bulan sebelumnya
                    curr_idx = list_bulan.index(selected_month)
                    prev_month_text = ""
                    
                    if curr_idx > 2: # Jika bukan Januari (Januari di index 2)
                        nama_bulan_lalu = list_bulan[curr_idx - 1]
                        total_lalu = len(
                            df_year[df_year['Bulan_Nama'] == nama_bulan_lalu]
                            )

                        
                        selisih = total_skrg - total_lalu
                        if selisih > 0:
                            perubahan_text = f"mengalami <b style='color:red;'>kenaikan</b> sebanyak {selisih} kasus"
                        elif selisih < 0:
                            perubahan_text = f"mengalami <b style='color:green;'>penurunan</b> sebanyak {abs(selisih)} kasus"
                        else:
                            perubahan_text = "stabil (sama dengan bulan sebelumnya)"
                            
                        prev_month_text = f"Jika dibandingkan dengan bulan {nama_bulan_lalu}, jumlah gangguan {perubahan_text}."
                    else:
                        prev_month_text = "Ini merupakan bulan awal dalam data yang diunggah, sehingga belum ada perbandingan dengan bulan sebelumnya."

                    st.markdown(f"""
                    <div style="background-color: #ffffff; padding: 20px; border-radius: 10px; border-left: 5px solid #1f4e78; box-shadow: 2px 2px 5px rgba(0,0,0,0.05);">
                        <h4>Analisis Intensitas Bulan {selected_month}</h4>
                        <p>Berdasarkan dashboard visual di atas, dapat disimpulkan bahwa:</p>
                        <ul>
                            <li>Total aktivitas IT Support pada bulan {selected_month} adalah sebanyak <b>{total_skrg} kasus</b>. {prev_month_text}</li>
                            <li>Distribusi lokasi menunjukkan bahwa <b>{top_loc}</b> adalah area dengan tingkat laporan gangguan tertinggi.</li>
                            <li>Frekuensi gangguan cenderung meningkat pada jam sibuk, terutama pada pukul <b>{peak_time} WIB</b>.</li>
                        </ul>
                    </div>
                    """, unsafe_allow_html=True) 

    except Exception as e:
        st.error(f"Terjadi kesalahan pengolahan data: {e}")
else:
    st.image("https://img.freepik.com/free-vector/data-report-concept-illustration_114360-883.jpg", width=400)
