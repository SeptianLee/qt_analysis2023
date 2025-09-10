import streamlit as st
import pandas as pd
import matplotlib.pyplot as plt
import seaborn as sns
import numpy as np
import os

# Konfigurasi halaman
st.set_page_config(
    page_title="Analisis Data Quotation",
    page_icon="📊",
    layout="wide",
    initial_sidebar_state="expanded"
)

# Judul aplikasi
st.title("📊 Dashboard Analisis Data Quotation")
st.markdown("---")

# Sidebar untuk informasi
with st.sidebar:
    st.header("Informasi Dataset")
    st.info("""
    **Dataset:** QT_REPORT_2023.xlsx
    
    **Lokasi:** Direktori yang sama dengan aplikasi
    
    **Kolom yang digunakan:**
    - DATE: Tanggal quotation
    - TO: Penerima quotation
    - SUBJECT: Subjek quotation
    - QUANTITY: Kuantitas
    """)

# Fungsi untuk memuat data
@st.cache_data
def load_data():
    try:
        # Membaca file dari direktori yang sama
        file_path = "QT_REPORT_2023.xlsx"
        qt = pd.read_excel(file_path)
        
        # Preprocessing data
        qt['DATE'] = pd.to_datetime(qt['DATE'], format='%d/%m/%Y', errors='coerce')
        qt['Year_Month'] = qt['DATE'].dt.to_period('M')
        
        return qt
    except FileNotFoundError:
        st.error("File QT_REPORT_2023.xlsx tidak ditemukan di direktori yang sama dengan aplikasi.")
        return None
    except Exception as e:
        st.error(f"Terjadi kesalahan saat memuat data: {str(e)}")
        return None

    # --- 3. bersihkan & normalisasi kolom QUANTITY ---
    # simpan raw uppercase untuk deteksi keywords
    qt['QUANTITY_raw'] = qt['QUANTITY'].astype(str).str.strip()
    qt['QUANTITY_up'] = qt['QUANTITY_raw'].str.upper()

    # ekstrak angka (misal "23 unit" -> 23). Jika tidak ada angka -> NaN
    qt['quantity_num'] = qt['QUANTITY_up'].str.extract(r'(\d+)').astype(float)

    # deteksi kata 'VARIOUS' atau variasi (anggap kata 'VARIOUS' menandakan non-numeric)
    qt['is_various'] = qt['QUANTITY_up'].str.contains('VARIOUS', na=False)

    # deteksi kata 'UNIT' bila perlu
    qt['has_unit_word'] = qt['QUANTITY_up'].str.contains(r'\bUNIT\b', na=False)

    # set kategori quantity_type: 'numeric' / 'various' / 'other'
    qt['quantity_type'] = np.where(qt['quantity_num'].notna(), 'numeric',
                            np.where(qt['is_various'], 'various', 'other'))


# Fungsi untuk membuat visualisasi
def create_monthly_chart(monthly_counts):
    if len(monthly_counts) == 0:
        return None, None
    
    # Format label bulan
    month_labels = [period.strftime('%b-%Y') for period in monthly_counts.index]
    
    # Buat figure
    fig, ax = plt.subplots(figsize=(10, 5))
    
    # Buat bar chart
    bars = ax.bar(range(len(monthly_counts)), monthly_counts.values, 
                  color=plt.cm.viridis(np.linspace(0, 1, len(monthly_counts))),
                  edgecolor='black', linewidth=0.5, alpha=0.8)
    
    # Tambahkan nilai di atas setiap bar
    for i, bar in enumerate(bars):
        height = bar.get_height()
        ax.text(bar.get_x() + bar.get_width()/2., height + 0.05,
                f'{int(height)}', ha='center', va='bottom', fontweight='bold')
    
    # Customize chart
    ax.set_title('Jumlah Quotation per Bulan', fontsize=14, fontweight='bold', pad=15)
    ax.set_xlabel('Bulan', fontsize=10, fontweight='bold')
    ax.set_ylabel('Jumlah Quotation', fontsize=10, fontweight='bold')
    
    # Atur ticks dan labels
    ax.set_xticks(range(len(monthly_counts)))
    ax.set_xticklabels(month_labels, rotation=45, ha='right')
    
    # Tambahkan grid
    ax.yaxis.grid(True, linestyle='--', alpha=0.7)
    
    # Hilangkan spines yang tidak diperlukan
    ax.spines[['top', 'right']].set_visible(False)
    
    return fig, month_labels

def create_trend_chart(monthly_counts, month_labels):
    if len(monthly_counts) == 0:
        return None
    
    fig, ax = plt.subplots(figsize=(10, 5))
    ax.plot(month_labels, monthly_counts.values, marker='o', 
             linewidth=2.5, markersize=8, color='steelblue')
    ax.set_title('Tren Quotation per Bulan', fontsize=14, fontweight='bold', pad=15)
    ax.set_xlabel('Bulan', fontsize=10, fontweight='bold')
    ax.set_ylabel('Jumlah Quotation', fontsize=10, fontweight='bold')
    ax.tick_params(axis='x', rotation=45)
    ax.grid(True, linestyle='--', alpha=0.7)
    
    # Tambahkan anotasi untuk nilai maksimum
    max_idx = monthly_counts.values.argmax()
    ax.annotate(f'Puncak: {monthly_counts.values[max_idx]}', 
                xy=(max_idx, monthly_counts.values[max_idx]),
                xytext=(max_idx, monthly_counts.values[max_idx] + 0.5),
                arrowprops=dict(arrowstyle='->', color='red'),
                ha='center', color='red', fontweight='bold')
    
    return fig

def create_top10_to_chart(top_to):
    if top_to.empty:
        return None
    
    fig, ax = plt.subplots(figsize=(10, 6))
    top_10_to = top_to.nlargest(10).sort_values()
    bars = ax.barh(range(len(top_10_to)), top_10_to.values, 
                   color=plt.cm.viridis(np.linspace(0.2, 0.8, 10)),
                   edgecolor='black', alpha=0.85, linewidth=0.7)
    
    # Add value labels
    max_val = top_10_to.max()
    for i, v in enumerate(top_10_to.values):
        if v > max_val * 0.15:
            ax.text(v - (max_val * 0.02), i, f'{v:,.0f}', 
                    va='center', fontweight='bold', color='white', fontsize=10)
        else:
            ax.text(v + (max_val * 0.01), i, f'{v:,.0f}', 
                    va='center', fontweight='bold', color='black', fontsize=10)
    
    # Styling
    ax.set_title('Top 10 Penerima Quotation', fontsize=14, fontweight='bold', pad=15)
    ax.set_xlabel('Jumlah Quotation', fontsize=10, fontweight='bold')
    ax.set_ylabel('Penerima (TO)', fontsize=10, fontweight='bold')
    
    # Set y-tick labels
    ax.set_yticks(range(len(top_10_to)))
    ax.set_yticklabels(top_10_to.index)
    
    # Remove chart borders
    for spine in ['top', 'right', 'bottom']:
        ax.spines[spine].set_visible(False)
    ax.spines['left'].set_visible(True)
    ax.spines['left'].set_color('#CCCCCC')
    
    # Add grid
    ax.set_axisbelow(True)
    ax.xaxis.grid(True, linestyle='--', alpha=0.4, color='#AAAAAA')
    
    return fig

def create_top10_subject_chart(top_subject):
    if top_subject.empty:
        return None
    
    # Urutkan nilai secara descending untuk visualisasi yang lebih baik
    top_subject_sorted = top_subject.nlargest(10).sort_values(ascending=True)
    
    fig, ax = plt.subplots(figsize=(10, 6))
    bars = ax.barh(range(len(top_subject_sorted)), top_subject_sorted.values, 
                   color=plt.cm.Blues(np.linspace(0.4, 0.8, len(top_subject_sorted))),
                   edgecolor='black', linewidth=0.5, alpha=0.8)
    
    # Tambahkan nilai di ujung setiap bar
    for i, (value, name) in enumerate(zip(top_subject_sorted.values, top_subject_sorted.index)):
        ax.text(value + max(top_subject_sorted.values)*0.01, i, 
                f'{int(value)}', va='center', fontweight='bold', fontsize=10)
    
    # Customize chart
    ax.set_title('Top 10 SUBJECT Quotation', fontsize=14, fontweight='bold', pad=15)
    ax.set_xlabel('Jumlah Quotation', fontsize=10, fontweight='bold')
    ax.set_ylabel('Subject', fontsize=10, fontweight='bold')
    
    # Atur y-ticks dan labels
    ax.set_yticks(range(len(top_subject_sorted)))
    ax.set_yticklabels(top_subject_sorted.index, fontsize=10)
    
    # Tambahkan grid
    ax.xaxis.grid(True, linestyle='--', alpha=0.7)
    
    # Hilangkan spines yang tidak diperlukan
    ax.spines[['top', 'right']].set_visible(False)
    
    # Atur batas x-axis untuk ruang teks nilai
    ax.set_xlim(0, max(top_subject_sorted.values) * 1.15)
    
    return fig, top_subject_sorted

def create_quantity_chart(qty_type_counts):
    if qty_type_counts.empty:
        return None
    
    fig, (ax1, ax2) = plt.subplots(1, 2, figsize=(12, 5))
    
    # Bar chart
    colors = ['#3498db', '#2ecc71', '#e74c3c', '#f39c12', '#9b59b6']
    bars = ax1.bar(range(len(qty_type_counts)), qty_type_counts.values, 
                   color=colors[:len(qty_type_counts)], 
                   edgecolor='black', alpha=0.8, linewidth=0.8)
    
    # Add value labels on bars
    for i, bar in enumerate(bars):
        height = bar.get_height()
        ax1.text(bar.get_x() + bar.get_width()/2., height + max(qty_type_counts.values)*0.01,
                f'{int(height)}', ha='center', va='bottom', fontweight='bold')
    
    # Customize bar chart
    ax1.set_title('Distribusi Tipe QUANTITY', fontsize=12, fontweight='bold', pad=15)
    ax1.set_xlabel('Tipe QUANTITY', fontsize=10, fontweight='bold')
    ax1.set_ylabel('Jumlah', fontsize=10, fontweight='bold')
    ax1.set_xticks(range(len(qty_type_counts)))
    ax1.set_xticklabels(qty_type_counts.index, rotation=45, ha='right')
    ax1.yaxis.grid(True, linestyle='--', alpha=0.7)
    ax1.spines[['top', 'right']].set_visible(False)
    
    # Pie chart
    wedges, texts, autotexts = ax2.pie(qty_type_counts.values, 
                                       labels=qty_type_counts.index, 
                                       autopct='%1.1f%%',
                                       colors=colors[:len(qty_type_counts)],
                                       startangle=90,
                                       textprops={'fontsize': 9})
    
    # Customize pie chart
    ax2.set_title('Proporsi Tipe QUANTITY', fontsize=12, fontweight='bold', pad=15)
    
    # Make autopct text more readable
    for autotext in autotexts:
        autotext.set_color('white')
        autotext.set_fontweight('bold')
    
    return fig

# Memuat data
qt = load_data()

if qt is not None:
    # Hitung berbagai metrik untuk visualisasi
    monthly_counts = qt['Year_Month'].value_counts().sort_index()
    
    # Buat rentang bulan lengkap
    start_date = qt['DATE'].min().replace(day=1)
    end_date = qt['DATE'].max().replace(day=1)
    all_months = pd.period_range(start=start_date, end=end_date, freq='M')
    monthly_counts = monthly_counts.reindex(all_months, fill_value=0)
    
    # Hitung metrik lainnya
    top_to = qt['TO'].value_counts()
    top_subject = qt['SUBJECT '].value_counts()  # Perbaiki nama kolom (hapus spasi jika ada)
    qty_type_counts = qt['QUANTITY'].value_counts()
    
    st.markdown("---")
    
    # Visualisasi 1: Jumlah Quotation per Bulan
    st.subheader("📈 Jumlah Quotation per Bulan")
    
    monthly_fig, month_labels = create_monthly_chart(monthly_counts)
    if monthly_fig:
        st.pyplot(monthly_fig)
        
        # Statistik Bulanan
        max_idx = monthly_counts.values.argmax()
        col1, col2, col3 = st.columns(3)
        col1.metric("Rata-rata per Bulan", f"{monthly_counts.mean():.2f}")
        col2.metric("Bulan Tertinggi", f"{month_labels[max_idx]} ({monthly_counts.values[max_idx]})")
        col3.metric("Total Quotation", f"{monthly_counts.sum()}")
        
        # Tren Quotation per Bulan (Line Plot)
        st.subheader("📊 Tren Quotation per Bulan")
        trend_fig = create_trend_chart(monthly_counts, month_labels)
        if trend_fig:
            st.pyplot(trend_fig)
    else:
        st.warning("Tidak ada data bulanan yang valid untuk divisualisasikan.")
    
    st.markdown("---")
    
    # Visualisasi 2: Top 10 TO
    st.subheader("👥 Top 10 Penerima Quotation")
    
    top_to_fig = create_top10_to_chart(top_to)
    if top_to_fig:
        st.pyplot(top_to_fig)
        
        # Additional summary
        total_top10 = top_to.nlargest(10).sum()
        total_all = top_to.sum()
        percentage = (total_top10 / total_all) * 100
        
        col1, col2 = st.columns(2)
        col1.metric("Total Quotation Top 10", f"{total_top10:,.0f}")
        col2.metric("Persentase dari Total", f"{percentage:.1f}%")
    else:
        st.warning("Tidak ada data penerima quotation yang valid.")
    
    st.markdown("---")
    
    # Visualisasi 3: Top 10 Subject
    st.subheader("🏷️ Top 10 Subject Quotation")
    
    top_subject_fig, top_subject_sorted = create_top10_subject_chart(top_subject)
    if top_subject_fig:
        st.pyplot(top_subject_fig)
        
        # Tampilkan persentase dari total
        total_subjects = top_subject.sum()
        st.write("**Distribusi Top 10 Subject:**")
        for i, (name, value) in enumerate(top_subject_sorted.items(), 1):
            percentage = (value / total_subjects) * 100
            st.write(f"{i}. {name}: {value} quotation ({percentage:.1f}%)")
    else:
        st.warning("Tidak ada data subject yang valid.")
    
    st.markdown("---")
    
    # Visualisasi 4: Distribusi Tipe QUANTITY
    st.subheader("🔢 Distribusi Tipe QUANTITY")
    
    quantity_fig = create_quantity_chart(qty_type_counts)
    if quantity_fig:
        st.pyplot(quantity_fig)
        
        # Display summary statistics
        st.write("**Statistik Distribusi Tipe QUANTITY:**")
        total = qty_type_counts.sum()
        for i, (qty_type, count) in enumerate(qty_type_counts.items()):
            percentage = (count / total) * 100
            st.write(f"{i+1}. {qty_type}: {count} ({percentage:.1f}%)")
    else:
        st.warning("Tidak ada data quantity yang valid.")

# Footer
st.markdown("---")
st.markdown(
    """
    <style>
    .footer {
        text-align: center;
        padding: 10px;
    }
    </style>
    <div class="footer">
        <p>Dashboard Analisis Data Quotation • Dibuat dengan Streamlit</p>
    </div>
    """,
    unsafe_allow_html=True
)




