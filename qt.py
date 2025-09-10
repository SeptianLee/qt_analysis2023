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
    top_subject = qt['SUBJECT '].value_counts()
    qty_type_counts = qt['QUANTITY'].value_counts()
    
    # Tampilkan info dataset
    st.subheader("📋 Informasi Dataset")
    col1, col2, col3, col4 = st.columns(4)
    col1.metric("Jumlah Data", f"{len(qt):,}")
    col2.metric("Periode Awal", qt['DATE'].min().strftime('%d %b %Y'))
    col3.metric("Periode Akhir", qt['DATE'].max().strftime('%d %b %Y'))
    col4.metric("Jumlah Bulan", f"{len(monthly_counts)}")
    
    st.markdown("---")
    
    # Visualisasi 1: Jumlah Quotation per Bulan
    st.subheader("📈 Jumlah Quotation per Bulan")
    
    if len(monthly_counts) > 0:
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
        
        # Tampilkan chart di Streamlit
        st.pyplot(fig)
        
        # Statistik Bulanan
        max_idx = monthly_counts.values.argmax()
        col1, col2, col3 = st.columns(3)
        col1.metric("Rata-rata per Bulan", f"{monthly_counts.mean():.2f}")
        col2.metric("Bulan Tertinggi", f"{month_labels[max_idx]} ({monthly_counts.values[max_idx]})")
        col3.metric("Total Quotation", f"{monthly_counts.sum()}")
        
        # Tren Quotation per Bulan (Line Plot)
        st.subheader("📊 Tren Quotation per Bulan")
        
        fig2, ax2 = plt.subplots(figsize=(10, 5))
        ax2.plot(month_labels, monthly_counts.values, marker='o', 
                 linewidth=2.5, markersize=8, color='steelblue')
        ax2.set_title('Tren Quotation per Bulan', fontsize=14, fontweight='bold', pad=15)
        ax2.set_xlabel('Bulan', fontsize=10, fontweight='bold')
        ax2.set_ylabel('Jumlah Quotation', fontsize=10, fontweight='bold')
        ax2.tick_params(axis='x', rotation=45)
        ax2.grid(True, linestyle='--', alpha=0.7)
        
        # Tambahkan anotasi untuk nilai maksimum
        ax2.annotate(f'Puncak: {monthly_counts.values[max_idx]}', 
                    xy=(max_idx, monthly_counts.values[max_idx]),
                    xytext=(max_idx, monthly_counts.values[max_idx] + 0.5),
                    arrowprops=dict(arrowstyle='->', color='red'),
                    ha='center', color='red', fontweight='bold')
        
        st.pyplot(fig2)
        
    else:
        st.warning("Tidak ada data bulanan yang valid untuk divisualisasikan.")
    
    st.markdown("---")
    
    # Visualisasi 2: Top 10 TO
    st.subheader("👥 Top 10 Penerima Quotation")
    
    fig3, ax3 = plt.subplots(figsize=(10, 6))
    top_10_to = top_to.nlargest(10).sort_values()
    bars = ax3.barh(range(len(top_10_to)), top_10_to.values, 
                   color=plt.cm.viridis(np.linspace(0.2, 0.8, 10)),
                   edgecolor='black', alpha=0.85, linewidth=0.7)
    
    # Add value labels
    max_val = top_10_to.max()
    for i, v in enumerate(top_10_to.values):
        if v > max_val * 0.15:
            ax3.text(v - (max_val * 0.02), i, f'{v:,.0f}', 
                    va='center', fontweight='bold', color='white', fontsize=10)
        else:
            ax3.text(v + (max_val * 0.01), i, f'{v:,.0f}', 
                    va='center', fontweight='bold', color='black', fontsize=10)
    
    # Styling
    ax3.set_title('Top 10 Penerima Quotation', fontsize=14, fontweight='bold', pad=15)
    ax3.set_xlabel('Jumlah Quotation', fontsize=10, fontweight='bold')
    ax3.set_ylabel('Penerima (TO)', fontsize=10, fontweight='bold')
    
    # Set y-tick labels
    ax3.set_yticks(range(len(top_10_to)))
    ax3.set_yticklabels(top_10_to.index)
    
    # Remove chart borders
    for spine in ['top', 'right', 'bottom']:
        ax3.spines[spine].set_visible(False)
    ax3.spines['left'].set_visible(True)
    ax3.spines['left'].set_color('#CCCCCC')
    
    # Add grid
    ax3.set_axisbelow(True)
    ax3.xaxis.grid(True, linestyle='--', alpha=0.4, color='#AAAAAA')
    
    st.pyplot(fig3)
    
    # Additional summary
    total_top10 = top_to.nlargest(10).sum()
    total_all = top_to.sum()
    percentage = (total_top10 / total_all) * 100
    
    col1, col2 = st.columns(2)
    col1.metric("Total Quotation Top 10", f"{total_top10:,.0f}")
    col2.metric("Persentase dari Total", f"{percentage:.1f}%")
    
    st.markdown("---")
    
    # Visualisasi 3: Top 10 Subject
    st.subheader("🏷️ Top 10 Subject Quotation")
    
    if not top_subject.empty:
        # Urutkan nilai secara descending untuk visualisasi yang lebih baik
        top_subject_sorted = top_subject.nlargest(10).sort_values(ascending=True)
        
        fig4, ax4 = plt.subplots(figsize=(10, 6))
        bars = ax4.barh(range(len(top_subject_sorted)), top_subject_sorted.values, 
                       color=plt.cm.Blues(np.linspace(0.4, 0.8, len(top_subject_sorted))),
                       edgecolor='black', linewidth=0.5, alpha=0.8)
        
        # Tambahkan nilai di ujung setiap bar
        for i, (value, name) in enumerate(zip(top_subject_sorted.values, top_subject_sorted.index)):
            ax4.text(value + max(top_subject_sorted.values)*0.01, i, 
                    f'{int(value)}', va='center', fontweight='bold', fontsize=10)
        
        # Customize chart
        ax4.set_title('Top 10 SUBJECT Quotation', fontsize=14, fontweight='bold', pad=15)
        ax4.set_xlabel('Jumlah Quotation', fontsize=10, fontweight='bold')
        ax4.set_ylabel('Subject', fontsize=10, fontweight='bold')
        
        # Atur y-ticks dan labels
        ax4.set_yticks(range(len(top_subject_sorted)))
        ax4.set_yticklabels(top_subject_sorted.index, fontsize=10)
        
        # Tambahkan grid
        ax4.xaxis.grid(True, linestyle='--', alpha=0.7)
        
        # Hilangkan spines yang tidak diperlukan
        ax4.spines[['top', 'right']].set_visible(False)
        
        # Atur batas x-axis untuk ruang teks nilai
        ax4.set_xlim(0, max(top_subject_sorted.values) * 1.15)
        
        st.pyplot(fig4)
        
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
    
    fig5, (ax5, ax6) = plt.subplots(1, 2, figsize=(12, 5))
    
    # Bar chart
    colors = ['#3498db', '#2ecc71', '#e74c3c', '#f39c12', '#9b59b6']
    bars = ax5.bar(range(len(qty_type_counts)), qty_type_counts.values, 
                   color=colors[:len(qty_type_counts)], 
                   edgecolor='black', alpha=0.8, linewidth=0.8)
    
    # Add value labels on bars
    for i, bar in enumerate(bars):
        height = bar.get_height()
        ax5.text(bar.get_x() + bar.get_width()/2., height + max(qty_type_counts.values)*0.01,
                f'{int(height)}', ha='center', va='bottom', fontweight='bold')
    
    # Customize bar chart
    ax5.set_title('Distribusi Tipe QUANTITY', fontsize=12, fontweight='bold', pad=15)
    ax5.set_xlabel('Tipe QUANTITY', fontsize=10, fontweight='bold')
    ax5.set_ylabel('Jumlah', fontsize=10, fontweight='bold')
    ax5.set_xticks(range(len(qty_type_counts)))
    ax5.set_xticklabels(qty_type_counts.index, rotation=45, ha='right')
    ax5.yaxis.grid(True, linestyle='--', alpha=0.7)
    ax5.spines[['top', 'right']].set_visible(False)
    
    # Pie chart
    wedges, texts, autotexts = ax6.pie(qty_type_counts.values, 
                                       labels=qty_type_counts.index, 
                                       autopct='%1.1f%%',
                                       colors=colors[:len(qty_type_counts)],
                                       startangle=90,
                                       textprops={'fontsize': 9})
    
    # Customize pie chart
    ax6.set_title('Proporsi Tipe QUANTITY', fontsize=12, fontweight='bold', pad=15)
    
    # Make autopct text more readable
    for autotext in autotexts:
        autotext.set_color('white')
        autotext.set_fontweight('bold')
    
    st.pyplot(fig5)
    
    # Display summary statistics
    st.write("**Statistik Distribusi Tipe QUANTITY:**")
    total = qty_type_counts.sum()
    for i, (qty_type, count) in enumerate(qty_type_counts.items()):
        percentage = (count / total) * 100
        st.write(f"{i+1}. {qty_type}: {count} ({percentage:.1f}%)")

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
