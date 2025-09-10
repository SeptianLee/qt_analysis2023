import streamlit as st
import pandas as pd
import matplotlib.pyplot as plt
import seaborn as sns
import numpy as np
import re
import os

# -------------------------------
# Konfigurasi halaman
# -------------------------------
st.set_page_config(
    page_title="Analisis Data Quotation",
    page_icon="📊",
    layout="wide",
    initial_sidebar_state="expanded"
)

st.title("📊 Dashboard Analisis Data Quotation")
st.markdown("---")

# -------------------------------
# Sidebar untuk informasi
# -------------------------------
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

# -------------------------------
# Fungsi untuk memuat data
# -------------------------------
@st.cache_data
def load_data():
    try:
        file_path = "QT_REPORT_2023.xlsx"
        qt = pd.read_excel(file_path)

        # Rapikan nama kolom
        if isinstance(qt.columns, pd.MultiIndex):
            qt.columns = qt.columns.get_level_values(0).str.strip()
        else:
            qt.columns = qt.columns.str.strip()

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

# -------------------------------
# Fungsi ekstrak QUANTITY
# -------------------------------
def extract_quantity(text):
    if pd.isna(text): 
        return []
    text = str(text).lower()
    pattern = r'(\d+)\s*(unit|ea|set|pcs|orang|hari|each|ton)?'
    return [{'num': int(n), 'unit': u if u else 'unspecified'} 
            for n, u in re.findall(pattern, text)]

# -------------------------------
# Fungsi untuk expand data QUANTITY
# -------------------------------
def expand_data(qt):
    expanded = []
    for _, row in qt.iterrows():
        for q in extract_quantity(row['QUANTITY']):
            expanded.append({
                'DATE': row['DATE'],
                'TO': row['TO'],
                'SUBJECT': row['SUBJECT'] if 'SUBJECT' in qt.columns else None,
                'num': q['num'],
                'unit': q['unit']
            })
    return pd.DataFrame(expanded)

# -------------------------------
# Fungsi visualisasi
# -------------------------------
def create_monthly_chart(monthly_counts):
    if len(monthly_counts) == 0:
        return None, None
    month_labels = [period.strftime('%b-%Y') for period in monthly_counts.index]
    fig, ax = plt.subplots(figsize=(10, 5))
    bars = ax.bar(range(len(monthly_counts)), monthly_counts.values,
                  color=plt.cm.viridis(np.linspace(0, 1, len(monthly_counts))),
                  edgecolor='black', linewidth=0.5, alpha=0.8)
    for i, bar in enumerate(bars):
        height = bar.get_height()
        ax.text(bar.get_x() + bar.get_width()/2., height + 0.05,
                f'{int(height)}', ha='center', va='bottom', fontweight='bold')
    ax.set_title('Jumlah Quotation per Bulan', fontsize=14, fontweight='bold', pad=15)
    ax.set_xlabel('Bulan', fontsize=10, fontweight='bold')
    ax.set_ylabel('Jumlah Quotation', fontsize=10, fontweight='bold')
    ax.set_xticks(range(len(monthly_counts)))
    ax.set_xticklabels(month_labels, rotation=45, ha='right')
    ax.yaxis.grid(True, linestyle='--', alpha=0.7)
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
    ax.barh(top_10_to.index, top_10_to.values,
            color=plt.cm.viridis(np.linspace(0.2, 0.8, 10)),
            edgecolor='black', alpha=0.85, linewidth=0.7)
    ax.set_title('Top 10 Penerima Quotation', fontsize=14, fontweight='bold', pad=15)
    ax.set_xlabel('Jumlah Quotation', fontsize=10, fontweight='bold')
    ax.set_ylabel('Penerima (TO)', fontsize=10, fontweight='bold')
    ax.xaxis.grid(True, linestyle='--', alpha=0.4, color='#AAAAAA')
    return fig

def create_top10_subject_chart(top_subject):
    if top_subject.empty:
        return None, None
    top_subject_sorted = top_subject.nlargest(10).sort_values(ascending=True)
    fig, ax = plt.subplots(figsize=(10, 6))
    ax.barh(top_subject_sorted.index, top_subject_sorted.values,
            color=plt.cm.Blues(np.linspace(0.4, 0.8, len(top_subject_sorted))),
            edgecolor='black', linewidth=0.5, alpha=0.8)
    ax.set_title('Top 10 SUBJECT Quotation', fontsize=14, fontweight='bold', pad=15)
    ax.set_xlabel('Jumlah Quotation', fontsize=10, fontweight='bold')
    ax.set_ylabel('Subject', fontsize=10, fontweight='bold')
    ax.xaxis.grid(True, linestyle='--', alpha=0.7)
    return fig, top_subject_sorted

def create_quantity_chart(qty_type_counts):
    if qty_type_counts.empty:
        return None
    fig, (ax1, ax2) = plt.subplots(1, 2, figsize=(12, 5))
    colors = ['#3498db', '#2ecc71', '#e74c3c', '#f39c12', '#9b59b6']
    ax1.bar(qty_type_counts.index, qty_type_counts.values,
            color=colors[:len(qty_type_counts)], edgecolor='black', alpha=0.8, linewidth=0.8)
    ax1.set_title('Distribusi Tipe QUANTITY', fontsize=12, fontweight='bold', pad=15)
    ax1.set_xlabel('Tipe QUANTITY', fontsize=10, fontweight='bold')
    ax1.set_ylabel('Jumlah', fontsize=10, fontweight='bold')
    ax1.yaxis.grid(True, linestyle='--', alpha=0.7)
    wedges, _, autotexts = ax2.pie(qty_type_counts.values,
                                   labels=qty_type_counts.index,
                                   autopct='%1.1f%%',
                                   colors=colors[:len(qty_type_counts)],
                                   startangle=90,
                                   textprops={'fontsize': 9})
    for autotext in autotexts:
        autotext.set_color('white')
        autotext.set_fontweight('bold')
    ax2.set_title('Proporsi Tipe QUANTITY', fontsize=12, fontweight='bold', pad=15)
    return fig

# -------------------------------
# MAIN APP
# -------------------------------
qt = load_data()

if qt is not None:
    qt_expanded = expand_data(qt)

    # Ringkasan
    unit_summary = qt_expanded.groupby('unit')['num'].sum().sort_values(ascending=False)
    monthly_qty = (qt_expanded[qt_expanded['unit'] == 'unit']
                   .groupby(qt_expanded['DATE'].dt.to_period('M'))['num']
                   .sum().to_timestamp())
    top_to = qt['TO'].value_counts()
    top_subject = qt['SUBJECT'].value_counts() if 'SUBJECT' in qt.columns else pd.Series(dtype=int)
    qty_type_counts = qt['QUANTITY'].value_counts()

    st.subheader("📈 Jumlah Quotation per Bulan")
    monthly_fig, month_labels = create_monthly_chart(monthly_qty)
    if monthly_fig:
        st.pyplot(monthly_fig)
        trend_fig = create_trend_chart(monthly_qty, month_labels)
        if trend_fig: st.pyplot(trend_fig)

    st.subheader("👥 Top 10 Penerima Quotation")
    top_to_fig = create_top10_to_chart(top_to)
    if top_to_fig: st.pyplot(top_to_fig)

    st.subheader("🏷️ Top 10 Subject Quotation")
    top_subject_fig, top_subject_sorted = create_top10_subject_chart(top_subject)
    if top_subject_fig: st.pyplot(top_subject_fig)

    st.subheader("🔢 Distribusi Tipe QUANTITY")
    quantity_fig = create_quantity_chart(qty_type_counts)
    if quantity_fig: st.pyplot(quantity_fig)

st.markdown("---")
st.markdown(
    """
    <div style='text-align: center; padding: 10px;'>
        <p>Dashboard Analisis Data Quotation • Dibuat dengan Streamlit</p>
    </div>
    """,
    unsafe_allow_html=True
)
