"""
Streamlit Web Application: COA Lost & Found Operations Hub & BI Monitoring Dashboard
Designed for Team COA (Custodian of Assets) — HCM Warehouse Inventory Audit
"""

import os
import io
import re
import json
from datetime import datetime, timedelta
import numpy as np
import pandas as pd
import streamlit as st
import plotly.express as px
import plotly.graph_objects as go

# -----------------------------------------------------------------------------
# 1. PAGE CONFIGURATION & STYLING (GRAB BRAND IDENTITY)
# -----------------------------------------------------------------------------
st.set_page_config(
    page_title="Grab L&F Control Tower | Grab Vietnam COA",
    page_icon="🟢",
    layout="wide",
    initial_sidebar_state="expanded"
)

# Custom Grab Corporate CSS
st.markdown("""
<style>
    @import url('https://fonts.googleapis.com/css2?family=Inter:wght@400;500;600;700;800&display=swap');
    
    html, body, [class*="css"] {
        font-family: 'Inter', sans-serif;
    }
    
    /* Grab Header Banner */
    .main-header {
        background: linear-gradient(135deg, #00B14F 0%, #00843D 50%, #064E3B 100%);
        color: white;
        padding: 26px 36px;
        border-radius: 14px;
        margin-bottom: 24px;
        box-shadow: 0 10px 25px -5px rgba(0, 177, 79, 0.25), 0 8px 10px -6px rgba(0, 177, 79, 0.2);
        border-bottom: 4px solid #006837;
    }
    
    .main-header h1 {
        margin: 8px 0 4px 0;
        font-size: 28px;
        font-weight: 800;
        color: #ffffff;
        letter-spacing: -0.5px;
    }
    
    .main-header p {
        margin: 0;
        color: #E8F8F0;
        font-size: 14px;
        font-weight: 500;
    }
    
    /* Grab Badge Pills */
    .badge-pill {
        display: inline-block;
        padding: 4px 12px;
        border-radius: 9999px;
        font-size: 11px;
        font-weight: 700;
        letter-spacing: 0.5px;
        margin-right: 8px;
        text-transform: uppercase;
    }
    
    .badge-grab { background-color: #ffffff; color: #00843D; border: 1px solid #ffffff; }
    .badge-hub { background-color: rgba(255, 255, 255, 0.2); color: #ffffff; border: 1px solid rgba(255, 255, 255, 0.4); }
    .badge-success { background-color: #E8F8F0; color: #00843D; border: 1px solid #00B14F; }
    .badge-warning { background-color: #FEF3C7; color: #B45309; border: 1px solid #F59E0B; }
    .badge-danger { background-color: #FEE2E2; color: #B91C1C; border: 1px solid #EF4444; }
    .badge-info { background-color: #E0F2FE; color: #0369A1; border: 1px solid #0284C7; }
    
    /* Grab Metric Cards */
    .metric-card {
        background: #ffffff;
        border: 1px solid #E2E8F0;
        border-top: 4px solid #00B14F;
        border-radius: 12px;
        padding: 18px 20px;
        box-shadow: 0 2px 4px rgba(0,0,0,0.04);
        transition: transform 0.2s ease, box-shadow 0.2s ease, border-color 0.2s ease;
    }
    .metric-card:hover {
        transform: translateY(-3px);
        box-shadow: 0 10px 15px -3px rgba(0, 177, 79, 0.15);
        border-top-color: #00843D;
    }
    .metric-card-danger {
        border-top: 4px solid #EF4444;
    }
    .metric-card-warning {
        border-top: 4px solid #F59E0B;
    }
    .metric-card-blue {
        border-top: 4px solid #0284C7;
    }
    .metric-card-purple {
        border-top: 4px solid #8B5CF6;
    }
    .metric-title {
        font-size: 12px;
        font-weight: 700;
        color: #64748B;
        text-transform: uppercase;
        letter-spacing: 0.6px;
    }
    .metric-value {
        font-size: 28px;
        font-weight: 800;
        color: #0F172A;
        margin-top: 4px;
        letter-spacing: -0.5px;
    }
    .metric-subtitle {
        font-size: 12px;
        color: #64748B;
        margin-top: 4px;
    }
    
    /* Grab Streamlit Tabs */
    .stTabs [data-baseweb="tab-list"] {
        gap: 8px;
        border-bottom: 2px solid #E2E8F0;
    }
    .stTabs [data-baseweb="tab"] {
        padding: 12px 20px;
        font-weight: 700;
        color: #475569;
        border-radius: 8px 8px 0 0;
    }
    .stTabs [aria-selected="true"] {
        color: #00843D !important;
        border-bottom: 3px solid #00B14F !important;
        background-color: #F0FDF4 !important;
    }
    
    /* Grab Custom Primary Buttons */
    div.stButton > button:first-child {
        border-radius: 8px;
        font-weight: 600;
        transition: all 0.15s ease;
    }
    div.stButton > button:first-child:hover {
        border-color: #00B14F;
        color: #00843D;
    }
</style>
""", unsafe_allow_html=True)


# -----------------------------------------------------------------------------
# 2. DATA PIPELINE & CACHE
# -----------------------------------------------------------------------------
@st.cache_data
def load_default_cleaned_data():
    """Tải dữ liệu đã làm sạch chuẩn từ CSV (Ưu tiên cleaned_data.csv từ notebook)."""
    csv_paths = [
        'cleaned_data.csv',
        os.path.join('usecase', 'cleaned_data.csv'),
        'cleaned_lost_and_found_hcm.csv',
        os.path.join('usecase', 'cleaned_lost_and_found_hcm.csv')
    ]
    for p in csv_paths:
        if os.path.exists(p):
            df = pd.read_csv(p)
            df['cleaned_date'] = pd.to_datetime(df['cleaned_date'], errors='coerce')
            return df
    return None

def process_raw_excel(file_bytes_or_path):
    """Quy trình ETL làm sạch tự động cho file Excel thô mới (Khớp 100% logic Notebook)."""
    xls = pd.ExcelFile(file_bytes_or_path)
    target_sheet = None
    # Tìm sheet chứa dữ liệu HCM không phải sheet Báo cáo
    for s in xls.sheet_names:
        if 'HCM' in s and ('Hàng GE' in s or 'trước ngày' in s or '25' in s or 'GBC' in s):
            target_sheet = s
            break
    if not target_sheet:
        for s in xls.sheet_names:
            if 'HCM' in s and 'Báo cáo' not in s:
                target_sheet = s
                break
    if not target_sheet:
        target_sheet = xls.sheet_names[0]
        
    raw_df = pd.read_excel(xls, sheet_name=target_sheet)
    col0 = raw_df.columns[0]
    
    df = raw_df.rename(columns={
        col0: 'raw_timestamp',
        'ID': 'item_id',
        'Tên hàng hóa': 'item_name',
        'Ticket': 'ticket_id',
        'Dịch vụ': 'service',
        'Vị trí lưu trữ của COA ': 'storage_location',
        'Phân loại hàng hóa (theo cột C)': 'raw_category',
        'Note': 'raw_note'
    }).copy()
    
    # Chuẩn hóa ngày (Khớp Cell 7 Notebook)
    def clean_timestamp(val):
        if pd.isna(val) or val == '#REF!' or str(val).strip() == '':
            return pd.NaT
        if isinstance(val, (pd.Timestamp, datetime)):
            return pd.to_datetime(val)
        val_str = str(val).strip()
        try:
            return pd.to_datetime(val_str, dayfirst=True)
        except:
            return pd.NaT
            
    df['cleaned_date'] = df['raw_timestamp'].apply(clean_timestamp)
    
    # Khử trùng lặp (Khớp Cell 17 Notebook)
    df_clean = df.drop_duplicates().copy()
    tail_indices = [1639, 1640, 1641, 1642, 1643]
    df_clean = df_clean.drop(index=[i for i in tail_indices if i in df_clean.index])
    dup_biz_mask = df_clean.duplicated(subset=['item_id', 'ticket_id', 'item_name', 'storage_location'], keep='first')
    df_clean = df_clean[~dup_biz_mask].copy()
    
    # Cấp Synthetic ID (Khớp Cell 17 Notebook)
    def assign_op_id(row, idx):
        raw_id = str(row['item_id']).strip().lower()
        if pd.isna(row['item_id']) or raw_id in ['nan', 'không có', 'none', '']:
            return f"MANUAL-ITEM-{idx:04d}"
        try:
            return str(int(float(raw_id)))
        except:
            return str(row['item_id']).strip()
            
    df_clean['system_item_id'] = [assign_op_id(row, i) for i, (_, row) in enumerate(df_clean.iterrows(), start=1)]
    
    # Chuẩn hóa phân loại (Khớp Cell 17 Notebook)
    def std_cat(row):
        cat = str(row['raw_category']).strip()
        name = str(row['item_name']).strip().lower()
        if cat in ['nan', '', 'None']:
            if 'máy may' in name:
                return 'Thiết bị điện tử có giá'
            elif 'quà tết' in name or 'thực phẩm' in name:
                return 'Hàng hóa hàng ngày, giấy tờ khác (quần áo, mắt kính...)'
            return 'Hàng hóa hàng ngày, giấy tờ khác (quần áo, mắt kính...)'
        if 'Thiết bị điện tử' in cat:
            return 'Thiết bị điện tử có giá'
        elif 'Giấy tờ tùy thân' in cat or 'Tiền mặt' in cat:
            return 'Tiền mặt & giấy tờ cấp bởi CQCN'
        elif 'Thẻ Ngân hàng' in cat:
            return 'Thẻ Ngân hàng'
        elif 'Thuốc lá điện tử' in cat:
            return 'Thuốc lá điện tử'
        elif 'Hàng hóa hàng ngày' in cat:
            return 'Hàng hóa hàng ngày, giấy tờ khác (quần áo, mắt kính...)'
        return cat
        
    df_clean['standard_category'] = df_clean.apply(std_cat, axis=1)
    
    # Rule Engine chính sách (Khớp 100% Cell 22 Notebook)
    CUTOFF_DATE = pd.to_datetime('2023-09-25')
    def eval_policy(row):
        service = str(row['service']).strip().upper()
        cat = row['standard_category']
        date = row['cleaned_date']
        
        # Tính số ngày lưu kho đến 25/09/2023
        if pd.isna(date):
            days_in_storage = 999
        else:
            days_in_storage = (CUTOFF_DATE - date).days
            
        # Aging Status (Khớp Cell 22 Notebook)
        if 'CQCN' in cat or 'Thẻ Ngân hàng' in cat:
            status = 'Cần gửi CQNN'
        else:
            if 'GBC' in service:
                status = 'Quá hạn' if days_in_storage >= 30 else 'Còn hạn'
            else:
                status = 'Quá hạn' if days_in_storage >= 180 else 'Còn hạn'
                
        # Disposal Action (Khớp Cell 22 Notebook)
        if 'CQCN' in cat or 'Thẻ Ngân hàng' in cat:
            disposal_action = 'Bàn giao CQNN / Ngân hàng'
            legal_form = 'Phụ lục 4 (TTLT 18)'
            responsible_pic = 'COA + Legal + FPS'
        elif 'Thuốc lá điện tử' in cat:
            disposal_action = 'Tự tiêu hủy tại Hub (Chất thải nguy hại)'
            legal_form = 'Biên bản tiêu hủy nội bộ'
            pic = 'FPS + COA'
            responsible_pic = pic
        elif 'Thiết bị điện tử' in cat:
            disposal_action = 'Xóa dữ liệu & Thanh lý / Thuê Vendor hủy'
            legal_form = 'Phụ lục 3 (TTLT 18)'
            responsible_pic = 'GE + FPS + Procurement'
        else:
            disposal_action = 'Thuê Vendor tiêu hủy / Quyên góp'
            legal_form = 'Phụ lục 3 (TTLT 18)'
            responsible_pic = 'GE + FPS + Vendor'
            
        return pd.Series([days_in_storage, status, disposal_action, legal_form, responsible_pic],
                         index=['days_in_storage', 'aging_status', 'disposal_action', 'legal_form', 'responsible_pic'])
                         
    df_clean[['days_in_storage', 'aging_status', 'disposal_action', 'legal_form', 'responsible_pic']] = df_clean.apply(eval_policy, axis=1)
    return df_clean, len(raw_df) - len(df_clean)


# -----------------------------------------------------------------------------
# 3. SIDEBAR & CONTROL PANEL
# -----------------------------------------------------------------------------
with st.sidebar:
    st.markdown("### 🟢 GRAB CONTROL PANEL")
    st.caption("Grab Vietnam • Team COA (Custodian of Assets)")
    st.caption("Tổng Kho HCM (GBC & GE) • Tập kết từ các Hubs")
    
    # Nguồn dữ liệu
    data_source = st.radio(
        "Nguồn dữ liệu phân tích:",
        ["Dữ liệu mặc định (Đợt <25/09/2023)", "Tải lên file Excel mới (.xlsx)"],
        index=0
    )
    
    uploaded_file = None
    if data_source == "Tải lên file Excel mới (.xlsx)":
        uploaded_file = st.file_uploader("Chọn file Excel đợt mới:", type=['xlsx', 'xls'])
        
    st.markdown("---")
    st.markdown("#### 🔍 Bộ Lọc Đa Chiều (Multi-Filters)")

# Nạp dữ liệu
df_base = None
drop_count_note = 9
if uploaded_file is not None:
    with st.spinner("Đang chạy pipeline làm sạch tự động..."):
        df_base, drop_count_note = process_raw_excel(uploaded_file)
        st.sidebar.success(f"Đã xử lý file mới: {len(df_base)} kiện hàng!")
else:
    df_base = load_default_cleaned_data()

if df_base is None:
    st.error("Không tìm thấy dữ liệu mẫu! Vui lòng tải lên file Excel hoặc kiểm tra file cleaned_lost_and_found_hcm.csv.")
    st.stop()

# Bộ lọc tại Sidebar
services = ['Tất cả'] + sorted(list(df_base['service'].dropna().unique()))
selected_service = st.sidebar.selectbox("1. Dịch vụ Grab (GBC / GE):", services, index=0)

statuses = ['Tất cả'] + sorted(list(df_base['aging_status'].dropna().unique()))
selected_status = st.sidebar.selectbox("2. Trạng thái lưu kho (Aging):", statuses, index=0)

categories = ['Tất cả'] + sorted(list(df_base['standard_category'].dropna().unique()))
selected_category = st.sidebar.selectbox("3. Nhóm phân loại chuẩn:", categories, index=0)

actions = ['Tất cả'] + sorted(list(df_base['disposal_action'].dropna().unique()))
selected_action = st.sidebar.selectbox("4. Hướng xử lý nghiệp vụ:", actions, index=0)

# Bộ lọc Vị trí kho
locations = sorted([str(x) for x in df_base['storage_location'].dropna().unique()])
selected_locations = st.sidebar.multiselect("5. Vị trí kho / Bao / Thùng:", locations, placeholder="Chọn vị trí cụ thể...")

# Tìm kiếm từ khóa
search_kw = st.sidebar.text_input("6. Tìm kiếm từ khóa (Tên hàng, Ticket, ID):", placeholder="Ví dụ: iphone, cmnd, nồi chiên...")

st.sidebar.markdown("---")
st.sidebar.info("""
**Căn cứ pháp lý & Quy định Grab:**
- TTLT **18/2015/TTLT-BTTTT-BTC**
- SOP Lost & Found Grab (Review T5/2026)
- SLA: GBC (30 ngày) • GE (180 ngày)
""")

# Áp dụng bộ lọc
filtered_df = df_base.copy()

if selected_service != 'Tất cả':
    filtered_df = filtered_df[filtered_df['service'] == selected_service]

if selected_status != 'Tất cả':
    filtered_df = filtered_df[filtered_df['aging_status'] == selected_status]

if selected_category != 'Tất cả':
    filtered_df = filtered_df[filtered_df['standard_category'] == selected_category]

if selected_action != 'Tất cả':
    filtered_df = filtered_df[filtered_df['disposal_action'] == selected_action]

if selected_locations:
    filtered_df = filtered_df[filtered_df['storage_location'].astype(str).isin(selected_locations)]

if search_kw.strip():
    kw = search_kw.strip().lower()
    mask = (
        filtered_df['item_name'].astype(str).str.lower().str.contains(kw) |
        filtered_df['ticket_id'].astype(str).str.lower().str.contains(kw) |
        filtered_df['system_item_id'].astype(str).str.lower().str.contains(kw) |
        filtered_df['storage_location'].astype(str).str.lower().str.contains(kw)
    )
    filtered_df = filtered_df[mask]


# -----------------------------------------------------------------------------
# 4. HEADER & TOP KPI BANNER (GRAB THEME)
# -----------------------------------------------------------------------------
st.markdown("""
<div class="main-header">
    <div style="display: flex; justify-content: space-between; align-items: center;">
        <div>
            <span class="badge-pill badge-grab">GRAB VIETNAM</span>
            <span class="badge-pill badge-hub">TEAM COA • CUSTODIAN OF ASSETS</span>
            <span class="badge-pill badge-hub">TỔNG KHO HCM • GBC &amp; GE</span>
            <h1>🛵 GRAB LOST &amp; FOUND OPERATIONS CONTROL TOWER</h1>
            <p>Hệ thống giám sát điều phối dứt điểm 1,635 kiện hàng thất lạc GrabBike, GrabCar &amp; GrabExpress trên toàn TP.HCM</p>
        </div>
    </div>
</div>
""", unsafe_allow_html=True)

# KPI Cards Row
col1, col2, col3, col4, col5 = st.columns(5)

total_items = len(filtered_df)
overdue_items = int((filtered_df['aging_status'] == 'Quá hạn').sum())
cqnn_items = int((filtered_df['aging_status'] == 'Cần gửi CQNN').sum())
elec_items = int((filtered_df['standard_category'] == 'Thiết bị điện tử có giá').sum())
vape_items = int((filtered_df['standard_category'] == 'Thuốc lá điện tử').sum())

with col1:
    st.markdown(f"""
    <div class="metric-card">
        <div class="metric-title">📦 Tổng Kiện Hàng</div>
        <div class="metric-value" style="color: #00843D;">{total_items:,}</div>
        <div class="metric-subtitle">Kiện hàng vật lý thực tế</div>
    </div>
    """, unsafe_allow_html=True)

with col2:
    pct_overdue = (overdue_items / total_items * 100) if total_items > 0 else 0
    st.markdown(f"""
    <div class="metric-card metric-card-danger">
        <div class="metric-title">⏰ Quá Hạn Lưu Kho</div>
        <div class="metric-value" style="color: #DC2626;">{overdue_items:,}</div>
        <div class="metric-subtitle">Chiếm {pct_overdue:.1f}% tổng số kiện</div>
    </div>
    """, unsafe_allow_html=True)

with col3:
    st.markdown(f"""
    <div class="metric-card metric-card-warning">
        <div class="metric-title">🏛️ Cần Gửi CQNN / Thẻ</div>
        <div class="metric-value" style="color: #D97706;">{cqnn_items:,}</div>
        <div class="metric-subtitle">Biên bản Phụ lục 4 (TTLT 18)</div>
    </div>
    """, unsafe_allow_html=True)

with col4:
    st.markdown(f"""
    <div class="metric-card metric-card-blue">
        <div class="metric-title">💻 Thiết Bị Điện Tử</div>
        <div class="metric-value" style="color: #0284C7;">{elec_items:,}</div>
        <div class="metric-subtitle">Xóa sạch dữ liệu & thanh lý</div>
    </div>
    """, unsafe_allow_html=True)

with col5:
    st.markdown(f"""
    <div class="metric-card metric-card-purple">
        <div class="metric-title">⚠️ Thuốc Lá Điện Tử</div>
        <div class="metric-value" style="color: #7C3AED;">{vape_items:,}</div>
        <div class="metric-subtitle">Tự hủy chất thải nguy hại</div>
    </div>
    """, unsafe_allow_html=True)

st.markdown("<br>", unsafe_allow_html=True)


# -----------------------------------------------------------------------------
# 5. MAIN NAVIGATION TABS
# -----------------------------------------------------------------------------
tab1, tab2, tab3, tab4, tab5 = st.tabs([
    "📊 Phân Tích & Phân Luồng (Analytics)",
    "📅 Kế Hoạch Tác Chiến 7 Ngày (Blueprint)",
    "🔍 Tra Cứu Kiện Hàng & Xuất File (Explorer)",
    "📂 Tải Lên & Xử Lý Đợt Mới (ETL Pipeline)",
    "💡 Đề Xuất Cải Tiến & Template Chuẩn (Standards)"
])

# -----------------------------------------------------------------------------
# TAB 1: PHÂN TÍCH & MA TRẬN CHÍNH SÁCH
# -----------------------------------------------------------------------------
with tab1:
    st.subheader("1. Luồng Phân Bổ Hàng Hóa Theo Policy (SOP Grab & Thông Tư 18)")
    
    col_chart1, col_chart2 = st.columns([1, 1])
    
    with col_chart1:
        # Donut Chart Trạng Thái (Grab Palette)
        status_df = filtered_df['aging_status'].value_counts().reset_index()
        status_df.columns = ['Trạng thái', 'Số lượng']
        fig_donut = px.pie(
            status_df, 
            values='Số lượng', 
            names='Trạng thái',
            hole=0.55,
            color='Trạng thái',
            color_discrete_map={
                'Quá hạn': '#EF4444',
                'Còn hạn': '#00B14F',
                'Cần gửi CQNN': '#F59E0B'
            },
            title="Tỷ Trọng Trạng Thái Lưu Trữ (Aging Status)"
        )
        fig_donut.update_traces(textposition='inside', textinfo='percent+label')
        fig_donut.update_layout(showlegend=False, margin=dict(t=40, b=20, l=20, r=20))
        st.plotly_chart(fig_donut, use_container_width=True)
        
    with col_chart2:
        # Bar chart Hướng xử lý (Grab Palette)
        action_df = filtered_df['disposal_action'].value_counts().reset_index()
        action_df.columns = ['Hướng xử lý', 'Số lượng']
        fig_bar_action = px.bar(
            action_df,
            x='Số lượng',
            y='Hướng xử lý',
            orientation='h',
            color='Hướng xử lý',
            color_discrete_sequence=['#00B14F', '#0284C7', '#F59E0B', '#EF4444'],
            title="Số Lượng Kiện Hàng Theo Hướng Xử Lý Nghiệp Vụ"
        )
        fig_bar_action.update_layout(showlegend=False, margin=dict(t=40, b=20, l=20, r=20))
        st.plotly_chart(fig_bar_action, use_container_width=True)
        
    # Sunburst Hierarchy Flow
    st.markdown("#### 🌳 Luồng Phân Tầng Chi Tiết (Grab Service → Category → Action → Status)")
    if len(filtered_df) > 0:
        fig_sunburst = px.sunburst(
            filtered_df,
            path=['service', 'standard_category', 'disposal_action', 'aging_status'],
            title="Bản Đồ Phân Luồng Tài Sản Grab Lost & Found",
            color_discrete_sequence=['#00B14F', '#00843D', '#10B981', '#34D399', '#6EE7B7', '#0284C7', '#38BDF8']
        )
        fig_sunburst.update_layout(margin=dict(t=40, b=20, l=20, r=20))
        st.plotly_chart(fig_sunburst, use_container_width=True)
    
    # Top Vị trí lưu kho
    st.markdown("#### 📍 Top 12 Vị Trí Kho / Bao / Thùng Chiếm Nhiều Hàng Nhất")
    top_loc = filtered_df['storage_location'].value_counts().head(12).reset_index()
    top_loc.columns = ['Vị trí lưu kho', 'Số kiện']
    fig_loc = px.bar(
        top_loc,
        x='Vị trí lưu kho',
        y='Số kiện',
        text='Số kiện',
        color='Số kiện',
        color_continuous_scale=[[0, '#D1FAE5'], [0.5, '#34D399'], [1, '#00843D']],
        title="Phân Bổ Hàng Hóa Theo Vị Trí Thực Địa Tại Kho Lý Thường Kiệt"
    )
    fig_loc.update_traces(textposition='outside')
    fig_loc.update_layout(margin=dict(t=40, b=20, l=20, r=20))
    st.plotly_chart(fig_loc, use_container_width=True)
    
    # Ma trận tổng hợp Crosstab
    st.markdown("#### 📑 Ma Trận Chi Tiết (Crosstabulation Table)")
    pivot_table = pd.crosstab(
        index=[filtered_df['service'], filtered_df['standard_category']],
        columns=[filtered_df['aging_status'], filtered_df['disposal_action']],
        margins=True, margins_name='Tổng cộng'
    )
    st.dataframe(pivot_table, use_container_width=True)


# -----------------------------------------------------------------------------
# TAB 2: KẾ HOẠCH TÁC CHIẾN 7 NGÀY
# -----------------------------------------------------------------------------
with tab2:
    st.subheader("2. Kế Hoạch Tác Chiến Xử Lý Dứt Điểm Trong 7 Ngày Dương Lịch")
    
    # Capacity model banner
    st.markdown("""
    > **🎯 Định mức Năng lực (Capacity Planning Model):**
    > - **Tổng khối lượng**: 1,635 kiện hàng vật lý.
    > - **Thời gian đồng kiểm thực địa**: 5 ngày tập trung (Day 1 - Day 5). Bình quân: **~327 kiện/ngày**.
    > - **Năng lực định mức**: Mỗi cặp nhân sự (1 COA + 1 FPS) kiểm đếm **50 kiện/giờ** x ca 4 tiếng = **200 kiện/ca/ngày**.
    > - **Bố trí nhân sự**: **2 tổ đồng kiểm liên ngành** (Gồm 4 nhân sự: 2 COA + 2 FPS/GE). Đảm bảo công suất **400 kiện/ngày** (Dư buffer an toàn 22%).
    """)
    
    st.markdown("---")
    
    cqnn_c = int((df_base['aging_status'] == 'Cần gửi CQNN').sum())
    elec_c = int((df_base['standard_category'] == 'Thiết bị điện tử có giá').sum())
    vape_c = int((df_base['standard_category'] == 'Thuốc lá điện tử').sum())
    vendor_items = int((df_base['disposal_action'] == 'Thuê Vendor tiêu hủy / Quyên góp').sum())
    day4_c = int(np.ceil(vendor_items / 2))
    day5_c = vendor_items - day4_c
    
    days_data = [
        {
            "day": "Day 1 (Thứ 2)",
            "title": "Kick-off, Layout Kho & Tiêu Hủy Hàng Nguy Hại",
            "badge": "KHỞI ĐỘNG",
            "vol": f"{vape_c + 119} kiện",
            "pic": "Lead COA + Kho FPS",
            "details": f"Thiết lập khu vực đồng kiểm tại Kho LTK; Phân loại & tiêu hủy ngay Thuốc lá điện tử ({vape_c} items) + Thực phẩm/Rượu còn tồn đọng; Sort nhanh theo Thùng/Bao.",
            "deliverable": "Biên bản tiêu hủy hàng nguy hại nội bộ; Kho bãi được chia layout theo line xử lý."
        },
        {
            "day": "Day 2 (Thứ 3)",
            "title": "Đồng Kiểm Nhóm Pháp Lý & CQNN (P1 - Priority 1)",
            "badge": "PHÁP LÝ CAO",
            "vol": f"{cqnn_c} kiện",
            "pic": "COA + Legal + FPS",
            "details": f"Kiểm đếm 100% Giấy tờ tùy thân (CCCD, CMND, GPLX), Thẻ ATM, Tiền mặt ({cqnn_c} items); Đối soát thông tin chủ sở hữu nếu có trên Help Center.",
            "deliverable": "Niêm phong túi an toàn; Biên bản Phụ lục 4 (TTLT 18); Dự thảo Công văn gửi Công an / CQNN."
        },
        {
            "day": "Day 3 (Thứ 4)",
            "title": "Đồng Kiểm Thiết Bị Điện Tử & An Ninh Dữ Liệu (P2)",
            "badge": "GIÁ TRỊ CAO",
            "vol": f"{elec_c} kiện",
            "pic": "COA + GE + Kỹ thuật IT",
            "details": f"Kiểm tra tình trạng kỹ thuật Điện thoại, Laptop, Sạc dự phòng, Tai nghe ({elec_c} items); Xóa dữ liệu cá nhân, factory reset; Phân nhóm Thanh lý vs Tiêu hủy.",
            "deliverable": "Biên bản Phụ lục 3 (TTLT 18); Danh mục thiết bị đề xuất thanh lý thu hồi chi phí."
        },
        {
            "day": "Day 4 (Thứ 5)",
            "title": "Đồng Kiểm Hàng Tiêu Dùng Đợt 1 (Kho LTK - P3)",
            "badge": "GIẢI PHÓNG KHO",
            "vol": f"{day4_c} kiện",
            "pic": "Tổ 1 + Tổ 2 Đồng kiểm",
            "details": "Đồng kiểm các lô hàng tiêu dùng, quần áo, mũ nón tại Kho Lý Thường Kiệt & các Bao lớn (Bao 1 - Bao 15). Đóng gói vào bao tải quy chuẩn.",
            "deliverable": "Danh mục niêm phong đợt 1; Cân tổng trọng lượng phục vụ tính phí Vendor."
        },
        {
            "day": "Day 5 (Thứ 6)",
            "title": "Đồng Kiểm Hàng Tiêu Dùng Đợt 2 (Các Thùng Lẻ)",
            "badge": "HOÀN TẤT KIỂM KÊ",
            "vol": f"{day5_c} kiện",
            "pic": "Tổ 1 + Tổ 2 Đồng kiểm",
            "details": "Đồng kiểm toàn bộ các thùng lẻ còn lại (Thùng 01 - Thùng 30, Bì lẻ); Khử trùng dứt điểm danh sách tồn.",
            "deliverable": "Hoàn tất 100% kiểm đếm vật lý; Bảng đối soát khớp 1:1 với Data."
        },
        {
            "day": "Day 6 (Thứ 7)",
            "title": "Hội Đồng Phê Duyệt & Bàn Giao Xe Vendor Tiêu Hủy",
            "badge": "BÀN GIAO TIÊU HỦY",
            "vol": f"{len(df_base)} kiện",
            "pic": "Hội đồng xử lý + Procurement",
            "details": "Họp Hội đồng xử lý bưu gửi (GE + FPS + COA); Ký biên bản tổng Phụ lục 1; Xe Vendor đến bốc dỡ tiêu hủy dưới sự giám sát camera.",
            "deliverable": "Biên bản bàn giao rác thải công nghiệp có xác nhận của Vendor; Video/Ảnh nghiệm thu tiêu hủy."
        },
        {
            "day": "Day 7 (Chủ Nhật)",
            "title": "Cập Nhật Hệ Thống & Nghiệm Thu Dự Án",
            "badge": "NGHIỆM THU",
            "vol": "0 kiện (Đã sạch)",
            "pic": "Lead DA & Team COA",
            "details": "Cập nhật trạng thái 'ĐÃ XỬ LÝ' lên Dashboard/DB; Dọn dẹp trả lại mặt bằng kho; Lập báo cáo kết quả gửi Ban Giám đốc.",
            "deliverable": "Báo cáo hoàn thành dứt điểm đợt L&F; Dashboard nghiệm thu 100% Resolved."
        }
    ]
    
    for item in days_data:
        with st.expander(f"📌 {item['day']} — {item['title']} (Khối lượng: {item['vol']})", expanded=(item['day'].startswith("Day 1"))):
            col_d1, col_d2, col_d3 = st.columns([2, 1, 1])
            with col_d1:
                st.markdown(f"**Chi tiết công việc:** {item['details']}")
                st.markdown(f"**Đầu ra nghiệm thu:** `{item['deliverable']}`")
            with col_d2:
                st.markdown(f"**Nhân sự phụ trách:** `{item['pic']}`")
                st.markdown(f"**Phân loại ưu tiên:** `{item['badge']}`")
            with col_d3:
                is_done = st.checkbox("Đã hoàn thành", key=f"chk_{item['day']}")
                if is_done:
                    st.success("✅ Đã nghiệm thu")


# -----------------------------------------------------------------------------
# TAB 3: TRA CỨU KIỆN HÀNG & XUẤT FILE
# -----------------------------------------------------------------------------
with tab3:
    st.subheader(f"3. Tra Cứu Kiện Hàng Thực Tế ({len(filtered_df):,} kiện tìm thấy)")
    
    # Quick filter buttons
    st.markdown("**Bộ lọc nhanh các tài sản đặc biệt:**")
    q_col1, q_col2, q_col3, q_col4, q_col5 = st.columns(5)
    
    active_quick = None
    with q_col1:
        if st.button("📱 iPhone / IMEI"):
            active_quick = "iphone"
    with q_col2:
        if st.button("🪪 CCCD / CMND"):
            active_quick = "cmnd|cccd|gplx"
    with q_col3:
        if st.button("💵 Tiền Mặt"):
            active_quick = "tiền"
    with q_col4:
        if st.button("🍳 Nồi Chiên Không Dầu"):
            active_quick = "nồi chiên"
    with q_col5:
        if st.button("💊 Túi Thuốc Y Tế"):
            active_quick = "thuốc"
            
    display_df = filtered_df.copy()
    if active_quick:
        display_df = display_df[display_df['item_name'].astype(str).str.lower().str.contains(active_quick, regex=True, na=False)]
        st.info(f"Đang hiển thị {len(display_df)} kiện hàng liên quan đến từ khóa '{active_quick}'.")
        
    cols_to_show = [
        'system_item_id', 'ticket_id', 'item_name', 'service', 
        'standard_category', 'days_in_storage', 'aging_status', 
        'disposal_action', 'storage_location', 'responsible_pic'
    ]
    
    st.dataframe(
        display_df[cols_to_show].rename(columns={
            'system_item_id': 'Mã Vận Hành',
            'ticket_id': 'Ticket',
            'item_name': 'Tên Hàng Hóa',
            'service': 'Dịch Vụ',
            'standard_category': 'Phân Loại Chuẩn',
            'days_in_storage': 'Ngày Lưu Kho',
            'aging_status': 'Trạng Thái',
            'disposal_action': 'Hướng Xử Lý',
            'storage_location': 'Vị Trí Kho',
            'responsible_pic': 'Đơn Vị Phụ Trách'
        }),
        use_container_width=True,
        height=400
    )
    
    # Nút Tải Xuất Dữ Liệu
    col_dl1, col_dl2 = st.columns(2)
    with col_dl1:
        csv_buffer = display_df.to_csv(index=False, encoding='utf-8-sig').encode('utf-8-sig')
        st.download_button(
            label="📥 Tải Dữ Liệu Đang Xem (Định dạng CSV)",
            data=csv_buffer,
            file_name=f"lf_hcm_export_{datetime.now().strftime('%Y%m%d_%H%M%S')}.csv",
            mime="text/csv"
        )
    with col_dl2:
        excel_buffer = io.BytesIO()
        with pd.ExcelWriter(excel_buffer, engine='openpyxl') as writer:
            display_df.to_excel(writer, index=False, sheet_name='Cleaned_LF')
        st.download_button(
            label="📥 Tải Dữ Liệu Đang Xem (Định dạng Excel .xlsx)",
            data=excel_buffer.getvalue(),
            file_name=f"lf_hcm_export_{datetime.now().strftime('%Y%m%d_%H%M%S')}.xlsx",
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
        )


# -----------------------------------------------------------------------------
# TAB 4: TẢI LÊN & XỬ LÝ ĐỢT MỚI
# -----------------------------------------------------------------------------
with tab4:
    st.subheader("4. Pipeline Tự Động Xử Lý Dữ Liệu Cho Các Đợt L&F Tiếp Theo")
    st.markdown("""
    Hệ thống tích hợp sẵn thuật toán tự động nhận diện và làm sạch:
    - **Tự sửa lỗi Blank Header** của cột mốc thời gian.
    - **Tự động parse định dạng ngày** chuẩn Việt Nam `DD/MM/YYYY` (tránh bẫy chuẩn Mỹ) và gán mặc định Quá hạn cho dòng lỗi `#REF!`.
    - **Tự động phát hiện Missing ID / Dummy ID** và cấp mã định danh vận hành `MANUAL-ITEM-XXXX`.
    - **Tự động khử trùng lặp chuyển kho** (như Hub Gò Vấp dán đè ở cuối bảng).
    - **Áp dụng Rule Engine** phân loại chính sách theo SOP và TTLT 18.
    """)
    
    new_batch_file = st.file_uploader("Kéo thả file Excel đợt mới vào đây (.xlsx / .xls):", type=['xlsx', 'xls'], key="batch_uploader")
    
    if new_batch_file is not None:
        with st.spinner("Đang chạy toàn bộ quy trình ETL & Policy Rule Engine..."):
            new_cleaned_df, new_drop_count = process_raw_excel(new_batch_file)
            st.success(f"🎉 Hoàn tất làm sạch! Phát hiện & loại bỏ {new_drop_count} dòng trùng lặp, giữ lại chuẩn xác {len(new_cleaned_df)} kiện hàng vật lý.")
            
            c_res1, c_res2, c_res3 = st.columns(3)
            with c_res1:
                st.metric("Tổng Kiện Thực Tế", f"{len(new_cleaned_df):,} kiện")
            with c_res2:
                st.metric("Quá Hạn Lưu Kho", f"{(new_cleaned_df['aging_status'] == 'Quá hạn').sum():,} kiện")
            with c_res3:
                st.metric("Cần Bàn Giao CQNN", f"{(new_cleaned_df['aging_status'] == 'Cần gửi CQNN').sum():,} kiện")
                
            st.dataframe(new_cleaned_df.head(10), use_container_width=True)
            
            # Tải kết quả đợt mới
            res_csv = new_cleaned_df.to_csv(index=False, encoding='utf-8-sig').encode('utf-8-sig')
            st.download_button(
                "📥 Tải Về Dữ Liệu Sạch Đợt Mới (.CSV)",
                data=res_csv,
                file_name="new_batch_cleaned_lf.csv",
                mime="text/csv"
            )


# -----------------------------------------------------------------------------
# TAB 5: ĐỀ XUẤT CẢI TIẾN & TEMPLATE CHUẨN
# -----------------------------------------------------------------------------
with tab5:
    st.subheader("5. Chuẩn Hóa Nhập Liệu & Đề Xuất Tự Động Hóa Vận Hành")
    
    st.markdown("#### 📋 Đặc Tả Template Nhập Liệu Chuẩn Hóa (Standard Data Dictionary)")
    template_spec = pd.DataFrame([
        {
            "Tên Cột Chuẩn": "item_id",
            "Kiểu Dữ Liệu": "Varchar (Auto-generated)",
            "Ràng Buộc": "Bắt buộc duy nhất; format: LF-HCM-YYYYMM-XXXXX",
            "Lợi Ích Vận Hành": "Loại bỏ hoàn toàn tình trạng ID 'không có' hay dùng chung mã giữ chỗ 5253492."
        },
        {
            "Tên Cột Chuẩn": "intake_date",
            "Kiểu Dữ Liệu": "Date (YYYY-MM-DD)",
            "Ràng Buộc": "Bắt buộc; format ISO; Ngày tiếp nhận <= Ngày hiện tại",
            "Lợi Ích Vận Hành": "Triệt tiêu lỗi #REF! và lỗi đảo lộn ngày/tháng do bẫy US/VN format."
        },
        {
            "Tên Cột Chuẩn": "service_type",
            "Kiểu Dữ Liệu": "Dropdown (Enum)",
            "Ràng Buộc": "Chỉ chọn: ['Express', 'Bike', 'Car', 'Food', 'Mart']",
            "Lợi Ích Vận Hành": "Tự động áp dụng SLA lưu trữ tương ứng (30 ngày cho GBC, 180 ngày cho GE)."
        },
        {
            "Tên Cột Chuẩn": "ticket_id",
            "Kiểu Dữ Liệu": "String",
            "Ràng Buộc": "Bắt buộc với Express; Khuyến khích với Bike/Car",
            "Lợi Ích Vận Hành": "Liên kết trực tiếp tới cuốc xe và khách hàng trên Help Center."
        },
        {
            "Tên Cột Chuẩn": "category",
            "Kiểu Dữ Liệu": "Dropdown (Enum)",
            "Ràng Buộc": "5 Nhóm: [Điện tử, CQNN/Thẻ, Tiêu dùng, Nguy hại, Thực phẩm]",
            "Lợi Ích Vận Hành": "Không cho nhập tự do; map 1:1 sang hình thức xử lý pháp lý."
        },
        {
            "Tên Cột Chuẩn": "item_name",
            "Kiểu Dữ Liệu": "Text (Ngắn)",
            "Ràng Buộc": "Chỉ mô tả vật phẩm chính (ví dụ: 'Áo khoác gió')",
            "Lợi Ích Vận Hành": "Tách biệt hoàn toàn với chi tiết và chỉ thị xử lý."
        },
        {
            "Tên Cột Chuẩn": "item_details_and_notes",
            "Kiểu Dữ Liệu": "Text (Chi tiết)",
            "Ràng Buộc": "Ghi rõ IMEI, tiền mặt, thẻ ngân hàng, tình trạng mới/cũ",
            "Lợi Ích Vận Hành": "Giải quyết dứt điểm yêu cầu tách cột 'ghi chú xử lý' ra khỏi tên hàng."
        },
        {
            "Tên Cột Chuẩn": "warehouse_location",
            "Kiểu Dữ Liệu": "Barcode Mã Hóa",
            "Ràng Buộc": "Mã hóa: KHO-DAY-KE-THUNG (Ví dụ: LTK-A1-02-B05)",
            "Lợi Ích Vận Hành": "Quản lý vị trí chính xác đến từng kệ; quét mã vạch khi kiểm kê."
        }
    ])
    st.dataframe(template_spec, use_container_width=True)
    
    # Nút tải file Excel template mẫu
    empty_template = pd.DataFrame(columns=[
        'item_id', 'intake_date', 'service_type', 'ticket_id', 
        'category', 'item_name', 'item_details_and_notes', 'warehouse_location'
    ])
    tpl_buffer = io.BytesIO()
    with pd.ExcelWriter(tpl_buffer, engine='openpyxl') as writer:
        empty_template.to_excel(writer, index=False, sheet_name='Template_Nhap_Lieu')
        
    st.download_button(
        "📥 Tải File Mẫu Excel Nhập Liệu Chuẩn (.xlsx)",
        data=tpl_buffer.getvalue(),
        file_name="Template_Nhap_Lieu_Chuan_Lost_and_Found.xlsx",
        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
    )
    
    st.markdown("---")
    st.markdown("#### 🚀 3 Đề Xuất Tự Động Hóa Vận Hành Mang Tính Đột Phá")
    
    col_p1, col_p2, col_p3 = st.columns(3)
    with col_p1:
        st.markdown("""
        **1. Bot Cảnh Báo Sớm Lưu Kho (Aging Alert Bot)**
        - Tích hợp bot tự động gửi cảnh báo hàng tuần qua Slack/Email:
          - GBC: Cảnh báo tại mốc **T+20** (còn 10 ngày hết hạn).
          - GE: Cảnh báo tại **T+75** (hết hạn công khai) và **T+150** (chuẩn bị lập Hội đồng tiêu hủy).
        - **Hiệu quả**: Chuyển từ xử lý dồn ứ đa năm sang xử lý gối đầu hàng tháng (FIFO).
        """)
        
    with col_p2:
        st.markdown("""
        **2. Ứng Dụng Quét Mã Vạch Barcode Tại Kho**
        - Dán nhãn Barcode/QR Code cho từng kiện hàng ngay tại thời điểm tiếp nhận từ Tài xế/Hub.
        - Khi chuyển kho hoặc đồng kiểm: Thủ kho chỉ cần dùng máy quét cầm tay quét mã `BIP` thay vì gõ tay Excel.
        - **Hiệu quả**: Triệt tiêu 100% lỗi dán đè dòng chuyển kho như trường hợp Hub Gò Vấp.
        """)
        
    with col_p3:
        st.markdown("""
        **3. Bộ 3 Chỉ Số Giám Sát Sức Khỏe Tồn Kho (L&F Health KPI)**
        - `Aging Index`: Tỷ lệ hàng quá hạn lưu kho (Mục tiêu: < 5%).
        - `Disposal Cycle Time`: Thời gian từ lúc hết hạn đến khi tiêu hủy xong (Mục tiêu: < 7 ngày).
        - `Data Quality Rate`: Tỷ lệ bản ghi có đủ Ticket, Ngày và Phân loại (Mục tiêu: 100%).
        """)

# Footer
st.markdown("---")
st.caption("🛵 Grab Vietnam Operations Intelligence Platform • Team COA (Custodian of Assets) • GrabBike | GrabCar | GrabExpress • Built with Streamlit")
