"""
绿豆品牌销售业务与供应链全景看板
高级定制商业版 v3.1 (含双角色权限与数据持久化 + 图表高亮易读版)
"""

import streamlit as st
import pandas as pd
import plotly.express as px
import plotly.graph_objects as go
import os

# ──────────────────────────────────────────────
# 页面配置（必须在首行）
# ──────────────────────────────────────────────
st.set_page_config(
    page_title="绿豆品牌销售业务与供应链全景看板",
    page_icon="🌿",
    layout="wide",
    initial_sidebar_state="expanded",
)

# ──────────────────────────────────────────────
# 极致美化 CSS：苹果风高级哑光商务感
# ──────────────────────────────────────────────
custom_css = """
<style>
    @import url('https://fonts.googleapis.com/css2?family=Inter:wght@400;500;600;700&display=swap');
    html, body, [class*="css"] {
        font-family: 'Inter', 'Helvetica Neue', 'PingFang SC', 'Microsoft YaHei', sans-serif;
        background-color: #F4F7F9;
    }
    #MainMenu {visibility: hidden;}
    footer {visibility: hidden;}
    header[data-testid="stHeader"] {background: transparent;}
    button[kind="headerNoPadding"], section[data-testid="collapsedControl"] {
        display: block !important; visibility: visible !important; opacity: 1 !important;
    }
    [data-testid="stSidebar"] {
        background: linear-gradient(180deg, #1A2639 0%, #2A3B55 100%);
        box-shadow: 2px 0 12px rgba(0,0,0,0.1);
    }
    [data-testid="stSidebar"] * { color: #E2E8F0 !important; }
    [data-testid="stSidebar"] h1, [data-testid="stSidebar"] h2, [data-testid="stSidebar"] h3 {
        color: #FFFFFF !important; font-weight: 600 !important;
    }
    [data-testid="metric-container"] {
        background: #FFFFFF; border: 1px solid #EAEFEF; padding: 24px 20px !important;
        border-radius: 16px; box-shadow: 0 4px 16px rgba(18, 38, 63, 0.03);
        transition: all 0.3s ease; position: relative; overflow: hidden;
    }
    [data-testid="metric-container"]::before {
        content: ''; position: absolute; top: 0; left: 0; width: 4px; height: 100%;
        background: #3B82F6; border-radius: 4px 0 0 4px;
    }
    [data-testid="metric-container"]:hover {
        transform: translateY(-4px); box-shadow: 0 8px 24px rgba(18, 38, 63, 0.08);
    }
    [data-testid="metric-container"] label {
        color: #64748B !important; font-size: 0.9rem !important; font-weight: 500 !important;
    }
    [data-testid="metric-container"] [data-testid="stMetricValue"] {
        color: #0F172A !important; font-size: 2rem !important; font-weight: 700 !important; margin-top: 8px;
    }
    .stTabs [data-baseweb="tab-list"] {
        background: #FFFFFF; border-radius: 12px; padding: 6px;
        box-shadow: 0 2px 8px rgba(0,0,0,0.02); gap: 8px; margin-bottom: 24px;
    }
    .stTabs [data-baseweb="tab"] {
        border-radius: 8px !important; font-weight: 600 !important; color: #64748B !important;
        padding: 12px 24px !important; transition: all 0.2s ease; border: none; background: transparent;
    }
    .stTabs [aria-selected="true"] {
        background: #EBF5FF !important; color: #1D4ED8 !important;
    }
    [data-testid="stDataFrame"] {
        border-radius: 12px; overflow: hidden; border: 1px solid #E2E8F0; box-shadow: 0 2px 12px rgba(0,0,0,0.02);
    }
    .section-title {
        font-size: 1.35rem; font-weight: 700; color: #0F172A; margin: 28px 0 16px 0; display: flex; align-items: center;
    }
    .main-title {
        font-size: 2.2rem; font-weight: 800; color: #0F172A; letter-spacing: -0.03em; margin-bottom: 4px;
    }
    .main-subtitle { font-size: 1rem; color: #64748B; margin-bottom: 32px; font-weight: 400; }
</style>
"""
st.markdown(custom_css, unsafe_allow_html=True)

# ──────────────────────────────────────────────
# 访问控制与持久化文件路径
# ──────────────────────────────────────────────
CREDENTIALS = {"admin": "lvdou2026", "viewer": "view2026"}
DATA_FILE_PATH = "latest_business_data.xlsx"

if "authenticated" not in st.session_state:
    st.session_state.authenticated = False
if "username" not in st.session_state:
    st.session_state.username = ""

def show_login():
    st.markdown("""
    <div style="text-align:center; margin-top: 10vh; margin-bottom: 40px;">
        <div style="font-size:3.5rem; margin-bottom: 10px;">🌿</div>
        <div class="main-title">绿豆品牌业务与供应链看板</div>
        <p style="color:#64748B;">LVDOU BRAND BUSINESS INTELLIGENCE</p>
    </div>
    """, unsafe_allow_html=True)

    col1, col2, col3 = st.columns([1, 1.2, 1])
    with col2:
        with st.form("login_form"):
            username = st.text_input("登录账号", placeholder="请输入用户名")
            password = st.text_input("安全密码", type="password", placeholder="请输入密码")
            submitted = st.form_submit_button("进入看板", use_container_width=True)
            if submitted:
                if username in CREDENTIALS and CREDENTIALS[username] == password:
                    st.session_state.authenticated = True
                    st.session_state.username = username
                    st.rerun()
                else:
                    st.error("❌ 账号或密码错误，请核对。")

# ──────────────────────────────────────────────
# 核心数据引擎
# ──────────────────────────────────────────────
@st.cache_data(show_spinner=False)
def load_data_smartly(file_path_or_buffer, sheet_keywords, header_keywords):
    try:
        xls = pd.ExcelFile(file_path_or_buffer)
    except Exception:
        return pd.DataFrame()
        
    target_sheet = None
    for name in xls.sheet_names:
        if any(k in name for k in sheet_keywords):
            target_sheet = name
            break
            
    if not target_sheet: return pd.DataFrame() 
        
    df_raw = pd.read_excel(xls, sheet_name=target_sheet, header=None)
    header_idx = -1
    for idx, row in df_raw.iterrows():
        row_strs = [str(x) for x in row.values if pd.notna(x)]
        match_count = sum(1 for k in header_keywords if any(k in str(r) for r in row_strs))
        if match_count == len(header_keywords):
            header_idx = idx
            break
            
    if header_idx != -1:
        df = pd.read_excel(xls, sheet_name=target_sheet, header=header_idx)
        df = df.loc[:, ~df.columns.str.contains('^Unnamed')]
        
        valid_cols = [k for k in header_keywords if k in df.columns]
        if valid_cols:
            df = df.dropna(subset=valid_cols, how='all')
            
        mask = pd.Series(False, index=df.index)
        cols_to_check = df.columns[:10] 
        for col in cols_to_check:
            mask = mask | df[col].astype(str).str.contains('小计|合计|总计', na=False)
        df = df[~mask]
        return df
    return pd.DataFrame()

def style_dataframe(df, float_cols=None):
    if float_cols is None: float_cols = []
    df_clean = df.dropna(axis=1, how='all').fillna("")
    numeric_cols = df_clean.select_dtypes(include=['float64', 'int64']).columns
    format_dict = {}
    for col in numeric_cols:
        if any(f_c in col for f_c in float_cols): 
            format_dict[col] = "{:,.2f}"
        else:
            format_dict[col] = "{:,.0f}"
    return df_clean.style.format(format_dict, na_rep="")

# ──────────────────────────────────────────────
# 主看板构建
# ──────────────────────────────────────────────
def show_dashboard():
    with st.sidebar:
        st.markdown("### ⚙️ 数据控制台")
        role_name = "👑 超级管理员" if st.session_state.username == "admin" else "👤 访问访客"
        st.markdown(f"<p style='color:#94A3B8; font-size:0.85rem;'>当前角色：{role_name}</p>", unsafe_allow_html=True)
        st.markdown("<hr style='border-top:1px solid #334155; margin:15px 0;'>", unsafe_allow_html=True)
        
        if st.session_state.username == "admin":
            st.markdown("<p style='color:#E2E8F0; font-size:0.9rem;'>📁 <b>后台更新数据源</b></p>", unsafe_allow_html=True)
            uploaded_file = st.file_uploader("选择最新的 Excel 业务总表", type=["xlsx", "xls"])
            if uploaded_file:
                with open(DATA_FILE_PATH, "wb") as f:
                    f.write(uploaded_file.getvalue())
                st.success("✅ 数据已成功同步至云端！领导现在登录可直接查看最新数据。")
                st.cache_data.clear()
        else:
            st.info("💡 提示：您正处于只读模式。看板数据由管理员在后台统一更新。")
        
        st.markdown("<div style='margin-top:50vh;'></div>", unsafe_allow_html=True)
        if st.button("退出登录", use_container_width=True):
            st.session_state.authenticated = False
            st.session_state.username = ""
            st.rerun()

    st.markdown("""
        <div class="main-title">绿豆品牌销售业务与供应链全景看板</div>
        <div class="main-subtitle">LVDOU BRAND PANORAMIC INTELLIGENCE DASHBOARD</div>
    """, unsafe_allow_html=True)

    if os.path.exists(DATA_FILE_PATH):
        data_source = DATA_FILE_PATH
    else:
        data_source = None

    if not data_source:
        if st.session_state.username == "admin":
            st.warning("👈 系统当前为空。请在左侧栏上传您的《业务管理表 Excel》以初始化全景看板。")
        else:
            st.info("☕ 看板数据正在初始化维护中，请联系管理员上传最新数据。")
        st.stop()

    with st.spinner('⏳ 正在进行数据清洗与高维建模...'):
        df_sales = load_data_smartly(data_source, ['销售统计', '销售'], ['品牌', '区分'])
        df_inventory = load_data_smartly(data_source, ['成品库存'], ['产品名称', '实际库存量'])
        df_prepay = load_data_smartly(data_source, ['未执行订单', '预付款管理'], ['订单编号', '订单金额'])
        df_production = load_data_smartly(data_source, ['生产计划'], ['产品名称', '数量'])
        df_unproduced = load_data_smartly(data_source, ['备案未产'], ['产品编号', '产品名称'])
        df_ongoing = load_data_smartly(data_source, ['进行中业务'], ['项目名称', '业务阶段'])

    tab1, tab2, tab3, tab4 = st.tabs([
        "💰 全年销售业绩", 
        "📦 在库成品全景", 
        "💳 未执行订单与生产", 
        "🚀 核心业务推进"
    ])

    # ═══════════════════════════════════════════
    # Tab 1: 销售业绩
    # ═══════════════════════════════════════════
    with tab1:
        if not df_sales.empty:
            target_row = df_sales[df_sales['区分'].astype(str).str.contains('目标', na=False)]
            actual_row = df_sales[df_sales['区分'].astype(str).str.contains('实绩', na=False) & ~df_sales['区分'].astype(str).str.contains('达成率', na=False)]
            
            if not target_row.empty and not actual_row.empty:
                total_target = pd.to_numeric(str(target_row['2026全年'].iloc[0]).replace(',', ''), errors='coerce') if '2026全年' in target_row.columns else 0
                months_cols = [f"{i}月" for i in range(1, 13)]
                valid_months = [m for m in months_cols if m in target_row.columns]
                
                target_series = pd.to_numeric(target_row[valid_months].iloc[0].astype(str).str.replace(',', '', regex=False), errors='coerce').fillna(0)
                actual_series = pd.to_numeric(actual_row[valid_months].iloc[0].astype(str).str.replace(',', '', regex=False), errors='coerce').fillna(0)
                
                actual_data = actual_series.sum()
                achieve_rate = (actual_data / total_target * 100) if total_target > 0 else 0
                
                col1, col2, col3 = st.columns(3)
                col1.metric("📌 2026 全年总销售目标", f"¥ {total_target:,.0f}")
                col2.metric("✅ 当前累计完成实绩", f"¥ {actual_data:,.0f}")
                col3.metric("🔥 年度整体达成率", f"{achieve_rate:.1f}%")
                
                st.markdown('<div class="section-title">📊 2026 年月度目标与实绩走势</div>', unsafe_allow_html=True)
                df_plot = pd.DataFrame({'月份': valid_months, '目标销量': target_series, '实际销量': actual_series})
                fig_sales = px.bar(df_plot, x='月份', y=['目标销量', '实际销量'], barmode='group', 
                                   color_discrete_sequence=['#94A3B8', '#3B82F6'], text_auto=',.0f')
                fig_sales.update_layout(
                    plot_bgcolor='rgba(0,0,0,0)', paper_bgcolor='rgba(0,0,0,0)', 
                    legend_title="", yaxis_title="金额 (元)", margin=dict(t=20, b=10, l=10, r=10),
                    font=dict(color='#475569')
                )
                fig_sales.update_traces(textposition='outside')
                st.plotly_chart(fig_sales, use_container_width=True)

            df_history = df_sales[df_sales['区分'].astype(str).str.contains(r'2022|2023|2024|2025', na=False)].copy()
            if not df_history.empty:
                st.markdown("<hr>", unsafe_allow_html=True)
                st.markdown('<div class="section-title">📅 历年月别销售实绩回顾 (2023-2025)</div>', unsafe_allow_html=True)
                
                cols_to_keep = ['区分']
                total_col = [c for c in df_sales.columns if '全年' in str(c)]
                if total_col: cols_to_keep.append(total_col[0])
                cols_to_keep.extend([f"{i}月" for i in range(1, 13) if f"{i}月" in df_sales.columns])
                
                df_history_show = df_history[cols_to_keep].copy()
                df_history_show['年份_int'] = df_history_show['区分'].str.extract(r'(\d{4})').astype(float)
                df_history_show = df_history_show.sort_values(by='年份_int', ascending=False).drop(columns=['年份_int'])
                
                rename_dict = {'区分': '年份'}
                if total_col: rename_dict[total_col[0]] = '全年累计实绩额'
                df_history_show.rename(columns=rename_dict, inplace=True)
                
                for col in df_history_show.columns:
                    if col != '年份':
                        df_history_show[col] = pd.to_numeric(df_history_show[col].astype(str).str.replace(',', '', regex=False), errors='coerce').fillna(0)
                        
                st.dataframe(style_dataframe(df_history_show), use_container_width=True, hide_index=True)
        else:
            st.warning("⚠️ 无法抓取《销售统计》数据，请确认表头是否包含「品牌」和「区分」。")

    # ═══════════════════════════════════════════
    # Tab 2: 成品库存
    # ═══════════════════════════════════════════
    with tab2:
        if not df_inventory.empty:
            df_inv_clean = df_inventory.copy()
            for col in ['在库金额', '实际库存量', '出库单价']:
                if col in df_inv_clean.columns:
                    df_inv_clean[col] = df_inv_clean[col].astype(str).str.replace(',', '', regex=False)
                    df_inv_clean[col] = pd.to_numeric(df_inv_clean[col], errors='coerce').fillna(0)
            
            date_cols = [c for c in df_inv_clean.columns if '日期' in c or '时间' in c]
            for dc in date_cols:
                df_inv_clean[dc] = pd.to_datetime(df_inv_clean[dc], errors='coerce').dt.strftime('%Y-%m-%d')
                df_inv_clean[dc] = df_inv_clean[dc].fillna("") 
            
            total_inv_value = df_inv_clean['在库金额'].sum()
            total_inv_qty = df_inv_clean['实际库存量'].sum()
            
            col1, col2 = st.columns(2)
            col1.metric("💰 库存资金总额 (元)", f"¥ {total_inv_value:,.0f}")
            col2.metric("📦 物资总库存数 (件)", f"{total_inv_qty:,.0f} 件")
            
            # 【深度升级优化】库存资金预警柱状图
            st.markdown('<div class="section-title">⚠️ 库存资金预警 Top 10</div>', unsafe_allow_html=True)
            df_top_val = df_inv_clean[df_inv_clean['在库金额'] > 0].nlargest(10, '在库金额').copy()
            if not df_top_val.empty:
                # 算法：每隔12个字符换一行，保证Y轴文字清晰呈现
                def wrap_text(text, width=12):
                    text = str(text)
                    return '<br>'.join([text[i:i+width] for i in range(0, len(text), width)])
                df_top_val['换行名称'] = df_top_val['产品名称'].apply(wrap_text)
                
                # 自动分配莫兰迪等高级多色系区分每一根柱子，并悬挂具体数值
                fig_val = px.bar(
                    df_top_val, 
                    x='在库金额', 
                    y='换行名称', 
                    orientation='h', 
                    color='产品名称', # 触发多色系区分
                    text='在库金额',
                    color_discrete_sequence=px.colors.qualitative.Pastel # 柔和且高级的多色系
                )
                
                # 强制锁死文字大小，并将数值放在柱子外部最右侧，解决随柱子缩小的痛点
                fig_val.update_traces(
                    texttemplate='¥ %{text:,.0f}', 
                    textposition='outside', 
                    textfont=dict(color='#0F172A', size=13, family="sans-serif")
                )
                
                # 重新打开左侧 Y 轴显示名称，隐藏顶部没用的 X 轴数值
                fig_val.update_layout(
                    showlegend=False, 
                    plot_bgcolor='rgba(0,0,0,0)', paper_bgcolor='rgba(0,0,0,0)', 
                    yaxis={'categoryorder':'total ascending', 'showticklabels': True, 'title': '', 'tickfont': dict(size=12)}, 
                    xaxis={'showticklabels': False, 'title': '', 'showgrid': False}, 
                    margin=dict(t=10, b=10, l=10, r=80), height=550, # 高度增加，防止文字拥挤
                    uniformtext_minsize=12, uniformtext_mode='show' # 核心锁死大小属性
                )
                st.plotly_chart(fig_val, use_container_width=True)
            else:
                st.info("暂无库存资金预警数据。")
            
            st.markdown('<div class="section-title">📦 核心库存金额占比 Top 10</div>', unsafe_allow_html=True)
            df_top_val_pie = df_inv_clean[df_inv_clean['在库金额'] > 0].nlargest(10, '在库金额')
            if not df_top_val_pie.empty:
                fig_qty = px.pie(df_top_val_pie, names='产品名称', values='在库金额', hole=0.5, color_discrete_sequence=px.colors.qualitative.Pastel)
                fig_qty.update_layout(paper_bgcolor='rgba(0,0,0,0)', margin=dict(t=20, b=20, l=20, r=20), height=600, showlegend=False)
                fig_qty.update_traces(textinfo='label+percent', textposition='outside', textfont=dict(size=13))
                st.plotly_chart(fig_qty, use_container_width=True)
            else:
                st.info("暂无库存金额数据用于占比。")
            
            st.markdown("<hr>", unsafe_allow_html=True)
            st.markdown('<div class="section-title">📋 详细成品库存台账明细</div>', unsafe_allow_html=True)
            st.dataframe(style_dataframe(df_inv_clean, float_cols=['单价']), use_container_width=True, hide_index=True)
        else:
            st.warning("⚠️ 无法抓取《成品库存》数据。")

    # ═══════════════════════════════════════════
    # Tab 3: 订单与生产
    # ═══════════════════════════════════════════
    with tab3:
        st.markdown('<div class="section-title">💳 客户未执行订单明细</div>', unsafe_allow_html=True)
        if not df_prepay.empty:
            for c in ['订单金额', '已支付金额', '未支付金额', '已支付预付款', '未支付预付款']:
                if c in df_prepay.columns: 
                    df_prepay[c] = df_prepay[c].astype(str).str.replace(',', '', regex=False)
                    df_prepay[c] = pd.to_numeric(df_prepay[c], errors='coerce').fillna(0)
            total_order_amt = df_prepay['订单金额'].sum() if '订单金额' in df_prepay.columns else 0
            st.info(f"**💡 当前有效未执行订单总金额： ¥ {total_order_amt:,.0f}**")
            st.dataframe(style_dataframe(df_prepay), use_container_width=True, hide_index=True)
        else:
            st.warning("暂无未执行订单数据。")

        st.markdown("<hr>", unsafe_allow_html=True)
        st.markdown('<div class="section-title">⚙️ 当期工厂生产排期计划</div>', unsafe_allow_html=True)
        if not df_production.empty:
            for c in ['订单金额', '数量', '采购价', '单价']:
                if c in df_production.columns: 
                    df_production[c] = df_production[c].astype(str).str.replace(',', '', regex=False)
                    df_production[c] = pd.to_numeric(df_production[c], errors='coerce').fillna(0)
            total_prod_amt = df_production['订单金额'].sum() if '订单金额' in df_production.columns else 0
            total_prod_qty = df_production['数量'].sum() if '数量' in df_production.columns else 0
            st.success(f"**🛠️ 本期计划总产量： {total_prod_qty:,.0f} 件 | 投入资金预估： ¥ {total_prod_amt:,.0f}**")
            st.dataframe(style_dataframe(df_production, float_cols=['单价', '价']), use_container_width=True, hide_index=True)
        else:
            st.warning("暂无生产计划数据。")

    # ═══════════════════════════════════════════
    # Tab 4: 业务跟踪
    # ═══════════════════════════════════════════
    with tab4:
        col5, col6 = st.columns(2)
        with col5:
            st.markdown('<div class="section-title">⏳ 备案未产清单监控</div>', unsafe_allow_html=True)
            if not df_unproduced.empty:
                st.dataframe(style_dataframe(df_unproduced), use_container_width=True, height=600, hide_index=True)
            else:
                st.info("暂无备案未产数据")
                
        with col6:
            st.markdown('<div class="section-title">🏃 核心业务节点跟进</div>', unsafe_allow_html=True)
            if not df_ongoing.empty:
                st.dataframe(style_dataframe(df_ongoing), use_container_width=True, height=600, hide_index=True)
            else:
                st.info("暂无进行中业务数据")

if __name__ == "__main__":
    if not st.session_state.authenticated:
        show_login()
    else:
        show_dashboard()