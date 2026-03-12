"""
绿豆品牌销售业务与供应链全景看板
高级定制商业版
"""

import streamlit as st
import pandas as pd
import plotly.express as px
import plotly.graph_objects as go

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
    /* 全局字体优化 */
    @import url('https://fonts.googleapis.com/css2?family=Inter:wght@400;500;600;700&display=swap');
    html, body, [class*="css"] {
        font-family: 'Inter', 'Helvetica Neue', 'PingFang SC', 'Microsoft YaHei', sans-serif;
        background-color: #F4F7F9;
    }
    
    /* 隐藏 Streamlit 默认 header/footer，保留侧边栏展开按钮 */
    #MainMenu {visibility: hidden;}
    footer {visibility: hidden;}
    header[data-testid="stHeader"] {background: transparent;}
    button[kind="headerNoPadding"], section[data-testid="collapsedControl"] {
        display: block !important; visibility: visible !important; opacity: 1 !important;
    }

    /* 侧边栏高级渐变 */
    [data-testid="stSidebar"] {
        background: linear-gradient(180deg, #1A2639 0%, #2A3B55 100%);
        box-shadow: 2px 0 12px rgba(0,0,0,0.1);
    }
    [data-testid="stSidebar"] * { color: #E2E8F0 !important; }
    [data-testid="stSidebar"] h1, [data-testid="stSidebar"] h2, [data-testid="stSidebar"] h3 {
        color: #FFFFFF !important; font-weight: 600 !important;
    }

    /* KPI 核心指标卡片：苹果风微发光圆角设计 */
    [data-testid="metric-container"] {
        background: #FFFFFF;
        border: 1px solid #EAEFEF;
        padding: 24px 20px !important;
        border-radius: 16px;
        box-shadow: 0 4px 16px rgba(18, 38, 63, 0.03);
        transition: all 0.3s ease;
        position: relative;
        overflow: hidden;
    }
    [data-testid="metric-container"]::before {
        content: ''; position: absolute; top: 0; left: 0; width: 4px; height: 100%;
        background: #3B82F6; border-radius: 4px 0 0 4px;
    }
    [data-testid="metric-container"]:hover {
        transform: translateY(-4px);
        box-shadow: 0 8px 24px rgba(18, 38, 63, 0.08);
    }
    [data-testid="metric-container"] label {
        color: #64748B !important; font-size: 0.9rem !important; font-weight: 500 !important;
    }
    [data-testid="metric-container"] [data-testid="stMetricValue"] {
        color: #0F172A !important; font-size: 2rem !important; font-weight: 700 !important;
        margin-top: 8px;
    }

    /* 沉浸式标签页 (Tabs) 样式 */
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

    /* 数据表格无边框高级化 */
    [data-testid="stDataFrame"] {
        border-radius: 12px; overflow: hidden; border: 1px solid #E2E8F0;
        box-shadow: 0 2px 12px rgba(0,0,0,0.02);
    }
    
    /* 模块标题加重 */
    .section-title {
        font-size: 1.35rem; font-weight: 700; color: #0F172A;
        margin: 28px 0 16px 0; display: flex; align-items: center;
    }
    
    /* 顶部大标题 */
    .main-title {
        font-size: 2.2rem; font-weight: 800; color: #0F172A; letter-spacing: -0.03em; margin-bottom: 4px;
    }
    .main-subtitle {
        font-size: 1rem; color: #64748B; margin-bottom: 32px; font-weight: 400;
    }
</style>
"""
st.markdown(custom_css, unsafe_allow_html=True)

# ──────────────────────────────────────────────
# 访问控制 (可在此修改密码)
# ──────────────────────────────────────────────
CREDENTIALS = {"admin": "lvdou2026", "viewer": "view2026"}

if "authenticated" not in st.session_state:
    st.session_state.authenticated = False

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
# 核心数据引擎：智能雷达解析
# ──────────────────────────────────────────────
@st.cache_data(show_spinner=False)
def load_data_smartly(uploaded_file, sheet_keywords, header_keywords):
    """不受空行限制，智能寻找包含指定关键字的表头并精准读取"""
    xls = pd.ExcelFile(uploaded_file)
    target_sheet = None
    for name in xls.sheet_names:
        if any(k in name for k in sheet_keywords):
            target_sheet = name
            break
            
    if not target_sheet:
        return pd.DataFrame() 
        
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
        # 清洗：干掉无用列和带有“合计/小计”的干扰行
        df = df.loc[:, ~df.columns.str.contains('^Unnamed')]
        
        # 针对第一列或关键列进行合计清理
        clean_col = df.columns[0]
        for kw in header_keywords:
            if kw in df.columns:
                clean_col = kw
                break
        df = df.dropna(subset=[clean_col])
        df = df[~df[clean_col].astype(str).str.contains('小计|合计|总计', na=False)]
        return df
    return pd.DataFrame()

def style_dataframe(df):
    """全自动给财务数字加上千分位逗号，去掉小数"""
    df_clean = df.dropna(axis=1, how='all').fillna("")
    numeric_cols = df_clean.select_dtypes(include=['float64', 'int64']).columns
    format_dict = {col: "{:,.0f}" for col in numeric_cols}
    return df_clean.style.format(format_dict, na_rep="")

# ──────────────────────────────────────────────
# 主看板构建
# ──────────────────────────────────────────────
def show_dashboard():
    with st.sidebar:
        st.markdown("### ⚙️ 数据控制台")
        st.markdown(f"<p style='color:#94A3B8; font-size:0.85rem;'>欢迎回来，{st.session_state.username}</p>", unsafe_allow_html=True)
        st.markdown("<hr style='border-top:1px solid #334155; margin:15px 0;'>", unsafe_allow_html=True)
        uploaded_file = st.file_uploader("📁 上传最新业务总表 (Excel)", type=["xlsx", "xls"])
        
        st.markdown("<div style='margin-top:50vh;'></div>", unsafe_allow_html=True)
        if st.button("退出系统", use_container_width=True):
            st.session_state.authenticated = False
            st.rerun()

    st.markdown("""
        <div class="main-title">绿豆品牌销售业务与供应链全景看板</div>
        <div class="main-subtitle">LVDOU BRAND PANORAMIC INTELLIGENCE DASHBOARD</div>
    """, unsafe_allow_html=True)

    if not uploaded_file:
        st.info("👈 系统已就绪。请在左侧栏上传您的《业务管理表 Excel》以自动渲染全景数据。")
        st.stop()

    with st.spinner('⏳ 正在进行数据清洗与高维建模...'):
        # 智能提取各类表格
        df_sales = load_data_smartly(uploaded_file, ['销售统计', '销售'], ['品牌', '区分'])
        df_inventory = load_data_smartly(uploaded_file, ['成品库存'], ['产品名称', '实际库存量'])
        df_prepay = load_data_smartly(uploaded_file, ['未执行订单', '预付款管理'], ['订单编号', '订单金额'])
        df_production = load_data_smartly(uploaded_file, ['生产计划'], ['产品名称', '数量'])
        df_unproduced = load_data_smartly(uploaded_file, ['备案未产'], ['产品编号', '产品名称'])
        df_ongoing = load_data_smartly(uploaded_file, ['进行中业务'], ['项目名称', '业务阶段'])

    # 核心四大模块
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
                # KPI 计算
                total_target = pd.to_numeric(target_row['2026全年'].iloc[0], errors='coerce') if '2026全年' in target_row.columns else 0
                months_cols = [f"{i}月" for i in range(1, 13)]
                valid_months = [m for m in months_cols if m in target_row.columns]
                
                target_series = pd.to_numeric(target_row[valid_months].iloc[0], errors='coerce').fillna(0)
                actual_series = pd.to_numeric(actual_row[valid_months].iloc[0], errors='coerce').fillna(0)
                
                actual_data = actual_series.sum()
                achieve_rate = (actual_data / total_target * 100) if total_target > 0 else 0
                
                col1, col2, col3 = st.columns(3)
                col1.metric("📌 2026 全年总销售目标", f"¥ {total_target:,.0f}")
                col2.metric("✅ 当前累计完成实绩", f"¥ {actual_data:,.0f}")
                col3.metric("🔥 年度整体达成率", f"{achieve_rate:.1f}%")
                
                # 月度走势图
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

            # 历年历史数据提取
            df_history = df_sales[df_sales['区分'].astype(str).str.contains(r'2022|2023|2024|2025', na=False)].copy()
            if not df_history.empty:
                st.markdown("<hr>", unsafe_allow_html=True)
                st.markdown('<div class="section-title">📅 历年月别销售实绩回顾 (2023-2025)</div>', unsafe_allow_html=True)
                
                cols_to_keep = ['区分']
                total_col = [c for c in df_sales.columns if '全年' in str(c)]
                if total_col: cols_to_keep.append(total_col[0])
                cols_to_keep.extend([f"{i}月" for i in range(1, 13) if f"{i}月" in df_sales.columns])
                
                df_history_show = df_history[cols_to_keep].copy()
                rename_dict = {'区分': '年份'}
                if total_col: rename_dict[total_col[0]] = '全年累计实绩额'
                df_history_show.rename(columns=rename_dict, inplace=True)
                
                for col in df_history_show.columns:
                    if col != '年份':
                        df_history_show[col] = pd.to_numeric(df_history_show[col], errors='coerce').fillna(0)
                        
                st.dataframe(style_dataframe(df_history_show), use_container_width=True, hide_index=True)
        else:
            st.warning("⚠️ 无法抓取《销售统计》数据，请确认表头是否包含「品牌」和「区分」。")

    # ═══════════════════════════════════════════
    # Tab 2: 成品库存
    # ═══════════════════════════════════════════
    with tab2:
        if not df_inventory.empty:
            df_inv_clean = df_inventory.copy()
            df_inv_clean['在库金额'] = pd.to_numeric(df_inv_clean.get('在库金额', 0), errors='coerce').fillna(0)
            df_inv_clean['实际库存量'] = pd.to_numeric(df_inv_clean.get('实际库存量', 0), errors='coerce').fillna(0)
            
            total_inv_value = df_inv_clean['在库金额'].sum()
            total_inv_qty = df_inv_clean['实际库存量'].sum()
            
            col1, col2 = st.columns(2)
            col1.metric("💰 冻结资金总额 (元)", f"¥ {total_inv_value:,.0f}")
            col2.metric("📦 物资总库存数 (件)", f"{total_inv_qty:,.0f} 件")
            
            # 图表独占全宽，智能截断名称
            st.markdown('<div class="section-title">⚠️ 资金占用预警 Top 10</div>', unsafe_allow_html=True)
            df_top_val = df_inv_clean.nlargest(10, '在库金额').copy()
            df_top_val['产品简称'] = df_top_val['产品名称'].apply(lambda x: str(x)[:15] + '...' if len(str(x)) > 15 else str(x))
            
            fig_val = px.bar(df_top_val, x='在库金额', y='产品简称', orientation='h', 
                             color='在库金额', color_continuous_scale='Blues', text_auto=',.0f',
                             hover_data={'产品名称': True, '产品简称': False})
            fig_val.update_layout(
                plot_bgcolor='rgba(0,0,0,0)', paper_bgcolor='rgba(0,0,0,0)', 
                yaxis={'categoryorder':'total ascending'}, margin=dict(t=10, b=0, l=10, r=20), height=380
            )
            st.plotly_chart(fig_val, use_container_width=True)
            
            st.markdown('<div class="section-title">📦 核心库存件数占比 Top 10</div>', unsafe_allow_html=True)
            df_top_qty = df_inv_clean.nlargest(10, '实际库存量')
            fig_qty = px.pie(df_top_qty, names='产品名称', values='实际库存量', hole=0.5)
            fig_qty.update_layout(
                paper_bgcolor='rgba(0,0,0,0)', margin=dict(t=10, b=40, l=10, r=10), height=550,
                legend=dict(orientation="h", yanchor="top", y=-0.1, xanchor="center", x=0.5)
            )
            fig_qty.update_traces(textinfo='percent+value', textposition='inside')
            st.plotly_chart(fig_qty, use_container_width=True)
            
            st.markdown("<hr>", unsafe_allow_html=True)
            st.markdown('<div class="section-title">📋 详细成品库存台账明细</div>', unsafe_allow_html=True)
            st.dataframe(style_dataframe(df_inv_clean), use_container_width=True, hide_index=True)
        else:
            st.warning("⚠️ 无法抓取《成品库存》数据，请确认表头格式。")

    # ═══════════════════════════════════════════
    # Tab 3: 订单与生产
    # ═══════════════════════════════════════════
    with tab3:
        st.markdown('<div class="section-title">💳 客户未执行订单明细</div>', unsafe_allow_html=True)
        if not df_prepay.empty:
            for c in ['订单金额', '已支付金额', '未支付金额']:
                if c in df_prepay.columns: df_prepay[c] = pd.to_numeric(df_prepay[c], errors='coerce').fillna(0)
                    
            total_order_amt = df_prepay['订单金额'].sum() if '订单金额' in df_prepay.columns else 0
            st.info(f"**💡 当前有效未执行订单总金额： ¥ {total_order_amt:,.0f}**")
            st.dataframe(style_dataframe(df_prepay), use_container_width=True, hide_index=True)
        else:
            st.warning("暂无未执行订单数据。")

        st.markdown("<hr>", unsafe_allow_html=True)
        st.markdown('<div class="section-title">⚙️ 当期工厂生产排期计划</div>', unsafe_allow_html=True)
        if not df_production.empty:
            for c in ['订单金额', '数量', '采购价']:
                if c in df_production.columns: df_production[c] = pd.to_numeric(df_production[c], errors='coerce').fillna(0)
                    
            total_prod_amt = df_production['订单金额'].sum() if '订单金额' in df_production.columns else 0
            total_prod_qty = df_production['数量'].sum() if '数量' in df_production.columns else 0
            
            st.success(f"**🛠️ 本期计划总产量： {total_prod_qty:,.0f} 件 | 投入资金预估： ¥ {total_prod_amt:,.0f}**")
            st.dataframe(style_dataframe(df_production), use_container_width=True, hide_index=True)
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