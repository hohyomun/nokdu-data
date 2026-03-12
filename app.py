"""
绿豆品牌销售业务与供应链全景看板
Streamlit Dashboard - app.py
"""

import streamlit as st
import pandas as pd
import numpy as np
import plotly.graph_objects as go
import plotly.express as px
import io

# ──────────────────────────────────────────────
# 页面配置（必须是第一个 st 调用）
# ──────────────────────────────────────────────
st.set_page_config(
    page_title="绿豆品牌销售业务与供应链全景看板",
    page_icon="🌿"，
    layout="wide"，
    initial_sidebar_state="expanded"，
)

# ──────────────────────────────────────────────
# 全局 CSS（高级哑光商务感）
# ──────────────────────────────────────────────
st.markdown("""
<style>
/* 全局字体与背景 */
@import url('https://fonts.googleapis.com/css2?family=Inter:wght@300;400;500;600;700&display=swap');

html, body, [class*="css"] {
    font-family: 'Inter', -apple-system, BlinkMacSystemFont, sans-serif;
}

/* 主背景 */
.stApp {
    background-color: #F0F2F6;
}

/* 隐藏 Streamlit 默认 header / footer，但保留侧边栏开关按钮 */
#MainMenu { visibility: hidden; }
footer { visibility: hidden; }
header[data-testid="stHeader"] { background: transparent; }

/* 保留侧边栏展开/收起按钮 */
button[kind="headerNoPadding"],
section[data-testid="collapsedControl"],
div[data-testid="collapsedControl"] {
    display: block !important;
    visibility: visible !important;
    opacity: 1 !important;
}

/* 侧边栏样式 */
[data-testid="stSidebar"] {
    background: linear-gradient(180deg, #1C2B4A 0%, #243655 100%);
    padding-top: 1.5rem;
}
[data-testid="stSidebar"] * {
    color: #C8D6E8 !important;
}
[data-testid="stSidebar"] h1,
[data-testid="stSidebar"] h2，
[data-testid="stSidebar"] h3 {
    color: #FFFFFF !important;
}

/* KPI 指标卡片 */
[data-testid="metric-container"] {
    background: #FFFFFF;
    border-radius: 12px;
    padding: 18px 22px !important;
    box-shadow: 0 2px 12px rgba(28, 43, 74, 0.08);
    border-left: 4px solid #2563EB;
    margin-bottom: 8px;
}
[data-testid="metric-container"] label {
    color: #64748B !important;
    font-size: 0.82rem !important;
    font-weight: 500 !important;
    letter-spacing: 0.02em;
    text-transform: uppercase;
}
[data-testid="metric-container"] [data-testid="stMetricValue"] {
    color: #1C2B4A !important;
    font-size: 1.8rem !important;
    font-weight: 700 !important;
}
[data-testid="metric-container"] [data-testid="stMetricDelta"] {
    font-size: 0.85rem !important;
}

/* Tab 样式 */
[data-testid="stTabs"] [data-baseweb="tab-list"] {
    background: #FFFFFF;
    border-radius: 10px;
    padding: 4px;
    box-shadow: 0 1px 6px rgba(28,43,74,0.07);
    margin-bottom: 1.2rem;
    gap: 4px;
}
[data-testid="stTabs"] [data-baseweb="tab"] {
    border-radius: 8px !important;
    font-weight: 500 !important;
    color: #64748B !important;
    padding: 8px 18px !important;
    transition: all 0.2s;
}
[data-testid="stTabs"] [aria-selected="true"] {
    background: #2563EB !important;
    color: #FFFFFF !important;
}

/* 数据表格 */
[data-testid="stDataFrame"] {
    border-radius: 10px;
    overflow: hidden;
    box-shadow: 0 1px 8px rgba(28,43,74,0.07);
}

/* 分割线 */
hr {
    border: none;
    border-top: 2px solid #E2E8F0;
    margin: 1.5rem 0;
}

/* 模块标题 */
.section-title {
    font-size: 1.1rem;
    font-weight: 600;
    color: #1C2B4A;
    border-left: 4px solid #2563EB;
    padding-left: 12px;
    margin: 1.2rem 0 0.8rem 0;
}

/* 主标题 */
.main-title {
    font-size: 1.6rem;
    font-weight: 700;
    color: #1C2B4A;
    letter-spacing: -0.02em;
    margin-bottom: 0.2rem;
}
.main-subtitle {
    font-size: 0.88rem;
    color: #64748B;
    margin-bottom: 1.5rem;
}

/* 警告/高亮卡片 */
.highlight-card {
    background: linear-gradient(135deg, #2563EB 0%, #1D4ED8 100%);
    color: white;
    border-radius: 12px;
    padding: 16px 20px;
    margin-bottom: 1rem;
}
.highlight-card .label {
    font-size: 0.78rem;
    opacity: 0.85;
    text-transform: uppercase;
    letter-spacing: 0.05em;
    margin-bottom: 4px;
}
.highlight-card .value {
    font-size: 1.6rem;
    font-weight: 700;
}

/* 登录页 */
.login-container {
    max-width: 380px;
    margin: 80px auto;
    background: white;
    border-radius: 16px;
    padding: 40px 36px;
    box-shadow: 0 8px 40px rgba(28,43,74,0.14);
}
</style>
""", unsafe_allow_html=True)

# ──────────────────────────────────────────────
# 认证配置（可在此修改用户名密码）
# ──────────────────────────────────────────────
CREDENTIALS = {
    "admin": "lvdou2026",
    "viewer": "view2026",
}

# ──────────────────────────────────────────────
# Session State 初始化
# ──────────────────────────────────────────────
if "authenticated" not in st.session_state:
    st.session_state.authenticated = False
if "username" not in st.session_state:
    st.session_state.username = ""

# ──────────────────────────────────────────────
# 工具函数
# ──────────────────────────────────────────────

def fmt_number(val):
    """财务格式：千位分隔符,不保留小数"""
    try:
        return f"{int(round(float(val))):,}"
    except Exception:
        return str(val)

def fmt_pct(val):
    """百分比格式"""
    try:
        return f"{float(val)*100:.1f}%"
    except Exception:
        return str(val)

def truncate_name(name, max_len=18):
    """截断产品名称,超过 max_len 显示 …"""
    s = str(name)
    return s if len(s) <= max_len else s[:max_len] + "…"

def find_sheet(xls, keywords):
    """
    按关键字列表查找 Sheet（模糊匹配）。
    keywords: 列表,任意一个命中即可。
    返回第一个匹配的 sheet 名,否则 None。
    """
    for sheet_name in xls.sheet_names:
        for kw in keywords:
            if kw in sheet_name:
                return sheet_name
    return None

def smart_read_sheet(xls, sheet_name, header_keywords):
    """
    智能扫描 Sheet,找到包含所有 header_keywords 的行作为表头,
    然后返回从该行开始读取的 DataFrame。
    过滤掉小计/合计/空行。
    """
    # 不指定 header,全部读取
    df_raw = pd.read_excel(xls, sheet_name=sheet_name, header=None)

    header_row = None
    for i, row in df_raw.iterrows():
        row_str = " ".join([str(v) for v in row.values if pd.notna(v)])
        if all(kw in row_str for kw in header_keywords):
            header_row = i
            break

    if header_row is None:
        return None, f"未找到包含 {header_keywords} 的表头行"

    # 以该行为 header 重新读取
    df = pd.read_excel(xls, sheet_name=sheet_name, header=header_row)

    # 清理列名：去除 Unnamed 列和空列
    df.columns = [str(c).strip() if "Unnamed" not in str(c) else "" for c in df.columns]
    df = df.loc[:, df.columns != ""]

    # 过滤掉小计/合计/空行
    def is_noise_row(row):
        for val in row.values:
            v = str(val).strip()
            if v in ("小计", "合计", "总计", "nan", ""):
                continue
            return False
        return True  # 全是噪音值

    noise_keywords = ["小计", "合计", "总计", "汇总"]

    def row_is_noise(row):
        for val in row.values:
            v = str(val).strip()
            if any(kw in v for kw in noise_keywords):
                return True
        return False

    df = df[~df.apply(row_is_noise, axis=1)]
    df = df.dropna(how="all")
    df = df.reset_index(drop=True)

    return df, None

def to_numeric_safe(series):
    return pd.to_numeric(series, errors="coerce").fillna(0)


# ──────────────────────────────────────────────
# 登录页面
# ──────────────────────────────────────────────
def show_login():
    st.markdown("""
    <div style="text-align:center; margin-top: 60px; margin-bottom: 30px;">
        <div style="font-size:3rem;">🌿</div>
        <h2 style="color:#1C2B4A; font-weight:700; margin-bottom:4px;">绿豆品牌销售业务与供应链全景看板</h2>
        <p style="color:#64748B; font-size:0.9rem;">请输入授权账号登录</p>
    </div>
    """, unsafe_allow_html=True)

    col1, col2, col3 = st.columns([1, 1.4, 1])
    with col2:
        with st.container():
            st.markdown("""
            <div style="background:white; border-radius:16px; padding:32px 28px;
                        box-shadow:0 8px 40px rgba(28,43,74,0.12);">
            """, unsafe_allow_html=True)

            username = st.text_input("用户名", placeholder="请输入用户名", key="login_user")
            password = st.text_input("密码", type="password", placeholder="请输入密码", key="login_pass")

            if st.button("登 录", use_container_width=True, type="primary"):
                if username in CREDENTIALS and CREDENTIALS[username] == password:
                    st.session_state.authenticated = True
                    st.session_state.username = username
                    st.rerun()
                else:
                    st.error("用户名或密码错误,请重试。")

            st.markdown("</div>", unsafe_allow_html=True)

    st.markdown("""
    <p style="text-align:center; color:#94A3B8; font-size:0.78rem; margin-top:24px;">
    默认账号：admin / lvdou2026 &nbsp;|&nbsp; 只读账号：viewer / view2026
    </p>
    """, unsafe_allow_html=True)


# ──────────────────────────────────────────────
# 主看板
# ──────────────────────────────────────────────
def show_dashboard():
    # 侧边栏
    with st.sidebar:
        st.markdown("### 🌿 数据控制台")
        st.markdown(f"<small>当前用户：{st.session_state.username}</small>", unsafe_allow_html=True)
        st.divider()

        uploaded_file = st.file_uploader(
            "📂 上传每周数据 Excel",
            type=["xlsx", "xls"],
            help="上传包含多个 Sheet 的复杂 Excel 表格,上传后自动刷新所有数据"
        )

        if uploaded_file:
            st.success(f"✅ 已加载：{uploaded_file.name}")

        st.divider()
        st.markdown("<small>数据说明：请每周上传最新版 Excel,看板将自动解析所有 Sheet 并更新图表。</small>",
                    unsafe_allow_html=True)
        st.divider()

        if st.button("🚪 退出登录", use_container_width=True):
            st.session_state.authenticated = False
            st.rerun()

    # 主标题
    st.markdown("""
    <div class="main-title">🌿 绿豆品牌销售业务与供应链全景看板</div>
    <div class="main-subtitle">实时业务数据整合 · 智能分析 · 供应链全景</div>
    """, unsafe_allow_html=True)

    if uploaded_file is None:
        st.info("👈 请在左侧侧边栏上传 Excel 数据文件,看板将自动解析并展示所有数据。")
        st.stop()

    # 读取 Excel
    try:
        xls = pd.ExcelFile(uploaded_file)
    except Exception as e:
        st.error(f"Excel 文件读取失败：{e}")
        st.stop()

    with st.expander("📋 检测到的 Sheet 列表（点击展开）", expanded=False):
        st.write(xls.sheet_names)

    # ───────────────────────────
    # 四个 Tab
    # ───────────────────────────
    tab1, tab2, tab3, tab4 = st.tabs([
        "💰 全年销售业绩",
        "📦 在库成品全景",
        "💳 未执行订单与生产",
        "🚀 核心业务推进跟踪"
    ])

    # ═══════════════════════════════════════════
    # Tab 1 — 全年销售业绩
    # ═══════════════════════════════════════════
    with tab1:
        sheet_name = find_sheet(xls, ["销售统计"])
        if not sheet_name:
            st.warning("未找到包含「销售统计」的 Sheet,请确认 Excel 文件格式。")
        else:
            df_sales, err = smart_read_sheet(xls, sheet_name, ["品牌", "区分"])
            if err:
                st.warning(f"销售统计数据解析问题：{err}")
                df_sales = None

        if df_sales is not None and not df_sales.empty:
            # ——— 自动识别月度列 ———
            # 尝试找目标、实绩相关列
            all_cols = df_sales.columns.tolist()

            # 找年份行（包含 2026 字样的行作为全年数据）
            def find_year_rows(df, year_kw):
                """找包含 year_kw 的行"""
                rows = []
                for col in df.columns:
                    mask = df[col].astype(str).str.contains(str(year_kw), na=False)
                    if mask.any():
                        rows = df[mask].index.tolist()
                        return rows, col
                return [], None

            # ——— KPI 提取 ———
            # 尝试找目标/实绩列（常见命名：目标、实绩、完成、达成等）
            target_col = None
            actual_col = None
            for c in all_cols:
                c_str = str(c)
                if any(kw in c_str for kw in ["目标", "计划"]) and target_col is None:
                    target_col = c
                if any(kw in c_str for kw in ["实绩", "完成", "实际"]) and actual_col is None:
                    actual_col = c

            # 月份列（1月~12月）
            month_cols_target = {}
            month_cols_actual = {}
            for c in all_cols:
                c_str = str(c)
                for m in range(1, 13):
                    if f"{m}月" in c_str or f"{m:02d}月" in c_str:
                        if "目标" in c_str or "计划" in c_str:
                            month_cols_target[m] = c
                        elif "实绩" in c_str or "实际" in c_str or "完成" in c_str:
                            month_cols_actual[m] = c

            # 尝试找 2026 行
            row_2026 = None
            for idx, row in df_sales.iterrows():
                row_str = " ".join([str(v) for v in row.values])
                if "2026" in row_str:
                    row_2026 = row
                    break

            # 总目标/实绩
            total_target = 0
            total_actual = 0

            if row_2026 is not None:
                if target_col:
                    total_target = float(pd.to_numeric(row_2026.get(target_col, 0), errors="coerce") or 0)
                if actual_col:
                    total_actual = float(pd.to_numeric(row_2026.get(actual_col, 0), errors="coerce") or 0)

                # 如果专门的全年列找不到,尝试从月度列求和
                if total_target == 0 and month_cols_target:
                    for m, c in month_cols_target.items():
                        total_target += float(pd.to_numeric(row_2026.get(c, 0), errors="coerce") or 0)
                if total_actual == 0 and month_cols_actual:
                    for m, c in month_cols_actual.items():
                        total_actual += float(pd.to_numeric(row_2026.get(c, 0), errors="coerce") or 0)

            # 如果按行找不到,尝试按全列汇总
            if total_target == 0 and target_col:
                total_target = pd.to_numeric(df_sales[target_col], errors="coerce").sum()
            if total_actual == 0 and actual_col:
                total_actual = pd.to_numeric(df_sales[actual_col], errors="coerce").sum()

            achieve_rate = (total_actual / total_target * 100) if total_target > 0 else 0

            # ——— 顶部 KPI ———
            st.markdown('<div class="section-title">📊 2026 全年核心 KPI</div>', unsafe_allow_html=True)
            kpi1, kpi2, kpi3 = st.columns(3)
            with kpi1:
                st.metric("2026 全年销售目标", f"¥ {fmt_number(total_target)}")
            with kpi2:
                st.metric("当前累计已完成实绩", f"¥ {fmt_number(total_actual)}")
            with kpi3:
                delta_color = "normal" if achieve_rate >= 100 else "inverse"
                st.metric("年度整体达成率", f"{achieve_rate:.1f}%",
                          delta=f"{'✅ 达标' if achieve_rate >= 100 else '⚠️ 未达标'}")

            # ——— 月度对比柱状图 ———
            st.markdown("---")
            st.markdown('<div class="section-title">📈 2026 年 1-12 月 目标 vs 实际销量对比</div>',
                        unsafe_allow_html=True)

            # 构建月度数据
            months = [f"{m}月" for m in range(1, 13)]
            target_vals = []
            actual_vals = []

            for m in range(1, 13):
                t_val = 0
                a_val = 0
                if row_2026 is not None:
                    if m in month_cols_target:
                        t_val = float(pd.to_numeric(row_2026.get(month_cols_target[m], 0), errors="coerce") or 0)
                    if m in month_cols_actual:
                        a_val = float(pd.to_numeric(row_2026.get(month_cols_actual[m], 0), errors="coerce") or 0)
                target_vals.append(t_val)
                actual_vals.append(a_val)

            # 如果月度列找不到,展示原始数据并提示
            if all(v == 0 for v in target_vals) and all(v == 0 for v in actual_vals):
                st.info("💡 未能自动识别月度目标/实绩列（建议列名含 X月目标 或 X月实绩）,以下展示原始销售统计数据：")
                # 数字列展示
                num_cols = df_sales.select_dtypes(include=[np.number]).columns.tolist()
                if num_cols:
                    st.dataframe(df_sales[all_cols[:min(10, len(all_cols))]].style.hide(axis="index"),
                                 use_container_width=True)
            else:
                fig_bar = go.Figure()
                fig_bar.add_trace(go.Bar(
                    name="目标销量",
                    x=months,
                    y=target_vals,
                    marker_color="#93C5FD",
                    text=[fmt_number(v) if v > 0 else "" for v in target_vals],
                    textposition="outside",
                ))
                fig_bar.add_trace(go.Bar(
                    name="实际销量",
                    x=months,
                    y=actual_vals,
                    marker_color="#2563EB",
                    text=[fmt_number(v) if v > 0 else "" for v in actual_vals],
                    textposition="outside",
                ))
                fig_bar.update_layout(
                    barmode="group",
                    plot_bgcolor="white",
                    paper_bgcolor="white",
                    font=dict(family="Inter, sans-serif", size=13),
                    legend=dict(orientation="h", yanchor="bottom", y=1.02, xanchor="right", x=1),
                    margin=dict(l=20, r=20, t=40, b=20),
                    height=400,
                    yaxis=dict(showgrid=True, gridcolor="#F1F5F9"),
                    xaxis=dict(showgrid=False),
                )
                st.plotly_chart(fig_bar, use_container_width=True)

            # ——— 历史回溯数据 ———
            st.markdown("---")
            st.markdown('<div class="section-title">📂 历史年度数据回溯（2023 / 2024 / 2025）</div>',
                        unsafe_allow_html=True)

            history_rows = []
            for year in [2023, 2024, 2025]:
                for idx, row in df_sales.iterrows():
                    row_str = " ".join([str(v) for v in row.values])
                    if str(year) in row_str:
                        row_dict = {"年份": str(year)}
                        # 全年列
                        if actual_col and pd.notna(row.get(actual_col)):
                            row_dict["全年累计实绩额"] = fmt_number(
                                pd.to_numeric(row.get(actual_col, 0), errors="coerce") or 0)
                        elif target_col and pd.notna(row.get(target_col)):
                            row_dict["全年累计实绩额"] = fmt_number(
                                pd.to_numeric(row.get(target_col, 0), errors="coerce") or 0)
                        # 月度
                        for m in range(1, 13):
                            for mc in [month_cols_actual, month_cols_target]:
                                if m in mc and pd.notna(row.get(mc[m])):
                                    v = pd.to_numeric(row.get(mc[m], 0), errors="coerce") or 0
                                    row_dict[f"{m}月"] = fmt_number(v)
                                    break
                        history_rows.append(row_dict)
                        break

            if history_rows:
                df_history = pd.DataFrame(history_rows)
                st.dataframe(df_history.style.hide(axis="index"), use_container_width=True)
            else:
                st.info("未识别到 2023/2024/2025 历史行,请确认 Excel 中含年份标记。展示原始数据供参考：")
                st.dataframe(df_sales.head(20).style.hide(axis="index"), use_container_width=True)

        else:
            st.warning("销售统计 Sheet 数据为空或解析失败,请检查文件格式。")

    # ═══════════════════════════════════════════
    # Tab 2 — 在库成品全景
    # ═══════════════════════════════════════════
    with tab2:
        sheet_inv = find_sheet(xls, ["成品库存"])
        df_inv = None
        if not sheet_inv:
            st.warning("未找到包含「成品库存」的 Sheet。")
        else:
            df_inv, err = smart_read_sheet(xls, sheet_inv, ["产品名称", "实际库存量"])
            if err:
                st.warning(f"成品库存数据解析：{err}")
                df_inv = None

        if df_inv is not None and not df_inv.empty:
            # 找关键列
            product_col = next((c for c in df_inv.columns if "产品名称" in str(c)), None)
            qty_col = next((c for c in df_inv.columns if "实际库存量" in str(c)), None)
            amount_col = next((c for c in df_inv.columns if any(kw in str(c) for kw in ["金额", "库存金额", "总额", "价值"])), None)

            if product_col:
                df_inv[product_col] = df_inv[product_col].astype(str).str.strip()
                df_inv = df_inv[~df_inv[product_col].isin(["nan", "", "None"])]

            if qty_col:
                df_inv[qty_col] = to_numeric_safe(df_inv[qty_col])

            total_qty = df_inv[qty_col].sum() if qty_col else 0
            total_amount = df_inv[amount_col].apply(lambda x: pd.to_numeric(x, errors="coerce")).sum() if amount_col else 0

            # KPI
            st.markdown('<div class="section-title">📊 库存核心 KPI</div>', unsafe_allow_html=True)
            kpi_c1, kpi_c2 = st.columns(2)
            with kpi_c1:
                st.metric("冻结资金（在库总金额）",
                          f"¥ {fmt_number(total_amount)}" if total_amount > 0 else "（需含金额列）")
            with kpi_c2:
                st.metric("物资规模（成品总库存数）", f"{fmt_number(total_qty)} 件")

            st.markdown("---")

            # ——— 图表 1：资金占用预警 Top 10 横向柱状图 ———
            st.markdown('<div class="section-title">💰 资金占用预警 Top 10</div>', unsafe_allow_html=True)

            if amount_col and product_col:
                df_inv_sorted_amt = df_inv[[product_col, amount_col]].copy()
                df_inv_sorted_amt[amount_col] = to_numeric_safe(df_inv_sorted_amt[amount_col])
                df_inv_sorted_amt = df_inv_sorted_amt.sort_values(amount_col, ascending=False).head(10)

                # 截断名称（Y轴）,保留全名（hover）
                df_inv_sorted_amt["short_name"] = df_inv_sorted_amt[product_col].apply(truncate_name)
                df_inv_sorted_amt["full_name"] = df_inv_sorted_amt[product_col]

                fig_bar_h = go.Figure(go.Bar(
                    x=df_inv_sorted_amt[amount_col],
                    y=df_inv_sorted_amt["short_name"],
                    orientation="h",
                    marker=dict(
                        color=df_inv_sorted_amt[amount_col],
                        colorscale=[[0, "#93C5FD"], [1, "#1D4ED8"]],
                        showscale=False,
                    ),
                    customdata=df_inv_sorted_amt["full_name"],
                    hovertemplate="<b>%{customdata}</b><br>金额：¥ %{x:,.0f}<extra></extra>",
                    text=[fmt_number(v) for v in df_inv_sorted_amt[amount_col]],
                    textposition="outside",
                ))
                fig_bar_h.update_layout(
                    plot_bgcolor="white",
                    paper_bgcolor="white",
                    height=420,
                    margin=dict(l=10, r=60, t=20, b=20),
                    xaxis=dict(showgrid=True, gridcolor="#F1F5F9", tickformat=","),
                    yaxis=dict(autorange="reversed"),
                    font=dict(family="Inter, sans-serif", size=12),
                )
                st.plotly_chart(fig_bar_h, use_container_width=True)
            else:
                st.info("未识别到金额相关列,请确认列名含「金额」、「价值」等关键字。")

            st.markdown("---")

            # ——— 图表 2：库存件数分布占比 Top 10 饼图 ———
            st.markdown('<div class="section-title">📦 核心库存件数分布占比 Top 10</div>',
                        unsafe_allow_html=True)

            if qty_col and product_col:
                df_pie = df_inv[[product_col, qty_col]].copy()
                df_pie[qty_col] = to_numeric_safe(df_pie[qty_col])
                df_pie = df_pie[df_pie[qty_col] > 0].sort_values(qty_col, ascending=False).head(10)
                df_pie["short_name"] = df_pie[product_col].apply(truncate_name)

                fig_pie = go.Figure(go.Pie(
                    labels=df_pie["short_name"],
                    values=df_pie[qty_col],
                    customdata=df_pie[product_col],
                    hovertemplate="<b>%{customdata}</b><br>库存：%{value:,.0f} 件<br>占比：%{percent}<extra></extra>",
                    textinfo="percent",
                    hole=0.35,
                    marker=dict(colors=px.colors.qualitative.Set2),
                ))
                fig_pie.update_layout(
                    height=520,
                    legend=dict(
                        orientation="h",
                        yanchor="bottom",
                        y=-0.30,
                        xanchor="center",
                        x=0.5,
                        font=dict(size=11),
                    ),
                    margin=dict(l=20, r=20, t=20, b=80),
                    paper_bgcolor="white",
                    font=dict(family="Inter, sans-serif"),
                )
                st.plotly_chart(fig_pie, use_container_width=True)

            st.markdown("---")

            # ——— 全量库存明细 ———
            st.markdown('<div class="section-title">📋 成品库存明细表（全量）</div>', unsafe_allow_html=True)
            st.dataframe(df_inv.style.hide(axis="index"), use_container_width=True)

        else:
            st.warning("成品库存数据为空或解析失败。")

    # ═══════════════════════════════════════════
    # Tab 3 — 未执行订单与生产
    # ═══════════════════════════════════════════
    with tab3:
        # ——— 未执行订单 ———
        sheet_order = find_sheet(xls, ["未执行订单", "预付款管理"])
        df_order = None
        if sheet_order:
            df_order, err = smart_read_sheet(xls, sheet_order, ["订单编号", "订单金额"])
            if err:
                st.warning(f"未执行订单解析：{err}")
                df_order = None

        st.markdown('<div class="section-title">📋 客户未执行订单明细</div>', unsafe_allow_html=True)

        if df_order is not None and not df_order.empty:
            order_no_col = next((c for c in df_order.columns if "订单编号" in str(c)), None)
            order_amt_col = next((c for c in df_order.columns if "订单金额" in str(c)), None)

            if order_no_col:
                df_order[order_no_col] = df_order[order_no_col].astype(str).str.strip()
                df_order = df_order[~df_order[order_no_col].isin(["nan", "", "None"])]

            total_order_amt = 0
            if order_amt_col:
                df_order[order_amt_col] = to_numeric_safe(df_order[order_amt_col])
                total_order_amt = df_order[order_amt_col].sum()

            st.markdown(f"""
            <div class="highlight-card">
                <div class="label">📌 有效订单池总金额</div>
                <div class="value">¥ {fmt_number(total_order_amt)}</div>
            </div>
            """, unsafe_allow_html=True)

            st.dataframe(df_order.style.hide(axis="index"), use_container_width=True)
        else:
            st.warning("未找到「未执行订单」或「预付款管理」Sheet,或数据解析失败。")

        st.markdown("---")

        # ——— 生产计划 ———
        sheet_prod = find_sheet(xls, ["生产计划"])
        df_prod = None
        if sheet_prod:
            df_prod, err = smart_read_sheet(xls, sheet_prod, ["产品名称", "数量"])
            if err:
                st.warning(f"生产计划解析：{err}")
                df_prod = None

        st.markdown('<div class="section-title">🏭 当期工厂生产排期计划</div>', unsafe_allow_html=True)

        if df_prod is not None and not df_prod.empty:
            prod_name_col = next((c for c in df_prod.columns if "产品名称" in str(c)), None)
            prod_qty_col = next((c for c in df_prod.columns if "数量" in str(c)), None)
            prod_price_col = next((c for c in df_prod.columns
                                   if any(kw in str(c) for kw in ["单价", "价格", "金额", "成本"])), None)

            if prod_name_col:
                df_prod[prod_name_col] = df_prod[prod_name_col].astype(str).str.strip()
                df_prod = df_prod[~df_prod[prod_name_col].isin(["nan", "", "None"])]

            total_plan_qty = 0
            total_invest = 0
            if prod_qty_col:
                df_prod[prod_qty_col] = to_numeric_safe(df_prod[prod_qty_col])
                total_plan_qty = df_prod[prod_qty_col].sum()

            if prod_price_col:
                df_prod[prod_price_col] = to_numeric_safe(df_prod[prod_price_col])
                total_invest = df_prod[prod_price_col].sum()

            kpi_p1, kpi_p2 = st.columns(2)
            with kpi_p1:
                st.markdown(f"""
                <div class="highlight-card">
                    <div class="label">📦 本期计划总产量</div>
                    <div class="value">{fmt_number(total_plan_qty)} 件</div>
                </div>
                """, unsafe_allow_html=True)
            with kpi_p2:
                st.markdown(f"""
                <div class="highlight-card" style="background:linear-gradient(135deg,#059669,#047857)">
                    <div class="label">💰 投入资金预估</div>
                    <div class="value">¥ {fmt_number(total_invest)}</div>
                </div>
                """, unsafe_allow_html=True)

            st.dataframe(df_prod.style.hide(axis="index"), use_container_width=True)
        else:
            st.warning("未找到「生产计划」Sheet,或数据解析失败。")

    # ═══════════════════════════════════════════
    # Tab 4 — 核心业务推进跟踪
    # ═══════════════════════════════════════════
    with tab4:
        col_left, col_right = st.columns(2)

        # ——— 左列：备案未产 ———
        with col_left:
            st.markdown('<div class="section-title">📁 备案未产项目</div>', unsafe_allow_html=True)
            sheet_ba = find_sheet(xls, ["备案未产"])
            df_ba = None
            if sheet_ba:
                df_ba, err = smart_read_sheet(xls, sheet_ba, ["产品编号", "产品名称"])
                if err:
                    st.warning(f"备案未产解析：{err}")
                    df_ba = None

            if df_ba is not None and not df_ba.empty:
                ba_name_col = next((c for c in df_ba.columns if "产品名称" in str(c)), None)
                if ba_name_col:
                    df_ba[ba_name_col] = df_ba[ba_name_col].astype(str).str.strip()
                    df_ba = df_ba[~df_ba[ba_name_col].isin(["nan", "", "None"])]

                st.metric("备案未产产品数量", f"{len(df_ba)} 项")
                st.dataframe(df_ba.style.hide(axis="index"), use_container_width=True, height=500)
            else:
                st.warning("未找到「备案未产」Sheet,或数据解析失败。")

        # ——— 右列：进行中业务 ———
        with col_right:
            st.markdown('<div class="section-title">🚀 进行中业务跟踪</div>', unsafe_allow_html=True)
            sheet_biz = find_sheet(xls, ["进行中业务"])
            df_biz = None
            if sheet_biz:
                df_biz, err = smart_read_sheet(xls, sheet_biz, ["项目名称", "业务阶段"])
                if err:
                    st.warning(f"进行中业务解析：{err}")
                    df_biz = None

            if df_biz is not None and not df_biz.empty:
                proj_col = next((c for c in df_biz.columns if "项目名称" in str(c)), None)
                stage_col = next((c for c in df_biz.columns if "业务阶段" in str(c)), None)

                if proj_col:
                    df_biz[proj_col] = df_biz[proj_col].astype(str).str.strip()
                    df_biz = df_biz[~df_biz[proj_col].isin(["nan", "", "None"])]

                st.metric("进行中业务项目数", f"{len(df_biz)} 项")

                # 按业务阶段着色
                if stage_col:
                    stage_counts = df_biz[stage_col].value_counts()
                    fig_stage = go.Figure(go.Bar(
                        x=stage_counts.values,
                        y=stage_counts.index.astype(str),
                        orientation="h",
                        marker_color="#2563EB",
                        text=stage_counts.values,
                        textposition="outside",
                    ))
                    fig_stage.update_layout(
                        plot_bgcolor="white",
                        paper_bgcolor="white",
                        height=200,
                        margin=dict(l=10, r=40, t=10, b=10),
                        yaxis=dict(autorange="reversed"),
                        font=dict(family="Inter, sans-serif", size=12),
                    )
                    st.plotly_chart(fig_stage, use_container_width=True)

                st.dataframe(df_biz.style.hide(axis="index"), use_container_width=True, height=400)
            else:
                st.warning("未找到「进行中业务」Sheet,或数据解析失败。")


# ──────────────────────────────────────────────
# 主入口
# ──────────────────────────────────────────────
def main():
    if not st.session_state.authenticated:
        show_login()
    else:
        show_dashboard()


if __name__ == "__main__":
    main()
