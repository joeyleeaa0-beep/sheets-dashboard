import streamlit as st
import pandas as pd
import requests
import plotly.express as px

st.set_page_config(page_title="新媒体数据看板", page_icon="📊", layout="wide")

APP_ID = st.secrets["FEISHU_APP_ID"]
APP_SECRET = st.secrets["FEISHU_APP_SECRET"]
WIKI_TOKEN = "G0VbwCGGMiEVNlktEYwcoA4JnHd"

SHEETS = {
    "四地汇总": "0HBHQk",
    "细分项汇总": "1hOmjY",
}

MONTHS = ["1月","2月","3月","4月","5月","6月","7月","8月","9月","10月","11月","12月"]
CITIES = ["深圳", "上海", "成都", "天津"]

@st.cache_data(ttl=60)
def get_token():
    res = requests.post(
        "https://open.feishu.cn/open-apis/auth/v3/tenant_access_token/internal",
        json={"app_id": APP_ID, "app_secret": APP_SECRET}
    )
    return res.json().get("tenant_access_token")

@st.cache_data(ttl=60)
def get_spreadsheet_token():
    token = get_token()
    headers = {"Authorization": f"Bearer {token}"}
    res = requests.get(
        f"https://open.feishu.cn/open-apis/wiki/v2/spaces/get_node?token={WIKI_TOKEN}",
        headers=headers
    ).json()
    return res.get("data", {}).get("node", {}).get("obj_token")

@st.cache_data(ttl=60)
def read_sheet(spreadsheet_token, sheet_id, data_range="A1:Z500"):
    token = get_token()
    headers = {"Authorization": f"Bearer {token}"}
    url = f"https://open.feishu.cn/open-apis/sheets/v2/spreadsheets/{spreadsheet_token}/values/{sheet_id}!{data_range}?renderType=FORMATTED_VALUE"
    res = requests.get(url, headers=headers).json()
    values = res.get("data", {}).get("valueRange", {}).get("values", [])
    if not values or len(values) < 2:
        return pd.DataFrame()
    headers_row = [str(h) if h else f"列{i}" for i, h in enumerate(values[0])]
    data_rows = values[1:]
    df = pd.DataFrame(data_rows, columns=headers_row)
    return df

def to_num(series):
    return pd.to_numeric(series, errors="coerce").fillna(0)

def fill_merged(df, col):
    """处理合并单元格：向下填充"""
    if col in df.columns:
        df[col] = df[col].replace([None, "None", "", "nan", "/"], pd.NA)
        df[col] = df[col].ffill()
    return df

def clean_main_df(df):
    if df.empty:
        return df
    df = fill_merged(df, "月份")
    if "月份" in df.columns:
        df["月份"] = df["月份"].astype(str).str.strip().str.replace(" ", "")
    if "地区" in df.columns:
        df["地区"] = df["地区"].replace([None, "None", "", "nan"], pd.NA)
        df = df[df["地区"].isin(CITIES)]
    num_cols = ["投放金额", "总成交量", "销售量", "收购量", "直播成交", "视频成交"]
    for col in num_cols:
        if col in df.columns:
            df[col] = to_num(df[col])
    for col in df.columns:
        if "客资" in col and "成本" not in col:
            df["客资量"] = to_num(df[col])
            break
    return df

def clean_detail_df(df):
    # 数值转换（只处理数值列）
    num_cols = ["投放金额", "销售量", "收购量", "素材更新量", "直播场次"]
    for col in num_cols:
        if col in df.columns:
            df[col] = df[col].astype(str).str.replace(r'IFERROR.*', '0', regex=True)
            df[col] = df[col].str.replace("/", "0")
            df[col] = to_num(df[col])
    
    # 客资列单独处理
    for col in df.columns:
        if "客资" in col and "成本" not in col and "直播" not in col and "视频" not in col:
            df["客资量"] = df[col].astype(str).str.replace("/", "0")
            df["客资量"] = df["客资量"].str.replace(r'IFERROR.*', '0', regex=True)
            df["客资量"] = to_num(df["客资量"])
            break
    
    # 重新计算总成交量 = 销售量 + 收购量
    if "销售量" in df.columns and "收购量" in df.columns:
        df["总成交量"] = to_num(df["销售量"]) + to_num(df["收购量"])
    
    # 渠道分类（在数值转换之后做）
    if "渠道/平台" in df.columns:
        df["渠道分类"] = df["渠道/平台"].astype(str).apply(lambda x:
            "抖音" if "抖音" in x else
            "信息流" if "信息流" in x else
            "小红书" if "小红书" in x else
            "视频号" if "视频号" in x else
            "B站" if "B站" in x else
            "快手" if "快手" in x else x
        )

# ── 加载数据 ──
with st.spinner("正在连接飞书..."):
    try:
        spreadsheet_token = get_spreadsheet_token()
        if not spreadsheet_token:
            st.error("无法获取表格Token")
            st.stop()
    except Exception as e:
        st.error(f"连接失败：{e}")
        st.stop()

with st.spinner("正在加载数据..."):
    df_main_raw = read_sheet(spreadsheet_token, SHEETS["四地汇总"], "A2:K200")
    df_main = clean_main_df(df_main_raw.copy())
    df_detail_raw = read_sheet(spreadsheet_token, SHEETS["细分项汇总"], "A1:P500")
    df_detail = clean_detail_df(df_detail_raw.copy())

# ── 侧边栏筛选 ──
st.sidebar.header("🔍 筛选条件")
sel_city = st.sidebar.selectbox("筛选城市", ["全部城市"] + CITIES)
sel_month = st.sidebar.selectbox("筛选月份", ["全部月份"] + MONTHS)

def apply_filter(df, city, month):
    d = df.copy()
    if city != "全部城市" and "地区" in d.columns:
        d = d[d["地区"] == city]
    if month != "全部月份" and "月份" in d.columns:
        d = d[d["月份"] == month]
    return d

df_filtered = apply_filter(df_main, sel_city, sel_month)
df_detail_filtered = apply_filter(df_detail, sel_city, sel_month)

st.sidebar.markdown(f"📋 四地汇总：**{len(df_filtered)}** 条")
st.sidebar.markdown(f"📋 渠道明细：**{len(df_detail_filtered)}** 条")

# ── 核心指标 ──
st.title("📊 新媒体数据看板")
st.caption(f"城市：{sel_city} ｜ 月份：{sel_month}")

total_spend = to_num(df_filtered["投放金额"]).sum() if "投放金额" in df_filtered.columns else 0
total_keizi = to_num(df_filtered.get("客资量", pd.Series())).sum()
total_chengjiao = to_num(df_filtered["总成交量"]).sum() if "总成交量" in df_filtered.columns else 0
total_xiaoshou = to_num(df_filtered["销售量"]).sum() if "销售量" in df_filtered.columns else 0
total_shougou = to_num(df_filtered["收购量"]).sum() if "收购量" in df_filtered.columns else 0

keizi_cost = total_spend / total_keizi if total_keizi > 0 else 0
chengjiao_cost = total_spend / total_chengjiao if total_chengjiao > 0 else 0
chengjiao_rate = total_chengjiao / total_keizi * 100 if total_keizi > 0 else 0

col1, col2, col3, col4, col5 = st.columns(5)
col1.metric("总投放金额", f"¥{total_spend:,.0f}")
col2.metric("总客资量", f"{int(total_keizi):,}")
col3.metric("总成交量", f"{int(total_chengjiao):,}")
col4.metric("销售总量", f"{int(total_xiaoshou):,}")
col5.metric("收购总量", f"{int(total_shougou):,}")

col6, col7, col8 = st.columns(3)
col6.metric("客资成本", f"¥{keizi_cost:.2f}")
col7.metric("成交成本", f"¥{chengjiao_cost:.2f}")
col8.metric("成交率", f"{chengjiao_rate:.2f}%")

st.divider()

# ── Tab ──
tab1, tab2, tab3, tab4 = st.tabs(["🏙️ 分城市数据", "📡 分渠道数据", "📈 趋势分析", "📋 数据明细"])

with tab1:
    st.subheader("分城市经营对比")
    if "地区" in df_main.columns:
        df_city = apply_filter(df_main, "全部城市", sel_month)
        city_group = df_city.groupby("地区").agg(
            投放金额=("投放金额", "sum"),
            客资量=("客资量", "sum"),
            总成交量=("总成交量", "sum"),
        ).reset_index()
        city_group["客资成本"] = (city_group["投放金额"] / city_group["客资量"].replace(0, pd.NA)).round(2)
        city_group["成交成本"] = (city_group["投放金额"] / city_group["总成交量"].replace(0, pd.NA)).round(2)
        city_group["成交率"] = (city_group["总成交量"] / city_group["客资量"].replace(0, pd.NA) * 100).round(2)
        st.dataframe(city_group, use_container_width=True, hide_index=True)

        ca, cb, cc, cd = st.columns(4)
        with ca:
            st.plotly_chart(px.bar(city_group, x="地区", y="客资量", title="各城市客资量", color="地区"), use_container_width=True)
        with cb:
            st.plotly_chart(px.bar(city_group, x="地区", y="总成交量", title="各城市成交量", color="地区"), use_container_width=True)
        with cc:
            st.plotly_chart(px.bar(city_group, x="地区", y="客资成本", title="各城市客资成本", color="地区"), use_container_width=True)
        with cd:
            st.plotly_chart(px.bar(city_group, x="地区", y="成交成本", title="各城市成交成本", color="地区"), use_container_width=True)

with tab2:
    st.subheader("分渠道数据对比")
    if not df_detail_filtered.empty and "渠道分类" in df_detail_filtered.columns:
        channel_group = df_detail_filtered.groupby("渠道分类").agg(
            投放金额=("投放金额", "sum"),
            客资量=("客资量", "sum"),
            总成交量=("总成交量", "sum"),
        ).reset_index()
        channel_group["客资成本"] = (channel_group["投放金额"] / channel_group["客资量"].replace(0, pd.NA)).round(2)
        channel_group = channel_group.sort_values("客资量", ascending=False)

        st.dataframe(channel_group, use_container_width=True, hide_index=True)

        ra, rb = st.columns(2)
        with ra:
            st.plotly_chart(px.bar(channel_group, x="渠道分类", y="客资量",
                title="各渠道客资量对比", color="渠道分类"), use_container_width=True)
        with rb:
            st.plotly_chart(px.bar(channel_group, x="渠道分类", y="投放金额",
                title="各渠道投放金额对比", color="渠道分类"), use_container_width=True)

        rc, rd = st.columns(2)
        with rc:
            st.plotly_chart(px.bar(channel_group, x="渠道分类", y="总成交量",
                title="各渠道成交量对比", color="渠道分类"), use_container_width=True)
        with rd:
            st.plotly_chart(px.bar(channel_group, x="渠道分类", y="客资成本",
                title="各渠道客资成本对比", color="渠道分类"), use_container_width=True)

        # 饼图
        re, rf = st.columns(2)
        with re:
            st.plotly_chart(px.pie(channel_group, names="渠道分类", values="客资量",
                title="各渠道客资量占比"), use_container_width=True)
        with rf:
            st.plotly_chart(px.pie(channel_group, names="渠道分类", values="投放金额",
                title="各渠道投放占比"), use_container_width=True)
    else:
        st.info("暂无渠道数据")

with tab3:
    st.subheader("趋势分析")

    # 各城市客资成本月度趋势
    if "月份" in df_main.columns and "地区" in df_main.columns:
        df_trend = df_main.copy()
        if sel_city != "全部城市":
            df_trend = df_trend[df_trend["地区"] == sel_city]

        city_month = df_trend.groupby(["月份", "地区"]).agg(
            投放金额=("投放金额", "sum"),
            客资量=("客资量", "sum"),
            总成交量=("总成交量", "sum"),
        ).reset_index()
        city_month["客资成本"] = (city_month["投放金额"] / city_month["客资量"].replace(0, pd.NA)).round(2)
        city_month["成交成本"] = (city_month["投放金额"] / city_month["总成交量"].replace(0, pd.NA)).round(2)
        city_month["月份"] = pd.Categorical(city_month["月份"], categories=MONTHS, ordered=True)
        city_month = city_month.sort_values("月份")

        st.plotly_chart(px.line(city_month, x="月份", y="客资成本", color="地区",
            title="各城市客资成本月度趋势", markers=True), use_container_width=True)

        st.plotly_chart(px.line(city_month, x="月份", y="成交成本", color="地区",
            title="各城市成交成本月度趋势", markers=True), use_container_width=True)

        # 大盘投放与客资趋势
        total_month = df_trend.groupby("月份").agg(
            投放金额=("投放金额", "sum"),
            客资量=("客资量", "sum"),
            总成交量=("总成交量", "sum"),
        ).reset_index()
        total_month["月份"] = pd.Categorical(total_month["月份"], categories=MONTHS, ordered=True)
        total_month = total_month.sort_values("月份")

        st.plotly_chart(px.line(total_month, x="月份", y=["客资量", "总成交量"],
            title="月度客资与成交趋势", markers=True), use_container_width=True)

with tab4:
    st.subheader("数据明细")
    inner_tab1, inner_tab2 = st.tabs(["四地汇总", "渠道明细"])
    with inner_tab1:
        st.dataframe(df_filtered.dropna(axis=1, how='all'), use_container_width=True)
    with inner_tab2:
        st.dataframe(df_detail_filtered.dropna(axis=1, how='all'), use_container_width=True)
