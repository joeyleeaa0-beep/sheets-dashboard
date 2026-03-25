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
    "直播号": "2qReEC",
    "信息流": "3aoFSL",
    "小红书": "4jVWrG",
    "视频号": "5xwXLW",
    "B站": "6lywkx",
    "快手": "7szWUE",
}

MONTHS = ["1月","2月","3月","4月","5月","6月","7月","8月","9月","10月","11月","12月"]
CITIES = ["深圳", "上海", "成都", "天津"]

@st.cache_data(ttl=300)
def get_token():
    res = requests.post(
        "https://open.feishu.cn/open-apis/auth/v3/tenant_access_token/internal",
        json={"app_id": APP_ID, "app_secret": APP_SECRET}
    )
    return res.json().get("tenant_access_token")

@st.cache_data(ttl=300)
def get_spreadsheet_token():
    token = get_token()
    headers = {"Authorization": f"Bearer {token}"}
    res = requests.get(
        f"https://open.feishu.cn/open-apis/wiki/v2/spaces/get_node?token={WIKI_TOKEN}",
        headers=headers
    ).json()
    return res.get("data", {}).get("node", {}).get("obj_token")

@st.cache_data(ttl=300)
def read_sheet(spreadsheet_token, sheet_id, data_range="A1:Z500"):
    token = get_token()
    headers = {"Authorization": f"Bearer {token}"}
    # 用renderType=FORMULA获取计算结果而不是公式
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

def clean_main_df(df):
    if df.empty:
        return df
    
    # 处理月份列：合并单元格只有第一行有值，需要向下填充
    if "月份" in df.columns:
        df["月份"] = df["月份"].replace([None, "None", "", "nan"], pd.NA)
        df["月份"] = df["月份"].ffill()
        df["月份"] = df["月份"].astype(str).str.strip().str.replace(" ", "")
    
    # 处理地区列：同样可能有合并单元格
    if "地区" in df.columns:
        df["地区"] = df["地区"].replace([None, "None", "", "nan"], pd.NA)
        # 不向下填充地区，而是根据城市名过滤
        df = df[df["地区"].isin(CITIES)]
    
    # 转换数值列
    num_cols = ["投放金额", "总成交量", "销售量", "收购量", "直播成交", "视频成交"]
    for col in num_cols:
        if col in df.columns:
            df[col] = to_num(df[col])
    
    # 找客资列
    for col in df.columns:
        if "客资" in col and "成本" not in col:
            df["客资量"] = to_num(df[col])
            break
    
    return df

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
    df_raw = read_sheet(spreadsheet_token, SHEETS["四地汇总"], "A2:K200")
    df = clean_main_df(df_raw.copy())

# ── 侧边栏筛选 ──
st.sidebar.header("🔍 筛选条件")
sel_city = st.sidebar.selectbox("筛选城市", ["全部城市"] + CITIES)
sel_month = st.sidebar.selectbox("筛选月份", ["全部月份"] + MONTHS)

df_filtered = df.copy()
if sel_city != "全部城市" and "地区" in df_filtered.columns:
    df_filtered = df_filtered[df_filtered["地区"] == sel_city]
if sel_month != "全部月份" and "月份" in df_filtered.columns:
    df_filtered = df_filtered[df_filtered["月份"] == sel_month]

st.sidebar.markdown(f"📋 当前数据：**{len(df_filtered)}** 条")

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
tab1, tab2, tab3 = st.tabs(["🏙️ 分城市数据", "📈 月度趋势", "📋 数据明细"])

with tab1:
    st.subheader("分城市经营对比")
    if "地区" in df_filtered.columns:
        city_group = df_filtered.groupby("地区").agg(
            投放金额=("投放金额", "sum"),
            客资量=("客资量", "sum"),
            总成交量=("总成交量", "sum"),
        ).reset_index()
        city_group["客资成本"] = (city_group["投放金额"] / city_group["客资量"]).round(2)
        city_group["成交成本"] = (city_group["投放金额"] / city_group["总成交量"]).round(2)
        city_group["成交率"] = (city_group["总成交量"] / city_group["客资量"] * 100).round(2)
        st.dataframe(city_group, use_container_width=True, hide_index=True)

        ca, cb, cc = st.columns(3)
        with ca:
            st.plotly_chart(px.bar(city_group, x="地区", y="客资量", title="各城市客资量", color="地区"), use_container_width=True)
        with cb:
            st.plotly_chart(px.bar(city_group, x="地区", y="总成交量", title="各城市成交量", color="地区"), use_container_width=True)
        with cc:
            st.plotly_chart(px.bar(city_group, x="地区", y="客资成本", title="各城市客资成本", color="地区"), use_container_width=True)

with tab2:
    st.subheader("月度趋势")
    if "月份" in df.columns:
        df_trend = df.copy()
        if sel_city != "全部城市":
            df_trend = df_trend[df_trend["地区"] == sel_city]
        month_group = df_trend.groupby("月份").agg(
            投放金额=("投放金额", "sum"),
            客资量=("客资量", "sum"),
            总成交量=("总成交量", "sum"),
        ).reset_index()
        month_group["月份"] = pd.Categorical(month_group["月份"], categories=MONTHS, ordered=True)
        month_group = month_group.sort_values("月份")

        st.plotly_chart(px.line(month_group, x="月份", y=["投放金额", "客资量"], title="月度投放与客资趋势", markers=True), use_container_width=True)
        st.plotly_chart(px.bar(month_group, x="月份", y="总成交量", title="月度成交量", color="月份"), use_container_width=True)

with tab3:
    st.subheader("数据明细")
    st.dataframe(df_filtered.dropna(axis=1, how='all'), use_container_width=True)
