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
CHANNELS = ["抖音", "信息流", "小红书", "视频号", "B站", "快手"]

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
    if isinstance(series, pd.Series):
        s = series.astype(str)
        s = s.str.replace(r'IFERROR\(.*?\)', '0', regex=True)
        s = s.str.replace("/", "0")
        s = s.str.replace(",", "")
        return pd.to_numeric(s, errors="coerce").fillna(0)
    return 0

def fill_merged(df, col):
    if col in df.columns:
        df[col] = df[col].replace([None, "None", "", "nan", "/"], pd.NA)
        df[col] = df[col].ffill()
    return df

def apply_filter(df, city, month):
    if df is None or df.empty:
        return pd.DataFrame()
    d = df.copy()
    if city != "全部城市" and "地区" in d.columns:
        d = d[d["地区"] == city]
    if month != "全部月份" and "月份" in d.columns:
        d = d[d["月份"] == month]
    return d

def classify_channel(x):
    x = str(x)
    if "抖音" in x: return "抖音"
    if "信息流" in x: return "信息流"
    if "小红书" in x: return "小红书"
    if "视频号" in x: return "视频号"
    if "B站" in x: return "B站"
    if "快手" in x: return "快手"
    return x

def clean_main_df(df):
    if df.empty:
        return df
    df = fill_merged(df, "月份")
    if "月份" in df.columns:
        df["月份"] = df["月份"].astype(str).str.strip().str.replace(" ", "")
    if "地区" in df.columns:
        df["地区"] = df["地区"].replace([None, "None", "", "nan"], pd.NA)
        df = df[df["地区"].isin(CITIES)].copy()
    for col in ["投放金额", "总成交量", "销售量", "收购量"]:
        if col in df.columns:
            df[col] = to_num(df[col])
    for col in df.columns:
        if "客资" in col and "成本" not in col:
            df["客资量"] = to_num(df[col])
            break
    return df

def clean_detail_df(df):
    if df.empty:
        return pd.DataFrame()
    
    # 找真正的表头行
    header_idx = 0
    for i in range(min(5, len(df))):
        row_vals = [str(v) for v in df.iloc[i].values]
        if any("渠道" in v or "月份" in v for v in row_vals):
            header_idx = i
            break
    
    # 重建表头
    new_cols = []
    for i, v in enumerate(df.iloc[header_idx].values):
        s = str(v).strip() if v and str(v) != "None" else f"列{i}"
        new_cols.append(s)
    
    df = df.iloc[header_idx+1:].reset_index(drop=True)
    df.columns = new_cols
    
    # 填充合并单元格
    df = fill_merged(df, "月份")
    df = fill_merged(df, "地区")
    
    if "月份" in df.columns:
        df["月份"] = df["月份"].astype(str).str.strip().str.replace(" ", "")
    
    # 过滤合计行和无效行
    channel_col = None
    for col in df.columns:
        if "渠道" in col or "平台" in col:
            channel_col = col
            break
    
    if channel_col:
        df = df[~df[channel_col].astype(str).str.contains("合计", na=False)]
        df = df[df[channel_col].astype(str).str.strip().isin(["", "None", "nan"]) == False]
        df = df[df[channel_col].notna()].copy()
        df["渠道分类"] = df[channel_col].apply(classify_channel)
    
    # 过滤城市
    if "地区" in df.columns:
        df = df[df["地区"].isin(CITIES)].copy()
    
    # 数值转换（只处理已知数值列）
    for col in df.columns:
        if col in ["投放金额", "销售量", "收购量", "素材更新量", "直播场次"]:
            df[col] = to_num(df[col])
        elif "客资" in col and "成本" not in col and "直播" not in col and "视频" not in col:
            df["客资量"] = to_num(df[col])
    
    # 计算总成交量
    if "销售量" in df.columns and "收购量" in df.columns:
        df["总成交量"] = df["销售量"] + df["收购量"]
    
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
    df_main = clean_main_df(read_sheet(spreadsheet_token, SHEETS["四地汇总"], "A2:K200").copy())
    df_detail = clean_detail_df(read_sheet(spreadsheet_token, SHEETS["细分项汇总"], "A1:P500").copy())
    if df_detail is None:
        df_detail = pd.DataFrame()

# ── 侧边栏 ──
st.sidebar.header("🔍 筛选条件")
sel_city = st.sidebar.selectbox("筛选城市", ["全部城市"] + CITIES)
sel_month = st.sidebar.selectbox("筛选月份", ["全部月份"] + MONTHS)

df_filtered = apply_filter(df_main, sel_city, sel_month)
df_detail_filtered = apply_filter(df_detail, sel_city, sel_month)

st.sidebar.markdown(f"📋 四地汇总：**{len(df_filtered)}** 条")
st.sidebar.markdown(f"📋 渠道明细：**{len(df_detail_filtered)}** 条")

# ── 核心指标 ──
st.title("📊 新媒体数据看板")
st.caption(f"城市：{sel_city} ｜ 月份：{sel_month}")

total_spend = to_num(df_filtered["投放金额"]).sum() if "投放金额" in df_filtered.columns else 0
total_keizi = to_num(df_filtered["客资量"]).sum() if "客资量" in df_filtered.columns else 0
total_chengjiao = to_num(df_filtered["总成交量"]).sum() if "总成交量" in df_filtered.columns else 0
total_xiaoshou = to_num(df_filtered["销售量"]).sum() if "销售量" in df_filtered.columns else 0
total_shougou = to_num(df_filtered["收购量"]).sum() if "收购量" in df_filtered.columns else 0
keizi_cost = total_spend / total_keizi if total_keizi > 0 else 0
chengjiao_cost = total_spend / total_chengjiao if total_chengjiao > 0 else 0
chengjiao_rate = total_chengjiao / total_keizi * 100 if total_keizi > 0 else 0

c1,c2,c3,c4,c5 = st.columns(5)
c1.metric("总投放金额", f"¥{total_spend:,.0f}")
c2.metric("总客资量", f"{int(total_keizi):,}")
c3.metric("总成交量", f"{int(total_chengjiao):,}")
c4.metric("销售总量", f"{int(total_xiaoshou):,}")
c5.metric("收购总量", f"{int(total_shougou):,}")
c6,c7,c8 = st.columns(3)
c6.metric("客资成本", f"¥{keizi_cost:.2f}")
c7.metric("成交成本", f"¥{chengjiao_cost:.2f}")
c8.metric("成交率", f"{chengjiao_rate:.2f}%")

st.divider()

tab1, tab2, tab3, tab4 = st.tabs(["🏙️ 分城市数据", "📡 分渠道数据", "📈 趋势分析", "📋 数据明细"])

with tab1:
    st.subheader("分城市经营对比")
    df_city = apply_filter(df_main, "全部城市", sel_month)
    if not df_city.empty and "地区" in df_city.columns:
        cg = df_city.groupby("地区").agg(
            投放金额=("投放金额","sum"),
            客资量=("客资量","sum"),
            总成交量=("总成交量","sum"),
        ).reset_index()
        cg["客资成本"] = (cg["投放金额"]/cg["客资量"].replace(0,pd.NA)).round(2)
        cg["成交成本"] = (cg["投放金额"]/cg["总成交量"].replace(0,pd.NA)).round(2)
        cg["成交率%"] = (cg["总成交量"]/cg["客资量"].replace(0,pd.NA)*100).round(2)
        st.dataframe(cg, use_container_width=True, hide_index=True)
        ca,cb,cc,cd = st.columns(4)
        with ca: st.plotly_chart(px.bar(cg,x="地区",y="客资量",title="客资量",color="地区"),use_container_width=True)
        with cb: st.plotly_chart(px.bar(cg,x="地区",y="总成交量",title="成交量",color="地区"),use_container_width=True)
        with cc: st.plotly_chart(px.bar(cg,x="地区",y="客资成本",title="客资成本",color="地区"),use_container_width=True)
        with cd: st.plotly_chart(px.bar(cg,x="地区",y="成交成本",title="成交成本",color="地区"),use_container_width=True)

with tab2:
    st.subheader("分渠道数据对比")
    if not df_detail_filtered.empty and "渠道分类" in df_detail_filtered.columns:
        rg = df_detail_filtered.groupby("渠道分类").agg(
            投放金额=("投放金额","sum"),
            客资量=("客资量","sum"),
            总成交量=("总成交量","sum"),
        ).reset_index()
        rg["客资成本"] = (rg["投放金额"]/rg["客资量"].replace(0,pd.NA)).round(2)
        rg = rg.sort_values("客资量", ascending=False)
        st.dataframe(rg, use_container_width=True, hide_index=True)
        ra,rb = st.columns(2)
        with ra: st.plotly_chart(px.bar(rg,x="渠道分类",y="客资量",title="各渠道客资量",color="渠道分类"),use_container_width=True)
        with rb: st.plotly_chart(px.bar(rg,x="渠道分类",y="投放金额",title="各渠道投放金额",color="渠道分类"),use_container_width=True)
        rc,rd = st.columns(2)
        with rc: st.plotly_chart(px.bar(rg,x="渠道分类",y="总成交量",title="各渠道成交量",color="渠道分类"),use_container_width=True)
        with rd: st.plotly_chart(px.bar(rg,x="渠道分类",y="客资成本",title="各渠道客资成本",color="渠道分类"),use_container_width=True)
        re,rf = st.columns(2)
        with re: st.plotly_chart(px.pie(rg,names="渠道分类",values="客资量",title="各渠道客资量占比"),use_container_width=True)
        with rf: st.plotly_chart(px.pie(rg,names="渠道分类",values="投放金额",title="各渠道投放占比"),use_container_width=True)
    else:
        st.info("暂无渠道数据")

with tab3:
    st.subheader("趋势分析")
    df_trend = df_main.copy()
    if sel_city != "全部城市" and "地区" in df_trend.columns:
        df_trend = df_trend[df_trend["地区"] == sel_city]
    if not df_trend.empty and "月份" in df_trend.columns and "地区" in df_trend.columns:
        cm = df_trend.groupby(["月份","地区"]).agg(
            投放金额=("投放金额","sum"),
            客资量=("客资量","sum"),
            总成交量=("总成交量","sum"),
        ).reset_index()
        cm["客资成本"] = (cm["投放金额"]/cm["客资量"].replace(0,pd.NA)).round(2)
        cm["成交成本"] = (cm["投放金额"]/cm["总成交量"].replace(0,pd.NA)).round(2)
        cm["月份"] = pd.Categorical(cm["月份"], categories=MONTHS, ordered=True)
        cm = cm.sort_values("月份")
        st.plotly_chart(px.line(cm,x="月份",y="客资成本",color="地区",title="各城市客资成本月度趋势",markers=True),use_container_width=True)
        st.plotly_chart(px.line(cm,x="月份",y="成交成本",color="地区",title="各城市成交成本月度趋势",markers=True),use_container_width=True)
        tm = df_trend.groupby("月份").agg(
            客资量=("客资量","sum"),
            总成交量=("总成交量","sum"),
            投放金额=("投放金额","sum"),
        ).reset_index()
        tm["月份"] = pd.Categorical(tm["月份"], categories=MONTHS, ordered=True)
        tm = tm.sort_values("月份")
        st.plotly_chart(px.line(tm,x="月份",y=["客资量","总成交量"],title="月度客资与成交趋势",markers=True),use_container_width=True)

with tab4:
    st.subheader("数据明细")
    t1,t2 = st.tabs(["四地汇总","渠道明细"])
    with t1:
        st.dataframe(df_filtered.dropna(axis=1,how='all'), use_container_width=True)
    with t2:
        st.dataframe(df_detail_filtered.dropna(axis=1,how='all'), use_container_width=True)
