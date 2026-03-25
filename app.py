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

@st.cache_data(ttl=300)
def get_token():
    res = requests.post(
        "https://open.feishu.cn/open-apis/auth/v3/tenant_access_token/internal",
        json={"app_id": APP_ID, "app_secret": APP_SECRET}
    )
    return res.json().get("tenant_access_token")

@st.cache_data(ttl=300)
def get_spreadsheet_token():
    """通过Wiki API拿到电子表格的spreadsheetToken"""
    token = get_token()
    headers = {"Authorization": f"Bearer {token}"}
    res = requests.get(
        f"https://open.feishu.cn/open-apis/wiki/v2/spaces/get_node?token={WIKI_TOKEN}",
        headers=headers
    ).json()
    return res.get("data", {}).get("node", {}).get("obj_token")

@st.cache_data(ttl=300)
def read_sheet(spreadsheet_token, sheet_id, data_range="A1:Z500"):
    """读取指定sheet的数据"""
    token = get_token()
    headers = {"Authorization": f"Bearer {token}"}
    url = f"https://open.feishu.cn/open-apis/sheets/v2/spreadsheets/{spreadsheet_token}/values/{sheet_id}!{data_range}"
    res = requests.get(url, headers=headers).json()
    values = res.get("data", {}).get("valueRange", {}).get("values", [])
    if not values or len(values) < 2:
        return pd.DataFrame()
    headers_row = values[0]
    data_rows = values[1:]
    df = pd.DataFrame(data_rows, columns=headers_row)
    return df

def to_num(series):
    return pd.to_numeric(series, errors="coerce").fillna(0)

# ── 加载数据 ──
with st.spinner("正在连接飞书..."):
    try:
        spreadsheet_token = get_spreadsheet_token()
        if not spreadsheet_token:
            st.error("无法获取表格Token，请检查Wiki权限设置")
            st.stop()
        st.sidebar.success(f"✅ 连接成功")
    except Exception as e:
        st.error(f"连接失败：{e}")
        st.stop()

# 读取四地汇总表
@st.cache_data(ttl=300)
def load_main_data(spreadsheet_token):
    df = read_sheet(spreadsheet_token, SHEETS["四地汇总"], "A2:K200")
    return df

with st.spinner("正在加载数据..."):
    df_main = load_main_data(spreadsheet_token)

# ── 调试：显示原始数据 ──
st.subheader("📋 原始数据预览（调试用）")
st.write("列名：", list(df_main.columns))
st.dataframe(df_main.head(10))
```

---
