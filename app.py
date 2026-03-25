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
    token = get_token()
    headers = {"Authorization": f"Bearer {token}"}
    res = requests.get(
        f"https://open.feishu.cn/open-apis/wiki/v2/spaces/get_node?token={WIKI_TOKEN}",
        headers=headers
    ).json()
    st.sidebar.write("Wiki API响应：", res)
    return res.get("data", {}).get("node", {}).get("obj_token")

@st.cache_data(ttl=300)
def read_sheet(spreadsheet_token, sheet_id, data_range="A1:Z500"):
    token = get_token()
    headers = {"Authorization": f"Bearer {token}"}
    url = f"https://open.feishu.cn/open-apis/sheets/v2/spreadsheets/{spreadsheet_token}/values/{sheet_id}!{data_range}"
    res = requests.get(url, headers=headers).json()
    st.sidebar.write(f"Sheet {sheet_id} 响应：", res)
    values = res.get("data", {}).get("valueRange", {}).get("values", [])
    if not values or len(values) < 2:
        return pd.DataFrame()
    headers_row = [str(h) if h else f"列{i}" for i, h in enumerate(values[0])]
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
            st.error("无法获取表格Token，请检查Wiki权限")
            st.stop()
    except Exception as e:
        st.error(f"连接失败：{e}")
        st.stop()

with st.spinner("正在加载数据..."):
    df_main = read_sheet(spreadsheet_token, SHEETS["四地汇总"], "A2:K200")

st.title("📊 新媒体数据看板")
st.subheader("原始数据预览（调试）")
st.write("列名：", list(df_main.columns) if not df_main.empty else "空")
st.dataframe(df_main.head(20))
