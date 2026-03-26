import streamlit as st
import pandas as pd
import requests
import plotly.express as px
import plotly.graph_objects as go
import plotly.io as pio
from io import BytesIO
from reportlab.lib.pagesizes import A4
from reportlab.lib import colors
from reportlab.lib.units import cm
from reportlab.platypus import SimpleDocTemplate, Table, TableStyle, Paragraph, Spacer, Image, HRFlowable
from reportlab.lib.styles import getSampleStyleSheet, ParagraphStyle
from reportlab.pdfbase import pdfmetrics
from reportlab.pdfbase.ttfonts import TTFont
from reportlab.lib.enums import TA_CENTER, TA_LEFT
import datetime

st.set_page_config(page_title="新媒体数据看板", page_icon="📊", layout="wide")

st.markdown("""
<style>
    .stApp { background-color: #f8f9fc; }
    #MainMenu, footer, header { visibility: hidden; }
    [data-testid="metric-container"] {
        background: white;
        border: 1px solid #eef0f4;
        border-radius: 12px;
        padding: 20px 24px;
        box-shadow: 0 1px 4px rgba(0,0,0,0.05);
    }
    [data-testid="metric-container"] label { color: #6b7280 !important; font-size: 13px !important; font-weight: 500 !important; }
    [data-testid="metric-container"] [data-testid="stMetricValue"] { color: #111827 !important; font-size: 26px !important; font-weight: 700 !important; }
    .stTabs [data-baseweb="tab-list"] { gap: 8px; background: transparent; border-bottom: 1px solid #eef0f4; }
    .stTabs [data-baseweb="tab"] { background: transparent; color: #6b7280; font-weight: 500; border-radius: 6px 6px 0 0; padding: 8px 18px; }
    .stTabs [aria-selected="true"] { background: white !important; color: #111827 !important; border-bottom: 2px solid #4f46e5 !important; }
    [data-testid="stSidebar"] { background: white; border-right: 1px solid #eef0f4; }
    hr { border-color: #eef0f4; }
</style>
""", unsafe_allow_html=True)

APP_ID = st.secrets["FEISHU_APP_ID"]
APP_SECRET = st.secrets["FEISHU_APP_SECRET"]
WIKI_TOKEN = "G0VbwCGGMiEVNlktEYwcoA4JnHd"
SHEETS = {"四地汇总": "0HBHQk", "细分项汇总": "1hOmjY"}
MONTHS = ["1月","2月","3月","4月","5月","6月","7月","8月","9月","10月","11月","12月"]
CITIES = ["深圳", "上海", "成都", "天津"]
COLORS = ["#4f46e5","#06b6d4","#10b981","#f59e0b","#ef4444","#8b5cf6"]
PLOTLY_LAYOUT = dict(
    paper_bgcolor="white", plot_bgcolor="white",
    font=dict(family="sans-serif", size=12, color="#374151"),
    margin=dict(l=16, r=16, t=40, b=16),
    legend=dict(orientation="h", yanchor="bottom", y=1.02, xanchor="right", x=1),
    xaxis=dict(showgrid=False, linecolor="#eef0f4"),
    yaxis=dict(gridcolor="#f3f4f6", linecolor="#eef0f4"),
)

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
    return pd.DataFrame(values[1:], columns=headers_row)

def to_num(series):
    if isinstance(series, pd.Series):
        s = series.astype(str)
        s = s.str.replace(r'IFERROR\(.*?\)', '0', regex=True)
        s = s.str.replace("/", "0").str.replace(",", "")
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
    if df.empty: return df
    df = fill_merged(df, "月份")
    if "月份" in df.columns:
        df["月份"] = df["月份"].astype(str).str.strip().str.replace(" ", "")
    if "地区" in df.columns:
        df["地区"] = df["地区"].replace([None,"None","","nan"], pd.NA)
        df = df[df["地区"].isin(CITIES)].copy()
    for col in ["投放金额","总成交量","销售量","收购量"]:
        if col in df.columns:
            df[col] = to_num(df[col])
    for col in df.columns:
        if "客资" in col and "成本" not in col:
            df["客资量"] = to_num(df[col])
            break
    if "投放金额" in df.columns and "客资量" in df.columns:
        df["综合客资成本"] = (df["投放金额"] / df["客资量"].replace(0, pd.NA)).round(2)
    if "投放金额" in df.columns and "总成交量" in df.columns:
        df["综合成交成本"] = (df["投放金额"] / df["总成交量"].replace(0, pd.NA)).round(2)
    return df

def clean_detail_df(df):
    if df.empty: return pd.DataFrame()
    header_idx = 0
    for i in range(min(5, len(df))):
        row_vals = [str(v) for v in df.iloc[i].values]
        if any("渠道" in v or "月份" in v for v in row_vals):
            header_idx = i
            break
    new_cols = [str(v).strip() if v and str(v) != "None" else f"列{i}"
                for i, v in enumerate(df.iloc[header_idx].values)]
    df = df.iloc[header_idx+1:].reset_index(drop=True)
    df.columns = new_cols
    df = fill_merged(df, "月份")
    df = fill_merged(df, "地区")
    if "月份" in df.columns:
        df["月份"] = df["月份"].astype(str).str.strip().str.replace(" ", "")
    channel_col = next((c for c in df.columns if "渠道" in c or "平台" in c), None)
    if channel_col:
        df = df[~df[channel_col].astype(str).str.contains("合计", na=False)]
        df = df[df[channel_col].astype(str).str.strip().isin(["","None","nan"]) == False]
        df = df[df[channel_col].notna()].copy()
        df["渠道分类"] = df[channel_col].apply(classify_channel)
    if "地区" in df.columns:
        df = df[df["地区"].isin(CITIES)].copy()
    for col in df.columns:
        if col in ["投放金额","销售量","收购量","素材更新量","直播场次"]:
            df[col] = to_num(df[col])
        elif "客资" in col and "成本" not in col and "直播" not in col and "视频" not in col:
            df["客资量"] = to_num(df[col])
    if "销售量" in df.columns and "收购量" in df.columns:
        df["总成交量"] = df["销售量"] + df["收购量"]
    return df

def make_chart(fig):
    fig.update_layout(**PLOTLY_LAYOUT)
    return fig

def fig_to_image(fig, width=700, height=350):
    """把plotly图表转成图片bytes"""
    fig.update_layout(**PLOTLY_LAYOUT)
    img_bytes = pio.to_image(fig, format="png", width=width, height=height, scale=2)
    return BytesIO(img_bytes)

def generate_pdf(sel_city, sel_month, metrics, cg, rg, cm_df, ch_month_df):
    """生成PDF报告"""
    buffer = BytesIO()
    doc = SimpleDocTemplate(
        buffer, pagesize=A4,
        leftMargin=1.5*cm, rightMargin=1.5*cm,
        topMargin=1.5*cm, bottomMargin=1.5*cm
    )

    styles = getSampleStyleSheet()
    title_style = ParagraphStyle('title', fontSize=18, fontName='Helvetica-Bold',
                                  textColor=colors.HexColor('#111827'), spaceAfter=4)
    subtitle_style = ParagraphStyle('subtitle', fontSize=10, fontName='Helvetica',
                                     textColor=colors.HexColor('#6b7280'), spaceAfter=16)
    section_style = ParagraphStyle('section', fontSize=13, fontName='Helvetica-Bold',
                                    textColor=colors.HexColor('#111827'), spaceBefore=16, spaceAfter=8)

    story = []
    page_width = A4[0] - 3*cm

    # ── 标题 ──
    story.append(Paragraph("新媒体数据看板报告", title_style))
    story.append(Paragraph(
        f"城市：{sel_city} · 月份：{sel_month} · 生成时间：{datetime.datetime.now().strftime('%Y-%m-%d %H:%M')}",
        subtitle_style
    ))
    story.append(HRFlowable(width="100%", thickness=1, color=colors.HexColor('#eef0f4')))
    story.append(Spacer(1, 0.4*cm))

    # ── 核心指标 ──
    story.append(Paragraph("核心指标总览", section_style))
    metric_data = [
        ["总投放金额", "总客资量", "总成交量", "成交率"],
        [metrics['total_spend'], metrics['total_keizi'], metrics['total_chengjiao'], metrics['chengjiao_rate']],
        ["销售总量", "收购总量", "客资成本", "成交成本"],
        [metrics['total_xiaoshou'], metrics['total_shougou'], metrics['keizi_cost'], metrics['chengjiao_cost']],
    ]
    metric_table = Table(metric_data, colWidths=[page_width/4]*4)
    metric_table.setStyle(TableStyle([
        ('BACKGROUND', (0,0), (-1,0), colors.HexColor('#f3f4f6')),
        ('BACKGROUND', (0,2), (-1,2), colors.HexColor('#f3f4f6')),
        ('TEXTCOLOR', (0,0), (-1,0), colors.HexColor('#6b7280')),
        ('TEXTCOLOR', (0,2), (-1,2), colors.HexColor('#6b7280')),
        ('FONTNAME', (0,0), (-1,-1), 'Helvetica'),
        ('FONTNAME', (0,1), (-1,1), 'Helvetica-Bold'),
        ('FONTNAME', (0,3), (-1,3), 'Helvetica-Bold'),
        ('FONTSIZE', (0,0), (-1,0), 9),
        ('FONTSIZE', (0,2), (-1,2), 9),
        ('FONTSIZE', (0,1), (-1,1), 14),
        ('FONTSIZE', (0,3), (-1,3), 14),
        ('ALIGN', (0,0), (-1,-1), 'CENTER'),
        ('VALIGN', (0,0), (-1,-1), 'MIDDLE'),
        ('ROWBACKGROUND', (0,0), (-1,-1), [colors.HexColor('#f9fafb'), colors.white]),
        ('BOX', (0,0), (-1,-1), 0.5, colors.HexColor('#eef0f4')),
        ('INNERGRID', (0,0), (-1,-1), 0.5, colors.HexColor('#eef0f4')),
        ('ROWHEIGHT', (0,0), (-1,-1), 28),
    ]))
    story.append(metric_table)
    story.append(Spacer(1, 0.4*cm))

    # ── 分城市表格 ──
    if cg is not None and not cg.empty:
        story.append(HRFlowable(width="100%", thickness=1, color=colors.HexColor('#eef0f4')))
        story.append(Paragraph("分城市经营对比", section_style))
        city_cols = ["地区","投放金额","客资量","总成交量","客资成本","成交成本","成交率%"]
        city_cols = [c for c in city_cols if c in cg.columns]
        city_data = [city_cols] + cg[city_cols].values.tolist()
        city_table = Table(city_data, colWidths=[page_width/len(city_cols)]*len(city_cols))
        city_table.setStyle(TableStyle([
            ('BACKGROUND', (0,0), (-1,0), colors.HexColor('#4f46e5')),
            ('TEXTCOLOR', (0,0), (-1,0), colors.white),
            ('FONTNAME', (0,0), (-1,0), 'Helvetica-Bold'),
            ('FONTNAME', (0,1), (-1,-1), 'Helvetica'),
            ('FONTSIZE', (0,0), (-1,-1), 9),
            ('ALIGN', (0,0), (-1,-1), 'CENTER'),
            ('VALIGN', (0,0), (-1,-1), 'MIDDLE'),
            ('ROWBACKGROUND', (0,1), (-1,-1), [colors.white, colors.HexColor('#f9fafb')]),
            ('BOX', (0,0), (-1,-1), 0.5, colors.HexColor('#eef0f4')),
            ('INNERGRID', (0,0), (-1,-1), 0.5, colors.HexColor('#eef0f4')),
            ('ROWHEIGHT', (0,0), (-1,-1), 24),
        ]))
        story.append(city_table)

        # 城市对比图
        if not cg.empty:
            fig_city = px.bar(cg, x="地区", y=["客资量","总成交量"],
                barmode="group", title="各城市客资量与成交量对比",
                color_discrete_sequence=COLORS)
            story.append(Spacer(1, 0.3*cm))
            story.append(Image(fig_to_image(fig_city, 700, 300), width=page_width, height=page_width*300/700))

    # ── 分渠道表格 ──
    if rg is not None and not rg.empty:
        story.append(HRFlowable(width="100%", thickness=1, color=colors.HexColor('#eef0f4')))
        story.append(Paragraph("分渠道数据对比", section_style))
        ch_cols = ["渠道分类","投放金额","客资量","总成交量","客资成本"]
        ch_cols = [c for c in ch_cols if c in rg.columns]
        ch_data = [ch_cols] + rg[ch_cols].values.tolist()
        ch_table = Table(ch_data, colWidths=[page_width/len(ch_cols)]*len(ch_cols))
        ch_table.setStyle(TableStyle([
            ('BACKGROUND', (0,0), (-1,0), colors.HexColor('#4f46e5')),
            ('TEXTCOLOR', (0,0), (-1,0), colors.white),
            ('FONTNAME', (0,0), (-1,0), 'Helvetica-Bold'),
            ('FONTNAME', (0,1), (-1,-1), 'Helvetica'),
            ('FONTSIZE', (0,0), (-1,-1), 9),
            ('ALIGN', (0,0), (-1,-1), 'CENTER'),
            ('VALIGN', (0,0), (-1,-1), 'MIDDLE'),
            ('ROWBACKGROUND', (0,1), (-1,-1), [colors.white, colors.HexColor('#f9fafb')]),
            ('BOX', (0,0), (-1,-1), 0.5, colors.HexColor('#eef0f4')),
            ('INNERGRID', (0,0), (-1,-1), 0.5, colors.HexColor('#eef0f4')),
            ('ROWHEIGHT', (0,0), (-1,-1), 24),
        ]))
        story.append(ch_table)

        # 渠道对比图
        fig_ch = px.bar(rg, x="渠道分类", y="客资量",
            title="各渠道客资量对比", color="渠道分类", color_discrete_sequence=COLORS)
        story.append(Spacer(1, 0.3*cm))
        story.append(Image(fig_to_image(fig_ch, 700, 300), width=page_width, height=page_width*300/700))

    # ── 趋势图 ──
    if cm_df is not None and not cm_df.empty:
        story.append(HRFlowable(width="100%", thickness=1, color=colors.HexColor('#eef0f4')))
        story.append(Paragraph("趋势分析", section_style))
        fig_trend1 = px.line(cm_df, x="月份", y="客资成本", color="地区",
            title="各城市客资成本月度趋势", markers=True, color_discrete_sequence=COLORS)
        story.append(Image(fig_to_image(fig_trend1, 700, 300), width=page_width, height=page_width*300/700))
        story.append(Spacer(1, 0.3*cm))
        fig_trend2 = px.line(cm_df, x="月份", y="成交成本", color="地区",
            title="各城市成交成本月度趋势", markers=True, color_discrete_sequence=COLORS)
        story.append(Image(fig_to_image(fig_trend2, 700, 300), width=page_width, height=page_width*300/700))

    if ch_month_df is not None and not ch_month_df.empty:
        story.append(Spacer(1, 0.3*cm))
        fig_ch_trend = px.line(ch_month_df, x="月份", y="客资量", color="渠道分类",
            title="各渠道客资量月度趋势", markers=True, color_discrete_sequence=COLORS)
        story.append(Image(fig_to_image(fig_ch_trend, 700, 300), width=page_width, height=page_width*300/700))

    doc.build(story)
    buffer.seek(0)
    return buffer.getvalue()

# ── 加载数据 ──
with st.spinner("正在加载数据..."):
    try:
        spreadsheet_token = get_spreadsheet_token()
        if not spreadsheet_token:
            st.error("无法获取表格Token")
            st.stop()
        df_main = clean_main_df(read_sheet(spreadsheet_token, SHEETS["四地汇总"], "A2:K200").copy())
        df_detail = clean_detail_df(read_sheet(spreadsheet_token, SHEETS["细分项汇总"], "A1:P500").copy())
        if df_detail is None: df_detail = pd.DataFrame()
    except Exception as e:
        st.error(f"加载失败：{e}")
        st.stop()

# ── 侧边栏 ──
with st.sidebar:
    st.markdown("## 筛选条件")
    sel_city = st.selectbox("城市", ["全部城市"] + CITIES)
    sel_month = st.selectbox("月份", ["全部月份"] + MONTHS)
    st.divider()

    df_filtered = apply_filter(df_main, sel_city, sel_month)
    df_detail_filtered = apply_filter(df_detail, sel_city, sel_month)

    st.caption(f"四地汇总：{len(df_filtered)} 条")
    st.caption(f"渠道明细：{len(df_detail_filtered)} 条")
    st.divider()

    st.markdown("## 导出报告")
    if st.button("生成 PDF 报告", use_container_width=True):
        with st.spinner("正在生成报告..."):
            # 准备数据
            total_spend = to_num(df_filtered["投放金额"]).sum() if "投放金额" in df_filtered.columns else 0
            total_keizi = to_num(df_filtered["客资量"]).sum() if "客资量" in df_filtered.columns else 0
            total_chengjiao = to_num(df_filtered["总成交量"]).sum() if "总成交量" in df_filtered.columns else 0
            total_xiaoshou = to_num(df_filtered["销售量"]).sum() if "销售量" in df_filtered.columns else 0
            total_shougou = to_num(df_filtered["收购量"]).sum() if "收购量" in df_filtered.columns else 0
            keizi_cost = total_spend / total_keizi if total_keizi > 0 else 0
            chengjiao_cost = total_spend / total_chengjiao if total_chengjiao > 0 else 0
            chengjiao_rate = total_chengjiao / total_keizi * 100 if total_keizi > 0 else 0

            metrics = {
                'total_spend': f"¥{total_spend:,.0f}",
                'total_keizi': f"{int(total_keizi):,}",
                'total_chengjiao': f"{int(total_chengjiao):,}",
                'chengjiao_rate': f"{chengjiao_rate:.2f}%",
                'total_xiaoshou': f"{int(total_xiaoshou):,}",
                'total_shougou': f"{int(total_shougou):,}",
                'keizi_cost': f"¥{keizi_cost:.2f}",
                'chengjiao_cost': f"¥{chengjiao_cost:.2f}",
            }

            # 城市对比
            df_city = apply_filter(df_main, "全部城市", sel_month)
            cg = None
            if not df_city.empty and "地区" in df_city.columns:
                cg = df_city.groupby("地区").agg(投放金额=("投放金额","sum"), 客资量=("客资量","sum"), 总成交量=("总成交量","sum")).reset_index()
                cg["客资成本"] = (cg["投放金额"]/cg["客资量"].replace(0,pd.NA)).round(2)
                cg["成交成本"] = (cg["投放金额"]/cg["总成交量"].replace(0,pd.NA)).round(2)
                cg["成交率%"] = (cg["总成交量"]/cg["客资量"].replace(0,pd.NA)*100).round(2)

            # 渠道对比
            rg = None
            if not df_detail_filtered.empty and "渠道分类" in df_detail_filtered.columns:
                rg = df_detail_filtered.groupby("渠道分类").agg(投放金额=("投放金额","sum"), 客资量=("客资量","sum"), 总成交量=("总成交量","sum")).reset_index()
                rg["客资成本"] = (rg["投放金额"]/rg["客资量"].replace(0,pd.NA)).round(2)
                rg = rg.sort_values("客资量", ascending=False)

            # 趋势
            cm_df = None
            if not df_main.empty and "月份" in df_main.columns:
                df_trend = df_main.copy()
                if sel_city != "全部城市" and "地区" in df_trend.columns:
                    df_trend = df_trend[df_trend["地区"] == sel_city]
                cm_df = df_trend.groupby(["月份","地区"]).agg(投放金额=("投放金额","sum"), 客资量=("客资量","sum"), 总成交量=("总成交量","sum")).reset_index()
                cm_df["客资成本"] = (cm_df["投放金额"]/cm_df["客资量"].replace(0,pd.NA)).round(2)
                cm_df["成交成本"] = (cm_df["投放金额"]/cm_df["总成交量"].replace(0,pd.NA)).round(2)
                cm_df["月份"] = pd.Categorical(cm_df["月份"], categories=MONTHS, ordered=True)
                cm_df = cm_df.sort_values("月份")

            ch_month_df = None
            if not df_detail.empty and "月份" in df_detail.columns and "渠道分类" in df_detail.columns:
                df_ch = apply_filter(df_detail, sel_city, "全部月份")
                ch_month_df = df_ch.groupby(["月份","渠道分类"]).agg(客资量=("客资量","sum")).reset_index()
                ch_month_df["月份"] = pd.Categorical(ch_month_df["月份"], categories=MONTHS, ordered=True)
                ch_month_df = ch_month_df.sort_values("月份")

            try:
                pdf_bytes = generate_pdf(sel_city, sel_month, metrics, cg, rg, cm_df, ch_month_df)
                st.download_button(
                    label="⬇️ 下载 PDF 报告",
                    data=pdf_bytes,
                    file_name=f"新媒体看板报告_{sel_city}_{sel_month}_{datetime.datetime.now().strftime('%Y%m%d')}.pdf",
                    mime="application/pdf",
                    use_container_width=True
                )
            except Exception as e:
                st.error(f"PDF生成失败：{e}")

# ── 主页面 ──
st.markdown(f"""
<div style="padding:8px 0 20px 0;border-bottom:1px solid #eef0f4;margin-bottom:24px;">
    <h2 style="margin:0;color:#111827;font-weight:700;">📊 新媒体数据看板</h2>
    <p style="margin:4px 0 0 0;color:#6b7280;font-size:14px;">{sel_city} · {sel_month} · 数据每60秒自动更新</p>
</div>
""", unsafe_allow_html=True)

total_spend = to_num(df_filtered["投放金额"]).sum() if "投放金额" in df_filtered.columns else 0
total_keizi = to_num(df_filtered["客资量"]).sum() if "客资量" in df_filtered.columns else 0
total_chengjiao = to_num(df_filtered["总成交量"]).sum() if "总成交量" in df_filtered.columns else 0
total_xiaoshou = to_num(df_filtered["销售量"]).sum() if "销售量" in df_filtered.columns else 0
total_shougou = to_num(df_filtered["收购量"]).sum() if "收购量" in df_filtered.columns else 0
keizi_cost = total_spend / total_keizi if total_keizi > 0 else 0
chengjiao_cost = total_spend / total_chengjiao if total_chengjiao > 0 else 0
chengjiao_rate = total_chengjiao / total_keizi * 100 if total_keizi > 0 else 0

c1,c2,c3,c4 = st.columns(4)
c1.metric("总投放金额", f"¥{total_spend:,.0f}")
c2.metric("总客资量", f"{int(total_keizi):,}")
c3.metric("总成交量", f"{int(total_chengjiao):,}")
c4.metric("成交率", f"{chengjiao_rate:.2f}%")
c5,c6,c7,c8 = st.columns(4)
c5.metric("销售总量", f"{int(total_xiaoshou):,}")
c6.metric("收购总量", f"{int(total_shougou):,}")
c7.metric("客资成本", f"¥{keizi_cost:.2f}")
c8.metric("成交成本", f"¥{chengjiao_cost:.2f}")

st.divider()

tab1, tab2, tab3, tab4 = st.tabs(["🏙️ 分城市", "📡 分渠道", "📈 趋势分析", "📋 数据明细"])

with tab1:
    st.subheader("分城市经营对比")
    df_city = apply_filter(df_main, "全部城市", sel_month)
    if not df_city.empty and "地区" in df_city.columns:
        cg = df_city.groupby("地区").agg(投放金额=("投放金额","sum"), 客资量=("客资量","sum"), 总成交量=("总成交量","sum")).reset_index()
        cg["客资成本"] = (cg["投放金额"]/cg["客资量"].replace(0,pd.NA)).round(2)
        cg["成交成本"] = (cg["投放金额"]/cg["总成交量"].replace(0,pd.NA)).round(2)
        cg["成交率%"] = (cg["总成交量"]/cg["客资量"].replace(0,pd.NA)*100).round(2)
        st.dataframe(cg, use_container_width=True, hide_index=True)
        ca,cb,cc,cd = st.columns(4)
        with ca:
            fig = px.bar(cg,x="地区",y="客资量",title="客资量",color="地区",color_discrete_sequence=COLORS)
            st.plotly_chart(make_chart(fig),use_container_width=True)
        with cb:
            fig = px.bar(cg,x="地区",y="总成交量",title="成交量",color="地区",color_discrete_sequence=COLORS)
            st.plotly_chart(make_chart(fig),use_container_width=True)
        with cc:
            fig = px.bar(cg,x="地区",y="客资成本",title="客资成本",color="地区",color_discrete_sequence=COLORS)
            st.plotly_chart(make_chart(fig),use_container_width=True)
        with cd:
            fig = px.bar(cg,x="地区",y="成交成本",title="成交成本",color="地区",color_discrete_sequence=COLORS)
            st.plotly_chart(make_chart(fig),use_container_width=True)

with tab2:
    st.subheader("分渠道数据对比")
    if not df_detail_filtered.empty and "渠道分类" in df_detail_filtered.columns:
        rg = df_detail_filtered.groupby("渠道分类").agg(投放金额=("投放金额","sum"), 客资量=("客资量","sum"), 总成交量=("总成交量","sum")).reset_index()
        rg["客资成本"] = (rg["投放金额"]/rg["客资量"].replace(0,pd.NA)).round(2)
        rg = rg.sort_values("客资量", ascending=False)
        st.dataframe(rg, use_container_width=True, hide_index=True)
        ra,rb = st.columns(2)
        with ra:
            fig = px.bar(rg,x="渠道分类",y="客资量",title="各渠道客资量",color="渠道分类",color_discrete_sequence=COLORS)
            st.plotly_chart(make_chart(fig),use_container_width=True)
        with rb:
            fig = px.bar(rg,x="渠道分类",y="投放金额",title="各渠道投放金额",color="渠道分类",color_discrete_sequence=COLORS)
            st.plotly_chart(make_chart(fig),use_container_width=True)
        rc,rd = st.columns(2)
        with rc:
            fig = px.bar(rg,x="渠道分类",y="总成交量",title="各渠道成交量",color="渠道分类",color_discrete_sequence=COLORS)
            st.plotly_chart(make_chart(fig),use_container_width=True)
        with rd:
            fig = px.bar(rg,x="渠道分类",y="客资成本",title="各渠道客资成本",color="渠道分类",color_discrete_sequence=COLORS)
            st.plotly_chart(make_chart(fig),use_container_width=True)
        re,rf = st.columns(2)
        with re:
            fig = px.pie(rg,names="渠道分类",values="客资量",title="各渠道客资量占比",color_discrete_sequence=COLORS)
            fig.update_traces(textposition='inside', textinfo='percent+label')
            st.plotly_chart(make_chart(fig),use_container_width=True)
        with rf:
            fig = px.pie(rg,names="渠道分类",values="投放金额",title="各渠道投放占比",color_discrete_sequence=COLORS)
            fig.update_traces(textposition='inside', textinfo='percent+label')
            st.plotly_chart(make_chart(fig),use_container_width=True)
        st.divider()
        st.subheader("渠道月度趋势")
        df_ch_trend = apply_filter(df_detail, sel_city, "全部月份")
        if not df_ch_trend.empty and "渠道分类" in df_ch_trend.columns:
            ch_month = df_ch_trend.groupby(["月份","渠道分类"]).agg(客资量=("客资量","sum"), 投放金额=("投放金额","sum")).reset_index()
            ch_month["月份"] = pd.Categorical(ch_month["月份"], categories=MONTHS, ordered=True)
            ch_month = ch_month.sort_values("月份")
            fig = px.line(ch_month,x="月份",y="客资量",color="渠道分类",title="各渠道客资量月度趋势",markers=True,color_discrete_sequence=COLORS)
            st.plotly_chart(make_chart(fig),use_container_width=True)
            fig2 = px.line(ch_month,x="月份",y="投放金额",color="渠道分类",title="各渠道投放金额月度趋势",markers=True,color_discrete_sequence=COLORS)
            st.plotly_chart(make_chart(fig2),use_container_width=True)
    else:
        st.info("暂无渠道数据")

with tab3:
    st.subheader("趋势分析")
    df_trend = df_main.copy()
    if sel_city != "全部城市" and "地区" in df_trend.columns:
        df_trend = df_trend[df_trend["地区"] == sel_city]
    if not df_trend.empty and "月份" in df_trend.columns:
        cm_df = df_trend.groupby(["月份","地区"]).agg(投放金额=("投放金额","sum"), 客资量=("客资量","sum"), 总成交量=("总成交量","sum")).reset_index()
        cm_df["客资成本"] = (cm_df["投放金额"]/cm_df["客资量"].replace(0,pd.NA)).round(2)
        cm_df["成交成本"] = (cm_df["投放金额"]/cm_df["总成交量"].replace(0,pd.NA)).round(2)
        cm_df["月份"] = pd.Categorical(cm_df["月份"], categories=MONTHS, ordered=True)
        cm_df = cm_df.sort_values("月份")
        fig = px.line(cm_df,x="月份",y="客资成本",color="地区",title="各城市客资成本月度趋势",markers=True,color_discrete_sequence=COLORS)
        st.plotly_chart(make_chart(fig),use_container_width=True)
        fig2 = px.line(cm_df,x="月份",y="成交成本",color="地区",title="各城市成交成本月度趋势",markers=True,color_discrete_sequence=COLORS)
        st.plotly_chart(make_chart(fig2),use_container_width=True)
        tm = df_trend.groupby("月份").agg(客资量=("客资量","sum"), 总成交量=("总成交量","sum"), 投放金额=("投放金额","sum")).reset_index()
        tm["月份"] = pd.Categorical(tm["月份"], categories=MONTHS, ordered=True)
        tm = tm.sort_values("月份")
        fig3 = px.line(tm,x="月份",y=["客资量","总成交量"],title="月度客资与成交趋势",markers=True,color_discrete_sequence=COLORS)
        st.plotly_chart(make_chart(fig3),use_container_width=True)

with tab4:
    st.subheader("数据明细")
    t1,t2 = st.tabs(["四地汇总","渠道明细"])
    with t1:
        st.dataframe(df_filtered.dropna(axis=1,how='all'), use_container_width=True)
    with t2:
        st.dataframe(df_detail_filtered.dropna(axis=1,how='all'), use_container_width=True)
