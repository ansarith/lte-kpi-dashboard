import streamlit as st
import pandas as pd
import plotly.express as px
import plotly.graph_objects as go
from pptx import Presentation
from pptx.util import Inches
import io
from PIL import Image
import tempfile
import os

st.title("📊 LTE KPI Dashboard (Kaleido-free PPT export)")

RUNNING_ON_CLOUD = os.environ.get("STREAMLIT_SERVER_PORT") is not None

# ------------------ Load Data ------------------
@st.cache_data
def load_data(path):
    df = pd.read_excel(path)
    df["Period start time"] = pd.to_datetime(df["Period start time"], errors="coerce")
    return df

DATA_PATH = "4G_Main_KPIs_Report_SRAN21B-Sarith-2025_10_13-Site KCH2567RBR & 2070_LB.xlsx"
df = load_data(DATA_PATH)

# ------------------ Filters ------------------
kpi_columns = [c for c in df.columns if c not in ["Period start time","LNBTS name","LNCEL name"]]
selected_kpis = st.multiselect("Select KPI(s)", options=kpi_columns, default=kpi_columns[:4])
enodeb_selected = st.multiselect("Select LNBTS name", sorted(df["LNBTS name"].unique()))
cell_selected = st.multiselect(
    "Select LNCEL name",
    sorted(df[df["LNBTS name"].isin(enodeb_selected)]["LNCEL name"].unique()) if enodeb_selected else sorted(df["LNCEL name"].unique())
)
daily_option = st.checkbox("📅 Daily Aggregation")
group_option = st.checkbox("🏙️ Group by Site")

# ------------------ Aggregate Data ------------------
plot_df = df.copy()
if enodeb_selected:
    plot_df = plot_df[plot_df["LNBTS name"].isin(enodeb_selected)]
if cell_selected:
    plot_df = plot_df[plot_df["LNCEL name"].isin(cell_selected)]

def aggregate_data(df, kpis, daily=False, group=False):
    for kpi in kpis:
        df[kpi] = pd.to_numeric(df[kpi], errors="coerce")
    agg_dict = {kpi: "sum" if "Volume" in kpi or "RRC" in kpi else "mean" for kpi in kpis}
    time_col = "Date" if daily else "Period start time"
    if daily:
        df["Date"] = df["Period start time"].dt.normalize()
    group_cols = [time_col] + ([] if group else ["LNCEL name"] if "LNCEL name" in df.columns else [])
    grouped = df.groupby(group_cols, as_index=False).agg(agg_dict)
    grouped["Time_str"] = grouped[time_col].dt.strftime("%Y-%m-%d" if daily else "%Y-%m-%d %H:%M")
    return grouped

plot_df = aggregate_data(plot_df, selected_kpis, daily_option, group_option)

# ------------------ Create Figures ------------------
figures = []
colors = px.colors.qualitative.Dark24
cols = st.columns(2)

if not plot_df.empty:
    for idx, kpi in enumerate(selected_kpis[:4]):
        fig = go.Figure()
        if not group_option and "LNCEL name" in plot_df.columns:
            for i, cell in enumerate(plot_df["LNCEL name"].unique()):
                cell_df = plot_df[plot_df["LNCEL name"] == cell]
                fig.add_trace(go.Scatter(x=cell_df["Time_str"], y=cell_df[kpi],
                                         mode="lines+markers", name=cell,
                                         line=dict(color=colors[i % len(colors)])))
        else:
            fig.add_trace(go.Scatter(x=plot_df["Time_str"], y=plot_df[kpi],
                                     mode="lines+markers", name=kpi))
        fig.update_layout(height=420, title=dict(text=kpi, x=0.5),
                          hovermode="x unified", legend=dict(x=1.02,y=1),
                          margin=dict(l=40,r=120,t=60,b=40))
        cols[idx % 2].plotly_chart(fig, use_container_width=True)
        figures.append(fig)

# ------------------ PPT Export (Kaleido-free) ------------------
import plotly.graph_objects as go
from pptx import Presentation
from pptx.util import Inches
import io
import asyncio
from pyppeteer import launch
from tempfile import NamedTemporaryFile

async def fig_to_png(fig):
    """Convert Plotly figure to PNG via headless browser."""
    with NamedTemporaryFile(suffix=".html", delete=False) as f:
        fig.write_html(f.name)
        browser = await launch(headless=True, args=['--no-sandbox'])
        page = await browser.newPage()
        await page.goto(f"file://{f.name}")
        png_bytes = await page.screenshot()
        await browser.close()
    return png_bytes

def create_ppt_with_screenshots(figures):
    prs = Presentation()
    prs.slide_width = Inches(13.33)
    prs.slide_height = Inches(7.5)
    positions = [(Inches(0.5),Inches(0.5)), (Inches(6.9),Inches(0.5)),
                 (Inches(0.5),Inches(4.0)), (Inches(6.9),Inches(4.0))]
    chart_width = Inches(6.08)
    chart_height = Inches(3.04)

    for idx, fig in enumerate(figures):
        if idx % 4 == 0:
            slide = prs.slides.add_slide(prs.slide_layouts[5])
        png_bytes = asyncio.get_event_loop().run_until_complete(fig_to_png(fig))
        img_buf = io.BytesIO(png_bytes)
        slide.shapes.add_picture(img_buf, positions[idx%4][0], positions[idx%4][1],
                                 width=chart_width, height=chart_height)

    ppt_buffer = io.BytesIO()
    prs.save(ppt_buffer)
    ppt_buffer.seek(0)
    return ppt_buffer
if figures:
    ppt_file = create_ppt(figures)
    if ppt_file:
        st.download_button("📊 Download PowerPoint Report", ppt_file,
                           "LTE_KPI_Report.pptx",
                           "application/vnd.openxmlformats-officedocument.presentationml.presentation")
else:
    st.warning("⚠️ No data available for the selected filters.")

