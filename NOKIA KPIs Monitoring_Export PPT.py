import os, sys, io
import streamlit as st
import pandas as pd
import plotly.graph_objects as go
import plotly.express as px
from pptx import Presentation
from pptx.util import Inches
import matplotlib.pyplot as plt

st.set_page_config(layout="wide")

# ----------------------------------------------------
st.write("🚀 Running file:", os.path.abspath(__file__))
st.write("🟢 Python executable:", sys.executable)
# ----------------------------------------------------

# ---------------- LOAD DATA ----------------
@st.cache_data
def load_data(path):
    df = pd.read_excel(path)
    df["Period start time"] = pd.to_datetime(df["Period start time"], errors="coerce")

    percentage_kpis = [col for col in df.columns if "%" in col or "Rate" in col]
    for col in percentage_kpis:
        if col in df.columns and pd.api.types.is_numeric_dtype(df[col]):
            if df[col].max() <= 1.0:
                df[col] = df[col] * 100

    return df

DATA_PATH = "4G_Main_KPIs_Report_SRAN21B-Sarith-2025_10_13-Site KCH2567RBR & 2070_LB.xlsx"
df = load_data(DATA_PATH)

st.title("📊 LTE KPI Dashboard")

# ---------------- KPI SELECTION ----------------
kpi_columns = [col for col in df.columns if col not in ["Period start time","LNBTS name","LNCEL name"]]
selected_kpis = st.multiselect("Select KPI(s)", options=kpi_columns, default=kpi_columns[:4])

# ---------------- SITE FILTER ----------------
enodeb_selected = st.multiselect("Select LNBTS name", options=sorted(df["LNBTS name"].unique()))

# ---------------- CELL FILTER ----------------
if enodeb_selected:
    cell_options = sorted(df[df["LNBTS name"].isin(enodeb_selected)]["LNCEL name"].unique())
else:
    cell_options = sorted(df["LNCEL name"].unique())

cell_selected = st.multiselect("Select LNCEL name", options=cell_options)

# ---------------- OPTIONS ----------------
daily_option = st.checkbox("📅 Daily Aggregation")
group_option = st.checkbox("🏙️ Group by Site")

# ---------------- FILTER DATAFRAME ----------------
plot_df = df.copy()

if enodeb_selected:
    plot_df = plot_df[plot_df["LNBTS name"].isin(enodeb_selected)]

if cell_selected:
    plot_df = plot_df[plot_df["LNCEL name"].isin(cell_selected)]

# ---------------- AGGREGATION ----------------
def aggregate_data(df, kpis, daily=False, group=False):

    for kpi in kpis:
        df[kpi] = pd.to_numeric(df[kpi], errors="coerce")

    agg_dict = {}

    for kpi in kpis:
        if kpi in [
            "PDCP SDU Volume, DL",
            "PDCP SDU Volume, UL",
            "Total LTE data volume, DL + UL",
            "Avg RRC conn UE",
            "RRC Connected UEs Max (M8051C56)"
        ]:
            agg_dict[kpi] = "sum"
        else:
            agg_dict[kpi] = "mean"

    if daily:
        df["Date"] = df["Period start time"].dt.normalize()
        time_col = "Date"
    else:
        time_col = "Period start time"

    if not group:
        group_cols = [time_col]

        if "LNCEL name" in df.columns:
            group_cols.append("LNCEL name")

        grouped = df.groupby(group_cols, as_index=False).agg(agg_dict)

    else:
        grouped = df.groupby([time_col], as_index=False).agg(agg_dict)

    return grouped

plot_df = aggregate_data(plot_df, selected_kpis, daily_option, group_option)

time_col = "Date" if daily_option else "Period start time"
plot_df[time_col] = pd.to_datetime(plot_df[time_col], errors="coerce")
plot_df = plot_df.dropna(subset=[time_col])

plot_df["Time_str"] = plot_df[time_col].dt.strftime(
    "%Y-%m-%d" if daily_option else "%Y-%m-%d %H:%M"
)

# ---------------- DASHBOARD ----------------
figures_png = []

if not plot_df.empty:

    cols = st.columns(2)

    for idx, selected_kpi in enumerate(selected_kpis[:4]):

        plt.figure(figsize=(10,4))

        if not group_option and "LNCEL name" in plot_df.columns:

            for cell in plot_df["LNCEL name"].unique():

                cell_df = plot_df[plot_df["LNCEL name"] == cell]

                plt.plot(
                    cell_df["Time_str"],
                    cell_df[selected_kpi],
                    marker="o",
                    label=cell
                )

        else:

            plt.plot(
                plot_df["Time_str"],
                plot_df[selected_kpi],
                marker="o",
                label=selected_kpi
            )

        plt.title(selected_kpi)
        plt.xticks(rotation=45)
        plt.legend()
        plt.tight_layout()

        buf = io.BytesIO()
        plt.savefig(buf, format="png")
        buf.seek(0)
        plt.close()

        figures_png.append(buf)

        cols[idx % 2].image(buf, use_container_width=True)

# ---------------- CREATE PPT ----------------
def create_ppt(figures_png):

    prs = Presentation()

    prs.slide_width = Inches(13.33)
    prs.slide_height = Inches(7.5)

    positions = [
        (Inches(0.5), Inches(0.5)),
        (Inches(6.9), Inches(0.5)),
        (Inches(0.5), Inches(4.0)),
        (Inches(6.9), Inches(4.0))
    ]

    chart_width = Inches(6.08)
    chart_height = Inches(3.04)

    for idx, buf in enumerate(figures_png):

        if idx % 4 == 0:
            slide = prs.slides.add_slide(prs.slide_layouts[5])

        pos_idx = idx % 4

        slide.shapes.add_picture(
            buf,
            positions[pos_idx][0],
            positions[pos_idx][1],
            width=chart_width,
            height=chart_height
        )

    ppt_buffer = io.BytesIO()
    prs.save(ppt_buffer)
    ppt_buffer.seek(0)

    return ppt_buffer

# ---------------- DOWNLOAD PPT ----------------
if figures_png:

    ppt_file = create_ppt(figures_png)

    st.download_button(
        "📊 Download PowerPoint Report",
        data=ppt_file,
        file_name="LTE_KPI_Report.pptx",
        mime="application/vnd.openxmlformats-officedocument.presentationml.presentation"
    )

else:

    st.warning("⚠️ No data available for the selected filters.")
