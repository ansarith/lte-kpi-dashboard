import streamlit as st
import pandas as pd
import matplotlib.pyplot as plt
import io
from pptx import Presentation
from pptx.util import Inches

st.set_page_config(layout="wide")

# ---------------- LOAD DATA ----------------
@st.cache_data
def load_data(path):
    df = pd.read_excel(path)
    df["Period start time"] = pd.to_datetime(df["Period start time"], errors="coerce")
    # Convert ratios to %
    for col in df.columns:
        if "%" in col or "Rate" in col:
            if df[col].max() <= 1.0:
                df[col] = df[col] * 100
    return df

DATA_PATH = "4G_Main_KPIs_Report_SRAN21B-Sarith-2025_10_13-Site KCH2567RBR & 2070_LB.xlsx"
df = load_data(DATA_PATH)

st.title("📊 LTE KPI Dashboard")

# ---------------- SELECT KPI / SITE / CELL ----------------
kpi_columns = [c for c in df.columns if c not in ["Period start time","LNBTS name","LNCEL name"]]
selected_kpis = st.multiselect("Select KPI(s)", options=kpi_columns, default=kpi_columns[:4])
enodeb_selected = st.multiselect("Select LNBTS name", options=sorted(df["LNBTS name"].unique()))
if enodeb_selected:
    cell_options = sorted(df[df["LNBTS name"].isin(enodeb_selected)]["LNCEL name"].unique())
else:
    cell_options = sorted(df["LNCEL name"].unique())
cell_selected = st.multiselect("Select LNCEL name", options=cell_options)

daily_option = st.checkbox("📅 Daily Aggregation")
group_option = st.checkbox("🏙️ Group by Site")

# ---------------- FILTER ----------------
plot_df = df.copy()
if enodeb_selected:
    plot_df = plot_df[plot_df["LNBTS name"].isin(enodeb_selected)]
if cell_selected:
    plot_df = plot_df[plot_df["LNCEL name"].isin(cell_selected)]

# ---------------- AGGREGATION ----------------
def aggregate_data(df, kpis, daily=False, group=False):
    for kpi in kpis:
        df[kpi] = pd.to_numeric(df[kpi], errors="coerce")
    agg_dict = {kpi: "sum" if "Volume" in kpi or "RRC" in kpi else "mean" for kpi in kpis}

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
plot_df = plot_df.dropna(subset=[time_col])
plot_df["Time_str"] = plot_df[time_col].dt.strftime("%Y-%m-%d" if daily_option else "%Y-%m-%d %H:%M")

# ---------------- PLOT & EXPORT PNG ----------------
st.subheader("📈 KPI Charts")
figures_png = []  # store PNG bytes for PPT

for kpi in selected_kpis[:4]:
    plt.figure(figsize=(10,4))
    if not group_option and "LNCEL name" in plot_df.columns:
        for cell in plot_df["LNCEL name"].unique():
            cell_df = plot_df[plot_df["LNCEL name"]==cell]
            plt.plot(cell_df["Time_str"], cell_df[kpi], marker='o', label=cell)
    else:
        plt.plot(plot_df["Time_str"], plot_df[kpi], marker='o', label=kpi)
    plt.title(kpi)
    plt.xticks(rotation=45)
    plt.tight_layout()
    plt.legend()
    
    # Save PNG to buffer
    buf = io.BytesIO()
    plt.savefig(buf, format='png')
    buf.seek(0)
    figures_png.append((kpi, buf))  # store for PPT
    plt.close()
    
    # Show in Streamlit
    st.image(buf, caption=kpi, use_column_width=True)

# ---------------- CREATE PPT ----------------
def create_ppt(figures_png):
    prs = Presentation()
    prs.slide_width = Inches(13.33)
    prs.slide_height = Inches(7.5)
    positions = [(Inches(0.5), Inches(0.5)), (Inches(6.9), Inches(0.5)),
                 (Inches(0.5), Inches(4.0)), (Inches(6.9), Inches(4.0))]
    chart_width = Inches(6.08)
    chart_height = Inches(3.04)

    for idx, (kpi, buf) in enumerate(figures_png):
        if idx % 4 == 0:
            slide = prs.slides.add_slide(prs.slide_layouts[5])
        pos_idx = idx % 4
        slide.shapes.add_picture(buf, positions[pos_idx][0], positions[pos_idx][1],
                                 width=chart_width, height=chart_height)

    ppt_buffer = io.BytesIO()
    prs.save(ppt_buffer)
    ppt_buffer.seek(0)
    return ppt_buffer

if figures_png:
    ppt_file = create_ppt(figures_png)
    st.download_button("📊 Download PowerPoint Report",
                       data=ppt_file,
                       file_name="LTE_KPI_Report.pptx",
                       mime="application/vnd.openxmlformats-officedocument.presentationml.presentation")

