import pandas as pd
import plotly.express as px
import plotly.io as pio
import streamlit as st

# --------------------------------
# Page setup
# --------------------------------
st.set_page_config(page_title="Spend & Leads Dashboard", layout="wide")

# Consistent professional theme
pio.templates.default = "plotly_dark"

# Light UI polish (spacing + typography)
st.markdown(
    """
    <style>
      .block-container { padding-top: 1.2rem; padding-bottom: 2rem; }
      section[data-testid="stSidebar"] { padding-top: 1rem; }
      h1, h2, h3 { letter-spacing: 0.2px; }
      [data-testid="stCaptionContainer"] { opacity: 0.85; }
      [data-testid="stMetricValue"] { font-size: 1.7rem; }
    </style>
    """,
    unsafe_allow_html=True,
)

st.title("Spend and Leads Dashboard")
st.caption("Interactive performance dashboard with filters, rankings, breakdowns, and export.")

# --------------------------------
# Upload
# --------------------------------
uploaded_file = st.file_uploader("Upload Excel file", type=["xlsx"])

if not uploaded_file:
    st.info("Upload your Excel file to view the dashboard.")
    st.stop()

# --------------------------------
# Load & clean
# --------------------------------
df = pd.read_excel(uploaded_file)
df.columns = df.columns.str.strip()

# Rename columns (matches your Excel headers)
df = df.rename(columns={
    "Brand": "brand",
    "Destination": "destination",
    "Leads": "leads",
    "Spent (GBP)": "spent_gbp",
    "Month": "month"
})

required_cols = {"brand", "destination", "leads", "spent_gbp", "month"}
missing = required_cols - set(df.columns)
if missing:
    st.error(f"Missing required columns in Excel: {', '.join(sorted(missing))}")
    st.stop()

# Types
df["leads"] = pd.to_numeric(df["leads"], errors="coerce").fillna(0).astype(int)
df["spent_gbp"] = pd.to_numeric(df["spent_gbp"], errors="coerce").fillna(0.0)

df["month"] = df["month"].astype(str).str.strip()
df["brand"] = df["brand"].astype(str).str.strip()
df["destination"] = df["destination"].astype(str).str.strip()

# Month order (Power BI-like)
month_order = [
    "January","February","March","April","May","June",
    "July","August","September","October","November","December"
]
df["month"] = pd.Categorical(df["month"], categories=month_order, ordered=True)

# --------------------------------
# Sidebar filters
# --------------------------------
st.sidebar.header("Filters")

available_months = [m for m in month_order if m in df["month"].dropna().unique().tolist()]
if not available_months:
    available_months = sorted(df["month"].dropna().astype(str).unique().tolist())

month = st.sidebar.selectbox("Month", available_months)
d = df[df["month"] == month].copy()

brand = st.sidebar.selectbox("Brand", ["All"] + sorted(d["brand"].dropna().unique()))
if brand != "All":
    d = d[d["brand"] == brand]

destination = st.sidebar.selectbox("Destination", ["All"] + sorted(d["destination"].dropna().unique()))
if destination != "All":
    d = d[d["destination"] == destination]

# --------------------------------
# Download filtered data
# --------------------------------
st.download_button(
    "Download filtered data (CSV)",
    d.to_csv(index=False).encode("utf-8"),
    file_name=f"filtered_{month}_{brand}_{destination}.csv".replace(" ", "_"),
    mime="text/csv"
)

st.divider()

# --------------------------------
# KPI cards
# --------------------------------
total_spend = float(d["spent_gbp"].sum())
total_leads = int(d["leads"].sum())

k1, k2, k3, k4 = st.columns(4)
k1.metric("Total Spend", f"£{total_spend:,.2f}")
k2.metric("Total Leads", f"{total_leads:,}")
k3.metric("Brands", f"{d['brand'].nunique():,}")
k4.metric("Destinations", f"{d['destination'].nunique():,}")

st.divider()

# --------------------------------
# Brand rankings (Top 3)
# --------------------------------
st.subheader("Top Brands")

r1, r2 = st.columns(2)

top3_spend = (
    d.groupby("brand", as_index=False)["spent_gbp"]
    .sum()
    .sort_values("spent_gbp", ascending=False)
    .head(3)
)

top3_leads = (
    d.groupby("brand", as_index=False)["leads"]
    .sum()
    .sort_values("leads", ascending=False)
    .head(3)
)

with r1:
    st.markdown("**Top 3 by Spend**")
    if top3_spend.empty:
        st.write("No data available for this selection.")
    else:
        for i, row in top3_spend.reset_index(drop=True).iterrows():
            st.write(f"**#{i+1} {row['brand']}** — £{row['spent_gbp']:,.2f}")

with r2:
    st.markdown("**Top 3 by Leads**")
    if top3_leads.empty:
        st.write("No data available for this selection.")
    else:
        for i, row in top3_leads.reset_index(drop=True).iterrows():
            st.write(f"**#{i+1} {row['brand']}** — {int(row['leads']):,} leads")

st.divider()

# --------------------------------
# Brand performance charts
# --------------------------------
st.subheader("Brand Performance")

c1, c2 = st.columns(2)

spend_by_brand = (
    d.groupby("brand", as_index=False)["spent_gbp"]
    .sum()
    .sort_values("spent_gbp", ascending=True)
)
fig_spend_brand = px.bar(
    spend_by_brand,
    x="spent_gbp",
    y="brand",
    orientation="h",
    title="Spend by Brand"
)
fig_spend_brand.update_layout(xaxis_title="Spend (GBP)", yaxis_title="Brand")
c1.plotly_chart(fig_spend_brand, use_container_width=True)

leads_by_brand = (
    d.groupby("brand", as_index=False)["leads"]
    .sum()
    .sort_values("leads", ascending=True)
)
fig_leads_brand = px.bar(
    leads_by_brand,
    x="leads",
    y="brand",
    orientation="h",
    title="Leads by Brand"
)
fig_leads_brand.update_layout(xaxis_title="Leads", yaxis_title="Brand")
c2.plotly_chart(fig_leads_brand, use_container_width=True)

st.divider()

# --------------------------------
# Decomposition (Power BI-style drill-down)
# --------------------------------
st.subheader("Decomposition: Brand to Destination")
st.caption("Clean drill-down view for analysis. Use the options below to switch levels.")

view = st.radio(
    "Breakdown level",
    ["Brand summary", "Destination summary", "Brand to destination detail"],
    horizontal=True
)

if view == "Brand summary":
    brand_summary = (
        d.groupby("brand", as_index=False)
        .agg(spent_gbp=("spent_gbp", "sum"), leads=("leads", "sum"))
        .sort_values("spent_gbp", ascending=False)
    )
    st.dataframe(brand_summary, use_container_width=True)

elif view == "Destination summary":
    dest_summary = (
        d.groupby("destination", as_index=False)
        .agg(spent_gbp=("spent_gbp", "sum"), leads=("leads", "sum"))
        .sort_values("spent_gbp", ascending=False)
    )
    st.dataframe(dest_summary, use_container_width=True)

else:
    detail = (
        d.groupby(["brand", "destination"], as_index=False)
        .agg(spent_gbp=("spent_gbp", "sum"), leads=("leads", "sum"))
        .sort_values("spent_gbp", ascending=False)
    )
    st.dataframe(detail, use_container_width=True)

st.divider()

# --------------------------------
# Top destinations (simple, readable)
# --------------------------------
st.subheader("Top Destinations")

top_n = st.slider("Number of destinations to show", 5, 30, 10)

top_dest = (
    d.groupby("destination", as_index=False)
    .agg(spent_gbp=("spent_gbp", "sum"), leads=("leads", "sum"))
    .sort_values("spent_gbp", ascending=False)
    .head(top_n)
)

colA, colB = st.columns(2)

fig_dest_spend = px.bar(
    top_dest.sort_values("spent_gbp"),
    x="spent_gbp",
    y="destination",
    orientation="h",
    title="Top Destinations by Spend"
)
fig_dest_spend.update_layout(xaxis_title="Spend (GBP)", yaxis_title="Destination")
colA.plotly_chart(fig_dest_spend, use_container_width=True)

fig_dest_leads = px.bar(
    top_dest.sort_values("leads"),
    x="leads",
    y="destination",
    orientation="h",
    title="Top Destinations by Leads"
)
fig_dest_leads.update_layout(xaxis_title="Leads", yaxis_title="Destination")
colB.plotly_chart(fig_dest_leads, use_container_width=True)

# --------------------------------
# Detail table at the bottom (optional but useful)
# --------------------------------
with st.expander("Show full filtered data table"):
    st.dataframe(d, use_container_width=True)


