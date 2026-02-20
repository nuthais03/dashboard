import io
import math
import pandas as pd
import streamlit as st
import matplotlib.pyplot as plt
from reportlab.lib.pagesizes import A4
from reportlab.pdfgen import canvas
from reportlab.lib.utils import ImageReader

# -----------------------------
# Page setup + small CSS polish
# -----------------------------
st.set_page_config(page_title="Spend & Leads Dashboard (Matplotlib)", layout="wide")

st.markdown(
    """
    <style>
      .block-container { padding-top: 1.2rem; padding-bottom: 2rem; }
      section[data-testid="stSidebar"] { padding-top: 1rem; }
      h1, h2, h3 { letter-spacing: 0.2px; }
      [data-testid="stCaptionContainer"] { opacity: 0.85; }
      [data-testid="stMetricValue"] { font-size: 1.7rem; }
      a { text-decoration: none; }
    </style>
    """,
    unsafe_allow_html=True,
)

# Anchor for "go to upload"
st.markdown('<div id="upload"></div>', unsafe_allow_html=True)

st.title("Spend, Leads & Messages Dashboard (Matplotlib)")
st.caption("Power BI-like filters + editable inputs + PDF export (no monthly trend).")

# -----------------------------
# Template (CSV) generator
# -----------------------------
TEMPLATE_COLUMNS = [
    "Month",
    "Brand",
    "Destination",
    "Spent (GBP)",
    "Leads",
    "Messages",
    "Impressions",
    "Converted Leads",
]
def template_csv_bytes() -> bytes:
    temp = pd.DataFrame(columns=TEMPLATE_COLUMNS)
    return temp.to_csv(index=False).encode("utf-8")

# -----------------------------
# Upload area (always visible)
# -----------------------------
top_left, top_right = st.columns([2, 1])
with top_left:
    uploaded_file = st.file_uploader("Upload Excel file (.xlsx)", type=["xlsx"], key="uploader")

with top_right:
    st.download_button(
        "Download CSV Template",
        data=template_csv_bytes(),
        file_name="dashboard_template.csv",
        mime="text/csv",
        help="Use this template to prepare your data (you can upload Excel now; CSV template is for later)."
    )

if not uploaded_file:
    st.info("Upload your Excel file to view the dashboard.")
    st.stop()

# -----------------------------
# Load & standardize columns
# -----------------------------
df = pd.read_excel(uploaded_file)
df.columns = df.columns.astype(str).str.strip()

# Accept flexible headers but output consistent internal names
rename_map = {
    "Month": "month",
    "Brand": "brand",
    "Destination": "destination",
    "Spent (GBP)": "spent_gbp",
    "Spend (GBP)": "spent_gbp",
    "Spent": "spent_gbp",
    "Spend": "spent_gbp",
    "Leads": "leads",
    "Messages": "messages",
    "Impressions": "impressions",
    "Converted Leads": "converted_leads",
    "Converted": "converted_leads",
}
df = df.rename(columns={c: rename_map.get(c, c) for c in df.columns})

required_cols = {"month", "brand", "destination", "spent_gbp", "leads"}
missing = required_cols - set(df.columns)
if missing:
    st.error(f"Missing required columns in Excel: {', '.join(sorted(missing))}")
    st.stop()

# Optional columns: create if missing
if "messages" not in df.columns:
    df["messages"] = 0
if "impressions" not in df.columns:
    df["impressions"] = 0
if "converted_leads" not in df.columns:
    df["converted_leads"] = 0

# Types
df["month"] = df["month"].astype(str).str.strip()
df["brand"] = df["brand"].astype(str).str.strip()
df["destination"] = df["destination"].astype(str).str.strip()

df["spent_gbp"] = pd.to_numeric(df["spent_gbp"], errors="coerce").fillna(0.0)
df["leads"] = pd.to_numeric(df["leads"], errors="coerce").fillna(0).astype(int)
df["messages"] = pd.to_numeric(df["messages"], errors="coerce").fillna(0).astype(int)
df["impressions"] = pd.to_numeric(df["impressions"], errors="coerce").fillna(0).astype(int)
df["converted_leads"] = pd.to_numeric(df["converted_leads"], errors="coerce").fillna(0).astype(int)

# -----------------------------
# Sidebar filters
# -----------------------------
st.sidebar.header("Filters")

month_order = [
    "January","February","March","April","May","June",
    "July","August","September","October","November","December"
]
# Keep user's months list stable even if not full month names
available_months = [m for m in month_order if m in df["month"].unique().tolist()]
if not available_months:
    available_months = sorted(df["month"].dropna().unique().tolist())

month = st.sidebar.selectbox("Month", available_months)
d = df[df["month"] == month].copy()

brand = st.sidebar.selectbox("Brand", ["All"] + sorted(d["brand"].dropna().unique()))
if brand != "All":
    d = d[d["brand"] == brand]

destination = st.sidebar.selectbox("Destination", ["All"] + sorted(d["destination"].dropna().unique()))
if destination != "All":
    d = d[d["destination"] == destination]

# -----------------------------
# Editable data (manual input)
# - user can type converted leads, impressions, messages, etc.
# - formulas computed after editing
# -----------------------------
st.subheader("Data Entry (Editable)")
st.caption("You can manually edit Messages / Impressions / Converted Leads here. Formulas will update automatically.")

editable_cols = ["month", "brand", "destination", "spent_gbp", "leads", "messages", "impressions", "converted_leads"]
d_edit = d[editable_cols].copy()

edited = st.data_editor(
    d_edit,
    use_container_width=True,
    hide_index=True,
    num_rows="dynamic",
    column_config={
        "spent_gbp": st.column_config.NumberColumn("Spent (GBP)", min_value=0.0, step=1.0),
        "leads": st.column_config.NumberColumn("Leads", min_value=0, step=1),
        "messages": st.column_config.NumberColumn("Messages", min_value=0, step=1),
        "impressions": st.column_config.NumberColumn("Impressions", min_value=0, step=100),
        "converted_leads": st.column_config.NumberColumn("Converted Leads", min_value=0, step=1),
    },
    key="editor_matplot"
)

# Clean edited values (in case)
edited["spent_gbp"] = pd.to_numeric(edited["spent_gbp"], errors="coerce").fillna(0.0)
edited["leads"] = pd.to_numeric(edited["leads"], errors="coerce").fillna(0).astype(int)
edited["messages"] = pd.to_numeric(edited["messages"], errors="coerce").fillna(0).astype(int)
edited["impressions"] = pd.to_numeric(edited["impressions"], errors="coerce").fillna(0).astype(int)
edited["converted_leads"] = pd.to_numeric(edited["converted_leads"], errors="coerce").fillna(0).astype(int)

# -----------------------------
# Derived metrics (formulas)
# -----------------------------
def safe_div(n, d):
    return (n / d) if d not in (0, None) else 0.0

edited["cpl"] = edited.apply(lambda r: safe_div(r["spent_gbp"], r["leads"]), axis=1)
edited["conversion_rate"] = edited.apply(lambda r: safe_div(r["converted_leads"], r["leads"]), axis=1)

# KPIs from edited data (not original)
total_spend = float(edited["spent_gbp"].sum())
total_leads = int(edited["leads"].sum())
total_messages = int(edited["messages"].sum())
total_impressions = int(edited["impressions"].sum())
total_converted = int(edited["converted_leads"].sum())

overall_cpl = safe_div(total_spend, total_leads)
overall_cr = safe_div(total_converted, total_leads)

st.divider()

# -----------------------------
# KPI cards + “go to upload”
# -----------------------------
k1, k2, k3, k4, k5, k6 = st.columns(6)
k1.metric("Total Spend", f"£{total_spend:,.2f}")
k2.metric("Total Leads", f"{total_leads:,}")
k3.metric("Total Messages", f"{total_messages:,}")
k4.metric("Impressions", f"{total_impressions:,}")
k5.metric("CPL", f"£{overall_cpl:,.2f}")
k6.metric("Conversion Rate", f"{overall_cr*100:,.2f}%")

# “Click leads → go to upload area” (Streamlit metrics aren’t clickable, so we give a button right under)
st.markdown("[⬆️ Go to Upload area](#upload)")

# Download edited data
st.download_button(
    "Download edited data (CSV)",
    data=edited.to_csv(index=False).encode("utf-8"),
    file_name=f"edited_{month}_{brand}_{destination}.csv".replace(" ", "_"),
    mime="text/csv"
)

st.divider()

# -----------------------------
# Charts helpers (Matplotlib)
# -----------------------------
def barh_chart(df_plot, label_col, value_col, title, xlabel):
    df_plot = df_plot.copy()
    df_plot = df_plot.sort_values(value_col, ascending=True)

    fig, ax = plt.subplots(figsize=(8, 4.8))
    ax.barh(df_plot[label_col], df_plot[value_col])
    ax.set_title(title)
    ax.set_xlabel(xlabel)
    ax.set_ylabel("")
    ax.grid(axis="x", linestyle="--", alpha=0.3)
    fig.tight_layout()
    return fig

# -----------------------------
# Two dashboards: Leads + Messages placeholder
# -----------------------------
tab_leads, tab_messages = st.tabs(["Leads Dashboard", "Messages Dashboard (placeholder)"])

with tab_leads:
    st.subheader("Company-wise and Destination-wise View (Leads)")
    top_n = st.slider("Top N items", 5, 30, 10, key="topn_leads")

    # Company-wise (Brand)
    by_brand = edited.groupby("brand", as_index=False).agg(
        spent_gbp=("spent_gbp", "sum"),
        leads=("leads", "sum"),
        converted_leads=("converted_leads", "sum"),
        impressions=("impressions", "sum")
    )
    by_brand["cpl"] = by_brand.apply(lambda r: safe_div(r["spent_gbp"], r["leads"]), axis=1)
    by_brand["conversion_rate"] = by_brand.apply(lambda r: safe_div(r["converted_leads"], r["leads"]), axis=1)

    # Destination-wise
    by_dest = edited.groupby("destination", as_index=False).agg(
        spent_gbp=("spent_gbp", "sum"),
        leads=("leads", "sum"),
        converted_leads=("converted_leads", "sum"),
        impressions=("impressions", "sum")
    )
    by_dest["cpl"] = by_dest.apply(lambda r: safe_div(r["spent_gbp"], r["leads"]), axis=1)
    by_dest["conversion_rate"] = by_dest.apply(lambda r: safe_div(r["converted_leads"], r["leads"]), axis=1)

    cA, cB = st.columns(2)

    # Charts: Leads by brand / Leads by destination
    with cA:
        st.markdown("### Leads by Company (Brand)")
        fig = barh_chart(by_brand.sort_values("leads", ascending=False).head(top_n), "brand", "leads", "Top Companies by Leads", "Leads")
        st.pyplot(fig, use_container_width=True)

    with cB:
        st.markdown("### Leads by Destination")
        fig = barh_chart(by_dest.sort_values("leads", ascending=False).head(top_n), "destination", "leads", "Top Destinations by Leads", "Leads")
        st.pyplot(fig, use_container_width=True)

    st.markdown("### Summary Tables")
    t1, t2 = st.columns(2)

    with t1:
        show_brand = by_brand.sort_values("spent_gbp", ascending=False).copy()
        show_brand["conversion_rate"] = (show_brand["conversion_rate"] * 100).round(2).astype(str) + "%"
        st.dataframe(show_brand, use_container_width=True, hide_index=True)

    with t2:
        show_dest = by_dest.sort_values("spent_gbp", ascending=False).copy()
        show_dest["conversion_rate"] = (show_dest["conversion_rate"] * 100).round(2).astype(str) + "%"
        st.dataframe(show_dest, use_container_width=True, hide_index=True)

with tab_messages:
    st.subheader("Messages Dashboard (Coming Soon)")
    st.caption("This space is reserved. Later we can connect Messages data and build charts similar to Leads.")

    # Keep a visible placeholder layout
    p1, p2, p3 = st.columns(3)
    p1.metric("Total Messages", f"{total_messages:,}")
    p2.metric("Messages (CPL placeholder)", "—")
    p3.metric("Message Conversion placeholder", "—")

    st.info("When you’re ready, tell me what column or source you use for Messages conversion, and I’ll build the full Messages dashboard here.")

st.divider()

# -----------------------------
# PDF Summary Report (download)
# -----------------------------
st.subheader("Download PDF Summary Report")
st.caption("Creates a quick one-page PDF summary of KPIs + top tables + one chart snapshot.")

def fig_to_png_bytes(fig) -> bytes:
    buf = io.BytesIO()
    fig.savefig(buf, format="png", dpi=200, bbox_inches="tight")
    plt.close(fig)
    return buf.getvalue()

def build_pdf_bytes() -> bytes:
    pdf_buf = io.BytesIO()
    c = canvas.Canvas(pdf_buf, pagesize=A4)
    w, h = A4

    # Title
    c.setFont("Helvetica-Bold", 14)
    c.drawString(36, h - 40, "Spend, Leads & Messages — Summary Report")

    c.setFont("Helvetica", 10)
    c.drawString(36, h - 58, f"Filters: Month={month} | Brand={brand} | Destination={destination}")

    # KPIs
    y = h - 90
    c.setFont("Helvetica-Bold", 11)
    c.drawString(36, y, "Key Metrics")
    c.setFont("Helvetica", 10)
    y -= 16
    c.drawString(36, y, f"Total Spend: £{total_spend:,.2f}")
    y -= 14
    c.drawString(36, y, f"Total Leads: {total_leads:,}")
    y -= 14
    c.drawString(36, y, f"Total Messages: {total_messages:,}")
    y -= 14
    c.drawString(36, y, f"Impressions: {total_impressions:,}")
    y -= 14
    c.drawString(36, y, f"CPL: £{overall_cpl:,.2f}")
    y -= 14
    c.drawString(36, y, f"Conversion Rate: {overall_cr*100:,.2f}%")

    # Top table (Brands by spend)
    y -= 22
    c.setFont("Helvetica-Bold", 11)
    c.drawString(36, y, "Top Companies (by Spend)")
    y -= 14
    c.setFont("Helvetica", 9)

    top_brands_pdf = (
        edited.groupby("brand", as_index=False)["spent_gbp"]
        .sum()
        .sort_values("spent_gbp", ascending=False)
        .head(5)
    )

    for _, r in top_brands_pdf.iterrows():
        c.drawString(36, y, f"{r['brand']}: £{float(r['spent_gbp']):,.2f}")
        y -= 12

    # Add one chart image (Leads by destination top 10)
    y -= 12
    chart_df = (
        edited.groupby("destination", as_index=False)["leads"]
        .sum()
        .sort_values("leads", ascending=False)
        .head(10)
    )
    fig = barh_chart(chart_df, "destination", "leads", "Top Destinations by Leads", "Leads")
    img_bytes = fig_to_png_bytes(fig)
    img = ImageReader(io.BytesIO(img_bytes))

    # place image
    img_w = w - 72
    img_h = 240
    c.drawImage(img, 36, max(36, y - img_h), width=img_w, height=img_h, preserveAspectRatio=True, mask="auto")

    c.showPage()
    c.save()
    return pdf_buf.getvalue()

pdf_bytes = build_pdf_bytes()

st.download_button(
    "Download PDF Report",
    data=pdf_bytes,
    file_name=f"summary_report_{month}.pdf".replace(" ", "_"),
    mime="application/pdf"
)
