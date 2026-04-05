import streamlit as st
import pandas as pd
import matplotlib.pyplot as plt
from io import BytesIO

@st.cache_data
def load_data():
    return pd.read_excel("Demo_Payments_Report_SOA.xlsx", engine="openpyxl")

df = load_data()

st.set_page_config(page_title="Financial Dashboard", layout="wide")
st.title("📊 Financial Tracking Dashboard")

# Sidebar filters
st.sidebar.header("🔍 Filter Data")
vessels = st.sidebar.multiselect("Select Vessel (Consumer)", options=df["Consumer"].dropna().unique())
suppliers = st.sidebar.multiselect("Select Supplier", options=df["Supplier"].dropna().unique())
categories = st.sidebar.multiselect("Select Category", options=df["Category"].dropna().unique())
date_range = st.sidebar.date_input("Select Invoice Date Range", [])

# Apply filters
filtered_df = df.copy()
if vessels:
    filtered_df = filtered_df[filtered_df["Consumer"].isin(vessels)]
if suppliers:
    filtered_df = filtered_df[filtered_df["Supplier"].isin(suppliers)]
if categories:
    filtered_df = filtered_df[filtered_df["Category"].isin(categories)]
if len(date_range) == 2:
    filtered_df = filtered_df[
        (pd.to_datetime(filtered_df["Invoice Date"]) >= pd.to_datetime(date_range[0])) &
        (pd.to_datetime(filtered_df["Invoice Date"]) <= pd.to_datetime(date_range[1]))
    ]

# Summary metrics
st.subheader("📌 Summary Metrics")
col1, col2, col3, col4 = st.columns(4)
col1.metric("Total OPEX", f"{filtered_df['OPEX'].sum():,.2f}")
col2.metric("Total PAID", f"{filtered_df['PAID'].sum():,.2f}")
col3.metric("Total Non-OPEX", f"{filtered_df['Non-Opex'].sum():,.2f}")
col4.metric("Total Balance", f"{filtered_df['Balance'].sum():,.2f}")

# Charts
st.subheader("📈 Visualizations")

# Bar chart: Expenses by Category
category_summary = filtered_df.groupby("Category")[["OPEX", "Non-Opex", "PAID"]].sum()
st.markdown("**Expenses by Category**")
fig1, ax1 = plt.subplots()
category_summary.plot(kind="bar", stacked=True, ax=ax1)
ax1.set_ylabel("Amount")
ax1.set_title("Expenses by Category")
plt.xticks(rotation=45, ha="right")
plt.tight_layout()
st.pyplot(fig1)

# Pie chart: OPEX vs Non-OPEX
st.markdown("**OPEX vs Non-OPEX Distribution**")
opex_total = filtered_df["OPEX"].sum()
non_opex_total = filtered_df["Non-Opex"].sum()
fig2, ax2 = plt.subplots()
ax2.pie([opex_total, non_opex_total], labels=["OPEX", "Non-OPEX"], autopct="%1.1f%%", startangle=90)
ax2.axis("equal")
st.pyplot(fig2)

# Bar chart: Expenses by Vessel
st.markdown("**Expenses by Vessel**")
vessel_summary = filtered_df.groupby("Consumer")[["OPEX", "Non-Opex", "PAID"]].sum()
fig3, ax3 = plt.subplots()
vessel_summary.plot(kind="bar", ax=ax3)
ax3.set_ylabel("Amount")
ax3.set_title("Expenses by Vessel")
plt.xticks(rotation=0)
plt.tight_layout()
st.pyplot(fig3)

# Data table
st.subheader("📄 Filtered Data Table")
st.dataframe(filtered_df)

# Download filtered data
output = BytesIO()
with pd.ExcelWriter(output, engine="openpyxl") as writer:
    filtered_df.to_excel(writer, index=False)
output.seek(0)
st.download_button("📥 Download Filtered Data as Excel", data=output.getvalue(), file_name="filtered_data.xlsx")
