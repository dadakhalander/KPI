# app.py

import pandas as pd
import numpy as np
import streamlit as st
import matplotlib.pyplot as plt
import seaborn as sns
import plotly.express as px

# -----------------------------
# Page Config
# -----------------------------
st.set_page_config(
    page_title="Confectionary Sales Dashboard",
    layout="wide"
)

st.title("Confectionary Sales Dashboard")
st.markdown("Interactive dashboard for revenue, profit, units sold, and confectionary trends.")

# -----------------------------
# Load Data
# -----------------------------
excel_file = "Confectionary_[4564].xlsx"

@st.cache_data
def load_data(file):
    sheets = pd.ExcelFile(file).sheet_names
    df = pd.read_excel(file, sheet_name=sheets[0])
    return df, sheets

df, sheets = load_data(excel_file)

st.sidebar.header("Filters")

# Convert date column
df['Date'] = pd.to_datetime(df['Date'])
df['Month'] = df['Date'].dt.month

# Profit Margin Calculation
df["Profit Margin"] = df.apply(
    lambda row: row["Profit(£)"] / row["Revenue(£)"]
    if row["Revenue(£)"] > 0 else 0,
    axis=1
)

# Sidebar Filters
countries = st.sidebar.multiselect(
    "Select Country",
    options=df['Country(UK)'].unique(),
    default=df['Country(UK)'].unique()
)

confectionaries = st.sidebar.multiselect(
    "Select Confectionary",
    options=df['Confectionary'].unique(),
    default=df['Confectionary'].unique()
)

filtered_df = df[
    (df['Country(UK)'].isin(countries)) &
    (df['Confectionary'].isin(confectionaries))
]

# -----------------------------
# Display Data
# -----------------------------
st.subheader("Dataset Preview")
st.dataframe(filtered_df.head())

st.subheader("Available Sheets")
st.write(sheets)

# -----------------------------
# KPI Section
# -----------------------------
total_revenue = filtered_df['Revenue(£)'].sum()
total_profit = filtered_df['Profit(£)'].sum()
avg_margin = filtered_df['Profit Margin'].mean()
peak_month_overall = filtered_df.groupby('Month')['Units Sold'].sum().idxmax()

col1, col2, col3, col4 = st.columns(4)

col1.metric("Total Revenue", f"£{total_revenue:,.2f}")
col2.metric("Total Profit", f"£{total_profit:,.2f}")
col3.metric("Average Profit Margin", f"{avg_margin:.2%}")
col4.metric("Peak Month Overall", peak_month_overall)

# -----------------------------
# Unique Values
# -----------------------------
st.subheader("Unique Values in Columns")

selected_column = st.selectbox("Select Column", filtered_df.columns)

st.write(filtered_df[selected_column].unique())

# -----------------------------
# Distribution & Skewness
# -----------------------------
st.subheader("Distribution & Skewness")

numeric_cols = ['Units Sold', 'Cost(£)', 'Profit(£)', 'Revenue(£)']

selected_numeric_col = st.selectbox(
    "Select Numeric Column for Distribution",
    numeric_cols
)

skewness = filtered_df[selected_numeric_col].skew()

st.write(f"Skewness of {selected_numeric_col}: {skewness:.2f}")

fig, ax = plt.subplots(figsize=(8, 5))
sns.histplot(filtered_df[selected_numeric_col], kde=True, ax=ax)
ax.set_title(f"Distribution of {selected_numeric_col}")
st.pyplot(fig)

# -----------------------------
# Revenue Summary
# -----------------------------
st.subheader("Total Revenue by Country & Confectionary")

revenue_summary = filtered_df.groupby(
    ['Country(UK)', 'Confectionary']
)['Revenue(£)'].sum().reset_index()

st.dataframe(revenue_summary)

# -----------------------------
# Units Sold by Country & Confectionary
# -----------------------------
st.subheader("Units Sold by Country and Confectionary")

fig_bar = px.bar(
    filtered_df,
    x="Country(UK)",
    y="Units Sold",
    color="Confectionary",
    barmode="group",
    title="Units Sold by Country and Confectionary"
)

st.plotly_chart(fig_bar, use_container_width=True)

# -----------------------------
# Profit Margin Bubble Chart
# -----------------------------
st.subheader("Bubble Chart of Profit Margin")

margin_summary = filtered_df.groupby(
    ['Country(UK)', 'Confectionary']
)['Profit Margin'].mean().reset_index()

fig_scatter = px.scatter(
    margin_summary,
    x="Country(UK)",
    y="Profit Margin",
    size="Profit Margin",
    color="Confectionary",
    hover_name="Confectionary",
    title="Bubble Chart of Profit Margin"
)

st.plotly_chart(fig_scatter, use_container_width=True)

# -----------------------------
# Highest vs Lowest Profit Margin
# -----------------------------
st.subheader("Highest vs Lowest Profit Margin per Country")

highest_margin = margin_summary.loc[
    margin_summary.groupby('Country(UK)')['Profit Margin'].idxmax()
]

lowest_margin = margin_summary.loc[
    margin_summary.groupby('Country(UK)')['Profit Margin'].idxmin()
]

margin_compare = highest_margin[['Country(UK)', 'Profit Margin']].merge(
    lowest_margin[['Country(UK)', 'Profit Margin']],
    on='Country(UK)',
    suffixes=('_Highest', '_Lowest')
)

fig_compare, ax = plt.subplots(figsize=(8, 5))

margin_compare.plot(
    x='Country(UK)',
    y=['Profit Margin_Highest', 'Profit Margin_Lowest'],
    kind='bar',
    ax=ax
)

ax.set_title("Highest vs Lowest Profit Margin per Country")
ax.set_ylabel("Profit Margin")
ax.set_xticklabels(margin_compare['Country(UK)'], rotation=0)

st.pyplot(fig_compare)

# -----------------------------
# Peak Month per Confectionary
# -----------------------------
st.subheader("Peak Month per Confectionary")

peak_months = filtered_df.groupby(
    ['Confectionary', 'Month']
)['Units Sold'].sum().reset_index()

peak_months = peak_months.loc[
    peak_months.groupby('Confectionary')['Units Sold'].idxmax()
]

st.dataframe(peak_months)

# -----------------------------
# Monthly Sales by Confectionary & Country
# -----------------------------
st.subheader("Monthly Sales by Confectionary & Country")

fig_monthly = px.bar(
    filtered_df,
    x="Month",
    y="Units Sold",
    color="Country(UK)",
    facet_col="Confectionary",
    facet_col_wrap=3,
    title="Monthly Sales by Confectionary & Country"
)

fig_monthly.update_layout(
    height=700,
    width=1000
)

st.plotly_chart(fig_monthly, use_container_width=True)
