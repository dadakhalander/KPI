```python
# ============================================================
# STREAMLIT KPI DASHBOARD FOR CONFECTIONARY SALES DATA
# FULL CLEAN PIPELINE | EDA + KPI + INTERACTIVE VISUALIZATION
# ============================================================

# ============================================================
# REQUIRED INSTALLATION
# pip install -r requirements.txt
# streamlit run app.py
# ============================================================

# ============================================================
# IMPORT LIBRARIES
# ============================================================
import streamlit as st
import pandas as pd
import numpy as np
import plotly.express as px
import plotly.graph_objects as go
from plotly.subplots import make_subplots
import missingno as msno
import matplotlib.pyplot as plt
import seaborn as sns
from io import BytesIO

# ============================================================
# PAGE CONFIGURATION
# ============================================================
st.set_page_config(
    page_title="Confectionary Sales Dashboard",
    page_icon="🍬",
    layout="wide",
    initial_sidebar_state="expanded"
)

# ============================================================
# CUSTOM CSS
# ============================================================
st.markdown("""
<style>
.metric-card {
    background-color: #111827;
    padding: 20px;
    border-radius: 12px;
    text-align: center;
    color: white;
    box-shadow: 0px 4px 8px rgba(0,0,0,0.2);
}
.section-header {
    font-size: 28px;
    font-weight: bold;
    color: #4B5563;
    margin-top: 20px;
    margin-bottom: 20px;
}
</style>
""", unsafe_allow_html=True)

# ============================================================
# LOAD DATA FUNCTION
# ============================================================
@st.cache_data
def load_data():
    excel_file = "Confectionary_[4564].xlsx"
    xls = pd.ExcelFile(excel_file)

    df = pd.read_excel(excel_file, sheet_name=xls.sheet_names[0])

    # Rename columns
    df.columns = [
        'Date',
        'Country',
        'Confectionary',
        'Units_Sold',
        'Cost',
        'Profit',
        'Revenue'
    ]

    # Convert data types
    df['Date'] = pd.to_datetime(df['Date'])

    # Fill missing values
    df['Units_Sold'].fillna(df['Units_Sold'].median(), inplace=True)
    df['Cost'].fillna(df['Cost'].median(), inplace=True)
    df['Profit'].fillna(df['Profit'].median(), inplace=True)

    # Feature Engineering
    df['Year'] = df['Date'].dt.year
    df['Month'] = df['Date'].dt.month_name()
    df['Month_Num'] = df['Date'].dt.month
    df['Quarter'] = df['Date'].dt.quarter
    df['Profit_Margin'] = (df['Profit'] / df['Revenue']) * 100

    return df

df = load_data()

# ============================================================
# SIDEBAR FILTERS
# ============================================================
st.sidebar.title("Dashboard Filters")

selected_countries = st.sidebar.multiselect(
    "Select Country",
    options=sorted(df['Country'].unique()),
    default=sorted(df['Country'].unique())
)

selected_products = st.sidebar.multiselect(
    "Select Confectionary Product",
    options=sorted(df['Confectionary'].unique()),
    default=sorted(df['Confectionary'].unique())
)

selected_years = st.sidebar.multiselect(
    "Select Year",
    options=sorted(df['Year'].unique()),
    default=sorted(df['Year'].unique())
)

# Apply filters
filtered_df = df[
    (df['Country'].isin(selected_countries)) &
    (df['Confectionary'].isin(selected_products)) &
    (df['Year'].isin(selected_years))
]

# ============================================================
# KPI CALCULATIONS
# ============================================================
total_revenue = filtered_df['Revenue'].sum()
total_profit = filtered_df['Profit'].sum()
total_cost = filtered_df['Cost'].sum()
total_units = filtered_df['Units_Sold'].sum()
avg_profit_margin = filtered_df['Profit_Margin'].mean()

# ============================================================
# HEADER
# ============================================================
st.title("🍬 Confectionary Sales Executive Dashboard")
st.markdown("Comprehensive KPI dashboard with interactive business intelligence visualizations.")

# ============================================================
# KPI CARDS
# ============================================================
col1, col2, col3, col4, col5 = st.columns(5)

col1.metric("Total Revenue", f"£{total_revenue:,.2f}")
col2.metric("Total Profit", f"£{total_profit:,.2f}")
col3.metric("Total Cost", f"£{total_cost:,.2f}")
col4.metric("Units Sold", f"{total_units:,.0f}")
col5.metric("Avg Profit Margin", f"{avg_profit_margin:.2f}%")

# ============================================================
# REVENUE & PROFIT TRENDS
# ============================================================
st.markdown("## 📈 Revenue and Profit Trends")

col1, col2 = st.columns(2)

revenue_trend = filtered_df.groupby('Date')['Revenue'].sum().reset_index()
profit_trend = filtered_df.groupby('Date')['Profit'].sum().reset_index()

fig1 = px.line(
    revenue_trend,
    x='Date',
    y='Revenue',
    title='Revenue Trend Over Time',
    markers=True
)

fig2 = px.line(
    profit_trend,
    x='Date',
    y='Profit',
    title='Profit Trend Over Time',
    markers=True,
    color_discrete_sequence=['green']
)

col1.plotly_chart(fig1, use_container_width=True)
col2.plotly_chart(fig2, use_container_width=True)

# ============================================================
# COUNTRY ANALYSIS
# ============================================================
st.markdown("## 🌍 Country Performance Analysis")

col1, col2 = st.columns(2)

country_rev = filtered_df.groupby('Country')['Revenue'].sum().reset_index()
country_profit = filtered_df.groupby('Country')['Profit'].sum().reset_index()

fig3 = px.bar(
    country_rev,
    x='Country',
    y='Revenue',
    title='Revenue by Country',
    color='Revenue',
    text_auto=True
)

fig4 = px.bar(
    country_profit,
    x='Country',
    y='Profit',
    title='Profit by Country',
    color='Profit',
    text_auto=True
)

col1.plotly_chart(fig3, use_container_width=True)
col2.plotly_chart(fig4, use_container_width=True)

# ============================================================
# PRODUCT ANALYSIS
# ============================================================
st.markdown("## 🍫 Product Performance Analysis")

col1, col2 = st.columns(2)

product_rev = filtered_df.groupby('Confectionary')['Revenue'].sum().reset_index()

fig5 = px.pie(
    product_rev,
    names='Confectionary',
    values='Revenue',
    title='Revenue Share by Product'
)

fig6 = px.bar(
    product_rev,
    x='Confectionary',
    y='Revenue',
    title='Revenue by Product',
    color='Revenue'
)

col1.plotly_chart(fig5, use_container_width=True)
col2.plotly_chart(fig6, use_container_width=True)

# ============================================================
# MONTHLY HEATMAP
# ============================================================
st.markdown("## 🔥 Monthly Revenue Heatmap")

month_order = [
    'January', 'February', 'March', 'April', 'May', 'June',
    'July', 'August', 'September', 'October', 'November', 'December'
]

heatmap_data = filtered_df.pivot_table(
    values='Revenue',
    index='Month',
    columns='Year',
    aggfunc='sum'
).reindex(month_order)

fig7 = px.imshow(
    heatmap_data,
    title='Monthly Revenue Heatmap',
    color_continuous_scale='Viridis',
    aspect='auto'
)

st.plotly_chart(fig7, use_container_width=True)

# ============================================================
# SCATTER ANALYSIS
# ============================================================
st.markdown("## 📊 Profit vs Revenue Analysis")

fig8 = px.scatter(
    filtered_df,
    x='Revenue',
    y='Profit',
    color='Confectionary',
    size='Units_Sold',
    hover_data=['Country'],
    title='Profit vs Revenue by Product'
)

st.plotly_chart(fig8, use_container_width=True)

# ============================================================
# BOX PLOT ANALYSIS
# ============================================================
st.markdown("## 📦 Revenue Distribution by Product")

fig9 = px.box(
    filtered_df,
    x='Confectionary',
    y='Revenue',
    color='Confectionary',
    title='Revenue Distribution'
)

st.plotly_chart(fig9, use_container_width=True)

# ============================================================
# CORRELATION MATRIX
# ============================================================
st.markdown("## 🔗 Correlation Analysis")

numeric_df = filtered_df[['Units_Sold', 'Cost', 'Profit', 'Revenue', 'Profit_Margin']]
corr = numeric_df.corr()

fig10 = px.imshow(
    corr,
    text_auto=True,
    title='Correlation Matrix'
)

st.plotly_chart(fig10, use_container_width=True)

# ============================================================
# DATA EXPLORER
# ============================================================
st.markdown("## 📋 Raw Data Explorer")

st.dataframe(filtered_df, use_container_width=True)

# ============================================================
# DOWNLOAD FILTERED DATA
# ============================================================
csv = filtered_df.to_csv(index=False).encode('utf-8')

st.download_button(
    label="⬇️ Download Filtered Data as CSV",
    data=csv,
    file_name="filtered_confectionary_sales.csv",
    mime="text/csv"
)

# ============================================================
# FOOTER
# ============================================================
st.markdown("---")
st.caption("Developed using Streamlit + Plotly | Business Intelligence Dashboard")
```
