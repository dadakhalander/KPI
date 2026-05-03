# ============================================================
# CONFECTIONARY SALES STREAMLIT DASHBOARD
# FULL PROFESSIONAL INTERACTIVE BUSINESS DASHBOARD
# ============================================================

# ------------------------------------------------------------
# STEP 1: INSTALL REQUIRED LIBRARIES
# ------------------------------------------------------------
# Run this in terminal / Colab once:
# !pip install streamlit pandas openpyxl plotly streamlit-option-menu streamlit-aggrid --quiet

# ------------------------------------------------------------
# STEP 2: IMPORT LIBRARIES
# ------------------------------------------------------------
import streamlit as st
import pandas as pd
import numpy as np
import plotly.express as px
import plotly.graph_objects as go
from plotly.subplots import make_subplots
from streamlit_option_menu import option_menu
from st_aggrid import AgGrid, GridOptionsBuilder

# ------------------------------------------------------------
# STEP 3: PAGE CONFIGURATION
# ------------------------------------------------------------
st.set_page_config(
    page_title="Confectionary Sales Dashboard",
    page_icon="🍬",
    layout="wide",
    initial_sidebar_state="expanded"
)

# ------------------------------------------------------------
# STEP 4: CUSTOM CSS
# ------------------------------------------------------------
st.markdown("""
    <style>
    .main {
        background-color: #f8f9fa;
    }
    .stMetric {
        background-color: #ffffff;
        padding: 15px;
        border-radius: 10px;
        box-shadow: 2px 2px 8px rgba(0,0,0,0.1);
    }
    </style>
""", unsafe_allow_html=True)

# ------------------------------------------------------------
# STEP 5: LOAD DATA
# ------------------------------------------------------------
@st.cache_data
def load_data():
    excel_file = "Confectionary_[4564].xlsx"
    df = pd.read_excel(excel_file)

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

    # Data Cleaning
    df['Date'] = pd.to_datetime(df['Date'])
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

# ------------------------------------------------------------
# STEP 6: SIDEBAR NAVIGATION
# ------------------------------------------------------------
with st.sidebar:
    selected = option_menu(
        menu_title="Dashboard Menu",
        options=[
            "Overview",
            "Sales Analysis",
            "Product Insights",
            "Country Analysis",
            "Advanced Analytics",
            "Raw Data"
        ],
        icons=[
            "house",
            "graph-up",
            "box-seam",
            "globe",
            "bar-chart-line",
            "table"
        ],
        menu_icon="cast",
        default_index=0
    )

# ------------------------------------------------------------
# STEP 7: GLOBAL FILTERS
# ------------------------------------------------------------
st.sidebar.header("Filter Data")

country_filter = st.sidebar.multiselect(
    "Select Country",
    options=df['Country'].unique(),
    default=df['Country'].unique()
)

product_filter = st.sidebar.multiselect(
    "Select Product",
    options=df['Confectionary'].unique(),
    default=df['Confectionary'].unique()
)

year_filter = st.sidebar.multiselect(
    "Select Year",
    options=sorted(df['Year'].unique()),
    default=sorted(df['Year'].unique())
)

filtered_df = df[
    (df['Country'].isin(country_filter)) &
    (df['Confectionary'].isin(product_filter)) &
    (df['Year'].isin(year_filter))
]

# ------------------------------------------------------------
# STEP 8: KPI CALCULATIONS
# ------------------------------------------------------------
total_revenue = filtered_df['Revenue'].sum()
total_profit = filtered_df['Profit'].sum()
total_cost = filtered_df['Cost'].sum()
total_units = filtered_df['Units_Sold'].sum()
avg_profit_margin = filtered_df['Profit_Margin'].mean()

# ============================================================
# OVERVIEW PAGE
# ============================================================
if selected == "Overview":

    st.title("🍬 Confectionary Sales Executive Dashboard")

    # KPI Cards
    col1, col2, col3, col4, col5 = st.columns(5)

    col1.metric("Total Revenue", f"£{total_revenue:,.2f}")
    col2.metric("Total Profit", f"£{total_profit:,.2f}")
    col3.metric("Total Cost", f"£{total_cost:,.2f}")
    col4.metric("Units Sold", f"{total_units:,.0f}")
    col5.metric("Profit Margin", f"{avg_profit_margin:.2f}%")

    st.markdown("---")

    # Revenue Trend
    revenue_trend = filtered_df.groupby('Date')['Revenue'].sum().reset_index()

    fig = px.line(
        revenue_trend,
        x='Date',
        y='Revenue',
        title='Revenue Trend Over Time',
        markers=True
    )

    st.plotly_chart(fig, use_container_width=True)

    # Profit Trend
    profit_trend = filtered_df.groupby('Date')['Profit'].sum().reset_index()

    fig2 = px.line(
        profit_trend,
        x='Date',
        y='Profit',
        title='Profit Trend Over Time',
        markers=True,
        color_discrete_sequence=['green']
    )

    st.plotly_chart(fig2, use_container_width=True)

# ============================================================
# SALES ANALYSIS PAGE
# ============================================================
elif selected == "Sales Analysis":

    st.title("📈 Sales Analysis")

    col1, col2 = st.columns(2)

    # Monthly Revenue
    monthly_rev = filtered_df.groupby(['Month', 'Month_Num'])['Revenue'].sum().reset_index()
    monthly_rev = monthly_rev.sort_values('Month_Num')

    fig = px.bar(
        monthly_rev,
        x='Month',
        y='Revenue',
        title='Monthly Revenue Analysis',
        color='Revenue'
    )

    col1.plotly_chart(fig, use_container_width=True)

    # Monthly Profit
    monthly_profit = filtered_df.groupby(['Month', 'Month_Num'])['Profit'].sum().reset_index()
    monthly_profit = monthly_profit.sort_values('Month_Num')

    fig2 = px.bar(
        monthly_profit,
        x='Month',
        y='Profit',
        title='Monthly Profit Analysis',
        color='Profit'
    )

    col2.plotly_chart(fig2, use_container_width=True)

    # Scatter Plot
    fig3 = px.scatter(
        filtered_df,
        x='Revenue',
        y='Profit',
        color='Confectionary',
        size='Units_Sold',
        hover_data=['Country'],
        title='Revenue vs Profit Analysis'
    )

    st.plotly_chart(fig3, use_container_width=True)

# ============================================================
# PRODUCT INSIGHTS PAGE
# ============================================================
elif selected == "Product Insights":

    st.title("🍫 Product Insights")

    product_rev = filtered_df.groupby('Confectionary')['Revenue'].sum().reset_index()

    col1, col2 = st.columns(2)

    fig = px.pie(
        product_rev,
        names='Confectionary',
        values='Revenue',
        title='Revenue Share by Product'
    )

    col1.plotly_chart(fig, use_container_width=True)

    fig2 = px.bar(
        product_rev,
        x='Confectionary',
        y='Revenue',
        title='Revenue by Product',
        color='Revenue'
    )

    col2.plotly_chart(fig2, use_container_width=True)

    # Box Plot
    fig3 = px.box(
        filtered_df,
        x='Confectionary',
        y='Profit',
        color='Confectionary',
        title='Profit Distribution by Product'
    )

    st.plotly_chart(fig3, use_container_width=True)

# ============================================================
# COUNTRY ANALYSIS PAGE
# ============================================================
elif selected == "Country Analysis":

    st.title("🌍 Country Performance Analysis")

    country_rev = filtered_df.groupby('Country')['Revenue'].sum().reset_index()

    col1, col2 = st.columns(2)

    fig = px.bar(
        country_rev,
        x='Country',
        y='Revenue',
        title='Revenue by Country',
        color='Revenue'
    )

    col1.plotly_chart(fig, use_container_width=True)

    country_profit = filtered_df.groupby('Country')['Profit'].sum().reset_index()

    fig2 = px.bar(
        country_profit,
        x='Country',
        y='Profit',
        title='Profit by Country',
        color='Profit'
    )

    col2.plotly_chart(fig2, use_container_width=True)

    # Heatmap
    heatmap_data = filtered_df.pivot_table(
        values='Revenue',
        index='Month',
        columns='Country',
        aggfunc='sum'
    )

    fig3 = px.imshow(
        heatmap_data,
        title='Monthly Revenue Heatmap by Country',
        color_continuous_scale='Viridis'
    )

    st.plotly_chart(fig3, use_container_width=True)

# ============================================================
# ADVANCED ANALYTICS PAGE
# ============================================================
elif selected == "Advanced Analytics":

    st.title("📊 Advanced Analytics")

    numeric_df = filtered_df[['Units_Sold', 'Cost', 'Profit', 'Revenue', 'Profit_Margin']]

    corr = numeric_df.corr()

    fig = px.imshow(
        corr,
        text_auto=True,
        title='Correlation Matrix'
    )

    st.plotly_chart(fig, use_container_width=True)

    # Multi Dashboard
    fig2 = make_subplots(
        rows=2, cols=2,
        subplot_titles=(
            'Revenue Trend',
            'Profit Trend',
            'Country Revenue',
            'Product Revenue Share'
        ),
        specs=[
            [{"type": "scatter"}, {"type": "scatter"}],
            [{"type": "bar"}, {"type": "pie"}]
        ]
    )

    daily_rev = filtered_df.groupby('Date')['Revenue'].sum().reset_index()
    daily_profit = filtered_df.groupby('Date')['Profit'].sum().reset_index()

    fig2.add_trace(
        go.Scatter(x=daily_rev['Date'], y=daily_rev['Revenue'], mode='lines'),
        row=1, col=1
    )

    fig2.add_trace(
        go.Scatter(x=daily_profit['Date'], y=daily_profit['Profit'], mode='lines'),
        row=1, col=2
    )

    fig2.add_trace(
        go.Bar(x=country_rev['Country'], y=country_rev['Revenue']),
        row=2, col=1
    )

    fig2.add_trace(
        go.Pie(labels=product_rev['Confectionary'], values=product_rev['Revenue']),
        row=2, col=2
    )

    fig2.update_layout(height=900)

    st.plotly_chart(fig2, use_container_width=True)

# ============================================================
# RAW DATA PAGE
# ============================================================
elif selected == "Raw Data":

    st.title("📋 Raw Dataset Explorer")

    st.write("Filtered Dataset Shape:", filtered_df.shape)

    gb = GridOptionsBuilder.from_dataframe(filtered_df)
    gb.configure_pagination()
    gb.configure_side_bar()
    gb.configure_default_column(groupable=True, value=True, enableRowGroup=True)

    gridOptions = gb.build()

    AgGrid(
        filtered_df,
        gridOptions=gridOptions,
        enable_enterprise_modules=True,
        fit_columns_on_grid_load=True
    )

    # Download Option
    csv = filtered_df.to_csv(index=False).encode('utf-8')

    st.download_button(
        label="Download Filtered Data as CSV",
        data=csv,
        file_name='filtered_confectionary_data.csv',
        mime='text/csv'
    )

# ============================================================
# FOOTER
# ============================================================
st.markdown("---")
st.caption("Developed with Streamlit | Interactive Confectionary Business Intelligence Dashboard")
