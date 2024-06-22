import streamlit as st
import pandas as pd
import plotly.express as px
import matplotlib.pyplot as plt
import seaborn as sns

# Function to load data
@st.cache
def load_data():
    data = pd.read_excel("superstore_sales.xlsx")
    return data

# Function to preprocess data (if needed)
def preprocess_data(data):
    # Example preprocessing steps if required
    data['order_date'] = pd.to_datetime(data['order_date'])  # Convert to datetime if needed
    data['month_year'] = data['order_date'].dt.strftime('%Y-%m')  # Extract month-year
    return data

# Function to plot overall sales trend
def plot_overall_sales_trend(data):
    sales_by_month = data.groupby('month_year')['sales'].sum().reset_index()
    fig = px.line(sales_by_month, x='month_year', y='sales', title='Overall Sales Trend')
    return fig

# Function to plot sales by category
def plot_sales_by_category(data):
    sales_by_category = data.groupby('category')['sales'].sum().reset_index()
    fig = px.pie(sales_by_category, values='sales', names='category', 
                 title='Sales Analysis by Category', hole=0.5,
                 color_discrete_sequence=px.colors.qualitative.Pastel)
    fig.update_traces(textposition='inside', textinfo='percent+label')
    return fig

# Function to plot sales by sub-category
def plot_sales_by_subcategory(data):
    sales_by_subcategory = data.groupby('sub_category')['sales'].sum().reset_index()
    fig = px.bar(sales_by_subcategory, x='sub_category', y='sales', 
                 title='Sales Analysis by Sub-Category')
    return fig

# Function to plot monthly profits
def plot_monthly_profits(data):
    profit_by_month = data.groupby('month_year')['profit'].sum().reset_index()
    fig = px.line(profit_by_month, x='month_year', y='profit', title='Monthly Profit Analysis')
    return fig

# Function to plot profit by category
def plot_profit_by_category(data):
    profit_by_category = data.groupby('category')['profit'].sum().reset_index()
    fig = px.pie(profit_by_category, values='profit', names='category', 
                 title='Profit Analysis by Category', hole=0.5,
                 color_discrete_sequence=px.colors.qualitative.Pastel)
    fig.update_traces(textposition='inside', textinfo='percent+label')
    return fig

# Function to plot profit by sub-category
def plot_profit_by_subcategory(data):
    profit_by_subcategory = data.groupby('sub_category')['profit'].sum().reset_index()
    fig = px.bar(profit_by_subcategory, x='sub_category', y='profit', 
                 title='Profit Analysis by Sub-Category')
    return fig

# Function to plot sales and profit analysis by customer segment
def plot_sales_profit_by_segment(data):
    sales_profit_by_segment = data.groupby('segment').agg({'sales': 'sum', 'profit': 'sum'}).reset_index()
    color_palette = px.colors.qualitative.Pastel
    fig = px.bar(sales_profit_by_segment, x='segment', y=['sales', 'profit'], 
                 title='Sales and Profit Analysis by Customer Segment', 
                 barmode='group', color_discrete_sequence=color_palette)
    fig.update_layout(xaxis_title='Customer Segment', yaxis_title='Amount')
    return fig

# Function to display sales to profit ratio by segment
def display_sales_to_profit_ratio(data):
    sales_profit_by_segment = data.groupby('segment').agg({'sales': 'sum', 'profit': 'sum'}).reset_index()
    sales_profit_by_segment['Sales_to_Profit_Ratio'] = sales_profit_by_segment['sales'] / sales_profit_by_segment['profit']
    st.write("Sales to Profit Ratio by Customer Segment:")
    st.dataframe(sales_profit_by_segment[['segment', 'Sales_to_Profit_Ratio']])

# Function to display top 10 products by sales
def display_top_10_products_by_sales(data):
    prod_sales = data.groupby('product_name')['sales'].sum().reset_index()
    prod_sales.sort_values(by='sales', ascending=False, inplace=True)
    top_10_products = prod_sales.head(10)
    st.write("Top 10 Products by Sales:")
    st.dataframe(top_10_products)

# Function to display most selling products
def display_most_selling_products(data):
    best_selling_prods = data.groupby('product_name')['quantity'].sum().reset_index()
    best_selling_prods.sort_values(by='quantity', ascending=False, inplace=True)
    most_selling_products = best_selling_prods.head(10)
    st.write("Most Selling Products:")
    st.dataframe(most_selling_products)

# Function to display most preferred ship mode
def display_most_preferred_ship_mode(data):
    st.title("Most Preferred Ship Mode")
    plt.figure(figsize=(10, 6))
    sns.countplot(x='ship_mode', data=data)
    plt.xlabel('Ship Mode')
    plt.ylabel('Count')
    plt.xticks(rotation=45)
    st.pyplot()

# Function to display most profitable category and sub-category
def display_most_profitable_category_subcategory(data):
    cat_subcat = data.groupby(['category', 'sub_category'])['profit'].sum().reset_index()
    cat_subcat.sort_values(by='profit', ascending=False, inplace=True)
    most_profitable = cat_subcat.head(10)
    st.write("Most Profitable Category and Sub-Category:")
    st.dataframe(most_profitable)

def main():
    st.title("Superstore Sales Analysis Dashboard")

    # Load data
    data = load_data()

    if data.empty:
        st.error("Failed to load data. Please check the data source.")
        return

    # Preprocess data (if needed)
    data = preprocess_data(data)

    # Sidebar navigation
    st.sidebar.title('Dashboard Navigation')
    sections = ['Overall Sales Trend', 'Sales by Category', 'Sales by Sub-Category', 
                'Monthly Profits', 'Profit by Category', 'Profit by Sub-Category', 
                'Sales and Profit by Customer Segment', 'Sales to Profit Ratio', 
                'Top 10 Products by Sales', 'Most Selling Products', 'Most Preferred Ship Mode', 
                'Most Profitable Category and Sub-Category']

    section = st.sidebar.radio('Go to', sections)

    if section == 'Overall Sales Trend':
        st.subheader('Overall Sales Trend')
        fig = plot_overall_sales_trend(data)
        st.plotly_chart(fig)

    elif section == 'Sales by Category':
        st.subheader('Sales by Category')
        fig = plot_sales_by_category(data)
        st.plotly_chart(fig)

    elif section == 'Sales by Sub-Category':
        st.subheader('Sales by Sub-Category')
        fig = plot_sales_by_subcategory(data)
        st.plotly_chart(fig)

    elif section == 'Monthly Profits':
        st.subheader('Monthly Profits')
        fig = plot_monthly_profits(data)
        st.plotly_chart(fig)

    elif section == 'Profit by Category':
        st.subheader('Profit by Category')
        fig = plot_profit_by_category(data)
        st.plotly_chart(fig)

    elif section == 'Profit by Sub-Category':
        st.subheader('Profit by Sub-Category')
        fig = plot_profit_by_subcategory(data)
        st.plotly_chart(fig)

    elif section == 'Sales and Profit by Customer Segment':
        st.subheader('Sales and Profit by Customer Segment')
        fig = plot_sales_profit_by_segment(data)
        st.plotly_chart(fig)

    elif section == 'Sales to Profit Ratio':
        st.subheader('Sales to Profit Ratio')
        display_sales_to_profit_ratio(data)

    elif section == 'Top 10 Products by Sales':
        st.subheader('Top 10 Products by Sales')
        display_top_10_products_by_sales(data)

    elif section == 'Most Selling Products':
        st.subheader('Most Selling Products')
        display_most_selling_products(data)

    elif section == 'Most Preferred Ship Mode':
        st.subheader('Most Preferred Ship Mode')
        display_most_preferred_ship_mode(data)

    elif section == 'Most Profitable Category and Sub-Category':
        st.subheader('Most Profitable Category and Sub-Category')
        display_most_profitable_category_subcategory(data)

if __name__ == "__main__":
    main()
