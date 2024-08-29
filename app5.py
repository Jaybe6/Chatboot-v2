import pandas as pd
import streamlit as st
import matplotlib.pyplot as plt
import seaborn as sns


# Function to load data
def load_data(file_path):
    try:
        df = pd.read_excel(file_path)
        return df
    except FileNotFoundError:
        st.error(f"The file was not found at {file_path}")
        return None
    except Exception as e:
        st.error(f"An error occurred: {e}")
        return None


# Function to handle user queries
def handle_query(df, query):
    query = query.lower()

    # Define keyword groups
    top_keywords = ['top', 'best', 'highest', 'top 5']
    bottom_keywords = ['bottom', 'worst', 'lowest', 'bottom 5']
    company_keywords = ['company', 'companies']
    revenue_keywords = ['revenue', 'sales']
    profit_keywords = ['profit', 'income', 'earnings']
    count_keywords = ['how many', 'number of', 'count']

    # Handling queries for the top 5 companies by revenue or profit
    if any(keyword in query for keyword in top_keywords) and any(keyword in query for keyword in company_keywords):
        if any(keyword in query for keyword in revenue_keywords):
            top_5_revenue = df.nlargest(5, 'Revenue')[['Company', 'Revenue']]
            st.write("Top 5 companies by revenue:")
            st.write(top_5_revenue)
        elif any(keyword in query for keyword in profit_keywords):
            top_5_profit = df.nlargest(5, 'Profit')[['Company', 'Profit']]
            st.write("Top 5 companies by profit:")
            st.write(top_5_profit)
        else:
            # Default to top 5 by revenue if no specific measure is mentioned
            top_5_revenue = df.nlargest(5, 'Revenue')[['Company', 'Revenue']]
            st.write("Top 5 companies by revenue:")
            st.write(top_5_revenue)
        return True

    # Handling queries for the bottom 5 companies by revenue or profit
    elif any(keyword in query for keyword in bottom_keywords) and any(keyword in query for keyword in company_keywords):
        if any(keyword in query for keyword in revenue_keywords):
            bottom_5_revenue = df.nsmallest(5, 'Revenue')[['Company', 'Revenue']]
            st.write("Bottom 5 companies by revenue:")
            st.write(bottom_5_revenue)
        elif any(keyword in query for keyword in profit_keywords):
            bottom_5_profit = df.nsmallest(5, 'Profit')[['Company', 'Profit']]
            st.write("Bottom 5 companies by profit:")
            st.write(bottom_5_profit)
        else:
            # Default to bottom 5 by revenue if no specific measure is mentioned
            bottom_5_revenue = df.nsmallest(5, 'Revenue')[['Company', 'Revenue']]
            st.write("Bottom 5 companies by revenue:")
            st.write(bottom_5_revenue)
        return True

    # Handling queries for total number of companies
    elif any(keyword in query for keyword in count_keywords) and any(keyword in query for keyword in company_keywords):
        num_companies = df['Company'].nunique()
        st.write(f"Total number of companies: {num_companies}")
        return True

    # Handling queries for highest profit
    elif any(keyword in query for keyword in ['highest', 'maximum', 'max']) and 'profit' in query:
        highest_profit = df.loc[df['Profit'].idxmax()][['Company', 'Profit']]
        st.write("Company with the highest profit:")
        st.write(highest_profit)
        return True

    # If the query contains a specific company name
    else:
        company_name = query.strip().title()  # Capitalize each word for consistent matching
        if company_name in df['Company'].values:
            display_company_info(df, company_name)
            plot_charts(df, company_name)
            return True

    # If no matching query is found
    return False


# Function to display company information
def display_company_info(df, company_name):
    company_data = df[df['Company'] == company_name]
    if company_data.empty:
        st.write(f"<span style='color: white;'>No data found for company: {company_name}</span>",
                 unsafe_allow_html=True)
        return None
    else:
        st.write(f"<span style='color: white;'>Data for company: {company_name}</span>", unsafe_allow_html=True)
        st.write(company_data)
        return company_data


# Function to plot charts
def plot_charts(df, company_name):
    # Filter data for specific company
    company_data = df[df['Company'] == company_name]

    if company_data.empty:
        st.write("<span style='color: white;'>No data to display for the specified company.</span>",
                 unsafe_allow_html=True)
        return

    # Remove duplicate entries for charting
    unique_df = df.drop_duplicates(subset=['Company'])

    # Set a black background for the plots
    plt.style.use('dark_background')

    # Bar Chart: User Input vs. Top 5 Companies by Revenue
    st.subheader('Revenue Model')
    top_5 = unique_df.nlargest(5, 'Revenue')
    top_5_and_user = pd.concat([top_5, company_data]).drop_duplicates(subset=['Company'])

    plt.figure(figsize=(16, 8))  # Increase the size of the bar chart
    bar_plot = sns.barplot(x='Company', y='Revenue', data=top_5_and_user, palette='viridis')
    plt.title('Revenue Model', color='white', fontsize=20)
    plt.xticks(rotation=45, ha='right', color='white')  # Rotate labels for better readability
    plt.tight_layout()

    # Decrease the width of the columns and increase the spacing
    for index, patch in enumerate(bar_plot.patches):
        patch.set_width(0.6)  # Adjust width of bars
        patch.set_edgecolor('black')  # Add black edge color to bars

    # Add data labels to the bar chart with increased width and bold font
    for p in bar_plot.patches:
        bar_plot.annotate(
            format(p.get_height(), '.0f'),
            (p.get_x() + p.get_width() / 2., p.get_height()),
            ha='center', va='center',
            xytext=(0, 10),
            textcoords='offset points',
            color='white',
            fontsize=12,
            fontweight='bold'
        )

    st.pyplot(plt)

    # Pie Chart: User Input vs. Top 5 Companies by Profit
    st.subheader('Market Share')
    top_5_profit = unique_df.nlargest(5, 'Profit')
    top_5_profit_and_user = pd.concat([top_5_profit, company_data]).drop_duplicates(subset=['Company'])

    plt.figure(figsize=(8, 8))  # Increase the size of the pie chart
    plt.pie(top_5_profit_and_user['Profit'], labels=top_5_profit_and_user['Company'], autopct='%1.1f%%',
            colors=sns.color_palette("dark:#5A9_r", len(top_5_profit_and_user)))
    plt.title('Market Share', color='white', fontsize=20)
    plt.tight_layout()
    st.pyplot(plt)

    # Histogram: Distribution of Profit
    st.subheader('Profit Distribution')
    plt.figure(figsize=(16, 8))  # Increase the size of the histogram
    bins = 20
    counts, bin_edges, patches = plt.hist(df['Profit'], bins=bins, color='skyblue', edgecolor='black', alpha=0.7,
                                          label='Profit Distribution')
    plt.title('Profit Distribution', color='white', fontsize=20)
    plt.xlabel('Profit', color='white')
    plt.ylabel('Frequency', color='white')

    # Add data labels to the histogram
    for count, bin_edge, patch in zip(counts, bin_edges, patches):
        height = patch.get_height()
        plt.text(
            bin_edge + (patch.get_width() / 2), height,
            f'{int(height)}',
            ha='center', va='bottom',
            color='white',
            fontsize=12,
            fontweight='bold'
        )

    plt.tight_layout()
    st.pyplot(plt)


# Streamlit app
def main():
    st.title('Excel Data Insights Chatbot')

    # Set background color for Streamlit dashboard
    st.markdown(
        """
        <style>
        .main {
            background-color: black;
        }
        .stButton button {
            background-color: #262730;
            color: white;
        }
        .stTextInput input {
            background-color: #262730;
            color: white;
        }
        .stTextInput label {
            color: white;
        }
        .stSubheader {
            color: #FFD700; /* Gold color for better visibility on black */
        }
        </style>
        """,
        unsafe_allow_html=True
    )

    # File path for the Excel file
    file_path = "C:\\Users\\Joydip\\Desktop\\Data212.xlsx"

    # Load data from the specified file path
