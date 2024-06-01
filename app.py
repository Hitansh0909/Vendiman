import pandas as pd
import streamlit as st
import plotly.express as px
import base64

# Error Handling
try:
    # Load data from Excel with caching
    @st.cache
    def load_data(file_path):
        return pd.read_excel(file_path, sheet_name='Sheet1', header=2)

    data = load_data('Book1.xlsx')

    # Sidebar for filtering
    st.sidebar.header('Filters')

    # Create dropdown menus for filtering
    selected_segment = st.sidebar.selectbox('Select Client Segment:', ['All'] + data['Client Segment (Client Sector)'].unique().tolist(), key='segment_filter')
    selected_account_type = st.sidebar.selectbox('Select Account Type:', ['All'] + data['Account Type'].unique().tolist(), key='account_type_filter')
    selected_commercial_type = st.sidebar.selectbox('Select Commercial Type:', ['All'] + data['Commercial Type'].unique().tolist(), key='commercial_type_filter')

    # Multi-select filter for Client Segment
    selected_segments = st.sidebar.multiselect('Select Client Segments:', data['Client Segment (Client Sector)'].unique().tolist())
    if selected_segments:
        filtered_data = data[data['Client Segment (Client Sector)'].isin(selected_segments)]
    else:
        filtered_data = data.copy()

    # Filter data based on selections
    if selected_segment != 'All':
        filtered_data = filtered_data[filtered_data['Client Segment (Client Sector)'] == selected_segment]
    if selected_account_type != 'All':
        filtered_data = filtered_data[filtered_data['Account Type'] == selected_account_type]
    if selected_commercial_type != 'All':
        filtered_data = filtered_data[filtered_data['Commercial Type'] == selected_commercial_type]

    # Machine Status Distribution (Stacked Bar Chart)
    st.subheader('Machine Status Distribution')
    machine_status_counts = filtered_data['Machine Status'].value_counts()
    st.bar_chart(machine_status_counts)

    # Monthly Service Charge Trend (Line Chart)
    if 'Monthly Service Charge' in filtered_data.columns:  # Check if column exists
        st.subheader('Monthly Service Charge Trend')
        service_charge_trend = filtered_data.groupby('Client Segment (Client Sector)')['Monthly Service Charge'].mean().reset_index()
        fig = px.line(service_charge_trend, x='Client Segment (Client Sector)', y='Monthly Service Charge')
        st.plotly_chart(fig)
    else:
        st.warning("Monthly Service Charge column not found. Skipping line chart.")

    st.header('Client Health and Activity')

    # Sidebar for additional filtering
    selected_client_sub_name = st.sidebar.selectbox('Select Client Sub Name:', ['All'] + data['Client Sub Name'].unique().tolist(), key='client_sub_name')
    selected_agreement_status = st.sidebar.selectbox('Select Agreement Status:', ['All'] + data['Agreement Status'].unique().tolist(), key='agreement_status')

    # Filter data based on additional selections
    if selected_client_sub_name != 'All':
        filtered_data = filtered_data[filtered_data['Client Sub Name'] == selected_client_sub_name]
    if selected_agreement_status != 'All':
        filtered_data = filtered_data[filtered_data['Agreement Status'] == selected_agreement_status]

    # Credit Terms vs. Service Charge (Scatter Plot)
    if 'Customer Credit Terms in Days' in filtered_data.columns and 'Monthly Service Charge' in filtered_data.columns:
        st.subheader('Credit Terms vs. Monthly Service Charge')
        fig = px.scatter(filtered_data, x='Customer Credit Terms in Days', y='Monthly Service Charge')
        st.plotly_chart(fig)
    else:
        st.warning("Required columns for scatter plot not found. Skipping.")

    # Agreement Status Distribution (Pie Chart)
    st.subheader('Agreement Status Distribution')
    agreement_status_counts = filtered_data['Agreement Status'].value_counts().reset_index()
    agreement_status_counts.columns = ['Agreement Status', 'Count']
    fig = px.pie(agreement_status_counts, names='Agreement Status', values='Count', title='Agreement Status Distribution')
    st.plotly_chart(fig)

    st.header('Machine Status and Live Status')

    # Sidebar for machine status and live status filtering
    selected_machine_status = st.sidebar.selectbox('Select Machine Status:', ['All'] + data['Machine Status'].unique().tolist(), key='machine_status')
    selected_live_status = st.sidebar.selectbox('Select Live Status as per Remob:', ['All'] + data['Live Status as per Remob'].unique().tolist(), key='live_status')

    # Filter data based on machine status and live status
    if selected_machine_status != 'All':
        filtered_data = filtered_data[filtered_data['Machine Status'] == selected_machine_status]
    if selected_live_status != 'All':
        filtered_data = filtered_data[filtered_data['Live Status as per Remob'] == selected_live_status]

    # Live Status by Machine Status (Stacked Bar Chart)
    st.subheader('Live Status Distribution by Machine Status')
    live_status_counts = filtered_data['Live Status as per Remob'].value_counts().reset_index()
    live_status_counts.columns = ['Live Status as per Remob', 'Count']
    fig = px.bar(live_status_counts, x='Live Status as per Remob', y='Count')
    st.plotly_chart(fig)

    # Activated vs. Deactivated Date Trend (Line Chart)
    if 'Activated Date' in filtered_data.columns and 'Deactivated Date' in filtered_data.columns:
        st.subheader('Activation vs. Deactivation Trend')
        filtered_data['Activated Date'] = pd.to_datetime(filtered_data['Activated Date'], errors='coerce')
        filtered_data['Deactivated Date'] = pd.to_datetime(filtered_data['Deactivated Date'], errors='coerce')
        filtered_data['Activation Year'] = filtered_data['Activated Date'].dt.year
        filtered_data['Deactivation Year'] = filtered_data['Deactivated Date'].dt.year.fillna(0)

        # Group by activation and deactivation years
        activation_deactivation_trend = filtered_data.groupby(['Activation Year', 'Deactivation Year']).size().reset_index(name='Count')

        # Plot the Activation vs. Deactivation Trend
        if not activation_deactivation_trend.empty:
            st.subheader('Activation vs. Deactivation Trend')
            fig = px.line(activation_deactivation_trend, x='Activation Year', y='Count', color='Deactivation Year', title='Activation vs. Deactivation Trend')
            st.plotly_chart(fig)
        else:
            st.warning("No data available for Activation vs. Deactivation Trend.")

    # Monthly Service Charge Distribution (Box Plot)
    st.subheader('Monthly Service Charge Distribution')
    fig = px.box(filtered_data, x='Client Segment (Client Sector)', y='Monthly Service Charge')
    st.plotly_chart(fig)

    # Summary Statistics
    st.subheader('Summary Statistics')
    st.write(filtered_data.describe())

    # Download Filtered Data
    def get_table_download_link(df):
        csv = df.to_csv(index=False)
        b64 = base64.b64encode(csv.encode()).decode()
        href = f'<a href="data:file/csv;base64,{b64}" download="filtered_data.csv">Download filtered data as CSV</a>'
        return href

    st.markdown(get_table_download_link(filtered_data), unsafe_allow_html=True)

except Exception as e:
    st.error(f"An error occurred: {e}")

# Animated GIF for Security
st.sidebar.subheader("Security")
st.sidebar.markdown("![Security GIF](https://www.google.com/url?sa=i&url=https%3A%2F%2Fdribbble.com%2Fshots%2F1470721-DATA-SECURITY-GIF&psig=AOvVaw2WCHZHwktUYa5nNxtHppyk&ust=1717328831711000&source=images&cd=vfe&opi=89978449&ved=0CBEQjRxqFwoTCIDbssmquoYDFQAAAAAdAAAAABAJ)")

# Display the entire DataFrame at the end of the code without filters
st.subheader("Entire DataFrame")
st.write(data)
