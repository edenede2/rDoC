import streamlit as st
import pandas as pd
import plotly.express as px
import openpyxl

# Function to load and process the Excel file
@st.cache_resource  # Updated caching function
def load_excel(file):
    df = pd.read_excel(file, header=0)
    # Drop the row with NaN values which seems to be the second row in the file
    df = df.drop(1)
    # Reshape the DataFrame for easier processing
    df = df.melt(id_vars=[df.columns[0]], var_name='Metric_Segment', value_name='Value')
    df[['Metric', 'Segment']] = df['Metric_Segment'].str.split(' ', expand=True)
    df = df.drop('Metric_Segment', axis=1)
    # Convert values to numeric, errors coerced to NaN
    df['Value'] = pd.to_numeric(df['Value'], errors='coerce')
    return df

# Streamlit app
def main():
    st.title("Excel Data Analysis App")

    # Upload Excel file
    uploaded_file = st.file_uploader("Upload your Excel file", type=["xlsx"])
    if uploaded_file is not None:
        df = load_excel(uploaded_file)
        st.write("DataFrame Preview:")
        st.write(df.head())

        # Metric selection
        metrics = df['Metric'].unique()
        selected_metric = st.selectbox("Select a Metric", metrics)

        # Filter DataFrame for the selected metric
        df_filtered = df[df['Metric'] == selected_metric]

        # Subject exclusion or isolation
        subjects = df[df.columns[0]].unique()
        exclude_subjects = st.multiselect("Exclude Subjects", subjects)
        isolate_subject = st.selectbox("Or, Isolate a Single Subject (overrides exclusion)", ['None'] + list(subjects))

        if isolate_subject != 'None':
            df_filtered = df_filtered[df_filtered[df.columns[0]] == isolate_subject]
        elif exclude_subjects:
            df_filtered = df_filtered[~df_filtered[df.columns[0]].isin(exclude_subjects)]

        # Calculating mean values across subjects for each segment
        segment_means = df_filtered.groupby('Segment')['Value'].mean().reset_index()

        # Plot type selection
        plot_types = ['line', 'bar', 'scatter']
        plot_type = st.selectbox("Select Plot Type", plot_types)

        # Create the plot
        if plot_type == 'line':
            fig = px.line(segment_means, x='Segment', y='Value', title=f"{selected_metric} over Segments")
        elif plot_type == 'bar':
            fig = px.bar(segment_means, x='Segment', y='Value', title=f"{selected_metric} over Segments")
        elif plot_type == 'scatter':
            fig = px.scatter(segment_means, x='Segment', y='Value', title=f"{selected_metric} over Segments")

        st.plotly_chart(fig)

if __name__ == "__main__":
    main()
