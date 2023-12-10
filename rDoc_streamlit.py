import streamlit as st
import pandas as pd
import plotly.express as px
import openpyxl

# Function to load and process the Excel file
@st.cache_resource  # Updated caching function
def load_excel(file):
    # Read the file, skipping the first three rows and setting the first column as index
    df = pd.read_excel(file, header=[0, 1], skiprows=[2], index_col=0)
    # Create a multi-index using the metric and segment names
    df.columns = pd.MultiIndex.from_tuples([tuple(c.split()) for c in df.columns])
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
        metrics = df.columns.get_level_values(0).unique()
        selected_metric = st.selectbox("Select a Metric", metrics)

        # Segment names for the selected metric
        segments = df[selected_metric].columns

        # Subject exclusion or isolation
        subjects = df.index.unique()
        exclude_subjects = st.multiselect("Exclude Subjects", subjects)
        isolate_subject = st.selectbox("Or, Isolate a Single Subject (overrides exclusion)", ['None'] + list(subjects))

        # Filtering DataFrame based on subject selection
        if isolate_subject != 'None':
            df_filtered = df.loc[df.index == isolate_subject, selected_metric]
        else:
            df_filtered = df.loc[~df.index.isin(exclude_subjects), selected_metric]

        # Calculating mean values across subjects for each segment
        segment_means = df_filtered.mean()

        # Plot type selection
        plot_types = ['line', 'bar', 'scatter']
        plot_type = st.selectbox("Select Plot Type", plot_types)

        # Create the plot
        fig = None
        if plot_type == 'line':
            fig = px.line(x=segments, y=segment_means.values, title=f"{selected_metric} over Segments")
        elif plot_type == 'bar':
            fig = px.bar(x=segments, y=segment_means.values, title=f"{selected_metric} over Segments")
        elif plot_type == 'scatter':
            fig = px.scatter(x=segments, y=segment_means.values, title=f"{selected_metric} over Segments")

        st.plotly_chart(fig)

if __name__ == "__main__":
    main()
