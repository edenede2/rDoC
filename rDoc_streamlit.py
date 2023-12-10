import streamlit as st
import pandas as pd
import plotly.express as px
import openpyxl

# Function to load and process the Excel file
@st.cache_resource  # Updated caching function
def load_excel(file):
    # Assuming the first row contains the metrics and the second row contains segment names
    df = pd.read_excel(file, header=[0, 1])
    df.columns = ['_'.join(col).strip() for col in df.columns.values]
    # Dropping any columns that are 'Unnamed'
    df = df.loc[:, ~df.columns.str.contains('^Unnamed')]
    return df

# Streamlit app
def main():
    st.title("Excel Data Analysis App")

    # Upload Excel file
    uploaded_file = st.file_uploader("Upload your Excel file", type=["xlsx"])
    if uploaded_file is not None:
        df = load_excel(uploaded_file)
        st.write("DataFrame Preview:")
        st.write(df)

        # Metric selection (excluding 'Unnamed' options)
        metric_options = [col for col in df.columns if 'Unnamed' not in col and col != df.columns[0]]
        selected_metric = st.selectbox("Select a Metric", metric_options)

        # Plot type selection
        plot_types = ['line', 'bar', 'scatter']
        plot_type = st.selectbox("Select Plot Type", plot_types)

        # Subject exclusion or isolation
        subjects = df.iloc[:, 0].unique()
        exclude_subjects = st.multiselect("Exclude Subjects", subjects)
        isolate_subject = st.selectbox("Or, Isolate a Single Subject (overrides exclusion)", ['None'] + list(subjects))

        # Data processing for visualization
        if isolate_subject != 'None':
            df_filtered = df[df.iloc[:, 0] == isolate_subject]
        else:
            df_filtered = df if not exclude_subjects else df[~df.iloc[:, 0].isin(exclude_subjects)]
        
        # Calculating mean across all subjects for each segment
        segment_means = df_filtered.mean()

        # Extracting the relevant segment means for the selected metric
        relevant_means = segment_means[[col for col in segment_means.index if selected_metric in col]]

        # Display graph
        fig = None
        if plot_type == 'line':
            fig = px.line(x=relevant_means.index, y=relevant_means.values, title=f"{selected_metric} over Segments")
        elif plot_type == 'bar':
            fig = px.bar(x=relevant_means.index, y=relevant_means.values, title=f"{selected_metric} over Segments")
        elif plot_type == 'scatter':
            fig = px.scatter(x=relevant_means.index, y=relevant_means.values, title=f"{selected_metric} over Segments")

        if fig:
            st.plotly_chart(fig)

if __name__ == "__main__":
    main()
