import streamlit as st
import pandas as pd
import plotly.express as px
import openpyxl
import numpy as np

# Function to load and process the Excel file
@st.cache_resource  # Updated caching function
def load_excel(file):
    df = pd.read_excel(file, header=[0, 1], skiprows=[2], index_col=0)
    return df

# Function to detect outliers using IQR
def detect_outliers_iqr(df, segment):
    Q1 = df[segment].quantile(0.25)
    Q3 = df[segment].quantile(0.75)
    IQR = Q3 - Q1
    outliers = df[(df[segment] < (Q1 - 1.5 * IQR)) | (df[segment] > (Q3 + 1.5 * IQR))]
    return outliers.index

# Function to detect outliers using STD
def detect_outliers_std(df, segment):
    mean = df[segment].mean()
    std = df[segment].std()
    outliers = df[(df[segment] < (mean - 2 * std)) | (df[segment] > (mean + 2 * std))]
    return outliers.index

# Streamlit app
def main():
    st.title("Excel Data Analysis App with Outlier Detection")

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

        # Exclude or isolate subjects
        if isolate_subject != 'None':
            df_filtered = df.loc[df.index == isolate_subject, selected_metric]
        else:
            df_filtered = df.loc[~df.index.isin(exclude_subjects), selected_metric]

        # Outlier detection options
        outlier_method = st.selectbox("Select Outlier Detection Method", ["IQR", "STD"])
        exclude_outliers = st.checkbox("Exclude Outliers")
        outlier_indices = set()

        if exclude_outliers:
            for segment in segments:
                if outlier_method == "IQR":
                    outlier_indices.update(detect_outliers_iqr(df_filtered, segment))
                elif outlier_method == "STD":
                    outlier_indices.update(detect_outliers_std(df_filtered, segment))

            df_filtered = df_filtered.drop(index=outlier_indices, errors='ignore')

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

        # Highlight outliers if not excluded
        if not exclude_outliers and outlier_indices:
            for outlier in outlier_indices:
                outlier_data = df.loc[outlier, selected_metric]
                fig.add_trace(px.scatter(x=outlier_data.index, y=outlier_data.values, marker_color='red').data[0])

        st.plotly_chart(fig)

if __name__ == "__main__":
    main()
