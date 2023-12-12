import streamlit as st
import pandas as pd
import plotly.express as px
import plotly.graph_objs as go
import openpyxl
import numpy as np
from io import BytesIO
import base64

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
    return outliers

# Function to detect outliers using STD
def detect_outliers_std(df, segment):
    mean = df[segment].mean()
    std = df[segment].std()
    outliers = df[(df[segment] < (mean - 2 * std)) | (df[segment] > (mean + 2 * std))]
    return outliers

def calculate_summary(df, included_segments):
    summary_df = pd.DataFrame()

    # Iterate over the first level of columns (metrics)
    for metric in df.columns.get_level_values(0).unique():
        # Extract the DataFrame for the metric, ensuring it's always a DataFrame
        metric_df = df[[metric]]  # Note the double brackets

        # Check if the segments are in the DataFrame
        valid_segments = [seg for seg in included_segments if seg in metric_df.columns.get_level_values(1)]

        # Calculate summary statistics for the valid segments
        summary_stats = metric_df[metric][valid_segments].agg(['mean', 'std', 'count', 'sem', 'min', 'max'])

        # Add 'out high' and 'out low'
        summary_stats.loc['out high'] = summary_stats['mean'] + 2.5 * summary_stats['std']
        summary_stats.loc['out low'] = summary_stats['mean'] - 2.5 * summary_stats['std']

        # Create a multi-level column
        summary_stats.columns = pd.MultiIndex.from_product([[metric], summary_stats.columns])

        # Append to the summary DataFrame
        summary_df = pd.concat([summary_df, summary_stats], axis=1)

    return summary_df

def download_link(object_to_download, download_filename, download_link_text):
    """
    Generates a link to download the given object_to_download.
    """
    if isinstance(object_to_download, pd.DataFrame):
        # Create a BytesIO object and write the dataframe to it as Excel
        towrite = BytesIO()
        object_to_download.to_excel(towrite, index=False, engine='openpyxl')  # Make sure to use the 'openpyxl' engine
        towrite.seek(0)  # Move back to the beginning after writing
        b64 = base64.b64encode(towrite.read()).decode()  # Read and encode the Excel file
    else:
        b64 = base64.b64encode(object_to_download.encode()).decode()

    return f'<a href="data:application/octet-stream;base64,{b64}" download="{download_filename}">{download_link_text}</a>'

# Streamlit app
def main():
    st.title("Excel Data Analysis App with Outlier Detection and Segment Removal")

    # Upload Excel file
    uploaded_file = st.file_uploader("Upload your Excel file", type=["xlsx"])
    if uploaded_file is not None:
        df = load_excel(uploaded_file)
        st.write("DataFrame Preview:")
        st.write(df.head())

        # Metric selection
        metrics = df.columns.get_level_values(0).unique()
        selected_metric = st.selectbox("Select a Metric", metrics)

        # Segment names for the selected metric and option to exclude specific segments
        all_segments = df[selected_metric].columns.tolist()  # Convert to list
        included_segments = st.multiselect("Include Segments", all_segments, default=all_segments)

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
        outlier_info = []

        if exclude_outliers:
            for segment in included_segments:
                if outlier_method == "IQR":
                    outliers = detect_outliers_iqr(df_filtered, segment)
                elif outlier_method == "STD":
                    outliers = detect_outliers_std(df_filtered, segment)
                
                # Track outlier information
                for index in outliers.index:
                    outlier_info.append((index, segment))

            # Remove outliers from the DataFrame
            outlier_indices = [info[0] for info in outlier_info]
            df_filtered = df_filtered.drop(index=outlier_indices, errors='ignore')

        else:
            for segment in included_segments:
                if outlier_method == "IQR":
                    outliers = detect_outliers_iqr(df_filtered, segment)
                elif outlier_method == "STD":
                    outliers = detect_outliers_std(df_filtered, segment)
                
                # Track outlier information
                for index in outliers.index:
                    outlier_info.append((index, segment))

            # Remove outliers from the DataFrame
            outlier_indices = [info[0] for info in outlier_info]
        
        # Calculating mean values across subjects for each segment
        segment_means = df_filtered[included_segments].mean()
        # Calculating standard deviation for each segment
            # After calculating standard deviation
        segment_se = df_filtered[included_segments].sem()
        st.write("Standard error for each segment:", segment_se) # Debugging line

        # Alternatively, for standard error, you can use:
        # segment_se = df_filtered[included_segments].sem()
        # segment_std = df_filtered[included_segments].std()
        
        
        # Plot type selection with additional options
        plot_types = ['line (with outliers)', 'scatter (individual values)', 'bar']
        plot_type = st.selectbox("Select Plot Type", plot_types)

        # Create the plot based on selected type
        fig = go.Figure()

        if plot_type == 'scatter (individual values)':
            for segment in included_segments:
                for subj in df_filtered.index:
                    fig.add_trace(go.Scatter(x=[segment], y=[df_filtered.loc[subj, segment]], 
                                             mode='markers', name=subj))

        elif plot_type == 'line (with outliers)':
            # Calculating mean values across subjects for each segment
            segment_means = df_filtered[included_segments].mean()

            # Adding the line plot with error bars
            fig.add_trace(go.Scatter(
                x=included_segments, 
                y=segment_means.values, 
                mode='lines', 
                name='Mean',
                error_y=dict(
                    type='percent', # or 'percent' for percentage-based error bars
                    array=segment_se.values, # or segment_se.values for standard error
                    visible=True
                )
            ))
            # Adding outlier lines
            if exclude_outliers:
                for segment in included_segments:
                    outliers = detect_outliers_iqr(df_filtered, segment) if outlier_method == "IQR" \
                               else detect_outliers_std(df_filtered, segment)
                    for index in outliers.index:
                        fig.add_trace(go.Scatter(x=[segment, segment], 
                                                 y=[outliers.loc[index, segment], outliers.loc[index, segment]], 
                                                 mode='lines', marker_color='red', name=f'Outlier - {index}'))

        elif plot_type == 'bar':
            # Calculating mean values across subjects for each segment
            segment_means = df_filtered[included_segments].mean()
            fig = px.bar(x=included_segments, y=segment_means.values, 
                         title=f"{selected_metric} over Segments")

        st.plotly_chart(fig)

        # Displaying removed outliers
        if outlier_info:
            st.write("Outliers:")
            for subject, segment in outlier_info:
                st.write(f"Subject: {subject}, Segment: {segment}")
        
        # Calculate summary statistics after data processing
        metrics = df.columns.get_level_values(0).unique()  # Fetch unique metrics from the DataFrame
        summary_stats = calculate_summary(df_filtered, included_segments)
        
        # Generate download link for the summary dataframe
        tmp_download_link = download_link(summary_stats, 'HRV_Summary.xlsx', 'Download Excel file with Summary')
        st.markdown(tmp_download_link, unsafe_allow_html=True)
        
if __name__ == "__main__":
    main()
