import streamlit as st
import pandas as pd
import plotly.express as px
import openpyxl

# Function to load and process the Excel file
@st.cache_resource  # Updated caching function
def load_excel(file):
    # Reading the Excel file and using the first row for multi-index columns
    df = pd.read_excel(file, header=[0, 1])
    # Flatten the MultiIndex for easier processing
    df.columns = ['_'.join(col).strip() for col in df.columns.values]
    # Convert all relevant columns to numeric, ignoring errors to skip non-numeric columns
    for col in df.columns[1:]:  # Assuming first column is not numeric
        df[col] = pd.to_numeric(df[col], errors='coerce')
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

        # Extract metric names (assuming they are prefixes in the column names)
        metric_names = list(set(col.split('_')[0] for col in df.columns if '_' in col))
        selected_metric = st.selectbox("Select a Metric", metric_names)

        # Extract segment names for the selected metric
        segments = [col for col in df.columns if col.startswith(selected_metric)]

        # Plot type selection
        plot_types = ['line', 'bar', 'scatter']
        plot_type = st.selectbox("Select Plot Type", plot_types)

        # Subject exclusion or isolation
        subjects = df['Subject'].unique()
        exclude_subjects = st.multiselect("Exclude Subjects", subjects)
        isolate_subject = st.selectbox("Or, Isolate a Single Subject (overrides exclusion)", ['None'] + list(subjects))

        # Data processing for visualization
        if isolate_subject != 'None':
            df_filtered = df[df['Subject'] == isolate_subject]
        else:
            df_filtered = df if not exclude_subjects else df[~df['Subject'].isin(exclude_subjects)]

        # Calculating mean across all subjects for each segment under the selected metric
        segment_means = df_filtered[segments].mean()

        # Create the plot
        fig = None
        if plot_type in ['line', 'bar', 'scatter']:
            segment_labels = [seg.split('_')[1] for seg in segments]  # Extracting segment labels
            if plot_type == 'line':
                fig = px.line(x=segment_labels, y=segment_means.values, title=f"{selected_metric} over Segments")
            elif plot_type == 'bar':
                fig = px.bar(x=segment_labels, y=segment_means.values, title=f"{selected_metric} over Segments")
            elif plot_type == 'scatter':
                fig = px.scatter(x=segment_labels, y=segment_means.values, title=f"{selected_metric} over Segments")

            # Update layout for better readability
            fig.update_layout(xaxis_title="Segment", yaxis_title=selected_metric, xaxis={'type':'category'})

            st.plotly_chart(fig)

if __name__ == "__main__":
    main()
