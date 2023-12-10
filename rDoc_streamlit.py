import streamlit as st
import pandas as pd
import plotly.express as px
import openpyxl

# Function to load and process the Excel file
@st.cache_resource  # Updated caching function
def load_excel(file):
    # Reading the Excel file with the second row as header
    df = pd.read_excel(file, header=1)
    # Ensure that only numeric data is used in calculations
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

        # Metric selection (excluding 'Unnamed' options and the first column)
        metric_options = df.columns[1:].tolist()  # Assuming first column is for subjects
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
        segment_means = df_filtered.mean(numeric_only=True)

        # Creating the plot
        fig = None
        if plot_type in ['line', 'bar', 'scatter']:
            if plot_type == 'line':
                fig = px.line(x=df.columns[1:], y=segment_means.values, title=f"{selected_metric} over Segments")
            elif plot_type == 'bar':
                fig = px.bar(x=df.columns[1:], y=segment_means.values, title=f"{selected_metric} over Segments")
            elif plot_type == 'scatter':
                fig = px.scatter(x=df.columns[1:], y=segment_means.values, title=f"{selected_metric} over Segments")

            # Update layout for better readability
            fig.update_layout(xaxis_title="Segment", yaxis_title=selected_metric, xaxis={'type':'category'})

            st.plotly_chart(fig)

if __name__ == "__main__":
    main()
