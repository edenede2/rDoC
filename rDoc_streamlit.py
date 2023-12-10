import streamlit as st
import pandas as pd
import plotly.express as px

# Function to load and process the Excel file
@st.cache
def load_excel(file):
    df = pd.read_excel(file)
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

        # Metric selection
        metric_options = df.columns.tolist()[1:]  # Assuming first column is not a metric
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
            df_filtered = df[~df.iloc[:, 0].isin(exclude_subjects)]

        segment_means = df_filtered.groupby(df_filtered.columns[0])[selected_metric].mean()

        # Display graph
        if plot_type == 'line':
            fig = px.line(segment_means, x=segment_means.index, y=selected_metric, title=f"{selected_metric} over Segments")
        elif plot_type == 'bar':
            fig = px.bar(segment_means, x=segment_means.index, y=selected_metric, title=f"{selected_metric} over Segments")
        elif plot_type == 'scatter':
            fig = px.scatter(segment_means, x=segment_means.index, y=selected_metric, title=f"{selected_metric} over Segments")

        st.plotly_chart(fig)

if __name__ == "__main__":
    main()
