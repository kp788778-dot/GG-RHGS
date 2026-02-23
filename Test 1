import streamlit as st
import pandas as pd

st.title("Simple Data Explorer")

# Upload CSV
uploaded_file = st.file_uploader("Upload a CSV file", type=["csv"])

if uploaded_file is not None:
  df = pd.read_csv(uploaded_file)
  st.write("Here’s a preview of your data:")
  st.dataframe(df.head())

  # Add a chart
  st.line_chart(df.select_dtypes(include=["number"]))
