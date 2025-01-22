import streamlit as st
from pathlib import Path
from components.sidebar import render_sidebar
from modules.file_handler import read_excel_file
from pages import topic_modeling, word_cloud, sentiment_analysis

st.set_page_config(layout="wide")

css_path = Path(__file__).parent / 'styles.css'
with open(css_path) as f:
    st.markdown(f'<style>{f.read()}</style>', unsafe_allow_html=True)

uploaded_file, topic, cloud, sentiment = render_sidebar()

if uploaded_file:
    df = read_excel_file(uploaded_file)
    st.dataframe(df.head(5).select(df.columns[:5]))

if topic:
    topic_modeling.show_topic_modeling()
elif cloud:
    word_cloud.show_word_cloud()
elif sentiment:
    sentiment_analysis.show_sentiment_analysis()
else:
    st.header("Welcome")
    st.write("Select an option from the sidebar")