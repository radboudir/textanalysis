import streamlit as st

def render_sidebar():
    with st.sidebar:
        st.title("Menu")
        uploaded_file = st.file_uploader("Upload File", type=["xlsx", "xls"], 
                                       key="file_uploader",
                                       label_visibility="hidden")
        if uploaded_file:
            st.write(f"Filename: {uploaded_file.name}")
            st.write(f"Size: {uploaded_file.size} bytes")
            
        topic = st.button("Topic Modeling")
        cloud = st.button("Word Cloud")
        sentiment = st.button("Sentiment Analysissss")
        
        return uploaded_file, topic, cloud, sentiment
