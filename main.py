# python3 -m streamlit run main.py

import streamlit as st
import pandas as pd
import re
from pptx import Presentation
import os
import io
import base64

def find_acronyms_in_slide(slide, slide_number):
    acronyms = []
    for shape in slide.shapes:
        if shape.has_text_frame:
            for paragraph in shape.text_frame.paragraphs:
                for run in paragraph.runs:
                    words = run.text.split(' ')
                    for i, word in enumerate(words):
                        if re.match(r'^[A-Z]{2,5}$', word):  # simple rule for acronyms
                            before_words = ' '.join(words[max(0, i-5):i])
                            after_words = ' '.join(words[i+1:min(i+6, len(words))])
                            acronyms.append((word, slide_number, before_words, after_words))
    return acronyms

def read_pptx(file):
    prs = Presentation(file)
    data = []
    for i, slide in enumerate(prs.slides):
        data += find_acronyms_in_slide(slide, i+1)
    return pd.DataFrame(data, columns=['Acronym', 'Slide', 'Before Words', 'After Words'])

st.set_option('deprecation.showfileUploaderEncoding', False)  # disable Streamlit's warning message
uploaded_file = st.file_uploader("Choose a PowerPoint file", type="pptx")
if uploaded_file is not None:
    df = read_pptx(uploaded_file)

    if st.button('Show Data'):
        # Create two columns
        col1, col2 = st.columns(2)

        # In the first column, display count of each unique acronym and total count of acronyms
        with col1:
            st.write("Count of Each Unique Acronym:")
            st.write(df['Acronym'].value_counts().to_frame())

            st.write(f"Total count of acronyms: {df['Acronym'].count()}")

        # In the second column, show the dataframe
        with col2:
            st.header("Table of Acronyms:")
            st.write(df)

    if st.button('Export to CSV'):
        csv = df.to_csv(index=False)
        b64 = base64.b64encode(csv.encode()).decode()  # some strings
        href = f'<a href="data:file/csv;base64,{b64}" download="acronyms.csv">Download CSV File</a>'
        st.markdown(href, unsafe_allow_html=True)

