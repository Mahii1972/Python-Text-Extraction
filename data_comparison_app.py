import streamlit as st
import pandas as pd
import re
import pdfplumber
from bs4 import BeautifulSoup

# Function to clean phone numbers
def clean_phone_number(phone):
    return re.sub(r'\D', '', str(phone).split('.')[0])

# Function to extract text from PDF
def extract_text_from_pdf(uploaded_file):
    text = ''
    with pdfplumber.open(uploaded_file) as pdf:
        for page in pdf.pages:
            text += page.extract_text()
    return text

# Function to extract text from TXT
def extract_text_from_txt(uploaded_file):
    text = uploaded_file.read().decode("utf-8") 
    return text

# Function to extract text from HTML
def extract_text_from_html(uploaded_file):
    soup = BeautifulSoup(uploaded_file, 'html.parser')
    text = soup.get_text()
    return text

# Function to extract text from Excel
def extract_text_from_excel(uploaded_file):
    df = pd.read_excel(uploaded_file)
    return df.to_string()

# Function to compare data
def compare_data(base_df, text):
    results = []
    for index, row in base_df.iterrows():
        phone = clean_phone_number(row['Mobile No']) if pd.notnull(row['Mobile No']) else ''
        email = row['Mail ID'].lower() if pd.notnull(row['Mail ID']) else ''
        name = row['Passenger Name'].lower() if pd.notnull(row['Passenger Name']) else ''
        agency = row['Travel Agency'].lower() if pd.notnull(row['Travel Agency']) else ''

        phone_match = re.search(phone, text) if phone else None
        email_match = re.search(re.escape(email), text, re.IGNORECASE) if email else None
        name_match = re.search(re.escape(name), text, re.IGNORECASE) if name else None
        agency_match = re.search(re.escape(agency), text, re.IGNORECASE) if agency else None

        if phone_match:
            context_start = max(phone_match.start() - 10, 0)
            context_end = phone_match.end() + 10
            context = text[context_start:context_end]
            results.append(['Phone', phone_match.group(), context])
        if email_match:
            context_start = max(email_match.start() - 10, 0)
            context_end = email_match.end() + 10
            context = text[context_start:context_end]
            results.append(['Email', email_match.group(), context])
        if name_match:
            context_start = max(name_match.start() - 10, 0)
            context_end = name_match.end() + 10
            context = text[context_start:context_end]
            results.append(['Name', name_match.group(), context])
        if agency_match:
            context_start = max(agency_match.start() - 10, 0)
            context_end = agency_match.end() + 10
            context = text[context_start:context_end]
            results.append(['Agency', agency_match.group(), context])

    return pd.DataFrame(results, columns=['Match Type', 'Match String', 'File Context'])

# Streamlit App
st.title("Data Comparison App")

# File Uploads
st.sidebar.header("Upload Files")
base_file = st.sidebar.file_uploader("Upload Base Excel File", type=['xlsx', 'xls'])
compare_file = st.sidebar.file_uploader("Upload File to Compare", type=['pdf', 'txt', 'htm', 'html', 'xlsx', 'xls'])

if base_file and compare_file:
    try:
        base_df = pd.read_excel(base_file)
        base_df['Mobile No'] = base_df['Mobile No'].apply(clean_phone_number)
        base_df['Mail ID'] = base_df['Mail ID'].str.lower()
        base_df['Passenger Name'] = base_df['Passenger Name'].str.lower()
        base_df['Travel Agency'] = base_df['Travel Agency'].str.lower()

        file_extension = compare_file.name.split('.')[-1].lower()

        # Extract text based on file type
        if file_extension == 'pdf':
            compare_text = extract_text_from_pdf(compare_file)
        elif file_extension == 'txt':
            compare_text = extract_text_from_txt(compare_file)
        elif file_extension in ['htm', 'html']:
            compare_text = extract_text_from_html(compare_file)
        elif file_extension in ['xls', 'xlsx']:
            compare_text = extract_text_from_excel(compare_file)
        else:
            st.error("Unsupported file type")

        # Perform Comparison
        if compare_text:
            comparison_results = compare_data(base_df, compare_text)

            # Display Results
            if not comparison_results.empty:
                st.subheader("Comparison Results:")
                st.dataframe(comparison_results)
            else:
                st.info("No matches found.")
    except Exception as e:
        st.error(f"An error occurred: {e}")