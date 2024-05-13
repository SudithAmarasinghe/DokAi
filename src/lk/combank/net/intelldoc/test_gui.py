import streamlit as st
from datetime import datetime

class Sidebar:
    def __init__(self):
        pass
    
    def run(self):
        st.sidebar.header("Select options")
        company_names = {
                "commercial bank": ["2021", "2022", "2023", "2024"],
                "company B": ["2021", "2022", "2023", "2024"],
                "company C": ["2021", "2022", "2023", "2024"],
                "company D": ["2021", "2022", "2023", "2024"],
                "company E": ["2021", "2022", "2023", "2024"]
            }
        selected_option = st.sidebar.radio("Select company type:", ("PLC", "PVT"))
        selected_company = st.sidebar.selectbox("Select company", list(company_names.keys()))
        years = company_names[selected_company]
        start_date_default = datetime(int(years[0]), 1, 1)
        end_date_default = datetime(int(years[-1]), 12, 31)
        start_date = st.sidebar.date_input("Select start date", start_date_default)
        end_date = st.sidebar.date_input("Select end date", end_date_default)

        if end_date < start_date:
            st.sidebar.error("End date cannot be before start date. Please select a valid end date.")
        else:
            uploaded_file_placeholder = st.sidebar.empty()
            uploaded_file = uploaded_file_placeholder.file_uploader("Upload a PDF", type=["pdf"])
            return(selected_company, selected_option, start_date, end_date, uploaded_file)
