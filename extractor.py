import streamlit as st
from pdf2image import convert_from_path
import base64
import tempfile
import os
from PyPDF2 import PdfReader, PdfWriter
import pandas as pd
from st_aggrid import AgGrid, GridOptionsBuilder
from io import BytesIO
from datetime import datetime
import sys
from pathlib import Path

wd = Path(__file__).parent.parent.resolve()
sys.path.append(str(wd))

st.set_page_config(layout="wide")

from sidebar import Sidebar

class ExtractorGUI:

    def __init__(self):
        super()
 
    def get_total_pages(self, pdf_path):
        pdf = PdfReader(pdf_path)
        total_pages = len(pdf.pages)
        return total_pages

    def extract_pages(self, pdf_path, start_page, end_page):
        pdf = PdfReader(pdf_path)
        pdf_writer = PdfWriter()
        # Add selected pages to PDF writer
        for page_num in range(start_page - 1, end_page):
            pdf_writer.add_page(pdf.pages[page_num])
        # Write extracted pages to a new PDF file
        with open("extracted_pages.pdf", "wb") as output_pdf:
            pdf_writer.write(output_pdf)

    # # Function to display Excel data using st_aggrid
    def display_excel(self, filename):
        df = pd.read_excel(filename)
        gob = GridOptionsBuilder.from_dataframe(df)
        grid_options = gob.build()

        # Display the AgGrid
        editable_df = AgGrid(df, gridOptions=grid_options, height=500, width='100%', editable=True)
        return editable_df

    # Main function to extract PDF data
    def extract(self):
        @st.cache_data
        
        def get_img_as_base64(file):
            with open(file, "rb") as f:
                data = f.read()
            return base64.b64encode(data).decode()
        
        # #background and sidebar images
        img = get_img_as_base64("app_images/background_img.png")
        himg = get_img_as_base64("app_images/wallp_sidebr.png")

        footer_html = """
        <footer style="position:fixed; bottom:0;right:0; width:100%; text-align:center; background-image: url('app_images/fbg.png'); background-size: cover; padding: 2px; color: white;">
            <p>&#169 Commercial bank 2024</p>
        </footer>
        """
        with open('style1.css') as f:
            css = f.read()

        page_bg_img = f"""<style>{css}
                                [data-testid="stAppViewContainer"] > .main {{
                                background-image: url("data:image/png;base64,{img}");
                                background-size: cover;
                                background-repeat: no-repeat;
                                background-attachment: scroll;
                                }}
                                [data-testid="stSidebar"]{{
                                background-image: url("data:image/png;base64,{himg}");
                                background-size: cover;
                                
                                }}
                                </style>"""
        st.markdown(page_bg_img, unsafe_allow_html=True)
        st.markdown(footer_html, unsafe_allow_html=True)
    
        # Display title and hint message
        title_text = "<h1><span style='font-family:Montserrat, sans-serif; font-size: 62px;'>Do<span style='font-size: 80px;color:#4ab847'>K</span><span style='color:#006eb9'>.AI</span</span></h1>"
        st.markdown(title_text, unsafe_allow_html=True)
        hint_slot = st.empty()
        hint_slot.write("Hint💡: select options from the sidebar and upload a PDF file")

        sidebar = Sidebar()
        selected_company, selected_option, start_date, end_date, uploaded_file = sidebar.run()

        if uploaded_file is not None:
                file_name = uploaded_file.name
                hint_slot.empty()
                temp_dir = tempfile.TemporaryDirectory()
                temp_pdf_path = os.path.join(temp_dir.name, file_name)
                with open(temp_pdf_path, "wb") as f:
                    f.write(uploaded_file.getbuffer())

                col1, col2 = st.columns(spec=[12,10], gap="medium")
                with col1:
                    pdf_data = uploaded_file.read()
                    encoded_pdf = base64.b64encode(pdf_data).decode('utf-8')
                    st.write(f'<iframe src="data:application/pdf;base64,{encoded_pdf}" width="100%" height="1000px"></iframe>', unsafe_allow_html=True)
                
                with col2:
                    page_range_input = st.text_input("Enter page range (e.g., 1-5)", "1-" + str(self.get_total_pages(temp_pdf_path)))

                    try:
                        start_page, end_page = map(int, page_range_input.split('-'))
                        if not (1 <= start_page <= end_page <= self.get_total_pages(temp_pdf_path)):
                            st.warning("Invalid page range! Please enter a valid page range.")
                        else:
                            if st.button("Extract Pages"):
                                self.extract_pages(temp_pdf_path, start_page, end_page)

                                with st.spinner("Converting PDF to images..."):
                                    images = convert_from_path(temp_pdf_path)

                                output_dir = os.path.join(os.path.dirname(__file__), "images")
                                os.makedirs(output_dir, exist_ok=True)
                                saved_image_paths = []

                                progress_text = "Saving images..."
                                progress_bar = st.progress(0, text=progress_text)

                                for idx in range(start_page - 1, end_page):
                                    img_filename = f"page_{idx + 1}.jpg"
                                    img_path = os.path.join(output_dir, img_filename)
                                    images[idx].save(img_path)
                                    saved_image_paths.append(img_path)

                                progress_bar.empty()

                                filename = "example.xlsx"
                                file_extension = "xlsx"
                                download_folder = "download"
                                filepath = os.path.join(download_folder, filename)

                                editable_df = self.display_excel(filepath)
                                edited_df = editable_df['data']
                                excel_buffer = BytesIO()
                                edited_df.to_excel(excel_buffer, index=False)
                                excel_buffer.seek(0)

                                col1, col2 = st.columns(spec=[3,1])
                                with col2:
                                    st.button("Override")
                                with col1:
                                    st.download_button(label="Download Excel File", data=excel_buffer, file_name="example.xlsx", mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet")
                    
                                temp_dir.cleanup()
                    except ValueError:
                        st.warning("Invalid page range format! Please enter a valid page range in the format 'start-end'.")

if __name__ == "__main__":
    ext = ExtractorGUI()
    ext.extract()      

