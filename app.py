import streamlit as st
from pathlib import Path
import os

# Import your refactored scripts
# Make sure these files (e.g., converter_script.py) are in the same directory as app.py
# or in a directory included in your Python path.
from converter_script import run_converter_process
# from merger_script import run_merger_process
# from formatter_script import run_formatter_process
# from standardizer_script import run_standardizer_process

st.set_page_config(layout="wide")
st.title("📊 Financial Statement Data Retriever")

st.markdown("""
This application helps you extract, merge, format, and standardize financial statement data from PDFs.
""")

# --- User Inputs ---
st.header("1. Configuration")

company_folder_name = st.text_input(
    "Enter the company folder name (e.g., PVIAM, General_Mills):",
    value="PVIAM" # Default value for convenience
)

periods_input = st.text_input(
    "Enter periods to process (e.g., 2021, 2022, 2023, 2024), separated by commas:",
    value="2021, 2022, 2023, 2024" # Default value
)
periods_to_process = [p.strip() for p in periods_input.split(',') if p.strip()]

extraction_method = st.radio(
    "Choose PDF text extraction method:",
    ('OCR', 'Direct'),
    index=0 # Default to OCR
)

page_range_input = st.text_input(
    "Enter page range to extract (e.g., 50-90, leave blank for all pages):",
    value="" # Default to all pages
)

start_page = None
end_page = None
if page_range_input:
    try:
        if '-' in page_range_input:
            start_str, end_str = page_range_input.split('-')
            start_page = int(start_str)
            end_page = int(end_str)
            if start_page > end_page:
                st.warning("Warning: Start page is greater than end page. Processing all pages.")
                start_page = None
                end_page = None
        else:
            st.warning("Invalid page range format. Processing all pages.")
    except ValueError:
        st.warning("Invalid page number in range. Processing all pages.")

# --- Google API Key (Optional, if not set as environment variable) ---
# You might want to add a text input for the API key if it's not an environment variable
# google_api_key = st.text_input("Enter your Google API Key (optional, if not set as env var):", type="password")
# if google_api_key:
os.environ["GOOGLE_API_KEY"] = "AIzaSyBAAP0sKv5nJOGySzVmd66I23vUDtjjI-s"


st.markdown("---")
st.header("2. Run Workflow")

if st.button("Start Financial Data Processing"):
    if not company_folder_name:
        st.error("Please enter a company folder name.")
    elif not periods_to_process:
        st.error("Please enter at least one period to process.")
    else:
        st.info("Starting the financial data processing workflow...")
        
        # Display current configuration
        st.subheader("Current Configuration:")
        st.write(f"- Company Folder: **{company_folder_name}**")
        st.write(f"- Periods: **{', '.join(periods_to_process)}**")
        st.write(f"- Extraction Method: **{extraction_method}**")
        st.write(f"- Page Range: **{page_range_input if page_range_input else 'All Pages'}**")
        
        st.subheader("Processing Output:")
        output_area = st.empty() # Placeholder for dynamic output

        # --- Step 1: PDF to Text (OCR or Direct) ---
        output_area.write("### Step 1: Converting PDF to Text files...")
        # You would call your refactored PDF to text function here
        # For now, we'll simulate it or call the converter_script's first part
        # Assuming you have a `run_pdf_to_text_process` function in `pdf_to_text_script.py`
        # For this example, we'll just use the `run_converter_process` which includes LLM extraction
        
        # Placeholder for PDF to Text (if separated from LLM extraction)
        try:
            pdf_to_text_log = run_pdf_to_text_process(company_folder_name, periods_to_process, extraction_method)
            output_area.write(pdf_to_text_log)
        except Exception as e:
            st.error(f"Error in PDF to Text conversion: {e}")
            st.stop() # Stop execution if this critical step fails

        # --- Step 2: LLM Extraction (from 1_test_converter.ipynb) ---
        output_area.write("### Step 2: Extracting data using Gemini 2.5 Flash...")
        try:
            llm_extraction_log = run_converter_process(
                company_folder_name, 
                periods_to_process, 
                extraction_method.lower(), # Pass 'ocr' or 'direct'
                start_page, 
                end_page
            )
            output_area.markdown(f"```\n{llm_extraction_log}\n```") # Display log in a code block
        except Exception as e:
            st.error(f"Error during LLM data extraction: {e}")
            st.stop()

        # --- Step 3: Merging Excel Files (from 2_excel_merger.ipynb) ---
        output_area.write("### Step 3: Merging Excel Files...")
        try:
            merger_log = run_merger_process(company_folder_name, periods_to_process)
            output_area.markdown(f"```\n{merger_log}\n```")
        except Exception as e:
            st.error(f"Error during Excel merging: {e}")
            st.stop()

        # --- Step 4: Formatting Excel Files (from 3_excel_formatter.ipynb) ---
        output_area.write("### Step 4: Formatting Excel Files...")
        try:
            formatter_log = run_formatter_process(company_folder_name)
            output_area.markdown(f"```\n{formatter_log}\n```")
        except Exception as e:
            st.error(f"Error during Excel formatting: {e}")
            st.stop()

        # --- Step 5: Standardizing Excel Files (from 4_excel_standardization.ipynb) ---
        output_area.write("### Step 5: Standardizing Excel Files...")
        try:
            standardizer_log = run_standardizer_process(company_folder_name)
            output_area.markdown(f"```\n{standardizer_log}\n```")
        except Exception as e:
            st.error(f"Error during Excel standardization: {e}")
            st.stop()

        st.success("Workflow completed successfully!")
