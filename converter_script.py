import fitz         # PyMuPDF
from PIL import Image
import pytesseract
import io
from pathlib import Path
import os

# Step 1: 
# point to your tesseract exe if not in PATH
# Ensure this path is correct for your system
# Example: pytesseract.pytesseract.tesseract_cmd = r"C:\Program Files\Tesseract-OCR\tesseract.exe"
# If Tesseract is in your system's PATH, this line might not be strictly necessary.
# However, it's good practice to explicitly set it for robustness.
try:
    pytesseract.pytesseract.tesseract_cmd = r"C:\Program Files\Tesseract-OCR\tesseract.exe"
except pytesseract.TesseractNotFoundError:
    print("Tesseract-OCR is not found in the specified path or system PATH. OCR functionality may not work.")
    print("Please install Tesseract-OCR or update the path in pdf_to_text_script.py.")


def run_pdf_to_text_process(company_folder_name, periods_to_process, extraction_method):
    company_base_path = Path(r"D:\Visual Studio Projects\Financial Statement Data Retriever") / company_folder_name
    base_pdf_dir = company_base_path / "financial_statements"
    ocr_dir = company_base_path / "text_statements"

    ocr_dir.mkdir(parents=True, exist_ok=True)

    results = []
    results.append(f"--- Starting PDF Text Extraction Process ({extraction_method.upper()} method) ---")

    for period in periods_to_process:
        pdf_path = base_pdf_dir / f"{period}.pdf"
        out_txt = ocr_dir / f"{period}_ocr.txt"

        if not pdf_path.exists():
            error_message = f"Warning: PDF file not found for {period} at {pdf_path}. Skipping text extraction for this period."
            print(error_message)
            results.append(error_message)
            continue

        status_message = f"\nProcessing PDF for period: {period} ({pdf_path}) using {extraction_method.upper()}..."
        print(status_message)
        results.append(status_message)
        
        doc = None
        try:
            doc = fitz.open(pdf_path)
            with out_txt.open("w", encoding="utf-8") as fout:
                for pageno in range(len(doc)):
                    page = doc.load_page(pageno)
                    
                    if extraction_method.lower() == "ocr":
                        pix = page.get_pixmap(dpi=650)
                        img_bytes = pix.tobytes("png")
                        img = Image.open(io.BytesIO(img_bytes))
                        text = pytesseract.image_to_string(img, lang="vie", config="--psm 3") # Assuming 'vie' language pack
                    else: # extraction_method.lower() == "direct"
                        text = page.get_text("text")

                    fout.write(f"--- PAGE {pageno+1} ---\n")
                    fout.write(text + "\n\n")
            status_message = f"Text output for {period} saved to: {out_txt}"
            print(status_message)
            results.append(status_message)

        except Exception as e:
            error_message = f"An error occurred during {extraction_method.upper()} for {period} at {pdf_path}: {e}. Skipping this period."
            print(error_message)
            results.append(error_message)
        finally:
            if doc:
                doc.close()
    
    results.append("\n--- PDF Text Extraction Process Complete ---")
    return "\n".join(results)



# Step 2: This script uses the LLM to extract financial data from the text files.
import os
import json
import pandas as pd
from pathlib import Path
from langchain_google_genai import ChatGoogleGenerativeAI
from langchain_core.prompts import ChatPromptTemplate
from langchain_core.output_parsers import StrOutputParser
import re

# --- Helper function to extract text from a specific page range ---
def extract_pages(text_content, start_page=None, end_page=None):
    if start_page is None and end_page is None:
        return text_content

    lines = text_content.split('\n')
    filtered_lines = []
    current_page = 0
    in_desired_range = False

    for line in lines:
        page_header_match = re.match(r'--- PAGE (\d+) ---', line)
        if page_header_match:
            current_page = int(page_header_match.group(1))
            if (start_page is None or current_page >= start_page) and \
               (end_page is None or current_page <= end_page):
                in_desired_range = True
                filtered_lines.append(line)
            else:
                in_desired_range = False
        elif in_desired_range:
            filtered_lines.append(line)
    
    return "\n".join(filtered_lines)

def run_converter_process(company_folder_name, periods_to_process, extraction_method, start_page, end_page):
    # Ensure GOOGLE_API_KEY is set in the environment or passed securely
    # For Streamlit, it's often better to set it as an environment variable
    # or pass it securely. This line ensures it's available for this function.
    os.environ["GOOGLE_API_KEY"] = "AIzaSyD1f3CDdw71J98b4LEFFM6IUY893qfnqdg" # Removed hardcoded key

    company_base_path = Path(r"D:\Visual Studio Projects\Financial Statement Data Retriever") / company_folder_name
    json_dir = company_base_path / "json_statements"
    excel_dir = company_base_path / "excel_statements"
    ocr_dir = company_base_path / "text_statements"

    json_dir.mkdir(parents=True, exist_ok=True)
    excel_dir.mkdir(parents=True, exist_ok=True)

    llm = ChatGoogleGenerativeAI(model="gemini-2.5-flash", temperature=0.05)
    prompt_template = ChatPromptTemplate.from_messages(
        [
            ("system", "You are an expert financial analyst. Your task is to extract various line items and their values from the provided text. "
                       "Output the extracted data as a JSON array of objects, where each object has 'item_number' (if there is item number, or else leave blank),'statement_type', 'item', 'year', and 'value'. "
                       "Ensure values are numeric (remove commas, currency symbols, etc.) or leave empty if not found."
                       "Ensure that the line items, as well as the name of the statements are the same as the language being used in the text."
                       "Sometimes there can be grammatial error and line item numering error, make sure to fix it as well, don't be too rigid"
                       "ONLY take the current year from this statement, not the last years."
                       "Make sure that the line items are in proper form, that is no FULL CAPITALIZTATION, and only First Letter Capitalization"),
            ("human", "Extract the 3 main statements from the Financial report, put the columns as ['statement_type', 'item', 'year', 'value] :\n\n{text}")
        ]
    )
    output_parser = StrOutputParser()
    chain = prompt_template | llm | output_parser

    results = []
    for period in periods_to_process:
        ocr_text_file_path = ocr_dir / f"{period}_ocr.txt"
        output_json_file_path = json_dir / f"{period}_financial_statements_raw.json"

        status_message = f"Processing period: {period}"
        print(status_message)
        results.append(status_message)

        if not ocr_text_file_path.exists():
            error_message = f"Error: OCR text file not found for {period} at {ocr_text_file_path}. Skipping LLM extraction for this period."
            print(error_message)
            results.append(error_message)
            continue
        
        try:
            with ocr_text_file_path.open("r", encoding="utf-8") as f:
                ocr_content = f.read()
        except Exception as e:
            error_message = f"Error reading OCR text file for {period} at {ocr_text_file_path}: {e}. Skipping LLM extraction."
            print(error_message)
            results.append(error_message)
            continue

        filtered_ocr_content = extract_pages(ocr_content, start_page, end_page)
        
        if not filtered_ocr_content.strip():
            warning_message = f"No content found in the specified page range ({start_page}-{end_page}) for {period}. Skipping LLM extraction."
            print(warning_message)
            results.append(warning_message)
            continue

        try:
            status_message = f"Sending text for {period} (pages {start_page}-{end_page} if specified) to Gemini 2.5 Flash for extraction..."
            print(status_message)
            results.append(status_message)
            llm_response = chain.invoke({"text": filtered_ocr_content})
            status_message = f"Received response from Gemini for {period}."
            print(status_message)
            results.append(status_message)

            with output_json_file_path.open("w", encoding="utf-8") as f:
                f.write(llm_response)
            status_message = f"Successfully saved raw LLM output for {period} to: {output_json_file_path}"
            print(status_message)
            results.append(status_message)

        except Exception as e:
            error_message = f"An error occurred during LLM invocation for {period}: {e}"
            print(error_message)
            results.append(error_message)
            continue
    
    results.append("\n--- LLM Extraction Process Complete ---")
    return "\n".join(results)

# Step 3: This script merges the individual Excel files generated 
# by the converter into a single long-format DataFrame and then separates them 
# by statement type

import pandas as pd
import numpy as np
from pathlib import Path
import os

def run_merger_process(company_folder_name, periods_to_process):
    company_base_path = Path(r"D:\Visual Studio Projects\Financial Statement Data Retriever") / company_folder_name
    base_dir = company_base_path / "excel_statements"
    period_statements_dir = company_base_path / "period_statements"

    period_statements_dir.mkdir(parents=True, exist_ok=True)

    results = []
    results.append(f"{' BEGINNING CONCATENATING EACH PERIODS ':=^100}")

    financial_statements = []
    found_files_count = 0
    for period in periods_to_process:
        statement_path = base_dir / f"{period}_financial_statements.xlsx"
        if not statement_path.exists():
            msg = f'Warning: Excel file not found for period {period} at {statement_path}. Skipping this period.'
            results.append(msg)
            continue
        try:
            df_statement = pd.read_excel(statement_path)
            financial_statements.append(df_statement)
            found_files_count += 1
        except Exception as e:
            msg = f'Error reading {statement_path}: {e}. Skipping this period.'
            results.append(msg)
            continue

    if found_files_count > 0:
        results.append(f'Successfully read in {found_files_count} years of financial statements in Excel \n')
    else:
        results.append(f'No financial statements were successfully read from Excel files. Please check paths and file existence.')
        return "\n".join(results) # Exit early if no files found

    row_length = sum(len(df_statement) for df_statement in financial_statements)
    concatenated_df = pd.concat(financial_statements, ignore_index=True)

    if len(concatenated_df) == row_length:
        results.append(f'1) SUCCESS: Concatenated successfully dataframes from all periods. Total rows: {len(concatenated_df)}')
    else:
        results.append(f'1) ERROR: There are missing rows or an issue during concatenation. Expected {row_length} rows, got {len(concatenated_df)}.')

    if 'statement_type' in concatenated_df.columns:
        concatenated_df['statement_type'] = concatenated_df['statement_type'].astype(str).str.title()
        results.append("Applied proper casing to 'statement_type' column.")

    results.append("\n--- Saving Full Concatenated DataFrame ---")
    full_concatenated_output_path = period_statements_dir / "all_periods_concatenated.xlsx"
    try:
        concatenated_df.to_excel(full_concatenated_output_path, index=False)
        results.append(f"Successfully saved full concatenated DataFrame to: {full_concatenated_output_path}")
    except Exception as e:
        results.append(f"ERROR: Could not save full concatenated DataFrame: {e}")
    results.append("------------------------------------------")

    results.append(f"\n{' SEPARATING BY STATEMENT TYPE AND SAVING ':=^100}")
    
    unique_statement_types = concatenated_df['statement_type'].unique()
    
    if len(unique_statement_types) > 0:
        results.append(f"Found {len(unique_statement_types)} unique statement types: {', '.join(unique_statement_types)}")
        for st_type in unique_statement_types:
            df_filtered = concatenated_df[concatenated_df['statement_type'] == st_type].copy()
            output_file_path = period_statements_dir / f"{st_type}.xlsx"
            try:
                df_filtered.to_excel(output_file_path, index=False)
                results.append(f"  - Successfully saved '{st_type}' to: {output_file_path}")
            except Exception as e:
                results.append(f"  - ERROR: Could not save '{st_type}' to {output_file_path}: {e}")
    else:
        results.append("No unique 'statement_type' found in the concatenated data. No individual files created.")

    results.append("\n--- Statement Separation and Saving Complete ---")
    return "\n".join(results)

# Step 4: This script takes the merged data, cleans values, 
# and pivots it into a wide format (items as index, years as columns).

import pandas as pd
from pathlib import Path
import os

def run_formatter_process(company_folder_name, periods_to_process):
    company_base_path = Path(r"D:\Visual Studio Projects\Financial Statement Data Retriever") / company_folder_name
    period_statements_dir = company_base_path / "period_statements"
    final_statements_dir = company_base_path / "final_statements"

    final_statements_dir.mkdir(parents=True, exist_ok=True)

    results = []
    results.append("--- Starting Financial Statement Reformatting ---")

    # Process the 'all_periods_concatenated.xlsx' file first if it exists
    all_periods_file_path = period_statements_dir / "all_periods_concatenated.xlsx"
    if all_periods_file_path.exists():
        results.append(f"\nProcessing combined file: {all_periods_file_path.name}")
        try:
            df_long = pd.read_excel(all_periods_file_path)

            required_columns = ['item', 'year', 'value', 'statement_type']
            if not all(col in df_long.columns for col in required_columns):
                msg = f"Warning: Skipping {all_periods_file_path.name} — missing required columns ({', '.join(required_columns)})."
                results.append(msg)
            else:
                def clean_value(x):
                    s = str(x).strip()
                    if s == 'nan' or s == '' or s.lower() == 'n/a':
                        return pd.NA
                    if s.startswith('(') and s.endswith(')'):
                        s = '-' + s[1:-1]
                    s = s.replace(',', '').replace(' ', '')
                    s = pd.Series([s]).replace(r'[^\d\.\-]', '', regex=True).iloc[0]
                    return pd.to_numeric(s, errors='coerce')

                df_long['value'] = df_long['value'].apply(clean_value)
                df_long = df_long.dropna(subset=['item', 'year'])
                df_long['item'] = df_long['item'].astype(str)
                df_long['year'] = df_long['year'].astype(str)

                # Group by statement_type, item, and year to handle potential duplicates
                df_grouped = (
                    df_long
                    .sort_values(['statement_type', 'item', 'year'])
                    .groupby(['statement_type', 'item', 'year'], as_index=False)
                    .agg({'value': 'first'}) # Take the first non-null value
                )

                # Iterate through unique statement types to create separate wide-format Excel files
                for st_type in df_grouped['statement_type'].unique():
                    df_statement_type = df_grouped[df_grouped['statement_type'] == st_type]
                    
                    # Pivot to wide format
                    df_wide = df_statement_type.pivot_table(index='item', columns='year', values='value', aggfunc='first')
                    df_wide.columns.name = None

                    # Reorder columns to follow periods_to_process if available
                    if periods_to_process and isinstance(periods_to_process, (list, tuple)):
                        ordered = [str(p) for p in periods_to_process if str(p) in df_wide.columns]
                        remaining = [c for c in df_wide.columns if c not in ordered]
                        df_wide = df_wide.reindex(columns=ordered + remaining)

                    output_file = final_statements_dir / f"{st_type}.xlsx"
                    df_wide.to_excel(output_file)
                    results.append(f"Successfully reformatted and saved '{st_type}' to: {output_file}")
        except Exception as e:
            results.append(f"Error processing {all_periods_file_path.name}: {e}")
    else:
        results.append(f"Warning: Combined file '{all_periods_file_path.name}' not found. Skipping formatting.")


    results.append("\n--- Financial Statement Reformatting Complete ---")
    return "\n".join(results)

# Step 5: This script uses the LLM to standardize 
# the line item names across different financial statements.

import pandas as pd
from pathlib import Path
import os
import json
from langchain_google_genai import ChatGoogleGenerativeAI
from langchain_core.prompts import ChatPromptTemplate
from langchain_core.output_parsers import StrOutputParser

def run_standardizer_process(company_folder_name):
    # Ensure GOOGLE_API_KEY is set in the environment or passed securely
    # For Streamlit, it's best to set it in the Streamlit app or as an env var
    # os.environ["GOOGLE_API_KEY"] = "YOUR_GOOGLE_API_KEY" # Removed hardcoded key

    company_base_path = Path(r"D:\Visual Studio Projects\Financial Statement Data Retriever") / company_folder_name
    input_dir = company_base_path / "final_statements"
    output_dir = company_base_path / "final_statements_standardized"

    output_dir.mkdir(parents=True, exist_ok=True)

    llm = ChatGoogleGenerativeAI(model="gemini-2.5-flash", temperature=0.5)
    prompt_template = ChatPromptTemplate.from_messages(
        [
            ("system", "You are an expert financial analyst specializing in financial statements. "
                       "Your task is to standardize financial statement line items. "
                       "You will be given a list of items, which may contain variations due to OCR errors, slightly different phrasing, or garbled numbering. "
                       "For each group of semantically similar items, identify them and propose a single, concise, and commonly accepted standardized name. "
                       "The standardized name should be in proper case (first letter of each word capitalized). "
                       "Prioritize standardized names that include a line item number if available among the original items. "
                       "If a 'total' type item, ensure its standardized name clearly reflects it as a total. "
                       "Make sure that the item that has the same name, isn't containing different values from each other. Since there could be sub-items that have the same name but bear different values for their parent item."
                       "Output the mapping as a JSON array of objects. Each object in the array should represent a standardized item and contain two keys: "
                       "'standardized_item' (the proposed standardized name) and 'original_items' (a list of all original items that map to this standardized name). "
                       "The order of objects in the JSON array MUST represent the logical order of items in a financial statement (e.g., assets before liabilities, short-term before long-term, and within sections, by line item number if present). "
                       "Ensure all original items from the input list are present in your output mapping under their respective standardized items."),
            ("human", "Standardize the following financial statement items:\n\n{items_list_json}")
        ]
    )
    output_parser = StrOutputParser()
    chain = prompt_template | llm | output_parser

    results = []
    results.append("--- Starting Financial Statement Item Standardization ---")

    for file_path in input_dir.glob("*.xlsx"):
        results.append(f"\nProcessing file for standardization: {file_path.name}")
        try:
            df_wide = pd.read_excel(file_path, index_col=0)

            if df_wide.empty:
                msg = f"  Warning: {file_path.name} is empty. Skipping standardization."
                results.append(msg)
                continue

            items_to_standardize = df_wide.index.astype(str).unique().tolist()

            if not items_to_standardize:
                msg = f"  No items found in {file_path.name} to standardize. Skipping."
                results.append(msg)
                continue

            results.append(f"  Found {len(items_to_standardize)} unique items. Sending to Gemini for standardization...")
            
            items_list_json = json.dumps(items_to_standardize, ensure_ascii=False, indent=2)
            llm_response = chain.invoke({"items_list_json": items_list_json})
            results.append(f"  Received standardization mapping from Gemini for {file_path.name}.")

            cleaned_json_string = llm_response.strip()
            if cleaned_json_string.startswith("```json"):
                cleaned_json_string = cleaned_json_string[len("```json"):].strip()
            if cleaned_json_string.endswith("```"):
                cleaned_json_string = cleaned_json_string[:-len("```")].strip()

            standardization_groups = json.loads(cleaned_json_string)
            
            item_mapping = {}
            ordered_standardized_items = []

            for group in standardization_groups:
                standardized_name = group['standardized_item']
                ordered_standardized_items.append(standardized_name)
                for original_item in group['original_items']:
                    item_mapping[original_item] = standardized_name

            df_temp = df_wide.rename(index=item_mapping)
            df_aggregated = df_temp.groupby(df_temp.index).sum()
            df_standardized = df_aggregated.reindex(ordered_standardized_items)

            output_file_path = output_dir / file_path.name
            
            df_standardized.to_excel(output_file_path)
            results.append(f"  Successfully standardized and saved '{file_path.name}' to: {output_file_path}")
            results.append(f"  Final standardized DataFrame shape: {df_standardized.shape}")
            results.append(f"  Final standardized DataFrame head:\n{df_standardized.head().to_string()}")

        except json.JSONDecodeError as e:
            results.append(f"  ERROR: JSON decoding failed for LLM response for {file_path.name}: {e}")
            results.append(f"  LLM Response (raw):\n{llm_response}")
            continue
        except Exception as e:
            results.append(f"  ERROR processing {file_path.name}: {e}")
            continue

    results.append("\n--- Financial Statement Item Standardization Complete ---")
    return "\n".join(results)
