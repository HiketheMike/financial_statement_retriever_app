import os
import json
import pandas as pd
from pathlib import Path
from langchain_google_genai import ChatGoogleGenerativeAI
from langchain_core.prompts import ChatPromptTemplate
from langchain_core.output_parsers import StrOutputParser
import re

# Define the GitHub repository name for display purposes
REPO_NAME = "financial_statement_retriever_app"

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
    company_base_path = Path(company_folder_name)
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
    processed_any_period = False # Track if any period was successfully processed
    for period in periods_to_process:
        ocr_text_file_path = ocr_dir / f"{period}_ocr.txt"
        output_json_file_path = json_dir / f"{period}_financial_statements_raw.json"
        output_excel_file_path = excel_dir / f"{period}_financial_statements.xlsx"

        # Helper to format path for display
        def format_github_path(p: Path):
            return f"{REPO_NAME}/{str(p).replace('\\', '/')}"

        status_message = f"Processing period: {period}"
        print(status_message)
        results.append(status_message)

        if not ocr_text_file_path.exists():
            # Changed: Use format_github_path for display
            error_message = f"Error: OCR text file not found for {period} at {format_github_path(ocr_text_file_path)}. Skipping LLM extraction for this period."
            print(error_message)
            results.append(error_message)
            continue
        
        try:
            with ocr_text_file_path.open("r", encoding="utf-8") as f:
                ocr_content = f.read()
        except Exception as e:
            # Changed: Use format_github_path for display
            error_message = f"Error reading OCR text file for {period} at {format_github_path(ocr_text_file_path)}: {e}. Skipping LLM extraction."
            print(error_message)
            results.append(error_message)
            continue

        filtered_ocr_content = extract_pages(ocr_content, start_page, end_page)
        
        if not filtered_ocr_content.strip():
            warning_message = f"No content found in the specified page range ({start_page}-{end_page}) for {period}. Skipping LLM extraction."
            print(warning_message)
            results.append(warning_message)
            continue

        llm_response = None
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
            # Changed: Use format_github_path for display
            status_message = f"Successfully saved raw LLM output for {period} to: {format_github_path(output_json_file_path)}"
            results.append(status_message)

            # --- Convert to Pandas DataFrame and Save to Excel ---
            cleaned_json_string = llm_response.strip()
            if cleaned_json_string.startswith("```json"):
                cleaned_json_string = cleaned_json_string[len("```json"):].strip()
            if cleaned_json_string.endswith("```"):
                cleaned_json_string = cleaned_json_string[:-len("```")].strip()

            extracted_data = json.loads(cleaned_json_string)

            if not isinstance(extracted_data, list):
                results.append(f"Warning: Parsed JSON for {period} was not a simple array. Attempting to recover.")
                if isinstance(extracted_data, dict) and "financial_statements" in extracted_data:
                    extracted_data = extracted_data["financial_statements"]
                elif isinstance(extracted_data, dict) and "data" in extracted_data:
                    extracted_data = extracted_data["data"]
                else:
                    extracted_data = []

            if extracted_data:
                df = pd.DataFrame(extracted_data)
                if 'value' in df.columns:
                    df['value'] = df['value'].astype(str).str.replace(',', '').str.strip()
                    df['value'] = pd.to_numeric(df['value'], errors='coerce')
                
                if 'year' not in df.columns:
                    df['year'] = period
                else:
                    df['year'] = df['year'].astype(str)

                df.to_excel(output_excel_file_path, index=False)
                # Changed: Use format_github_path for display
                results.append(f"Successfully extracted {len(df)} financial items for {period}, cleaned, and saved to: {format_github_path(output_excel_file_path)}")
                processed_any_period = True # Mark as successful for at least one period
            else:
                results.append(f"No financial data was extracted or parsed successfully for {period}. Excel file not created.")

        except json.JSONDecodeError as e:
            results.append(f"Error decoding JSON from LLM response for {period}: {e}")
            results.append(f"LLM Response (raw):\n{llm_response}")
        except Exception as e:
            error_message = f"An error occurred during LLM invocation or Excel conversion for {period}: {e}"
            print(error_message)
            results.append(error_message)

    results.append("\n--- LLM Extraction Process Complete ---")
    
    if not processed_any_period:
        raise ValueError("No financial data was successfully extracted and converted to Excel for any period.")
        
    return "\n".join(results)