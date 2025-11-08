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
    # Change: Use Path(company_folder_name) to make it relative to the repo root
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