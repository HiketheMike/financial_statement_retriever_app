import pandas as pd
from pathlib import Path
import os
import json
from langchain_google_genai import ChatGoogleGenerativeAI
from langchain_core.prompts import ChatPromptTemplate
from langchain_core.output_parsers import StrOutputParser

# Define the GitHub repository name for display purposes
REPO_NAME = "financial_statement_retriever_app"

def run_standardizer_process(company_folder_name):
    company_base_path = Path(company_folder_name)
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

    found_files_to_standardize = False
    processed_any_file_successfully = False

    # Helper to format path for display
    def format_github_path(p: Path):
        return f"{REPO_NAME}/{str(p).replace('\\', '/')}"

    for file_path in input_dir.glob("*.xlsx"):
        found_files_to_standardize = True
        # Changed: Use format_github_path for display
        results.append(f"\nProcessing file for standardization: {format_github_path(file_path)}")
        try:
            df_wide = pd.read_excel(file_path, index_col=0)

            if df_wide.empty:
                # Changed: Use format_github_path for display
                msg = f"  Warning: {format_github_path(file_path)} is empty. Skipping standardization."
                results.append(msg)
                continue

            items_to_standardize = df_wide.index.astype(str).unique().tolist()

            if not items_to_standardize:
                # Changed: Use format_github_path for display
                msg = f"  No items found in {format_github_path(file_path)} to standardize. Skipping."
                results.append(msg)
                continue

            results.append(f"  Found {len(items_to_standardize)} unique items. Sending to Gemini for standardization...")
            
            items_list_json = json.dumps(items_to_standardize, ensure_ascii=False, indent=2)
            llm_response = chain.invoke({"items_list_json": items_list_json})
            # Changed: Use format_github_path for display
            results.append(f"  Received standardization mapping from Gemini for {format_github_path(file_path)}.")

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
            # Changed: Use format_github_path for display
            results.append(f"  Successfully standardized and saved '{format_github_path(file_path)}' to: {format_github_path(output_file_path)}")
            results.append(f"  Final standardized DataFrame shape: {df_standardized.shape}")
            results.append(f"  Final standardized DataFrame head:\n{df_standardized.head().to_string()}")
            processed_any_file_successfully = True

        except json.JSONDecodeError as e:
            # Changed: Use format_github_path for display
            results.append(f"  ERROR: JSON decoding failed for LLM response for {format_github_path(file_path)}: {e}")
            results.append(f"  LLM Response (raw):\n{llm_response}")
            # Do not re-raise here, allow other files to be processed
        except Exception as e:
            # Changed: Use format_github_path for display
            results.append(f"  ERROR processing {format_github_path(file_path)}: {e}")
            # Do not re-raise here, allow other files to be processed

    if not found_files_to_standardize:
        # Changed: Use format_github_path for display
        raise FileNotFoundError(f"No Excel files found in '{format_github_path(input_dir)}' to standardize. Please ensure previous steps completed.")
    if not processed_any_file_successfully:
        raise ValueError("No financial statements were successfully standardized. Check logs for errors.")

    results.append("\n--- Financial Statement Item Standardization Complete ---")
    return "\n".join(results)