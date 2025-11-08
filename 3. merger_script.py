import pandas as pd
import numpy as np
from pathlib import Path
import os

# Define the GitHub repository name for display purposes
REPO_NAME = "financial_statement_retriever_app"

# Helper to format path for display
def format_github_path(p: Path):
    # Get the string representation of the path
    path_str = str(p)
    
    # If it starts with the current working directory, remove that prefix
    cwd_str = str(Path.cwd())
    if path_str.startswith(cwd_str):
        # Remove the cwd prefix, and handle potential separator differences
        path_str = path_str[len(cwd_str):].lstrip(os.sep).lstrip('/')
    
    # Ensure forward slashes for GitHub style
    path_str = path_str.replace('\\', '/')
    
    # Prepend REPO_NAME
    return f"{REPO_NAME}/{path_str}"

def run_merger_process(company_folder_name, periods_to_process):
    company_base_path = Path(company_folder_name)
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
            # Changed: Use the refined format_github_path for display
            msg = f'Warning: Excel file not found for period {period} at {format_github_path(statement_path)}. Skipping this period.'
            results.append(msg)
            continue
        try:
            df_statement = pd.read_excel(statement_path)
            financial_statements.append(df_statement)
            found_files_count += 1
        except Exception as e:
            # Changed: Use the refined format_github_path for display
            msg = f'Error reading {format_github_path(statement_path)}: {e}. Skipping this period.'
            results.append(msg)
            continue

    if found_files_count > 0:
        results.append(f'Successfully read in {found_files_count} years of financial statements in Excel \n')
    else:
        msg = f'No financial statements were successfully read from Excel files. Please check paths and file existence.'
        results.append(msg)
        raise ValueError(msg) # Raise error if no files found

    row_length = sum(len(df_statement) for df_statement in financial_statements)
    concatenated_df = pd.concat(financial_statements, ignore_index=True)

    if len(concatenated_df) == row_length:
        results.append(f'1) SUCCESS: Concatenated successfully dataframes from all periods. Total rows: {len(concatenated_df)}')
    else:
        results.append(f'1) ERROR: There are missing rows or an issue during concatenation. Expected {row_length} rows, got {len(concatenated_df)}.')
        # This is a warning, not necessarily a hard stop, but could be upgraded to an error if desired.

    if 'statement_type' in concatenated_df.columns:
        concatenated_df['statement_type'] = concatenated_df['statement_type'].astype(str).str.title()
        results.append("Applied proper casing to 'statement_type' column.")

    results.append("\n--- Saving Full Concatenated DataFrame ---")
    full_concatenated_output_path = period_statements_dir / "all_periods_concatenated.xlsx"
    try:
        concatenated_df.to_excel(full_concatenated_output_path, index=False)
        # Changed: Use the refined format_github_path for display
        results.append(f"Successfully saved full concatenated DataFrame to: {format_github_path(full_concatenated_output_path)}")
    except Exception as e:
        results.append(f"ERROR: Could not save full concatenated DataFrame: {e}")
        raise # Re-raise the exception if saving fails
    results.append("------------------------------------------")

    results.append(f"\n{' SEPARATING BY STATEMENT TYPE AND SAVING ':=^100}")
    
    unique_statement_types = concatenated_df['statement_type'].unique()
    
    if len(unique_statement_types) > 0:
        results.append(f"Found {len(unique_statement_types)} unique statement types: {', '.join(unique_statement_types)}")
        processed_any_statement_type = False
        for st_type in unique_statement_types:
            df_filtered = concatenated_df[concatenated_df['statement_type'] == st_type].copy()
            output_file_path = period_statements_dir / f"{st_type}.xlsx"
            try:
                df_filtered.to_excel(output_file_path, index=False)
                # Changed: Use the refined format_github_path for display
                results.append(f"  - Successfully saved '{st_type}' to: {format_github_path(output_file_path)}")
                processed_any_statement_type = True
            except Exception as e:
                # Changed: Use the refined format_github_path for display
                results.append(f"  - ERROR: Could not save '{st_type}' to {format_github_path(output_file_path)}: {e}")
        if not processed_any_statement_type:
            raise ValueError("No individual statement type files could be saved after concatenation.")
    else:
        results.append("No unique 'statement_type' found in the concatenated data. No individual files created.")
        raise ValueError("No unique 'statement_type' found in the concatenated data.")

    results.append("\n--- Statement Separation and Saving Complete ---")
    return "\n".join(results)