import pandas as pd
from pathlib import Path
import os

def run_formatter_process(company_folder_name, periods_to_process):
    company_base_path = Path(company_folder_name)
    period_statements_dir = company_base_path / "period_statements"
    final_statements_dir = company_base_path / "final_statements"

    final_statements_dir.mkdir(parents=True, exist_ok=True)

    results = []
    results.append("--- Starting Financial Statement Reformatting ---")

    all_periods_file_path = period_statements_dir / "all_periods_concatenated.xlsx"
    
    if not all_periods_file_path.exists():
        msg = f"Error: Combined file '{all_periods_file_path.name}' not found. Cannot proceed with formatting."
        results.append(msg)
        raise FileNotFoundError(msg)

    results.append(f"\nProcessing combined file: {all_periods_file_path.name}")
    
    processed_any_statement = False
    try:
        df_long = pd.read_excel(all_periods_file_path)

        required_columns = ['item', 'year', 'value', 'statement_type']
        if not all(col in df_long.columns for col in required_columns):
            msg = f"Error: Skipping {all_periods_file_path.name} — missing required columns ({', '.join(required_columns)}). Cannot format."
            results.append(msg)
            raise ValueError(msg)
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

            df_grouped = (
                df_long
                .sort_values(['statement_type', 'item', 'year'])
                .groupby(['statement_type', 'item', 'year'], as_index=False)
                .agg({'value': 'first'})
            )

            for st_type in df_grouped['statement_type'].unique():
                df_statement_type = df_grouped[df_grouped['statement_type'] == st_type]
                
                df_wide = df_statement_type.pivot_table(index='item', columns='year', values='value', aggfunc='first')
                df_wide.columns.name = None

                if periods_to_process and isinstance(periods_to_process, (list, tuple)):
                    ordered = [str(p) for p in periods_to_process if str(p) in df_wide.columns]
                    remaining = [c for c in df_wide.columns if c not in ordered]
                    df_wide = df_wide.reindex(columns=ordered + remaining)

                output_file = final_statements_dir / f"{st_type}.xlsx"
                df_wide.to_excel(output_file)
                results.append(f"Successfully reformatted and saved '{st_type}' to: {output_file}")
                processed_any_statement = True
    except Exception as e:
        results.append(f"Error processing {all_periods_file_path.name}: {e}")
        raise # Re-raise the exception

    if not processed_any_statement:
        raise ValueError("No financial statements were successfully formatted and saved.")

    results.append("\n--- Financial Statement Reformatting Complete ---")
    return "\n".join(results)