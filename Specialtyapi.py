import os
import snowflake.connector
import pandas as pd
import numpy as np
import json

# --- New Part: Extract NPI Number column from Provider sheet and save as JSON ---
provider_excel_path = r"Excel Files/Output.xlsx"
provider_sheet = "Provider"
try:
    provider_df = pd.read_excel(provider_excel_path, sheet_name=provider_sheet)
    if 'NPI Number' in provider_df.columns:
        def clean_npi(npi):
            if pd.isnull(npi):
                return None
            npi_str = str(npi)
            # Remove trailing .0 if present
            if npi_str.endswith('.0'):
                npi_str = npi_str[:-2]
            return npi_str
        npi_list = provider_df['NPI Number'].dropna().apply(clean_npi).tolist()
        with open("Excel Files/npi_list.json", "w") as f:
            json.dump(npi_list, f, indent=2)
        print(f"NPI list extracted and saved to Excel Files/npi_list.json")
        print(f"Number of NPI Numbers found: {len(npi_list)}")
    else:
        print(f"Column 'NPI Number' not found in {provider_excel_path} sheet {provider_sheet}")
except Exception as e:
    print(f"Error extracting NPI list: {e}")

# --- Read NPI list from JSON for Snowflake query ---
with open("Excel Files/npi_list.json", "r") as f:
    npi_list = json.load(f)

# Prepare the NPI list for SQL IN clause
npi_in_clause = ", ".join([f"'{{}}'".format(npi) for npi in npi_list])

# Connect to Snowflake using SSO (external browser authentication)
conn = snowflake.connector.connect(
    user="dhruv.bhattacharjee@zocdoc.com",
    account="OLIKNSY-ZOCDOC_001",
    warehouse="USER_QUERY_WH",
    database='CISTERN',
    schema='PROVIDER_PREFILL',  # updated schema
    role="PROD_OPS_PUNE_ROLE",
    authenticator='externalbrowser'
)

try:
    cs = conn.cursor()
    query = f"""
    SELECT * FROM merged_provider
    WHERE NPI:value::string IN ({npi_in_clause})
    """
    cs.execute(query)
    results = cs.fetchall()
    columns = [desc[0] for desc in cs.description]
    df = pd.DataFrame(results, columns=columns)
    # Remove timezone info from all datetime columns
    for col in df.select_dtypes(include=['datetimetz']).columns:
        df[col] = df[col].dt.tz_localize(None)
    for col in df.columns:
        if df[col].dtype == 'object':
            if df[col].apply(lambda x: hasattr(x, 'tzinfo') and x.tzinfo is not None).any():
                df[col] = df[col].apply(lambda x: x.tz_localize(None) if hasattr(x, 'tzinfo') and x.tzinfo is not None else x)
    # Select only the required columns
    selected_columns = ['NPI', 'SPECIALTIES', 'FIRST_NAME', 'LAST_NAME', 'SUFFIX']
    df_selected = df[selected_columns].copy()
    # Extract 'value' from JSON strings if present, special handling for SPECIALTIES
    def extract_value(val, colname):
        if isinstance(val, str):
            try:
                parsed = json.loads(val)
                if colname == 'SPECIALTIES' and isinstance(parsed, list) and len(parsed) > 0:
                    first = parsed[0]
                    if isinstance(first, dict) and 'value' in first:
                        return first['value']
                if isinstance(parsed, dict) and 'value' in parsed:
                    return parsed['value']
            except Exception:
                pass
        return val
    for col in selected_columns:
        df_selected[col] = df_selected[col].apply(lambda x: extract_value(x, col))
    # Drop rows where SPECIALTIES is blank or null
    df_selected = df_selected[df_selected['SPECIALTIES'].notnull() & (df_selected['SPECIALTIES'] != '')]
    # Save to Excel
    output_path = r"Excel Files/snowflake.xlsx"
    df_selected.to_excel(output_path, index=False)
    print(f"Results saved to {output_path}")
    # Remove duplicate rows based on NPI and overwrite the file
    df_nodedup = df_selected.drop_duplicates(subset=['NPI'], keep='first')
    df_nodedup.to_excel(output_path, index=False)
    print(f"Duplicates removed based on NPI. Final file saved to {output_path}")
finally:
    cs.close()
    conn.close()
