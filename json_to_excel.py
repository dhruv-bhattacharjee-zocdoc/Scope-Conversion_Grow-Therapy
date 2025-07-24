import json
import pandas as pd

def convert_json_to_excel(json_path, excel_path):
    with open(json_path, 'r', encoding='utf-8') as f:
        data = json.load(f)
    locations = data['practice_locations']
    df = pd.json_normalize(locations)
    for col in df.columns:
        if df[col].apply(lambda x: isinstance(x, list)).any():
            df[col] = df[col].apply(lambda x: ', '.join(x) if isinstance(x, list) else x)
    df.to_excel(excel_path, index=False)
    print(f"Conversion complete! Excel file saved as '{excel_path}'")

if __name__ == "__main__":
    convert_json_to_excel('Excel Files/output.json', 'Excel Files/json_to_excel.xlsx') 