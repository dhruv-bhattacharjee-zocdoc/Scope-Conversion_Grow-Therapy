import requests
import json
import pandas as pd
from json_to_excel import convert_json_to_excel
import os

def run_api():
    # Read unique Practice IDs from Excel file
    excel_path = r'C:\Users\dhruv.bhattacharjee\Desktop\PDO Data Transposition\Scope Conversion_Grow Therapy\Excel Files\Grow Therapy - Sample Data.xlsx'
    df = pd.read_excel(excel_path)
    unique_ids = df['Practice ID'].dropna().astype(str).unique().tolist()

    print("Unique monolith practice IDs:")
    print(unique_ids)

    # Step 1: Get practice_id from monolith_practice_id
    monolith_url = 'https://provider-reference-v1.east.zocdoccloud.com/provider-reference/v1/practice/ids-by-monolith-ids~batchGet'
    headers = {
        'accept': 'application/json',
        'Content-Type': 'application/json'
    }
    monolith_data = {
        "monolith_practice_ids": unique_ids
    }

    monolith_response = requests.post(monolith_url, headers=headers, json=monolith_data)
    print(f"Monolith Status Code: {monolith_response.status_code}")
    try:
        monolith_json = monolith_response.json()
        # Extract all unique practice_ids from the response (new structure)
        practice_ids = [item['practice_id'] for item in monolith_json.get('practice_ids', []) if 'practice_id' in item]
        practice_ids = list(set(practice_ids))
        if not practice_ids:
            raise ValueError("No practice_id found in monolith response.")
        # Print mapping of monolith_practice_id to practice_id (without header line)
        for item in monolith_json.get('practice_ids', []):
            print(f"Monolith ID: {item.get('monolith_practice_id')} -> Practice Cloud ID: {item.get('practice_id')}")
    except Exception as e:
        print("Failed to extract practice_id:", e)
        with open(r'C:\Users\dhruv.bhattacharjee\Desktop\PDO Data Transposition\Scope Conversion_Grow Therapy\Excel Files\output.json', 'w', encoding='utf-8') as f:
            f.write(f"Failed to extract practice_id: {e}\nResponse: {monolith_response.text}")
        return

    # Step 2: Use the extracted practice_ids in the second request
    location_url = 'https://provider-reference-v1.east.zocdoccloud.com/provider-reference/v1/practice/location~batchGet'
    location_data = {
        "practice_ids": practice_ids
    }

    location_response = requests.post(location_url, headers=headers, json=location_data)
    print(f"Location Status Code: {location_response.status_code}")
    try:
        location_json = location_response.json()
        output_json_path = r'C:\Users\dhruv.bhattacharjee\Desktop\PDO Data Transposition\Scope Conversion_Grow Therapy\Excel Files\output.json'
        output_excel_path = r'C:\Users\dhruv.bhattacharjee\Desktop\PDO Data Transposition\Scope Conversion_Grow Therapy\Excel Files\json_to_excel.xlsx'
        with open(output_json_path, 'w', encoding='utf-8') as f:
            json.dump(location_json, f, ensure_ascii=False, indent=2)
        # Call the conversion function
        convert_json_to_excel(output_json_path, output_excel_path)
        # Delete the JSON file after Excel is created
        if os.path.exists(output_json_path):
            os.remove(output_json_path)
            print(f"Deleted JSON file: {output_json_path}")
    except Exception:
        print("Location Response Text:", location_response.text)
        with open(r'C:\Users\dhruv.bhattacharjee\Desktop\PDO Data Transposition\Scope Conversion_Grow Therapy\Excel Files\output.json', 'w', encoding='utf-8') as f:
            f.write(location_response.text)

if __name__ == "__main__":
    run_api()
