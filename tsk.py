import pandas as pd
import os
import re
from pathlib import Path

def validate_item_master():
    # ==========================================
    # 1. CONFIGURATION (Update these variables)
    # ==========================================
    excel_path = r'C:\Users\m.emaan\OneDrive - Casa Mia LLC\Documents\Extraction_task\Item_Master.xlsx'
    folder_path = r'C:\Users\m.emaan\OneDrive - Casa Mia LLC\Documents\Extraction_task\Teka' # Unzip the folder first
    target_brand = input("Enter the brand you want to sort by: ").strip()
    
    # Update these if your Excel column headers are different
    brand_column_name = 'Brand' 
    # Optional: If you have a column in Excel to compare the reference against
    excel_ref_column = 'Reference_Number' 

    # ==========================================
    # 2. PART 1: MAPPING THE ITEM MASTER
    # ==========================================
    print(f"Loading Item Master and filtering for '{target_brand}'...")
    try:
        df = pd.read_excel(excel_path)
    except Exception as e:
        print(f"Error loading Excel file: {e}")
        return

    # Assume Column A is the Product Code (index 0)
    col_a_name = df.columns[0]
    
    # Filter by Brand (case-insensitive)
    df_filtered = df[df[brand_column_name].astype(str).str.strip().str.lower() == target_brand.lower()]
    
    if df_filtered.empty:
        print(f"No items found for brand '{target_brand}'. Please check the spelling.")
        return

    # Create a set of product codes from the Excel sheet for fast lookup
    item_master_codes = set(df_filtered[col_a_name].astype(str).str.strip())
    print(f"Found {len(item_master_codes)} products for '{target_brand}' in the Item Master.")

    # Scan the folder for files
    print("Scanning folder for product files...")
    found_files = {}
    for root, _, files in os.walk(folder_path):
        for file in files:
            # Assumes the filename without extension is the product code
            code = Path(file).stem.strip() 
            found_files[code] = os.path.join(root, file)

    found_codes = set(found_files.keys())

    # Map Presence/Absence
    present_in_folder = item_master_codes.intersection(found_codes)
    absent_from_folder = item_master_codes - found_codes
    
    print(f"Match Results: {len(present_in_folder)} found, {len(absent_from_folder)} missing in folder.")

    # ==========================================
    # 3. PART 2: REFERENCE COMPARISON
    # ==========================================
    print("Extracting references from files...")
    results = []
    
    # Process files that matched
    for code in present_in_folder:
        filepath = found_files[code]
        extracted_ref = None
        
        # Read the file to find ;REFERENCE;
        try:
            with open(filepath, 'r', encoding='utf-8', errors='ignore') as f:
                content = f.read()
                # Regex looks for ;REFERENCE; and captures the text right after it 
                # until it hits another semicolon, space, or new line.
                match = re.search(r'REFERENCE([^\s;]+)', content)
                if match:
                    extracted_ref = match.group(1)
                else:
                    extracted_ref = "Tag found"
                    "d, but no value" if "REFERENCE" in content else "Tag missing"
        except Exception as e:
            extracted_ref = "Error reading file content"

        # Check against Excel (if the reference column exists in the sheet)
        expected_ref = None
        match_status = "N/A"
        if excel_ref_column in df_filtered.columns:
            expected_ref = df_filtered[df_filtered[col_a_name] == code][excel_ref_column].values[0]
            match_status = "True" if str(extracted_ref) == str(expected_ref) else "False"

        results.append({
            'Product Code (Col A)': code,
            'Brand': target_brand,
            'Status': 'Present in Folder',
            'Extracted File Reference': extracted_ref,
            'Expected Excel Reference': expected_ref,
            'Reference Match': match_status
        })

    # Log the missing items as well
    for code in absent_from_folder:
        results.append({
            'Product Code (Col A)': code,
            'Brand': target_brand,
            'Status': 'Missing from Folder',
            'Extracted File Reference': 'N/A',
            'Expected Excel Reference': 'N/A',
            'Reference Match': 'N/A'
        })

    # ==========================================
    # 4. EXPORT RESULTS
    # ==========================================
    results_df = pd.DataFrame(results)
    output_filename = f'{target_brand}_validation_report.csv'
    results_df.to_csv(output_filename, index=False)
    print(f"\nTask Complete! Report saved to your current directory as: {output_filename}")

if __name__ == "__main__":
    validate_item_master()