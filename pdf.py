import pandas as pd
import os
import shutil
from pathlib import Path

def organize_pdfs_by_brand():
    # ==========================================
    # 1. CONFIGURATION (Update your paths here)
    # ==========================================
    excel_path = r'C:\Users\m.emaan\OneDrive - Casa Mia LLC\Documents\Extraction_task\Item_Master.xlsx'
    input_folder = r'C:\Users\m.emaan\OneDrive - Casa Mia LLC\Documents\Extraction_task\brands\Gessi_duplicated_pdfs'
    
    # This is where the new Brand folders and matched PDFs will be saved
    output_folder = r'C:\Users\m.emaan\OneDrive - Casa Mia LLC\Documents\Extraction_task\output' 
    
    # Report filename
    report_filename = 'brand_match_summary.csv'

    # ==========================================
    # 2. LOAD AND MAP THE ITEM MASTER
    # ==========================================
    print("Loading Item Master...")
    try:
        df = pd.read_excel(excel_path)
    except Exception as e:
        print(f"Error loading Excel file: {e}")
        return

    # Assuming Column A (Index 0) is Code, Column D (Index 3) is Brand
    # Drop any rows where the Code or Brand is missing to prevent errors
    df_clean = df.dropna(subset=[df.columns[0], df.columns[3]])

    # Create a dictionary mapping: {'Code': 'Brand'}
    # .astype(str).str.strip() ensures we don't have hidden spaces ruining the match
    codes = df_clean.iloc[:, 0].astype(str).str.strip()
    brands = df_clean.iloc[:, 3].astype(str).str.strip()
    
    code_to_brand = dict(zip(codes, brands))
    print(f"Mapped {len(code_to_brand)} unique products to their brands.")

    # ==========================================
    # 3. SCAN FOLDERS AND COPY MATCHED PDFS
    # ==========================================
    print("Scanning input folders for PDFs...")
    
    # Create the main output directory if it doesn't exist
    os.makedirs(output_folder, exist_ok=True)
    
    # Dictionary to keep track of how many files matched per brand
    brand_stats = {}
    files_processed = 0

    # Walk through every sub-folder in the input directory
    for root, dirs, files in os.walk(input_folder):
        for file in files:
            # We only care about PDFs
            if file.lower().endswith('.pdf'):
                # Extract the code (filename without '.pdf')
                code = Path(file).stem.strip()
                
                # Check if this PDF code exists in our Excel mapping
                if code in code_to_brand:
                    brand = code_to_brand[code]
                    
                    # 1. Create the specific brand folder inside the output folder
                    brand_out_dir = os.path.join(output_folder, brand)
                    os.makedirs(brand_out_dir, exist_ok=True)
                    
                    # 2. Define source and destination paths
                    src_path = os.path.join(root, file)
                    dst_path = os.path.join(brand_out_dir, file)
                    
                    # 3. Copy the file (only if it hasn't been copied already)
                    if not os.path.exists(dst_path):
                        shutil.copy2(src_path, dst_path)
                        
                        # Update our counting statistics
                        brand_stats[brand] = brand_stats.get(brand, 0) + 1
                        files_processed += 1

    print(f"\nFinished processing. {files_processed} matched PDFs successfully copied.")

    # ==========================================
    # 4. GENERATE THE SUMMARY REPORT
    # ==========================================
    if brand_stats:
        # Convert our tracking dictionary into a clean Pandas DataFrame
        report_df = pd.DataFrame(list(brand_stats.items()), columns=['Brand', 'Matched PDFs Saved'])
        
        # Sort it alphabetically by Brand for easier reading
        report_df = report_df.sort_values(by='Brand')
        
        # Save to CSV in the output folder
        report_path = os.path.join(output_folder, report_filename)
        report_df.to_csv(report_path, index=False)
        print(f"Match summary report saved to: {report_path}")
    else:
        print("No matches were found. Report not generated.")

if __name__ == "__main__":
    organize_pdfs_by_brand()