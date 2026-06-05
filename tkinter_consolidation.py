import pandas as pd
import os
import glob

def process_files(input_folder, reference_template_path, output_prefix="consolidated_file"):
    # 1. Load Reference Headers
    ref_df = pd.read_excel(reference_template_path) # Or pd.read_csv
    ref_headers = list(ref_df.columns)
    
    all_data = []
    
    # 2. Get all files in directory
    files = glob.glob(os.path.join(input_folder, "*.xlsx")) # Change extension if needed
    
    for file in files:
        if file == reference_template_path:
            continue
            
        df = pd.read_excel(file)
        
        # 3. Header Validation
        if list(df.columns) != ref_headers:
            print(f"Error: Headers do not match in file: {file}")
            continue
        
        all_data.append(df)
        print(f"Successfully validated: {file}")

    if not all_data:
        print("No valid files to process.")
        return

    # 4. Data Consolidation
    consolidated_df = pd.concat(all_data, ignore_index=True)
    
    # 5. Split and Save
    max_rows = 1000000
    file_count = 1
    
    for i in range(0, len(consolidated_df), max_rows):
        chunk = consolidated_df.iloc[i : i + max_rows]
        output_filename = f"{output_prefix}_{file_count}.xlsx"
        
        chunk.to_excel(output_filename, index=False)
        print(f"Saved: {output_filename} with {len(chunk)} rows.")
        file_count += 1

# Execution
process_files(
    input_folder=r"C:\Users\Aung\Documents\Templates", 
    reference_template_path=r"C:\Users\Aung\Documents\Reference.xlsx"

)
