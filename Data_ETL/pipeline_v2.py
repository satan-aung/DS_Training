import pandas as pd
import logging
import os
from pathlib import Path
from Data_ETL.extract_v3 import extract
from Data_ETL.transform_v2 import transform_data
# from upload_database import upload_with_error_handling, db_config
# from read_database import read_data

logging.basicConfig(level=logging.INFO,
                    format='%(asctime)s - %(levelname)s - %(message)s')
logger = logging.getLogger(__name__)

def load_to_file(df, output_path, format='csv', **kwargs):

    output_dir = os.path.dirname(output_path)
    if output_dir and not os.path.exists(output_dir):
        os.makedirs(output_dir)
        logger.info(f"Created directory: {output_dir}")
    
    if format == 'csv':
        df.to_csv(output_path, index=False, **kwargs)
        logger.info(f"Saved CSV to: {output_path} (rows: {len(df)})")
    
    elif format == 'excel':
        df.to_excel(output_path, index=False, **kwargs)
        logger.info(f"Saved Excel to: {output_path} (rows: {len(df)})")
    
    elif format == 'json':
        df.to_json(output_path, orient='records', indent=2, **kwargs)
        logger.info(f"Saved JSON to: {output_path} (rows: {len(df)})")
    
    else:
        raise ValueError(f"Unsupported format: {format}")
    
    return output_path

def add_unique_id_simple(df, id_column='BreachID', prefix='DS', start=1):
    
    df_copy = df.copy()
    
    if prefix:
        df_copy[id_column] = [f"{prefix}{i:08d}" for i in range(start, start + len(df_copy))]
    else:
        df_copy[id_column] = range(start, start + len(df_copy))
    
    return df_copy

def run_pipeline(input_folder, output_path):

    logger.info(f"Starting ETL pipeline: {input_folder} -> {output_path}")
    
    df_raw = extract(input_folder)
    logger.info(f"Extracted {len(df_raw)} rows from {input_folder}")
    
    df_clean = transform_data(df_raw)
    logger.info(f"Cleaned data: {len(df_clean)} rows remaining")
    
    # output_df = df_clean
    # rows_to_update = []
    # rows_to_remove = []

    # output_df_clean = output_df.copy()
    # if 'Name' in output_df_clean.columns:
    #     if 'SSN' in output_df_clean.columns:
    #         name_missing = output_df_clean['Name'].isna() | (output_df_clean['Name'].astype(str).str.strip() == '')
    #         empty_rows = pd.Series([True] * len(output_df_clean))
    #         ssn_missing = output_df_clean['SSN'].isna() | (output_df_clean['SSN'].astype(str).str.strip() == '')
    #         rows_to_remove = name_missing & ssn_missing
    #         empty_rows = empty_rows & rows_to_remove
            
    #         ssn_present = output_df_clean['SSN'].notna() & (output_df_clean['SSN'].astype(str).str.strip() != '')
    
    #         rows_to_update = name_missing & ssn_present

    #         output_df_clean = output_df_clean[~empty_rows]
    # else:
    #     ssn_present = output_df_clean['SSN'].notna() & (output_df_clean['SSN'].astype(str).str.strip() != '')
    #     rows_to_update = ssn_present
    
    # if rows_to_update.sum() > 0:
    #     output_df_clean.loc[rows_to_update, 'Name'] = 'Unknown, Unknown'
    #     output_df_clean.loc[rows_to_update, 'Emp ID'] = ''
    #     output_df_clean.loc[rows_to_update, 'DOB'] = ''
    
    #     print(f"Updated {rows_to_update.sum()} rows with 'Unknown, Unknown' in {'Name'}")

    df_addedid = add_unique_id_simple(df_clean, id_column='Breach ID', prefix='DS', start=1)

    final_columns = ['Doc ID', 'Source', 'Breach ID', 'Name', 'Emp ID', 'DOB', 'SSN']
    df_standard = df_addedid[final_columns]

    if df_standard.empty:
        raise ValueError("No PII columns detected")
    else:
        if 'Name' not in df_standard:
            if 'SSN' not in df_standard:
                logging.info("No name columns detected")
            else:
                logging.info("SSN with no name columns")

    empty_cols = []
    for col in df_standard.columns:
        if df_standard[col].isna().all():
            empty_cols.append(col)
        elif (df_standard[col].astype(str).str.strip() == '').all():
            empty_cols.append(col)
        
    if empty_cols:
        df_standard = df_standard.drop(columns=empty_cols)
        print(f"Removed {len(empty_cols)} empty columns: {empty_cols}")

    # columns_to_check = ['Emp ID', 'DOB', 'SSN']
    # df_copy = df_standard.copy()
    # original_rows = len(df_copy)
    
    # all_empty = pd.Series([True] * len(df_copy))
    
    # for col in columns_to_check:
    #     if col in df_copy.columns:
    #         col_empty = (df_copy[col].isna() | 
    #                         (df_copy[col].astype(str).str.strip() == '') |
    #                         (df_copy[col].astype(str).str.strip() == 'nan'))
    #         all_empty = all_empty & col_empty
    
    # df_final = df_copy[~all_empty.values]
    # rows_removed = original_rows - len(df_clean)
    
    # print(f"Removed {rows_removed} rows where ALL of {columns_to_check} were empty")
    # print(f"Remaining: {len(df_final)} rows")

    load_to_file(df_standard, output_path, format='csv')
    
    logger.info("Pipeline completed successfully!")
    return df_standard

if __name__ == "__main__":
    import sys
    
    if len(sys.argv) < 2:
        print("Usage: python pipeline.py <input_folder> [output_csv_path]")
        print("Example: python pipeline.py test_data output/cleaned_data.csv")
        sys.exit(1)
    
    input_folder = sys.argv[1]
    
    if len(sys.argv) >= 3:
        output_path = sys.argv[2]
    else:
        input_name = Path(input_folder).stem
        output_path = f"cleaned_{input_name}.csv"
    
    if not os.path.exists(input_folder):
        print(f"Error: Folder not found - {input_folder}")
        sys.exit(1)
    
    result_df = run_pipeline(input_folder, output_path)

    # success = upload_with_error_handling(result_df, 'template_data', db_config, if_exists='replace')
    # read_data()
    
    print("\n" + "="*50)
    print(f"ETL Pipeline Complete!")
    print(f"Input: {input_folder}")
    print(f"Output: {output_path}")
    print(f"Rows: {len(result_df)}")
    print(f"Columns: {list(result_df.columns)}")
    print("="*50)