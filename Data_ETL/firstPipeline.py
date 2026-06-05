import pandas as pd
import pdfplumber
import logging
import os
import re
from pathlib import Path

logging.basicConfig(level=logging.INFO,
                    format='%(asctime)s - %(levelname)s - %(message)s')
logger = logging.getLogger(__name__)

def extract_from_csv(file_path):
    return pd.read_csv(file_path, dtype=str)

def extract_from_excel(file_path):
    return pd.read_excel(file_path, dtype=str)

def extract_from_pdf(file_path):
    data = []
    with pdfplumber.open(file_path) as pdf:
        for page in pdf.pages:
            tables = page.extract_tables()
            for table in tables:
                if table:
                    headers = [h.strip() if h else f"col_{i}" for i, h in enumerate(table[0])]
                    for row in table[1:]:
                        if any(row):
                            data.append(dict(zip(headers, row)))
    return pd.DataFrame(data, dtype=str)

def extract(file_path):
    ext = os.path.splitext(file_path)[1].lower()
    if ext == '.csv':
        return extract_from_csv(file_path)
    elif ext == '.xlsx':
        return extract_from_excel(file_path)
    elif ext == '.pdf':
        return extract_from_pdf(file_path)
    else:
        raise ValueError(f"Unsupported file type: {ext}")

def clean_employee_name(name):
    if pd.isna(name):
        return None
    name = str(name).strip()
    name = re.sub(r'\s+', ' ', name)
    return name.title()

def clean_employee_id(emp_id):
    if pd.isna(emp_id):
        return None
    emp_id = str(emp_id).strip()
    if emp_id.isdigit():
        return emp_id
    else:
        return emp_id.upper()

def normalize_dataframe(df):
    df.columns = [col.strip() for col in df.columns]
    name_col = None
    id_col = None
    for col in df.columns:
        if col.lower() == 'employee name':
            name_col = col
        elif col.lower() == 'empl id':
            id_col = col
    
    if name_col is None or id_col is None:
        raise KeyError(f"Required columns not found. Found: {df.columns.tolist()}")
    
    df['Employee_Name_Cleaned'] = df[name_col].apply(clean_employee_name)
    df['Empl_ID_Cleaned'] = df[id_col].apply(clean_employee_id)

    df = df.dropna(subset=['Employee_Name_Cleaned', 'Empl_ID_Cleaned'], how='all')

    for col in df.select_dtypes(include=['object']).columns:
        df[col] = df[col].astype(str).str.strip()
        df[col] = df[col].replace('', None)
    
    return df

def load_to_csv(df, output_path, format='csv', **kwargs):

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

def run_pipeline_to_csv(input_file, output_csv_path):

    logger.info(f"Starting ETL pipeline: {input_file} -> {output_csv_path}")
    
    df_raw = extract(input_file)
    logger.info(f"Extracted {len(df_raw)} rows from {input_file}")
    
    df_clean = normalize_dataframe(df_raw)
    logger.info(f"Cleaned data: {len(df_clean)} rows remaining")
    
    output_df = df_clean

    cleaned_columns = ['Employee_Name_Cleaned', 'Empl_ID_Cleaned']
    output_df = df_clean[cleaned_columns]
    
    load_to_csv(output_df, output_csv_path, format='csv')
    
    logger.info("Pipeline completed successfully!")
    return output_df

def dataframe_to_csv(df, output_path, **kwargs):
    
    output_dir = os.path.dirname(output_path)
    if output_dir and not os.path.exists(output_dir):
        os.makedirs(output_dir)
    
    df.to_csv(output_path, index=False, **kwargs)
    print(f"✅ Data saved to {output_path} ({len(df)} rows, {len(df.columns)} columns)")
    return output_path

if __name__ == "__main__":
    import sys
    
    if len(sys.argv) < 2:
        print("Usage: python pipeline.py <input_file> [output_csv_path]")
        print("Example: python pipeline.py employees.xlsx output/cleaned_data.csv")
        sys.exit(1)
    
    input_file = sys.argv[1]
    
    if len(sys.argv) >= 3:
        output_path = sys.argv[2]
    else:
        input_name = Path(input_file).stem
        output_path = f"cleaned_{input_name}.csv"
    
    if not os.path.exists(input_file):
        print(f"Error: File not found - {input_file}")
        sys.exit(1)
    
    result_df = run_pipeline_to_csv(input_file, output_path)
    
    print("\n" + "="*50)
    print(f"✅ ETL Pipeline Complete!")
    print(f"📁 Input: {input_file}")
    print(f"💾 Output: {output_path}")
    print(f"📊 Rows: {len(result_df)}")
    print(f"📋 Columns: {list(result_df.columns)}")
    print("="*50)