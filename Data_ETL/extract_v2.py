import pandas as pd
import os
import logging
from pathlib import Path

def detect_pii(df):
    pii_patterns = {
        'name_columns': ['full_name', 'fullname', 'employeefullname', 'name'],
        'empid_columns': ['emplid', 'empid', 'empid#', 'idnumber', 'emp#', 'id_number', 'empno'],
        'dob_columns': ['dob', 'dateofbirth', 'birthdate', 'dobirth', 'bdate'],
        'ssn_columns': ['ssn', 'socialsecurity', 'ssnumber', 'taxid', 'tax_id', 'socsec']
    }
    
    name_col = next((col for col in pii_patterns['name_columns'] if col in df.columns), None)
    id_col = next((col for col in pii_patterns['empid_columns'] if col in df.columns), None)
    dob_col = next((col for col in pii_patterns['dob_columns'] if col in df.columns), None)
    ssn_col = next((col for col in pii_patterns['ssn_columns'] if col in df.columns), None)
    
    mapping = {}
    if name_col:
        mapping[name_col] = 'Name'
    if id_col:
        mapping[id_col] = 'Emp ID'
    if dob_col:
        mapping[dob_col] = 'DOB'
    if ssn_col:
        mapping[ssn_col] = 'SSN'
    
    df_mapped = df.rename(columns=mapping)

    return df_mapped


def extract_from_excel(file_path):
    current_file_dict = pd.read_excel(file_path, sheet_name=None)
    all_row_list = []

    for sheet_name, df in current_file_dict.items():
        df.columns = df.columns.str.strip().str.lower().str.replace(" ","")

        df_result = detect_pii(df)
        df_result['Doc ID'] = file_path.name
        df_result['Sheet Name'] = sheet_name

        all_row_list.append(df_result)
    df_final = pd.concat(all_row_list, ignore_index=True)
    
    return df_final


""" def extract_from_pdf(file_path):
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
    return pd.DataFrame(data, dtype=str) """

def extract(folder_path):
    folder = Path(folder_path)
    supported_file_types = [".xls", ".xlsx", ".xlsm", ".xlsb", ".odf", ".ods", ".odt"]
    df = []

    for file_path in folder.iterdir():
        
        ext = os.path.splitext(file_path)[1].lower()
        if ext in supported_file_types:
            all_data = (extract_from_excel(file_path)) 

            df.append(all_data)
            
        else:
            raise ValueError(f"Unsupported file type: {ext}")
        
    df_final = pd.concat(df, ignore_index=True)
    
    return df_final