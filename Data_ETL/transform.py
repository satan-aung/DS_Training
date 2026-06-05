import pandas as pd
import logging
import re
from datetime import datetime

def detect_pii(df):
    pii_patterns = {
        'name_columns': ['full_name', 'fullname', 'employeefullname' 'name'],
        'empid_columns': ['emplid', 'empid', 'empid#', 'idnumber', 'emp#', 'id_number'],
        'dob_columns': ['dob', 'dateofbirth', 'birthdate', 'dobirth'],
        'ssn_columns': ['ssn', 'socialsecurity', 'ssnumber', 'taxid', 'tax_id']
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
    standard_cols = ['Name','Emp ID', 'DOB', 'SSN']
    df_final = df_mapped[[c for c in standard_cols if c in df_mapped.columns]]

    if df_final.empty:
        raise ValueError("No PII columns detected")
    else:
        if 'Name' not in df_final:
            if 'SSN' not in df_final:
                logging.info("No name columns detected")
            else:
                logging.info("SSN with no name columns")
    
    return df_final

def clean_name(name):
    if pd.isna(name):
        return None
    name = str(name).strip()
    name = re.sub(r'\s+', ' ', name)
    return name.title()

def clean_id(emp_id):
    if pd.isna(emp_id):
        return None
    emp_id = str(emp_id).strip()
    if emp_id.isdigit():
        return emp_id
    else:
        return emp_id.upper()

def clean_dob(dob, df_row=None, notes_column='Notes'):
    if pd.isna(dob):
        return None
    try:
        date_obj = pd.to_datetime(dob, errors='coerce')
        if pd.notna(date_obj) and date_obj.year <= datetime.now().year:
            return date_obj.strftime('%m/%d/%Y')
    except:
        pass
    return None

def clean_ssn(ssn):
    if pd.isna(ssn):
        return None
    digits = re.sub(r'\D', '', str(ssn))
    if len(digits) == 9:
        if digits[:3] not in ['000', '666'] and digits[3:5] != '00' and digits[5:] != '0000':
            return f"{digits[:3]}-{digits[3:5]}-{digits[5:9]}"
    return None

def clean_data(df):

    if 'Name' in df.columns:
        df['Name'] = df['Name'].apply(clean_name)
    if 'Emp ID' in df.columns:
        df['Emp ID'] = df['Emp ID'].apply(clean_id)
    if 'DOB' in df.columns:
        df['DOB'] = df['DOB'].apply(clean_dob)
    if 'SSN' in df.columns:
        df['SSN'] = df['SSN'].apply(clean_ssn)

    logging.info(f"Rows after cleaning: {len(df)}")

    return df

def transform_data(df):
    try:
        logging.info("Starting detection")

        df.columns = df.columns.str.strip().str.lower().str.replace(" ","")
        df = detect_pii(df)
        df = clean_data(df)
        return df

    except Exception as e:
        logging.error(f"Transformation failed: {e}")
        raise