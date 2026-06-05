import pandas as pd
import numpy as np
from datetime import datetime, timedelta
import random
from faker import Faker

# Initialize Faker for realistic data
fake = Faker()

# Set random seed for reproducibility
np.random.seed(42)
random.seed(42)

# ============================================================
# FILE 1: HR System Export (Messy Column Names)
# ============================================================
def generate_hr_messy_data(rows=120):
    """Generate messy HR data with inconsistent column names"""
    
    data = {
        'Employee Full Name': [fake.name() for _ in range(rows)],
        'Emp ID #': [f"EMP{random.randint(10000, 99999)}" for _ in range(rows)],
        'Date of Birth (DOB)': [],
        'SSN (Social Security)': [],
        'Hire Date': [],
        'Annual Salary $': [],
        'Dept': [],
        'Email Address': [],
        'Phone #': [],
        'Status': []
    }
    
    for _ in range(rows):
        # Random dates with some errors
        year = random.choice([random.randint(1950, 2000), 1880, 2025, 1800])
        month = random.randint(1, 12)
        day = random.randint(1, 28)
        data['Date of Birth (DOB)'].append(f"{month}/{day}/{year}")
        
        # SSN with different formats
        ssn_format = random.choice(['ddd-dd-dddd', 'ddddddddd', 'ddd.dd.dddd', 'ddd dd dddd'])
        if ssn_format == 'ddd-dd-dddd':
            ssn = f"{random.randint(100, 999)}-{random.randint(10, 99)}-{random.randint(1000, 9999)}"
        elif ssn_format == 'ddddddddd':
            ssn = f"{random.randint(100000000, 999999999)}"
        elif ssn_format == 'ddd.dd.dddd':
            ssn = f"{random.randint(100, 999)}.{random.randint(10, 99)}.{random.randint(1000, 9999)}"
        else:
            ssn = f"{random.randint(100, 999)} {random.randint(10, 99)} {random.randint(1000, 9999)}"
        data['SSN (Social Security)'].append(ssn)
        
        # Hire date with various formats
        hire_format = random.choice(['YYYY-MM-DD', 'MM/DD/YYYY', 'DD-Mon-YYYY', 'YYYYMMDD'])
        hire_date = fake.date_between(start_date='-20y', end_date='today')
        if hire_format == 'YYYY-MM-DD':
            data['Hire Date'].append(hire_date.strftime('%Y-%m-%d'))
        elif hire_format == 'MM/DD/YYYY':
            data['Hire Date'].append(hire_date.strftime('%m/%d/%Y'))
        elif hire_format == 'DD-Mon-YYYY':
            data['Hire Date'].append(hire_date.strftime('%d-%b-%Y'))
        else:
            data['Hire Date'].append(hire_date.strftime('%Y%m%d'))
        
        # Salary with currency symbols
        salary = random.randint(30000, 150000)
        salary_format = random.choice(['$' + str(salary), str(salary) + ' USD', str(salary), f"${salary:,}"])
        data['Annual Salary $'].append(salary_format)
        
        data['Dept'].append(random.choice(['IT', 'HR', 'Finance', 'Marketing', 'Sales', 'Operations', 'R&D', None]))
        data['Email Address'].append(fake.email())
        data['Phone #'].append(fake.phone_number())
        data['Status'].append(random.choice(['Active', 'Inactive', 'Terminated', 'On Leave', '', None]))
    
    # Add some empty rows
    for i in random.sample(range(rows), 10):
        for col in ['Employee Full Name', 'Emp ID #']:
            data[col][i] = ''
    
    # Add some duplicate headers in data (messy)
    data[''] = [np.nan] * rows  # empty column
    data['Unnamed: 1'] = [np.nan] * rows
    
    return pd.DataFrame(data)

# ============================================================
# FILE 2: Payroll System Export (Different Naming Convention)
# ============================================================
def generate_payroll_messy_data(rows=150):
    """Generate messy payroll data with different column names"""
    
    data = {
        'fullname': [fake.name() for _ in range(rows)],
        'id_number': [f"ID{random.randint(10000, 99999)}" for _ in range(rows)],
        'birthdate': [],
        'tax_id': [],
        'start_date': [],
        'salary_amount': [],
        'department_name': [],
        'email': [],
        'contact_no': [],
        'employment_type': []
    }
    
    for _ in range(rows):
        # Birthdate with errors
        year = random.choice([random.randint(1960, 2002), 1899, 2030, 1776])
        data['birthdate'].append(f"{year}-{random.randint(1, 12)}-{random.randint(1, 28)}")
        
        # Tax ID with various formats
        tax_format = random.choice(['###-##-####', '##########', '### ## ####'])
        if tax_format == '###-##-####':
            tax_id = f"{random.randint(100, 999)}-{random.randint(10, 99)}-{random.randint(1000, 9999)}"
        elif tax_format == '##########':
            tax_id = f"{random.randint(1000000000, 9999999999)}"
        else:
            tax_id = f"{random.randint(100, 999)} {random.randint(10, 99)} {random.randint(1000, 9999)}"
        data['tax_id'].append(tax_id)
        
        data['start_date'].append(fake.date_between(start_date='-15y', end_date='today').strftime('%d/%m/%Y'))
        data['salary_amount'].append(random.randint(35000, 200000))
        data['department_name'].append(random.choice(['IT', 'HR', 'Finance', 'Marketing', 'Sales']))
        data['email'].append(fake.email())
        data['contact_no'].append(fake.phone_number())
        data['employment_type'].append(random.choice(['Full-time', 'Part-time', 'Contract', 'Temporary']))
    
    # Add some data errors
    for i in random.sample(range(rows), 15):
        data['fullname'][i] = ''
        data['tax_id'][i] = 'invalid_ssn'
    
    # Add messy column names with spaces
    data[' extra space '] = [random.choice(['A', 'B', 'C']) for _ in range(rows)]
    data['MESSY!@#'] = [random.randint(1, 100) for _ in range(rows)]
    
    return pd.DataFrame(data)

# ============================================================
# FILE 3: Legacy System Export (Inconsistent & Abbreviated)
# ============================================================
def generate_legacy_messy_data(rows=110):
    """Generate messy legacy data with abbreviations and inconsistencies"""
    
    first_names = ['John', 'Jane', 'Bob', 'Alice', 'Charlie', 'Diana', 'Edward', 'Fiona', 'George', 'Hannah']
    last_names = ['Smith', 'Johnson', 'Williams', 'Brown', 'Jones', 'Garcia', 'Miller', 'Davis', 'Rodriguez', 'Martinez']
    
    data = {
        'nm': [f"{random.choice(first_names)} {random.choice(last_names)}" for _ in range(rows)],
        'empno': [f"E{random.randint(1000, 9999)}" for _ in range(rows)],
        'bdate': [],
        'socsec': [],
        'hdate': [],
        'sal': [],
        'dept': [],
        'eml': [],
        'ph': [],
        'stat': []
    }
    
    for _ in range(rows):
        # Birthdate with DD/MM/YYYY format (European)
        day = random.randint(1, 28)
        month = random.randint(1, 12)
        year = random.choice([random.randint(1950, 2002), 1901, 2026])
        data['bdate'].append(f"{day}/{month}/{year}")
        
        # SSN with missing digits
        ssn_len = random.choice([9, 8, 10, 11])
        if ssn_len == 9:
            ssn = f"{random.randint(100, 999)}-{random.randint(10, 99)}-{random.randint(1000, 9999)}"
        elif ssn_len == 8:
            ssn = f"{random.randint(100, 999)}-{random.randint(10, 99)}-{random.randint(100, 999)}"
        else:
            ssn = f"{random.randint(1000, 9999)}-{random.randint(100, 999)}-{random.randint(100, 999)}"
        data['socsec'].append(ssn)
        
        data['hdate'].append(fake.date_between(start_date='-10y', end_date='today').strftime('%Y%m%d'))
        data['sal'].append(f"${random.randint(40000, 180000)}")
        data['dept'].append(random.choice(['IT', 'HR', 'FIN', 'MKT', 'SALES', 'OPS']))
        data['eml'].append(fake.email())
        data['ph'].append(fake.phone_number())
        data['stat'].append(random.choice(['A', 'I', 'T', 'L', '']))
    
    # Add some messy data
    for i in random.sample(range(rows), 20):
        if random.choice([True, False]):
            data['nm'][i] = ''
        else:
            data['socsec'][i] = 'N/A'
    
    # Add duplicate and empty columns
    data[''] = [np.nan] * rows
    data['another column'] = [random.random() for _ in range(rows)]
    data['   '] = [np.nan] * rows
    
    return pd.DataFrame(data)

# ============================================================
# Generate and Save All Files
# ============================================================

print("="*70)
print("GENERATING MESSY DATA FILES")
print("="*70)

# Generate File 1: HR System Export
print("\n📁 Generating File 1: HR_System_Export.xlsx (120 rows)")
df1 = generate_hr_messy_data(120)
df1.to_excel('HR_System_Export.xlsx', index=False)
print(f"   ✓ Saved: HR_System_Export.xlsx")
print(f"   ✓ Shape: {df1.shape}")
print(f"   ✓ Columns: {df1.columns.tolist()[:5]}... (total {len(df1.columns)} columns)")

# Generate File 2: Payroll System Export
print("\n📁 Generating File 2: Payroll_System_Export.xlsx (150 rows)")
df2 = generate_payroll_messy_data(150)
df2.to_excel('Payroll_System_Export.xlsx', index=False)
print(f"   ✓ Saved: Payroll_System_Export.xlsx")
print(f"   ✓ Shape: {df2.shape}")
print(f"   ✓ Columns: {df2.columns.tolist()[:5]}... (total {len(df2.columns)} columns)")

# Generate File 3: Legacy_System_Export
print("\n📁 Generating File 3: Legacy_System_Export.xlsx (110 rows)")
df3 = generate_legacy_messy_data(110)
df3.to_excel('Legacy_System_Export.xlsx', index=False)
print(f"   ✓ Saved: Legacy_System_Export.xlsx")
print(f"   ✓ Shape: {df3.shape}")
print(f"   ✓ Columns: {df3.columns.tolist()[:5]}... (total {len(df3.columns)} columns)")

print("\n" + "="*70)
print("✅ ALL FILES GENERATED SUCCESSFULLY!")
print("="*70)
print("\n📋 Files created:")
print("   1. HR_System_Export.xlsx (120 rows)")
print("   2. Payroll_System_Export.xlsx (150 rows)")
print("   3. Legacy_System_Export.xlsx (110 rows)")
print("\n💡 Each file contains messy data with:")
print("   - Different column naming conventions")
print("   - Various date formats")
print("   - SSNs in different formats")
print("   - Missing/empty values")
print("   - Special characters in column names")
print("   - Invalid dates (future, very old)")
print("   - Extra whitespace columns")

# ============================================================
# Preview the generated files
# ============================================================

print("\n" + "="*70)
print("DATA PREVIEW")
print("="*70)

print("\n📊 HR_System_Export.xlsx - First 5 rows:")
print(df1.head().to_string())

print("\n📊 Payroll_System_Export.xlsx - First 5 rows:")
print(df2.head().to_string())

print("\n📊 Legacy_System_Export.xlsx - First 5 rows:")
print(df3.head().to_string())

# ============================================================
# Optional: Create a ZIP file with all three
# ============================================================
import zipfile
import os

def create_zip():
    """Create a zip file containing all three Excel files"""
    with zipfile.ZipFile('Messy_Data_Files.zip', 'w') as zipf:
        zipf.write('HR_System_Export.xlsx')
        zipf.write('Payroll_System_Export.xlsx')
        zipf.write('Legacy_System_Export.xlsx')
    
    print("\n📦 Created ZIP file: Messy_Data_Files.zip")
    print("   (Contains all 3 Excel files)")

create_zip()

print("\n" + "="*70)
print("🎉 DONE! Files are ready for download")
print("="*70)