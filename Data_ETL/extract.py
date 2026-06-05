import pandas as pd
import os
import pdfplumber

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