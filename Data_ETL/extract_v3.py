import pandas as pd
import pdfplumber
import xml.etree.ElementTree as ET
import re
import os
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

# ─────────────────────────────────────────────
# 1. EXCEL  (.xlsx / .xlsm / .xls / .ods)
# ─────────────────────────────────────────────

def read_excel(path: str, sheet_name=None) -> pd.DataFrame:
    """
    Reads an Excel file into a DataFrame.
    - .xlsx / .xlsm  → openpyxl (default)
    - .xls (legacy)  → xlrd
    - .ods           → odf
    """
    ext = os.path.splitext(path)[1].lstrip(".").lower()

    engine_map = {
        "xls":  "xlrd",
        "ods":  "odf",
        "xlsm": "openpyxl",
        "xlsx": "openpyxl",
    }
    engine = engine_map.get(ext, "openpyxl")

    current_file_dict = pd.read_excel(path, sheet_name=sheet_name, engine=engine)
    all_row_list = []

    for sheet_name, df in current_file_dict.items():
        df.columns = df.columns.str.strip().str.lower().str.replace(" ","")

        df_result = detect_pii(df)
        df_result['Doc ID'] = path.name # use file name as Doc ID
        df_result['Source'] = sheet_name

        all_row_list.append(df_result)
    df = pd.concat(all_row_list, ignore_index=True)
    
    print(f"[Excel] {df.shape[0]} rows × {df.shape[1]} cols  |  engine={engine}")
    return df


# ─────────────────────────────────────────────
# 2. CSV  (.csv / .tsv / pipe-delimited …)
# ─────────────────────────────────────────────

def read_csv(path: str, sep: str | None = None, encoding: str = "utf-8") -> pd.DataFrame:
    """
    Reads a CSV (or any delimiter-separated) file.
    Pass sep=None to let pandas sniff the delimiter automatically.
    """
    df = pd.read_csv(
        path,
        sep=sep,              # None → auto-detect
        engine="python",      # required for sep=None
        encoding=encoding,
        on_bad_lines="skip",  # skip corrupted rows
    )

    print(f"[CSV]   {df.shape[0]} rows × {df.shape[1]} cols  |  sep={repr(df.columns)[:40]}")

    df.columns = df.columns.str.strip().str.lower().str.replace(" ","")

    df_result = detect_pii(df)
    df_result['Doc ID'] = path.name # use file name as Doc ID
    df_result['Source'] = "csv"

    return df_result

# ─────────────────────────────────────────────
# 3. PDF  (.pdf)  — table extraction
# ─────────────────────────────────────────────

def read_pdf(path: str, pages: str = "all") -> pd.DataFrame:
    """
    Extracts tables from a PDF using pdfplumber.
    Merges all tables found across the requested pages.
    Falls back to raw text parsing if no tables are detected.
    """
    all_tables = []

    with pdfplumber.open(path) as pdf:
        page_range = pdf.pages if pages == "all" else [pdf.pages[p] for p in pages]

        for page in page_range:
            tables = page.extract_tables()
            for table in tables:
                if table:
                    df = pd.DataFrame(table[1:], columns=table[0])  # first row = header
                    all_tables.append(df)
                    all_tables['Doc ID'] = path.name
                    all_tables['Source'] = f"Page {page.page_number}"

    if all_tables:
        combined = pd.concat(all_tables, ignore_index=True)
        print(f"[PDF]   {combined.shape[0]} rows × {combined.shape[1]} cols  |  {len(all_tables)} table(s) found")
        return combined

    # ── Fallback: no tables found → extract raw text into a single-column DF ──
    print("[PDF]   No tables detected — falling back to raw text extraction")
    lines = []
    with pdfplumber.open(path) as pdf:
        for page in pdf.pages:
            text = page.extract_text()
            if text:
                for line in text.splitlines():
                    if line.strip():
                        lines.append({"text": line, "Doc ID": path.name})
                        lines.append({"text": line, "Source": f"Page {page.page_number}"})

    df_data = pd.DataFrame(lines)
    print(f"[PDF]   Extracted {df_data.shape[0]} lines of text into DataFrame")
    return df_data

# ─────────────────────────────────────────────
# 4. XML  (.xml)
# ─────────────────────────────────────────────

def read_xml(path: str, record_tag: str | None = None) -> pd.DataFrame:
    """
    Parses an XML file into a DataFrame.

    Strategy:
      1. pandas.read_xml()  — works great for flat, uniform XML.
      2. Manual ElementTree — fallback for nested / irregular XML;
         flattens each child element of `record_tag` into a row.

    Args:
        path:       Path to the .xml file.
        record_tag: Tag name that represents a single record, e.g. "item",
                    "row", "record". Auto-detected if None.
    """
    # ── Attempt 1: pandas native ───────────────────────────────────────────
    try:
        df = pd.read_xml(path)
        print(f"[XML]   {df.shape[0]} rows × {df.shape[1]} cols  |  pandas native")
        return df
    except Exception:
        pass  # fall through to manual parse

    # ── Attempt 2: ElementTree — flatten children of record_tag ───────────
    tree = ET.parse(path)
    root = tree.getroot()

    # Auto-detect record tag: most-common direct child tag
    if record_tag is None:
        tag_counts: dict[str, int] = {}
        for child in root:
            tag = re.sub(r"\{.*?\}", "", child.tag)  # strip namespace
            tag_counts[tag] = tag_counts.get(tag, 0) + 1
        record_tag = max(tag_counts, key=tag_counts.get)
        print(f"[XML]   Auto-detected record tag: <{record_tag}>")

    records = []
    for elem in root.iter(record_tag):
        row: dict = {**elem.attrib}                   # element attributes
        for child in elem:
            tag = re.sub(r"\{.*?\}", "", child.tag)   # strip namespace
            row[tag] = child.text
        records.append(row)

    df = pd.DataFrame(records)
    print(f"[XML]   {df.shape[0]} rows × {df.shape[1]} cols  |  ElementTree fallback")
    df.columns = df.columns.str.strip().str.lower().str.replace(" ","")

    df_result = detect_pii(df)
    df_result['Doc ID'] = path.name # use file name as Doc ID

    return df_result

def read_txt(path: str) -> pd.DataFrame:
    file_path = Path(path)
    
    with open(file_path, "r", encoding="utf-8") as f:
        lines = [line.strip() for line in f.readlines() if line.strip()]
    
    df = pd.DataFrame({
        "text":     lines,
        "Doc ID":   file_path.stem,
        "Source":   "text"
    })

    print(f"[TXT]   {df.shape[0]} rows × {df.shape[1]} cols  |  raw text lines")
    
    return df

# ─────────────────────────────────────────────
# 5. UNIVERSAL DISPATCHER
# ─────────────────────────────────────────────

def read_file(path: str, **kwargs) -> pd.DataFrame:
    """Auto-routes to the right reader based on file extension."""
    # ext = os.path.split(".", 1)[-1].lower()
    ext = os.path.splitext(path)[1].lstrip(".").lower()
    dispatch = {
        "xlsx": read_excel,
        "xlsm": read_excel,
        "xls":  read_excel,
        "ods":  read_excel,
        "csv":  read_csv,
        "tsv":  lambda p, **kw: read_csv(p, sep="\t", **kw),
        "pdf":  read_pdf,
        "xml":  read_xml,
        "txt":  read_txt,
    }
    reader = dispatch.get(ext)
    if reader is None:
        raise ValueError(f"Unsupported extension: .{ext}")
    return reader(path, **kwargs)


def extract(folder_path):
    folder = Path(folder_path)
    df = []

    for file_path in folder.iterdir():
        
        all_data = (read_file(file_path)) 
        print(f"Extracted {all_data.shape[0]} rows from {file_path.name}")
        df.append(all_data)
 
          
    df_final = pd.concat(df, ignore_index=True)
    
    return df_final
