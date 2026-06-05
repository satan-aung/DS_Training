"""
extract_v3.py — File extraction layer for the ETL pipeline.

Supported formats: .xlsx / .xlsm / .xls / .ods / .csv / .tsv / .pdf / .xml / .txt

Bugs fixed vs previous version:
  - read_pdf: all_tables['Doc ID'] was assigning to a list, not the DataFrame.
  - read_pdf fallback: two separate appends per line produced duplicate rows.
  - read_xml (native path): detect_pii() and Doc ID were never applied.
  - read_excel: sheet_name loop variable shadowed the parameter; single-sheet
    returns (non-dict) from pd.read_excel now handled correctly.
  - All readers now normalise `path` to a Path object for consistent .name access.
"""

import re
import os
import xml.etree.ElementTree as ET
from pathlib import Path
from typing import Optional, Union

import pandas as pd
import pdfplumber


# ─────────────────────────────────────────────────────────────────────────────
# PII column detection
# ─────────────────────────────────────────────────────────────────────────────

def detect_pii(df: pd.DataFrame) -> pd.DataFrame:
    """
    Detects PII columns by fuzzy name matching and renames them to canonical
    labels: Name, Emp ID, DOB, SSN.
    """
    pii_patterns: dict[str, list[str]] = {
        "Name":   ["full_name", "fullname", "employeefullname", "name"],
        "Emp ID": ["emplid", "empid", "empid#", "idnumber", "emp#", "id_number", "empno"],
        "DOB":    ["dob", "dateofbirth", "birthdate", "dobirth", "bdate"],
        "SSN":    ["ssn", "socialsecurity", "ssnumber", "taxid", "tax_id", "socsec"],
    }

    col_set = set(df.columns)
    mapping: dict[str, str] = {}
    for canonical, aliases in pii_patterns.items():
        match = next((a for a in aliases if a in col_set), None)
        if match:
            mapping[match] = canonical

    return df.rename(columns=mapping)


# ─────────────────────────────────────────────────────────────────────────────
# 1. EXCEL  (.xlsx / .xlsm / .xls / .ods)
# ─────────────────────────────────────────────────────────────────────────────

def read_excel(
    path: Union[str, Path],
    sheet_name: Optional[str] = None,
) -> pd.DataFrame:
    """
    Reads an Excel file into a DataFrame.
      - .xlsx / .xlsm  → openpyxl
      - .xls (legacy)  → xlrd
      - .ods           → odf
    Reads all sheets when sheet_name is None (default).
    """
    path = Path(path)
    ext = path.suffix.lstrip(".").lower()

    engine_map = {"xls": "xlrd", "ods": "odf", "xlsm": "openpyxl", "xlsx": "openpyxl"}
    engine = engine_map.get(ext, "openpyxl")

    result = pd.read_excel(path, sheet_name=sheet_name, engine=engine)

    # pd.read_excel returns a dict when sheet_name=None, else a single DataFrame.
    # FIX: handle both cases instead of assuming dict.
    if isinstance(result, pd.DataFrame):
        sheets: dict[str, pd.DataFrame] = {sheet_name or path.stem: result}
    else:
        sheets = result  # type: ignore[assignment]

    all_rows: list[pd.DataFrame] = []
    for sname, df in sheets.items():  # FIX: renamed loop var to avoid shadowing param
        df.columns = df.columns.str.strip().str.lower().str.replace(" ", "", regex=False)
        df = detect_pii(df)
        df["Doc ID"] = path.name
        df["Source"] = sname
        all_rows.append(df)

    combined = pd.concat(all_rows, ignore_index=True)
    print(f"[Excel] {combined.shape[0]} rows × {combined.shape[1]} cols  |  engine={engine}")
    return combined


# ─────────────────────────────────────────────────────────────────────────────
# 2. CSV  (.csv / .tsv / pipe-delimited …)
# ─────────────────────────────────────────────────────────────────────────────

def read_csv(
    path: Union[str, Path],
    sep: Optional[str] = None,
    encoding: str = "utf-8",
) -> pd.DataFrame:
    """
    Reads a CSV (or any delimiter-separated) file.
    Pass sep=None to let pandas sniff the delimiter automatically.
    """
    path = Path(path)

    df = pd.read_csv(
        path,
        sep=sep,
        engine="python",       # required for sep=None auto-detect
        encoding=encoding,
        on_bad_lines="skip",
    )

    df.columns = df.columns.str.strip().str.lower().str.replace(" ", "", regex=False)
    df = detect_pii(df)
    df["Doc ID"] = path.name
    df["Source"] = "csv"

    print(f"[CSV]   {df.shape[0]} rows × {df.shape[1]} cols")
    return df


# ─────────────────────────────────────────────────────────────────────────────
# 3. PDF  (.pdf)  — table extraction
# ─────────────────────────────────────────────────────────────────────────────

def read_pdf(path: Union[str, Path], pages: str = "all") -> pd.DataFrame:
    """
    Extracts tables from a PDF using pdfplumber.
    Merges all tables found across the requested pages.
    Falls back to raw text parsing if no tables are detected.
    """
    path = Path(path)
    all_tables: list[pd.DataFrame] = []

    with pdfplumber.open(path) as pdf:
        page_range = pdf.pages if pages == "all" else [pdf.pages[p] for p in pages]

        for page in page_range:
            for table in page.extract_tables():
                if not table:
                    continue
                df = pd.DataFrame(table[1:], columns=table[0])  # row 0 = header
                # FIX: set metadata on the DataFrame, not on the list
                df["Doc ID"] = path.name
                df["Source"] = f"Page {page.page_number}"
                all_tables.append(df)

    if all_tables:
        combined = pd.concat(all_tables, ignore_index=True)
        print(f"[PDF]   {combined.shape[0]} rows × {combined.shape[1]} cols  |  {len(all_tables)} table(s)")
        return combined

    # ── Fallback: no tables found → extract raw text ──────────────────────
    print("[PDF]   No tables detected — falling back to raw text extraction")
    lines: list[dict] = []
    with pdfplumber.open(path) as pdf:
        for page in pdf.pages:
            text = page.extract_text()
            if text:
                for line in text.splitlines():
                    if line.strip():
                        # FIX: single dict per line with all metadata fields together
                        lines.append({
                            "text":   line,
                            "Doc ID": path.name,
                            "Source": f"Page {page.page_number}",
                        })

    df = pd.DataFrame(lines)
    print(f"[PDF]   Extracted {df.shape[0]} text lines")
    return df


# ─────────────────────────────────────────────────────────────────────────────
# 4. XML  (.xml)
# ─────────────────────────────────────────────────────────────────────────────

def read_xml(
    path: Union[str, Path],
    record_tag: Optional[str] = None,
) -> pd.DataFrame:
    """
    Parses an XML file into a DataFrame.
    Tries pandas.read_xml() first; falls back to ElementTree for nested XML.

    Args:
        path:       Path to the .xml file.
        record_tag: Tag that represents a single record (e.g. "item", "row").
                    Auto-detected if None.
    """
    path = Path(path)

    # ── Attempt 1: pandas native ──────────────────────────────────────────
    try:
        df = pd.read_xml(path)
        df.columns = df.columns.str.strip().str.lower().str.replace(" ", "", regex=False)
        # FIX: apply detect_pii and add Doc ID for the native path too
        df = detect_pii(df)
        df["Doc ID"] = path.name
        print(f"[XML]   {df.shape[0]} rows × {df.shape[1]} cols  |  pandas native")
        return df
    except Exception:
        pass  # fall through to ElementTree

    # ── Attempt 2: ElementTree — flatten children of record_tag ──────────
    tree = ET.parse(path)
    root = tree.getroot()

    if record_tag is None:
        tag_counts: dict[str, int] = {}
        for child in root:
            tag = re.sub(r"\{.*?\}", "", child.tag)  # strip namespace
            tag_counts[tag] = tag_counts.get(tag, 0) + 1
        record_tag = max(tag_counts, key=tag_counts.get)
        print(f"[XML]   Auto-detected record tag: <{record_tag}>")

    records: list[dict] = []
    for elem in root.iter(record_tag):
        row: dict = {**elem.attrib}
        for child in elem:
            tag = re.sub(r"\{.*?\}", "", child.tag)
            row[tag] = child.text
        records.append(row)

    df = pd.DataFrame(records)
    df.columns = df.columns.str.strip().str.lower().str.replace(" ", "", regex=False)
    df = detect_pii(df)
    df["Doc ID"] = path.name
    print(f"[XML]   {df.shape[0]} rows × {df.shape[1]} cols  |  ElementTree fallback")
    return df


# ─────────────────────────────────────────────────────────────────────────────
# 5. TXT  (.txt)
# ─────────────────────────────────────────────────────────────────────────────

def read_txt(path: Union[str, Path]) -> pd.DataFrame:
    """Reads a plain-text file; each non-empty line becomes one row."""
    path = Path(path)

    with open(path, "r", encoding="utf-8") as f:
        lines = [line.strip() for line in f if line.strip()]

    df = pd.DataFrame({"text": lines, "Doc ID": path.name, "Source": "text"})
    print(f"[TXT]   {df.shape[0]} rows × {df.shape[1]} cols  |  raw text lines")
    return df


# ─────────────────────────────────────────────────────────────────────────────
# 6. UNIVERSAL DISPATCHER
# ─────────────────────────────────────────────────────────────────────────────

def read_file(path: Union[str, Path], **kwargs) -> pd.DataFrame:
    """Auto-routes to the correct reader based on file extension."""
    path = Path(path)
    ext = path.suffix.lstrip(".").lower()

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
        raise ValueError(f"Unsupported file extension: .{ext}")
    return reader(path, **kwargs)


def extract(folder_path: Union[str, Path]) -> pd.DataFrame:
    """
    Reads every supported file in a folder and returns a combined DataFrame.
    Unsupported or unreadable files are skipped with a warning.
    """
    folder = Path(folder_path)
    frames: list[pd.DataFrame] = []

    for file_path in sorted(folder.iterdir()):
        if not file_path.is_file():
            continue
        try:
            df = read_file(file_path)
            print(f"Extracted {df.shape[0]} rows from {file_path.name}")
            frames.append(df)
        except ValueError as e:
            print(f"[SKIP]  {file_path.name}: {e}")
        except Exception as e:
            print(f"[ERROR] {file_path.name}: {e}")

    if not frames:
        raise ValueError(f"No supported files found in: {folder}")

    return pd.concat(frames, ignore_index=True)
