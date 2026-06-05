import pandas as pd
import pdfplumber          # pip install pdfplumber
import xml.etree.ElementTree as ET
import re

# ─────────────────────────────────────────────
# 1. EXCEL  (.xlsx / .xlsm / .xls / .ods)
# ─────────────────────────────────────────────

def read_excel(path: str, sheet_name=0) -> pd.DataFrame:
    """
    Reads an Excel file into a DataFrame.
    - .xlsx / .xlsm  → openpyxl (default)
    - .xls (legacy)  → xlrd
    - .ods           → odf
    """
    ext = path.rsplit(".", 1)[-1].lower()

    engine_map = {
        "xls":  "xlrd",
        "ods":  "odf",
        "xlsm": "openpyxl",
        "xlsx": "openpyxl",
    }
    engine = engine_map.get(ext, "openpyxl")

    df = pd.read_excel(path, sheet_name=sheet_name, engine=engine)
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
    return df


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
                lines.extend(text.splitlines())

    return pd.DataFrame({"text": [l for l in lines if l.strip()]})


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
    return df


# ─────────────────────────────────────────────
# 5. UNIVERSAL DISPATCHER
# ─────────────────────────────────────────────

def read_file(path: str, **kwargs) -> pd.DataFrame:
    """Auto-routes to the right reader based on file extension."""
    ext = path.rsplit(".", 1)[-1].lower()
    dispatch = {
        "xlsx": read_excel,
        "xlsm": read_excel,
        "xls":  read_excel,
        "ods":  read_excel,
        "csv":  read_csv,
        "tsv":  lambda p, **kw: read_csv(p, sep="\t", **kw),
        "pdf":  read_pdf,
        "xml":  read_xml,
    }
    reader = dispatch.get(ext)
    if reader is None:
        raise ValueError(f"Unsupported extension: .{ext}")
    return reader(path, **kwargs)


# ─────────────────────────────────────────────
# Usage
# ─────────────────────────────────────────────
if __name__ == "__main__":
    df_excel = read_file("data.xlsx", sheet_name="Sheet1")
    df_csv   = read_file("data.csv")
    df_pdf   = read_file("report.pdf", pages="all")
    df_xml   = read_file("data.xml", record_tag="item")   # or omit record_tag for auto-detect

    print(df_excel.head())
    print(df_csv.head())
    print(df_pdf.head())
    print(df_xml.head())