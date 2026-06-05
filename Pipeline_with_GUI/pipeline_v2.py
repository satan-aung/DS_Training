"""
pipeline_v2.py — Orchestrates the full ETL pipeline.

Usage:
    python pipeline_v2.py <input_folder> [output_path]

Bugs fixed vs previous version:
  - final_columns selection raised KeyError when a PII column was absent;
    now selects only columns that actually exist in the DataFrame.
  - `'Name' not in df_standard` is now explicit: `'Name' not in df_standard.columns`.
  - Output format is now auto-detected from the output file extension.
  - Removed dead commented-out code blocks.
"""

import logging
import os
import sys
from pathlib import Path

import pandas as pd

from Data_ETL.extract_v3 import extract
from Data_ETL.transform_v2 import transform_data

logging.basicConfig(
    level=logging.INFO,
    format="%(asctime)s - %(levelname)s - %(message)s",
)
logger = logging.getLogger(__name__)

# Canonical column order for the final output
_STANDARD_COLUMNS = ["Doc ID", "Source", "Breach ID", "Name", "Emp ID", "DOB", "SSN"]

# Map file extensions to load_to_file format strings
_EXT_TO_FORMAT = {"csv": "csv", "json": "json", "xlsx": "excel", "xls": "excel"}


# ─────────────────────────────────────────────────────────────────────────────
# Helpers
# ─────────────────────────────────────────────────────────────────────────────

def load_to_file(df: pd.DataFrame, output_path: str, format: str = "csv", **kwargs) -> str:
    """Persists a DataFrame to disk in the requested format."""
    output_dir = os.path.dirname(output_path)
    if output_dir and not os.path.exists(output_dir):
        os.makedirs(output_dir)
        logger.info(f"Created directory: {output_dir}")

    if format == "csv":
        df.to_csv(output_path, index=False, **kwargs)
    elif format == "excel":
        df.to_excel(output_path, index=False, **kwargs)
    elif format == "json":
        df.to_json(output_path, orient="records", indent=2, **kwargs)
    else:
        raise ValueError(f"Unsupported output format: {format!r}")

    logger.info(f"Saved {format.upper()} → {output_path}  ({len(df):,} rows)")
    return output_path


def add_unique_id(
    df: pd.DataFrame,
    id_column: str = "Breach ID",
    prefix: str = "DS",
    start: int = 1,
) -> pd.DataFrame:
    """Adds a zero-padded unique ID column (e.g. DS00000001) to the DataFrame."""
    df = df.copy()
    if prefix:
        df[id_column] = [f"{prefix}{i:08d}" for i in range(start, start + len(df))]
    else:
        df[id_column] = range(start, start + len(df))
    return df


# ─────────────────────────────────────────────────────────────────────────────
# Pipeline
# ─────────────────────────────────────────────────────────────────────────────

def run_pipeline(input_folder: str, output_path: str) -> pd.DataFrame:
    """
    Full ETL pipeline:
      1. Extract  — read all supported files in input_folder
      2. Transform — clean PII fields
      3. Enrich   — add unique Breach IDs
      4. Filter   — keep only standard columns; drop fully-empty ones
      5. Load     — write to output_path
    """
    logger.info(f"Starting ETL pipeline: {input_folder} → {output_path}")

    # ── Extract ───────────────────────────────────────────────────────────────
    df_raw = extract(input_folder)
    logger.info(f"Extracted {len(df_raw):,} rows from {input_folder}")

    # ── Transform ─────────────────────────────────────────────────────────────
    df_clean = transform_data(df_raw)
    logger.info(f"Transformed: {len(df_clean):,} rows")

    # ── Add Breach IDs ────────────────────────────────────────────────────────
    df_with_id = add_unique_id(df_clean, id_column="Breach ID", prefix="DS", start=1)

    # ── Select standard columns (only those present) ──────────────────────────
    # FIX: filter to existing columns to prevent KeyError
    present = [c for c in _STANDARD_COLUMNS if c in df_with_id.columns]
    df_standard = df_with_id[present]

    if df_standard.empty:
        raise ValueError("No PII columns detected in the extracted data.")

    # FIX: use explicit .columns membership check
    if "Name" not in df_standard.columns:
        if "SSN" not in df_standard.columns:
            logger.warning("No Name or SSN columns detected")
        else:
            logger.info("SSN present but no Name column found")

    # ── Drop fully-empty columns ──────────────────────────────────────────────
    empty_cols = [
        col for col in df_standard.columns
        if df_standard[col].isna().all()
        or (df_standard[col].astype(str).str.strip() == "").all()
    ]
    if empty_cols:
        df_standard = df_standard.drop(columns=empty_cols)
        logger.info(f"Dropped {len(empty_cols)} empty column(s): {empty_cols}")

    # ── Load ──────────────────────────────────────────────────────────────────
    ext = Path(output_path).suffix.lstrip(".").lower()
    fmt = _EXT_TO_FORMAT.get(ext, "csv")
    load_to_file(df_standard, output_path, format=fmt)

    logger.info(
        f"Pipeline complete — {len(df_standard):,} rows, "
        f"{len(df_standard.columns)} columns"
    )
    return df_standard


# ─────────────────────────────────────────────────────────────────────────────
# CLI entry point
# ─────────────────────────────────────────────────────────────────────────────

if __name__ == "__main__":
    if len(sys.argv) < 2:
        print("Usage: python pipeline_v2.py <input_folder> [output_path]")
        print("Example: python pipeline_v2.py test_data output/cleaned_data.csv")
        sys.exit(1)

    input_folder = sys.argv[1]
    output_path = (
        sys.argv[2] if len(sys.argv) >= 3
        else f"cleaned_{Path(input_folder).stem}.csv"
    )

    if not os.path.exists(input_folder):
        print(f"Error: Folder not found — {input_folder}")
        sys.exit(1)

    result = run_pipeline(input_folder, output_path)

    print("\n" + "=" * 50)
    print("ETL Pipeline Complete!")
    print(f"  Input:   {input_folder}")
    print(f"  Output:  {output_path}")
    print(f"  Rows:    {len(result):,}")
    print(f"  Columns: {list(result.columns)}")
    print("=" * 50)
