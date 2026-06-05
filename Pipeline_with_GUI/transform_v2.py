"""
transform_v2.py — Transform / cleaning layer for the ETL pipeline.

Changes vs previous version:
  - Removed unused `from sqlalchemy import true`.
  - Removed the large commented-out detect_pii block (now lives in extract_v3).
  - Extracted a `CLEANERS` dispatch dict to make clean_data() more concise.
  - Minor type-hint and readability improvements.
"""

import re
import logging
from datetime import datetime
from typing import Optional

import pandas as pd

logger = logging.getLogger(__name__)


# ─────────────────────────────────────────────────────────────────────────────
# Field-level cleaners
# ─────────────────────────────────────────────────────────────────────────────

def clean_name(name) -> Optional[str]:
    if pd.isna(name):
        return None
    name = re.sub(r"\s+", " ", str(name).strip())
    return name.title() or None


def clean_id(emp_id) -> Optional[str]:
    if pd.isna(emp_id):
        return None
    emp_id = str(emp_id).strip()
    if not emp_id:
        return None
    return emp_id if emp_id.isdigit() else emp_id.upper()


def clean_dob(dob) -> Optional[str]:
    if pd.isna(dob):
        return None
    date_obj = pd.to_datetime(dob, errors="coerce", format=None)
    if pd.notna(date_obj) and date_obj.year <= datetime.now().year:
        return date_obj.strftime("%m/%d/%Y")
    return None


def clean_ssn(ssn) -> Optional[str]:
    """Returns a formatted SSN (###-##-####) or None if invalid."""
    if pd.isna(ssn):
        return None
    digits = re.sub(r"\D", "", str(ssn))
    if len(digits) != 9:
        return None
    invalid_prefixes = {"000", "666"}
    if (
        digits[:3] not in invalid_prefixes
        and digits[3:5] != "00"
        and digits[5:] != "0000"
        and digits != "123456789"
    ):
        return f"{digits[:3]}-{digits[3:5]}-{digits[5:9]}"
    return None


# ─────────────────────────────────────────────────────────────────────────────
# Row-level cleaning
# ─────────────────────────────────────────────────────────────────────────────

_CLEANERS = {
    "Name":   clean_name,
    "Emp ID": clean_id,
    "DOB":    clean_dob,
    "SSN":    clean_ssn,
}


def clean_data(df: pd.DataFrame) -> pd.DataFrame:
    """Applies type-specific cleaning to each recognised PII column in-place."""
    for col, fn in _CLEANERS.items():
        if col in df.columns:
            df[col] = df[col].apply(fn)

    logger.info(f"clean_data: {len(df)} rows processed")
    return df


# ─────────────────────────────────────────────────────────────────────────────
# Public entry point
# ─────────────────────────────────────────────────────────────────────────────

def transform_data(df: pd.DataFrame) -> pd.DataFrame:
    """Entry point for the transform stage of the pipeline."""
    try:
        logger.info("Starting transform stage")
        return clean_data(df)
    except Exception as e:
        logger.error(f"Transformation failed: {e}")
        raise
