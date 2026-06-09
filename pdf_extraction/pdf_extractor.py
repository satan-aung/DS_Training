"""
pdf_extractor.py — Core PDF extraction engine
=============================================
Extracts structured tabular rows from PDF files, classifying each row
as either clean data or removed noise (headers, footers, pattern mismatches).

Dependencies:
    pip install pdfplumber pandas openpyxl
"""

from __future__ import annotations

import os
import re
from dataclasses import dataclass, field
from typing import Callable, Dict, List, Optional, Tuple

import pdfplumber
import pandas as pd


# ══════════════════════════════════════════════════════════════════════════════
#  Configuration data classes
# ══════════════════════════════════════════════════════════════════════════════

@dataclass
class ColumnDef:
    """
    Definition of one expected output column.

    Parameters
    ----------
    name    : display name used as the DataFrame column header
    pattern : optional regex string; if supplied, a cell must match it to be
              accepted.  Leave empty to accept any non-empty value.
    """
    name: str
    pattern: str = ""

    def __post_init__(self) -> None:
        self._compiled: Optional[re.Pattern] = None
        p = self.pattern.strip()
        if p:
            try:
                self._compiled = re.compile(p, re.IGNORECASE)
            except re.error as exc:
                raise ValueError(
                    f"Invalid regex for column '{self.name}': {exc}"
                ) from exc

    def matches(self, value: str) -> bool:
        """Return True when value satisfies this column's pattern constraint."""
        if self._compiled is None:
            return True                         # no constraint → always accept
        return bool(self._compiled.search(str(value)))


@dataclass
class ExtractionConfig:
    """
    All settings required for one extraction run.

    Parameters
    ----------
    columns             : ordered list of ColumnDef objects
    skip_row_patterns   : full-row regex strings; any matching row is removed
    skip_header_rows    : rows to discard from the top of each extracted block
    skip_footer_rows    : rows to discard from the bottom of each extracted block
    min_columns_matched : fraction 0–1 of column patterns that must pass
                          (1.0 = all columns must match)
    table_strategy      : pdfplumber extraction strategy –
                          "auto"  try line-based then text-based
                          "lines" line-based only
                          "text"  text-based only
    """
    columns: List[ColumnDef]
    skip_row_patterns: List[str] = field(default_factory=list)
    skip_header_rows: int = 0
    skip_footer_rows: int = 0
    min_columns_matched: float = 1.0
    table_strategy: str = "auto"

    def __post_init__(self) -> None:
        self._skip_compiled: List[re.Pattern] = []
        for p in self.skip_row_patterns:
            p = p.strip()
            if p:
                try:
                    self._skip_compiled.append(re.compile(p, re.IGNORECASE))
                except re.error:
                    pass  # bad user pattern → silently ignore

    @property
    def column_names(self) -> List[str]:
        return [c.name for c in self.columns]

    def row_matches_skip(self, text: str) -> Optional[str]:
        """
        Return the matched pattern string if this row should be skipped,
        otherwise None.
        """
        for pat in self._skip_compiled:
            if pat.search(text):
                return pat.pattern
        return None


# ══════════════════════════════════════════════════════════════════════════════
#  Extractor
# ══════════════════════════════════════════════════════════════════════════════

class PDFExtractor:
    """
    Extracts structured rows from a list of PDF files.

    Usage
    -----
    config = ExtractionConfig(
        columns=[
            ColumnDef("Date",   r"\\d{2}/\\d{2}/\\d{4}"),
            ColumnDef("Amount", r"^\\d+(\\.\\d{1,2})?$"),
            ColumnDef("Ref"),
        ]
    )
    extractor = PDFExtractor(config, log_callback=print)
    clean_df, removed_df = extractor.extract_from_paths(["report.pdf"])
    """

    # pdfplumber table-extraction setting candidates (tried in order for "auto")
    _TABLE_CANDIDATES: List[Dict] = [
        {"vertical_strategy": "lines",        "horizontal_strategy": "lines"},
        {"vertical_strategy": "lines_strict",  "horizontal_strategy": "lines_strict"},
        {"vertical_strategy": "text",          "horizontal_strategy": "text"},
        {"vertical_strategy": "explicit",      "horizontal_strategy": "text"},
    ]

    def __init__(
        self,
        config: ExtractionConfig,
        log_callback: Optional[Callable[..., None]] = None,
        progress_callback: Optional[Callable[[float], None]] = None,
    ) -> None:
        self.config    = config
        self._n        = len(config.columns)
        self._log_cb   = log_callback   or (lambda *_: None)
        self._prog_cb  = progress_callback or (lambda _: None)

    # ── Public API ───────────────────────────────────────────────────────────

    def extract_from_paths(
        self, paths: List[str]
    ) -> Tuple[pd.DataFrame, pd.DataFrame]:
        """
        Process all paths and return ``(clean_df, removed_df)``.

        ``clean_df``   — rows that passed all filters, with columns from config
        ``removed_df`` — rows that were discarded, with extra meta columns:
                         source_file · page · row_index · reason
        """
        clean_rows:   List[List[str]] = []
        removed_rows: List[Dict]      = []
        total = len(paths)

        for idx, path in enumerate(paths, 1):
            fname = os.path.basename(path)
            self._log(f"[{idx}/{total}] Processing: {fname}")
            try:
                c, r = self._process_file(path)
                clean_rows.extend(c)
                removed_rows.extend(r)
                self._log(
                    f"         ✓  {len(c)} data rows  |  {len(r)} removed rows"
                )
            except Exception as exc:
                self._log(f"         ✗  Error in '{fname}': {exc}", "ERROR")
                raise
            finally:
                self._prog_cb(idx / total * 100)

        cols      = self.config.column_names
        meta_cols = ["source_file", "page", "row_index", "reason"]

        clean_df = (
            pd.DataFrame(clean_rows, columns=cols)
            if clean_rows
            else pd.DataFrame(columns=cols)
        )
        removed_df = (
            pd.DataFrame(removed_rows)
            if removed_rows
            else pd.DataFrame(columns=meta_cols + cols)
        )
        return clean_df, removed_df

    # ── File / page processing ───────────────────────────────────────────────

    def _process_file(
        self, path: str
    ) -> Tuple[List[List[str]], List[Dict]]:
        clean, removed = [], []
        fname = os.path.basename(path)
        with pdfplumber.open(path) as pdf:
            for pnum, page in enumerate(pdf.pages, 1):
                c, r = self._process_page(page, fname, pnum)
                clean.extend(c)
                removed.extend(r)
        return clean, removed

    def _process_page(
        self,
        page:  "pdfplumber.page.Page",
        fname: str,
        pnum:  int,
    ) -> Tuple[List[List[str]], List[Dict]]:
        clean, removed = [], []
        raw_rows = self._extract_raw_rows(page)
        if not raw_rows:
            return clean, removed

        n_rows = len(raw_rows)
        h = self.config.skip_header_rows
        f = self.config.skip_footer_rows
        valid_range = range(h, n_rows - f if f > 0 else n_rows)

        for ridx, raw in enumerate(raw_rows):
            # ── Header / footer skip ────────────────────────────────────────
            if ridx not in valid_range:
                reason = "header" if ridx < h else "footer"
                removed.append(self._make_removed(raw, fname, pnum, ridx, reason))
                continue

            cells    = self._normalize(raw)
            row_text = " ".join(cells)

            # ── Skip-pattern check ──────────────────────────────────────────
            matched_pat = self.config.row_matches_skip(row_text)
            if matched_pat:
                removed.append(
                    self._make_removed(
                        raw, fname, pnum, ridx,
                        f"skip pattern: {matched_pat}"
                    )
                )
                continue

            # ── Empty row check ─────────────────────────────────────────────
            if not any(v.strip() for v in cells):
                removed.append(
                    self._make_removed(raw, fname, pnum, ridx, "empty row")
                )
                continue

            # ── Column-pattern validation ───────────────────────────────────
            ok, reason = self._validate(cells)
            if ok:
                clean.append(cells)
            else:
                removed.append(self._make_removed(raw, fname, pnum, ridx, reason))

        return clean, removed

    # ── Table extraction strategies ──────────────────────────────────────────

    def _extract_raw_rows(
        self, page: "pdfplumber.page.Page"
    ) -> List[List[Optional[str]]]:
        """
        Try pdfplumber table extraction using the configured strategy;
        fall back to plain-text line splitting if no tables are found.
        """
        strategy = self.config.table_strategy

        if strategy == "lines":
            candidates = [self._TABLE_CANDIDATES[0]]
        elif strategy == "text":
            candidates = [self._TABLE_CANDIDATES[2]]
        else:  # "auto"
            candidates = self._TABLE_CANDIDATES

        for settings in candidates:
            try:
                tables = page.extract_tables(settings)
                if tables:
                    rows: List[List[Optional[str]]] = []
                    for tbl in tables:
                        rows.extend(tbl or [])
                    if rows:
                        return rows
            except Exception:
                continue  # try next strategy

        # ── Plain-text fallback: split by 2+ spaces or tabs ─────────────────
        text = page.extract_text() or ""
        rows = []
        for line in text.splitlines():
            line = line.strip()
            if line:
                rows.append(re.split(r"\s{2,}|\t", line))
        return rows

    # ── Row utilities ────────────────────────────────────────────────────────

    def _normalize(self, raw: List[Optional[str]]) -> List[str]:
        """
        Convert a raw cell list (any length, may contain None) to exactly
        ``n_cols`` clean stripped strings.
        """
        cells = [str(v).strip() if v is not None else "" for v in (raw or [])]
        cells = cells[: self._n]                    # truncate if too wide
        while len(cells) < self._n:                 # pad if too narrow
            cells.append("")
        return cells

    def _validate(self, cells: List[str]) -> Tuple[bool, str]:
        """Return (True, '') or (False, reason_string)."""
        if not self.config.columns:
            return True, ""

        matched = sum(
            1
            for i, col in enumerate(self.config.columns)
            if col.matches(cells[i] if i < len(cells) else "")
        )
        frac = matched / self._n if self._n else 1.0

        if frac < self.config.min_columns_matched:
            return False, f"pattern mismatch ({matched}/{self._n} cols matched)"
        return True, ""

    def _make_removed(
        self,
        raw:    List[Optional[str]],
        fname:  str,
        page:   int,
        ridx:   int,
        reason: str,
    ) -> Dict:
        cells = self._normalize(raw)
        rec: Dict = {
            "source_file": fname,
            "page":        page,
            "row_index":   ridx,
            "reason":      reason,
        }
        for i, col in enumerate(self.config.columns):
            rec[col.name] = cells[i] if i < len(cells) else ""
        return rec

    # ── Internal logging ─────────────────────────────────────────────────────

    def _log(self, msg: str, level: str = "INFO") -> None:
        self._log_cb(msg, level)