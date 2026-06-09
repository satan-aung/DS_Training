"""
gui.py — PDF Data Extractor GUI
================================
A full-featured Tkinter front-end for pdf_extractor.py.

Run directly:
    python gui.py

Features
--------
• Select individual PDF files or an entire folder
• Configure column names and optional per-column regex validation patterns
• Define skip-row patterns (headers, footers, unwanted rows)
• Skip first / last N rows per page
• Table detection strategy selector (auto / lines / text)
• Threaded extraction with live progress bar and colour-coded log
• Export clean data as Excel / CSV / JSON
• Export removed rows always as Excel (with source file, page, reason)
• Save / load configuration as JSON for reuse
• FULLY RESIZABLE: Adjust window size and drag middle divider
"""

from __future__ import annotations

import json
import os
import threading
from datetime import datetime
from pathlib import Path
from typing import List, Optional

import pandas as pd
import tkinter as tk
from tkinter import filedialog, messagebox, scrolledtext, ttk

from pdf_extraction.pdf_extractor import ColumnDef, ExtractionConfig, PDFExtractor


# ══════════════════════════════════════════════════════════════════════════════
#  Constants
# ══════════════════════════════════════════════════════════════════════════════

APP_TITLE   = "PDF Data Extractor"
APP_VERSION = "2.0"  # Updated for resizable layout

PRESET_REGEX_PATTERNS = {
    "PII Complete (5 cols)": (
        r'^(?P<name>.+?)\s{2,}'
        r'(?P<ssn>\d{3}-\d{2}-\d{4})\s{2,}'
        r'(?P<dob>\d{4}-\d{2}-\d{2})\s{2,}'
        r'(?P<phone>\(\d{3}\)\s\d{3}-\d{4})\s{2,}'
        r'(?P<email>\S+@\S+)$'
    ),
    "PII strict": (
        r"^(?P<name>.+?)\s{2,}"
        r"(?P<ssn>\d{3}-\d{2}-\d{4})\s{2,}"
        r"(?P<dob>\d{4}-\d{2}-\d{2})\s{2,}"
        r"(?P<phone>\(\d{3}\)\s\d{3}-\d{4})\s{2,}"
        r"(?P<email>\S+@\S+)$"
    ),
    "Email": r"^\S+@\S+\.\S+$",
    "Phone": r"^\(\d{3}\)\s\d{3}-\d{4}$",
}

# PII Column names for the complete pattern
PII_COLUMN_NAMES = ["name", "ssn", "dob", "phone", "email"]

# Pre-filled skip patterns (regex, one per line)
DEFAULT_SKIP = (
    r"Page \d+ of \d+"  + "\n"
    r"Generated (on|by)"+ "\n"
    r"Confidential"     + "\n"
    r"^\s*Total\s*$"    + "\n"
    r"^\s*Grand Total"
)

EXPORT_FMTS = ["Excel (.xlsx)", "CSV (.csv)", "JSON (.json)"]
EXT_MAP = {
    "Excel (.xlsx)": ".xlsx",
    "CSV (.csv)":    ".csv",
    "JSON (.json)":  ".json",
}
FT_MAP = {
    "Excel (.xlsx)": [("Excel workbook", "*.xlsx")],
    "CSV (.csv)":    [("CSV file",       "*.csv")],
    "JSON (.json)":  [("JSON file",      "*.json")],
}


# ══════════════════════════════════════════════════════════════════════════════
#  Reusable sub-widget
# ══════════════════════════════════════════════════════════════════════════════

class ColumnRow(ttk.Frame):
    """One editable row in the column-configuration table."""

    def __init__(self, parent: tk.Widget, index: int, **kw) -> None:
        super().__init__(parent, **kw)
        self.index = index

        # Use grid for better resizing within the row
        self.columnconfigure(0, weight=0)  # index label
        self.columnconfigure(1, weight=1)  # column name
        self.columnconfigure(2, weight=3)  # regex pattern

        ttk.Label(self, text=str(index), width=4, anchor="center").grid(
            row=0, column=0, padx=(2, 4), sticky="w"
        )
        
        self.name_var    = tk.StringVar(value=f"Column{index}")
        self.pattern_var = tk.StringVar()
        
        name_entry = ttk.Entry(self, textvariable=self.name_var, width=15)
        name_entry.grid(row=0, column=1, padx=2, sticky="ew")
        
        pattern_entry = ttk.Entry(self, textvariable=self.pattern_var)
        pattern_entry.grid(row=0, column=2, padx=2, sticky="ew")

    @property
    def name(self) -> str:
        return self.name_var.get().strip()

    @property
    def pattern(self) -> str:
        return self.pattern_var.get().strip()


# ══════════════════════════════════════════════════════════════════════════════
#  Main Application
# ══════════════════════════════════════════════════════════════════════════════

class App(tk.Tk):

    def __init__(self) -> None:
        super().__init__()
        self.title(f"{APP_TITLE}  v{APP_VERSION}")
        self.geometry("1400x900")  # Slightly wider default
        self.minsize(1024, 768)    # Larger minimum size for usability

        # ── Runtime state ──────────────────────────────────────────────────
        self.selected_paths: List[str]         = []
        self.clean_df:  Optional[pd.DataFrame] = None
        self.removed_df: Optional[pd.DataFrame] = None
        self._col_rows: List[ColumnRow]         = []
        self._busy = False

        self._build_styles()
        self._build_menu()
        self._build_ui()

        # Bind resize event to adjust scrollable regions
        self.bind("<Configure>", self._on_window_resize)

    # ── Styles ───────────────────────────────────────────────────────────────

    def _build_styles(self) -> None:
        s = ttk.Style(self)
        try:
            s.theme_use("clam")
        except tk.TclError:
            pass
        s.configure(
            "Run.TButton",
            font=("Segoe UI", 11, "bold"),
            foreground="white",
            background="#0078D4",
            padding=7,
        )
        s.configure("Action.TButton", font=("Segoe UI", 9), padding=4)
        s.configure("TLabelframe.Label", font=("Segoe UI", 9, "bold"))
        
        # Configure pane sash (divider) for better visibility
        s.configure("TSash", background="#cccccc", gripcount=10)

    # ── Menu bar ─────────────────────────────────────────────────────────────

    def _build_menu(self) -> None:
        bar = tk.Menu(self)
        fm  = tk.Menu(bar, tearoff=0)
        fm.add_command(label="Save Configuration…", command=self._save_config)
        fm.add_command(label="Load Configuration…", command=self._load_config)
        fm.add_separator()
        fm.add_command(label="Exit", command=self.quit)
        bar.add_cascade(label="File", menu=fm)

        hm = tk.Menu(bar, tearoff=0)
        hm.add_command(label="About", command=self._show_about)
        bar.add_cascade(label="Help", menu=hm)
        self.config(menu=bar)

    # ── Top-level layout (FULLY RESIZABLE) ─────────────────────────────────────

    def _build_ui(self) -> None:
        # Main vertical container that fills the window
        main_container = ttk.Frame(self)
        main_container.pack(fill=tk.BOTH, expand=True, padx=8, pady=8)

        # Horizontal PanedWindow for left/right panels (ADJUSTABLE DIVIDER)
        self.main_pane = ttk.PanedWindow(main_container, orient=tk.HORIZONTAL)
        self.main_pane.pack(fill=tk.BOTH, expand=True)

        # Left panel (configuration) - weight 40% initially
        left_frame = ttk.Frame(self.main_pane)
        self.main_pane.add(left_frame, weight=40)

        # Right panel (results) - weight 60% initially
        right_frame = ttk.Frame(self.main_pane)
        self.main_pane.add(right_frame, weight=60)

        # Build both panels
        self._build_left(left_frame)
        self._build_right(right_frame)

    def _on_window_resize(self, event=None) -> None:
        """Handle window resize events to update scrollable regions."""
        if hasattr(self, '_col_canvas'):
            # Update canvas scroll region
            self._col_canvas.configure(
                scrollregion=self._col_canvas.bbox("all")
            )

    # ══════════════════════════════════════════════════════════════════════════
    #  LEFT PANEL — configuration (RESIZABLE)
    # ══════════════════════════════════════════════════════════════════════════

    def _build_left(self, parent: tk.Widget) -> None:
        # Make parent expandable
        parent.columnconfigure(0, weight=1)
        parent.rowconfigure(0, weight=0)  # Source
        parent.rowconfigure(1, weight=1)  # Columns (expandable)
        parent.rowconfigure(2, weight=0)  # Skip patterns
        parent.rowconfigure(3, weight=0)  # Settings
        parent.rowconfigure(4, weight=0)  # Run button

        # ── 1 · Source ───────────────────────────────────────────────────────
        src_frame = ttk.LabelFrame(parent, text="1 · Source")
        src_frame.grid(row=0, column=0, sticky="ew", padx=4, pady=4)
        src_frame.columnconfigure(0, weight=1)

        btn_row = ttk.Frame(src_frame)
        btn_row.grid(row=0, column=0, sticky="ew", padx=6, pady=(6, 2))
        btn_row.columnconfigure(0, weight=0)
        btn_row.columnconfigure(1, weight=0)
        btn_row.columnconfigure(2, weight=1)
        
        ttk.Button(
            btn_row, text="📄  Select PDF File(s)…",
            command=self._pick_files, style="Action.TButton",
        ).grid(row=0, column=0, padx=(0, 6), sticky="w")
        
        ttk.Button(
            btn_row, text="📁  Select Folder…",
            command=self._pick_folder, style="Action.TButton",
        ).grid(row=0, column=1, sticky="w")

        self._src_lbl = ttk.Label(
            src_frame, text="No source selected.", foreground="#888"
        )
        self._src_lbl.grid(row=1, column=0, sticky="w", padx=6, pady=(2, 6))

        # ── 2 · Column configuration (EXPANDABLE) ─────────────────────────────
        col_lf = ttk.LabelFrame(parent, text="2 · Column Configuration")
        col_lf.grid(row=1, column=0, sticky="nsew", padx=4, pady=4)
        col_lf.columnconfigure(0, weight=1)
        col_lf.rowconfigure(3, weight=1)  # Scrollable area expands

        # Number of columns row
        n_row = ttk.Frame(col_lf)
        n_row.grid(row=0, column=0, sticky="ew", padx=6, pady=4)
        n_row.columnconfigure(0, weight=0)
        n_row.columnconfigure(1, weight=1)
        
        ttk.Label(n_row, text="Number of columns:").grid(row=0, column=0, sticky="w")
        self._ncols = tk.IntVar(value=5)  # Default to 5 for PII pattern
        spin = ttk.Spinbox(
            n_row, from_=1, to=30, textvariable=self._ncols,
            width=5, command=self._refresh_cols,
        )
        spin.grid(row=0, column=1, sticky="w", padx=6)
        spin.bind("<Return>",   lambda _: self._refresh_cols())
        spin.bind("<FocusOut>", lambda _: self._refresh_cols())

        # Table header row
        hdr = ttk.Frame(col_lf)
        hdr.grid(row=1, column=0, sticky="ew", padx=6, pady=(0, 1))
        hdr.columnconfigure(0, weight=0)  # #
        hdr.columnconfigure(1, weight=1)  # Column Name
        hdr.columnconfigure(2, weight=3)  # Regex Pattern
        
        ttk.Label(hdr, text="#", width=4, anchor="center").grid(row=0, column=0, sticky="w")
        ttk.Label(hdr, text="Column Name", anchor="w").grid(row=0, column=1, sticky="ew", padx=2)
        ttk.Label(
            hdr, text="Regex Pattern (empty = accept all values)", anchor="w"
        ).grid(row=0, column=2, sticky="ew", padx=2)

        # Preset patterns row
        preset_row = ttk.Frame(col_lf)
        preset_row.grid(row=2, column=0, sticky="ew", padx=6, pady=(2, 4))
        preset_row.columnconfigure(0, weight=0)
        for i, label in enumerate(("PII Complete (5 cols)", "PII strict", "Email", "Phone")):
            ttk.Button(
                preset_row,
                text=label,
                command=lambda lbl=label: self._insert_preset_pattern(lbl),
                style="Action.TButton",
            ).grid(row=0, column=i, padx=(0 if i == 0 else 6, 0), sticky="w")

        ttk.Separator(col_lf, orient="horizontal").grid(row=3, column=0, sticky="ew", padx=6, pady=1)

        # Scrollable ColumnRow list (EXPANDS)
        wrap = ttk.Frame(col_lf)
        wrap.grid(row=4, column=0, sticky="nsew", padx=6, pady=(0, 4))
        wrap.columnconfigure(0, weight=1)
        wrap.rowconfigure(0, weight=1)
        
        self._col_canvas = tk.Canvas(wrap, highlightthickness=0)
        vsb = ttk.Scrollbar(wrap, orient="vertical", command=self._col_canvas.yview)
        self._col_canvas.configure(yscrollcommand=vsb.set)
        
        self._col_canvas.grid(row=0, column=0, sticky="nsew")
        vsb.grid(row=0, column=1, sticky="ns")

        self._col_inner = ttk.Frame(self._col_canvas)
        self._col_win   = self._col_canvas.create_window(
            (0, 0), window=self._col_inner, anchor="nw"
        )
        
        # Configure inner frame columns to expand
        self._col_inner.columnconfigure(0, weight=1)
        
        def _configure_scrollregion(event):
            self._col_canvas.configure(scrollregion=self._col_canvas.bbox("all"))
        
        self._col_inner.bind("<Configure>", _configure_scrollregion)
        self._col_canvas.bind(
            "<Configure>",
            lambda e: self._col_canvas.itemconfig(self._col_win, width=e.width),
        )

        # ── 3 · Skip-row patterns ─────────────────────────────────────────────
        sk_lf = ttk.LabelFrame(parent, text="3 · Skip Row Patterns (regex, one per line)")
        sk_lf.grid(row=2, column=0, sticky="ew", padx=4, pady=4)
        sk_lf.columnconfigure(0, weight=1)
        
        self._skip_text = tk.Text(sk_lf, height=5, font=("Consolas", 9), wrap=tk.WORD)
        self._skip_text.grid(row=0, column=0, sticky="ew", padx=6, pady=4)
        self._skip_text.insert("1.0", DEFAULT_SKIP)

        # ── 4 · Extraction settings ───────────────────────────────────────────
        es_lf = ttk.LabelFrame(parent, text="4 · Extraction Settings")
        es_lf.grid(row=3, column=0, sticky="ew", padx=4, pady=4)
        es_lf.columnconfigure(0, weight=1)
        
        g = ttk.Frame(es_lf)
        g.grid(row=0, column=0, sticky="ew", padx=6, pady=4)
        g.columnconfigure(0, weight=0)
        g.columnconfigure(1, weight=1)

        settings_rows = [
            ("Skip first N rows / page:",       "_skip_h",   0,  50),
            ("Skip last N rows / page:",        "_skip_f",   0,  50),
            ("Min % of columns that must match:", "_min_m",  100, 100),
        ]
        for row_idx, (label, attr, default, max_val) in enumerate(settings_rows):
            ttk.Label(g, text=label, anchor="w").grid(
                row=row_idx, column=0, sticky="w", pady=2
            )
            var = tk.IntVar(value=default)
            setattr(self, attr, var)
            ttk.Spinbox(g, from_=0, to=max_val, textvariable=var, width=8).grid(
                row=row_idx, column=1, padx=8, sticky="w"
            )

        ttk.Label(g, text="Table detection strategy:", anchor="w").grid(
            row=3, column=0, sticky="w", pady=2
        )
        self._strategy = tk.StringVar(value="auto")
        ttk.Combobox(
            g, textvariable=self._strategy,
            values=["auto", "lines", "text"],
            width=10, state="readonly",
        ).grid(row=3, column=1, padx=8, sticky="w")

        # ── Run button ────────────────────────────────────────────────────────
        self._run_btn = ttk.Button(
            parent, text="▶   Extract Data",
            command=self._run_extraction, style="Run.TButton",
        )
        self._run_btn.grid(row=4, column=0, sticky="ew", padx=4, pady=(10, 4))

        # Populate initial column rows (5 columns for PII)
        self._refresh_cols()

    # ══════════════════════════════════════════════════════════════════════════
    #  RIGHT PANEL — progress, log, export (RESIZABLE)
    # ══════════════════════════════════════════════════════════════════════════

    def _build_right(self, parent: tk.Widget) -> None:
        parent.columnconfigure(0, weight=1)
        parent.rowconfigure(1, weight=1)  # Log expands
        parent.rowconfigure(2, weight=0)  # Summary
        parent.rowconfigure(3, weight=0)  # Export

        # ── Progress bar ──────────────────────────────────────────────────────
        pf = ttk.LabelFrame(parent, text="Progress")
        pf.grid(row=0, column=0, sticky="ew", padx=4, pady=4)
        pf.columnconfigure(0, weight=1)
        
        self._prog = tk.DoubleVar()
        ttk.Progressbar(
            pf, variable=self._prog, maximum=100
        ).grid(row=0, column=0, sticky="ew", padx=6, pady=(6, 2))
        
        self._prog_lbl = ttk.Label(pf, text="Ready.", foreground="#555")
        self._prog_lbl.grid(row=1, column=0, sticky="w", padx=6, pady=(0, 6))

        # ── Log / errors (EXPANDS) ──────────────────────────────────────────
        lf = ttk.LabelFrame(parent, text="Log / Errors")
        lf.grid(row=1, column=0, sticky="nsew", padx=4, pady=4)
        lf.columnconfigure(0, weight=1)
        lf.rowconfigure(1, weight=1)

        tb = ttk.Frame(lf)
        tb.grid(row=0, column=0, sticky="ew", padx=4, pady=2)
        tb.columnconfigure(0, weight=1)
        
        ttk.Button(tb, text="Clear Log", command=self._clear_log, width=10).grid(
            row=0, column=4, padx=(0, 4)
        )
        
        # Legend
        legend_frame = ttk.Frame(tb)
        legend_frame.grid(row=0, column=0, sticky="w")
        
        for color, text in [("#111", "INFO"), ("#1e8449", "SUCCESS"), 
                            ("#ca6f1e", "WARN"), ("#c0392b", "ERROR")]:
            ttk.Label(legend_frame, text=f"■ {text}  ", foreground=color, 
                     font=("Segoe UI", 8)).pack(side=tk.LEFT)

        self._log_widget = scrolledtext.ScrolledText(
            lf,
            state="disabled",
            font=("Consolas", 9),
            wrap=tk.WORD,
        )
        self._log_widget.grid(row=1, column=0, sticky="nsew", padx=4, pady=(0, 4))
        self._log_widget.tag_config("INFO",    foreground="#111111")
        self._log_widget.tag_config("ERROR",   foreground="#c0392b")
        self._log_widget.tag_config("SUCCESS", foreground="#1e8449")
        self._log_widget.tag_config("WARN",    foreground="#ca6f1e")

        # ── Results summary ───────────────────────────────────────────────────
        sf = ttk.LabelFrame(parent, text="Results Summary")
        sf.grid(row=2, column=0, sticky="ew", padx=4, pady=4)
        
        self._summary_lbl = ttk.Label(
            sf, text="No extraction run yet.", foreground="#888"
        )
        self._summary_lbl.grid(row=0, column=0, sticky="w", padx=6, pady=6)

        # ── Export (stays at bottom) ──────────────────────────────────────────
        ef = ttk.LabelFrame(parent, text="Export")
        ef.grid(row=3, column=0, sticky="ew", padx=4, pady=4)
        ef.columnconfigure(0, weight=1)

        fmt_row = ttk.Frame(ef)
        fmt_row.grid(row=0, column=0, sticky="w", padx=6, pady=(6, 2))
        
        ttk.Label(fmt_row, text="Clean data format:").pack(side=tk.LEFT, padx=(0, 8))
        self._fmt = tk.StringVar(value="Excel (.xlsx)")
        for fmt in EXPORT_FMTS:
            ttk.Radiobutton(
                fmt_row, text=fmt, variable=self._fmt, value=fmt
            ).pack(side=tk.LEFT, padx=4)

        btn_row = ttk.Frame(ef)
        btn_row.grid(row=1, column=0, sticky="ew", padx=6, pady=(4, 8))
        
        ttk.Button(
            btn_row, text="💾  Export Clean Data",
            command=self._export_clean, style="Action.TButton",
        ).pack(side=tk.LEFT, padx=(0, 6))
        
        ttk.Button(
            btn_row, text="🗑  Export Removed Rows (.xlsx)",
            command=self._export_removed, style="Action.TButton",
        ).pack(side=tk.LEFT, padx=(0, 6))
        
        ttk.Button(
            btn_row, text="📦  Export Both to Folder…",
            command=self._export_all, style="Action.TButton",
        ).pack(side=tk.LEFT)

    # ══════════════════════════════════════════════════════════════════════════
    #  Column row management
    # ══════════════════════════════════════════════════════════════════════════

    def _insert_preset_pattern(self, name: str) -> None:
        """Insert a preset regex pattern into the first column-pattern box."""
        if not self._col_rows:
            self._logmsg("Create at least one column row before inserting a pattern.", "WARN")
            return

        pattern = PRESET_REGEX_PATTERNS.get(name, "")
        if not pattern:
            self._logmsg(f"No preset pattern found for '{name}'.", "WARN")
            return

        # Special handling for PII Complete pattern
        if name == "PII Complete (5 cols)":
            # Set number of columns to 5
            self._ncols.set(5)
            self._refresh_cols()
            
            # Set column names
            for i, col_name in enumerate(PII_COLUMN_NAMES):
                if i < len(self._col_rows):
                    self._col_rows[i].name_var.set(col_name)
            
            # Set pattern in first row (will be applied to all rows through validation)
            self._col_rows[0].pattern_var.set(pattern)
            self._logmsg(
                f"Applied PII Complete pattern with columns: {', '.join(PII_COLUMN_NAMES)}", 
                "SUCCESS"
            )
        else:
            self._col_rows[0].pattern_var.set(pattern)
            self._logmsg(f"Inserted preset pattern into the first regex box: {name}.", "SUCCESS")

    def _refresh_cols(self) -> None:
        try:
            n = int(self._ncols.get())
        except (tk.TclError, ValueError):
            return

        # Snapshot current values before destroying widgets
        saved = [(r.name, r.pattern) for r in self._col_rows]

        for w in self._col_inner.winfo_children():
            w.destroy()
        self._col_rows.clear()

        # Configure inner frame columns for proper expansion
        self._col_inner.columnconfigure(0, weight=1)
        
        for i in range(1, n + 1):
            row = ColumnRow(self._col_inner, index=i)
            row.pack(fill=tk.X, expand=True, pady=1)
            
            # Restore saved values if available
            if i - 1 < len(saved):
                row.name_var.set(saved[i - 1][0] or f"Column{i}")
                row.pattern_var.set(saved[i - 1][1])
            self._col_rows.append(row)
        
        # Update canvas scroll region
        self._col_canvas.configure(scrollregion=self._col_canvas.bbox("all"))

    # ══════════════════════════════════════════════════════════════════════════
    #  Source selection (unchanged but included for completeness)
    # ══════════════════════════════════════════════════════════════════════════

    def _pick_files(self) -> None:
        paths = filedialog.askopenfilenames(
            title="Select PDF Files",
            filetypes=[("PDF files", "*.pdf"), ("All files", "*.*")],
        )
        if not paths:
            return
        self.selected_paths = list(paths)
        names  = ", ".join(os.path.basename(p) for p in paths[:3])
        suffix = " …" if len(paths) > 3 else ""
        self._src_lbl.config(
            text=f"✓  {len(paths)} file(s): {names}{suffix}",
            foreground="#1a7a1a",
        )
        self._logmsg(f"Selected {len(paths)} PDF file(s).")

    def _pick_folder(self) -> None:
        folder = filedialog.askdirectory(title="Select Folder Containing PDFs")
        if not folder:
            return
        pdfs = sorted(str(p) for p in Path(folder).glob("*.pdf"))
        self.selected_paths = pdfs
        if pdfs:
            self._src_lbl.config(
                text=(
                    f"✓  Folder: {os.path.basename(folder)}"
                    f"  ({len(pdfs)} PDF file(s))"
                ),
                foreground="#1a7a1a",
            )
            self._logmsg(
                f"Folder '{folder}'  →  {len(pdfs)} PDF file(s) found."
            )
        else:
            self._src_lbl.config(
                text="⚠  No PDF files found in that folder.",
                foreground="#c0392b",
            )
            self._logmsg("No PDF files found in the selected folder.", "WARN")

    # ══════════════════════════════════════════════════════════════════════════
    #  Extraction (unchanged)
    # ══════════════════════════════════════════════════════════════════════════

    def _build_config(self) -> ExtractionConfig:
        """Read GUI state and return an ExtractionConfig. Raises ValueError on bad input."""
        columns = []
        for row in self._col_rows:
            if not row.name:
                raise ValueError(f"Column {row.index} has no name — please fill it in.")
            columns.append(ColumnDef(name=row.name, pattern=row.pattern))

        raw_skip  = self._skip_text.get("1.0", "end").strip()
        skip_pats = [ln.strip() for ln in raw_skip.splitlines() if ln.strip()]

        return ExtractionConfig(
            columns=columns,
            skip_row_patterns=skip_pats,
            skip_header_rows=max(0, self._skip_h.get()),
            skip_footer_rows=max(0, self._skip_f.get()),
            min_columns_matched=max(0, min(100, self._min_m.get())) / 100.0,
            table_strategy=self._strategy.get(),
        )

    def _run_extraction(self) -> None:
        if self._busy:
            messagebox.showinfo("Busy", "An extraction is already in progress.")
            return
        if not self.selected_paths:
            messagebox.showwarning(
                "No Source", "Please select PDF files or a folder first."
            )
            return
        try:
            config = self._build_config()
        except ValueError as exc:
            messagebox.showerror("Configuration Error", str(exc))
            return

        self._busy = True
        self._run_btn.config(state="disabled")
        self._prog.set(0)
        self._prog_lbl.config(text="Starting…")
        self._logmsg("─" * 60)
        self._logmsg(f"Extraction started on {len(self.selected_paths)} file(s).")

        threading.Thread(
            target=self._worker, args=(config,), daemon=True
        ).start()

    def _worker(self, config: ExtractionConfig) -> None:
        """Runs in a background thread; posts results back via after()."""
        try:
            extractor = PDFExtractor(
                config=config,
                log_callback=lambda msg, lvl="INFO": self.after(
                    0, self._logmsg, msg, lvl
                ),
                progress_callback=lambda pct: self.after(0, self._set_prog, pct),
            )
            clean, removed = extractor.extract_from_paths(self.selected_paths)
            self.after(0, self._on_done, clean, removed)
        except Exception as exc:
            self.after(0, self._on_fail, str(exc))

    def _on_done(self, clean: pd.DataFrame, removed: pd.DataFrame) -> None:
        self.clean_df   = clean
        self.removed_df = removed
        nc, nr = len(clean), len(removed)

        self._prog.set(100)
        self._prog_lbl.config(
            text=f"Done — {nc} rows extracted,  {nr} rows removed."
        )
        self._logmsg(
            f"Extraction complete ─ {nc} clean rows,  {nr} removed rows.",
            "SUCCESS",
        )

        col_names = "  |  ".join(clean.columns.tolist()) if not clean.empty else "—"
        self._summary_lbl.config(
            text=(
                f"Clean rows: {nc}   ·   Removed rows: {nr}\n"
                f"Columns ({len(clean.columns)}):  {col_names}"
            ),
            foreground="#111",
        )
        self._busy = False
        self._run_btn.config(state="normal")

    def _on_fail(self, msg: str) -> None:
        self._prog_lbl.config(text="Extraction failed — see log.")
        self._logmsg(f"Extraction failed: {msg}", "ERROR")
        messagebox.showerror("Extraction Error", msg)
        self._busy = False
        self._run_btn.config(state="normal")

    # ══════════════════════════════════════════════════════════════════════════
    #  Log helpers (unchanged)
    # ══════════════════════════════════════════════════════════════════════════

    def _logmsg(self, msg: str, level: str = "INFO") -> None:
        """Thread-safe log append (must be called from the main thread via after())."""
        ts = datetime.now().strftime("%H:%M:%S")
        self._log_widget.config(state="normal")
        self._log_widget.insert("end", f"[{ts}] {msg}\n", level)
        self._log_widget.see("end")
        self._log_widget.config(state="disabled")

    def _clear_log(self) -> None:
        self._


# ══════════════════════════════════════════════════════════════════════════════
#  Entry point
# ══════════════════════════════════════════════════════════════════════════════

if __name__ == "__main__":
    app = App()
    app.mainloop()