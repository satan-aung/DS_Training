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
• Export clean data as  Excel / CSV / JSON
• Export removed rows always as Excel (with source file, page, reason)
• Save / load configuration as JSON for reuse
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
APP_VERSION = "1.0"

PRESET_REGEX_PATTERNS = {
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

        ttk.Label(self, text=str(index), width=3, anchor="center").pack(
            side=tk.LEFT, padx=(2, 4)
        )
        self.name_var    = tk.StringVar(value=f"Column{index}")
        self.pattern_var = tk.StringVar()
        ttk.Entry(self, textvariable=self.name_var,    width=18).pack(
            side=tk.LEFT, padx=2
        )
        ttk.Entry(self, textvariable=self.pattern_var, width=36).pack(
            side=tk.LEFT, padx=2, fill=tk.X, expand=True
        )

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
        self.geometry("1300x900")
        self.minsize(960, 700)

        # ── Runtime state ──────────────────────────────────────────────────
        self.selected_paths: List[str]         = []
        self.clean_df:  Optional[pd.DataFrame] = None
        self.removed_df: Optional[pd.DataFrame] = None
        self._col_rows: List[ColumnRow]         = []
        self._busy = False

        self._build_styles()
        self._build_menu()
        self._build_ui()

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

    # ── Top-level layout ─────────────────────────────────────────────────────

    def _build_ui(self) -> None:
        pane = ttk.PanedWindow(self, orient=tk.HORIZONTAL)
        pane.pack(fill=tk.BOTH, expand=True, padx=8, pady=8)

        left  = ttk.Frame(pane)
        right = ttk.Frame(pane)
        pane.add(left,  weight=2)
        pane.add(right, weight=3)

        self._build_left(left)
        self._build_right(right)

    # ══════════════════════════════════════════════════════════════════════════
    #  LEFT PANEL — configuration
    # ══════════════════════════════════════════════════════════════════════════

    def _build_left(self, parent: tk.Widget) -> None:

        # ── 1 · Source ───────────────────────────────────────────────────────
        src_frame = ttk.LabelFrame(parent, text="1 · Source")
        src_frame.pack(fill=tk.X, padx=4, pady=4)

        btn_row = ttk.Frame(src_frame)
        btn_row.pack(fill=tk.X, padx=6, pady=(6, 2))
        ttk.Button(
            btn_row, text="📄  Select PDF File(s)…",
            command=self._pick_files, style="Action.TButton", width=22,
        ).pack(side=tk.LEFT, padx=(0, 6))
        ttk.Button(
            btn_row, text="📁  Select Folder…",
            command=self._pick_folder, style="Action.TButton", width=16,
        ).pack(side=tk.LEFT)

        self._src_lbl = ttk.Label(
            src_frame, text="No source selected.", foreground="#888", wraplength=440
        )
        self._src_lbl.pack(padx=6, pady=(2, 6), anchor="w")

        # ── 2 · Column configuration ─────────────────────────────────────────
        col_lf = ttk.LabelFrame(parent, text="2 · Column Configuration")
        col_lf.pack(fill=tk.BOTH, expand=True, padx=4, pady=4)

        n_row = ttk.Frame(col_lf)
        n_row.pack(fill=tk.X, padx=6, pady=4)
        ttk.Label(n_row, text="Number of columns:").pack(side=tk.LEFT)
        self._ncols = tk.IntVar(value=3)
        spin = ttk.Spinbox(
            n_row, from_=1, to=30, textvariable=self._ncols,
            width=5, command=self._refresh_cols,
        )
        spin.pack(side=tk.LEFT, padx=6)
        spin.bind("<Return>",   lambda _: self._refresh_cols())
        spin.bind("<FocusOut>", lambda _: self._refresh_cols())

        # Table header row
        hdr = ttk.Frame(col_lf)
        hdr.pack(fill=tk.X, padx=6, pady=(0, 1))
        ttk.Label(hdr, text="#",           width=3,  anchor="center").pack(side=tk.LEFT, padx=(2, 4))
        ttk.Label(hdr, text="Column Name", width=18, anchor="w").pack(side=tk.LEFT, padx=2)
        ttk.Label(
            hdr,
            text="Regex Pattern  (type your rule here; empty = accept all values)",
            anchor="w",
        ).pack(side=tk.LEFT, padx=2)

        preset_row = ttk.Frame(col_lf)
        preset_row.pack(fill=tk.X, padx=6, pady=(2, 4))
        ttk.Label(preset_row, text="Insert preset patterns:", foreground="#555").pack(side=tk.LEFT)
        for label in ("PII strict", "Email", "Phone"):
            ttk.Button(
                preset_row,
                text=label,
                command=lambda lbl=label: self._insert_preset_pattern(lbl),
                style="Action.TButton",
            ).pack(side=tk.LEFT, padx=(6, 0))

        ttk.Separator(col_lf, orient="horizontal").pack(fill=tk.X, padx=6, pady=1)

        # Scrollable ColumnRow list
        wrap = ttk.Frame(col_lf)
        wrap.pack(fill=tk.BOTH, expand=True, padx=6, pady=(0, 4))
        self._col_canvas = tk.Canvas(wrap, highlightthickness=0)
        vsb = ttk.Scrollbar(wrap, orient="vertical", command=self._col_canvas.yview)
        self._col_canvas.configure(yscrollcommand=vsb.set)
        self._col_canvas.pack(side=tk.LEFT, fill=tk.BOTH, expand=True)
        vsb.pack(side=tk.RIGHT, fill=tk.Y)

        self._col_inner = ttk.Frame(self._col_canvas)
        self._col_win   = self._col_canvas.create_window(
            (0, 0), window=self._col_inner, anchor="nw"
        )
        self._col_inner.bind(
            "<Configure>",
            lambda e: self._col_canvas.configure(
                scrollregion=self._col_canvas.bbox("all")
            ),
        )
        self._col_canvas.bind(
            "<Configure>",
            lambda e: self._col_canvas.itemconfig(self._col_win, width=e.width),
        )
        # Mousewheel scrolling
        self._col_canvas.bind_all(
            "<MouseWheel>",
            lambda e: self._col_canvas.yview_scroll(int(-1 * e.delta / 120), "units"),
        )

        # ── 3 · Skip-row patterns ─────────────────────────────────────────────
        sk_lf = ttk.LabelFrame(parent, text="3 · Skip Row Patterns  (regex, one per line)")
        sk_lf.pack(fill=tk.X, padx=4, pady=4)
        self._skip_text = tk.Text(sk_lf, height=5, font=("Consolas", 9), wrap=tk.NONE)
        self._skip_text.pack(fill=tk.X, padx=6, pady=4)
        self._skip_text.insert("1.0", DEFAULT_SKIP)

        # ── 4 · Extraction settings ───────────────────────────────────────────
        es_lf = ttk.LabelFrame(parent, text="4 · Extraction Settings")
        es_lf.pack(fill=tk.X, padx=4, pady=4)
        g = ttk.Frame(es_lf)
        g.pack(fill=tk.X, padx=6, pady=4)

        settings_rows = [
            ("Skip first N rows / page:",       "_skip_h",   0,  50),
            ("Skip last N rows / page:",        "_skip_f",   0,  50),
            ("Min % of columns that must match:", "_min_m",  100, 100),
        ]
        for row_idx, (label, attr, default, max_val) in enumerate(settings_rows):
            ttk.Label(g, text=label, anchor="w", width=35).grid(
                row=row_idx, column=0, sticky="w", pady=2
            )
            var = tk.IntVar(value=default)
            setattr(self, attr, var)
            ttk.Spinbox(g, from_=0, to=max_val, textvariable=var, width=6).grid(
                row=row_idx, column=1, padx=8, sticky="w"
            )

        ttk.Label(g, text="Table detection strategy:", anchor="w", width=35).grid(
            row=3, column=0, sticky="w", pady=2
        )
        self._strategy = tk.StringVar(value="auto")
        ttk.Combobox(
            g, textvariable=self._strategy,
            values=["auto", "lines", "text"],
            width=8, state="readonly",
        ).grid(row=3, column=1, padx=8, sticky="w")

        # ── Run button ────────────────────────────────────────────────────────
        self._run_btn = ttk.Button(
            parent, text="▶   Extract Data",
            command=self._run_extraction, style="Run.TButton",
        )
        self._run_btn.pack(fill=tk.X, padx=4, pady=(10, 4))

        # Populate initial column rows
        self._refresh_cols()

    # ══════════════════════════════════════════════════════════════════════════
    #  RIGHT PANEL — progress, log, export
    # ══════════════════════════════════════════════════════════════════════════

    def _build_right(self, parent: tk.Widget) -> None:

        # ── Progress bar ──────────────────────────────────────────────────────
        pf = ttk.LabelFrame(parent, text="Progress")
        pf.pack(fill=tk.X, padx=4, pady=4)
        self._prog = tk.DoubleVar()
        ttk.Progressbar(
            pf, variable=self._prog, maximum=100, length=500
        ).pack(fill=tk.X, padx=6, pady=(6, 2))
        self._prog_lbl = ttk.Label(pf, text="Ready.", foreground="#555")
        self._prog_lbl.pack(padx=6, pady=(0, 6), anchor="w")

        # ── Log / errors ──────────────────────────────────────────────────────
        lf = ttk.LabelFrame(parent, text="Log / Errors")
        lf.pack(fill=tk.BOTH, expand=True, padx=4, pady=4)

        tb = ttk.Frame(lf)
        tb.pack(fill=tk.X, padx=4, pady=2)
        ttk.Button(tb, text="Clear Log", command=self._clear_log, width=10).pack(
            side=tk.RIGHT
        )
        ttk.Label(
            tb,
            text="  ■ INFO   ",
            foreground="#111",
            font=("Segoe UI", 8),
        ).pack(side=tk.RIGHT)
        ttk.Label(
            tb,
            text="  ■ SUCCESS   ",
            foreground="#1e8449",
            font=("Segoe UI", 8),
        ).pack(side=tk.RIGHT)
        ttk.Label(
            tb,
            text="  ■ WARN   ",
            foreground="#ca6f1e",
            font=("Segoe UI", 8),
        ).pack(side=tk.RIGHT)
        ttk.Label(
            tb,
            text="  ■ ERROR",
            foreground="#c0392b",
            font=("Segoe UI", 8),
        ).pack(side=tk.RIGHT)

        self._log_widget = scrolledtext.ScrolledText(
            lf,
            state="disabled",
            font=("Consolas", 9),
            height=18,
            wrap=tk.WORD,
        )
        self._log_widget.pack(fill=tk.BOTH, expand=True, padx=4, pady=(0, 4))
        self._log_widget.tag_config("INFO",    foreground="#111111")
        self._log_widget.tag_config("ERROR",   foreground="#c0392b")
        self._log_widget.tag_config("SUCCESS", foreground="#1e8449")
        self._log_widget.tag_config("WARN",    foreground="#ca6f1e")

        # ── Results summary ───────────────────────────────────────────────────
        sf = ttk.LabelFrame(parent, text="Results Summary")
        sf.pack(fill=tk.X, padx=4, pady=4)
        self._summary_lbl = ttk.Label(
            sf, text="No extraction run yet.", foreground="#888"
        )
        self._summary_lbl.pack(padx=6, pady=6, anchor="w")

        # ── Export ────────────────────────────────────────────────────────────
        ef = ttk.LabelFrame(parent, text="Export")
        ef.pack(fill=tk.X, padx=4, pady=4)

        fmt_row = ttk.Frame(ef)
        fmt_row.pack(fill=tk.X, padx=6, pady=(6, 2))
        ttk.Label(fmt_row, text="Clean data format:").pack(side=tk.LEFT, padx=(0, 8))
        self._fmt = tk.StringVar(value="Excel (.xlsx)")
        for fmt in EXPORT_FMTS:
            ttk.Radiobutton(
                fmt_row, text=fmt, variable=self._fmt, value=fmt
            ).pack(side=tk.LEFT, padx=4)

        btn_row = ttk.Frame(ef)
        btn_row.pack(fill=tk.X, padx=6, pady=(4, 8))
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

        for i in range(1, n + 1):
            row = ColumnRow(self._col_inner, index=i)
            row.pack(fill=tk.X, pady=1)
            if i - 1 < len(saved):
                row.name_var.set(saved[i - 1][0] or f"Column{i}")
                row.pattern_var.set(saved[i - 1][1])
            self._col_rows.append(row)

    # ══════════════════════════════════════════════════════════════════════════
    #  Source selection
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
    #  Extraction
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
    #  Log helpers
    # ══════════════════════════════════════════════════════════════════════════

    def _logmsg(self, msg: str, level: str = "INFO") -> None:
        """Thread-safe log append (must be called from the main thread via after())."""
        ts = datetime.now().strftime("%H:%M:%S")
        self._log_widget.config(state="normal")
        self._log_widget.insert("end", f"[{ts}] {msg}\n", level)
        self._log_widget.see("end")
        self._log_widget.config(state="disabled")

    def _clear_log(self) -> None:
        self._log_widget.config(state="normal")
        self._log_widget.delete("1.0", "end")
        self._log_widget.config(state="disabled")

    def _set_prog(self, pct: float) -> None:
        self._prog.set(pct)
        self._prog_lbl.config(text=f"Processing…  {pct:.0f}%")

    # ══════════════════════════════════════════════════════════════════════════
    #  Export
    # ══════════════════════════════════════════════════════════════════════════

    def _check_data(self) -> bool:
        if self.clean_df is None:
            messagebox.showwarning("No Data", "Run an extraction first.")
            return False
        return True

    def _export_clean(self) -> None:
        if not self._check_data():
            return
        fmt  = self._fmt.get()
        ext  = EXT_MAP[fmt]
        path = filedialog.asksaveasfilename(
            defaultextension=ext,
            filetypes=FT_MAP[fmt],
            initialfile=f"clean_data{ext}",
            title="Save Clean Data",
        )
        if path:
            self._write_df(self.clean_df, path, fmt)
            self._logmsg(f"Clean data saved  →  {path}", "SUCCESS")

    def _export_removed(self) -> None:
        if not self._check_data():
            return
        if self.removed_df is None or self.removed_df.empty:
            messagebox.showinfo("Nothing to Export", "No removed rows were recorded.")
            return
        path = filedialog.asksaveasfilename(
            defaultextension=".xlsx",
            filetypes=[("Excel workbook", "*.xlsx")],
            initialfile="removed_data.xlsx",
            title="Save Removed Rows",
        )
        if path:
            self.removed_df.to_excel(path, index=False)
            self._logmsg(f"Removed rows saved  →  {path}", "SUCCESS")

    def _export_all(self) -> None:
        if not self._check_data():
            return
        folder = filedialog.askdirectory(title="Select Output Folder")
        if not folder:
            return

        ts  = datetime.now().strftime("%Y%m%d_%H%M%S")
        fmt = self._fmt.get()
        ext = EXT_MAP[fmt]

        clean_path   = os.path.join(folder, f"clean_data_{ts}{ext}")
        removed_path = os.path.join(folder, f"removed_data_{ts}.xlsx")

        self._write_df(self.clean_df, clean_path, fmt)
        self._logmsg(f"Clean data saved    →  {clean_path}", "SUCCESS")

        if self.removed_df is not None and not self.removed_df.empty:
            self.removed_df.to_excel(removed_path, index=False)
            self._logmsg(f"Removed data saved  →  {removed_path}", "SUCCESS")
        else:
            self._logmsg("No removed rows to export.", "WARN")

        messagebox.showinfo("Export Complete", f"Files saved to:\n{folder}")

    @staticmethod
    def _write_df(df: pd.DataFrame, path: str, fmt: str) -> None:
        if "Excel" in fmt:
            df.to_excel(path, index=False)
        elif "CSV" in fmt:
            df.to_csv(path, index=False)
        elif "JSON" in fmt:
            df.to_json(path, orient="records", indent=2)

    # ══════════════════════════════════════════════════════════════════════════
    #  Config persistence
    # ══════════════════════════════════════════════════════════════════════════

    def _save_config(self) -> None:
        path = filedialog.asksaveasfilename(
            defaultextension=".json",
            filetypes=[("JSON config", "*.json")],
            initialfile="extractor_config.json",
            title="Save Configuration",
        )
        if not path:
            return
        cfg = {
            "n_cols":           self._ncols.get(),
            "columns":          [{"name": r.name, "pattern": r.pattern} for r in self._col_rows],
            "skip_patterns":    self._skip_text.get("1.0", "end").strip(),
            "skip_header_rows": self._skip_h.get(),
            "skip_footer_rows": self._skip_f.get(),
            "min_match_pct":    self._min_m.get(),
            "table_strategy":   self._strategy.get(),
        }
        with open(path, "w", encoding="utf-8") as fh:
            json.dump(cfg, fh, indent=2)
        self._logmsg(f"Config saved  →  {path}", "SUCCESS")

    def _load_config(self) -> None:
        path = filedialog.askopenfilename(
            title="Load Configuration",
            filetypes=[("JSON config", "*.json")],
        )
        if not path:
            return
        try:
            with open(path, "r", encoding="utf-8") as fh:
                cfg = json.load(fh)
        except Exception as exc:
            messagebox.showerror("Load Error", f"Could not read config:\n{exc}")
            return

        self._ncols.set(cfg.get("n_cols", 3))
        self._refresh_cols()

        for i, col in enumerate(cfg.get("columns", [])):
            if i < len(self._col_rows):
                self._col_rows[i].name_var.set(col.get("name", f"Column{i + 1}"))
                self._col_rows[i].pattern_var.set(col.get("pattern", ""))

        self._skip_text.delete("1.0", "end")
        self._skip_text.insert("1.0", cfg.get("skip_patterns", ""))
        self._skip_h.set(cfg.get("skip_header_rows", 0))
        self._skip_f.set(cfg.get("skip_footer_rows", 0))
        self._min_m.set(cfg.get("min_match_pct", 100))
        self._strategy.set(cfg.get("table_strategy", "auto"))

        self._logmsg(f"Config loaded  ←  {path}", "SUCCESS")

    # ── About ─────────────────────────────────────────────────────────────────

    def _show_about(self) -> None:
        messagebox.showinfo(
            "About",
            f"{APP_TITLE}  v{APP_VERSION}\n\n"
            "Extract structured tabular data from PDF files.\n\n"
            "Files:\n"
            "  gui.py          — this GUI\n"
            "  pdf_extractor.py — extraction engine\n\n"
            "Dependencies:\n"
            "  pdfplumber, pandas, openpyxl\n\n"
            "Tip: use File → Save Configuration to reuse\n"
            "your column and pattern settings.",
        )


# ══════════════════════════════════════════════════════════════════════════════
#  Entry point
# ══════════════════════════════════════════════════════════════════════════════

if __name__ == "__main__":
    app = App()
    app.mainloop()