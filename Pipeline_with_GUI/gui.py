"""
gui.py — Tkinter GUI for the PII ETL Pipeline.

Run:
    python gui.py

Requires all pipeline files in the same directory:
    extract_v3.py  |  transform_v2.py  |  pipeline_v2.py
"""

import logging
import os
import queue
import threading
from pathlib import Path
from typing import Optional

import tkinter as tk
from tkinter import filedialog, messagebox, ttk

from Data_ETL.pipeline_v2 import run_pipeline


# ─────────────────────────────────────────────────────────────────────────────
# Thread-safe logging bridge
# ─────────────────────────────────────────────────────────────────────────────

class _QueueHandler(logging.Handler):
    """Forwards log records into a queue so the GUI can display them safely."""

    def __init__(self, q: "queue.Queue[str]") -> None:
        super().__init__()
        self.queue = q

    def emit(self, record: logging.LogRecord) -> None:
        self.queue.put(self.format(record))


# ─────────────────────────────────────────────────────────────────────────────
# Application
# ─────────────────────────────────────────────────────────────────────────────

class PipelineApp:
    # ── Palette ───────────────────────────────────────────────────────────────
    BG      = "#0d1117"   # main background
    BG2     = "#161b22"   # card / panel background
    BG3     = "#21262d"   # input / treeview background
    BORDER  = "#30363d"
    FG      = "#c9d1d9"
    FG_DIM  = "#8b949e"
    ACCENT  = "#58a6ff"   # blue
    SUCCESS = "#3fb950"   # green
    ERROR   = "#f85149"   # red
    WARN    = "#d29922"   # amber
    LOG_FG  = "#79c0ff"   # log text colour

    FONT_UI   = ("Segoe UI",        10)
    FONT_BOLD = ("Segoe UI",        10, "bold")
    FONT_H    = ("Segoe UI",        20, "bold")
    FONT_MONO = ("Cascadia Code",    9)          # falls back gracefully
    FONT_TINY = ("Segoe UI",         8)

    def __init__(self, root: tk.Tk) -> None:
        self.root = root
        self.log_queue: "queue.Queue[str]" = queue.Queue()
        self.result_df = None

        self._configure_root()
        self._apply_styles()
        self._build_ui()
        self._attach_logger()
        self._poll_log()          # start periodic log drain

    # ── Root setup ────────────────────────────────────────────────────────────

    def _configure_root(self) -> None:
        self.root.title("PII ETL Pipeline")
        self.root.geometry("980x700")
        self.root.minsize(740, 540)
        self.root.configure(bg=self.BG)

    # ── ttk styles ────────────────────────────────────────────────────────────

    def _apply_styles(self) -> None:
        s = ttk.Style(self.root)
        s.theme_use("clam")

        B, B2, B3 = self.BG, self.BG2, self.BG3
        FG, FD, AC = self.FG, self.FG_DIM, self.ACCENT
        BD = self.BORDER

        # Base
        s.configure(".",               background=B,  foreground=FG,  font=self.FONT_UI,
                                       borderwidth=0, focuscolor=AC)
        s.configure("TFrame",          background=B)
        s.configure("Card.TFrame",     background=B2)

        # Labels
        s.configure("TLabel",          background=B,  foreground=FG)
        s.configure("Dim.TLabel",      background=B,  foreground=FD,  font=self.FONT_TINY)
        s.configure("CardDim.TLabel",  background=B2, foreground=FD,  font=self.FONT_TINY)
        s.configure("Head.TLabel",     background=B,  foreground=AC,  font=self.FONT_H)

        # Primary / accent button
        s.configure("Accent.TButton",
                     background=AC,  foreground=B,
                     font=self.FONT_BOLD,
                     padding=(18, 8),
                     relief="flat")
        s.map("Accent.TButton",
              background=[("active", "#79b8ff"), ("disabled", B3)],
              foreground=[("disabled", FD)])

        # Ghost / secondary button
        s.configure("Ghost.TButton",
                     background=B3,  foreground=FG,
                     font=self.FONT_UI,
                     padding=(10, 6),
                     relief="flat")
        s.map("Ghost.TButton",
              background=[("active", BD)],
              foreground=[("active", FG)])

        # Notebook
        s.configure("TNotebook",       background=B,  borderwidth=0,  tabmargins=0)
        s.configure("TNotebook.Tab",   background=B3, foreground=FD,
                                       font=("Segoe UI", 9, "bold"),
                                       padding=(16, 7))
        s.map("TNotebook.Tab",
              background=[("selected", B2)],
              foreground=[("selected", FG)])

        # Treeview
        s.configure("Treeview",
                     background=B2,   fieldbackground=B2, foreground=FG,
                     font=self.FONT_MONO, rowheight=22)
        s.configure("Treeview.Heading",
                     background=B3,   foreground=AC,
                     font=("Segoe UI", 9, "bold"),
                     relief="flat")
        s.map("Treeview",
              background=[("selected", "#1f3358")],
              foreground=[("selected", FG)])

        # Progress bar
        s.configure("TProgressbar",
                     troughcolor=B3,  background=AC,  thickness=3)

        # Scrollbars
        s.configure("TScrollbar",
                     background=B3,   troughcolor=B,  arrowcolor=FD,  borderwidth=0)

    # ── UI Layout ─────────────────────────────────────────────────────────────

    def _build_ui(self) -> None:
        # ── Header ────────────────────────────────────────────────────────────
        hdr = tk.Frame(self.root, bg=self.BG)
        hdr.pack(fill="x", padx=24, pady=(20, 0))

        tk.Label(hdr, text="PII ETL Pipeline",
                 bg=self.BG, fg=self.ACCENT, font=self.FONT_H).pack(side="left")
        tk.Label(hdr, text="  extract · transform · load",
                 bg=self.BG, fg=self.FG_DIM,
                 font=("Segoe UI", 10)).pack(side="left", pady=(9, 0))

        tk.Frame(self.root, bg=self.BORDER, height=1).pack(fill="x", padx=24, pady=(14, 0))

        # ── Config card ───────────────────────────────────────────────────────
        card = tk.Frame(self.root, bg=self.BG2)
        card.pack(fill="x", padx=24, pady=(16, 0))

        inner = tk.Frame(card, bg=self.BG2)
        inner.pack(fill="x", padx=20, pady=16)
        inner.columnconfigure(1, weight=1)

        self._field_row(inner, row=0, label="INPUT FOLDER",
                        attr="input_var",  browse_cmd=self._browse_input)
        self._field_row(inner, row=1, label="OUTPUT FILE",
                        attr="output_var", browse_cmd=self._browse_output,
                        default="output/cleaned_data.csv")

        # ── Action bar ────────────────────────────────────────────────────────
        action = tk.Frame(self.root, bg=self.BG)
        action.pack(fill="x", padx=24, pady=14)

        self.run_btn = ttk.Button(action, text="▶  Run Pipeline",
                                  style="Accent.TButton",
                                  command=self._start_pipeline)
        self.run_btn.pack(side="left")

        ttk.Button(action, text="Clear Log",
                   style="Ghost.TButton",
                   command=self._clear_log).pack(side="left", padx=(8, 0))

        self.status_var = tk.StringVar(value="Ready.")
        tk.Label(action, textvariable=self.status_var,
                 bg=self.BG, fg=self.FG_DIM,
                 font=("Segoe UI", 9)).pack(side="left", padx=14)

        self.progress = ttk.Progressbar(action, mode="indeterminate", length=160)
        self.progress.pack(side="right")

        # ── Notebook: Log | Preview ───────────────────────────────────────────
        self.nb = ttk.Notebook(self.root)
        self.nb.pack(fill="both", expand=True, padx=24, pady=(0, 0))

        self._build_log_tab()
        self._build_preview_tab()

        # ── Status bar ────────────────────────────────────────────────────────
        bar = tk.Frame(self.root, bg=self.BG2, height=28)
        bar.pack(fill="x", side="bottom")
        bar.pack_propagate(False)

        self.stats_var = tk.StringVar(value="  No data loaded.")
        tk.Label(bar, textvariable=self.stats_var,
                 bg=self.BG2, fg=self.FG_DIM,
                 font=("Segoe UI", 9),
                 anchor="w").pack(side="left", padx=14, fill="y")

    def _field_row(
        self,
        parent: tk.Frame,
        row: int,
        label: str,
        attr: str,
        browse_cmd,
        default: str = "",
    ) -> None:
        """Renders a label + entry + Browse button row inside a grid."""
        pad_top = 10 if row > 0 else 0

        tk.Label(parent, text=label,
                 bg=self.BG2, fg=self.FG_DIM,
                 font=self.FONT_TINY).grid(
            row=row * 2, column=0, columnspan=3, sticky="w", pady=(pad_top, 0))

        var = tk.StringVar(value=default)
        setattr(self, attr, var)

        entry = tk.Entry(
            parent,
            textvariable=var,
            bg=self.BG3, fg=self.FG,
            insertbackground=self.FG,
            relief="flat",
            font=self.FONT_MONO,
            highlightthickness=1,
            highlightcolor=self.ACCENT,
            highlightbackground=self.BORDER,
        )
        entry.grid(row=row * 2 + 1, column=0, columnspan=2,
                   sticky="ew", pady=(3, 0), ipady=6)
        parent.columnconfigure(0, weight=1)

        ttk.Button(parent, text="Browse…",
                   style="Ghost.TButton",
                   command=browse_cmd).grid(
            row=row * 2 + 1, column=2, padx=(8, 0), pady=(3, 0), sticky="ns")

    def _build_log_tab(self) -> None:
        frame = tk.Frame(self.nb, bg=self.BG)
        self.nb.add(frame, text=" LOG ")

        self.log_text = tk.Text(
            frame,
            bg=self.BG2, fg=self.LOG_FG,
            font=self.FONT_MONO,
            wrap="word",
            relief="flat",
            borderwidth=0,
            selectbackground="#1f3358",
            selectforeground=self.FG,
            insertbackground=self.ACCENT,
            state="disabled",
        )
        sb = ttk.Scrollbar(frame, command=self.log_text.yview)
        self.log_text.configure(yscrollcommand=sb.set)
        sb.pack(side="right", fill="y")
        self.log_text.pack(fill="both", expand=True, padx=1, pady=1)

        # Colour tags for each log level
        self.log_text.tag_configure("INFO",    foreground=self.LOG_FG)
        self.log_text.tag_configure("WARNING", foreground=self.WARN)
        self.log_text.tag_configure("ERROR",   foreground=self.ERROR)
        self.log_text.tag_configure("CRITICAL",foreground=self.ERROR)
        self.log_text.tag_configure("SUCCESS", foreground=self.SUCCESS)
        self.log_text.tag_configure("SEP",     foreground=self.BORDER)

    def _build_preview_tab(self) -> None:
        frame = tk.Frame(self.nb, bg=self.BG)
        self.nb.add(frame, text=" DATA PREVIEW ")

        hint = tk.Label(
            frame,
            text="Run the pipeline to see a preview of the output (first 500 rows).",
            bg=self.BG, fg=self.FG_DIM, font=("Segoe UI", 9),
        )
        hint.pack(anchor="w", padx=10, pady=(6, 2))
        self._preview_hint = hint

        self.tree = ttk.Treeview(frame, show="headings", selectmode="browse")
        sx = ttk.Scrollbar(frame, orient="horizontal", command=self.tree.xview)
        sy = ttk.Scrollbar(frame, orient="vertical",   command=self.tree.yview)
        self.tree.configure(xscrollcommand=sx.set, yscrollcommand=sy.set)

        sy.pack(side="right",  fill="y")
        sx.pack(side="bottom", fill="x")
        self.tree.pack(fill="both", expand=True)

    # ── Browse dialogs ────────────────────────────────────────────────────────

    def _browse_input(self) -> None:
        path = filedialog.askdirectory(title="Select Input Folder")
        if path:
            self.input_var.set(path)

    def _browse_output(self) -> None:
        path = filedialog.asksaveasfilename(
            title="Save Output As",
            defaultextension=".csv",
            filetypes=[
                ("CSV",   "*.csv"),
                ("Excel", "*.xlsx"),
                ("JSON",  "*.json"),
                ("All",   "*.*"),
            ],
        )
        if path:
            self.output_var.set(path)

    # ── Pipeline execution ────────────────────────────────────────────────────

    def _start_pipeline(self) -> None:
        in_folder = self.input_var.get().strip()
        out_path  = self.output_var.get().strip()

        if not in_folder:
            messagebox.showwarning("Missing Input", "Please select an input folder.")
            return
        if not Path(in_folder).is_dir():
            messagebox.showerror("Not Found", f"Folder does not exist:\n{in_folder}")
            return
        if not out_path:
            messagebox.showwarning("Missing Output", "Please specify an output file path.")
            return

        self.run_btn.configure(state="disabled")
        self.progress.start(10)
        self._set_status("Running pipeline…")
        self._log("─" * 58, tag="SEP")
        self._log(f"▶  Pipeline started")
        self._log(f"   Input:  {in_folder}")
        self._log(f"   Output: {out_path}")
        self._log("─" * 58, tag="SEP")

        threading.Thread(
            target=self._worker,
            args=(in_folder, out_path),
            daemon=True,
        ).start()

    def _worker(self, in_folder: str, out_path: str) -> None:
        try:
            df = run_pipeline(in_folder, out_path)
            self.result_df = df
            self.root.after(0, self._on_success, df, out_path)
        except Exception as exc:
            self.root.after(0, self._on_error, str(exc))

    def _on_success(self, df, out_path: str) -> None:
        self.progress.stop()
        self.run_btn.configure(state="normal")
        fname = Path(out_path).name
        self._set_status(f"✓  {len(df):,} rows saved to {fname}")
        self._log(f"✓  Saved {len(df):,} rows → {out_path}", tag="SUCCESS")
        self._log("─" * 58, tag="SEP")
        self._populate_preview(df)
        self.nb.select(1)   # switch to preview tab
        self.stats_var.set(
            f"  {len(df):,} rows  ·  {len(df.columns)} columns  ·  {out_path}"
        )

    def _on_error(self, msg: str) -> None:
        self.progress.stop()
        self.run_btn.configure(state="normal")
        self._set_status("✗  Pipeline failed — see log for details")
        self._log(f"✗  ERROR: {msg}", tag="ERROR")
        self._log("─" * 58, tag="SEP")
        messagebox.showerror("Pipeline Error", msg)

    # ── Data preview ──────────────────────────────────────────────────────────

    def _populate_preview(self, df) -> None:
        self._preview_hint.configure(
            text=f"Showing first {min(500, len(df)):,} of {len(df):,} rows."
        )
        self.tree.delete(*self.tree.get_children())
        cols = list(df.columns)
        self.tree["columns"] = cols

        for col in cols:
            self.tree.heading(col, text=col, anchor="w")
            width = max(90, min(220, len(col) * 13))
            self.tree.column(col, width=width, minwidth=60, anchor="w")

        for _, row in df.head(500).iterrows():
            self.tree.insert("", "end", values=[str(v) if v is not None else "" for v in row])

    # ── Logging helpers ───────────────────────────────────────────────────────

    def _attach_logger(self) -> None:
        handler = _QueueHandler(self.log_queue)
        handler.setFormatter(
            logging.Formatter(
                "%(asctime)s  %(levelname)-8s  %(message)s",
                datefmt="%H:%M:%S",
            )
        )
        root_logger = logging.getLogger()
        root_logger.addHandler(handler)
        root_logger.setLevel(logging.INFO)

    def _poll_log(self) -> None:
        """Drains the log queue on the main thread every 100 ms."""
        try:
            while True:
                msg = self.log_queue.get_nowait()
                # Colour-code by level keyword in the formatted string
                if "ERROR" in msg or "CRITICAL" in msg:
                    tag = "ERROR"
                elif "WARNING" in msg:
                    tag = "WARNING"
                else:
                    tag = "INFO"
                self._log(msg, tag=tag)
        except queue.Empty:
            pass
        self.root.after(100, self._poll_log)

    def _log(self, msg: str, tag: str = "INFO") -> None:
        self.log_text.configure(state="normal")
        self.log_text.insert("end", msg + "\n", tag)
        self.log_text.see("end")
        self.log_text.configure(state="disabled")

    def _clear_log(self) -> None:
        self.log_text.configure(state="normal")
        self.log_text.delete("1.0", "end")
        self.log_text.configure(state="disabled")

    def _set_status(self, msg: str) -> None:
        self.status_var.set(msg)


# ─────────────────────────────────────────────────────────────────────────────
# Entry point
# ─────────────────────────────────────────────────────────────────────────────

def main() -> None:
    root = tk.Tk()
    PipelineApp(root)
    root.mainloop()


if __name__ == "__main__":
    main()
