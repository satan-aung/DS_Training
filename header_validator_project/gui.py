import tkinter as tk
from tkinter import ttk, filedialog, messagebox
import pandas as pd
from validator import validate_file
from exporter import export_csv, export_excel

class HeaderValidatorApp(tk.Tk):
    def __init__(self):
        super().__init__()
        self.title("Header Validator")
        self.geometry("1000x700")

        self.ref_file = ""
        self.target_files = []
        self.results = []

        ttk.Button(self, text="Select Reference File", command=self.select_ref).pack(pady=5)
        ttk.Button(self, text="Select Files", command=self.select_files).pack(pady=5)

        frm = ttk.Frame(self)
        frm.pack()

        ttk.Label(frm, text="Sheet Name").grid(row=0,column=0)
        ttk.Label(frm, text="Header Row").grid(row=0,column=1)

        self.sheet_entry = ttk.Entry(frm, width=30)
        self.sheet_entry.insert(0,"Business,Individual")
        self.sheet_entry.grid(row=1,column=0)

        self.row_entry = ttk.Entry(frm, width=10)
        self.row_entry.insert(0,"1,1")
        self.row_entry.grid(row=1,column=1)

        ttk.Button(self, text="Validate", command=self.validate).pack(pady=10)

        cols=("File","Status","Sheet","Remark")
        self.tree=ttk.Treeview(self, columns=cols, show="headings")
        for c in cols:
            self.tree.heading(c,text=c)
            self.tree.column(c,width=200)
        self.tree.pack(fill="both", expand=True)

        ttk.Button(self, text="Export CSV", command=self.save_csv).pack(side="left", padx=10, pady=10)
        ttk.Button(self, text="Export Excel", command=self.save_excel).pack(side="left", padx=10, pady=10)

    def select_ref(self):
        self.ref_file = filedialog.askopenfilename(filetypes=[("Excel","*.xlsx *.xlsm *.xls")])

    def select_files(self):
        self.target_files = filedialog.askopenfilenames(filetypes=[("Excel","*.xlsx *.xlsm *.xls")])

    def validate(self):
        if not self.ref_file:
            messagebox.showerror("Error","Select reference file")
            return

        sheet_names=[s.strip() for s in self.sheet_entry.get().split(",")]
        rows=[int(x.strip()) for x in self.row_entry.get().split(",")]
        cfg=dict(zip(sheet_names, rows))

        self.results.clear()
        for item in self.tree.get_children():
            self.tree.delete(item)

        for f in self.target_files:
            res=validate_file(self.ref_file, f, cfg)
            for status,sheet,remark in res:
                row=[f.split("/")[-1],status,sheet,remark]
                self.results.append(row)
                self.tree.insert("", "end", values=row)

    def save_csv(self):
        if not self.results:
            return
        p=filedialog.asksaveasfilename(defaultextension=".csv")
        if p:
            export_csv(pd.DataFrame(self.results, columns=["File","Status","Sheet","Remark"]), p)

    def save_excel(self):
        if not self.results:
            return
        p=filedialog.asksaveasfilename(defaultextension=".xlsx")
        if p:
            export_excel(pd.DataFrame(self.results, columns=["File","Status","Sheet","Remark"]), p)
