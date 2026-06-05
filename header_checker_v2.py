import tkinter as tk
from tkinter import filedialog, ttk, messagebox
import pandas as pd
import os

class HeaderCheckerApp:
    def __init__(self, root):
        self.root = root
        self.root.title("Header Consistency Checker")
        self.root.geometry("700x500")

        # Inputs
        tk.Label(root, text="Reference File:").grid(row=0, column=0, padx=5, pady=5)
        self.ref_entry = tk.Entry(root, width=50)
        self.ref_entry.grid(row=0, column=1)
        tk.Button(root, text="Browse", command=self.browse_ref).grid(row=0, column=2)

        tk.Label(root, text="Target Folder/Files:").grid(row=1, column=0, padx=5, pady=5)
        self.target_entry = tk.Entry(root, width=50)
        self.target_entry.grid(row=1, column=1)
        tk.Button(root, text="Browse", command=self.browse_targets).grid(row=1, column=2)

        tk.Label(root, text="Header Row Index (0-based):").grid(row=2, column=0)
        self.header_row = tk.Entry(root, width=10)
        self.header_row.insert(0, "0")
        self.header_row.grid(row=2, column=1, sticky="w")

        tk.Button(root, text="Run Validation", command=self.validate).grid(row=3, column=1, pady=20)
        
        # Output Table
        self.tree = ttk.Treeview(root, columns=("File", "Status", "Remark"), show="headings")
        self.tree.heading("File", text="File Name")
        self.tree.heading("Status", text="Status")
        self.tree.heading("Remark", text="Remark")
        self.tree.grid(row=4, column=0, columnspan=3, padx=10, pady=10)

        tk.Button(root, text="Export Results to CSV", command=self.export).grid(row=5, column=1)

    def browse_ref(self): self.ref_entry.insert(0, filedialog.askopenfilename())
    def browse_targets(self): self.target_entry.insert(0, filedialog.askdirectory())

    def validate(self):
        self.tree.delete(*self.tree.get_children())
        ref_path = self.ref_entry.get()
        target_dir = self.target_entry.get()
        header_idx = int(self.header_row.get())
        
        ref_df = pd.read_excel(ref_path, header=header_idx)
        ref_cols = list(ref_df.columns.astype(str))
        ref_sheets = pd.ExcelFile(ref_path).sheet_names

        for f in os.listdir(target_dir):
            if f.endswith(('.xlsx', '.xls')):
                path = os.path.join(target_dir, f)
                try:
                    target_xl = pd.ExcelFile(path)
                    target_cols = list(pd.read_excel(path, sheet_name=0, header=header_idx).columns.astype(str))
                    
                    errors = []
                    if set(ref_sheets) != set(target_xl.sheet_names):
                        errors.append(f"Sheet name mismatch. Expected: {ref_sheets}")
                    
                    missing = [c for c in ref_cols if c not in target_cols]
                    extra = [c for c in target_cols if c not in ref_cols]
                    
                    if not errors and not missing and not extra:
                        self.tree.insert("", "end", values=(f, "Pass", "All headers and sheet names are correct."))
                    else:
                        msg = []
                        if missing: msg.append(f"Missing: {missing}")
                        if extra: msg.append(f"Extra/Incorrect: {extra}")
                        self.tree.insert("", "end", values=(f, "Fail", "; ".join(msg)))
                except Exception as e:
                    self.tree.insert("", "end", values=(f, "Error", str(e)))

    def export(self):
        data = [self.tree.item(child)["values"] for child in self.tree.get_children()]
        df = pd.DataFrame(data, columns=["File", "Status", "Remark"])
        df.to_csv("validation_results.csv", index=False)
        messagebox.showinfo("Success", "Results saved to validation_results.csv")

if __name__ == "__main__":
    root = tk.Tk()
    app = HeaderCheckerApp(root)
    root.mainloop()
