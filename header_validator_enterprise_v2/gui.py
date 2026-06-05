import tkinter as tk
from tkinter import ttk,filedialog,messagebox
from concurrent.futures import ThreadPoolExecutor
import os,pandas as pd
from validator import validate_file
from exporter import export_excel,export_csv
from profile_manager import save_profile,load_profile

class HeaderValidatorApp(tk.Tk):
    def __init__(self):
        super().__init__()
        self.title('Header Validator Enterprise V2')
        self.geometry('1200x800')

        self.ref=''
        self.files=[]
        self.results=[]

        top=ttk.Frame(self); top.pack(fill='x')

        ttk.Button(top,text='Reference File',command=self.pick_ref).pack(side='left')
        ttk.Button(top,text='Select Files',command=self.pick_files).pack(side='left')
        ttk.Button(top,text='Select Folder',command=self.pick_folder).pack(side='left')
        ttk.Button(top,text='Save Profile',command=self.save_prof).pack(side='left')
        ttk.Button(top,text='Load Profile',command=self.load_prof).pack(side='left')

        cfg=ttk.Frame(self); cfg.pack(fill='x')
        ttk.Label(cfg,text='Sheets').grid(row=0,column=0)
        ttk.Label(cfg,text='Header Rows').grid(row=0,column=1)
        self.sheets=ttk.Entry(cfg,width=50)
        self.rows=ttk.Entry(cfg,width=20)
        self.sheets.insert(0,'Business,Individual')
        self.rows.insert(0,'1,1')
        self.sheets.grid(row=1,column=0)
        self.rows.grid(row=1,column=1)

        self.pb=ttk.Progressbar(self,mode='determinate')
        self.pb.pack(fill='x',padx=5,pady=5)

        ttk.Button(self,text='Validate',command=self.run_validation).pack()

        nb=ttk.Notebook(self); nb.pack(fill='both',expand=True)
        self.all_tree=self.make_tree(nb,'All')
        self.pass_tree=self.make_tree(nb,'PASS')
        self.fail_tree=self.make_tree(nb,'FAIL')

        bottom=ttk.Frame(self); bottom.pack(fill='x')
        ttk.Button(bottom,text='Export Excel',command=self.exp_xlsx).pack(side='left')
        ttk.Button(bottom,text='Export CSV',command=self.exp_csv).pack(side='left')

    def make_tree(self,nb,title):
        f=ttk.Frame(nb)
        nb.add(f,text=title)
        tree=ttk.Treeview(f,columns=('File','Status','Sheet','Remark'),show='headings')
        for c in ('File','Status','Sheet','Remark'):
            tree.heading(c,text=c)
            tree.column(c,width=250)
        tree.pack(fill='both',expand=True)
        return tree

    def pick_ref(self):
        self.ref=filedialog.askopenfilename()

    def pick_files(self):
        self.files=list(filedialog.askopenfilenames())

    def pick_folder(self):
        folder=filedialog.askdirectory()
        if folder:
            self.files=[]
            for r,_,fs in os.walk(folder):
                for f in fs:
                    if f.lower().endswith(('.xlsx','.xlsm','.xls')):
                        self.files.append(os.path.join(r,f))

    def cfg(self):
        return dict(zip(
            [x.strip() for x in self.sheets.get().split(',')],
            [int(x.strip()) for x in self.rows.get().split(',')]
        ))

    def run_validation(self):
        if not self.ref:
            messagebox.showerror('Error','Select reference file')
            return

        self.results=[]
        total=max(len(self.files),1)
        self.pb['value']=0

        def work(f):
            return f, validate_file(self.ref,f,self.cfg())

        with ThreadPoolExecutor(max_workers=min(32,(os.cpu_count() or 4)*4)) as ex:
            for i,(fname,res) in enumerate(ex.map(work,self.files),1):
                for st,sh,rm in res:
                    row=[os.path.basename(fname),st,sh,rm]
                    self.results.append(row)
                self.pb['value']=i/total*100
                self.update_idletasks()

        self.refresh()

    def refresh(self):
        for t in (self.all_tree,self.pass_tree,self.fail_tree):
            for i in t.get_children():
                t.delete(i)

        for r in self.results:
            self.all_tree.insert('', 'end', values=r)
            if r[1]=='PASS':
                self.pass_tree.insert('', 'end', values=r)
            else:
                self.fail_tree.insert('', 'end', values=r)

    def exp_xlsx(self):
        if not self.results: return
        p=filedialog.asksaveasfilename(defaultextension='.xlsx')
        if p:
            export_excel(pd.DataFrame(self.results,columns=['File','Status','Sheet','Remark']),p)

    def exp_csv(self):
        if not self.results: return
        p=filedialog.asksaveasfilename(defaultextension='.csv')
        if p:
            export_csv(pd.DataFrame(self.results,columns=['File','Status','Sheet','Remark']),p)

    def save_prof(self):
        p=filedialog.asksaveasfilename(defaultextension='.json')
        if p:
            save_profile(p,{'sheets':self.sheets.get(),'rows':self.rows.get()})

    def load_prof(self):
        p=filedialog.askopenfilename(filetypes=[('JSON','*.json')])
        if p:
            d=load_profile(p)
            self.sheets.delete(0,'end'); self.sheets.insert(0,d['sheets'])
            self.rows.delete(0,'end'); self.rows.insert(0,d['rows'])
