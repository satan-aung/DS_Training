import pandas as pd

def export_excel(df,path):
    with pd.ExcelWriter(path,engine='xlsxwriter') as w:
        summary=pd.DataFrame({
            'Metric':['Total','Passed','Failed'],
            'Value':[len(df),len(df[df.Status=='PASS']),len(df[df.Status=='FAIL'])]
        })
        summary.to_excel(w,sheet_name='Summary',index=False)
        df.to_excel(w,sheet_name='Results',index=False)

def export_csv(df,path):
    df.to_csv(path,index=False)
