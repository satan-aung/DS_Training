import pandas as pd

def export_csv(df, path):
    df.to_csv(path, index=False)

def export_excel(df, path):
    with pd.ExcelWriter(path, engine="xlsxwriter") as writer:
        df.to_excel(writer, sheet_name="Validation Results", index=False)
