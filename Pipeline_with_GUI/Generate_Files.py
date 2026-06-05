import pandas as pd
import xml.etree.ElementTree as ET
from Data_ETL.extract_v3 import read_excel

# Sample data
df_1 = pd.DataFrame(read_excel(r"D:/Exercises/Python/Pipeline_Testing/Test_Data/DOC00000001.xlsx"))
df_2 = pd.DataFrame(read_excel(r"D:/Exercises/Python/Pipeline_Testing/Test_Data/DOC00000002.xlsx"))
df_3 = pd.DataFrame(read_excel(r"D:/Exercises/Python/Pipeline_Testing/Test_Data/DOC00000003.xlsx"))

# 1. Save as Text (.txt)
df_1.to_csv("DOC00000007.txt", index=False, sep='\t')

# 2. Save as XML (.xml)
root = ET.Element("Employees")
for _, row in df_2.iterrows():
    emp = ET.SubElement(root, "Employee")
    for col in df_2.columns:
        child = ET.SubElement(emp, col.replace(" ", "_"))
        child.text = str(row[col])
tree = ET.ElementTree(root)
tree.write("DOC00000008.xml", encoding="utf-8", xml_declaration=True)

# 3. Save as PDF (.pdf)
# Requires: pip install fpdf
from fpdf import FPDF
pdf = FPDF()
pdf.add_page()
pdf.set_font("Arial", size=10)
for _, row in df_3.iterrows():
    pdf.cell(200, 10, txt=str(row.to_dict()), ln=True)
pdf.output("DOC00000009.pdf")

print("Files generated: DOC00000007.txt, DOC00000008.xml, DOC00000009.pdf")
