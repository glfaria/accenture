import os
import pandas as pd
from collections import defaultdict
from pptx import Presentation
from pptx.enum.shapes import MSO_SHAPE_TYPE
from PyPDF2 import PdfMerger  # OK usar no Linux para unir PDFs

# === CONFIGURAÇÕES ===
excel_path = "Skills.xlsx"
template_path = "CV.pptx"
output_dir = "output"
output_path = "CV_limpo.pptx"
pptx_dir = os.path.join(output_dir, "pptx")
pdf_dir = os.path.join(output_dir, "pdf")
final_pdf = os.path.join(output_dir, "CVs_final.pdf")

os.makedirs(pptx_dir, exist_ok=True)
os.makedirs(pdf_dir, exist_ok=True)

# === LER EXCEL ===
df = pd.read_excel(excel_path)
df.columns = df.columns.str.strip()  # remove espaços extras
grouped = defaultdict(list)
for _, row in df.iterrows():
    grouped[row["Worker Name"]].append(row)
