import glob
import pandas as pd
from fpdf import FPDF
from pathlib import Path

filepaths = glob.glob(f'Invocies\*.xlsx')

for path in filepaths:
    df = pd.read_excel(path)
    pdf = FPDF(orientation='P', unit='mm', format='A4')
    pdf.add_page()
    filename = Path(path).stem.split("-")[0]
    pdf.set_font(family='Times', style='' , size=12)
    pdf.cell(w=0, h=12, txt=f'Invoice no. {filename}', border=0, ln=1, align='L')
    pdf.output(f"PDF's/{filename}.pdf")


