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
    pdf.set_font(family='Times', style='B', size=12)
    pdf.cell(w=0, h=12, txt=f'Invoice no. {filename}', border=0, ln=1, align='L')
    pdf.cell(w=0,h=12,txt=f'Date: {Path(path).stem.split("-")[1]}', border=0,align='L', ln=1)

    columns = list(df.columns)
    columns = [name.replace('_', " ").title() for name in columns]
    pdf.set_font(family='Times', style='B', size=12)
    pdf.cell(w=25, h=12, txt=(columns[0]), border=1, align='L')
    pdf.cell(w=50, h=12, txt=(columns[1]), border=1, align='L')
    pdf.cell(w=40, h=12, txt=str(columns[2]), border=1, align='L')
    pdf.cell(w=30, h=12, txt=str(columns[3]), border=1, align='L')
    pdf.cell(w=25, h=12, txt=str(columns[4]), border=1, align='L', ln=1)

    for index, row in df.iterrows():
        pdf.set_font(family='Times', style='', size=12)
        pdf.cell(w=25, h=12, txt=str(row['product_id']), border=1, align='L')
        pdf.cell(w=50, h=12, txt=row['product_name'],border=1, align='L')
        pdf.cell(w=40, h=12, txt=str(row['amount_purchased']), border=1, align='L')
        pdf.cell(w=30, h=12, txt=str(row['price_per_unit']), border=1, align='L')
        pdf.cell(w=25, h=12, txt=str(row['total_price']), border=1, align='L', ln=1)

    pdf.set_font(family='Times', style='B', size=12)
    pdf.cell(w=25, h=12, txt='', border=1, align='L')
    pdf.cell(w=50, h=12, txt='', border=1, align='L')
    pdf.cell(w=40, h=12, txt='', border=1, align='L')
    pdf.cell(w=30, h=12, txt='', border=1, align='L')
    pdf.cell(w=25, h=12, txt=str(df['total_price'].sum()), border=1, align='L', ln=1)

    pdf.set_font(family='Times', style='B', size=12)
    pdf.cell(w=0, h=15, txt=f"The total due amount is {df['total_price'].sum()} Euros.", border=0, align='L', ln=1)

    pdf.set_font(family='Times', style='B', size=12)
    pdf.cell(w=22, h=12, txt="PythonHow", border=0, align='L')
    pdf.image('images/image3.png' )

    pdf.output(f"PDF's/{filename}.pdf")


