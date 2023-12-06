import pandas as pd
from fpdf import FPDF
import glob
from pathlib import Path

filepaths = glob.glob("invoices/*.xlsx")

for filepath in filepaths:

    df = pd.read_excel(filepath, sheet_name="Sheet 1")

    pdf = FPDF(orientation="P", unit="mm", format="A4")
    pdf.add_page()

    invoice_number, date = Path(filepath).stem.split("-")

    pdf.set_font(family="Times", size=16, style="B")
    pdf.cell(w=50, h=8, txt=f"Invoice nr. {invoice_number}", ln=1)

    pdf.set_font(family="Times", size=14)
    pdf.cell(w=50, h=8, txt=f"Date {date}", ln=2)

    # adding header of the table is outside the loop
    columns = [item.replace("_", " ").title() for item in df.columns]
    pdf.set_font(family="Times", size=12, style="B")
    pdf.cell(w=30, h=8, txt=columns[0], border=1)
    pdf.cell(w=65, h=8, txt=columns[1], border=1)
    pdf.cell(w=40, h=8, txt=columns[2], border=1)
    pdf.cell(w=35, h=8, txt=columns[3], border=1)
    pdf.cell(w=25, h=8, txt=columns[4], border=1, ln=1)

    # adding rows from the table to the pdf cell by cell for whole row, then do new line
    for index, row in df.iterrows():
        pdf.set_font(family="Times", size=10)
        pdf.set_text_color(80, 80, 80)

        pdf.cell(w=30, h=8, txt=str(row['product_id']), border=1)
        pdf.cell(w=65, h=8, txt=str(row['product_name']), border=1)
        pdf.cell(w=40, h=8, txt=str(row['amount_purchased']), border=1)
        pdf.cell(w=35, h=8, txt=str(row['price_per_unit']), border=1)
        pdf.cell(w=25, h=8, txt=str(row['total_price']), border=1, ln=1)

    # adding total price
    pdf.set_font(family="Times", size=10)
    pdf.set_text_color(80, 80, 80)

    sum = df['total_price'].sum()

    pdf.cell(w=30, h=8, txt="", border=1)
    pdf.cell(w=65, h=8, txt="", border=1)
    pdf.cell(w=40, h=8, txt="", border=1)
    pdf.cell(w=35, h=8, txt="", border=1)
    pdf.cell(w=25, h=8, txt=str(sum), border=1, ln=1)

    # add total sum copy
    pdf.set_font(family="Times", size=13, style='B')
    pdf.cell(w=25, h=8, txt=f"The total price for this invoice is {sum}", ln=1)

    # add company name and logo
    pdf.set_font(family="Times", size=12, style="B")
    pdf.cell(w=25, h=8, txt=f"PythonHow")
    pdf.image("pythonhow.png", w=10)

    pdf.output(f"PDFs/{Path(filepath).stem}.pdf")
