import pandas as pd
from fpdf import FPDF
import glob
from pathlib import Path

filepaths = glob.glob("invoices/*.xlsx")

for filepath in filepaths:
    df = pd.read_excel(filepath, sheet_name="Sheet 1")
    pdf = FPDF(orientation="P", unit="mm", format="A4")
    pdf.add_page()

    invoice_number = Path(filepath).stem.split("-")[0]

    pdf.set_font(family="Times", size=16, style="B")
    pdf.cell(w=50, h=8, txt=f"Invoice nr. {invoice_number}")
    pdf.output(f"PDFs/{Path(filepath).stem}.pdf")