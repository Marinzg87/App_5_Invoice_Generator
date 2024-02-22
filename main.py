import pandas as pd
import glob
from fpdf import FPDF
from pathlib import Path

# Create a list of files
filepaths = glob.glob("invoices/*.xlsx")

# Go through the list
for filepath in filepaths:
    # Create a data frame
    df = pd.read_excel(filepath, sheet_name="Sheet 1")
    # Create a PDF instance
    pdf = FPDF(orientation="P", unit="mm", format="A4")
    pdf.add_page()
    # Get the invoice number and date
    filename = Path(filepath).stem
    invoice_nr, date = filename.split("-")
    # Invoice numer
    pdf.set_font(family="Times", size=16, style="B")
    pdf.cell(w=50, h=8, txt=f"Invoice nr. {invoice_nr}", ln=1)
    # Invoice date
    pdf.set_font(family="Times", size=16, style="B")
    pdf.cell(w=50, h=8, txt=f"Date {date}")


    # Produce the PDF output
    pdf.output(f"PDFs/{filename}.pdf")
