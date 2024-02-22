import pandas as pd
import glob
from fpdf import FPDF
from pathlib import Path

# Create a list of files
filepaths = glob.glob("invoices/*.xlsx")

# Go through the list
for filepath in filepaths:
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
    pdf.cell(w=50, h=8, txt=f"Date {date}", ln=1)

    # Create a data frame
    df = pd.read_excel(filepath, sheet_name="Sheet 1")

    # Create the column list
    columns = df.columns  # we can iterate with this index object
    columns = [item.replace("_", " ").title() for item in columns]

    # Add the column headers to the table
    pdf.set_font(family="Times", style="B", size=10)
    pdf.set_text_color(80, 80, 80)
    pdf.cell(w=30, h=8, txt=columns[0], border=1)
    pdf.cell(w=60, h=8, txt=columns[1], border=1)
    pdf.cell(w=40, h=8, txt=columns[2], border=1)
    pdf.cell(w=30, h=8, txt=columns[3], border=1)
    pdf.cell(w=30, h=8, txt=columns[4], border=1, ln=1)

    # Go thorough the data frame and add the rows to the table
    for index, row in df.iterrows():
        pdf.set_font(family="Times", size=10)
        pdf.set_text_color(80, 80, 80)
        pdf.cell(w=30, h=8, txt=str(row["product_id"]), border=1)
        pdf.cell(w=60, h=8, txt=str(row["product_name"]), border=1)
        pdf.cell(w=40, h=8, txt=str(row["amount_purchased"]), border=1)
        pdf.cell(w=30, h=8, txt=str(row["price_per_unit"]), border=1)
        pdf.cell(w=30, h=8, txt=str(row["total_price"]), border=1, ln=1)

    # Add the total sum
    total_sum = df["total_price"].sum()
    pdf.set_font(family="Times", size=10)
    pdf.set_text_color(80, 80, 80)
    pdf.cell(w=30, h=8, txt="", border=1)
    pdf.cell(w=60, h=8, txt="", border=1)
    pdf.cell(w=40, h=8, txt="", border=1)
    pdf.cell(w=30, h=8, txt="", border=1)
    pdf.cell(w=30, h=8, txt=str(total_sum), border=1, ln=1)

    # Add total sum sentence
    pdf.set_font(family="Times", style="B", size=10)
    pdf.set_text_color(80, 80, 80)
    pdf.cell(w=30, h=8, txt=f"The total price is {total_sum} EUR", ln=1)

    # Add company name and logo
    pdf.set_font(family="Times", style="B", size=14)
    pdf.set_text_color(80, 80, 80)
    pdf.cell(w=30, h=8, txt="PythonHow")
    pdf.image("pythonhow.png", w=10)

    # Produce the PDF output
    pdf.output(f"PDFs/{filename}.pdf")
