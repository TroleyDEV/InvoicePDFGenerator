import glob
from pathlib import Path

import pandas as pd
from fpdf import FPDF

filepaths = glob.glob("invoices/*.xlsx")

# Loop for files
for filepath in filepaths:

    pdf = FPDF(orientation="P", unit="mm", format="A4")
    pdf.set_auto_page_break(auto=True)
    pdf.set_margins(10, 10, 10)
    pdf.add_page()

    filename = Path(filepath).stem
    invoice_nr, date = filename.split('-')

    pdf.set_font(family="Arial", size=16, style="B")
    pdf.cell(w=50, h=8, txt=f"Invoice nr.{invoice_nr}", ln=1)

    pdf.set_font(family="Arial", size=16, style="B")
    pdf.cell(w=50, h=8, txt=f"Date: {date}", ln=1)
    pdf.ln()

    df = pd.read_excel(filepath, sheet_name="Sheet 1")

    # Get column names from Data Frames
    columns = df.columns

    # Replace "_" for white spaces in column names
    columns = [col.replace("_", " ") for col in columns]

    pdf.set_font(family="Arial", size=14, style="B")

    # Declare width for cells
    columns_w = [pdf.get_string_width(col) + 6 for col in columns]

    # Create cell for each column title
    pdf.cell(w=columns_w[0], h=8, txt=columns[0].title(), border=1, align="C")
    pdf.cell(w=columns_w[1], h=8, txt=columns[1].title(), border=1, align="C")
    pdf.cell(w=columns_w[2], h=8, txt=columns[2].title(), border=1, align="C")
    pdf.cell(w=columns_w[3], h=8, txt=columns[3].title(), border=1, align="C")
    pdf.cell(w=0, h=8, txt=columns[4].title(), border=1, ln=1, align="C")

    # Loop for creating data tables from Excel data
    for index, row in df.iterrows():
        pdf.set_font(family="Times", size=9, style="B")
        pdf.set_text_color(80, 80, 80)

        pdf.cell(w=columns_w[0], h=8, txt=str(row["product_id"]), border=1)
        pdf.cell(w=columns_w[1], h=8, txt=str(row["product_name"]), border=1)
        pdf.cell(w=columns_w[2], h=8, txt=str(row["amount_purchased"]), border=1)
        pdf.cell(w=columns_w[3], h=8, txt=str(row["price_per_unit"]), border=1)
        pdf.cell(w=0, h=8, txt=str(row["total_price"]), border=1, ln=1)

    # Create row for total price info
    pdf.cell(w=columns_w[0], h=8, txt="", border=1)
    pdf.cell(w=columns_w[1], h=8, txt="", border=1)
    pdf.cell(w=columns_w[2], h=8, txt="", border=1)
    pdf.cell(w=columns_w[3], h=8, txt="", border=1)
    pdf.cell(w=0, h=8, txt=str(df["total_price"].sum()), border=1, ln=1)

    pdf.ln(20)

    # Set information of total amount and Company name, logo
    pdf.set_font(family="Arial", size=16, style="B")
    pdf.set_text_color(0, 0, 0)
    pdf.cell(w=0, h=8, txt=f"The total due amount is {df["total_price"].sum()} Euros.", ln=1)

    pdf.cell(w=32, h=8, txt="PythonHow")
    pdf.image(name="pythonhow.png", w=8, h=8)

    pdf.output(f"PDFs/{filename}.pdf")
