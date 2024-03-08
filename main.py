import glob
from pathlib import Path

import pandas as pd
from fpdf import FPDF

filepaths = glob.glob("invoices/*.xlsx")

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
    column_names = df.columns

    # Replace "_" for white spaces in column names
    formatted_columns = [col.replace("_", " ") for col in column_names]

    # Now you can access the formatted column names directly
    product_id, product_name, amount_purchased, price_per_unit, total_price = formatted_columns

    pdf.set_font(family="Arial", size=14, style="B")

    # Declare width for cells
    product_id_w = pdf.get_string_width(product_id) + 6
    product_name_w = pdf.get_string_width(product_name) + 6
    amount_purchased_w = pdf.get_string_width(amount_purchased) + 6
    price_per_unit_w = pdf.get_string_width(price_per_unit) + 6

    # Create cell for each column title
    pdf.cell(w=product_id_w, h=8, txt=product_id.title(), border=1, align="C")
    pdf.cell(w=product_name_w, h=8, txt=product_name.title(), border=1, align="C")
    pdf.cell(w=amount_purchased_w, h=8, txt=amount_purchased.title(), border=1, align="C")
    pdf.cell(w=price_per_unit_w, h=8, txt=price_per_unit.title(), border=1, align="C")
    pdf.cell(w=0, h=8, txt=total_price.title(), border=1, ln=1, align="C")

    # Reset total price from previous document
    total = 0

    # Loop for creating data tables from Excel data
    for index, row in df.iterrows():
        pdf.set_font(family="Times", size=10)
        pdf.set_text_color(80, 80, 80)

        pdf.cell(w=product_id_w, h=8, txt=str(row["product_id"]), border=1)
        pdf.cell(w=product_name_w, h=8, txt=str(row["product_name"]), border=1)
        pdf.cell(w=amount_purchased_w, h=8, txt=str(row["amount_purchased"]), border=1)
        pdf.cell(w=price_per_unit_w, h=8, txt=str(row["price_per_unit"]), border=1)
        pdf.cell(w=0, h=8, txt=str(row["total_price"]), border=1, ln=1)

        # Calculate total price
        total = total + row["total_price"]

    # Create row for total price info
    pdf.cell(w=product_id_w, h=8, txt="", border=1)
    pdf.cell(w=product_name_w, h=8, txt="", border=1)
    pdf.cell(w=amount_purchased_w, h=8, txt="", border=1)
    pdf.cell(w=price_per_unit_w, h=8, txt="", border=1)
    pdf.cell(w=0, h=8, txt=str(total), border=1, ln=1)

    pdf.output(f"PDFs/{filename}.pdf")
