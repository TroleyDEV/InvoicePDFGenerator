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

    # Replace "_" for white spaces
    product_id = column_names[0]
    product_id = product_id.replace("_", " ")

    product_name = column_names[1]
    product_name = product_name.replace("_", " ")

    amount_purchased = column_names[2]
    amount_purchased = amount_purchased.replace("_", " ")

    price_per_unit = column_names[3]
    price_per_unit = price_per_unit.replace("_", " ")

    total_price = column_names[4]
    total_price = total_price.replace("_", " ")

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

    for index, row in df.iterrows():
        pdf.set_font(family="Times", size=10)
        pdf.set_text_color(80, 80, 80)

        pdf.cell(w=product_id_w, h=8, txt=f"{row["product_id"]}", border=1)
        pdf.cell(w=product_name_w, h=8, txt=f"{row["product_name"]}", border=1)
        pdf.cell(w=amount_purchased_w, h=8, txt=f"{row["amount_purchased"]}", border=1)
        pdf.cell(w=price_per_unit_w, h=8, txt=f"{row["price_per_unit"]}", border=1)
        pdf.cell(w=0, h=8, txt=f"{row["total_price"]}", border=1, ln=1)

        # Calculate total price
        total = total + row["total_price"]

    # Create row for total price info
    pdf.cell(w=product_id_w, h=8, txt="", border=1)
    pdf.cell(w=product_name_w, h=8, txt="", border=1)
    pdf.cell(w=amount_purchased_w, h=8, txt="", border=1)
    pdf.cell(w=price_per_unit_w, h=8, txt="", border=1)
    pdf.cell(w=0, h=8, txt=f"{total}", border=1, ln=1)

    pdf.output(f"PDFs/{filename}.pdf")
