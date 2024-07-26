import pandas as pd
import glob
from fpdf import FPDF
from pathlib import Path

filepaths = glob.glob("invoices/*.xlsx")


for i in filepaths:
    pdf = FPDF(orientation="P", unit="mm", format="A4")
    pdf.add_page()

    filename = Path(i).stem
    invoice_nr, invoice_date = filename[:5], filename[6:]

    pdf.set_font(family="Times", size=16, style="B")
    pdf.cell(w=50, h=8, txt=f"Invoice nr: {invoice_nr}", ln=1)

    pdf.set_font(family="Times", size=16, style="B")
    pdf.cell(w=50, h=8, txt=f"Date: {invoice_date}", ln=1)

    df = pd.read_excel(i, sheet_name="Sheet 1")

    columns_table = list(df.columns)
    columns_table = [item.replace("_", " ").title() for item in columns_table]
    pdf.set_font(family="Times", style="B", size=10)

    pdf.cell(w=30, h=8, txt=columns_table[0], border=1)
    pdf.cell(w=70, h=8, txt=columns_table[1], border=1)
    pdf.cell(w=32, h=8, txt=columns_table[2], border=1)
    pdf.cell(w=30, h=8, txt=columns_table[3], border=1)
    pdf.cell(w=30, h=8, txt=columns_table[4], border=1, ln=1)
    total_sum = df['total_price'].sum()

    for index, row in df.iterrows():
        total_sum = df['total_price'].sum()
        pdf.set_font(family="Times", size=10)
        pdf.set_text_color(80, 80, 80)
        pdf.cell(w=30, h=8, txt=str(row["product_id"]), border=1)
        pdf.cell(w=70, h=8, txt=str(row["product_name"]), border=1)
        pdf.cell(w=32, h=8, txt=str(row["amount_purchased"]), border=1)
        pdf.cell(w=30, h=8, txt=str(row["price_per_unit"]), border=1)
        pdf.cell(w=30, h=8, txt=str(row["total_price"]), border=1, ln=1)

    pdf.set_font(family="Times", size=10)
    pdf.set_text_color(80, 80, 80)
    pdf.cell(w=30, h=8, txt="", border=1)
    pdf.cell(w=70, h=8, txt="", border=1)
    pdf.cell(w=32, h=8, txt="", border=1)
    pdf.cell(w=30, h=8, txt="", border=1)
    pdf.cell(w=30, h=8, txt=str(total_sum), border=1, ln=1)


    pdf.set_font(family="Times",style="B", size=14)
    pdf.cell(w=30, h=8, txt=f"The total price is {total_sum}.", ln=1)

    pdf.set_font(family="Times", style="B", size=14)
    pdf.cell(w=30, h=8, txt=f"PythonHow")
    pdf.image("pythonhow.png")


    pdf.output(f"PDFs/{filename}.pdf")


