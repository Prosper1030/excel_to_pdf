import pandas as pd
import glob
from fpdf import FPDF
from pathlib import Path

filepaths = glob.glob("invoices/*.xlsx")

for filepath in filepaths:
    pdf = FPDF(orientation="P", unit="mm", format="A4")
    pdf.add_page()

    filename = Path(filepath).stem
    invoice_nr, date = filename.split("-")

    pdf.set_font(family="Times", size=16, style="B")
    pdf.cell(w=50, h=8, txt=f"Invoice nr.{invoice_nr}", ln=1)

    pdf.set_font(family="Times", size=16, style="B")
    pdf.cell(w=50, h=8, txt=f"Dat: {date}", ln=1)

    df = pd.read_excel(filepath, sheet_name="Sheet 1")

    # Add headers
    columns = list(df.columns)
    column_dict = columns
    columns = [item.replace("_", " ").title() for item in columns]
    pdf.set_font(family="Times", size=10, style="B")
    pdf.set_text_color(80, 80, 80)
    w_data = [30, 60, 40, 30, 30]

    for i in range(5):
        if i != 4:
            pdf.cell(w=w_data[i], h=8, txt=columns[i], border=1, ln=0)
        else:
            pdf.cell(w=w_data[i], h=8, txt=columns[i], border=1, ln=1)

    # Add rows to table
    for index, row in df.iterrows():
        pdf.set_font(family="Times", size=10)
        for i in range(5):
            if i != 4:
                pdf.cell(w=w_data[i], h=8, txt=str(row[column_dict[i]]), border=1, ln=0)
            else:
                pdf.cell(w=w_data[i], h=8, txt=str(row[column_dict[i]]), border=1, ln=1)

    total_sum = df["total_price"].sum()
    pdf.set_font(family="Times", size=10)
    pdf.set_text_color(80, 80, 80)
    for i in range(5):
        if i != 4:
            pdf.cell(w=w_data[i], h=8, txt="", border=1, ln=0)
        else:
            pdf.cell(w=w_data[i], h=8, txt=str(total_sum), border=1, ln=1)
    # Add total sum sentence
    pdf.set_font(family="Times", size=10, style="B")
    pdf.cell(w=30, h=8, txt=f"The total price is {total_sum}", ln=1)
    # Add company name and logo
    pdf.set_font(family="Times", size=14, style="B")
    pdf.cell(w=25, h=8, txt=f"PythonHow")
    pdf.image("pythonhow.png", w=10)

    pdf.output(f"PDFs/{filename}.pdf")