import pandas as pd
import glob
from fpdf import FPDF
from pathlib import Path

filepaths = glob.glob("invoices/*.xlsx")

for filepath in filepaths:
    # Create PDF file
    pdf = FPDF(orientation="P", unit="mm", format="A4")
    
    # Add page to that PDF file
    pdf.add_page()

    # Retrieve the name of the file and place in variable
    filename = Path(filepath).stem

    # Establish two variables from the filename variable (that has been split by '-') for invoice number and invoice date
    invoice_nr, invoice_date = filename.split("-")
    
    # Set font family, font size and font style 
    pdf.set_font(family="Times", size=16, style="B")


    pdf.cell(w=50, h=8, ln=1, txt=f"Invoice nr. {invoice_nr}")
    pdf.cell(w=50, h=8,  ln=1, txt=f"{invoice_date}")

    # read the data frame
    df = pd.read_excel(filepath, sheet_name="Sheet 1")

    # Add a header
    columns = df.columns
    columns = [item.replace("_", " ").title() for item in columns]
    pdf.set_font(family="Times", size=10, style="B")
    pdf.set_text_color(80,80,80)
    pdf.cell(w=30, h=8, txt=str(columns[0]), border=1)
    pdf.cell(w=70, h=8, txt=str(columns[1]), border=1)
    pdf.cell(w=35, h=8, txt=str(columns[2]), border=1)
    pdf.cell(w=30, h=8, txt=str(columns[3]), border=1)
    pdf.cell(w=30, h=8, txt=str(columns[4]), border=1, ln=1)

    # Add rows to the table
    for index, row in df.iterrows():
        pdf.set_font(family="Times", size=10)
        pdf.set_text_color(80,80,80)
        pdf.cell(w=30, h=8, txt=str(row["product_id"]), border=1)
        pdf.cell(w=70, h=8, txt=str(row["product_name"]), border=1)
        pdf.cell(w=35, h=8, txt=str(row["amount_purchased"]), border=1)
        pdf.cell(w=30, h=8, txt=str(row["price_per_unit"]), border=1)
        pdf.cell(w=30, h=8, txt=str(row["total_price"]), border=1, ln=1)

    # Last line of table with total sum
    total_sum = df["total_price"].sum()
    pdf.set_font(family="Times", size=10)
    pdf.set_text_color(80,80,80)
    pdf.cell(w=30, h=8, txt="", border=1)
    pdf.cell(w=70, h=8, txt="", border=1)
    pdf.cell(w=35, h=8, txt="", border=1)
    pdf.cell(w=30, h=8, txt="", border=1)
    pdf.cell(w=30, h=8, txt=str(total_sum), border=1, ln=1)

    # Add total sum sentence
    pdf.set_font(family="Times", size=15, style="B")
    pdf.cell(w=30, h=8, txt=f"The total price is {total_sum}", ln=1)

    # Add company name and logo
    pdf.set_font(family="Times", size=15, style="B")
    pdf.cell(w=30, h=8, txt=f"PythonWow")
    pdf.image("python_wow.png", w=10)

    # Where the pdf will be located
    pdf.output(f"PDFs/{filename}.pdf")
    




