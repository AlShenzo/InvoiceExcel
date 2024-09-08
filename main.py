import pandas as pd
import glob
from fpdf import FPDF
from pathlib import Path

filepaths = glob.glob('invoices/*.xlsx')
# choose the files in the invoice folder specify the xlsx files


for filepath in filepaths:

    # we can create pdf file
    pdf = FPDF(orientation='p', unit='mm', format='A4')
    pdf.add_page()

    filename = Path(filepath).stem
    # we import pathlib with Path, we extract the Path() which gives us the file name in form of invoice/filename
    # we then put .stem to isolate the filename
    invoice_nr, date = filename.split('-')
    # when we split we get [10001, 2023.1.18]
    # we then put [0] to get 10001
    # we could duplicate and write
    # date = filename.split('-')[1] as separate and keep the invoice [0]
    # but we can write it as above.

    pdf.set_font(family='Times', size=16, style='B')
    pdf.cell(w=50, h=8, txt=f'Invoice nr.{invoice_nr}', ln=1)
    # ln=1 to create a break line
    pdf.set_font(family='Times', size=16, style='B')
    pdf.cell(w=50, h=8, txt=f'Date: {date}', ln=1)

    df = pd.read_excel(filepath, sheet_name='Sheet 1')
    # we need to provide filepath for each filepath, and which sheet

    # now add header
    columns = df.columns
    # panda reads the columns headers, and we convert into a list
    # but we don't actually need it into a list
    # we can iterate it over index object
    columns = [item.replace('_', ' ').title() for item in columns]
    # we replace the _ with space and capitalize first letter of each time
    # for loop for do that for all 3 files
    pdf.set_font(family='Times', size=10, style='B')
    pdf.set_text_color(80, 80, 80)
    pdf.cell(w=30, h=8, txt=str(columns[0]), border=1)
    pdf.cell(w=60, h=8, txt=str(columns[1]), border=1)
    pdf.cell(w=40, h=8, txt=str(columns[2]), border=1)
    pdf.cell(w=30, h=8, txt=str(columns[3]), border=1)
    pdf.cell(w=30, h=8, txt=str(columns[4]), border=1, ln=1)

    # add rows
    for index, row in df.iterrows():
        pdf.set_font(family='Times', size=10)
        pdf.set_text_color(80, 80, 80)
        pdf.cell(w=30, h=8, txt=str(row['product_id']), border=1)
        pdf.cell(w=60, h=8, txt=str(row['product_name']), border=1)
        pdf.cell(w=40, h=8, txt=str(row['amount_purchased']), border=1)
        pdf.cell(w=30, h=8, txt=str(row['price_per_unit']), border=1)
        pdf.cell(w=30, h=8, txt=str(row['total_price']), border=1, ln=1)
        # border =1 to add border
        # after the last cell we give space to the next row

    # now we add 1 cell for sum of column, outside for loop
    total_sum = df['total_price'].sum()
    pdf.cell(w=30, h=8, txt='', border=1)  # empty strings to create empty cells
    pdf.cell(w=60, h=8, txt='', border=1)
    pdf.cell(w=40, h=8, txt='', border=1)
    pdf.cell(w=30, h=8, txt='', border=1)
    pdf.cell(w=30, h=8, txt=str(total_sum), border=1, ln=1)

    # add total sum sentence
    pdf.set_font(family='Times', size=10, style='B')
    pdf.cell(w=30, h=8, txt=f'The total price is {total_sum}', ln=1)

    # company logo
    pdf.set_font(family='Times', size=10, style='B')
    pdf.cell(w=25, h=8, txt=f'PythonHow')
    pdf.image('pythonhow.png', w=10)

    pdf.output(f"PDFs/{filename}.pdf")
