import pandas as pd
import glob
from fpdf import FPDF
from pathlib import Path

filepaths = glob.glob('invoices/*.xlsx')
# choose the files in the invoice folder specify the xlsx files


for filepath in filepaths:
    df = pd.read_excel(filepath, sheet_name='Sheet 1')
    # we need to provide filepath for each filepath, and which sheet

    # we can create pdf file
    pdf = FPDF(orientation='p', unit='mm', format='A4')
    pdf.add_page()
    filename = Path(filepath).stem
    # we import pathlib with Path, we extract the Path() which gives us the file name in form of invoice/filename
    # we then put .stem to isolate the filename
    invoice_nr = filename.split('-')[0]
    # when we split we get [10001, 2023.1.18]
    # we then put [0] to get 10001
    pdf.set_font(family='Times', size=16, style='B')
    pdf.cell(w=50, h=8, txt=f'Invoice nr.{invoice_nr}')
    pdf.output(f"PDFs/{filename}.pdf")
