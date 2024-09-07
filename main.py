import pandas as pd
import glob

filepaths = glob.glob('invoices/*.xlsx')
# choose the files in the invoice folder specify the xlsx files


for filepath in filepaths:
    df = pd.read_excel(filepath,sheet_name='Sheet 1')
    # we need to provide filepath for each filepath, and which sheet

