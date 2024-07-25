import pandas as pd
import glob

filepaths = glob.glob("invoices/*.xlsx")


for i in filepaths:
    df = pd.read_excel(i, sheet_name="Sheet 1")
    print(df)

