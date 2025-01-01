import pandas as pd

class ExcelReader:
    def __init__(self, excel_file):
        self.excel_file = excel_file
        pd.set_option('display.max_columns', None)

    def read_excel(self):
        df = pd.read_excel(self.excel_file)
        df.columns = df.columns.str.strip()
        df['COURSE TITLE OG'] = df['COURSE TITLE']
        df['COURSE TITLE'] = df['COURSE TITLE'].str.upper()
        return df