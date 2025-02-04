import pandas as pd
from utils import with_loading_animation
from config import EXCEL_FILENAME

class ExcelReader:
    def __init__(self, excel_file=EXCEL_FILENAME):
        self.excel_file = excel_file
        pd.set_option('display.max_columns', None)

    @with_loading_animation("Reading Excel file")
    def read_excel(self):
        df = pd.read_excel(self.excel_file)
        # Remove leading/trailing whitespaces from column names
        df.columns = df.columns.str.strip()
        return df