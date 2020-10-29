import pandas as pd


class TextCreator:
    def __init__(self):
        self.origin = None
        self.width = None
        self.height = None
        self.row_num = None
        self.col_num = None
        self.config = None

    def table_data_fill(self, data, table):
        for row in range(self.row_num):
            for col in range(self.col_num):
                cell = table.cell(row, col)
                cell.text = data.iloc[row, col]
        return
