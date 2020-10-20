import pandas as pd


class TableCreator:
    def __init__(self, config: dict, data: pd.DataFrame):
        self.origin = config["origin"]
        self.width = config["width"]
        self.height = config["height"]
        self.row_num = len(data)
        self.col_num = len(data.columns)
        self.data = data
        self.config = config

    def table_data_fill(self, table):
        for row in range(self.row_num):
            for col in range(self.col_num):
                cell = table.cell(row, col)
                cell.text = self.data.iloc[row, col]
        return
