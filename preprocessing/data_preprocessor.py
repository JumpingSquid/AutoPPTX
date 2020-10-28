import pandas as pd
import numpy as np

from pptx.chart.data import CategoryChartData


class DataProcessor:

    def __init__(self, data):
        self.data = data

    @staticmethod
    def pandas_to_ppt_table(dataframe):
        assert isinstance(dataframe, pd.DataFrame)

        data = CategoryChartData()
        col_lst = dataframe.columns.to_list()
        data.categories = col_lst

        for i, r in dataframe.iterrows():
            data.add_series(i, tuple([r[col] for col in col_lst]))

        return data


class DataContainer:

    def __init__(self):
        self.data = None
