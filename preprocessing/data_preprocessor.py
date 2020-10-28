import pandas as pd
import numpy as np

from pptx.chart.data import CategoryChartData


class DataProcessor:

    def __init__(self, prs_object_pool):
        self.prs_object_pool = prs_object_pool
        self.data_container = DataContainer()

    def _data_preprocess_execute(self):
        for obj_container in self.prs_object_pool:
            uid = obj_container.uid
            data = obj_container.data
            self.data_container.add_data(uid, data)

    def data_container_export(self):
        return self.data_container

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
        self.data = {}

    def add_data(self, uid, data):
        self.data[uid] = data
