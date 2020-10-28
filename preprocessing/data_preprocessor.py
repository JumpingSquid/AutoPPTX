import pandas as pd
import numpy as np

from pptx.chart.data import CategoryChartData


class DataProcessor:

    def __init__(self, prs_object_pool):
        self.prs_object_pool = prs_object_pool
        self.data_container = DataContainer()

        # direct execution
        self._data_preprocess_execute()

    def _data_preprocess_execute(self):
        for uid in self.prs_object_pool:
            obj_container = self.prs_object_pool[uid]
            uid = obj_container.uid
            data = obj_container.data
            if obj_container.obj_type == 'chart' and isinstance(data, pd.DataFrame):
                data = self.pandas_to_ppt_chart_data(data)
            self.data_container.add_data(uid, data)

    def data_container_export(self):
        return self.data_container

    @staticmethod
    def pandas_to_ppt_chart_data(dataframe):
        assert isinstance(dataframe, pd.DataFrame)

        data = CategoryChartData()
        col_lst = dataframe.columns.to_list()
        data.categories = col_lst

        for i, r in dataframe.iterrows():
            data.add_series(str(i), tuple([r[col] for col in col_lst]))

        return data


class DataContainer:

    def __init__(self):
        self.data = {}

    def add_data(self, uid, data):
        self.data[uid] = data

    def get_data(self, uid):
        return self.data[uid]
