import pandas as pd

from pptx.util import Pt

from base import ObjectWorker
from utils.pptx_params import PrsParamsManager


class TableWorker(ObjectWorker):
    def __init__(self, prs=False):
        super(ObjectWorker, self).__init__()

        if prs:
            # initializing a basic presentation file when no existing file is given
            self.prs = self.create_prs()

        self.uid_pool = []
        self.params = PrsParamsManager()

    def create_table(self, data: pd.DataFrame, slide, obj_format, position, uid):
        x, y, w, h = position
        row_num = len(data) + 1
        col_num = len(data.columns) + 1

        col_lst = data.columns.to_list()
        index_lst = data.index.to_list()
        table = slide.shapes.add_table(row_num, col_num, x, y, w, h).table

        for col in range(1, col_num):
            cell = table.cell(0, col)
            cell.text = str(col_lst[col-1])

        for row in range(1, row_num):
            cell = table.cell(row, 0)
            cell.text = str(index_lst[row-1])

        for row in range(1, row_num):
            for col in range(1, col_num):
                cell = table.cell(row, col)
                cell.text = str(data.iloc[row-1, col-1])

        obj_format = self.default_object_format(obj_format)
        self.table_format_setter(table, obj_format)

        self.uid_pool.append(uid)

        return slide

    def table_format_setter(self, table, table_format):
        for row in table.rows:
            for cell in row.cells:
                for paragraph in cell.text_frame.paragraphs:
                    if "alignment" in table_format:
                        paragraph.alignment = self.params.alignment[table_format['alignment']]
                    for run in paragraph.runs:
                        if 'font_size' in table_format:
                            run.font.size = Pt(table_format['font_size'])
        return table

    def default_object_format(self, table_format):
        default_table_format = self.params.get_table_format()

        if table_format is None:
            return default_table_format

        for t in table_format:
            if t in default_table_format:
                default_table_format[t] = table_format[t]

        return table_format

