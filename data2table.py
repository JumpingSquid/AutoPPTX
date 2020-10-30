import pandas as pd

from base import ObjectWorker
from pptx.util import Inches, Pt
from pptx.enum.text import PP_PARAGRAPH_ALIGNMENT
from pptx.enum.lang import MSO_LANGUAGE_ID


class TableWorker(ObjectWorker):
    def __init__(self, prs=False):
        super(ObjectWorker, self).__init__(prs)

    def create_table(self, data, slide, obj_format, position, uid):
        x, y, w, h = position
        row_num = data['row_num']
        col_num = data['col_num']
        table = slide.shapes.add_table(row_num, col_num, x, y, w, h).table
        for row in range(row_num):
            for col in range(col_num):
                cell = table.cell(row, col)
                cell.text = data.iloc[row, col]

        table = self.table_format_setter(table, obj_format)

        self.uid_pool.append(uid)

        return table

    def table_format_setter(self, table, table_format):
        for row in table.rows:
            for cell in row.cells:
                for paragraph in cell.text_frame.paragraphs:
                    if "alignment" in table_format:
                        paragraph.alignment = self.alignment[table_format['alignment']]
                    for run in paragraph.runs:
                        if 'font_size' in table_format:
                            run.font.size = Pt(table_format['font_size'])
        return table

    @staticmethod
    def default_object_format(table_format):
        default_table_format = {'font_size': 24, 'alignment': 'left'}

        if table_format is None:
            return default_table_format

        for t in table_format:
            if t in default_table_format:
                default_table_format[t] = table_format[t]

        return table_format

