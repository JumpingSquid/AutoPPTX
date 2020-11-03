"""
PrsParamsManager is used to store and manage all parameters. Although this design requires multiple initialization,
it will allow the user to access all parameters in one place.

In the future, a more sophisticated implementation is needed if more parameter is added. In particular, it should avoid
redundant loading of all parameters.
"""


from pptx.util import Pt
from pptx.util import Inches
from pptx.dml.color import RGBColor
from pptx.enum.text import PP_PARAGRAPH_ALIGNMENT
from pptx.enum.lang import MSO_LANGUAGE_ID


class PrsParamsManager:

    def __init__(self):
        self.alignment = dict(left=PP_PARAGRAPH_ALIGNMENT.LEFT, right=PP_PARAGRAPH_ALIGNMENT.RIGHT,
                              center=PP_PARAGRAPH_ALIGNMENT.CENTER)
        self.text_lang = dict(tc=MSO_LANGUAGE_ID.TRADITIONAL_CHINESE)
        self.color_map_dict = dict(
            sunshine=[[212, 161, 167], [235, 214, 120], [254, 221, 98], [235, 171, 120], [235, 131, 95],
                      [245, 63, 43]],
            forest=[[87, 151, 115], [145, 96, 38], [166, 164, 167], [229, 233, 239], [89, 56, 0]],
            sea=[[23, 55, 94], [103, 149, 222], [112, 193, 226], [3, 193, 226], [3, 193, 161]],
            sea_reverse=[[3, 193, 161], [3, 193, 226], [112, 193, 226], [103, 149, 222], [180, 214, 219]],
            coldwarm=[[245, 227, 234], [212, 161, 167], [235, 214, 120], [112, 193, 226], [103, 149, 222],
                      [180, 214, 219]])

        self._default_chart_format = dict(chart_title="", chart_type="bar", legend_bool=True, label_bool=True,
                                          chart_bool=True, colormap=None, legend_font_size=Pt(12),
                                          label_font_size=Pt(12), chart_font_size=Pt(12), label_number_format="0.0")
        self._default_text_format = dict(alignment="left", font="title_font", font_size=12,
                                         font_color=RGBColor(0, 0, 0))
        self._default_textbox_format = dict(color="no_fill")
        self._default_table_format = dict(font_size=24, alignment='left')

    def get_chart_format(self):
        return self._default_chart_format

    def get_text_format(self):
        return self._default_text_format

    def get_textbox_format(self):
        return self._default_textbox_format

    def get_table_format(self):
        return self._default_table_format

    def set_default_chart_format(self, key=None, value=None, format_dict=None):
        if (key in self._default_chart_format) and (value is not None):
            self._default_chart_format[key] = value
        if format_dict is not None:
            for fkey in format_dict:
                if fkey in format_dict:
                    self._default_chart_format[fkey] = format_dict[fkey]
        return 1

    def set_default_text_format(self, key=None, value=None, format_dict=None):
        if (key in self._default_text_format) and (value is not None):
            self._default_text_format[key] = value
        if format_dict is not None:
            for fkey in format_dict:
                if fkey in format_dict:
                    self._default_text_format[fkey] = format_dict[fkey]
        return 1

    def set_default_textbox_format(self, key=None, value=None, format_dict=None):
        if (key in self._default_textbox_format) and (value is not None):
            self._default_textbox_format[key] = value
        if format_dict is not None:
            for fkey in format_dict:
                if fkey in format_dict:
                    self._default_textbox_format[fkey] = format_dict[fkey]
        return 1

    def set_default_table_format(self, key=None, value=None, format_dict=None):
        if (key in self._default_table_format) and (value is not None):
            self._default_table_format[key] = value
        if format_dict is not None:
            for fkey in format_dict:
                if fkey in format_dict:
                    self._default_table_format[fkey] = format_dict[fkey]
        return 1

    def color_map(self, index):
        if index in self.color_map_dict:
            return self.color_map_dict[index]
        elif 0 <= index < len(self.color_map_dict):
            key_lst = [k for k in self.color_map_dict.keys()]
            return self.color_map_dict[key_lst[index]]
        else:
            raise ValueError("index should be the color map name or the index")

    @staticmethod
    def color(r, g, b):
        return RGBColor(r, g, b)

    @staticmethod
    def textbox(key):
        text_dict = dict(title_font="微軟正黑體", comment_font="微軟正黑體", chart_appendix_font="微軟正黑體",
                         sample_warn_font="微軟正黑體", sample_size_font="微軟正黑體", sample_warn_text="Sample size < 15",
                         sample_size_text="總樣本數:")
        return text_dict[key]
