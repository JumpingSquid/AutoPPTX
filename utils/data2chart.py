# This module is used to create chart from a data
# The requirement of the data is to have only two columns
# The first column should be the categorical
# The second column can be numerical or categorical variable
# Two output format: 1. (data, slide) -> slide, 2. data -> presentation

from pptx.chart.data import CategoryChartData
from pptx.dml.color import RGBColor
from pptx.enum.chart import XL_CHART_TYPE
from pptx.enum.chart import XL_LEGEND_POSITION, XL_DATA_LABEL_POSITION
from pptx.util import Pt
from pptx import Presentation
from pptx.util import Inches


class ChatCreator:

    def __init__(self,
                 origin=None,
                 size=None,
                 chart_category=None,
                 chart_series=None,
                 chart_type="line",
                 chart_format=None
                 ):

        if chart_series is not None:
            self.chart_series = chart_series
            if chart_category is not None:
                self.chart_category = chart_category

        self.origin = origin
        self.size = size
        self.chart_type = chart_type

        self.chart_format = self.chart_format_infer()

        if chart_format is not None:
            for k in chart_format:
                self.chart_format[k] = chart_format[k]

    def chart_create(self, slide):
        # define categorical data
        chart_data = CategoryChartData()
        chart_data.categories = [cat for cat in self.chart_category]

        # define value data -
        for series in self.chart_series:
            chart_data.add_series(series[0], series[1])

        # set the position of the chart
        x, y = self.origin[0], self.origin[1]
        cx, cy = self.size[0], self.size[1]

        # add chart to slide
        chart_type_dict = {"hist": XL_CHART_TYPE.COLUMN_CLUSTERED,
                           "stackbar": XL_CHART_TYPE.BAR_STACKED_100,
                           "stackcolumn": XL_CHART_TYPE.COLUMN_STACKED_100,
                           "bar": XL_CHART_TYPE.BAR_CLUSTERED,
                           "pie": XL_CHART_TYPE.PIE,
                           "line": XL_CHART_TYPE.LINE}

        graphic_frame = slide.shapes.add_chart(
            chart_type_dict[self.chart_type], x, y, cx, cy, chart_data
        )

        # chart default theme setting
        if self.chart_type == "pie":
            self.piechartformat(graphic_frame.chart)
        elif self.chart_type == "line":
            self.linechartformat(graphic_frame.chart)
        else:
            self.chartformat(graphic_frame.chart)

        return slide

    def chart_create_on_slide(self, slide):
        if self.chart_type in ["hist", "stackbar", "bar", "pie", "line", "stackcolumn"]:
            return self.chart_create(slide)
        else:
            raise ValueError("Unknow chart type: " + str(self.chart_type) + "is given")

    def chartformat(self, chart):
        chart.value_axis.visible = False
        chart.value_axis.has_major_gridlines = False
        chart.font.size = self.chart_format["chart_font_size"]

        # set title
        if self.chart_format["chart_title"] is not None:
            chart.has_title = True
            chart.chart_title.has_text_frame = True
            chart.chart_title.text_frame.text = self.chart_format["chart_title"]
            chart.chart_title.text_frame.paragraphs[0].font.size = Pt(13)

        # set legends
        chart.has_legend = self.chart_format["legend_bool"]
        if self.chart_format["legend_bool"]:
            chart.legend.position = XL_LEGEND_POSITION.BOTTOM
            chart.legend.font.size = self.chart_format["legend_font_size"]

            chart.legend.include_in_layout = False

        # set data labels
        chart.plots[0].has_data_labels = self.chart_format["label_bool"]
        if self.chart_format["label_bool"]:
            labels = chart.plots[0].data_labels
            labels.number_format_is_linked = False
            labels.number_format = self.chart_format["label_number_format"]
            labels.font.size = self.chart_format["label_font_size"]

        # color
        if self.chart_format["colormap"] is None:
            color_r = 0
            color_g = 0
            color_b = 0
            colorgradient = len(chart.series)
            for series in chart.series:
                fill = series.format.fill  # fill the legend as well
                fill.solid()
                fill.fore_color.rgb = RGBColor(color_r, color_g, color_b)
                color_r += int(255/colorgradient)
                color_g += int(255/colorgradient)
                color_b += int(255/colorgradient)
        else:
            # color map format
            # [[color_R1, color_G1, color_B1], [color_R2, color_G2, color_B2],...]
            colormap = self.chart_format["colormap"]
            colormap_index = 0
            for series in chart.series:
                fill = series.format.fill  # fill the legend as well
                fill.solid()
                fill.fore_color.rgb = RGBColor(colormap[colormap_index % len(colormap)][0],
                                               colormap[colormap_index % len(colormap)][1],
                                               colormap[colormap_index % len(colormap)][2])
                colormap_index += 1
        return

    def linechartformat(self, chart):
        chart.value_axis.visible = False
        chart.value_axis.has_major_gridlines = False
        chart.font.size = self.chart_format["chart_font_size"]

        # set legends
        chart.has_legend = self.chart_format["legend_bool"]
        if self.chart_format["legend_bool"]:
            chart.legend.position = XL_LEGEND_POSITION.BOTTOM
            chart.legend.font.size = self.chart_format["legend_font_size"]
            chart.legend.include_in_layout = False

        # set data labels
        chart.plots[0].has_data_labels = self.chart_format["label_bool"]
        if self.chart_format["label_bool"]:
            labels = chart.plots[0].data_labels
            labels.number_format_is_linked = False
            labels.number_format = self.chart_format["label_number_format"]
            labels.font.size = self.chart_format["label_font_size"]

        # color
        if self.chart_format["colormap"] is None:
            color_r = 0
            color_g = 0
            color_b = 0
            colorgradient = len(chart.series)
            for series in chart.series:
                fill = series.format.fill  # fill the legend as well
                fill.solid()
                fill.fore_color.rgb = RGBColor(color_r, color_g, color_b)
                color_r += int(255/colorgradient)
                color_g += int(255/colorgradient)
                color_b += int(255/colorgradient)
        else:
            # color map format
            # [[color_R1, color_G1, color_B1], [color_R2, color_G2, color_B2],...]
            colormap = self.chart_format["colormap"]
            colormap_index = 0
            for series in chart.series:
                fill = series.format.fill  # fill the legend as well
                fill.solid()
                fill.fore_color.rgb = RGBColor(colormap[colormap_index % len(colormap)][0],
                                               colormap[colormap_index % len(colormap)][1],
                                               colormap[colormap_index % len(colormap)][2])
                colormap_index += 1
        return

    def piechartformat(self, chart):
        chart.has_legend = self.chart_format["legend_bool"]
        chart.font.size = self.chart_format["chart_font_size"]
        # set title
        if self.chart_format["chart_title"] is not None:
            chart.has_title = True
            chart.chart_title.has_text_frame = True
            chart.chart_title.text_frame.text = self.chart_format["chart_title"]
            chart.chart_title.text_frame.paragraphs[0].font.size = Pt(13)

        if self.chart_format["legend_bool"]:
            chart.legend.position = XL_LEGEND_POSITION.BOTTOM
            chart.legend.include_in_layout = False
            chart.legend.font.size = self.chart_format["legend_font_size"]

        chart.plots[0].has_data_labels = self.chart_format["label_bool"]
        if self.chart_format["label_bool"]:
            data_labels = chart.plots[0].data_labels
            data_labels.number_format = self.chart_format["label_number_format"]
            data_labels.position = XL_DATA_LABEL_POSITION.OUTSIDE_END

        if self.chart_format["colormap"] is not None:
            colormap = self.chart_format["colormap"]
            colormap_index = 0
            for point in chart.series[0].points:
                fill = point.format.fill  # fill the legend as well
                fill.solid()
                fill.fore_color.rgb = RGBColor(colormap[colormap_index % len(colormap)][0],
                                               colormap[colormap_index % len(colormap)][1],
                                               colormap[colormap_index % len(colormap)][2])
                colormap_index += 1

    def newslide(self) -> Presentation():
        prs = Presentation()
        # slide size: 16:9
        prs.slide_width = Inches(13.33)
        prs.slide_height = Inches(7.5)
        bullet_slide_layout = prs.slide_layouts[6]
        slide = prs.slides.add_slide(bullet_slide_layout)
        self.origin = [Inches(0), Inches(1.65)]
        self.size = [Inches(13), Inches(6)]
        self.chart_create(slide)
        return prs

    def pandas_to_ppt_series(self, data):
        import pandas
        if isinstance(data, pandas.DataFrame):
            chart_series = []
            for col in data.columns:
                chart_series.append((str(col), data[col]))

            self.chart_category = data.index.to_list()
            self.chart_series = chart_series

        elif isinstance(data, pandas.Series):
            self.chart_category = data.index.to_list()
            self.chart_series = [("series_1", tuple(data.values.to_list()))]

    def chart_format_infer(self):
        chart_format = {"legend_bool": True,
                        "label_bool": True,
                        "chart_bool": True,
                        "colormap": None,
                        "legend_font_size": Pt(9),
                        "label_font_size": Pt(9),
                        "chart_font_size": Pt(9),
                        "label_number_format": "0.0"}

        if all([all(float(number_ele).is_integer() for number_ele in x[1]) for x in self.chart_series]):
            chart_format["label_number_format"] = "0"

        return chart_format
