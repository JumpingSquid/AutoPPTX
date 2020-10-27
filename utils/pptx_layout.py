"""
Pptx layout contains two modules: Designer and Manager
Designer: arrange the location of objects, and pass the design structure to Manager
Manager: receive the design structure from Designer, use the data from DataPreprocessor to create the prs accordingly

Design structure (previously the config):
Every objects on the slide, will be given an unique id. Designer will arrange the location of every objects, and
provide a json format file to the manger, specifying the location of each uid
"""


from pptx.dml.color import RGBColor
from pptx.enum.text import PP_PARAGRAPH_ALIGNMENT
from pptx.enum.lang import MSO_LANGUAGE_ID
from pptx.util import Inches, Pt

import data2chart
from data2table import TableCreator
from utils.pptx_params import textbox


class PrsLayoutManager:

    def __init__(self, presentation):
        self.presentation = presentation
        self.prs_height = presentation.slide_height
        self.prs_width = presentation.slide_width
        self.data_container = None
        self.layout_design = None
        self.table_creator = TableCreator()

        # read in the title config
        # if no title in the config, use the blank slide instead
        if self.layout_design["title"] is not None:
            self.title_config = self.layout_design["title"]
            bullet_slide_layout = self.presentation.slide_layouts[5]
        else:
            bullet_slide_layout = self.presentation.slide_layouts[6]

        self.slide = self.presentation.slides.add_slide(bullet_slide_layout)

        # read in the chart config
        self.chart_config = self.layout_design["chart"]

        # without specified chart id, the chart is created based on the order of the chart list
        if not self.layout_design["setting"]["chart_id_sepcifed"]:
            self._chart_id = 0
            self.chart_rescale = self.chart_config[self._chart_id]["rescale"]
            self.chart_num_on_slide = len(self.chart_config)

        # beta: allow for customized layout for charts, but require the user to provide fully defined layout
        # the user should provide a list contains (1) origin and (2) the box size for each chart
        if "custom_layout" not in self.layout_design:
            self.custom_layout_flag = False
            self.chart_origin_anchor = [Inches(0), Inches(1.65)]
            print("INFO: ", self.chart_num_on_slide, "chart(s) required")
            self.chart_box_size = [self.prs_width / max(3, self.chart_num_on_slide) * self.chart_rescale[0],
                                   self.prs_height * 0.7 * self.chart_rescale[1]]
        elif "custom_layout" in self.layout_design:
            self.custom_layout_flag = True
            self.custom_layout_config = self.layout_design["custom_layout"]
            self.chart_origin_anchor = self.layout_design["custom_layout"]["origin"][self._chart_id]
            print("INFO: ", "Self-defined layout is used")
            self.chart_box_size = self.layout_design["custom_layout"]["size"][self._chart_id]

        # store the object that exists on the slide
        self.object_pool = []

    def slide_chart_layout(self, slide, chart_uid):
        # load the format of the chart
        chart_created_config = self.chart_config[chart_uid]
        if "format" not in chart_created_config:
            chart_created_config["format"] = None

        print("INFO: chart type", chart_created_config["type"], "require only 1 chart, full space will be used")
        chartcreator = data2chart.ChartCreator(chart_format=chart_created_config["format"])
        chartcreator.add_chart(slide, chart_type=chart_created_config["type"])
        return slide

    def add_chart_on_slide(self, slide, chart_uid, location):
        # try to create chart, if there is ZeroDivisionError for creating chart data, skip the chart first
        # otherwise stop the task
        self.slide_chart_layout(slide, chart_uid)
        self.chart_rescale = self.chart_config[self._chart_id]["rescale"]
        print(self._chart_id, "chart is added")
        return slide

    def add_text_on_slide(self, slide, text_uid, location):
        # create text box
        shapes = slide.shapes
        title_shape = shapes.title
        text_frame = title_shape.text_frame
        text_frame.clear()
        p = text_frame.paragraphs[0]

        # text box content is determined here
        p.text = self.title_config
        p.alignment = PP_PARAGRAPH_ALIGNMENT.LEFT
        p.font.name = textbox("title_font")
        p.font.size = Pt(40)
        p.font.language_id = MSO_LANGUAGE_ID.TRADITIONAL_CHINESE
        return slide

    def add_table_on_slide(self, slide, table_uid, location):
        # the position of the comment
        # location should be the arguments!
        left = self.table_creator.origin[0]
        top = self.table_creator.origin[1]
        width = self.table_creator.width
        height = self.table_creator.height
        row_num = self.table_creator.row_num
        col_num = self.table_creator.col_num
        table = slide.shapes.add_table(row_num, col_num, left, top, width, height).table
        self.table_creator.table_data_fill(table)
        return slide

    # create a warning tag on the chart where the sample size is below 15

    # draw the layout design on the screen, assist checking the number and the location of object created
    def layout_describe(self):

        def square_painter(draw_lst, x, y, lx, ly):
            draw_lst += [(x, y+y_bias) for y_bias in range(ly)]
            draw_lst += [(x+lx, y+y_bias) for y_bias in range(ly)]
            draw_lst += [(x+x_bias, y) for x_bias in range(lx)]
            draw_lst += [(x+x_bias, y+ly) for x_bias in range(lx)]
            return draw_lst

        draw_point = []
        draw_point = square_painter(draw_point, 0, 0, 55, 17)
        for obj in self.layout_design:
            draw_point = square_painter(draw_point,
                                        int(obj[0][0]/Inches(13.33)*55),
                                        int(obj[0][1]/Inches(7.5)*17),
                                        int(obj[1][0]/Inches(13.33)*55),
                                        int(obj[1][1]/Inches(7.5)*17))

        for y_painter in range(0, 18):
            drawer = ""
            for x_painter in range(0, 56):
                if (x_painter, y_painter) in draw_point:
                    drawer += "#"
                else:
                    drawer += " "
            print(drawer)

    # set the data container
    def _set_data_container(self, data_container):
        self.data_container = data_container


class PrsLayoutDesigner:
    def __init__(self):
        self.data = None
