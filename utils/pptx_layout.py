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
from utils.pptx_params import textbox


class PrsLayoutManager:

    def __init__(self, presentation):
        self.presentation = presentation
        self.prs_height = presentation.slide_height
        self.prs_width = presentation.slide_width
        self.data_container = None
        self.layout_design = None

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

    def _set_processed_data(self, processed_data):
        self.processed_data = processed_data

    def slide_chart_layout(self, slide, chart_uid):
        # load the format of the chart
        chart_created_config = self.chart_config[chart_uid]
        if "format" not in chart_created_config:
            chart_created_config["format"] = None

        # basic chart type: hist, stacked bar, stacked column, pie
        print("INFO: chart type", chart_created_config["type"], "require only 1 chart, full space will be used")
        chartcreator = data2chart.ChartCreator(chart_format=chart_created_config["format"])

        if chart_created_config["type"] in ["hist", "stackbar", "stackcolumn"]:
            chartcreator.add_chart(slide)
            self.sample_size_legend(slide, self.data_container.samplecount, left=self.chart_origin_anchor[0])
            self.object_pool.append((self.chart_origin_anchor, self.chart_box_size, chart_created_config["type"]))
            self.chart_origin_anchor = [self.chart_origin_anchor[0] + self.chart_box_size[0], Inches(1.65)]

        # basic chart type: pie
        elif chart_created_config["type"] in ["pie", "p"]:
            chartcreator.add_chart(slide)
            self.sample_size_legend(slide, self.data_container.totalsample, left=self.chart_origin_anchor[0])
            self.object_pool.append((self.chart_origin_anchor, self.chart_box_size, chart_created_config["type"]))

            # new function: add appendix under the chart created
            if chart_created_config["appendix"] is not None:
                appendix_bias = self.chart_origin_anchor[1] + self.chart_box_size[1]
                for appendix_ele in chart_created_config["appendix"]:
                    appendix_bias = self.chart_appendix(slide,
                                                        appendix_ele,
                                                        left=self.chart_origin_anchor[0],
                                                        top=appendix_bias,
                                                        width=self.chart_box_size[0])

            self.chart_origin_anchor = [self.chart_origin_anchor[0] + self.chart_box_size[0], Inches(1.65)]

        elif chart_created_config["type"] in ["sep_bar", "sep_b"]:
            chartcreator.add_chart(slide)
            self.sample_size_legend(slide, self.data_container.totalsample, left=self.chart_origin_anchor[0])
            self.object_pool.append((self.chart_origin_anchor, self.chart_box_size, chart_created_config["type"]))

            # new function: add appendix under the chart created
            if chart_created_config["appendix"] is not None:
                appendix_bias = self.chart_origin_anchor[1] + self.chart_box_size[1]
                for appendix_ele in chart_created_config["appendix"]:
                    appendix_bias = self.chart_appendix(slide,
                                                        appendix_ele,
                                                        left=self.chart_origin_anchor[0],
                                                        top=appendix_bias,
                                                        width=self.chart_box_size[0])

            self.chart_origin_anchor = [self.chart_origin_anchor[0] + self.chart_box_size[0], Inches(1.65)]

        elif chart_created_config["type"] == ["line", "l"]:
            chartcreator.add_chart(slide)
            self.sample_size_legend(slide, self.data_container.samplecount, left=self.chart_origin_anchor[0])
            self.chart_origin_anchor = [self.chart_origin_anchor[0] + self.chart_box_size[0], Inches(1.65)]

        elif chart_created_config["type"] in ["single_category_multiple_bar", "scmb"]:
            chart_num = self.data_container.category_num
            print("INFO: chart type", chart_created_config["type"], "requires", chart_num, "charts")
            self.sample_size_legend(slide, self.data_container.samplecount, left=self.chart_origin_anchor[0])
            chart_size_split_num = chart_num // 3
            if chart_num % 3 != 0:
                chart_size_split_num += 1
            self.chart_box_size = [self.chart_box_size[0] / chart_size_split_num, self.chart_box_size[1] / 3.2]
            for chart_id in range(0, chart_num):
                chartcreator = data2chart.ChartCreator(chart_format=chart_created_config["format"])
                slide = chartcreator.add_chart(slide)
                if chart_id % 3 != 2:
                    self.chart_origin_anchor = [self.chart_origin_anchor[0],
                                                self.chart_origin_anchor[1] + self.chart_box_size[1]]
                else:
                    self.chart_origin_anchor = [self.chart_origin_anchor[0] + self.chart_box_size[0],
                                                Inches(1.65)]
            if chart_num % 3 != 0:
                self.chart_origin_anchor = [self.chart_origin_anchor[0] + self.chart_box_size[0], Inches(1.65)]
            else:
                self.chart_origin_anchor = [self.chart_origin_anchor[0], Inches(1.65)]

        elif chart_created_config["type"] in ["multiple_value_column_bar", "mv"]:
            chartcreator = data2chart.ChartCreator(chart_format=chart_created_config["format"])
            slide = chartcreator.add_chart(slide)
            self.chart_origin_anchor = [self.chart_origin_anchor[0] + self.chart_box_size[0], Inches(1.65)]

        elif chart_created_config["type"] in ["multiple_category_multiple_value_column_bar", "mcmv"]:
            chart_num = self.data_container.category_num
            print("INFO: chart type", chart_created_config["type"], "requires", chart_num, "charts")

            chart_size_split_num = chart_num

            self.chart_box_size = [self.chart_box_size[0], self.chart_box_size[1] / chart_size_split_num]
            for chart_id in range(0, chart_num):
                slide = chartcreator.add_chart(slide)
                self.chart_origin_anchor = [self.chart_origin_anchor[0],
                                            self.chart_origin_anchor[1] + self.chart_box_size[1]]

            self.chart_origin_anchor = [self.chart_origin_anchor[0] + self.chart_box_size[0], Inches(1.65)]

        elif chart_created_config["type"] in ["multiple_dummy_one_column_bar", "mdc"]:
            chart_num = len(self.chart_config[self._chart_id]["cat_columns"])
            print("INFO: chart type", chart_created_config["type"], "requires", chart_num, "charts")

            self.chart_box_size = [size_ele / chart_num for size_ele in self.chart_box_size]

            for chart_id in range(0, chart_num):
                chartcreator = data2chart.ChartCreator(chart_format=chart_created_config["format"])
                slide = chartcreator.add_chart(slide)
                self.chart_origin_anchor = [self.chart_origin_anchor[0] + self.chart_box_size[0], Inches(1.65)]

        else:
            raise ValueError("ERROR: Unknow type of chart,", chart_created_config["type"], "is introduced")

        return slide

    def add_chart_on_slide(self):
        # try to create chart, if there is ZeroDivisionError for creating chart data, skip the chart first
        # otherwise stop the task
        try:
            self.slide_chart_layout(self.slide)
            self.chart_rescale = self.chart_config[self._chart_id]["rescale"]
            print(self._chart_id, "chart is added")
        except ZeroDivisionError:
            print("WARN: the chart-", self._chart_id, "incurs ZeroDivisionError, skipped automatically")
            pass
        if self._chart_id < self.chart_num_on_slide - 1:
            self._chart_id += 1
            # reset the chart size and anchor after creating the chart if no pre-defined layout
            # o/w navigating to the next object's spec
            if not self.custom_layout_flag:
                self.chart_box_size = [self.prs_width / max(3, self.chart_num_on_slide) * self.chart_rescale[0],
                                       self.prs_height * 0.7 * self.chart_rescale[1]]
            else:
                self.chart_origin_anchor = self.custom_layout_config["origin"][self._chart_id]
                self.chart_box_size = self.custom_layout_config["size"][self._chart_id]
        return self.presentation

    def add_title_on_slide(self):
        # create title
        shapes = self.slide.shapes
        title_shape = shapes.title
        text_frame = title_shape.text_frame
        text_frame.clear()
        p = text_frame.paragraphs[0]
        p.text = self.title_config
        p.alignment = PP_PARAGRAPH_ALIGNMENT.LEFT
        p.font.name = textbox("title_font")
        p.font.size = Pt(40)
        p.font.language_id = MSO_LANGUAGE_ID.TRADITIONAL_CHINESE
        return self.presentation

    def add_table_on_slide(self, tablecreator):
        # the position of the comment
        left = tablecreator.origin[0]
        top = tablecreator.origin[1]
        width = tablecreator.width
        height = tablecreator.height
        row_num = tablecreator.row_num
        col_num = tablecreator.col_num
        table = self.slide.shapes.add_table(row_num, col_num, left, top, width, height).table
        tablecreator.table_data_fill(table)
        return self.presentation

    # create a warning tag on the chart where the sample size is below 15
    @staticmethod
    def sample_size_warning(slide, left=Inches(8.33), top=Inches(1)):
        # the position of the comment
        width = Inches(2.66)
        height = Inches(0.37)
        txBox = slide.shapes.add_textbox(left, top, width, height)

        # the color of the comment
        fill = txBox.fill
        fill.solid()
        fill.fore_color.rgb = RGBColor(200, 200, 200)

        # the content and the format of the comment
        tf = txBox.text_frame
        tf.clear()  # not necessary for newly-created shape
        p = tf.paragraphs[0]
        run = p.add_run()
        run.text = textbox("sample_warn_text")
        font = run.font
        font.name = textbox("sample_warn_font")
        font.size = Pt(16)
        return slide

    # draw the layout setting on the screen, assist checking the number and the location of object created
    def layout_describe(self):

        def square_painter(draw_lst, x, y, lx, ly):
            draw_lst += [(x, y+y_bias) for y_bias in range(ly)]
            draw_lst += [(x+lx, y+y_bias) for y_bias in range(ly)]
            draw_lst += [(x+x_bias, y) for x_bias in range(lx)]
            draw_lst += [(x+x_bias, y+ly) for x_bias in range(lx)]
            return draw_lst

        draw_point = []
        draw_point = square_painter(draw_point, 0, 0, 55, 17)
        for object in self.object_pool:
            draw_point = square_painter(draw_point,
                                        int(object[0][0]/Inches(13.33)*55),
                                        int(object[0][1]/Inches(7.5)*17),
                                        int(object[1][0]/Inches(13.33)*55),
                                        int(object[1][1]/Inches(7.5)*17))

        for y_painter in range(0, 18):
            drawer = ""
            for x_painter in range(0, 56):
                if (x_painter, y_painter) in draw_point:
                    drawer += "#"
                else:
                    drawer += " "
            print(drawer)

    # add the sample size legend for the chart created
    @staticmethod
    def sample_size_legend(slide, samplecount, left=Inches(0), top=Inches(7.2)):
        width = Inches(2)
        height = Inches(0.27)
        txBox = slide.shapes.add_textbox(left, top, width, height)

        # the color of the comment
        fill = txBox.fill
        fill.solid()
        fill.fore_color.rgb = RGBColor(255, 255, 255)

        # the content and the format of the comment
        tf = txBox.text_frame
        tf.clear()  # not necessary for newly-created shape
        p = tf.paragraphs[0]
        run = p.add_run()
        if isinstance(samplecount, str):
            run.text = samplecount
        else:
            run.text = textbox("sample_size_text") + str(samplecount)
        font = run.font
        font.name = textbox("sample_size_font")
        p.alignment = PP_PARAGRAPH_ALIGNMENT.LEFT
        font.size = Pt(10)


class PrsLayoutDesigner:
    def __init__(self):
        self.data = None
